// MIME: página da web salva em arquivo único (.mht, .mhtml) e e-mail (.eml).
// As partes são decodificadas, as imagens de dentro da mensagem viram
// endereços locais e o HTML segue pelo mesmo caminho das páginas da web.
import { schema } from '../editor/schema.js';
import { blankModel, finalizeDoc } from './common.js';
import { htmlToBlocks } from './html.js';

const nodes = schema.nodes;
const M = schema.marks;
const LATIN1 = new TextDecoder('latin1');

function decoderFor(charset) {
  try { return new TextDecoder((charset || 'utf-8').trim().toLowerCase().replace(/^"|"$/g, '')); } catch { return new TextDecoder('utf-8'); }
}

// =?utf-8?B?...?= e =?iso-8859-1?Q?...?= nos cabeçalhos
export function decodeWords(s) {
  return String(s || '').replace(/(=\?[^?]+\?[bq]\?[^?]*\?=)\s+(?==\?)/gi, '$1').replace(/=\?([^?]+)\?([bq])\?([^?]*)\?=/gi, (m, cs, enc, data) => {
    try {
      let bytes;
      if (enc.toLowerCase() === 'b') bytes = Uint8Array.from(atob(data), (c) => c.charCodeAt(0));
      else bytes = qpBytes(data.replace(/_/g, ' '));
      return decoderFor(cs).decode(bytes);
    } catch { return m; }
  });
}

function qpBytes(s) {
  const out = [];
  for (let i = 0; i < s.length; i++) {
    const c = s[i];
    if (c === '=') {
      if (s[i + 1] === '\r' && s[i + 2] === '\n') { i += 2; continue; }
      if (s[i + 1] === '\n') { i += 1; continue; }
      const hx = s.substr(i + 1, 2);
      if (/^[0-9a-f]{2}$/i.test(hx)) { out.push(parseInt(hx, 16)); i += 2; continue; }
    }
    out.push(s.charCodeAt(i) & 0xff);
  }
  return Uint8Array.from(out);
}

function parseHeaders(text) {
  const headers = new Map();
  const lines = text.replace(/\r\n/g, '\n').split('\n');
  let cur = null;
  for (const line of lines) {
    if (/^[ \t]/.test(line) && cur) { headers.set(cur, headers.get(cur) + ' ' + line.trim()); continue; }
    const m = /^([^:]+):\s*(.*)$/.exec(line);
    if (!m) continue;
    cur = m[1].trim().toLowerCase();
    headers.set(cur, headers.has(cur) ? headers.get(cur) + ', ' + m[2] : m[2]);
  }
  return headers;
}

function param(value, name) {
  const m = new RegExp('(?:^|;)\\s*' + name + '\\s*=\\s*("([^"]*)"|[^;\\s]+)', 'i').exec(value || '');
  return m ? (m[2] != null ? m[2] : m[1]) : null;
}

// Uma parte: cabeçalhos, tipo e conteúdo já decodificado; multipart vira lista
export function parseMime(raw) {
  const split = /\r?\n\r?\n/.exec(raw);
  const head = split ? raw.slice(0, split.index) : raw;
  const body = split ? raw.slice(split.index + split[0].length) : '';
  const headers = parseHeaders(head);
  const ctype = headers.get('content-type') || 'text/plain';
  const type = ctype.split(';')[0].trim().toLowerCase();
  const part = { headers, type, charset: param(ctype, 'charset'), name: decodeWords(param(headers.get('content-disposition'), 'filename') || param(ctype, 'name') || '') };
  if (type.startsWith('multipart/')) {
    const boundary = param(ctype, 'boundary');
    part.parts = [];
    if (boundary) {
      const sep = '--' + boundary;
      const pieces = body.split(sep);
      for (let k = 1; k < pieces.length; k++) {
        let p = pieces[k];
        if (p.startsWith('--')) break;
        p = p.replace(/^\r?\n/, '').replace(/\r?\n$/, '');
        part.parts.push(parseMime(p));
      }
    }
    return part;
  }
  const cte = (headers.get('content-transfer-encoding') || '').trim().toLowerCase();
  if (cte === 'base64') {
    try { part.bytes = Uint8Array.from(atob(body.replace(/[^A-Za-z0-9+/=]/g, '')), (c) => c.charCodeAt(0)); } catch { part.bytes = new Uint8Array(0); }
  } else if (cte === 'quoted-printable') part.bytes = qpBytes(body);
  else part.bytes = Uint8Array.from(body, (c) => c.charCodeAt(0) & 0xff);
  return part;
}

function leaves(p, out = []) {
  if (p.parts) for (const c of p.parts) leaves(c, out);
  else out.push(p);
  return out;
}

function partText(p) {
  let cs = p.charset;
  if (!cs && p.type === 'text/html') {
    const head = LATIN1.decode(p.bytes.subarray(0, 2048));
    cs = (/<meta[^>]+charset=["']?([\w-]+)/i.exec(head) || [])[1];
  }
  return decoderFor(cs).decode(p.bytes);
}

// imagens e anexos citados no HTML por Content-Location ou cid:
function resourceMap(all) {
  const map = new Map();
  for (const p of all) {
    if (!p.bytes || p.type.startsWith('text/html')) continue;
    const url = URL.createObjectURL(new Blob([p.bytes], { type: p.type }));
    const loc = p.headers.get('content-location');
    const cid = (p.headers.get('content-id') || '').replace(/^<|>$/g, '').trim();
    if (loc) { map.set(loc.trim(), url); map.set(loc.trim().split('/').pop(), url); }
    if (cid) map.set('cid:' + cid, url);
  }
  return map;
}

function resolver(map, baseLoc) {
  return (src) => {
    if (!src) return null;
    if (map.has(src)) return map.get(src);
    if (baseLoc) {
      try { const abs = new URL(src, baseLoc).href; if (map.has(abs)) return map.get(abs); } catch { /* endereço inválido */ }
    }
    const tail = src.split(/[?#]/)[0].split('/').pop();
    return map.get(tail) || null;
  };
}

async function htmlPartBlocks(htmlPart, all, model) {
  const html = partText(htmlPart);
  const dom = new DOMParser().parseFromString(html, 'text/html');
  if (!model.properties.title) model.properties.title = (dom.querySelector('title') || {}).textContent || '';
  const map = resourceMap(all);
  return htmlToBlocks(dom.body, model, { resolve: resolver(map, htmlPart.headers.get('content-location')) });
}

function textBlocks(text) {
  return text.replace(/\r\n?/g, '\n').split('\n').map((line) => nodes.paragraph.create(null, line ? [schema.text(line)] : []));
}

export async function readMht(bytes) {
  const root = parseMime(LATIN1.decode(bytes));
  const all = leaves(root);
  const start = param(root.headers.get('content-type'), 'start');
  const htmlPart = (start && all.find((p) => (p.headers.get('content-id') || '').includes(start.replace(/^<|>$/g, '')))) || all.find((p) => p.type === 'text/html');
  const { model } = await blankModel();
  model.format = 'mht';
  const blocks = htmlPart ? await htmlPartBlocks(htmlPart, all, model) : textBlocks(all.filter((p) => p.type.startsWith('text/')).map(partText).join('\n'));
  return { model, doc: finalizeDoc(schema.node('doc', null, blocks.length ? blocks : [nodes.paragraph.create()]), model) };
}

function fmtSize(n) {
  return n >= 1048576 ? (n / 1048576).toFixed(1).replace('.', ',') + ' MB' : Math.max(1, Math.round(n / 1024)) + ' KB';
}

export async function readEml(bytes) {
  const root = parseMime(LATIN1.decode(bytes));
  const all = leaves(root);
  const { model } = await blankModel();
  model.format = 'eml';
  const h = root.headers;
  const subject = decodeWords(h.get('subject') || '').trim();
  model.properties.title = subject;
  const gray = [M.color.create({ a: { val: '595959' } })];
  const head = [nodes.paragraph.create({ outlineLvl: 0, spacing: { before: '0', after: '120', line: '240', lineRule: 'auto' } }, [schema.text(subject || '(sem assunto)', [M.b.create(), M.sz.create({ val: 32 })])])];
  let date = h.get('date') || '';
  const d = date ? new Date(date) : null;
  if (d && !Number.isNaN(d.getTime())) date = d.toLocaleString('pt-BR', { dateStyle: 'full', timeStyle: 'short' });
  for (const [label, key] of [['De', 'from'], ['Para', 'to'], ['Cc', 'cc'], ['Data', null]]) {
    const v = key ? decodeWords(h.get(key) || '').trim() : date;
    if (!v) continue;
    head.push(nodes.paragraph.create({ spacing: { before: '0', after: '0', line: '240', lineRule: 'auto' } }, [schema.text(label + ': ', [M.b.create(), ...gray]), schema.text(v, gray)]));
  }
  // corpo: prefere HTML, senão texto; o que sobra são anexos
  const alt = all.filter((p) => !/attachment/i.test(p.headers.get('content-disposition') || ''));
  const htmlPart = alt.find((p) => p.type === 'text/html');
  const textPart = alt.find((p) => p.type === 'text/plain');
  const line = nodes.paragraph.create({ pBdr: { bottom: { val: 'single', sz: '6', space: '1', color: 'BFBFBF' } }, spacing: { before: '0', after: '240', line: '240', lineRule: 'auto' } });
  const body = htmlPart ? await htmlPartBlocks(htmlPart, all, model) : textPart ? textBlocks(partText(textPart)) : [];
  const attachments = all.filter((p) => p !== htmlPart && p !== textPart && (p.name || /attachment/i.test(p.headers.get('content-disposition') || '')) && !(htmlPart && p.headers.get('content-id') && /^image\//.test(p.type) && !/attachment/i.test(p.headers.get('content-disposition') || '')));
  const tail = [];
  if (attachments.length) {
    tail.push(nodes.paragraph.create({ pBdr: { top: { val: 'single', sz: '6', space: '1', color: 'BFBFBF' } }, spacing: { before: '240', after: '0', line: '240', lineRule: 'auto' } }, [schema.text('Anexos', [M.b.create()])]));
    for (const a of attachments) tail.push(nodes.paragraph.create(null, [schema.text('📎 ' + (a.name || 'anexo') + ' (' + fmtSize(a.bytes ? a.bytes.length : 0) + ')')]));
  }
  const blocks = [...head, line, ...body, ...tail];
  return { model, doc: finalizeDoc(schema.node('doc', null, blocks), model) };
}
