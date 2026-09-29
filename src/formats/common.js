// Utilidades compartilhadas pelos formatos de importação e exportação.
import { EditorState } from 'prosemirror-state';
import { readDocx } from '../docx/reader.js';
import { blankDocxBytes } from '../docx/template.js';
import { resolvePastedLists } from '../editor/setup.js';
import { R } from '../editor/render.js';

export async function blankModel(author) {
  const bytes = await blankDocxBytes(author || '');
  return readDocx(bytes);
}

// Resolve marcadores provisórios de listas (vindos de HTML e Markdown)
export function finalizeDoc(doc, model, listMap) {
  const prev = R.model;
  R.model = model;
  try {
    const state = EditorState.create({ doc });
    const tr = state.tr;
    resolvePastedLists(tr, model, listMap);
    return tr.doc;
  } finally {
    R.model = prev;
  }
}

export function decodeText(bytes) {
  let bom = null;
  let enc = 'utf-8';
  let start = 0;
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) { bom = 'utf-8'; start = 3; }
  else if (bytes[0] === 0xff && bytes[1] === 0xfe) { bom = 'utf-16le'; enc = 'utf-16le'; start = 2; }
  else if (bytes[0] === 0xfe && bytes[1] === 0xff) { bom = 'utf-16be'; enc = 'utf-16be'; start = 2; }
  const body = bytes.subarray(start);
  let text;
  if (enc === 'utf-8') {
    try { text = new TextDecoder('utf-8', { fatal: true }).decode(body); }
    catch { enc = 'windows-1252'; text = new TextDecoder('windows-1252').decode(body); }
  } else {
    text = new TextDecoder(enc).decode(body);
  }
  const eol = /\r\n/.test(text) ? '\r\n' : /\r/.test(text) && !/\n/.test(text) ? '\r' : '\n';
  return { text, encoding: enc, bom: !!bom, eol };
}

export function encodeText(text, info) {
  const enc = (info && info.encoding) || 'utf-8';
  if (enc === 'utf-16le' || enc === 'utf-16be') {
    const buf = new Uint8Array(2 + text.length * 2);
    const le = enc === 'utf-16le';
    buf[0] = le ? 0xff : 0xfe;
    buf[1] = le ? 0xfe : 0xff;
    for (let i = 0; i < text.length; i++) {
      const c = text.charCodeAt(i);
      buf[2 + i * 2] = le ? c & 0xff : c >> 8;
      buf[3 + i * 2] = le ? c >> 8 : c & 0xff;
    }
    return buf;
  }
  if (enc === 'windows-1252') {
    // mantém a codificação original quando todos os caracteres cabem nela
    const out = new Uint8Array(text.length);
    const map = { '€': 0x80, '‚': 0x82, 'ƒ': 0x83, '„': 0x84, '…': 0x85, '†': 0x86, '‡': 0x87, 'ˆ': 0x88, '‰': 0x89, 'Š': 0x8a, '‹': 0x8b, 'Œ': 0x8c, 'Ž': 0x8e, '‘': 0x91, '’': 0x92, '“': 0x93, '”': 0x94, '•': 0x95, '–': 0x96, '—': 0x97, '˜': 0x98, '™': 0x99, 'š': 0x9a, '›': 0x9b, 'œ': 0x9c, 'ž': 0x9e, 'Ÿ': 0x9f };
    let ok = true;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      const c = ch.charCodeAt(0);
      if (c < 0x80 || (c >= 0xa0 && c <= 0xff)) out[i] = c;
      else if (map[ch] != null) out[i] = map[ch];
      else { ok = false; break; }
    }
    if (ok) return out;
  }
  const u8 = new TextEncoder().encode(text);
  if (info && info.bom) {
    const b = new Uint8Array(u8.length + 3);
    b.set([0xef, 0xbb, 0xbf]);
    b.set(u8, 3);
    return b;
  }
  return u8;
}

export function dirName(path) {
  if (!path) return '';
  const i = Math.max(path.lastIndexOf('\\'), path.lastIndexOf('/'));
  return i >= 0 ? path.slice(0, i) : '';
}

export function fileUrl(path) {
  const p = path.replace(/\\/g, '/');
  return 'file:///' + p.split('/').map((s, i) => (i === 0 && /^[a-z]:$/i.test(s) ? s : encodeURIComponent(s))).join('/');
}

export function resolveRelative(baseDir, src) {
  if (!src) return src;
  if (/^(https?:|data:|blob:|file:)/i.test(src)) return src;
  if (/^[a-z]:[\\/]/i.test(src) || src.startsWith('\\\\')) return fileUrl(src);
  if (!baseDir) return src;
  const parts = (baseDir.replace(/\\/g, '/') + '/' + decodeURI(src).replace(/\\/g, '/')).split('/');
  const out = [];
  for (const p of parts) {
    if (p === '..') out.pop();
    else if (p !== '.' && p !== '') out.push(p);
  }
  return fileUrl(out.join('/'));
}

export function loadImageSize(src, timeout = 4000) {
  return new Promise((resolve) => {
    const img = new Image();
    const t = setTimeout(() => resolve(null), timeout);
    img.onload = () => { clearTimeout(t); resolve({ w: img.naturalWidth, h: img.naturalHeight }); };
    img.onerror = () => { clearTimeout(t); resolve(null); };
    img.src = src;
  });
}

// tamanho em EMU limitado à largura do texto
export function fitSize(w, h, maxWidthPx) {
  if (!w || !h) return { cx: 0, cy: 0 };
  let ww = w, hh = h;
  if (ww > maxWidthPx) { hh = (hh * maxWidthPx) / ww; ww = maxWidthPx; }
  return { cx: Math.round(ww * 9525), cy: Math.round(hh * 9525) };
}

export function contentWidthPx(model) {
  const s = model.section;
  return (s.pgSz.w - s.pgMar.left - s.pgMar.right - (s.pgMar.gutter || 0)) / 15;
}
