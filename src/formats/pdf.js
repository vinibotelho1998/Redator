// PDF: lido com o pdf.js e remontado em parágrafos editáveis, como o Word faz
// ao abrir um PDF. O texto volta com fonte, tamanho, negrito, itálico, cor e
// links; títulos ganham nível de estrutura; listas com marcador viram listas;
// imagens entram onde aparecem; cabeçalho e rodapé repetidos em todas as
// páginas vão para o cabeçalho e o rodapé do documento; página sem texto, como
// um desenho, entra como imagem.
import { schema } from '../editor/schema.js';
import { R } from '../editor/render.js';
import { newHeaderFooter } from '../docx/writer.js';
import { blankModel, finalizeDoc } from './common.js';

const nodes = schema.nodes;
const M = schema.marks;

// marcadores de lista, inclusive os do Symbol e do Wingdings que chegam na área de uso privado
const BULLET = /^([•‣⁃∙▪▫●○◦■□◆➢–—·*\-])\s+(?=\S)/;
const NUMBERED = /^(\(?(\d{1,3}(\.\d{1,3})*|[a-z]|[ivx]{1,5})[.)])\s+(?=\S)/i;
// marcador que o PDF desenha como um item separado do texto
const LONE_BULLET = /^[•●▪◦■·–-]$/;

let loading = null;

function loadPdfJs() {
  if (window.RedatorPdf) return Promise.resolve(window.RedatorPdf);
  if (!loading) {
    loading = new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = new URL('../dist/pdf.js', location.href).href;
      s.onload = () => (window.RedatorPdf ? resolve(window.RedatorPdf) : reject(new Error('O leitor de PDF não carregou.')));
      s.onerror = () => { loading = null; reject(new Error('Não foi possível carregar o leitor de PDF.')); };
      document.head.appendChild(s);
    });
  }
  return loading;
}

// ---------------------------------------------------------------- geometria

function mul(m1, m2) {
  return [
    m1[0] * m2[0] + m1[2] * m2[1], m1[1] * m2[0] + m1[3] * m2[1],
    m1[0] * m2[2] + m1[2] * m2[3], m1[1] * m2[2] + m1[3] * m2[3],
    m1[0] * m2[4] + m1[2] * m2[5] + m1[4], m1[1] * m2[4] + m1[3] * m2[5] + m1[5],
  ];
}

function apply(m, x, y) {
  return [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];
}

function median(list) {
  if (!list.length) return 0;
  const s = [...list].sort((a, b) => a - b);
  return s[s.length >> 1];
}

const twip = (pt) => String(Math.round(pt * 20));

// ---------------------------------------------------------------- fontes

const FONT_ALIASES = {
  helvetica: 'Arial', arialmt: 'Arial', times: 'Times New Roman', timesroman: 'Times New Roman', courier: 'Courier New',
  timesnewroman: 'Times New Roman', timesnewromanps: 'Times New Roman', couriernew: 'Courier New', couriernewps: 'Courier New',
  arialnarrow: 'Arial Narrow', zapfdingbats: 'Wingdings', symbolmt: 'Symbol',
};

function genericFont(generic) {
  return generic === 'serif' ? 'Times New Roman' : generic === 'monospace' ? 'Courier New' : 'Arial';
}

// nomes sem família de verdade, como os que o driver de impressora gera
const OPAQUE_FONT = /^(TT[0-9A-F]+o?\d*|T\d+|F\d+|C\d+_\d+|CID\w*|Font\d*|Unnamed.*|[A-Z]{1,2}\d*)$/i;

function fontFamilyOf(raw, generic, known) {
  let n = String(raw || '').replace(/^[A-Z]{6}\+/, '');
  if (!n) return genericFont(generic);
  n = n.split(/[-,]/)[0].replace(/(PSMT|MT|PS)$/, '');
  const alias = FONT_ALIASES[n.toLowerCase()];
  if (alias) return alias;
  if (OPAQUE_FONT.test(n)) return genericFont(generic);
  // "TimesNewRoman" vira "Times New Roman", "SegoeUI" vira "Segoe UI"
  n = n.replace(/(Bold|Italic|Oblique|Regular|Semibold|SemiBold|Light|Medium|Black|Heavy|Book|Roman)+$/, '') || n;
  const fam = n.replace(/([a-z])([A-Z])/g, '$1 $2').replace(/([A-Z]+)([A-Z][a-z])/g, '$1 $2').trim();
  // fonte que não existe neste PC: usa uma parecida para a quebra de linha sair igual
  if (known && known.size && !known.has(fam.toLowerCase()) && !known.has(n.toLowerCase())) return genericFont(generic);
  return fam;
}

function fontStyle(page, fontName, styles, cache, known) {
  if (cache.has(fontName)) return cache.get(fontName);
  let name = '';
  let bold = false;
  let italic = false;
  try {
    if (page.commonObjs.has(fontName)) {
      const f = page.commonObjs.get(fontName);
      name = f.name || '';
      bold = !!(f.bold || f.black);
      italic = !!f.italic;
    }
  } catch { /* fonte sem dados */ }
  const low = name.toLowerCase().replace(/^[a-z]{6}\+/, '');
  if (/bold|black|heavy|semibold|demibold|,b$/.test(low)) bold = true;
  if (/italic|oblique|-it$|boldit$|,i$|,bi$/.test(low)) italic = true;
  const generic = styles && styles[fontName] ? styles[fontName].fontFamily : '';
  const st = { family: fontFamilyOf(name, generic, known), bold, italic };
  cache.set(fontName, st);
  return st;
}

// ---------------------------------------------------------------- operadores da página

function hexOf(args) {
  const a = args && args[0];
  if (typeof a === 'string') return a.replace('#', '').slice(0, 6).toUpperCase();
  const src = args && args.length >= 3 && typeof args[0] === 'number' ? args : a && a.length >= 3 ? a : null;
  if (!src) return null;
  return [0, 1, 2].map((k) => Math.max(0, Math.min(255, Math.round(src[k]))).toString(16).padStart(2, '0')).join('').toUpperCase();
}

// Percorre a lista de operadores: cor e visibilidade de cada caractere
// desenhado, na ordem do conteúdo, e as imagens com a posição na página.
function walkOperators(pdfjs, opList) {
  const OPS = pdfjs.OPS;
  const { fnArray, argsArray } = opList;
  let ctm = [1, 0, 0, 1, 0, 0];
  let fill = '000000';
  let stroke = '000000';
  let mode = 0;
  const stack = [];
  const chars = [];
  const images = [];
  let paths = 0;
  const fills = [];
  const segs = [];
  const fillOps = new Set([OPS.fill, OPS.eoFill, OPS.fillStroke, OPS.eoFillStroke, OPS.closeFillStroke, OPS.closeEOFillStroke].filter((x) => x != null));
  const strokeOps = new Set([OPS.stroke, OPS.closeStroke, OPS.fillStroke, OPS.eoFillStroke, OPS.closeFillStroke, OPS.closeEOFillStroke].filter((x) => x != null));
  const textOps = new Set([OPS.showText, OPS.showSpacedText, OPS.nextLineShowText, OPS.nextLineSetSpacingShowText].filter((x) => x != null));
  for (let i = 0; i < fnArray.length; i++) {
    const fn = fnArray[i];
    const args = argsArray[i];
    if (fn === OPS.save || fn === OPS.paintFormXObjectBegin) {
      stack.push({ ctm, fill, stroke, mode });
      if (fn === OPS.paintFormXObjectBegin && args && args[0] && args[0].length === 6) ctm = mul(ctm, args[0]);
    } else if (fn === OPS.restore || fn === OPS.paintFormXObjectEnd) {
      const s = stack.pop();
      if (s) ({ ctm, fill, stroke, mode } = s);
    } else if (fn === OPS.transform) {
      ctm = mul(ctm, args);
    } else if (fn === OPS.setFillRGBColor) {
      fill = hexOf(args) || fill;
    } else if (fn === OPS.setStrokeRGBColor) {
      stroke = hexOf(args) || stroke;
    } else if (fn === OPS.setTextRenderingMode) {
      mode = args[0];
    } else if (textOps.has(fn)) {
      const glyphs = (args || []).find((x) => Array.isArray(x)) || [];
      const hidden = mode === 3 || mode === 7;
      for (const g of glyphs) {
        if (!g || typeof g !== 'object' || !g.unicode) continue;
        for (const ch of g.unicode.normalize('NFKC')) if (!/\s/.test(ch)) chars.push({ ch, color: fill, hidden });
      }
    } else if (fn === OPS.paintImageXObject || fn === OPS.paintInlineImageXObject || fn === OPS.paintImageXObjectRepeat) {
      images.push({ ref: args[0], inline: fn === OPS.paintInlineImageXObject, ctm });
    } else if (fn === OPS.constructPath) {
      paths++;
      // retângulos preenchidos: faixas atrás de títulos e células
      const mm = args && args[2];
      if (fillOps.has(args[0]) && mm && mm.length === 4 && fill && fill !== 'FFFFFF') fills.push({ mm: [mm[0], mm[1], mm[2], mm[3]], ctm, color: fill });
      // trechos retos do caminho: bordas de tabela
      const data = args && args[1] && args[1][0];
      const stroked = strokeOps.has(args[0]) && stroke !== 'FFFFFF';
      const filled = fillOps.has(args[0]) && fill !== 'FFFFFF';
      if (data && data.length && typeof data[0] === 'number' && (stroked || filled)) {
        // subcaminhos com os pontos já na página
        const subs = [];
        let cur = null;
        for (let k = 0; k < data.length;) {
          const op = data[k++];
          if (op === 0) { cur = [apply(ctm, data[k], data[k + 1])]; subs.push(cur); k += 2; }
          else if (op === 1) { if (cur) cur.push(apply(ctm, data[k], data[k + 1])); k += 2; }
          else if (op === 2) { if (cur) cur.push(null); k += 6; }
          else if (op === 3) { if (cur) cur.push(null); k += 4; }
          else if (op === 4) { if (cur && cur.length) cur.closed = true; }
          else break;
        }
        for (const pts of subs) {
          if (pts.includes(null)) continue;
          if (stroked) {
            const ring = pts.closed ? [...pts, pts[0]] : pts;
            for (let k = 1; k < ring.length; k++) segs.push({ a: ring[k - 1], b: ring[k], color: stroke });
          }
          if (filled && pts.length >= 2) {
            const xs = pts.map((q) => q[0]);
            const ys = pts.map((q) => q[1]);
            const x0 = Math.min(...xs), x1 = Math.max(...xs), y0 = Math.min(...ys), y1 = Math.max(...ys);
            // retângulo fino desenhado como linha
            if (y1 - y0 < 2.5 && x1 - x0 > 3) segs.push({ a: [x0, (y0 + y1) / 2], b: [x1, (y0 + y1) / 2], color: fill });
            else if (x1 - x0 < 2.5 && y1 - y0 > 3) segs.push({ a: [(x0 + x1) / 2, y0], b: [(x0 + x1) / 2, y1], color: fill });
          }
        }
      }
    }
  }
  return { chars, images, paths, fills, segs };
}

// Casa os itens de texto com os caracteres desenhados para saber a cor
function colorItems(items, chars) {
  let gi = 0;
  for (const it of items) {
    const s = it.str.normalize('NFKC').replace(/\s+/g, '');
    if (!s) continue;
    let at = -1;
    const lim = Math.min(chars.length - s.length, gi + 400);
    for (let k = gi; k <= lim; k++) {
      let ok = true;
      for (let j = 0; j < s.length && ok; j++) if (chars[k + j].ch !== s[j]) ok = false;
      if (ok) { at = k; break; }
    }
    if (at < 0) continue;
    it.color = chars[at].color;
    it.hidden = chars[at].hidden;
    gi = at + s.length;
  }
}

// ---------------------------------------------------------------- imagens

function objFrom(page, ref) {
  const store = String(ref).startsWith('g_') ? page.commonObjs : page.objs;
  return new Promise((resolve) => {
    const t = setTimeout(() => resolve(null), 8000);
    try { store.get(ref, (v) => { clearTimeout(t); resolve(v); }); } catch { clearTimeout(t); resolve(null); }
  });
}

async function canvasBytes(canvas, opaque) {
  const jpeg = opaque && canvas.width * canvas.height > 250000;
  const blob = await new Promise((r) => canvas.toBlob(r, jpeg ? 'image/jpeg' : 'image/png', 0.9));
  if (!blob) return null;
  return { bytes: new Uint8Array(await blob.arrayBuffer()), ext: jpeg ? 'jpg' : 'png', type: jpeg ? 'image/jpeg' : 'image/png' };
}

async function imageToBytes(pdfjs, img) {
  if (!img || !img.width || !img.height) return null;
  const w = img.width;
  const h = img.height;
  const k = Math.min(1, 4000 / Math.max(w, h));
  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.round(w * k));
  canvas.height = Math.max(1, Math.round(h * k));
  const g = canvas.getContext('2d');
  let opaque = true;
  if (img.bitmap) {
    g.drawImage(img.bitmap, 0, 0, canvas.width, canvas.height);
    opaque = false;
  } else if (img.data) {
    const src = img.data;
    const rgba = new Uint8ClampedArray(w * h * 4);
    const kind = img.kind;
    const K = pdfjs.ImageKind || { GRAYSCALE_1BPP: 1, RGB_24BPP: 2, RGBA_32BPP: 3 };
    if (kind === K.RGBA_32BPP) { rgba.set(src.subarray ? src.subarray(0, rgba.length) : src.slice(0, rgba.length)); opaque = false; }
    else if (kind === K.RGB_24BPP) {
      for (let i = 0, j = 0; j < rgba.length; i += 3, j += 4) { rgba[j] = src[i]; rgba[j + 1] = src[i + 1]; rgba[j + 2] = src[i + 2]; rgba[j + 3] = 255; }
    } else if (kind === K.GRAYSCALE_1BPP) {
      const stride = (w + 7) >> 3;
      for (let y = 0; y < h; y++) {
        for (let x = 0; x < w; x++) {
          const on = src[y * stride + (x >> 3)] & (128 >> (x & 7));
          const j = (y * w + x) * 4;
          const v = on ? 255 : 0;
          rgba[j] = v; rgba[j + 1] = v; rgba[j + 2] = v; rgba[j + 3] = 255;
        }
      }
    } else return null;
    const tmp = document.createElement('canvas');
    tmp.width = w;
    tmp.height = h;
    tmp.getContext('2d').putImageData(new ImageData(rgba, w, h), 0, 0);
    g.drawImage(tmp, 0, 0, canvas.width, canvas.height);
  } else return null;
  return canvasBytes(canvas, opaque);
}

async function renderPage(page, scale) {
  const vp = page.getViewport({ scale });
  const canvas = document.createElement('canvas');
  canvas.width = Math.ceil(vp.width);
  canvas.height = Math.ceil(vp.height);
  const g = canvas.getContext('2d');
  g.fillStyle = '#fff';
  g.fillRect(0, 0, canvas.width, canvas.height);
  await page.render({ canvas, canvasContext: g, viewport: vp }).promise;
  return canvasBytes(canvas, true);
}

// ---------------------------------------------------------------- linhas

function buildLines(items, pageW) {
  let lines = [];
  let cur = null;
  for (const it of items) {
    const size = Math.max(cur ? cur.size : 0, it.size);
    const dy = cur ? Math.abs(it.y - cur.y) : Infinity;
    // na mesma linha de base vale qualquer distância, como as células de uma tabela
    const same = cur && it.x > cur.right - it.size * 0.8 && (dy < Math.min(cur.size, it.size) * 0.25 || (dy < size * 0.5 && it.x - cur.right < size * 14));
    if (same) {
      cur.items.push(it);
      cur.right = Math.max(cur.right, it.x + it.w);
    } else {
      cur = { items: [it], x: it.x, right: it.x + it.w, y: it.y, size: it.size };
      lines.push(cur);
    }
  }
  lines.forEach(finishLine);
  // células de tabela desenhadas fora de ordem: pedaços curtos na mesma linha de base viram uma linha só
  const joined = [];
  for (const l of lines) {
    l.frag = l.right - l.x;
    const m = joined.find((o) => Math.abs(o.y - l.y) < Math.min(o.size, l.size) * 0.25 && (l.x >= o.right - 1 || l.right <= o.x + 1)
      && Math.max(o.frag, l.frag) < pageW * 0.3);
    if (m) {
      m.items.push(...l.items);
      m.items.sort((a, b) => a.x - b.x);
      m.frag = Math.max(m.frag, l.frag);
      finishLine(m);
    } else joined.push(l);
  }
  lines = joined;
  // PDF que desenha o texto fora da ordem de leitura: ordena de cima para baixo
  let jumps = 0;
  for (let i = 1; i < lines.length; i++) {
    const a = lines[i - 1];
    const b = lines[i];
    if (b.top < a.top - a.size * 2 && !(b.x > a.right - 1)) jumps++;
  }
  if (jumps > Math.max(2, lines.length * 0.15)) {
    lines.sort((a, b) => a.y - b.y || a.x - b.x);
    const merged = [];
    for (const l of lines) {
      const m = merged.find((o) => Math.abs(o.y - l.y) < Math.min(o.size, l.size) * 0.25 && (l.x >= o.right - 1 || l.right <= o.x + 1));
      if (m) { m.items.push(...l.items); m.items.sort((a, b) => a.x - b.x); m.right = Math.max(m.right, l.right); finishLine(m); } else merged.push(l);
    }
    lines = merged;
  }
  return lines.filter((l) => l.text.trim());
}

function finishLine(l) {
  {
    // tamanho e linha de base do texto principal, sem índices sobrescritos
    const weight = new Map();
    for (const it of l.items) weight.set(it.size, (weight.get(it.size) || 0) + it.str.length);
    l.size = [...weight.entries()].sort((a, b) => b[1] - a[1])[0][0];
    const main = l.items.find((it) => it.size === l.size) || l.items[0];
    l.y = main.y;
    l.x = Math.min(...l.items.map((it) => it.x));
    l.top = l.y - l.size * 0.8;
    l.bottom = l.y + l.size * 0.25;
    l.text = l.items.reduce((s, it, i) => {
      const prev = l.items[i - 1];
      const gap = prev ? it.x - (prev.x + prev.w) : 0;
      return s + (prev && gap > it.size * 0.15 && !/\s$/.test(s) && !/^\s/.test(it.str) ? ' ' : '') + it.str;
    }, '');
    const chars = l.items.reduce((n, it) => n + it.str.trim().length, 0) || 1;
    l.bold = l.items.reduce((n, it) => n + (it.font.bold ? it.str.trim().length : 0), 0) / chars > 0.6;
    l.right = Math.max(...l.items.map((it) => it.x + it.w));
  }
}

// borda direita da coluna de cada linha: onde terminam as linhas cheias vizinhas
function columnRights(lines) {
  for (const l of lines) {
    const counts = new Map();
    for (const o of lines) {
      if (o.x >= l.right || o.right <= l.x) continue;
      const k = Math.round(o.right / 2) * 2;
      counts.set(k, (counts.get(k) || 0) + 1);
    }
    let best = l.right;
    for (const [k, c] of counts) if (c >= 2 && k > best) best = k;
    l.colRight = Math.max(best, l.right);
  }
}

function firstWordWidth(line) {
  const it = line.items.find((x) => x.str.trim());
  if (!it) return 0;
  const word = it.str.trimStart().split(/\s/)[0];
  return (it.w * word.length) / Math.max(1, it.str.length);
}

// ---------------------------------------------------------------- parágrafos

function startsList(line) {
  const t = line.text.trimStart();
  return BULLET.test(t) || NUMBERED.test(t);
}

function continues(para, prev, line) {
  if (line.page !== prev.page) return false;
  const size = Math.max(prev.size, line.size);
  const dy = line.y - prev.y;
  if (dy <= 0 || dy > size * 1.75) return false;
  if (para.lines.length >= 2 && dy > para.pitch * 1.3 + 0.5) return false;
  if (Math.abs(line.size - prev.size) > prev.size * 0.15) return false;
  if (prev.bold !== line.bold) return false;
  if (startsList(line)) return false;
  // a linha anterior terminou curta: a primeira palavra desta cabia nela
  if (prev.right + firstWordWidth(line) + prev.size * 0.3 < prev.colRight) return false;
  if (para.lines.length === 1) return line.x > prev.x - size * 6 && line.x < prev.x + size * 3;
  return Math.abs(line.x - para.lines[1].x) < size;
}

function buildParagraphs(lines) {
  const paras = [];
  let p = null;
  for (const line of lines) {
    const prev = p && p.lines[p.lines.length - 1];
    if (p && continues(p, prev, line)) {
      p.lines.push(line);
      if (p.lines.length === 2) p.pitch = line.y - prev.y;
      else p.pitch = median(p.lines.slice(1).map((l, i) => l.y - p.lines[i].y));
    } else {
      p = { lines: [line], pitch: line.size * 1.2, page: line.page };
      paras.push(p);
    }
  }
  for (const q of paras) finishPara(q);
  return paras;
}

function finishPara(p) {
  const first = p.lines[0];
  const last = p.lines[p.lines.length - 1];
  p.firstX = first.x;
  p.x = p.lines.length > 1 ? Math.min(...p.lines.slice(1).map((l) => l.x)) : first.x;
  p.right = Math.max(...p.lines.map((l) => l.right));
  p.colRight = Math.max(...p.lines.map((l) => l.colRight));
  p.top = first.top;
  p.bottom = last.bottom;
  p.size = first.size;
  p.bold = p.lines.every((l) => l.bold);
  p.text = p.lines.map((l) => l.text).join(' ').replace(/\s+/g, ' ').trim();
}

// ---------------------------------------------------------------- leitura

export async function readPdf(bytes, opts = {}) {
  const pdfjs = await loadPdfJs();
  const base = new URL('../dist/pdfjs/', location.href).href;
  const task = pdfjs.getDocument({
    data: bytes.slice(), cMapUrl: base + 'cmaps/', cMapPacked: true, standardFontDataUrl: base + 'standard_fonts/',
    isEvalSupported: false, useSystemFonts: true, fontExtraProperties: true,
  });
  task.onPassword = async (update, reason) => {
    const pw = opts.askPassword ? await opts.askPassword(reason === (pdfjs.PasswordResponses || {}).INCORRECT_PASSWORD) : null;
    if (pw == null) { task.destroy(); return; }
    update(pw);
  };
  let pdf;
  try {
    pdf = await task.promise;
  } catch (err) {
    if (err && err.name === 'PasswordException') throw new Error('O PDF é protegido por senha.');
    if (err && err.name === 'InvalidPDFException') throw new Error('O arquivo PDF está danificado ou incompleto.');
    throw err;
  }
  try {
    return await convert(pdfjs, pdf, opts);
  } finally {
    task.destroy();
  }
}

async function convert(pdfjs, pdf, opts) {
  const { model } = await blankModel();
  model.format = 'pdf';
  const pages = [];
  const fontCache = new Map();
  const known = new Set((opts.fonts || []).map((f) => f.toLowerCase()));
  for (let n = 1; n <= pdf.numPages; n++) {
    if (opts.onProgress) opts.onProgress(n, pdf.numPages);
    const page = await pdf.getPage(n);
    const vp = page.getViewport({ scale: 1 });
    const opList = await page.getOperatorList();
    const tc = await page.getTextContent();
    const ops = walkOperators(pdfjs, opList);
    const links = [];
    try {
      for (const a of await page.getAnnotations({ intent: 'display' })) {
        const href = a.subtype === 'Link' && (a.url || a.unsafeUrl);
        if (!href || !/^(https?:|mailto:|ftp:)/i.test(href)) continue;
        const [x1, y1] = apply(vp.transform, a.rect[0], a.rect[1]);
        const [x2, y2] = apply(vp.transform, a.rect[2], a.rect[3]);
        links.push({ x1: Math.min(x1, x2), x2: Math.max(x1, x2), y1: Math.min(y1, y2), y2: Math.max(y1, y2), href });
      }
    } catch { /* sem anotações */ }
    const items = [];
    for (const raw of tc.items) {
      if (!raw.str) continue;
      const tx = mul(vp.transform, raw.transform);
      const size = Math.hypot(tx[2], tx[3]);
      if (!size || Math.abs(tx[1]) > size * 0.05) continue;   // texto girado, como marca d'água
      const it = { str: raw.str, x: tx[4], y: tx[5], size: Math.round(size * 2) / 2, w: raw.width, font: fontStyle(page, raw.fontName, tc.styles, fontCache, known), page: n };
      const cx = it.x + it.w / 2;
      const cy = it.y - size * 0.3;
      const link = links.find((l) => cx >= l.x1 && cx <= l.x2 && cy >= l.y1 - 2 && cy <= l.y2 + 2);
      if (link) it.href = link.href;
      items.push(it);
    }
    colorItems(items, ops.chars);
    const visible = items.filter((it) => !it.hidden);
    const hiddenChars = items.length - visible.length;
    // imagens com a posição na página
    const imgs = [];
    for (const im of ops.images) {
      const m = mul(vp.transform, im.ctm);
      const pts = [[0, 0], [1, 0], [0, 1], [1, 1]].map(([x, y]) => apply(m, x, y));
      const xs = pts.map((q) => q[0]);
      const ys = pts.map((q) => q[1]);
      const box = { x: Math.min(...xs), y: Math.min(...ys), w: Math.max(...xs) - Math.min(...xs), h: Math.max(...ys) - Math.min(...ys) };
      if (box.w < 6 || box.h < 6) continue;
      imgs.push({ ...box, ref: im.ref, inline: im.inline });
    }
    // PDF digitalizado com texto reconhecido por baixo: fica o texto
    const scanned = hiddenChars > items.length / 2 && imgs.some((b) => b.w * b.h > vp.width * vp.height * 0.5);
    const useItems = scanned ? items : visible;
    const pageImgs = scanned ? imgs.filter((b) => b.w * b.h <= vp.width * vp.height * 0.5) : imgs;
    // slide: página deitada com fundo cobrindo tudo vira uma imagem fiel da página
    const pageArea = vp.width * vp.height;
    const bigFill = ops.fills.some((f) => {
      const m = mul(vp.transform, f.ctm);
      const [x1, y1] = apply(m, f.mm[0], f.mm[1]);
      const [x2, y2] = apply(m, f.mm[2], f.mm[3]);
      return Math.abs((x2 - x1) * (y2 - y1)) > pageArea * 0.8;
    });
    if (!scanned && vp.width > vp.height * 1.15 && (bigFill || imgs.some((b) => b.w * b.h > pageArea * 0.8))) {
      const shot = await renderPage(page, 1.6).catch(() => null);
      if (shot) {
        pages.push({ n, w: vp.width, h: vp.height, lines: [], imgs: [], tables: [], fills: [], render: shot, slide: true });
        page.cleanup();
        continue;
      }
    }
    for (const b of pageImgs) {
      const obj = b.inline ? b.ref : await objFrom(page, b.ref);
      b.data = await imageToBytes(pdfjs, obj).catch(() => null);
    }
    let render = null;
    const textChars = useItems.reduce((s, it) => s + it.str.trim().length, 0);
    if (textChars < 20 && !pageImgs.some((b) => b.data) && ops.paths > 5) render = await renderPage(page, 2).catch(() => null);
    const fills = [];
    const cellFills = [];
    for (const f of ops.fills) {
      const m = mul(vp.transform, f.ctm);
      const pts = [[f.mm[0], f.mm[1]], [f.mm[2], f.mm[1]], [f.mm[0], f.mm[3]], [f.mm[2], f.mm[3]]].map(([x, y]) => apply(m, x, y));
      const xs = pts.map((q) => q[0]);
      const ys = pts.map((q) => q[1]);
      const r = { x: Math.min(...xs), y: Math.min(...ys), right: Math.max(...xs), bottom: Math.max(...ys), color: f.color };
      if (r.right - r.x > 20 && r.bottom - r.y > 4 && r.bottom - r.y < 90) fills.push(r);
      if (r.right - r.x > 4 && r.bottom - r.y > 4 && (r.right - r.x) * (r.bottom - r.y) < vp.width * vp.height * 0.5) cellFills.push(r);
    }
    // tabelas desenhadas com linhas: o texto de dentro vai para as células
    const segs = ops.segs.map((sg) => {
      const [x1, y1] = apply(vp.transform, sg.a[0], sg.a[1]);
      const [x2, y2] = apply(vp.transform, sg.b[0], sg.b[1]);
      return { x1, y1, x2, y2, color: sg.color };
    });
    const tables = findTables(segs, useItems, cellFills, vp.height);
    for (const t of tables) t.page = n;
    const inTable = new Set(tables.flatMap((t) => t.cells.flatMap((c) => c.items)));
    const lines = buildLines(useItems.filter((it) => !inTable.has(it)), vp.width);
    for (const l of lines) l.page = n;
    pages.push({ n, w: vp.width, h: vp.height, lines, imgs: pageImgs.filter((b) => b.data), render, fills, tables });
    page.cleanup();
  }
  const prev = R.model;
  R.model = model;
  try {
    return layoutDocument(model, pages);
  } finally {
    R.model = prev;
  }
}

// ---------------------------------------------------------------- montagem

function layoutDocument(model, pages) {
  // tamanho de página mais comum: a capa às vezes é diferente
  const sizes = new Map();
  for (const p of pages) { const k = Math.round(p.w) + 'x' + Math.round(p.h); sizes.set(k, (sizes.get(k) || 0) + 1); }
  const common = [...sizes.entries()].sort((a, b) => b[1] - a[1])[0];
  const first = (common && pages.find((p) => Math.round(p.w) + 'x' + Math.round(p.h) === common[0])) || pages[0] || { w: 612, h: 792, lines: [], imgs: [] };
  const hf = extractHeaderFooter(pages);
  const allLines = pages.flatMap((p) => p.lines);
  // tamanho e fonte do texto principal, contando o que está nas tabelas
  const sizeW = new Map();
  const fontW = new Map();
  const tableItems = pages.flatMap((p) => p.tables.flatMap((t) => t.cells.flatMap((c) => c.items)));
  for (const l of [...allLines, { items: tableItems }]) {
    for (const it of l.items) {
      const n = it.str.trim().length;
      sizeW.set(it.size, (sizeW.get(it.size) || 0) + n);
      if (it.font.family) fontW.set(it.font.family, (fontW.get(it.font.family) || 0) + n);
    }
  }
  const top = (m) => [...m.entries()].sort((a, b) => b[1] - a[1])[0];
  const bodySize = (top(sizeW) || [11])[0];
  const bodyFont = (top(fontW) || ['Calibri'])[0];
  // margens a partir do texto
  const longLines = allLines.filter((l) => l.text.trim().length >= 8);
  const ref = longLines.length ? longLines : allLines;
  const left = ref.length ? Math.min(...ref.map((l) => l.x)) : 72;
  const right = ref.length ? first.w - Math.max(...ref.map((l) => l.right)) : 72;
  const topM = allLines.length ? Math.min(...allLines.map((l) => l.top)) : 72;
  const bottomM = allLines.length ? first.h - Math.max(...allLines.map((l) => l.bottom)) : 72;
  const mar = {
    left: Math.max(18, Math.min(left, first.w / 3)), right: Math.max(18, Math.min(right, first.w / 3)),
    top: Math.max(18, Math.min(topM, first.h / 3)), bottom: Math.max(18, Math.min(bottomM, first.h / 3)),
  };
  // apresentação: só imagens de slide, com margem pequena para ficarem grandes
  if (pages.length && pages.every((p) => p.slide)) { mar.left = 18; mar.right = 18; mar.top = 18; mar.bottom = 18; }
  model.section.setPageSize(first.w * 20, first.h * 20, first.w > first.h ? 'landscape' : null);
  model.section.setMargins({
    top: mar.top * 20, bottom: mar.bottom * 20, left: mar.left * 20, right: mar.right * 20,
    header: Math.max(18, hf.headerTop != null ? hf.headerTop : mar.top / 2) * 20, footer: Math.max(18, hf.footerBottom != null ? hf.footerBottom : mar.bottom / 2) * 20,
  });
  model.styles.setDefaultFont(bodyFont || 'Calibri', Math.round(bodySize * 2));
  model.styles.setDefaultParagraphSpacing(0, 0, 240);
  const ctx = { model, bodySize, bodyFont, mar, pageW: first.w, pageH: first.h, bulletNum: null, bulletLevels: [] };

  // parágrafos de cada página, emendando o que continua na página seguinte
  const paras = [];
  for (const pg of pages) {
    columnRights(pg.lines);
    const ps = buildParagraphs(pg.lines);
    // sem colunas lado a lado, a ordem é de cima para baixo, mesmo que o PDF desenhe fora de ordem
    const sideBySide = ps.some((a, i) => ps.some((b, j) => j > i && (a.right <= b.firstX || b.right <= a.firstX) && a.top < b.bottom && b.top < a.bottom));
    if (!sideBySide) ps.sort((a, b) => a.top - b.top);
    for (const q of ps) {
      const minX = Math.min(q.x, q.firstX);
      const f = pg.fills.find((r) => r.x <= minX + 2 && r.right >= q.right - 2 && r.y <= q.top + q.size * 0.3 && r.bottom >= q.bottom - q.size * 0.3);
      if (f) q.shd = f.color;
    }
    const last = paras[paras.length - 1];
    let pageBreak = false;
    if (last && ps.length && !pg.imgs.some((b) => b.y < ps[0].top) && !pg.tables.some((t) => t.y < ps[0].top) && joinsAcrossPages(last, ps[0], ctx)) {
      last.lines.push(...ps[0].lines);
      finishPara(last);
      ps.shift();
    } else if (last) {
      // página anterior terminada bem antes do fim: quebra de página intencional
      const prevPage = pages[pg.n - 2];
      const usedBottom = prevPage ? Math.max(0, ...prevPage.lines.map((l) => l.bottom), ...prevPage.imgs.map((b) => b.y + b.h), ...prevPage.tables.map((t) => t.bottom)) : 0;
      pageBreak = !!prevPage && usedBottom < prevPage.h - mar.bottom - (prevPage.h - mar.top - mar.bottom) * 0.25;
    }
    for (const q of ps) q.pageN = pg.n;
    // imagens entram antes do primeiro parágrafo que começa abaixo delas
    const blocks = ps.map((q) => ({ kind: 'para', top: q.top, para: q }));
    for (const b of pg.imgs) {
      const i = blocks.findIndex((x) => x.kind === 'para' && x.top > b.y + 1);
      const blk = { kind: 'img', top: b.y, img: b };
      if (i < 0) blocks.push(blk); else blocks.splice(i, 0, blk);
    }
    for (const t of pg.tables) {
      const i = blocks.findIndex((x) => x.top > t.y + 1);
      const blk = { kind: 'table', top: t.y, table: t };
      if (i < 0) blocks.push(blk); else blocks.splice(i, 0, blk);
    }
    // tabela que continua da página anterior, com as mesmas colunas
    const prevBlk = paras[paras.length - 1];
    if (prevBlk && prevBlk.kind === 'table' && blocks[0] && blocks[0].kind === 'table' && sameColumns(prevBlk.table, blocks[0].table)) {
      const t = prevBlk.table;
      (t.more = t.more || []).push(blocks[0].table);
      blocks.shift();
      pageBreak = false;
    }
    if (pg.render) blocks.push({ kind: 'img', top: 0, img: { x: 0, y: 0, w: pg.w, h: pg.h, data: pg.render, full: true } });
    if (pageBreak && blocks.length) {
      if (blocks[0].kind === 'para') blocks[0].para.breakBefore = true; else blocks[0].breakBefore = true;
    }
    paras.push(...blocks.map((b) => (b.kind === 'para' ? b.para : b)));
  }

  detectHeadings(paras, ctx);
  const out = [];
  let prevBottom = null;
  let prevPage = null;
  for (const p of paras) {
    if (p.kind === 'table') {
      const node = tableNode(p.table, ctx);
      if (node) {
        // duas tabelas coladas viram uma só no Word: um parágrafo vazio separa
        if (p.breakBefore) out.push(nodes.paragraph.create({ pageBreakBefore: true }));
        else if (out.length && out[out.length - 1].type === nodes.table) out.push(nodes.paragraph.create());
        out.push(node);
      }
      prevBottom = null;
      prevPage = null;
      continue;
    }
    if (p.kind === 'img') {
      out.push(imageParagraph(p.img, ctx, p.breakBefore));
      prevBottom = p.img.y + p.img.h;
      prevPage = null;
      continue;
    }
    const gap = prevPage === p.pageN && prevBottom != null ? p.lines[0].y - prevBottom : null;
    out.push(paragraphNode(p, ctx, gap));
    prevBottom = p.lines[p.lines.length - 1].y;
    prevPage = p.pageN;
  }
  if (!out.length) out.push(nodes.paragraph.create());
  const made = [];
  for (const part of ['header', 'footer']) {
    const lines = hf[part] || [];
    const imgs = hf[part + 'Imgs'] || [];
    const tbls = hf[part + 'Tables'] || [];
    if (!lines.length && !imgs.length && !tbls.length) continue;
    const entry = newHeaderFooter(model, part);
    const parts = [...lines.map((l) => ({ top: l.top, line: l })), ...imgs.map((b) => ({ top: b.y, img: b })), ...tbls.map((t) => ({ top: t.y, table: t }))].sort((a, b) => a.top - b.top);
    // elementos lado a lado, como logotipos numa faixa, dividem um parágrafo com tabulações
    const span = (x) => (x.img ? [x.img.x, x.img.x + x.img.w, x.img.y, x.img.y + x.img.h] : [x.line.x, x.line.right, x.line.top, x.line.bottom]);
    const rows = [];
    for (const x of parts) {
      if (x.table) { rows.push({ table: x.table }); continue; }
      const [x0, x1, y0, y1] = span(x);
      const row = rows.find((r) => r.items && y0 < r.bottom - 2 && y1 > r.top + 2 && r.items.every((o) => { const [a0, a1] = span(o); return x1 <= a0 + 1 || x0 >= a1 - 1; }));
      if (row) { row.items.push(x); row.top = Math.min(row.top, y0); row.bottom = Math.max(row.bottom, y1); } else rows.push({ items: [x], top: y0, bottom: y1 });
    }
    const blocks = [];
    for (const r of rows) {
      if (r.table) {
        const node = tableNode(r.table, { ...ctx, hfPage: r.table.page, numPages: pages.length, numState: { roles: r.table.numRoles || [], i: 0 } }, true);
        if (node) { if (blocks.length && blocks[blocks.length - 1].type === nodes.table) blocks.push(nodes.paragraph.create()); blocks.push(node); }
      } else if (r.items.length === 1) blocks.push(hfBlock(r.items[0]));
      else blocks.push(hfRow(r.items));
    }
    function hfRow(items) {
      items.sort((a, b) => span(a)[0] - span(b)[0]);
      const content = [];
      const tabs = [];
      items.forEach((x, k) => {
        const x0 = span(x)[0];
        if (k > 0) { content.push(nodes.tab.create()); tabs.push({ val: 'left', pos: twip(x0 - ctx.mar.left), leader: null }); }
        const para = hfBlock(x);
        para.content.forEach((c) => content.push(c));
      });
      const firstX = span(items[0])[0] - ctx.mar.left;
      const attrs = { tabs };
      if (firstX > 3) attrs.ind = { left: twip(firstX) };
      return nodes.paragraph.create(attrs, content);
    }
    // o cabeçalho precisa terminar num parágrafo
    if (!blocks.length || blocks[blocks.length - 1].type !== nodes.paragraph) blocks.push(nodes.paragraph.create({ spacing: { before: '0', after: '0', line: '20', lineRule: 'exact' } }));
    function hfBlock(x) {
      if (x.img) return imageParagraph(x.img, ctx);
      const p = { lines: [x.line], pageN: x.line.page };
      columnRights([x.line]);
      finishPara(p);
      return paragraphNode(p, { ...ctx, hfPage: x.line.page, numPages: pages.length, numState: { roles: x.line.numRoles || [], i: 0 } }, null, true);
    }
    entry.doc = schema.node('doc', null, blocks);
    entry.dirty = true;
    model.section.setReference(part, 'default', entry.rId);
    made.push(part);
  }
  if (hf.titlePg && made.length) {
    // a primeira página fica com cabeçalho e rodapé vazios
    model.section.setTitlePg(true);
    for (const part of made) {
      const entry = newHeaderFooter(model, part);
      entry.doc = schema.node('doc', null, [nodes.paragraph.create()]);
      entry.dirty = true;
      model.section.setReference(part, 'first', entry.rId);
    }
  }
  return { model, doc: finalizeDoc(schema.node('doc', null, out), model) };
}

function joinsAcrossPages(last, next, ctx) {
  if (!last.lines || last.kind === 'img') return false;
  const a = last.lines[last.lines.length - 1];
  const b = next.lines[0];
  if (Math.abs(a.size - b.size) > a.size * 0.15 || a.bold !== b.bold || startsList(b)) return false;
  if (a.right + firstWordWidth(b) + a.size * 0.3 < a.colRight) return false;
  if (/[.!?:]["”)]?$/.test(a.text.trim())) return false;
  return Math.abs(b.x - last.x) < a.size * 2 || Math.abs(b.x - ctx.mar.left) < a.size;
}

// Linhas no alto ou no pé que se repetem em pelo menos metade das páginas
function extractHeaderFooter(pages) {
  const res = { header: null, footer: null, headerTop: null, footerBottom: null };
  if (pages.length < 2) return res;
  const key = (l) => l.text.replace(/\d+/g, '#').replace(/\s+/g, ' ').trim();
  const zones = { header: (l, p) => l.bottom < p.h * 0.14, footer: (l, p) => l.top > p.h * 0.88 };
  for (const part of ['header', 'footer']) {
    const found = new Map();
    for (const p of pages) {
      const seen = new Map();
      for (const l of p.lines) if (zones[part](l, p) && l.text.trim()) seen.set(key(l), (l.text.match(/\d+/g) || []).map(Number));
      for (const [k, nums] of seen) { if (!found.has(k)) found.set(k, []); found.get(k).push({ page: p.n, nums }); }
    }
    const need = Math.max(2, Math.ceil(pages.length / 2));
    // números iguais em todas as páginas, ou que andam junto com a página
    const tracksPage = (list) => list[0].nums.map((v) => v - list[0].page).some((d) => list.every((x) => x.nums.includes(x.page + d)));
    const stable = (list) => list.every((x) => x.nums.join() === list[0].nums.join()) || tracksPage(list);
    const keys = new Set([...found.entries()].filter(([, list]) => list.length >= need && stable(list)).map(([k]) => k));
    if (!keys.size) continue;
    for (const p of pages) {
      const kept = [];
      const taken = [];
      for (const l of p.lines) (zones[part](l, p) && keys.has(key(l)) ? taken : kept).push(l);
      p.lines = kept;
      if (!res[part] && taken.length) {
        for (const l of taken) l.numRoles = numberRoles(found.get(key(l)), pages.length);
        res[part] = taken;
      }
    }
    const ls = res[part];
    if (part === 'header') res.headerTop = Math.min(...ls.map((l) => l.top));
    else res.footerBottom = pages[0].h - Math.max(...ls.map((l) => l.bottom));
    // a capa sem cabeçalho: primeira página diferente
    if (ls[0].page !== 1) res.titlePg = true;
  }
  // imagens repetidas no alto e no pé, como logotipos
  const imgZone = { header: (b, p) => b.y + b.h < p.h * 0.16, footer: (b, p) => b.y > p.h * 0.84 };
  const boxKey = (b) => [b.x, b.y, b.w, b.h].map((v) => Math.round(v / 3)).join(',');
  for (const part of ['header', 'footer']) {
    const count = new Map();
    for (const p of pages) {
      const seen = new Set();
      for (const b of p.imgs) if (imgZone[part](b, p)) seen.add(boxKey(b));
      for (const k of seen) count.set(k, (count.get(k) || 0) + 1);
    }
    const need = Math.max(2, Math.ceil(pages.length / 2));
    const keys = new Set([...count.entries()].filter(([, c]) => c >= need).map(([k]) => k));
    if (!keys.size) continue;
    const imgs = [];
    for (const p of pages) {
      const taken = p.imgs.filter((b) => imgZone[part](b, p) && keys.has(boxKey(b)));
      p.imgs = p.imgs.filter((b) => !taken.includes(b));
      if (!imgs.length && taken.length) { imgs.push(...taken); if (p.n !== 1) res.titlePg = true; }
    }
    res[part + 'Imgs'] = imgs;
  }
  // tabela de identificação repetida no alto ou no pé de cada página
  const tblZone = { header: (t, p) => t.bottom < p.h * 0.2, footer: (t, p) => t.y > p.h * 0.8 };
  const tblKey = (t) => t.cells.map((c) => c.items.map((it) => it.str).join('')).join('|').replace(/\d+/g, '#').replace(/\s+/g, ' ');
  for (const part of ['header', 'footer']) {
    const count = new Map();
    for (const p of pages) {
      const seen = new Set();
      for (const t of p.tables) if (tblZone[part](t, p)) seen.add(tblKey(t));
      for (const k of seen) count.set(k, (count.get(k) || 0) + 1);
    }
    const need = Math.max(2, Math.ceil(pages.length / 2));
    const keys = new Set([...count.entries()].filter(([, c]) => c >= need).map(([k]) => k));
    if (!keys.size) continue;
    const tbls = [];
    const inst = new Map();
    for (const p of pages) {
      const taken = p.tables.filter((t) => tblZone[part](t, p) && keys.has(tblKey(t)));
      p.tables = p.tables.filter((t) => !taken.includes(t));
      for (const t of taken) { const k = tblKey(t); if (!inst.has(k)) inst.set(k, []); inst.get(k).push({ page: p.n, nums: tableNumbers(t) }); }
      if (!tbls.length && taken.length) { tbls.push(...taken); if (p.n !== 1) res.titlePg = true; }
    }
    for (const t of tbls) t.numRoles = numberRoles(inst.get(tblKey(t)), pages.length);
    // imagem por trás da tabela, como um selo: fica de fora
    const imgs = res[part + 'Imgs'];
    if (imgs) res[part + 'Imgs'] = imgs.filter((b) => !tbls.some((t) => b.x < t.right && b.x + b.w > t.x && b.y < t.bottom && b.y + b.h > t.y));
    res[part + 'Tables'] = tbls;
  }
  return res;
}

function detectHeadings(paras, ctx) {
  const body = ctx.bodySize;
  const big = new Set();
  for (const p of paras) if (p.lines && p.size >= body * 1.15 && p.text.length <= 200 && p.lines.length <= 3) big.add(p.size);
  const ranks = [...big].sort((a, b) => b - a);
  for (let i = 0; i < paras.length; i++) {
    const p = paras[i];
    if (!p.lines || !p.text) continue;
    if (big.has(p.size) && p.text.length <= 200 && p.lines.length <= 3) {
      p.outline = Math.min(5, ranks.indexOf(p.size));
    } else if (p.bold && Math.abs(p.size - body) < 0.6 && p.lines.length === 1 && p.text.length <= 90 && !/[.;,]$/.test(p.text)) {
      const next = paras[i + 1];
      if (next && (!next.lines || !next.bold)) p.outline = Math.min(5, ranks.length);
    }
  }
}

function bulletNumId(ctx) {
  if (!ctx.bulletNum) {
    const chars = ['•', 'o', '▪'];
    ctx.bulletNum = ctx.model.numbering.createFromLevels(Array.from({ length: 9 }, (_, i) => ({ numFmt: 'bullet', lvlText: chars[i % 3], left: 720 * (i + 1), hanging: 360 })));
  }
  return ctx.bulletNum;
}

function marksFor(it, ctx, lineSize) {
  const ms = [];
  if (it.href) ms.push(M.hyperlink.create({ href: it.href }));
  const fam = it.font.family;
  if (fam && fam !== ctx.bodyFont) ms.push(M.rFonts.create({ a: { ascii: fam, hAnsi: fam, cs: fam } }));
  if (it.font.bold) ms.push(M.b.create());
  if (it.font.italic) ms.push(M.i.create());
  if (it.color && it.color !== '000000') ms.push(M.color.create({ a: { val: it.color } }));
  const sup = it.size < lineSize * 0.85 && it.y < it.lineY - lineSize * 0.15;
  const size = sup ? lineSize : it.size;
  const hp = Math.round(size * 2);
  if (hp !== Math.round(ctx.bodySize * 2)) ms.push(M.sz.create({ val: hp }));
  if (sup) ms.push(M.vertAlign.create({ val: 'superscript' }));
  return ms;
}

// campos PAGE e NUMPAGES no lugar dos números do cabeçalho e do rodapé
function pageFields(text, marks, ctx) {
  const out = [];
  const re = /\d+/g;
  let last = 0;
  let m;
  while ((m = re.exec(text))) {
    const st = ctx.numState;
    const role = st ? st.roles[st.i++] : null;
    const instr = role === 'page' ? ' PAGE ' : role === 'total' ? ' NUMPAGES ' : null;
    if (!instr) continue;
    if (m.index > last) out.push(schema.text(text.slice(last, m.index), marks));
    out.push(nodes.field_char.create({ type: 'begin' }, null, marks), nodes.instr_text.create({ text: instr }, null, marks),
      nodes.field_char.create({ type: 'separate' }, null, marks), schema.text(m[0], marks), nodes.field_char.create({ type: 'end' }, null, marks));
    last = m.index + m[0].length;
  }
  if (last < text.length) out.push(schema.text(text.slice(last), marks));
  return out;
}

function paragraphNode(p, ctx, gap, isHF) {
  const content = [];
  const tabs = [];
  let pendingSpace = false;
  let bullet = null;
  p.lines.forEach((line, li) => {
    let prevEnd = null;
    let prevStr = '';
    line.items.forEach((it, ii) => {
      let str = it.str;
      if (li === 0 && ii === 0 && !isHF) {
        const t = str.trimStart();
        const lone = line.items.length > 1 && LONE_BULLET.test(t.trim());
        const mb = lone ? null : BULLET.exec(t);
        if (mb || lone) {
          const rest = lone ? '' : t.slice(mb[0].length);
          bullet = { x: it.x, textX: rest ? it.x + (it.w * (it.str.length - rest.length)) / it.str.length : null };
          str = rest;
          if (!str) return;
        }
      }
      if (bullet && bullet.textX == null && str.trim()) bullet.textX = it.x;
      if (li > 0 && ii === 0) {
        // emenda de linhas: hífen de quebra some, o resto ganha um espaço
        const lastNode = content[content.length - 1];
        const lastText = lastNode && lastNode.isText ? lastNode.text : '';
        if (/\p{L}-$/u.test(lastText) && /^\p{Ll}/u.test(str)) content[content.length - 1] = schema.text(lastText.slice(0, -1), lastNode.marks);
        else if (lastText && !/\s$/.test(lastText) && !/^\s/.test(str)) pendingSpace = true;
      } else if (prevEnd != null) {
        const gapX = it.x - prevEnd;
        if (gapX > Math.max(it.size, line.size) * 2.2) {
          content.push(nodes.tab.create());
          tabs.push({ val: 'left', pos: twip(it.x - ctx.mar.left), leader: null });
          str = str.replace(/^\s+/, '');
          pendingSpace = false;
        } else if (gapX > it.size * 0.15 && !/\s$/.test(prevStr) && !/^\s/.test(str)) pendingSpace = true;
      }
      it.lineY = line.y;
      const marks = marksFor(it, ctx, line.size);
      if (pendingSpace) { str = ' ' + str; pendingSpace = false; }
      if (str) {
        if (isHF && ctx.hfPage) content.push(...pageFields(str, marks, ctx));
        else content.push(schema.text(str, marks));
      }
      prevEnd = it.x + it.w;
      prevStr = it.str;
    });
  });
  // tira espaços sobrando no começo e no fim
  while (content.length && content[0].isText && !content[0].text.trim()) content.shift();
  if (content.length && content[0].isText) content[0] = schema.text(content[0].text.replace(/^\s+/, ''), content[0].marks);
  const lastI = content.length - 1;
  if (lastI >= 0 && content[lastI].isText) {
    const t = content[lastI].text.replace(/\s+$/, '');
    if (t) content[lastI] = schema.text(t, content[lastI].marks); else content.pop();
  }
  const attrs = {};
  const size = p.size;
  const colLeft = ctx.mar.left;
  const colWidth = ctx.pageW - ctx.mar.left - ctx.mar.right;
  const center = (p.firstX + p.right) / 2;
  const colCenter = colLeft + colWidth / 2;
  const narrow = p.right - Math.min(p.firstX, p.x) < colWidth * 0.8;
  if (narrow && Math.abs(center - colCenter) < size * 1.5 && p.firstX > colLeft + size * 2) attrs.jc = 'center';
  else if (narrow && Math.abs(p.right - (ctx.pageW - ctx.mar.right)) < size && p.firstX > colLeft + colWidth * 0.3) attrs.jc = 'right';
  else if (p.lines.length >= 3 && p.lines.slice(0, -1).every((l) => Math.abs(l.right - p.colRight) < size * 0.4)) attrs.jc = 'both';
  if (!attrs.jc || attrs.jc === 'both') {
    const ind = {};
    if (bullet && bullet.textX) {
      const textX = bullet.textX;
      if (textX - colLeft > 2) ind.left = twip(textX - colLeft);
      ind.hanging = twip(Math.max(0, textX - bullet.x));
    } else {
      const leftIn = p.x - colLeft;
      if (leftIn > 3) ind.left = twip(leftIn);
      const fl = p.firstX - p.x;
      if (fl > 3) ind.firstLine = twip(fl);
      else if (fl < -3) ind.hanging = twip(-fl);
    }
    const rightIn = ctx.pageW - ctx.mar.right - p.colRight;
    if (p.lines.length >= 2 && rightIn > size * 2) ind.right = twip(rightIn);
    if (Object.keys(ind).length) attrs.ind = ind;
  }
  if (bullet && bullet.textX && !isHF) {
    attrs.numId = bulletNumId(ctx);
    const lvls = ctx.bulletLevels;
    let lvl = lvls.findIndex((x) => Math.abs(x - bullet.x) < 6);
    if (lvl < 0) { lvls.push(bullet.x); lvls.sort((a, b) => a - b); lvl = lvls.indexOf(bullet.x); }
    attrs.ilvl = Math.min(8, lvl);
  }
  const spacing = {};
  if (gap != null) {
    const extra = gap - size * 1.2;
    if (extra > 1.5) spacing.before = twip(extra);
  }
  if (p.lines.length >= 2) {
    const ratio = p.pitch / (size * 1.17);
    if (ratio > 1.12 || ratio < 0.9) { spacing.line = String(Math.round(Math.max(200, Math.min(720, ratio * 240)))); spacing.lineRule = 'auto'; }
  }
  if (Object.keys(spacing).length) attrs.spacing = spacing;
  if (tabs.length) attrs.tabs = tabs;
  if (p.outline != null && !isHF) { attrs.outlineLvl = p.outline; attrs.keepNext = true; }
  if (p.shd) attrs.shd = { val: 'clear', color: 'auto', fill: p.shd };
  if (p.breakBefore) attrs.pageBreakBefore = true;
  return nodes.paragraph.create(attrs, content);
}

let mediaN = 0;

function imageParagraph(img, ctx, breakBefore) {
  const { model } = ctx;
  const d = img.data;
  mediaN++;
  let path = 'word/media/pdf' + mediaN + '.' + d.ext;
  while (model.media.has(path)) { mediaN++; path = 'word/media/pdf' + mediaN + '.' + d.ext; }
  model.media.set(path, { url: URL.createObjectURL(new Blob([d.bytes], { type: d.type })), bytes: d.bytes, ext: d.ext, type: d.type });
  const maxW = ctx.pageW - ctx.mar.left - ctx.mar.right;
  const maxH = (ctx.pageH || 792) - ctx.mar.top - ctx.mar.bottom - 12;
  let w = img.w;
  let h = img.h;
  const k = Math.min(1, maxW / w, maxH / h);
  w *= k;
  h *= k;
  const cx = Math.round(w * 12700);
  const cy = Math.round(h * 12700);
  const attrs = {};
  const center = img.x + img.w / 2;
  if (img.full || Math.abs(center - ctx.pageW / 2) < 20) attrs.jc = 'center';
  else if (img.x - ctx.mar.left > 3) attrs.ind = { left: twip(Math.min(img.x - ctx.mar.left, maxW - w)) };
  if (breakBefore) attrs.pageBreakBefore = true;
  return nodes.paragraph.create(attrs, [nodes.drawing.create({ kind: 'image', mediaPath: path, src: model.media.get(path).url, cx, cy, ocx: cx, ocy: cy, ext: d.ext })]);
}

// ---------------------------------------------------------------- tabelas

function cluster(values, tol) {
  const sorted = [...values].sort((a, b) => a - b);
  const out = [];
  for (const v of sorted) {
    const last = out[out.length - 1];
    if (last && v - last.max <= tol) { last.max = v; last.sum += v; last.n++; } else out.push({ min: v, max: v, sum: v, n: 1 });
  }
  return out.map((c) => c.sum / c.n);
}

// Acha tabelas pelas linhas horizontais e verticais que se cruzam
function findTables(segs, items, fills, pageH) {
  const H = [];
  const V = [];
  for (const sg of segs) {
    const dx = Math.abs(sg.x2 - sg.x1);
    const dy = Math.abs(sg.y2 - sg.y1);
    if (dy < 1.5 && dx > 3) H.push({ y: (sg.y1 + sg.y2) / 2, x1: Math.min(sg.x1, sg.x2), x2: Math.max(sg.x1, sg.x2), color: sg.color });
    else if (dx < 1.5 && dy > 3) V.push({ x: (sg.x1 + sg.x2) / 2, y1: Math.min(sg.y1, sg.y2), y2: Math.max(sg.y1, sg.y2), color: sg.color });
  }
  if (H.length < 2 || V.length < 2 || H.length + V.length > 6000) return [];
  // grupos de linhas que se tocam
  const all = [...H.map((h) => ({ h })), ...V.map((v) => ({ v }))];
  const parent = all.map((_, i) => i);
  const find = (i) => { while (parent[i] !== i) { parent[i] = parent[parent[i]]; i = parent[i]; } return i; };
  const tol = 2.5;
  for (let i = 0; i < H.length; i++) {
    const h = H[i];
    for (let j = 0; j < V.length; j++) {
      const v = V[j];
      if (v.x >= h.x1 - tol && v.x <= h.x2 + tol && h.y >= v.y1 - tol && h.y <= v.y2 + tol) parent[find(i)] = find(H.length + j);
    }
  }
  const groups = new Map();
  all.forEach((e, i) => { const r = find(i); if (!groups.has(r)) groups.set(r, { H: [], V: [] }); if (e.h) groups.get(r).H.push(e.h); else groups.get(r).V.push(e.v); });
  const tables = [];
  for (const g of groups.values()) {
    if (g.H.length < 2 || g.V.length < 2) continue;
    // bordas duplas e quase coincidentes viram uma só
    const ys = cluster(g.H.map((h) => h.y), 4.5);
    const xs = cluster(g.V.map((v) => v.x), 4.5);
    if (ys.length < 2 || xs.length < 2 || (ys.length - 1) * (xs.length - 1) < 2) continue;
    const covH = (y, xa, xb) => g.H.some((h) => Math.abs(h.y - y) < 4.5 && h.x1 <= xa + (xb - xa) * 0.3 && h.x2 >= xb - (xb - xa) * 0.3);
    const covV = (x, ya, yb) => g.V.some((v) => Math.abs(v.x - x) < 4.5 && v.y1 <= ya + (yb - ya) * 0.3 && v.y2 >= yb - (yb - ya) * 0.3);
    const R = ys.length - 1;
    const C = xs.length - 1;
    const inner = items.filter((it) => it.str.trim() && it.x + it.w / 2 > xs[0] && it.x + it.w / 2 < xs[C] && it.y - it.size * 0.3 > ys[0] && it.y - it.size * 0.3 < ys[R]);
    // faixa sem linha vertical, mas com texto dos dois lados e nada atravessando: são colunas diferentes
    const splitsAt = (bx, y0, y1) => {
      const band = inner.filter((it) => { const cy = it.y - it.size * 0.3; return cy > y0 && cy < y1; });
      return band.some((it) => it.x + it.w <= bx + 1) && band.some((it) => it.x >= bx - 1) && !band.some((it) => it.x < bx - 1 && it.x + it.w > bx + 1);
    };
    const taken = Array.from({ length: R }, () => new Array(C).fill(false));
    const cells = [];
    for (let r = 0; r < R; r++) {
      for (let c = 0; c < C; c++) {
        if (taken[r][c]) continue;
        let cs = 1;
        while (c + cs < C && !covV(xs[c + cs], ys[r], ys[r + 1]) && !splitsAt(xs[c + cs], ys[r], ys[r + 1])) cs++;
        let rs = 1;
        while (r + rs < R && ![...Array(cs).keys()].some((k) => covH(ys[r + rs], xs[c + k], xs[c + k + 1]))) rs++;
        for (let a = 0; a < rs; a++) for (let b = 0; b < cs; b++) taken[r + a][c + b] = true;
        cells.push({ r, c, rs, cs, x: xs[c], right: xs[c + cs], y: ys[r], bottom: ys[r + rs], items: [] });
      }
    }
    // moldura de página não é tabela: célula mais alta que meia página
    if (cells.some((c) => c.bottom - c.y > pageH * 0.45)) continue;
    // toda linha da grade precisa começar ao menos uma célula
    if ([...Array(R).keys()].some((r) => !cells.some((c) => c.r === r))) continue;
    const colorCount = new Map();
    for (const e of [...g.H, ...g.V]) colorCount.set(e.color, (colorCount.get(e.color) || 0) + 1);
    const color = [...colorCount.entries()].sort((a, b) => b[1] - a[1])[0][0] || '000000';
    const t = { x: xs[0], right: xs[C], y: ys[0], bottom: ys[R], xs, ys, cells, color };
    for (const it of items) {
      const cx = it.x + it.w / 2;
      const cy = it.y - it.size * 0.3;
      if (cx < t.x || cx > t.right || cy < t.y || cy > t.bottom) continue;
      const cell = cells.find((c) => cx >= c.x && cx <= c.right && cy >= c.y && cy <= c.bottom);
      if (cell) cell.items.push(it);
    }
    // grade sem texto em pelo menos duas células é desenho, não tabela
    if (cells.filter((c) => c.items.length).length < 2) continue;
    for (const c of cells) {
      const area = (c.right - c.x) * (c.bottom - c.y);
      const byColor = new Map();
      for (const r of fills) {
        const w = Math.min(r.right, c.right) - Math.max(r.x, c.x);
        const h = Math.min(r.bottom, c.bottom) - Math.max(r.y, c.y);
        if (w > 0 && h > 0 && (r.right - r.x) * (r.bottom - r.y) < area * 4) byColor.set(r.color, (byColor.get(r.color) || 0) + w * h);
      }
      const best = [...byColor.entries()].sort((a, b) => b[1] - a[1])[0];
      if (best && best[1] >= area * 0.5 && best[0] !== 'FFFFFF') c.fill = best[0];
    }
    tables.push(t);
  }
  return tables;
}

function sameColumns(a, b) {
  const last = a.more && a.more.length ? a.more[a.more.length - 1] : a;
  return last.xs.length === b.xs.length && last.xs.every((x, i) => Math.abs(x - b.xs[i]) < 3);
}

function cellText(t, r) {
  return t.cells.filter((c) => c.r === r).sort((a, b) => a.c - b.c).map((c) => c.items.map((it) => it.str).join('').trim()).join('|');
}

function tableNode(t, ctx, isHF) {
  const parts = [t, ...(t.more || [])];
  const widths = t.xs.slice(1).map((x, i) => Math.round((x - t.xs[i]) * 20));
  const pad = 5.4;
  const head = cellText(t, 0);
  const rows = [];
  parts.forEach((part, pi) => {
    for (let r = 0; r < part.ys.length - 1; r++) {
      // cabeçalho repetido no alto da página seguinte: fica só o primeiro, marcado para repetir
      if (pi > 0 && r === 0 && head && cellText(part, 0) === head) { if (rows[0]) rows[0] = rows[0].type.create({ ...rows[0].attrs, header: true }, rows[0].content); continue; }
      const cells = part.cells.filter((c) => c.r === r).sort((a, b) => a.c - b.c).map((c) => {
        const cellW = c.right - c.x;
        const cctx = { ...ctx, mar: { ...ctx.mar, left: c.x + pad, right: ctx.pageW - (c.right - pad) } };
        const lines = buildLines(c.items, cellW);
        columnRights(lines);
        const ps = buildParagraphs(lines);
        let prevY = null;
        const content = ps.map((q) => {
          const node = paragraphNode(q, cctx, prevY != null ? q.lines[0].y - prevY : null, isHF);
          prevY = q.lines[q.lines.length - 1].y;
          return node;
        });
        const attrs = { colspan: c.cs, rowspan: c.rs, colwidth: widths.slice(c.c, c.c + c.cs).map((w) => w / 15), tcW: { w: String(widths.slice(c.c, c.c + c.cs).reduce((a, b) => a + b, 0)), type: 'dxa' } };
        if (c.fill) attrs.shd = { val: 'clear', color: 'auto', fill: c.fill };
        if (ps.length) {
          const above = ps[0].top - c.y;
          const below = c.bottom - ps[ps.length - 1].bottom;
          if (above > 4 && below > 4 && Math.abs(above - below) < Math.max(3, (above + below) * 0.2)) attrs.vAlign = 'center';
          else if (above > below + 6 && below < 4) attrs.vAlign = 'bottom';
        }
        return nodes.table_cell.create(attrs, content.length ? content : [nodes.paragraph.create()]);
      });
      rows.push(nodes.table_row.create({ height: String(Math.round((part.ys[r + 1] - part.ys[r]) * 20)), hRule: 'atLeast' }, cells));
    }
  });
  if (!rows.length) return null;
  const brd = { val: 'single', sz: '4', space: '0', color: t.color && t.color !== '000000' ? t.color : 'auto' };
  const tblBorders = { top: brd, left: brd, bottom: brd, right: brd, insideH: brd, insideV: brd };
  const ind = Math.round((t.x - ctx.mar.left) * 20);
  try {
    return nodes.table.create({ grid: widths, tblW: { w: '0', type: 'auto' }, tblBorders, tblInd: ind ? { w: String(ind), type: 'dxa' } : null }, rows);
  } catch {
    return null;
  }
}

// Posição de cada número de uma linha repetida: 'page' quando acompanha a
// página em todas as ocorrências, 'total' quando é sempre o total de páginas
function numberRoles(list, total) {
  if (!list || list.length < 2) return [];
  const n = list[0].nums.length;
  if (!list.every((x) => x.nums.length === n)) return [];
  const roles = [];
  for (let i = 0; i < n; i++) {
    const d = list[0].nums[i] - list[0].page;
    if (list.every((x) => x.nums[i] - x.page === d) && new Set(list.map((x) => x.page)).size > 1) roles.push('page');
    else if (list.every((x) => x.nums[i] === total)) roles.push('total');
    else roles.push(null);
  }
  return roles;
}

function tableNumbers(t) {
  const out = [];
  for (const c of [...t.cells].sort((a, b) => a.r - b.r || a.c - b.c)) {
    for (const l of buildLines(c.items, c.right - c.x)) for (const it of l.items) for (const m of it.str.match(/\d+/g) || []) out.push(Number(m));
  }
  return out;
}
