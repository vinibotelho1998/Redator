// Numeração e marcadores (numbering.xml).
import { parseXml, children, child, wa, isW, serialize, NS } from '../util/xml.js';
import { parsePPr, parseRPr } from './props.js';

const W_DECL = 'xmlns:w="' + NS.w + '"';

function parseLevel(el) {
  const lvl = {
    ilvl: parseInt(wa(el, 'ilvl'), 10) || 0,
    start: null, numFmt: 'decimal', format: null, lvlText: '', lvlJc: 'left', suff: 'tab',
    pPr: null, rPr: null, lvlRestart: null, isLgl: false, pStyle: null, picBulletId: null,
  };
  const handle = (c) => {
    if (!isW(c)) {
      if (c.localName === 'AlternateContent') {
        // usa o Fallback, que é compatível
        const fb = Array.from(c.children).find((x) => x.localName === 'Fallback');
        if (fb) for (const x of fb.children) handle(x);
      }
      return;
    }
    switch (c.localName) {
      case 'start': lvl.start = parseInt(wa(c, 'val'), 10); break;
      case 'numFmt': lvl.numFmt = wa(c, 'val') || 'decimal'; lvl.format = wa(c, 'format'); break;
      case 'lvlText': lvl.lvlText = wa(c, 'val') ?? ''; lvl.lvlTextNull = wa(c, 'null') === '1'; break;
      case 'lvlJc': lvl.lvlJc = wa(c, 'val') || 'left'; break;
      case 'suff': lvl.suff = wa(c, 'val') || 'tab'; break;
      case 'pPr': lvl.pPr = parsePPr(c); break;
      case 'rPr': lvl.rPr = parseRPr(c); break;
      case 'lvlRestart': lvl.lvlRestart = parseInt(wa(c, 'val'), 10); break;
      case 'isLgl': lvl.isLgl = wa(c, 'val') !== '0'; break;
      case 'pStyle': lvl.pStyle = wa(c, 'val'); break;
      case 'lvlPicBulletId': lvl.picBulletId = wa(c, 'val'); break;
      default: break;
    }
  };
  for (const c of el.children) handle(c);
  return lvl;
}

export class Numbering {
  constructor() {
    this.abstracts = new Map();
    this.nums = new Map();
    this.doc = null;
    this.text = null;
    this.dirty = false;
    this.styles = null;
  }

  static parse(text) {
    const n = new Numbering();
    if (!text) return n;
    n.text = text;
    try { n.doc = parseXml(text); } catch { return n; }
    const root = n.doc.documentElement;
    for (const el of children(root, 'abstractNum')) n.addAbstractEl(el);
    for (const el of children(root, 'num')) n.addNumEl(el);
    return n;
  }

  addAbstractEl(el) {
    const id = wa(el, 'abstractNumId');
    const abs = {
      id,
      levels: [],
      numStyleLink: null,
      styleLink: null,
      multiLevelType: null,
    };
    for (const c of el.children) {
      if (!isW(c)) continue;
      if (c.localName === 'lvl') { const l = parseLevel(c); abs.levels[l.ilvl] = l; }
      else if (c.localName === 'numStyleLink') abs.numStyleLink = wa(c, 'val');
      else if (c.localName === 'styleLink') abs.styleLink = wa(c, 'val');
      else if (c.localName === 'multiLevelType') abs.multiLevelType = wa(c, 'val');
    }
    this.abstracts.set(id, abs);
    return abs;
  }

  addNumEl(el) {
    const id = wa(el, 'numId');
    const absEl = child(el, 'abstractNumId');
    const num = { id, abstractNumId: absEl ? wa(absEl, 'val') : null, overrides: new Map() };
    for (const ov of children(el, 'lvlOverride')) {
      const ilvl = parseInt(wa(ov, 'ilvl'), 10) || 0;
      const so = child(ov, 'startOverride');
      const lv = child(ov, 'lvl');
      num.overrides.set(ilvl, {
        startOverride: so ? parseInt(wa(so, 'val'), 10) : null,
        lvl: lv ? parseLevel(lv) : null,
      });
    }
    this.nums.set(id, num);
    return num;
  }

  abstractFor(numId, depth = 0) {
    const num = this.nums.get(String(numId));
    if (!num) return null;
    let abs = this.abstracts.get(num.abstractNumId);
    if (abs && abs.numStyleLink && this.styles && depth < 4) {
      const st = this.styles.get(abs.numStyleLink);
      const linkedNum = st && st.pPr && st.pPr.numId;
      if (linkedNum && linkedNum !== String(numId)) {
        const other = this.abstractFor(linkedNum, depth + 1);
        if (other) abs = other;
      }
    }
    return abs || null;
  }

  level(numId, ilvl) {
    if (numId == null || numId === '0') return null;
    const num = this.nums.get(String(numId));
    if (!num) return null;
    const ov = num.overrides.get(ilvl || 0);
    if (ov && ov.lvl) return ov.lvl;
    const abs = this.abstractFor(numId);
    if (!abs) return null;
    return abs.levels[ilvl || 0] || null;
  }

  // Cria uma definição nova e devolve o numId da instância
  createList(kind, variant = 0) {
    const absId = String(this.nextAbstractId());
    const numId = String(this.nextNumId());
    const xml = kind === 'bullet' ? bulletAbstractXml(absId, variant) : numberAbstractXml(absId, variant);
    const el = parseXml('<w:root ' + W_DECL + '>' + xml + '</w:root>').documentElement.firstElementChild;
    this.ensureDoc();
    const root = this.doc.documentElement;
    const imported = this.doc.importNode(el, true);
    const firstNum = children(root, 'num')[0];
    if (firstNum) root.insertBefore(imported, firstNum);
    else root.appendChild(imported);
    this.addAbstractEl(imported);
    this.addNumInstance(numId, absId, null);
    this.dirty = true;
    return numId;
  }

  // Lista a partir de uma descrição de níveis (usada na importação de ODT)
  createCustomList(levels) {
    const absId = String(this.nextAbstractId());
    const numId = String(this.nextNumId());
    const FMT = { 1: 'decimal', a: 'lowerLetter', A: 'upperLetter', i: 'lowerRoman', I: 'upperRoman', '': 'none' };
    let x = '<w:abstractNum w:abstractNumId="' + absId + '"><w:multiLevelType w:val="hybridMultilevel"/>';
    for (let i = 0; i < 9; i++) {
      const l = levels[i] || levels[levels.length - 1] || { bullet: true, char: '•' };
      const left = l.left != null ? l.left : 720 * (i + 1);
      const hanging = l.hanging != null ? Math.max(0, l.hanging) : 360;
      if (l.bullet) {
        x += '<w:lvl w:ilvl="' + i + '"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="' + esc(l.char || '•') + '"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="' + left + '" w:hanging="' + hanging + '"/></w:pPr></w:lvl>';
      } else {
        const fmt = FMT[l.fmt] || 'decimal';
        const disp = Math.max(1, Math.min(i + 1, l.display || 1));
        let text = '';
        for (let k = i - disp + 1; k <= i; k++) text += (text ? '.' : '') + '%' + (k + 1);
        text = (l.prefix || '') + text + (l.suffix || '');
        x += '<w:lvl w:ilvl="' + i + '"><w:start w:val="' + (l.start || 1) + '"/><w:numFmt w:val="' + fmt + '"/><w:lvlText w:val="' + esc(text) + '"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="' + left + '" w:hanging="' + hanging + '"/></w:pPr></w:lvl>';
      }
    }
    x += '</w:abstractNum>';
    const el = parseXml('<w:root ' + W_DECL + '>' + x + '</w:root>').documentElement.firstElementChild;
    this.ensureDoc();
    const root = this.doc.documentElement;
    const imported = this.doc.importNode(el, true);
    const firstNum = children(root, 'num')[0];
    if (firstNum) root.insertBefore(imported, firstNum); else root.appendChild(imported);
    this.addAbstractEl(imported);
    this.addNumInstance(numId, absId, null);
    this.dirty = true;
    return numId;
  }

  // Nova instância reiniciando a numeração
  restartList(numId) {
    const num = this.nums.get(String(numId));
    if (!num) return numId;
    const newId = String(this.nextNumId());
    this.addNumInstance(newId, num.abstractNumId, 1);
    return newId;
  }

  addNumInstance(numId, absId, startOverride) {
    this.ensureDoc();
    const root = this.doc.documentElement;
    let xml = '<w:num w:numId="' + numId + '"><w:abstractNumId w:val="' + absId + '"/>';
    if (startOverride != null) {
      for (let i = 0; i < 9; i++) xml += '<w:lvlOverride w:ilvl="' + i + '"><w:startOverride w:val="' + (i === 0 ? startOverride : 1) + '"/></w:lvlOverride>';
    }
    xml += '</w:num>';
    const el = parseXml('<w:root ' + W_DECL + '>' + xml + '</w:root>').documentElement.firstElementChild;
    const imported = this.doc.importNode(el, true);
    const cleanup = children(root, 'numIdMacAtCleanup')[0];
    if (cleanup) root.insertBefore(imported, cleanup);
    else root.appendChild(imported);
    this.addNumEl(imported);
    this.dirty = true;
  }

  ensureDoc() {
    if (this.doc) return;
    this.doc = parseXml('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ' + W_DECL + ' xmlns:r="' + NS.r + '"></w:numbering>');
  }

  nextAbstractId() {
    let max = -1;
    for (const k of this.abstracts.keys()) max = Math.max(max, parseInt(k, 10) || 0);
    return max + 1;
  }

  nextNumId() {
    let max = 0;
    for (const k of this.nums.keys()) max = Math.max(max, parseInt(k, 10) || 0);
    return max + 1;
  }

  // Tipo aproximado da lista: bullet ou number
  kindOf(numId, ilvl = 0) {
    const l = this.level(numId, ilvl);
    if (!l) return null;
    return l.numFmt === 'bullet' ? 'bullet' : 'number';
  }

  toXml() {
    if (!this.doc) return null;
    let s = new XMLSerializer().serializeToString(this.doc);
    if (!s.startsWith('<?xml')) s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + s;
    return s;
  }

  get isEmpty() {
    return this.nums.size === 0;
  }
}

// ---------------------------------------------------------------- formatação de números

const ROMAN = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']];

export function toRoman(n) {
  if (n <= 0) return String(n);
  let s = '';
  for (const [v, r] of ROMAN) while (n >= v) { s += r; n -= v; }
  return s;
}

export function toLetter(n) {
  if (n <= 0) return String(n);
  const ch = String.fromCharCode(65 + ((n - 1) % 26));
  return ch.repeat(Math.floor((n - 1) / 26) + 1);
}

const UNITS = ['zero', 'um', 'dois', 'três', 'quatro', 'cinco', 'seis', 'sete', 'oito', 'nove', 'dez', 'onze', 'doze', 'treze', 'catorze', 'quinze', 'dezesseis', 'dezessete', 'dezoito', 'dezenove'];
const TENS = ['', '', 'vinte', 'trinta', 'quarenta', 'cinquenta', 'sessenta', 'setenta', 'oitenta', 'noventa'];
const ORD = ['', 'primeiro', 'segundo', 'terceiro', 'quarto', 'quinto', 'sexto', 'sétimo', 'oitavo', 'nono', 'décimo'];

function cardinal(n) {
  if (n < 20) return UNITS[n];
  if (n < 100) return TENS[Math.floor(n / 10)] + (n % 10 ? ' e ' + UNITS[n % 10] : '');
  return String(n);
}

const cap = (s) => s.charAt(0).toUpperCase() + s.slice(1);

export function formatNumber(n, fmt, format) {
  switch (fmt) {
    case 'decimal': return String(n);
    case 'decimalZero': return n < 10 && n >= 0 ? '0' + n : String(n);
    case 'upperRoman': return toRoman(n);
    case 'lowerRoman': return toRoman(n).toLowerCase();
    case 'upperLetter': return toLetter(n);
    case 'lowerLetter': return toLetter(n).toLowerCase();
    case 'ordinal': return n + 'º';
    case 'cardinalText': return cap(cardinal(n));
    case 'ordinalText': return cap(n <= 10 ? ORD[n] : n + 'º');
    case 'bullet': case 'none': return '';
    case 'decimalEnclosedCircle': case 'decimalEnclosedCircleChinese':
      return n >= 1 && n <= 20 ? String.fromCharCode(0x2460 + n - 1) : String(n);
    case 'decimalEnclosedParen': return '(' + n + ')';
    case 'decimalEnclosedFullstop': return n + '.';
    case 'chicago': { const s = ['*', '†', '‡', '§']; return s[(n - 1) % 4].repeat(Math.floor((n - 1) / 4) + 1); }
    case 'custom': {
      const first = (format || '').split(',')[0].trim();
      if (/^0+\d$/.test(first)) return String(n).padStart(first.length, '0');
      return String(n);
    }
    default: return String(n);
  }
}

// Caracteres de marcador das fontes de símbolo para Unicode
const SYMBOL_MAP = { 0xB7: '•', 0xA7: '♣', 0xA8: '♦', 0xA9: '♥', 0xAA: '♠', 0xDE: '⇒', 0xAE: '→', 0x2D: '−', 0x2A: '∗', 0xD8: '¬', 0xE0: '◊' };
const WINGDINGS_MAP = { 0xA7: '▪', 0xA8: '◻', 0x6E: '■', 0x6C: '●', 0x71: '❑', 0x76: '❖', 0xD8: '➢', 0xFC: '✓', 0xFB: '✗', 0xA1: '○', 0x9F: '•', 0x77: '⬥', 0x75: '◆', 0x6F: '□', 0x70: '◼', 0xE8: '➔', 0xF0: '⇨', 0x9E: '·', 0xAB: '★', 0x4A: '☺', 0xFE: '☑', 0xA4: '◉' };

export function bulletText(text, font) {
  if (!text) return { text: '', font: null };
  const f = (font || '').toLowerCase();
  const code = text.codePointAt(0);
  const low = code >= 0xF000 && code <= 0xF0FF ? code - 0xF000 : code;
  if (f === 'symbol' && SYMBOL_MAP[low]) return { text: SYMBOL_MAP[low], font: null };
  if (f.startsWith('wingdings') && WINGDINGS_MAP[low]) return { text: WINGDINGS_MAP[low], font: null };
  if (code >= 0xF000 && code <= 0xF0FF) return { text, font };
  if (f === 'symbol' || f.startsWith('wingdings') || f === 'webdings') {
    if (SYMBOL_MAP[low] && f === 'symbol') return { text: SYMBOL_MAP[low], font: null };
    return { text: String.fromCharCode(0xF000 + low), font };
  }
  if (text === '·' || text === '·') return { text: '•', font: null };
  return { text, font: null };
}

// ---------------------------------------------------------------- contadores

export class NumberingCounter {
  constructor(numbering) {
    this.n = numbering;
    this.counters = new Map();
    this.started = new Set();
  }

  next(numId, ilvl) {
    const n = this.n;
    if (!n || numId == null || numId === '0') return null;
    const num = n.nums.get(String(numId));
    if (!num) return null;
    const abs = n.abstractFor(numId);
    if (!abs) return null;
    ilvl = Math.max(0, Math.min(8, ilvl || 0));
    const lvlDef = (l) => n.level(numId, l) || { start: 0, numFmt: 'decimal', lvlText: '' };
    const key = abs.id;
    let ctr = this.counters.get(key);
    if (!ctr) { ctr = new Array(9).fill(null); this.counters.set(key, ctr); }
    if (!this.started.has(num.id)) {
      this.started.add(num.id);
      for (const [lv, ov] of num.overrides) {
        if (ov.startOverride != null && !Number.isNaN(ov.startOverride)) {
          ctr[lv] = ov.startOverride - 1;
          for (let d = lv + 1; d < 9; d++) ctr[d] = null;
        }
      }
    }
    const def = lvlDef(ilvl);
    if (ctr[ilvl] == null) ctr[ilvl] = (def.start ?? 0) - 1;
    ctr[ilvl] += 1;
    for (let l = ilvl + 1; l < 9; l++) {
      const d = n.level(numId, l);
      const restart = d && d.lvlRestart != null && !Number.isNaN(d.lvlRestart) ? d.lvlRestart : l;
      if (restart === 0) continue;
      if (ilvl + 1 <= restart) ctr[l] = null;
    }
    const text = (def.lvlText || '').replace(/%(\d)/g, (_m, d) => {
      const L = Number(d) - 1;
      if (L < 0 || L > 8) return '';
      const ld = lvlDef(L);
      const v = ctr[L] != null ? ctr[L] : (ld.start ?? 0);
      const fmt = def.isLgl && L < ilvl ? 'decimal' : ld.numFmt;
      return formatNumber(v, fmt, ld.format);
    });
    return { text, def, value: ctr[ilvl] };
  }
}

// ---------------------------------------------------------------- definições padrão

const BULLETS = [
  { text: '', font: 'Symbol' },
  { text: 'o', font: 'Courier New' },
  { text: '', font: 'Wingdings' },
];

export const BULLET_LIBRARY = [
  { text: '', font: 'Symbol', show: '•' },
  { text: 'o', font: 'Courier New', show: 'o' },
  { text: '', font: 'Wingdings', show: '▪' },
  { text: '', font: 'Wingdings', show: '❖' },
  { text: '', font: 'Wingdings', show: '➢' },
  { text: '', font: 'Wingdings', show: '✓' },
  { text: '–', font: 'Calibri', show: '–' },
  { text: '●', font: 'Calibri', show: '●' },
];

export const NUMBER_LIBRARY = [
  { fmt: 'decimal', text: '%1.', show: '1.  2.  3.' },
  { fmt: 'decimal', text: '%1)', show: '1)  2)  3)' },
  { fmt: 'upperRoman', text: '%1.', show: 'I.  II.  III.' },
  { fmt: 'upperLetter', text: '%1.', show: 'A.  B.  C.' },
  { fmt: 'lowerLetter', text: '%1)', show: 'a)  b)  c)' },
  { fmt: 'lowerLetter', text: '%1.', show: 'a.  b.  c.' },
  { fmt: 'lowerRoman', text: '%1.', show: 'i.  ii.  iii.' },
];

function esc(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
}

function bulletAbstractXml(id, variant) {
  const first = BULLET_LIBRARY[variant] || BULLET_LIBRARY[0];
  let x = '<w:abstractNum w:abstractNumId="' + id + '"><w:multiLevelType w:val="hybridMultilevel"/>';
  for (let i = 0; i < 9; i++) {
    const b = i === 0 ? first : BULLETS[i % 3];
    const left = 720 * (i + 1);
    x += '<w:lvl w:ilvl="' + i + '"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="' + esc(b.text) + '"/><w:lvlJc w:val="left"/>'
      + '<w:pPr><w:ind w:left="' + left + '" w:hanging="360"/></w:pPr>'
      + '<w:rPr><w:rFonts w:ascii="' + b.font + '" w:hAnsi="' + b.font + '" w:hint="default"/></w:rPr></w:lvl>';
  }
  return x + '</w:abstractNum>';
}

function numberAbstractXml(id, variant) {
  const first = NUMBER_LIBRARY[variant] || NUMBER_LIBRARY[0];
  const fmts = ['decimal', 'lowerLetter', 'lowerRoman'];
  let x = '<w:abstractNum w:abstractNumId="' + id + '"><w:multiLevelType w:val="hybridMultilevel"/>';
  for (let i = 0; i < 9; i++) {
    const fmt = i === 0 ? first.fmt : fmts[i % 3];
    const text = i === 0 ? first.text : '%' + (i + 1) + '.';
    const roman = fmt === 'lowerRoman' && i > 0;
    const left = 720 * (i + 1);
    x += '<w:lvl w:ilvl="' + i + '"><w:start w:val="1"/><w:numFmt w:val="' + fmt + '"/><w:lvlText w:val="' + esc(text) + '"/>'
      + '<w:lvlJc w:val="' + (roman ? 'right' : 'left') + '"/>'
      + '<w:pPr><w:ind w:left="' + left + '" w:hanging="' + (roman ? 180 : 360) + '"/></w:pPr></w:lvl>';
  }
  return x + '</w:abstractNum>';
}
