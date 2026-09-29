// Leitura e escrita das propriedades pPr, rPr, tblPr, trPr e tcPr.
import { children, child, wa, isW, onOff, serialize, wAttrs, attrsToXml, orderFragments, escAttr } from '../util/xml.js';

export const PPR_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange'];
export const RPR_ORDER = ['ins', 'del', 'moveFrom', 'moveTo', 'rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath', 'rPrChange'];
export const TBLPR_ORDER = ['tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual', 'tblStyleRowBandSize', 'tblStyleColBandSize', 'tblW', 'jc', 'tblCellSpacing', 'tblInd', 'tblBorders', 'shd', 'tblLayout', 'tblCellMar', 'tblLook', 'tblCaption', 'tblDescription', 'tblPrChange'];
export const TRPR_ORDER = ['cnfStyle', 'divId', 'gridBefore', 'gridAfter', 'wBefore', 'wAfter', 'cantSplit', 'trHeight', 'tblHeader', 'tblCellSpacing', 'jc', 'hidden', 'ins', 'del', 'trPrChange'];
export const TCPR_ORDER = ['cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge', 'tcBorders', 'shd', 'noWrap', 'tcMar', 'textDirection', 'tcFitText', 'vAlign', 'hideMark', 'headers', 'cellIns', 'cellDel', 'cellMerge', 'tcPrChange'];
export const BDR_ORDER = ['top', 'left', 'start', 'bottom', 'right', 'end', 'between', 'bar', 'insideH', 'insideV', 'tl2br', 'tr2bl'];
export const MAR_ORDER = ['top', 'left', 'start', 'bottom', 'right', 'end'];

export const TOGGLE_P = ['keepNext', 'keepLines', 'pageBreakBefore', 'widowControl', 'contextualSpacing'];
const TOGGLE_R = ['b', 'i', 'caps', 'smallCaps', 'strike', 'dstrike', 'vanish'];

const esc = escAttr;

function normInd(o) {
  const r = { ...o };
  if (r.start != null && r.left == null) { r.left = r.start; }
  if (r.end != null && r.right == null) { r.right = r.end; }
  delete r.start; delete r.end;
  if (r.startChars != null && r.leftChars == null) r.leftChars = r.startChars;
  if (r.endChars != null && r.rightChars == null) r.rightChars = r.endChars;
  delete r.startChars; delete r.endChars;
  return r;
}

export function parseBorders(el) {
  const b = {};
  for (const c of children(el)) {
    if (!isW(c)) continue;
    let side = c.localName;
    if (side === 'start') side = 'left';
    if (side === 'end') side = 'right';
    b[side] = wAttrs(c);
  }
  return b;
}

export function bordersXml(tag, b) {
  if (!b) return '';
  let s = '';
  for (const side of BDR_ORDER) if (b[side]) s += '<w:' + side + attrsToXml(b[side]) + '/>';
  return s ? '<w:' + tag + '>' + s + '</w:' + tag + '>' : '';
}

function parseMargins(el) {
  const m = {};
  for (const c of children(el)) {
    if (!isW(c)) continue;
    let side = c.localName;
    if (side === 'start') side = 'left';
    if (side === 'end') side = 'right';
    m[side] = wAttrs(c);
  }
  return m;
}

function marginsXml(tag, m) {
  if (!m) return '';
  let s = '';
  for (const side of MAR_ORDER) if (m[side]) s += '<w:' + side + attrsToXml(m[side]) + '/>';
  return s ? '<w:' + tag + '>' + s + '</w:' + tag + '>' : '';
}

// ---------------------------------------------------------------- parágrafo

export function parsePPr(el) {
  const a = {};
  if (!el) return a;
  const extra = [];
  for (const c of children(el)) {
    if (!isW(c)) { extra.push(serialize(c)); continue; }
    const n = c.localName;
    switch (n) {
      case 'pStyle': a.styleId = wa(c, 'val'); break;
      case 'jc': a.jc = wa(c, 'val'); break;
      case 'ind': a.ind = normInd(wAttrs(c)); break;
      case 'spacing': a.spacing = wAttrs(c); break;
      case 'numPr': {
        const il = child(c, 'ilvl'), ni = child(c, 'numId');
        if (ni) a.numId = wa(ni, 'val');
        if (il) a.ilvl = parseInt(wa(il, 'val'), 10) || 0;
        else if (ni) a.ilvl = 0;
        break;
      }
      case 'keepNext': case 'keepLines': case 'pageBreakBefore': case 'widowControl': case 'contextualSpacing':
        a[n] = onOff(c); break;
      case 'pBdr': a.pBdr = parseBorders(c); break;
      case 'shd': a.shd = wAttrs(c); break;
      case 'tabs': a.tabs = children(c, 'tab').map((t) => wAttrs(t)); break;
      case 'outlineLvl': a.outlineLvl = parseInt(wa(c, 'val'), 10); break;
      case 'rPr': { const r = parseRPr(c); if (r.length) a.markRPr = r; break; }
      case 'sectPr': a.sectPr = serialize(c); break;
      default: extra.push(serialize(c));
    }
  }
  if (extra.length) a.extra = extra;
  return a;
}

export function pPrXml(a) {
  if (!a) return '';
  const f = [];
  if (a.styleId) f.push('<w:pStyle w:val="' + esc(a.styleId) + '"/>');
  for (const k of TOGGLE_P) if (a[k] != null) f.push(a[k] ? '<w:' + k + '/>' : '<w:' + k + ' w:val="0"/>');
  if (a.numId != null || a.ilvl != null) {
    f.push('<w:numPr>' + (a.ilvl != null ? '<w:ilvl w:val="' + a.ilvl + '"/>' : '') + (a.numId != null ? '<w:numId w:val="' + esc(a.numId) + '"/>' : '') + '</w:numPr>');
  }
  if (a.pBdr) { const x = bordersXml('pBdr', a.pBdr); if (x) f.push(x); }
  if (a.shd) f.push('<w:shd' + attrsToXml(a.shd) + '/>');
  if (a.tabs && a.tabs.length) {
    const tabs = [...a.tabs].sort((x, y) => Number(x.pos) - Number(y.pos));
    f.push('<w:tabs>' + tabs.map((t) => '<w:tab' + attrsToXml(t) + '/>').join('') + '</w:tabs>');
  }
  if (a.spacing && Object.keys(a.spacing).length) f.push('<w:spacing' + attrsToXml(a.spacing) + '/>');
  if (a.ind && Object.keys(a.ind).length) f.push('<w:ind' + attrsToXml(a.ind) + '/>');
  if (a.jc) f.push('<w:jc w:val="' + esc(a.jc) + '"/>');
  if (a.outlineLvl != null && Number.isFinite(a.outlineLvl)) f.push('<w:outlineLvl w:val="' + a.outlineLvl + '"/>');
  if (a.markRPr && a.markRPr.length) { const r = rPrXml(a.markRPr); if (r) f.push(r); }
  if (a.sectPr) f.push(a.sectPr);
  if (a.extra) f.push(...a.extra);
  if (!f.length) return '';
  return '<w:pPr>' + orderFragments(f, PPR_ORDER).join('') + '</w:pPr>';
}

// ---------------------------------------------------------------- execução

// Devolve uma lista de {type, attrs} compatível com as marcas do esquema
export function parseRPr(el) {
  const out = [];
  if (!el) return out;
  const extra = [];
  for (const c of children(el)) {
    if (!isW(c)) { extra.push(serialize(c)); continue; }
    const n = c.localName;
    switch (n) {
      case 'rStyle': out.push({ type: 'rStyle', attrs: { val: wa(c, 'val') } }); break;
      case 'rFonts': out.push({ type: 'rFonts', attrs: { a: wAttrs(c) } }); break;
      case 'b': case 'i': case 'caps': case 'smallCaps': case 'strike': case 'dstrike': case 'vanish':
        out.push({ type: n, attrs: { val: onOff(c) } }); break;
      case 'color': out.push({ type: 'color', attrs: { a: wAttrs(c) } }); break;
      case 'sz': {
        const v = parseFloat(wa(c, 'val'));
        if (v > 0) out.push({ type: 'sz', attrs: { val: v } });
        break;
      }
      case 'highlight': out.push({ type: 'highlight', attrs: { val: wa(c, 'val') || 'none' } }); break;
      case 'u': { const a = wAttrs(c); if (!a.val) a.val = 'single'; out.push({ type: 'u', attrs: { a } }); break; }
      case 'shd': out.push({ type: 'shd', attrs: { a: wAttrs(c) } }); break;
      case 'vertAlign': out.push({ type: 'vertAlign', attrs: { val: wa(c, 'val') || 'baseline' } }); break;
      case 'spacing': {
        const v = parseInt(wa(c, 'val'), 10);
        if (Number.isFinite(v)) out.push({ type: 'spacing', attrs: { val: v } });
        else extra.push(serialize(c));
        break;
      }
      default: extra.push(serialize(c));
    }
  }
  if (extra.length) out.push({ type: 'rprExtra', attrs: { items: extra } });
  return out;
}

const typeName = (m) => (typeof m.type === 'string' ? m.type : m.type.name);

export function rPrXml(marks) {
  if (!marks) return '';
  const f = [];
  for (const m of marks) {
    const t = typeName(m);
    const a = m.attrs || {};
    switch (t) {
      case 'rStyle': if (a.val) f.push('<w:rStyle w:val="' + esc(a.val) + '"/>'); break;
      case 'rFonts': f.push('<w:rFonts' + attrsToXml(a.a) + '/>'); break;
      case 'color': f.push('<w:color' + attrsToXml(a.a) + '/>'); break;
      case 'sz': f.push('<w:sz w:val="' + Math.round(a.val) + '"/>'); break;
      case 'highlight': f.push('<w:highlight w:val="' + esc(a.val || 'yellow') + '"/>'); break;
      case 'u': f.push('<w:u' + attrsToXml(a.a || { val: 'single' }) + '/>'); break;
      case 'shd': f.push('<w:shd' + attrsToXml(a.a) + '/>'); break;
      case 'vertAlign': f.push('<w:vertAlign w:val="' + esc(a.val || 'baseline') + '"/>'); break;
      case 'spacing': f.push('<w:spacing w:val="' + Math.round(a.val || 0) + '"/>'); break;
      case 'rprExtra': if (a.items) f.push(...a.items); break;
      default:
        if (TOGGLE_R.includes(t)) f.push(a.val === false ? '<w:' + t + ' w:val="0"/>' : '<w:' + t + '/>');
    }
  }
  if (!f.length) return '';
  return '<w:rPr>' + orderFragments(f, RPR_ORDER).join('') + '</w:rPr>';
}

// ---------------------------------------------------------------- tabela

export function parseTblPr(el) {
  const a = {};
  if (!el) return a;
  const extra = [];
  for (const c of children(el)) {
    if (!isW(c)) { extra.push(serialize(c)); continue; }
    switch (c.localName) {
      case 'tblStyle': a.styleId = wa(c, 'val'); break;
      case 'tblW': a.tblW = wAttrs(c); break;
      case 'jc': a.jc = wa(c, 'val'); break;
      case 'tblInd': a.tblInd = wAttrs(c); break;
      case 'tblBorders': a.tblBorders = parseBorders(c); break;
      case 'shd': a.shd = wAttrs(c); break;
      case 'tblLayout': a.tblLayout = wa(c, 'type'); break;
      case 'tblCellMar': a.tblCellMar = parseMargins(c); break;
      case 'tblLook': a.tblLook = wAttrs(c); break;
      case 'tblpPr': a.tblpPr = wAttrs(c); break;
      default: extra.push(serialize(c));
    }
  }
  if (extra.length) a.extra = extra;
  return a;
}

export function tblPrXml(a) {
  const f = [];
  if (a.styleId) f.push('<w:tblStyle w:val="' + esc(a.styleId) + '"/>');
  if (a.tblpPr) f.push('<w:tblpPr' + attrsToXml(a.tblpPr) + '/>');
  if (a.tblW) f.push('<w:tblW' + attrsToXml(a.tblW) + '/>');
  if (a.jc) f.push('<w:jc w:val="' + esc(a.jc) + '"/>');
  if (a.tblInd) f.push('<w:tblInd' + attrsToXml(a.tblInd) + '/>');
  if (a.tblBorders) { const x = bordersXml('tblBorders', a.tblBorders); if (x) f.push(x); }
  if (a.shd) f.push('<w:shd' + attrsToXml(a.shd) + '/>');
  if (a.tblLayout) f.push('<w:tblLayout w:type="' + esc(a.tblLayout) + '"/>');
  if (a.tblCellMar) { const x = marginsXml('tblCellMar', a.tblCellMar); if (x) f.push(x); }
  if (a.tblLook) f.push('<w:tblLook' + attrsToXml(a.tblLook) + '/>');
  if (a.extra) f.push(...a.extra);
  return '<w:tblPr>' + orderFragments(f, TBLPR_ORDER).join('') + '</w:tblPr>';
}

export function parseTrPr(el) {
  const a = {};
  if (!el) return a;
  const extra = [];
  for (const c of children(el)) {
    if (!isW(c)) { extra.push(serialize(c)); continue; }
    switch (c.localName) {
      case 'trHeight': a.height = parseInt(wa(c, 'val'), 10) || null; a.hRule = wa(c, 'hRule'); break;
      case 'tblHeader': a.header = onOff(c); break;
      case 'cantSplit': a.cantSplit = onOff(c); break;
      case 'gridBefore': a.gridBefore = parseInt(wa(c, 'val'), 10) || 0; break;
      case 'gridAfter': a.gridAfter = parseInt(wa(c, 'val'), 10) || 0; break;
      default: extra.push(serialize(c));
    }
  }
  if (extra.length) a.extra = extra;
  return a;
}

export function trPrXml(a) {
  const f = [];
  if (a.gridBefore) f.push('<w:gridBefore w:val="' + a.gridBefore + '"/>');
  if (a.gridAfter) f.push('<w:gridAfter w:val="' + a.gridAfter + '"/>');
  if (a.cantSplit != null) f.push(a.cantSplit ? '<w:cantSplit/>' : '<w:cantSplit w:val="0"/>');
  if (a.height) f.push('<w:trHeight w:val="' + a.height + '"' + (a.hRule ? ' w:hRule="' + esc(a.hRule) + '"' : '') + '/>');
  if (a.header != null) f.push(a.header ? '<w:tblHeader/>' : '<w:tblHeader w:val="0"/>');
  if (a.extra) f.push(...a.extra);
  if (!f.length) return '';
  return '<w:trPr>' + orderFragments(f, TRPR_ORDER).join('') + '</w:trPr>';
}

export function parseTcPr(el) {
  const a = {};
  if (!el) return a;
  const extra = [];
  for (const c of children(el)) {
    if (!isW(c)) { extra.push(serialize(c)); continue; }
    switch (c.localName) {
      case 'tcW': a.tcW = wAttrs(c); break;
      case 'gridSpan': a.gridSpan = parseInt(wa(c, 'val'), 10) || 1; break;
      case 'vMerge': a.vMerge = wa(c, 'val') || 'continue'; break;
      case 'hMerge': a.hMerge = wa(c, 'val') || 'continue'; break;
      case 'tcBorders': a.tcBorders = parseBorders(c); break;
      case 'shd': a.shd = wAttrs(c); break;
      case 'vAlign': a.vAlign = wa(c, 'val'); break;
      case 'tcMar': a.tcMar = parseMargins(c); break;
      case 'noWrap': a.noWrap = onOff(c); break;
      case 'textDirection': a.textDirection = wa(c, 'val'); break;
      case 'cnfStyle': break; // o Word recalcula
      default: extra.push(serialize(c));
    }
  }
  if (extra.length) a.extra = extra;
  return a;
}

export function tcPrXml(a, gridSpan, vMerge) {
  const f = [];
  if (a.tcW) f.push('<w:tcW' + attrsToXml(a.tcW) + '/>');
  if (gridSpan > 1) f.push('<w:gridSpan w:val="' + gridSpan + '"/>');
  if (vMerge) f.push(vMerge === 'restart' ? '<w:vMerge w:val="restart"/>' : '<w:vMerge/>');
  if (a.tcBorders) { const x = bordersXml('tcBorders', a.tcBorders); if (x) f.push(x); }
  if (a.shd) f.push('<w:shd' + attrsToXml(a.shd) + '/>');
  if (a.noWrap) f.push('<w:noWrap/>');
  if (a.tcMar) { const x = marginsXml('tcMar', a.tcMar); if (x) f.push(x); }
  if (a.textDirection) f.push('<w:textDirection w:val="' + esc(a.textDirection) + '"/>');
  if (a.vAlign) f.push('<w:vAlign w:val="' + esc(a.vAlign) + '"/>');
  if (a.extra) f.push(...a.extra);
  if (!f.length) return '';
  return '<w:tcPr>' + orderFragments(f, TCPR_ORDER).join('') + '</w:tcPr>';
}

// ---------------------------------------------------------------- mesclagem para estilos

export function mergeInd(base, over) {
  if (!base) return over ? { ...over } : null;
  if (!over) return { ...base };
  const r = { ...base, ...over };
  if (over.firstLine != null && over.hanging == null) delete r.hanging;
  if (over.hanging != null && over.firstLine == null) delete r.firstLine;
  return r;
}

export function mergeTabs(base, over) {
  if (!over) return base ? base.slice() : null;
  const map = new Map();
  for (const t of base || []) map.set(String(t.pos), t);
  for (const t of over) {
    if (t.val === 'clear') map.delete(String(t.pos));
    else map.set(String(t.pos), t);
  }
  return [...map.values()].sort((a, b) => Number(a.pos) - Number(b.pos));
}

export function mergePPr(base, over) {
  if (!over) return { ...(base || {}) };
  const r = { ...(base || {}) };
  for (const k of Object.keys(over)) {
    const v = over[k];
    if (v == null) continue;
    switch (k) {
      case 'ind': r.ind = mergeInd(r.ind, v); break;
      case 'spacing': r.spacing = { ...(r.spacing || {}), ...v }; break;
      case 'pBdr': r.pBdr = { ...(r.pBdr || {}), ...v }; break;
      case 'tabs': r.tabs = mergeTabs(r.tabs, v); break;
      case 'extra': r.extra = [...(r.extra || []), ...v]; break;
      case 'numId': r.numId = v; if (over.ilvl == null) r.ilvl = r.ilvl ?? 0; break;
      default: r[k] = v;
    }
  }
  return r;
}

// Mapa tipo de marca -> atributos
export function runMap(list) {
  const m = {};
  for (const x of list || []) m[typeName(x)] = x.attrs;
  return m;
}

const THEME_FONT_KEYS = { ascii: 'asciiTheme', hAnsi: 'hAnsiTheme', eastAsia: 'eastAsiaTheme', cs: 'cstheme' };

export function mergeRun(base, over) {
  const r = { ...(base || {}) };
  for (const k of Object.keys(over || {})) {
    const v = over[k];
    if (k === 'rFonts' && r.rFonts) {
      const a = { ...(r.rFonts.a || {}) };
      for (const [key, val] of Object.entries(v.a || {})) {
        a[key] = val;
        if (THEME_FONT_KEYS[key] && !(v.a || {})[THEME_FONT_KEYS[key]]) delete a[THEME_FONT_KEYS[key]];
      }
      r.rFonts = { a };
    } else if (k === 'rprExtra' && r.rprExtra) {
      r.rprExtra = { items: [...(r.rprExtra.items || []), ...(v.items || [])] };
    } else {
      r[k] = v;
    }
  }
  return r;
}
