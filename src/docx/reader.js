// Leitor de .docx: pacote OOXML para o documento do editor.
import JSZip from 'jszip';
import { Fragment } from 'prosemirror-model';
import { schema } from '../editor/schema.js';
import {
  NS, parseXml, serialize, isW, isNs, wa, ra, child, children, childNs, findDeep, findAllDeep,
  resolvePartPath, relsPathFor, plainText, onOff,
} from '../util/xml.js';
import { parsePPr, parseRPr, parseTblPr, parseTrPr, parseTcPr } from './props.js';
import { StyleSheet } from './styles.js';
import { Numbering } from './numbering.js';
import { parseTheme } from './theme.js';
import { DocModel, Section } from '../model.js';

export const MIME = {
  png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', jpe: 'image/jpeg', jfif: 'image/jpeg', gif: 'image/gif',
  bmp: 'image/bmp', webp: 'image/webp', svg: 'image/svg+xml', tif: 'image/tiff', tiff: 'image/tiff',
  emf: 'image/x-emf', wmf: 'image/x-wmf', ico: 'image/x-icon', wdp: 'image/vnd.ms-photo',
};
export const NEEDS_CONVERSION = new Set(['emf', 'wmf', 'tif', 'tiff', 'wdp']);

export function extOf(path) {
  const m = /\.([a-z0-9]+)$/i.exec(path || '');
  return m ? m[1].toLowerCase() : '';
}

export function parseRels(text, basePart) {
  const map = new Map();
  if (!text) return map;
  let doc;
  try { doc = parseXml(text); } catch { return map; }
  for (const r of doc.getElementsByTagName('*')) {
    if (r.localName !== 'Relationship') continue;
    const id = r.getAttribute('Id');
    const type = r.getAttribute('Type') || '';
    const target = r.getAttribute('Target') || '';
    const mode = r.getAttribute('TargetMode');
    map.set(id, { id, type, target, mode, path: mode === 'External' ? target : resolvePartPath(basePart, target) });
  }
  return map;
}

function rootAttrString(el) {
  let s = '';
  for (const at of el.attributes) s += ' ' + at.name + '="' + at.value.replace(/&/g, '&amp;').replace(/"/g, '&quot;') + '"';
  return s;
}

function parseSettings(text) {
  const s = { defaultTabStop: 708, evenAndOdd: false, compat: 15, decimalSymbol: ',', text };
  if (!text) return s;
  let doc;
  try { doc = parseXml(text); } catch { return s; }
  const root = doc.documentElement;
  const dts = child(root, 'defaultTabStop');
  if (dts) s.defaultTabStop = parseInt(wa(dts, 'val'), 10) || 708;
  const eo = child(root, 'evenAndOddHeaders');
  s.evenAndOdd = !!eo && onOff(eo);
  s.compat = 12;
  const compat = child(root, 'compat');
  if (compat) {
    for (const cs of children(compat, 'compatSetting')) {
      if (wa(cs, 'name') === 'compatibilityMode') s.compat = parseInt(wa(cs, 'val'), 10) || 12;
    }
  }
  const ds = child(root, 'decimalSymbol');
  if (ds) s.decimalSymbol = wa(ds, 'val') || ',';
  s.mirrorMargins = !!child(root, 'mirrorMargins');
  const trk = child(root, 'trackRevisions');
  s.trackRevisions = !!trk && onOff(trk);
  return s;
}

function parseFontTable(text) {
  const map = new Map();
  if (!text) return map;
  let doc;
  try { doc = parseXml(text); } catch { return map; }
  for (const f of children(doc.documentElement, 'font')) {
    const fam = child(f, 'family');
    const alt = child(f, 'altName');
    map.set(wa(f, 'name'), { family: fam ? wa(fam, 'val') : null, altName: alt ? wa(alt, 'val') : null });
  }
  return map;
}

function parseCore(text) {
  const p = {};
  if (!text) return p;
  let doc;
  try { doc = parseXml(text); } catch { return p; }
  for (const el of doc.documentElement.children) {
    const v = el.textContent;
    switch (el.localName) {
      case 'title': p.title = v; break;
      case 'subject': p.subject = v; break;
      case 'creator': p.author = v; break;
      case 'keywords': p.keywords = v; break;
      case 'description': p.description = v; break;
      case 'lastModifiedBy': p.lastModifiedBy = v; break;
      case 'revision': p.revision = v; break;
      case 'created': p.created = v; break;
      case 'modified': p.modified = v; break;
      default: break;
    }
  }
  return p;
}

// ---------------------------------------------------------------- contexto de conversão

class Ctx {
  constructor(model, rels, partPath) {
    this.model = model;
    this.rels = rels;
    this.partPath = partPath;
    this.pending = [];
  }
  sub() {
    const c = new Ctx(this.model, this.rels, this.partPath);
    return c;
  }
}

const mk = (list) => list.map((m) => schema.mark(m.type, m.attrs));

function addMark(marks, mark) {
  return mark.addToSet(marks);
}

function opaque(el, level, kind, text, marks, id) {
  return schema.nodes.opaque_inline.create({ xml: serialize(el), level, kind, text: text || '', id: id ?? null }, null, marks || null);
}

// ---------------------------------------------------------------- blocos

export function convertBlocks(els, ctx) {
  const out = [];
  for (const el of els) {
    if (el.nodeType !== 1) continue;
    if (isW(el, 'p')) out.push(convertParagraph(el, ctx));
    else if (isW(el, 'tbl')) out.push(convertTable(el, ctx));
    else if (isW(el, 'sdt')) {
      const content = child(el, 'sdtContent');
      const inner = content ? convertBlocks(Array.from(content.children), ctx) : [];
      const pr = child(el, 'sdtPr');
      const endPr = child(el, 'sdtEndPr');
      if (inner.length) out.push(schema.nodes.sdt_block.create({ pr: pr ? serialize(pr) : null, endPr: endPr ? serialize(endPr) : null }, inner));
    } else if (isW(el, 'customXml') || isW(el, 'smartTag')) {
      out.push(...convertBlocks(Array.from(el.children), ctx));
    } else if (isW(el, 'sectPr')) {
      // tratado por quem chamou
    } else if (isW(el, 'bookmarkStart') || isW(el, 'bookmarkEnd') || isW(el, 'commentRangeStart') || isW(el, 'commentRangeEnd')
      || isW(el, 'permStart') || isW(el, 'permEnd') || isW(el, 'moveFromRangeStart') || isW(el, 'moveFromRangeEnd')
      || isW(el, 'moveToRangeStart') || isW(el, 'moveToRangeEnd') || isW(el, 'proofErr')) {
      if (!isW(el, 'proofErr')) ctx.pending.push(opaque(el, 'p', el.localName, '', null, wa(el, 'id')));
    } else if (isW(el, 'ins') || isW(el, 'moveTo')) {
      out.push(...convertBlocks(Array.from(el.children), ctx));
    } else if (isW(el, 'del') || isW(el, 'moveFrom')) {
      // bloco excluído com controle de alterações: mantém como está
      out.push(schema.nodes.opaque_block.create({ xml: serialize(el), kind: 'exclusão', text: plainText(el) }));
    } else if (el.localName === 'AlternateContent') {
      const ch = Array.from(el.children).find((c) => c.localName === 'Choice') || Array.from(el.children).find((c) => c.localName === 'Fallback');
      if (ch) out.push(...convertBlocks(Array.from(ch.children), ctx));
    } else {
      out.push(schema.nodes.opaque_block.create({ xml: serialize(el), kind: el.localName, text: plainText(el).slice(0, 2000) }));
    }
  }
  return out;
}

export function convertParagraph(p, ctx) {
  const attrs = parsePPr(child(p, 'pPr'));
  const inl = [];
  if (ctx.pending.length) { inl.push(...ctx.pending); ctx.pending = []; }
  convertInline(p, [], inl, ctx);
  try {
    return schema.nodes.paragraph.create(attrs, inl);
  } catch (err) {
    console.warn('Parágrafo inválido', err);
    return schema.nodes.paragraph.create(attrs);
  }
}

function convertInline(parent, marks, out, ctx) {
  for (const el of parent.children) {
    if (!isW(el)) {
      if (isNs(el, NS.m, 'oMath') || isNs(el, NS.m, 'oMathPara')) {
        out.push(opaque(el, 'p', 'math', mathText(el), marks));
      } else if (el.localName === 'AlternateContent') {
        const ch = Array.from(el.children).find((c) => c.localName === 'Choice');
        const fb = Array.from(el.children).find((c) => c.localName === 'Fallback');
        const target = ch && Array.from(ch.children).some((c) => isW(c)) ? ch : fb;
        if (target && Array.from(target.children).every((c) => isW(c, 'r') || isW(c, 'hyperlink'))) convertInline(target, marks, out, ctx);
        else out.push(opaque(el, 'p', 'alternate', '', marks));
      } else {
        out.push(opaque(el, 'p', el.localName, '', marks));
      }
      continue;
    }
    switch (el.localName) {
      case 'pPr': break;
      case 'r': convertRun(el, marks, out, ctx); break;
      case 'hyperlink': {
        const rId = ra(el, 'id');
        const rel = rId ? ctx.rels.get(rId) : null;
        const a = {};
        for (const at of el.attributes) if (at.localName !== 'id' || at.namespaceURI !== NS.r) { if (at.prefix === 'w' || at.namespaceURI === NS.w) a[at.localName] = at.value; }
        const link = schema.marks.hyperlink.create({ href: rel ? rel.target : null, anchor: wa(el, 'anchor'), a });
        convertInline(el, addMark(marks, link), out, ctx);
        break;
      }
      case 'fldSimple': {
        const instr = wa(el, 'instr') || '';
        out.push(schema.nodes.field_char.create({ type: 'begin' }, null, marks));
        out.push(schema.nodes.instr_text.create({ text: instr }, null, marks));
        out.push(schema.nodes.field_char.create({ type: 'separate' }, null, marks));
        convertInline(el, marks, out, ctx);
        out.push(schema.nodes.field_char.create({ type: 'end' }, null, marks));
        break;
      }
      case 'ins': case 'moveTo': {
        const a = {};
        for (const at of el.attributes) if (at.namespaceURI === NS.w) a[at.localName] = at.value;
        const m = schema.marks.ins.create({ a, tag: el.localName });
        convertInline(el, addMark(marks, m), out, ctx);
        break;
      }
      case 'del': case 'moveFrom': {
        const a = el.getAttributeNS(NS.w, 'author');
        const node = opaque(el, 'p', 'del', plainText(el), marks);
        out.push(node.type.create({ ...node.attrs, id: a }, null, node.marks));
        break;
      }
      case 'sdt': {
        const content = child(el, 'sdtContent');
        const pr = child(el, 'sdtPr');
        const endPr = child(el, 'sdtEndPr');
        const m = schema.marks.sdt.create({ pr: pr ? serialize(pr) : null, endPr: endPr ? serialize(endPr) : null });
        const before = out.length;
        if (content) convertInline(content, addMark(marks, m), out, ctx);
        if (out.length === before) out.push(opaque(el, 'p', 'sdt', '', marks));
        break;
      }
      case 'smartTag': case 'customXml': case 'dir': case 'bdo':
        convertInline(el, marks, out, ctx);
        break;
      case 'proofErr': break;
      case 'bookmarkStart': case 'bookmarkEnd': case 'commentRangeStart': case 'commentRangeEnd':
      case 'permStart': case 'permEnd': case 'moveFromRangeStart': case 'moveFromRangeEnd': case 'moveToRangeStart': case 'moveToRangeEnd':
        out.push(opaque(el, 'p', el.localName, '', null, wa(el, 'id') ?? wa(el, 'name')));
        break;
      default:
        out.push(opaque(el, 'p', el.localName, plainText(el), marks));
    }
  }
}

function mathText(el) {
  let s = '';
  const walk = (n) => {
    for (const c of n.children) {
      if (c.localName === 't') s += c.textContent;
      else if (c.localName === 'f') { s += '('; walk(child0(c, 'num')); s += ')/('; walk(child0(c, 'den')); s += ')'; }
      else if (c.localName === 'sSup') { walk(child0(c, 'e')); s += '^'; walk(child0(c, 'sup')); }
      else if (c.localName === 'sSub') { walk(child0(c, 'e')); s += '_'; walk(child0(c, 'sub')); }
      else if (c.localName === 'rad') { s += '√('; walk(child0(c, 'e')); s += ')'; }
      else walk(c);
    }
  };
  const child0 = (n, name) => Array.from(n.children).find((c) => c.localName === name) || { children: [] };
  walk(el);
  return s;
}

function convertRun(r, outerMarks, out, ctx) {
  const rPr = child(r, 'rPr');
  let marks = outerMarks;
  for (const m of mk(parseRPr(rPr))) marks = m.addToSet(marks);
  let text = '';
  const flush = () => {
    if (text) { out.push(schema.text(text, marks)); text = ''; }
  };
  for (const el of r.children) {
    if (!isW(el)) {
      flush();
      if (el.localName === 'AlternateContent') out.push(convertAlternate(el, marks, ctx));
      else if (isNs(el, NS.m, 'oMath')) out.push(opaque(el, 'r', 'math', mathText(el), marks));
      else out.push(opaque(el, 'r', el.localName, '', marks));
      continue;
    }
    switch (el.localName) {
      case 'rPr': break;
      case 't': text += el.textContent; break;
      case 'delText': text += el.textContent; break;
      case 'tab': flush(); out.push(schema.nodes.tab.create(null, null, marks)); break;
      case 'ptab': flush(); out.push(schema.nodes.tab.create({ ptab: wa(el, 'alignment') }, null, marks)); break;
      case 'br': {
        flush();
        const type = wa(el, 'type');
        out.push(schema.nodes.hard_break.create({ type: type === 'page' || type === 'column' ? type : null, clear: wa(el, 'clear') }, null, marks));
        break;
      }
      case 'cr': flush(); out.push(schema.nodes.hard_break.create(null, null, marks)); break;
      case 'noBreakHyphen': text += '‑'; break;
      case 'softHyphen': text += '­'; break;
      case 'sym': flush(); out.push(schema.nodes.sym.create({ font: wa(el, 'font'), char: wa(el, 'char') }, null, marks)); break;
      case 'drawing': flush(); out.push(convertDrawing(el, marks, ctx, serialize(el))); break;
      case 'pict': flush(); out.push(convertPict(el, marks, ctx, serialize(el))); break;
      case 'object': flush(); out.push(convertPict(el, marks, ctx, serialize(el), 'object')); break;
      case 'fldChar': {
        flush();
        const ff = child(el, 'ffData');
        out.push(schema.nodes.field_char.create({ type: wa(el, 'fldCharType') || 'begin', xml: ff || el.attributes.length > 1 ? serialize(el) : null }, null, marks));
        break;
      }
      case 'instrText': flush(); out.push(schema.nodes.instr_text.create({ text: el.textContent }, null, marks)); break;
      case 'delInstrText': flush(); out.push(schema.nodes.instr_text.create({ text: el.textContent }, null, marks)); break;
      case 'footnoteReference': case 'endnoteReference':
        flush();
        out.push(schema.nodes.footnote_ref.create({
          id: wa(el, 'id'),
          kind: el.localName === 'footnoteReference' ? 'footnote' : 'endnote',
          custom: wa(el, 'customMarkFollows') === '1' ? '1' : null,
        }, null, marks));
        break;
      case 'footnoteRef': case 'endnoteRef':
        flush(); out.push(opaque(el, 'r', 'noteMark', '', marks)); break;
      case 'commentReference':
        flush(); out.push(opaque(el, 'r', 'commentRef', '', marks, wa(el, 'id'))); break;
      case 'lastRenderedPageBreak': break;
      case 'separator': case 'continuationSeparator':
        flush(); out.push(opaque(el, 'r', el.localName, '', marks)); break;
      default:
        flush(); out.push(opaque(el, 'r', el.localName, '', marks));
    }
  }
  flush();
}

// ---------------------------------------------------------------- desenhos

function emu(v) {
  const n = parseInt(v, 10);
  return Number.isFinite(n) ? n : 0;
}

function parseAnchor(el) {
  const a = {
    behind: el.getAttribute('behindDoc') === '1' || el.getAttribute('behindDoc') === 'true',
    wrap: 'none', wrapText: null, relH: 'column', alignH: null, offH: 0, relV: 'paragraph', alignV: null, offV: 0,
    distL: emu(el.getAttribute('distL')), distR: emu(el.getAttribute('distR')), distT: emu(el.getAttribute('distT')), distB: emu(el.getAttribute('distB')),
    z: emu(el.getAttribute('relativeHeight')),
  };
  for (const c of el.children) {
    switch (c.localName) {
      case 'positionH': case 'positionV': {
        const h = c.localName === 'positionH';
        const al = childNs(c, NS.wp, 'align');
        const po = childNs(c, NS.wp, 'posOffset');
        const pct = Array.from(c.children).find((x) => x.localName === 'pctPosHOffset' || x.localName === 'pctPosVOffset');
        if (h) { a.relH = c.getAttribute('relativeFrom'); a.alignH = al ? al.textContent.trim() : null; a.offH = po ? emu(po.textContent) : 0; if (pct) a.pctH = emu(pct.textContent); }
        else { a.relV = c.getAttribute('relativeFrom'); a.alignV = al ? al.textContent.trim() : null; a.offV = po ? emu(po.textContent) : 0; if (pct) a.pctV = emu(pct.textContent); }
        break;
      }
      case 'wrapNone': a.wrap = 'none'; break;
      case 'wrapSquare': a.wrap = 'square'; a.wrapText = c.getAttribute('wrapText'); break;
      case 'wrapTight': a.wrap = 'tight'; a.wrapText = c.getAttribute('wrapText'); break;
      case 'wrapThrough': a.wrap = 'through'; a.wrapText = c.getAttribute('wrapText'); break;
      case 'wrapTopAndBottom': a.wrap = 'topAndBottom'; break;
      default: break;
    }
  }
  return a;
}

function colorFromSp(el) {
  if (!el) return null;
  const srgb = el.getElementsByTagNameNS(NS.a, 'srgbClr')[0];
  if (srgb) return '#' + srgb.getAttribute('val');
  const scheme = el.getElementsByTagNameNS(NS.a, 'schemeClr')[0];
  if (scheme) return 'theme:' + scheme.getAttribute('val');
  return null;
}

function shapeStyle(container) {
  const spPr = findDeep(container, 'spPr', NS.wps) || findDeep(container, 'spPr', NS.pic);
  if (!spPr) return { fill: null, line: null };
  let fill = null;
  let line = null;
  for (const c of spPr.children) {
    if (c.localName === 'solidFill') fill = colorFromSp(c);
    else if (c.localName === 'noFill') fill = 'none';
    else if (c.localName === 'ln') {
      if (Array.from(c.children).some((x) => x.localName === 'noFill')) line = 'none';
      else {
        const sf = Array.from(c.children).find((x) => x.localName === 'solidFill');
        line = { color: sf ? colorFromSp(sf) : '#000000', w: emu(c.getAttribute('w')) || 9525 };
      }
    }
  }
  return { fill, line };
}

function convertDrawing(drawingEl, marks, ctx, xml) {
  const container = drawingEl.firstElementChild;
  if (!container) return schema.nodes.opaque_inline.create({ xml, level: 'r', kind: 'drawing' }, null, marks);
  const isAnchor = container.localName === 'anchor';
  const extent = childNs(container, NS.wp, 'extent');
  const cx = extent ? emu(extent.getAttribute('cx')) : 0;
  const cy = extent ? emu(extent.getAttribute('cy')) : 0;
  const docPr = childNs(container, NS.wp, 'docPr');
  const graphicData = findDeep(container, 'graphicData', NS.a);
  const uri = graphicData ? graphicData.getAttribute('uri') || '' : '';
  const attrs = {
    kind: 'other', cx, cy, ocx: cx, ocy: cy,
    alt: docPr ? docPr.getAttribute('descr') || '' : '', title: docPr ? docPr.getAttribute('title') || '' : '', name: docPr ? docPr.getAttribute('name') || '' : '',
    xml, anchor: isAnchor ? parseAnchor(container) : null, srcPart: ctx.partPath,
  };
  const blip = graphicData ? findDeep(graphicData, 'blip', NS.a) : null;
  if (uri === NS.pic || (blip && !/wordprocessingShape|wordprocessingGroup|chart|diagram/.test(uri))) {
    attrs.kind = 'image';
    const rId = blip ? ra(blip, 'embed') || ra(blip, 'link') : null;
    const rel = rId ? ctx.rels.get(rId) : null;
    attrs.rId = rId;
    if (rel && rel.mode === 'External') attrs.src = rel.target;
    else if (rel) { attrs.mediaPath = rel.path; attrs.ext = extOf(rel.path); }
    const sr = findDeep(graphicData, 'srcRect', NS.a);
    if (sr && (sr.getAttribute('l') || sr.getAttribute('t') || sr.getAttribute('r') || sr.getAttribute('b'))) {
      attrs.crop = { l: emu(sr.getAttribute('l')), t: emu(sr.getAttribute('t')), r: emu(sr.getAttribute('r')), b: emu(sr.getAttribute('b')) };
    }
  } else if (/wordprocessingShape|wordprocessingGroup|wordprocessingCanvas/.test(uri)) {
    attrs.kind = /Group/.test(uri) ? 'group' : /Canvas/.test(uri) ? 'canvas' : 'shape';
    const boxes = findAllDeep(graphicData, 'txbxContent', NS.w);
    if (boxes.length) {
      const blocks = [];
      for (const b of boxes) blocks.push(...convertBlocks(Array.from(b.children), ctx.sub()));
      attrs.content = blocks.map((n) => n.toJSON());
    }
    const st = shapeStyle(graphicData);
    attrs.fill = st.fill;
    attrs.line = st.line;
    if (!boxes.length && blip) {
      const rId = ra(blip, 'embed');
      const rel = rId ? ctx.rels.get(rId) : null;
      if (rel && rel.mode !== 'External') { attrs.mediaPath = rel.path; attrs.ext = extOf(rel.path); }
    }
  } else if (/chart/.test(uri)) {
    attrs.kind = 'chart';
  } else if (/diagram/.test(uri)) {
    attrs.kind = 'smartart';
  }
  return schema.nodes.drawing.create(attrs, null, marks);
}

function cssLenToEmu(v) {
  const m = /(-?[\d.]+)\s*(pt|in|px|cm|mm)?/.exec(v || '');
  if (!m) return 0;
  const n = parseFloat(m[1]);
  switch (m[2] || 'px') {
    case 'pt': return Math.round(n * 12700);
    case 'in': return Math.round(n * 914400);
    case 'cm': return Math.round(n * 360000);
    case 'mm': return Math.round(n * 36000);
    default: return Math.round(n * 9525);
  }
}

function convertPict(el, marks, ctx, xml, kind) {
  const shape = findDeep(el, 'shape', NS.v) || findDeep(el, 'rect', NS.v) || findDeep(el, 'roundrect', NS.v) || findDeep(el, 'group', NS.v);
  const style = shape ? shape.getAttribute('style') || '' : '';
  const get = (k) => { const m = new RegExp('(?:^|;)\\s*' + k + '\\s*:\\s*([^;]+)').exec(style); return m ? m[1].trim() : null; };
  const cx = cssLenToEmu(get('width'));
  const cy = cssLenToEmu(get('height'));
  const attrs = { kind: kind || 'shape', cx, cy, ocx: cx, ocy: cy, xml, srcPart: ctx.partPath };
  const imagedata = findDeep(el, 'imagedata', NS.v);
  if (imagedata) {
    const rId = imagedata.getAttributeNS(NS.r, 'id') || imagedata.getAttributeNS(NS.r, 'pict') || imagedata.getAttributeNS(NS.o, 'relid');
    const rel = rId ? ctx.rels.get(rId) : null;
    if (rel && rel.mode !== 'External') { attrs.mediaPath = rel.path; attrs.ext = extOf(rel.path); attrs.kind = kind === 'object' ? 'object' : 'image'; }
  }
  const box = findDeep(el, 'txbxContent', NS.w);
  if (box) attrs.content = convertBlocks(Array.from(box.children), ctx.sub()).map((n) => n.toJSON());
  const pos = get('position');
  if (pos === 'absolute') {
    attrs.anchor = {
      behind: /z-index:\s*-/.test(style), wrap: 'none', relH: 'column', offH: cssLenToEmu(get('margin-left')), relV: 'paragraph', offV: cssLenToEmu(get('margin-top')),
      alignH: get('mso-position-horizontal'), alignV: get('mso-position-vertical'), distL: 0, distR: 0, distT: 0, distB: 0,
    };
  }
  if (shape && shape.getAttributeNS(NS.o, 'hr') === 't') { attrs.kind = 'hr'; attrs.cy = attrs.cy || 19050; }
  return schema.nodes.drawing.create(attrs, null, marks);
}

function convertAlternate(el, marks, ctx) {
  const xml = serialize(el);
  const choice = Array.from(el.children).find((c) => c.localName === 'Choice');
  const fallback = Array.from(el.children).find((c) => c.localName === 'Fallback');
  const drawing = choice && Array.from(choice.children).find((c) => isW(c, 'drawing'));
  if (drawing) return convertDrawing(drawing, marks, ctx, xml);
  const pict = fallback && Array.from(fallback.children).find((c) => isW(c, 'pict') || isW(c, 'object'));
  if (pict) return convertPict(pict, marks, ctx, xml);
  return schema.nodes.opaque_inline.create({ xml, level: 'r', kind: 'alternate' }, null, marks);
}

// ---------------------------------------------------------------- tabelas

function cellElements(tr) {
  const out = [];
  const walk = (parent) => {
    for (const c of parent.children) {
      if (isW(c, 'tc')) out.push(c);
      else if (isW(c, 'sdt')) { const content = child(c, 'sdtContent'); if (content) walk(content); }
      else if (isW(c, 'customXml')) walk(c);
    }
  };
  walk(tr);
  return out;
}

function rowElements(tbl) {
  const out = [];
  const walk = (parent) => {
    for (const c of parent.children) {
      if (isW(c, 'tr')) out.push(c);
      else if (isW(c, 'sdt')) { const content = child(c, 'sdtContent'); if (content) walk(content); }
      else if (isW(c, 'customXml')) walk(c);
    }
  };
  walk(tbl);
  return out;
}

function emptyPara() {
  return schema.nodes.paragraph.create();
}

export function convertTable(tbl, ctx) {
  const tattrs = parseTblPr(child(tbl, 'tblPr'));
  const gridEl = child(tbl, 'tblGrid');
  const grid = gridEl ? children(gridEl, 'gridCol').map((g) => parseInt(wa(g, 'w'), 10) || 0) : [];
  const rows = rowElements(tbl).map((tr) => {
    const rattrs = parseTrPr(child(tr, 'trPr'));
    const before = rattrs.gridBefore || 0;
    let col = before;
    const cells = cellElements(tr).map((tc) => {
      const cattrs = parseTcPr(child(tc, 'tcPr'));
      const span = cattrs.gridSpan || 1;
      const c = { tc, col, span, vMerge: cattrs.vMerge || null, cattrs, rowspan: 1, merged: false };
      col += span;
      return c;
    });
    return { tr, rattrs, cells, before, end: col };
  });
  // vMerge para rowspan
  for (let r = 0; r < rows.length; r++) {
    for (const c of rows[r].cells) {
      if (c.merged || c.vMerge !== 'restart') continue;
      let n = 1;
      for (let r2 = r + 1; r2 < rows.length; r2++) {
        const below = rows[r2].cells.find((x) => x.col === c.col && x.span === c.span && x.vMerge === 'continue' && !x.merged);
        if (!below) break;
        below.merged = true;
        n++;
      }
      c.rowspan = n;
    }
  }
  let width = grid.length;
  for (const r of rows) width = Math.max(width, r.end);
  while (grid.length < width) grid.push(grid.length ? grid[grid.length - 1] : 1440);
  const px = (tw) => Math.max(10, Math.round((tw / 15) * 100) / 100);
  // colunas cobertas por células mescladas verticalmente de linhas acima
  const above = rows.map(() => new Array(width).fill(false));
  for (let r = 0; r < rows.length; r++) {
    for (const c of rows[r].cells) {
      if (c.merged || c.rowspan <= 1) continue;
      for (let rr = r + 1; rr < r + c.rowspan && rr < rows.length; rr++) for (let k = c.col; k < c.col + c.span && k < width; k++) above[rr][k] = true;
    }
  }
  const phantom = (from, to) => schema.nodes.table_cell.create({ colspan: to - from, colwidth: grid.slice(from, to).map(px), phantom: true }, [emptyPara()]);
  const pmRows = [];
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    const starts = new Map();
    for (const c of row.cells) if (!c.merged) starts.set(c.col, c);
    const cells = [];
    let gap = null;
    const flushGap = (end) => { if (gap != null && end > gap) cells.push(phantom(gap, end)); gap = null; };
    for (let k = 0; k < width;) {
      const c = starts.get(k);
      if (c) {
        flushGap(k);
        const content = convertBlocks(Array.from(c.tc.children).filter((x) => !isW(x, 'tcPr')), ctx);
        if (!content.length) content.push(emptyPara());
        const { gridSpan, vMerge, hMerge, ...rest } = c.cattrs;
        cells.push(schema.nodes.table_cell.create({
          ...rest, colspan: c.span, rowspan: c.rowspan, colwidth: grid.slice(c.col, c.col + c.span).map(px),
        }, content));
        k += c.span;
        continue;
      }
      if (above[r][k]) { flushGap(k); k++; continue; }
      if (gap == null) gap = k;
      k++;
    }
    flushGap(width);
    const { gridBefore, gridAfter, ...rrest } = row.rattrs;
    if (!cells.length) cells.push(schema.nodes.table_cell.create({ colspan: 1, colwidth: [px(grid[0] || 1440)] }, [emptyPara()]));
    pmRows.push(schema.nodes.table_row.create({ ...rrest, gridBefore: gridBefore || null, gridAfter: gridAfter || null }, cells.filter(Boolean)));
  }
  if (!pmRows.length) pmRows.push(schema.nodes.table_row.create(null, [schema.nodes.table_cell.create({ colwidth: [px(grid[0] || 1440)] }, [emptyPara()])]));
  return schema.nodes.table.create({ ...tattrs, grid }, pmRows);
}

// ---------------------------------------------------------------- mídia

async function loadMedia(zip, model, rels) {
  const jobs = [];
  for (const rel of rels.values()) {
    if (!rel.type.endsWith('/image') || rel.mode === 'External') continue;
    if (model.media.has(rel.path)) continue;
    const f = zip.file(rel.path);
    if (!f) continue;
    const ext = extOf(rel.path);
    const entry = { url: null, bytes: null, ext, type: MIME[ext] || 'application/octet-stream', pending: false };
    model.media.set(rel.path, entry);
    jobs.push(f.async('uint8array').then((bytes) => {
      entry.bytes = bytes;
      if (NEEDS_CONVERSION.has(ext)) {
        entry.pending = true;
        entry.url = placeholderImage(ext);
      } else {
        entry.url = URL.createObjectURL(new Blob([bytes], { type: entry.type }));
      }
    }));
  }
  await Promise.all(jobs);
}

export function placeholderImage(ext) {
  const svg = '<svg xmlns="http://www.w3.org/2000/svg" width="120" height="80" viewBox="0 0 120 80"><rect x="0.5" y="0.5" width="119" height="79" fill="#f3f3f3" stroke="#c8c8c8"/><path d="M40 55l14-16 10 11 8-8 12 13z" fill="#bdbdbd"/><circle cx="46" cy="30" r="5" fill="#bdbdbd"/><text x="60" y="72" font-family="Segoe UI" font-size="9" text-anchor="middle" fill="#888">' + String(ext || '').toUpperCase() + '</text></svg>';
  return 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg);
}

// Converte EMF, WMF e TIFF em PNG com a ajuda do processo principal
export async function convertPendingMedia(model, host) {
  if (!host || !host.convertImage) return;
  for (const [path, entry] of model.media) {
    if (!entry.pending || !entry.bytes) continue;
    entry.pending = false;
    try {
      const png = await host.convertImage(entry.bytes, entry.ext, 0, 0);
      if (png && png.length) model.setMediaUrl(path, URL.createObjectURL(new Blob([png], { type: 'image/png' })));
    } catch (err) {
      console.warn('Conversão de imagem falhou', path, err);
    }
  }
}

// ---------------------------------------------------------------- partes auxiliares

async function loadPartDoc(zip, model, path) {
  const f = zip.file(path);
  if (!f) return null;
  const text = await f.async('string');
  const xml = parseXml(text);
  const rels = parseRels(await (zip.file(relsPathFor(path)) ? zip.file(relsPathFor(path)).async('string') : Promise.resolve(null)), path);
  await loadMedia(zip, model, rels);
  return { xml, rels, text };
}

async function loadHeaderFooter(zip, model, rId, kind) {
  const rel = model.rels.get(rId);
  if (!rel) return null;
  const part = await loadPartDoc(zip, model, rel.path);
  if (!part) return null;
  const ctx = new Ctx(model, part.rels, rel.path);
  const blocks = convertBlocks(Array.from(part.xml.documentElement.children), ctx);
  const doc = schema.node('doc', null, blocks.length ? blocks : [emptyPara()]);
  return { rId, kind, path: rel.path, doc, rootAttrs: rootAttrString(part.xml.documentElement), rels: part.rels, dirty: false };
}

async function loadNotes(zip, model, suffix, tag) {
  let rel = null;
  for (const r of model.rels.values()) if (r.type.endsWith(suffix)) rel = r;
  const map = new Map();
  if (!rel) return map;
  const part = await loadPartDoc(zip, model, rel.path);
  if (!part) return map;
  const ctx = new Ctx(model, part.rels, rel.path);
  for (const el of children(part.xml.documentElement, tag)) {
    const type = wa(el, 'type');
    if (type && type !== 'normal') continue;
    const blocks = convertBlocks(Array.from(el.children), ctx);
    map.set(wa(el, 'id'), { id: wa(el, 'id'), doc: schema.node('doc', null, blocks.length ? blocks : [emptyPara()]) });
  }
  return map;
}

async function loadComments(zip, model) {
  const map = new Map();
  let rel = null;
  for (const r of model.rels.values()) if (r.type.endsWith('/comments')) rel = r;
  if (!rel) return map;
  const f = zip.file(rel.path);
  if (!f) return map;
  const xml = parseXml(await f.async('string'));
  for (const el of children(xml.documentElement, 'comment')) {
    map.set(wa(el, 'id'), { id: wa(el, 'id'), author: wa(el, 'author') || '', initials: wa(el, 'initials') || '', date: wa(el, 'date') || '', text: plainText(el).trim() });
  }
  return map;
}

// ---------------------------------------------------------------- entrada principal

export async function readDocx(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const model = new DocModel('docx');
  model.zip = zip;
  const text = async (p) => { if (!p) return null; const f = zip.file(p); return f ? f.async('string') : null; };
  const rootRels = parseRels(await text('_rels/.rels'), '');
  let main = null;
  for (const r of rootRels.values()) if (r.type.endsWith('/officeDocument')) main = r.path;
  if (!main || !zip.file(main)) main = zip.file('word/document.xml') ? 'word/document.xml' : null;
  if (!main) throw new Error('O arquivo não parece ser um documento do Word.');
  model.mainPath = main;
  model.rels = parseRels(await text(relsPathFor(main)), main);
  const partOf = (suffix) => {
    for (const r of model.rels.values()) if (r.type.endsWith(suffix) && r.mode !== 'External') return r.path;
    return null;
  };
  const [stylesX, numberingX, themeX, settingsX, fontX, coreX] = await Promise.all([
    text(partOf('/styles')), text(partOf('/numbering')), text(partOf('/theme')), text(partOf('/settings')), text(partOf('/fontTable')), text('docProps/core.xml'),
  ]);
  model.theme = parseTheme(themeX);
  model.styles = StyleSheet.parse(stylesX);
  model.numbering = Numbering.parse(numberingX);
  model.numbering.styles = model.styles;
  model.settings = parseSettings(settingsX);
  model.fontTable = parseFontTable(fontX);
  model.properties = parseCore(coreX);
  await loadMedia(zip, model, model.rels);

  const docXml = parseXml(await text(main));
  const root = docXml.documentElement;
  model.rootAttrs = rootAttrString(root);
  const body = child(root, 'body');
  const ctx = new Ctx(model, model.rels, main);
  const blocks = body ? convertBlocks(Array.from(body.children), ctx) : [];
  if (ctx.pending.length) {
    const lastIdx = blocks.map((b) => b.type.name).lastIndexOf('paragraph');
    if (lastIdx >= 0) {
      const last = blocks[lastIdx];
      blocks[lastIdx] = last.copy(last.content.append(Fragment.from(ctx.pending)));
    }
    ctx.pending = [];
  }
  const sectEl = body ? children(body, 'sectPr').pop() : null;
  model.section = new Section(sectEl ? serialize(sectEl) : null);

  // referências herdadas de seções anteriores
  const effH = {};
  const effF = {};
  if (body) {
    for (const s of body.getElementsByTagNameNS(NS.w, 'sectPr')) {
      for (const h of children(s, 'headerReference')) effH[wa(h, 'type') || 'default'] = ra(h, 'id');
      for (const f of children(s, 'footerReference')) effF[wa(f, 'type') || 'default'] = ra(f, 'id');
    }
  }
  model.inheritedRefs = { headers: effH, footers: effF };
  const hdrIds = new Set([...Object.values(model.section.headers), ...Object.values(effH)]);
  const ftrIds = new Set([...Object.values(model.section.footers), ...Object.values(effF)]);
  for (const id of hdrIds) { const h = await loadHeaderFooter(zip, model, id, 'header'); if (h) model.headers.set(id, h); }
  for (const id of ftrIds) { const f = await loadHeaderFooter(zip, model, id, 'footer'); if (f) model.footers.set(id, f); }

  model.footnotes = await loadNotes(zip, model, '/footnotes', 'footnote');
  model.endnotes = await loadNotes(zip, model, '/endnotes', 'endnote');
  model.comments = await loadComments(zip, model);

  const doc = schema.node('doc', null, blocks.length ? blocks : [emptyPara()]);
  return { model, doc };
}
