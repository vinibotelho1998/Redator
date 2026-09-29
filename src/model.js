// Modelo do documento: tudo que acompanha o conteúdo editável.
import { StyleSheet } from './docx/styles.js';
import { Numbering } from './docx/numbering.js';
import { Theme } from './docx/theme.js';
import { parseXml, serialize, children, child, wa, ra, onOff, NS } from './util/xml.js';

export const PAPER_SIZES = [
  { id: 'A4', name: 'A4', w: 11906, h: 16838, desc: '21 cm x 29,7 cm' },
  { id: 'Letter', name: 'Carta', w: 12240, h: 15840, desc: '21,59 cm x 27,94 cm' },
  { id: 'Legal', name: 'Ofício', w: 12240, h: 20160, desc: '21,59 cm x 35,56 cm' },
  { id: 'A3', name: 'A3', w: 16838, h: 23811, desc: '29,7 cm x 42 cm' },
  { id: 'A5', name: 'A5', w: 8391, h: 11906, desc: '14,8 cm x 21 cm' },
  { id: 'B5', name: 'B5 (JIS)', w: 10319, h: 14571, desc: '18,2 cm x 25,7 cm' },
  { id: 'Executive', name: 'Executivo', w: 10440, h: 15120, desc: '18,41 cm x 26,67 cm' },
  { id: 'Tabloid', name: 'Tabloide', w: 15840, h: 24480, desc: '27,94 cm x 43,18 cm' },
];

export const MARGIN_PRESETS = [
  { id: 'normal', name: 'Normal', top: 1417, bottom: 1417, left: 1701, right: 1701 },
  { id: 'estreita', name: 'Estreita', top: 720, bottom: 720, left: 720, right: 720 },
  { id: 'moderada', name: 'Moderada', top: 1440, bottom: 1440, left: 1080, right: 1080 },
  { id: 'larga', name: 'Larga', top: 1440, bottom: 1440, left: 2880, right: 2880 },
  { id: 'abnt', name: 'ABNT', top: 1701, bottom: 1134, left: 1701, right: 1134 },
  { id: 'office2003', name: 'Padrão Office 2003', top: 1417, bottom: 1417, left: 1800, right: 1800 },
];

const W_DECL = 'xmlns:w="' + NS.w + '" xmlns:r="' + NS.r + '"';

export function defaultSectionXml() {
  return '<w:sectPr ' + W_DECL + '><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1701" w:bottom="1417" w:left="1701" w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>';
}

export class Section {
  constructor(xml) {
    this.doc = parseXml(xml || defaultSectionXml());
    this.el = this.doc.documentElement;
    this.read();
  }

  read() {
    const el = this.el;
    const pgSz = child(el, 'pgSz');
    const pgMar = child(el, 'pgMar');
    const num = (e, n, d) => { const v = parseInt(wa(e, n), 10); return Number.isFinite(v) ? v : d; };
    this.pgSz = { w: num(pgSz, 'w', 11906), h: num(pgSz, 'h', 16838), orient: wa(pgSz, 'orient') || null };
    this.pgMar = {
      top: num(pgMar, 'top', 1417), right: num(pgMar, 'right', 1701), bottom: num(pgMar, 'bottom', 1417),
      left: num(pgMar, 'left', 1701), header: num(pgMar, 'header', 708), footer: num(pgMar, 'footer', 708), gutter: num(pgMar, 'gutter', 0),
    };
    this.headers = {};
    this.footers = {};
    for (const h of children(el, 'headerReference')) this.headers[wa(h, 'type') || 'default'] = ra(h, 'id');
    for (const f of children(el, 'footerReference')) this.footers[wa(f, 'type') || 'default'] = ra(f, 'id');
    this.titlePg = !!child(el, 'titlePg') && onOff(child(el, 'titlePg'));
    const pnt = child(el, 'pgNumType');
    this.pgNumType = { fmt: wa(pnt, 'fmt') || 'decimal', start: pnt && wa(pnt, 'start') != null ? parseInt(wa(pnt, 'start'), 10) : null };
    const cols = child(el, 'cols');
    this.cols = cols ? parseInt(wa(cols, 'num'), 10) || 1 : 1;
    this.type = wa(child(el, 'type'), 'val') || 'nextPage';
    this.vAlign = wa(child(el, 'vAlign'), 'val') || 'top';
  }

  ensureChild(name, before) {
    let c = child(this.el, name);
    if (!c) {
      c = this.doc.createElementNS(NS.w, 'w:' + name);
      const ref = before ? before.map((b) => child(this.el, b)).find(Boolean) : null;
      this.el.insertBefore(c, ref || null);
    }
    return c;
  }

  setPageSize(w, h, orient) {
    const c = this.ensureChild('pgSz', ['pgMar', 'paperSrc', 'pgBorders', 'lnNumType', 'pgNumType', 'cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg', 'textDirection', 'bidi', 'rtlGutter', 'docGrid', 'printerSettings']);
    c.setAttributeNS(NS.w, 'w:w', String(Math.round(w)));
    c.setAttributeNS(NS.w, 'w:h', String(Math.round(h)));
    if (orient === 'landscape') c.setAttributeNS(NS.w, 'w:orient', 'landscape');
    else c.removeAttributeNS(NS.w, 'orient');
    this.read();
  }

  setMargins(m) {
    const c = this.ensureChild('pgMar', ['paperSrc', 'pgBorders', 'lnNumType', 'pgNumType', 'cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg', 'textDirection', 'bidi', 'rtlGutter', 'docGrid', 'printerSettings']);
    const cur = { ...this.pgMar, ...m };
    for (const k of ['top', 'right', 'bottom', 'left', 'header', 'footer', 'gutter']) c.setAttributeNS(NS.w, 'w:' + k, String(Math.round(cur[k])));
    this.read();
  }

  setTitlePg(on) {
    const c = child(this.el, 'titlePg');
    if (on && !c) this.ensureChild('titlePg', ['textDirection', 'bidi', 'rtlGutter', 'docGrid', 'printerSettings']);
    else if (!on && c) c.remove();
    this.read();
  }

  setReference(kind, type, rId) {
    const tag = kind === 'header' ? 'headerReference' : 'footerReference';
    let ref = children(this.el, tag).find((r) => (wa(r, 'type') || 'default') === type);
    if (!ref) {
      ref = this.doc.createElementNS(NS.w, 'w:' + tag);
      ref.setAttributeNS(NS.w, 'w:type', type);
      // referências vêm primeiro, cabeçalhos antes de rodapés
      const firstFooter = child(this.el, 'footerReference');
      const firstOther = Array.from(this.el.children).find((c) => c.localName !== 'headerReference' && c.localName !== 'footerReference');
      const before = kind === 'header' ? (firstFooter || firstOther) : firstOther;
      this.el.insertBefore(ref, before || null);
    }
    ref.setAttributeNS(NS.r, 'r:id', rId);
    this.read();
  }

  setPageNumbering(fmt, start) {
    const c = this.ensureChild('pgNumType', ['cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg', 'textDirection', 'bidi', 'rtlGutter', 'docGrid', 'printerSettings']);
    if (fmt) c.setAttributeNS(NS.w, 'w:fmt', fmt); else c.removeAttributeNS(NS.w, 'fmt');
    if (start != null) c.setAttributeNS(NS.w, 'w:start', String(start)); else c.removeAttributeNS(NS.w, 'start');
    this.read();
  }

  toXml() {
    return serialize(this.el);
  }

  clone() {
    return new Section(this.toXml());
  }
}

export class DocModel {
  constructor(format = 'docx') {
    this.format = format;
    this.zip = null;
    this.mainPath = 'word/document.xml';
    this.rootAttrs = null;
    this.rels = new Map();
    this.styles = new StyleSheet();
    this.numbering = new Numbering();
    this.numbering.styles = this.styles;
    this.theme = new Theme();
    this.settings = { defaultTabStop: 708, evenAndOdd: false, compat: 15, decimalSymbol: ',' };
    this.section = new Section();
    this.headers = new Map();
    this.footers = new Map();
    this.footnotes = new Map();
    this.endnotes = new Map();
    this.comments = new Map();
    this.media = new Map();
    this.fontTable = new Map();
    this.properties = {};
    this.extra = {};
    this.mediaListeners = new Set();
    this.changedParts = new Set();
  }

  mediaUrl(path) {
    const m = path ? this.media.get(path) : null;
    return m ? m.url : null;
  }

  setMediaUrl(path, url) {
    const m = this.media.get(path);
    if (!m) return;
    m.url = url;
    for (const fn of this.mediaListeners) fn(path, url);
  }

  onMedia(fn) {
    this.mediaListeners.add(fn);
    return () => this.mediaListeners.delete(fn);
  }

  // Cabeçalho ou rodapé de uma página
  headerFor(kind, pageIndex, isFirstOfSection) {
    const s = this.section;
    const inherited = (this.inheritedRefs && this.inheritedRefs[kind === 'header' ? 'headers' : 'footers']) || {};
    const refs = { ...inherited, ...(kind === 'header' ? s.headers : s.footers) };
    const map = kind === 'header' ? this.headers : this.footers;
    let type = 'default';
    if (isFirstOfSection && s.titlePg) type = 'first';
    else if (this.settings.evenAndOdd && (pageIndex + 1) % 2 === 0) type = 'even';
    const rId = refs[type];
    if (type === 'first' && !rId) return null;
    if (type === 'even' && !rId) return null;
    return rId ? map.get(rId) || null : null;
  }
}
