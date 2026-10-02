// OpenDocument Text (.odt): leitura e escrita.
import JSZip from 'jszip';
import { Node as PMNode } from 'prosemirror-model';
import { schema } from '../editor/schema.js';
import { R, effectiveNumPr, fontName, colorOf, shdColor, borderCss } from '../editor/render.js';
import { parseXml } from '../util/xml.js';
import { MIME } from '../docx/reader.js';
import { newHeaderFooter } from '../docx/writer.js';
import { blankModel, finalizeDoc, contentWidthPx } from './common.js';
import { cmToTwip, TWIP_PER_CM } from '../util/units.js';

const O = {
  office: 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
  style: 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
  text: 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
  table: 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
  draw: 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0',
  fo: 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
  xlink: 'http://www.w3.org/1999/xlink',
  svg: 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0',
  dc: 'http://purl.org/dc/elements/1.1/',
  meta: 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0',
  manifest: 'urn:oasis:names:tc:opendocument:xmlns:manifest:1.0',
};

const nodes = schema.nodes;
const marks = schema.marks;

const at = (el, ns, name) => (el && el.getAttributeNS(O[ns], name)) || null;
const kids = (el) => (el ? Array.from(el.children) : []);
const is = (el, ns, name) => el && el.namespaceURI === O[ns] && el.localName === name;

// Uma abertura sem edição conserva o pacote completo, inclusive partes que o
// editor ainda não interpreta. Mudanças no conteúdo ou nos recursos invalidam
// essa cópia; nesse caso os diagnósticos explicam as limitações antes de salvar.
function documentSnapshot(doc) {
  const json = doc.toJSON();
  if (json.attrs) {
    json.attrs = { ...json.attrs };
    delete json.attrs.pageProperties; // atributo de histórico, espelhado no modelo
    if (!Object.keys(json.attrs).length) delete json.attrs;
  }
  return json;
}

function resourceSnapshot(model) {
  const { dirty: _dirty, ...settings } = model.settings;
  const entries = (map) => [...map].map(([id, entry]) => [id, {
    doc: entry.doc ? documentSnapshot(entry.doc) : null,
    xml: entry.xml || null, type: entry.type || null,
    author: entry.author || null, date: entry.date || null,
  }]);
  return JSON.stringify({
    section: model.section.toXml(), settings,
    properties: model.properties, theme: model.theme,
    styles: model.styles.toXml(), numbering: model.numbering.toXml(),
    headers: entries(model.headers), footers: entries(model.footers),
    footnotes: entries(model.footnotes), endnotes: entries(model.endnotes),
    comments: entries(model.comments),
    media: [...model.media].map(([path, media]) => [path, media.ext, media.type]),
  });
}

export function bindOdtSource(model, doc) {
  const source = model.extra && model.extra.odt;
  if (!source) return;
  source.document = JSON.stringify(documentSnapshot(doc));
  source.resources = resourceSnapshot(model);
  source.media = new Map([...model.media].map(([path, media]) => [path, media.bytes && media.bytes.slice()]));
}

export function canPreserveOdtSource(model, doc, opts = {}) {
  const source = model.extra && model.extra.odt;
  if (!source || !source.bytes || source.flat || source.template !== !!opts.template) return false;
  if (source.document !== JSON.stringify(documentSnapshot(doc)) || source.resources !== resourceSnapshot(model)) return false;
  for (const [path, bytes] of source.media) {
    const current = model.media.get(path)?.bytes;
    if (!!bytes !== !!current || (bytes && (bytes.length !== current.length || bytes.some((v, i) => v !== current[i])))) return false;
  }
  return true;
}

function odtDiagnostics(content, stylesDoc, meta, S, zip) {
  const out = [];
  const add = (code, message) => out.push({ code, message });
  const docs = [content, stylesDoc].filter(Boolean);
  const has = (ns, names) => docs.some((doc) => names.some((name) => doc.getElementsByTagNameNS(O[ns], name).length));
  if (S.masters.size > 1) add('page-masters', 'O documento contém vários estilos de página; a edição preserva apenas o primeiro layout, cabeçalho e rodapé importados.');
  if (has('style', ['header-left', 'header-first', 'footer-left', 'footer-first'])) add('page-variants', 'Cabeçalhos e rodapés diferentes na primeira página ou nas páginas pares não são preservados.');
  if (has('office', ['annotation'])) add('comments', 'Comentários do documento original não foram importados e serão removidos na gravação.');
  if (has('draw', ['object', 'object-ole', 'applet', 'plugin', 'custom-shape', 'rect', 'ellipse', 'line', 'polygon', 'polyline', 'path', 'g', 'connector'])) add('objects', 'Objetos incorporados, gráficos ou formas do documento original não são preservados.');
  if (docs.some((doc) => [...doc.getElementsByTagNameNS(O.draw, 'frame')].some((el) => !kids(el).some((c) => is(c, 'draw', 'image') || is(c, 'draw', 'text-box'))))) {
    if (!out.some((d) => d.code === 'objects')) add('objects', 'Objetos do documento original não puderam ser importados e serão removidos na gravação.');
  }
  if (docs.some((doc) => [...doc.getElementsByTagNameNS(O.draw, 'image')].some((el) => !zip.file((at(el, 'xlink', 'href') || '').replace(/^\.\//, ''))))) add('images', 'Há imagens externas ou ausentes no pacote que não puderam ser importadas.');
  if (docs.some((doc) => [...doc.getElementsByTagNameNS(O.draw, 'frame')].some((el) => { const anchor = at(el, 'text', 'anchor-type'); return anchor && anchor !== 'as-char'; }))) add('image-position', 'Imagens e caixas flutuantes são convertidas em objetos alinhados ao texto.');
  if (has('text', ['tracked-changes', 'change', 'change-start', 'change-end'])) add('changes', 'O histórico de alterações controladas do documento original não é preservado.');
  if (has('text', ['bookmark', 'bookmark-start', 'bookmark-end', 'reference-mark', 'reference-mark-start'])) add('bookmarks', 'Marcadores e destinos de referências internas do documento original não são preservados.');
  if (meta) {
    const container = meta.getElementsByTagNameNS(O.office, 'meta')[0];
    const kept = (el) => is(el, 'dc', 'title') || is(el, 'dc', 'creator') || is(el, 'meta', 'initial-creator') || is(el, 'meta', 'generator') || is(el, 'dc', 'date');
    if (kids(container).some((el) => !kept(el))) add('metadata', 'Metadados adicionais, como propriedades personalizadas, descrição e histórico de edição, não são preservados.');
  }
  if (Object.keys(zip.files).some((path) => /^META-INF\/(?:documentsignatures|macrosignatures)\.xml$/i.test(path))) add('signatures', 'Assinaturas digitais do documento original serão removidas após a edição.');
  return out;
}

// comprimento ODF para twips
function lenTw(v) {
  if (v == null) return null;
  const m = /^(-?[\d.]+)\s*(cm|mm|in|pt|pc|px)?$/.exec(String(v).trim());
  if (!m) return null;
  const n = parseFloat(m[1]);
  switch (m[2] || 'cm') {
    case 'cm': return Math.round(n * TWIP_PER_CM);
    case 'mm': return Math.round((n / 10) * TWIP_PER_CM);
    case 'in': return Math.round(n * 1440);
    case 'pt': return Math.round(n * 20);
    case 'pc': return Math.round(n * 240);
    case 'px': return Math.round(n * 15);
    default: return null;
  }
}

function hexColor(v) {
  if (!v || v === 'transparent') return null;
  const m = /^#([0-9a-f]{6})$/i.exec(v.trim());
  return m ? m[1].toUpperCase() : null;
}

function odfBorder(v) {
  if (!v || v === 'none') return { val: 'nil' };
  const w = /([\d.]+)(pt|cm|mm|in|px)/.exec(v);
  const c = /#([0-9a-f]{6})/i.exec(v);
  const tw = w ? lenTw(w[1] + w[2]) : 10;
  const style = /double/.test(v) ? 'double' : /dotted/.test(v) ? 'dotted' : /dashed/.test(v) ? 'dashed' : 'single';
  return { val: style, sz: String(Math.max(2, Math.round((tw / 20) * 8))), space: '1', color: c ? c[1].toUpperCase() : 'auto' };
}

// ---------------------------------------------------------------- estilos ODF

class OdfStyles {
  constructor(contentDoc, stylesDoc) {
    this.map = new Map();
    this.defaults = new Map();
    this.lists = new Map();
    this.fonts = new Map();
    this.pageLayouts = new Map();
    this.masters = new Map();
    const scan = (doc, isAuto) => {
      if (!doc) return;
      for (const ff of doc.getElementsByTagNameNS(O.style, 'font-face')) {
        this.fonts.set(at(ff, 'style', 'name'), (at(ff, 'svg', 'font-family') || at(ff, 'style', 'name') || '').replace(/^['"]|['"]$/g, ''));
      }
      for (const container of [...doc.getElementsByTagNameNS(O.office, 'styles'), ...doc.getElementsByTagNameNS(O.office, 'automatic-styles')]) {
        for (const s of kids(container)) {
          if (is(s, 'style', 'style')) this.map.set(at(s, 'style', 'name'), this.readStyle(s, container.localName === 'automatic-styles'));
          else if (is(s, 'style', 'default-style')) this.defaults.set(at(s, 'style', 'family'), this.readStyle(s, false));
          else if (is(s, 'text', 'list-style')) this.lists.set(at(s, 'style', 'name'), this.readList(s));
          else if (is(s, 'style', 'page-layout')) {
            const p = kids(s).find((c) => is(c, 'style', 'page-layout-properties'));
            this.pageLayouts.set(at(s, 'style', 'name'), p);
          }
        }
      }
      for (const m of doc.getElementsByTagNameNS(O.style, 'master-page')) this.masters.set(at(m, 'style', 'name'), m);
    };
    scan(stylesDoc, false);
    scan(contentDoc, true);
  }

  readStyle(s, auto) {
    const st = {
      name: at(s, 'style', 'name'), family: at(s, 'style', 'family'), parent: at(s, 'style', 'parent-style-name'),
      display: at(s, 'style', 'display-name'), listStyle: at(s, 'style', 'list-style-name'), outline: at(s, 'style', 'default-outline-level'),
      auto, p: {}, t: {}, cell: {}, table: {}, col: {}, row: {}, graphic: {},
    };
    for (const c of kids(s)) {
      const target = { 'paragraph-properties': st.p, 'text-properties': st.t, 'table-cell-properties': st.cell, 'table-properties': st.table, 'table-column-properties': st.col, 'table-row-properties': st.row, 'graphic-properties': st.graphic }[c.localName];
      if (!target) continue;
      for (const a of c.attributes) target[a.prefix + ':' + a.localName] = a.value;
    }
    return st;
  }

  readList(s) {
    const levels = [];
    for (const l of kids(s)) {
      const level = Number(at(l, 'text', 'level')) || 1;
      const props = kids(l).find((c) => is(c, 'style', 'list-level-properties'));
      const align = props && kids(props).find((c) => is(c, 'style', 'list-level-label-alignment'));
      const lv = {
        bullet: is(l, 'text', 'list-level-style-bullet'),
        char: at(l, 'text', 'bullet-char') || '•',
        fmt: at(l, 'style', 'num-format') || '',
        prefix: at(l, 'style', 'num-prefix') || '',
        suffix: at(l, 'style', 'num-suffix') || '',
        start: Number(at(l, 'text', 'start-value')) || 1,
        display: Number(at(l, 'text', 'display-levels')) || 1,
        left: align ? lenTw(at(align, 'fo', 'margin-left')) : null,
        indent: align ? lenTw(at(align, 'fo', 'text-indent')) : null,
      };
      if (!align && props) {
        const before = lenTw(at(props, 'text', 'space-before')) || 0;
        const min = lenTw(at(props, 'text', 'min-label-width')) || 360;
        lv.left = before + min;
        lv.indent = -min;
      }
      levels[level - 1] = lv;
    }
    return levels;
  }

  chain(name) {
    const out = [];
    const seen = new Set();
    let cur = this.map.get(name);
    while (cur && !seen.has(cur.name)) { seen.add(cur.name); out.unshift(cur); cur = cur.parent ? this.map.get(cur.parent) : null; }
    return out;
  }

  merged(name, family, skipNamed) {
    const res = { p: {}, t: {}, cell: {}, table: {}, col: {}, row: {}, graphic: {}, listStyle: null, outline: null };
    const d = this.defaults.get(family);
    const list = this.chain(name).filter((s) => !skipNamed || s.auto);
    for (const s of list) {
      for (const k of ['p', 't', 'cell', 'table', 'col', 'row', 'graphic']) Object.assign(res[k], s[k]);
      if (s.listStyle) res.listStyle = s.listStyle;
      if (s.outline) res.outline = s.outline;
    }
    res.default = d || null;
    return res;
  }
}

function paraAttrsFromOdf(p) {
  const a = {};
  const align = p['fo:text-align'];
  if (align) a.jc = { start: 'left', left: 'left', center: 'center', end: 'right', right: 'right', justify: 'both' }[align] || null;
  const ind = {};
  if (p['fo:margin-left'] != null) ind.left = String(lenTw(p['fo:margin-left']) || 0);
  if (p['fo:margin-right'] != null) ind.right = String(lenTw(p['fo:margin-right']) || 0);
  if (p['fo:text-indent'] != null) {
    const ti = lenTw(p['fo:text-indent']) || 0;
    if (ti < 0) ind.hanging = String(-ti); else ind.firstLine = String(ti);
  }
  if (Object.keys(ind).length) a.ind = ind;
  const sp = {};
  if (p['fo:margin-top'] != null) sp.before = String(lenTw(p['fo:margin-top']) || 0);
  if (p['fo:margin-bottom'] != null) sp.after = String(lenTw(p['fo:margin-bottom']) || 0);
  const lh = p['fo:line-height'];
  if (lh) {
    if (/%$/.test(lh)) { sp.line = String(Math.round((parseFloat(lh) / 100) * 240)); sp.lineRule = 'auto'; }
    else if (lenTw(lh) != null) { sp.line = String(lenTw(lh)); sp.lineRule = 'exact'; }
  } else if (p['style:line-height-at-least']) {
    sp.line = String(lenTw(p['style:line-height-at-least']) || 240); sp.lineRule = 'atLeast';
  }
  if (Object.keys(sp).length) a.spacing = sp;
  const bg = hexColor(p['fo:background-color']);
  if (bg) a.shd = { val: 'clear', color: 'auto', fill: bg };
  const b = {};
  for (const side of ['top', 'bottom', 'left', 'right']) {
    const v = p['fo:border-' + side] || p['fo:border'];
    if (v && v !== 'none') b[side] = odfBorder(v);
  }
  if (Object.keys(b).length) a.pBdr = b;
  if (p['fo:break-before'] === 'page') a.pageBreakBefore = true;
  if (p['fo:keep-with-next'] === 'always') a.keepNext = true;
  return a;
}

function marksFromOdf(t, fonts) {
  const out = [];
  const font = t['style:font-name'] ? fonts.get(t['style:font-name']) || t['style:font-name'] : t['fo:font-family'] ? t['fo:font-family'].replace(/^['"]|['"]$/g, '') : null;
  if (font) out.push(marks.rFonts.create({ a: { ascii: font, hAnsi: font, cs: font } }));
  const size = t['fo:font-size'];
  if (size && !/%$/.test(size)) { const tw = lenTw(size); if (tw) out.push(marks.sz.create({ val: Math.round(tw / 10) })); }
  if (t['fo:font-weight']) out.push(marks.b.create({ val: /bold|[6-9]00/.test(t['fo:font-weight']) }));
  if (t['fo:font-style']) out.push(marks.i.create({ val: t['fo:font-style'] === 'italic' || t['fo:font-style'] === 'oblique' }));
  const u = t['style:text-underline-style'];
  if (u && u !== 'none') {
    const val = t['style:text-underline-type'] === 'double' ? 'double' : u === 'dotted' ? 'dotted' : u === 'dash' ? 'dash' : u === 'wave' ? 'wave' : 'single';
    out.push(marks.u.create({ a: { val } }));
  } else if (u === 'none') out.push(marks.u.create({ a: { val: 'none' } }));
  const lt = t['style:text-line-through-style'];
  if (lt && lt !== 'none') out.push(marks.strike.create({ val: true }));
  const color = hexColor(t['fo:color']);
  if (color) out.push(marks.color.create({ a: { val: color } }));
  const bg = hexColor(t['fo:background-color']);
  if (bg) out.push(marks.shd.create({ a: { val: 'clear', color: 'auto', fill: bg } }));
  const pos = t['style:text-position'];
  if (pos) {
    if (/^super/.test(pos) || /^[1-9]/.test(pos)) out.push(marks.vertAlign.create({ val: 'superscript' }));
    else if (/^sub/.test(pos) || /^-/.test(pos)) out.push(marks.vertAlign.create({ val: 'subscript' }));
  }
  if (t['fo:font-variant'] === 'small-caps') out.push(marks.smallCaps.create({ val: true }));
  if (t['fo:text-transform'] === 'uppercase') out.push(marks.caps.create({ val: true }));
  if (t['fo:letter-spacing'] && t['fo:letter-spacing'] !== 'normal') { const tw = lenTw(t['fo:letter-spacing']); if (tw) out.push(marks.spacing.create({ val: tw })); }
  return out;
}

function withMarks(base, add) {
  let set = base;
  for (const m of add) set = m.addToSet(set);
  return set;
}

// ---------------------------------------------------------------- leitura

class OdtReader {
  constructor(zip, model, styles) {
    this.zip = zip;
    this.model = model;
    this.S = styles;
    this.lists = new Map();
    this.lastListByStyle = new Map();
    this.noteId = 0;
    this.mediaN = 0;
    this.pending = [];
    this.maxW = contentWidthPx(model);
  }

  async media(href) {
    if (!href) return null;
    const f = this.zip.file(href.replace(/^\.\//, ''));
    if (!f) return null;
    const bytes = await f.async('uint8array');
    const ext = (href.split('.').pop() || 'png').toLowerCase();
    this.mediaN++;
    const path = 'word/media/odt' + this.mediaN + '.' + ext;
    const type = MIME[ext] || 'image/png';
    this.model.media.set(path, { url: URL.createObjectURL(new Blob([bytes], { type })), bytes, ext, type });
    return path;
  }

  listNumId(listStyleName, level, continueList, listEl) {
    const styleName = listStyleName || 'default';
    const key = styleName;
    if (continueList && this.lastListByStyle.has(key)) return this.lastListByStyle.get(key);
    const levels = this.S.lists.get(listStyleName) || [];
    const id = this.model.numbering.createCustomList(levels.map((l) => l && {
      bullet: l.bullet, char: l.char, fmt: l.fmt, text: l.bullet ? l.char : null, prefix: l.prefix, suffix: l.suffix, start: l.start, display: l.display, left: l.left, hanging: l.indent != null ? -l.indent : null,
    }));
    this.lastListByStyle.set(key, id);
    return id;
  }

  async blocks(parent, ctx) {
    const out = [];
    for (const el of kids(parent)) {
      if (is(el, 'text', 'p') || is(el, 'text', 'h')) out.push(await this.paragraph(el, ctx));
      else if (is(el, 'text', 'list')) out.push(...(await this.list(el, ctx, 0)));
      else if (is(el, 'table', 'table')) { out.push(await this.table(el, ctx)); }
      else if (is(el, 'text', 'section') || is(el, 'text', 'index-body')) out.push(...(await this.blocks(el, ctx)));
      else if (el.localName === 'table-of-content' || el.localName === 'alphabetical-index' || el.localName === 'illustration-index') {
        const body = kids(el).find((c) => c.localName === 'index-body');
        if (body) out.push(...(await this.blocks(body, ctx)));
      }
    }
    return out;
  }

  async list(el, ctx, depth, inherited) {
    const styleName = at(el, 'text', 'style-name') || (inherited && inherited.style) || null;
    const cont = at(el, 'text', 'continue-numbering') === 'true' || !!at(el, 'text', 'continue-list');
    const numId = depth === 0 ? this.listNumId(styleName, depth, cont, el) : inherited.numId;
    const out = [];
    for (const item of kids(el)) {
      if (!is(item, 'text', 'list-item') && !is(item, 'text', 'list-header')) continue;
      let first = true;
      for (const c of kids(item)) {
        if (is(c, 'text', 'p') || is(c, 'text', 'h')) {
          const p = await this.paragraph(c, ctx);
          const a = { ...p.attrs };
          if (first && is(item, 'text', 'list-item')) { a.numId = numId; a.ilvl = depth; first = false; }
          else { a.ind = { ...(a.ind || {}), left: String(720 * (depth + 1)) }; first = false; }
          if (!a.styleId) a.styleId = this.model.styles.get('ListParagraph') ? 'ListParagraph' : null;
          out.push(p.type.create(a, p.content));
        } else if (is(c, 'text', 'list')) {
          out.push(...(await this.list(c, ctx, depth + 1, { numId, style: styleName })));
        } else if (is(c, 'table', 'table')) out.push(await this.table(c, ctx));
      }
    }
    return out;
  }

  async paragraph(el, ctx) {
    const styleName = at(el, 'text', 'style-name');
    const isH = is(el, 'text', 'h');
    let attrs = {};
    let base = [];
    if (isH) {
      const level = Math.min(9, Number(at(el, 'text', 'outline-level')) || 1);
      attrs.styleId = this.model.styles.headingId(Math.min(6, level)) || 'Heading1';
      const m = this.S.merged(styleName, 'paragraph', true);
      attrs = { ...paraAttrsFromOdf(m.p), ...attrs };
      base = marksFromOdf(m.t, this.S.fonts);
    } else {
      const m = this.S.merged(styleName, 'paragraph', false);
      const chainNames = this.S.chain(styleName).map((s) => s.name);
      if (chainNames.includes('Title')) attrs.styleId = 'Title';
      else if (chainNames.includes('Subtitle')) attrs.styleId = 'Subtitle';
      else if (chainNames.some((n) => /Quotations|Quote/i.test(n))) attrs.styleId = 'Quote';
      else if (chainNames.some((n) => /Preformatted/i.test(n))) attrs.styleId = 'Code';
      attrs = { ...paraAttrsFromOdf(m.p), ...attrs };
      base = marksFromOdf(m.t, this.S.fonts);
    }
    const content = [];
    await this.inline(el, base, content, ctx);
    if (base.length) attrs.markRPr = base.map((mk) => ({ type: mk.type.name, attrs: mk.attrs }));
    return nodes.paragraph.create(attrs, content);
  }

  async inline(parent, active, out, ctx) {
    for (const n of parent.childNodes) {
      if (n.nodeType === 3) {
        const t = n.data.replace(/[\n\r\t]+/g, ' ').replace(/ {2,}/g, ' ');
        if (t) out.push(schema.text(t, active));
        continue;
      }
      if (n.nodeType !== 1) continue;
      const el = n;
      if (is(el, 'text', 'span')) {
        const m = this.S.merged(at(el, 'text', 'style-name'), 'text', false);
        await this.inline(el, withMarks(active, marksFromOdf(m.t, this.S.fonts)), out, ctx);
      } else if (is(el, 'text', 's')) {
        out.push(schema.text(' '.repeat(Number(at(el, 'text', 'c')) || 1), active));
      } else if (is(el, 'text', 'tab')) {
        out.push(nodes.tab.create(null, null, active));
      } else if (is(el, 'text', 'line-break')) {
        out.push(nodes.hard_break.create(null, null, active));
      } else if (is(el, 'text', 'a')) {
        const href = el.getAttributeNS(O.xlink, 'href') || '';
        const link = href.startsWith('#') ? marks.hyperlink.create({ anchor: href.slice(1) }) : marks.hyperlink.create({ href });
        const style = this.model.styles.get('Hyperlink') ? [marks.rStyle.create({ val: 'Hyperlink' })] : [];
        await this.inline(el, withMarks(active, [link, ...style]), out, ctx);
      } else if (is(el, 'text', 'note')) {
        const body = kids(el).find((c) => is(c, 'text', 'note-body'));
        const kind = at(el, 'text', 'note-class') === 'endnote' ? 'endnote' : 'footnote';
        this.noteId++;
        const id = String(this.noteId);
        const blocks = body ? await this.blocks(body, ctx) : [];
        if (!blocks.length) blocks.push(nodes.paragraph.create());
        const refStyle = this.model.styles.get('FootnoteReference') ? [marks.rStyle.create({ val: 'FootnoteReference' })] : [];
        const mark = nodes.opaque_inline.create({ xml: '<w:footnoteRef xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>', level: 'r', kind: 'noteMark' }, null, refStyle);
        const fp = blocks[0];
        const textStyle = this.model.styles.get('FootnoteText') ? 'FootnoteText' : null;
        blocks[0] = fp.type.create({ ...fp.attrs, styleId: fp.attrs.styleId || textStyle }, [mark, schema.text(' '), ...(fp.content.size ? fp.content.content : [])]);
        const map = kind === 'footnote' ? this.model.footnotes : this.model.endnotes;
        map.set(id, { id, doc: schema.node('doc', null, blocks), isNew: true });
        out.push(nodes.footnote_ref.create({ id, kind }, null, withMarks(active, refStyle)));
      } else if (is(el, 'draw', 'frame')) {
        out.push(...(await this.frame(el, active, ctx)));
      } else if (is(el, 'text', 'page-number') || is(el, 'text', 'page-count')) {
        const instr = is(el, 'text', 'page-number') ? ' PAGE ' : ' NUMPAGES ';
        out.push(nodes.field_char.create({ type: 'begin' }, null, active), nodes.instr_text.create({ text: instr }, null, active), nodes.field_char.create({ type: 'separate' }, null, active));
        if (el.textContent) out.push(schema.text(el.textContent, active));
        out.push(nodes.field_char.create({ type: 'end' }, null, active));
      } else if (is(el, 'text', 'soft-page-break') || is(el, 'text', 'bookmark') || is(el, 'text', 'bookmark-start') || is(el, 'text', 'bookmark-end') || el.localName === 'annotation' || el.localName === 'annotation-end') {
        // ignorados
      } else {
        await this.inline(el, active, out, ctx);
      }
    }
  }

  async frame(el, active, ctx) {
    const w = lenTw(at(el, 'svg', 'width'));
    const h = lenTw(at(el, 'svg', 'height'));
    const img = kids(el).find((c) => is(c, 'draw', 'image'));
    const box = kids(el).find((c) => is(c, 'draw', 'text-box'));
    const cx = w ? Math.round((w / 1440) * 914400) : 0;
    const cy = h ? Math.round((h / 1440) * 914400) : 0;
    const anchor = at(el, 'text', 'anchor-type');
    if (img) {
      const path = await this.media(img.getAttributeNS(O.xlink, 'href'));
      if (!path) return [];
      const title = kids(el).find((c) => c.localName === 'title' || c.localName === 'desc');
      return [nodes.drawing.create({ kind: 'image', mediaPath: path, cx, cy, ocx: cx, ocy: cy, alt: title ? title.textContent : '', anchor: anchor === 'as-char' ? null : null }, null, active)];
    }
    if (box) {
      const blocks = await this.blocks(box, ctx);
      return [nodes.drawing.create({ kind: 'shape', cx, cy, ocx: cx, ocy: cy, content: blocks.map((b) => b.toJSON()), line: { color: '#000000', w: 9525 } }, null, active)];
    }
    return [];
  }

  async table(el, ctx) {
    const tstyle = this.S.merged(at(el, 'table', 'style-name'), 'table', false);
    const cols = [];
    const rowsEl = [];
    const walk = (parent, header) => {
      for (const c of kids(parent)) {
        if (is(c, 'table', 'table-column')) {
          const rep = Number(at(c, 'table', 'number-columns-repeated')) || 1;
          const cs = this.S.merged(at(c, 'table', 'style-name'), 'table-column', false);
          const wTw = lenTw(cs.col['style:column-width']);
          for (let i = 0; i < Math.min(rep, 64); i++) cols.push(wTw);
        } else if (is(c, 'table', 'table-columns') || is(c, 'table', 'table-column-group')) walk(c, header);
        else if (is(c, 'table', 'table-header-rows')) walk(c, true);
        else if (is(c, 'table', 'table-rows') || is(c, 'table', 'table-row-group')) walk(c, header);
        else if (is(c, 'table', 'table-row')) {
          const rep = Math.min(Number(at(c, 'table', 'number-rows-repeated')) || 1, 200);
          for (let i = 0; i < rep; i++) rowsEl.push({ el: c, header });
        }
      }
    };
    walk(el, false);
    const pmRows = [];
    let maxCols = cols.length;
    for (const r of rowsEl) {
      let n = 0;
      for (const c of kids(r.el)) if (is(c, 'table', 'table-cell') || is(c, 'table', 'covered-table-cell')) n += Math.min(Number(at(c, 'table', 'number-columns-repeated')) || 1, 64);
      maxCols = Math.max(maxCols, n);
    }
    const total = Math.round(this.maxW * 15);
    const known = cols.filter((x) => x).reduce((s, x) => s + x, 0);
    const grid = [];
    for (let i = 0; i < maxCols; i++) grid.push(cols[i] || Math.round((total - known) / Math.max(1, maxCols - cols.filter((x) => x).length)) || 1440);
    for (const r of rowsEl) {
      const cells = [];
      let col = 0;
      for (const c of kids(r.el)) {
        if (is(c, 'table', 'covered-table-cell')) { col += Number(at(c, 'table', 'number-columns-repeated')) || 1; continue; }
        if (!is(c, 'table', 'table-cell')) continue;
        const rep = Math.min(Number(at(c, 'table', 'number-columns-repeated')) || 1, 64);
        for (let k = 0; k < rep; k++) {
          const span = Number(at(c, 'table', 'number-columns-spanned')) || 1;
          const rspan = Number(at(c, 'table', 'number-rows-spanned')) || 1;
          const cs = this.S.merged(at(c, 'table', 'style-name'), 'table-cell', false);
          const cattrs = { colspan: span, rowspan: rspan, colwidth: grid.slice(col, col + span).map((w) => w / 15) };
          const bg = hexColor(cs.cell['fo:background-color']);
          if (bg) cattrs.shd = { val: 'clear', color: 'auto', fill: bg };
          const b = {};
          for (const side of ['top', 'bottom', 'left', 'right']) { const v = cs.cell['fo:border-' + side] || cs.cell['fo:border']; if (v) b[side] = odfBorder(v); }
          if (Object.keys(b).length) cattrs.tcBorders = b;
          const va = cs.cell['style:vertical-align'];
          if (va === 'middle') cattrs.vAlign = 'center'; else if (va === 'bottom') cattrs.vAlign = 'bottom';
          let content = await this.blocks(c, ctx);
          if (!content.length) content = [nodes.paragraph.create()];
          cells.push(nodes.table_cell.create(cattrs, content));
          col += span;
        }
      }
      if (!cells.length) cells.push(nodes.table_cell.create({ colwidth: [grid[0] / 15] }, [nodes.paragraph.create()]));
      pmRows.push(nodes.table_row.create(r.header ? { header: true } : null, cells));
    }
    const tattrs = { grid, tblW: { w: '0', type: 'auto' }, tblLook: { val: '0000', firstRow: '0', lastRow: '0', firstColumn: '0', lastColumn: '0', noHBand: '1', noVBand: '1' } };
    const align = tstyle.table['table:align'];
    if (align === 'center') tattrs.jc = 'center';
    else if (align === 'right') tattrs.jc = 'right';
    const ml = lenTw(tstyle.table['fo:margin-left']);
    if (ml) tattrs.tblInd = { w: String(ml), type: 'dxa' };
    return nodes.table.create(tattrs, pmRows.length ? pmRows : [nodes.table_row.create(null, [nodes.table_cell.create(null, [nodes.paragraph.create()])])]);
  }
}

// OpenDocument plano (.fodt): um XML só, com as imagens dentro em base64.
// Vira um pacote comum na memória para seguir pela mesma leitura.
async function flatToZip(bytes) {
  const doc = parseXml(new TextDecoder('utf-8').decode(bytes).replace(/^\uFEFF/, ''));
  const zip = new JSZip();
  let n = 0;
  for (const img of Array.from(doc.getElementsByTagNameNS(O.draw, 'image'))) {
    const bin = Array.from(img.children).find((c) => c.localName === 'binary-data');
    if (!bin) continue;
    let data;
    try { data = Uint8Array.from(atob(bin.textContent.replace(/\s+/g, '')), (c) => c.charCodeAt(0)); } catch { continue; }
    const ext = data[0] === 0x89 ? 'png' : data[0] === 0xFF ? 'jpg' : data[0] === 0x47 ? 'gif' : data[0] === 0x3C ? 'svg' : 'png';
    const path = 'Pictures/plano' + (++n) + '.' + ext;
    zip.file(path, data);
    img.removeChild(bin);
    img.setAttributeNS(O.xlink, 'xlink:href', path);
  }
  const xml = new XMLSerializer().serializeToString(doc);
  zip.file('content.xml', xml);
  zip.file('styles.xml', xml);
  zip.file('meta.xml', xml);
  return zip.generateAsync({ type: 'uint8array' });
}

export async function readOdt(bytes) {
  // começa com "<": é o formato plano
  const start = bytes[0] === 0xEF ? 3 : 0;
  const flat = bytes[start] === 0x3C;
  const original = bytes.slice();
  if (flat) bytes = await flatToZip(bytes);
  const zip = await JSZip.loadAsync(bytes);
  const contentFile = zip.file('content.xml');
  if (!contentFile) throw new Error('O arquivo não parece ser um documento OpenDocument.');
  const content = parseXml(await contentFile.async('string'));
  const stylesDoc = zip.file('styles.xml') ? parseXml(await zip.file('styles.xml').async('string')) : null;
  const meta = zip.file('meta.xml') ? parseXml(await zip.file('meta.xml').async('string')) : null;
  const { model } = await blankModel();
  model.format = 'odt';
  const S = new OdfStyles(content, stylesDoc);
  // padrão do documento
  const dp = S.defaults.get('paragraph');
  if (dp) {
    const fnt = dp.t['style:font-name'] ? S.fonts.get(dp.t['style:font-name']) : null;
    const size = lenTw(dp.t['fo:font-size']);
    model.styles.setDefaultFont(fnt || null, size ? Math.round(size / 10) : null);
  }
  const std = S.map.get('Standard');
  if (std) {
    const fnt = std.t['style:font-name'] ? S.fonts.get(std.t['style:font-name']) : null;
    const size = lenTw(std.t['fo:font-size']);
    if (fnt || size) model.styles.setDefaultFont(fnt || null, size ? Math.round(size / 10) : null);
  }
  model.styles.setDefaultParagraphSpacing(0, 0, 240);
  // página
  const master = S.masters.get('Standard') || [...S.masters.values()][0];
  if (master) {
    const pl = S.pageLayouts.get(at(master, 'style', 'page-layout-name'));
    if (pl) {
      const w = lenTw(at(pl, 'fo', 'page-width'));
      const hgt = lenTw(at(pl, 'fo', 'page-height'));
      if (w && hgt) model.section.setPageSize(w, hgt, w > hgt ? 'landscape' : null);
      const mm = {};
      for (const side of ['top', 'bottom', 'left', 'right']) { const v = lenTw(at(pl, 'fo', 'margin-' + side)); if (v != null) mm[side] = v; }
      if (Object.keys(mm).length) model.section.setMargins(mm);
    }
  }
  const reader = new OdtReader(zip, model, S);
  const prev = R.model;
  R.model = model;
  try {
    const text = content.getElementsByTagNameNS(O.office, 'text')[0];
    const blocks = text ? await reader.blocks(text, {}) : [];
    if (master) {
      for (const [tag, kind] of [['header', 'header'], ['footer', 'footer']]) {
        const el = kids(master).find((c) => is(c, 'style', tag));
        if (!el) continue;
        const hb = await reader.blocks(el, {});
        if (!hb.length) continue;
        const entry = newHeaderFooter(model, kind);
        entry.doc = schema.node('doc', null, hb);
        model.section.setReference(kind, 'default', entry.rId);
      }
    }
    if (meta) {
      const t = meta.getElementsByTagNameNS(O.dc, 'title')[0];
      const c = meta.getElementsByTagNameNS(O.meta, 'initial-creator')[0] || meta.getElementsByTagNameNS(O.dc, 'creator')[0];
      if (t) model.properties.title = t.textContent;
      if (c) model.properties.author = c.textContent;
    }
    const doc = finalizeDoc(schema.node('doc', null, blocks.length ? blocks : [nodes.paragraph.create()]), model);
    const mime = zip.file('mimetype') ? (await zip.file('mimetype').async('string')).trim() : '';
    model.extra.odt = {
      bytes: original, flat, template: mime === 'application/vnd.oasis.opendocument.text-template',
      diagnostics: odtDiagnostics(content, stylesDoc, meta, S, zip),
    };
    bindOdtSource(model, doc);
    return { model, doc };
  } catch (err) {
    model.dispose();
    throw err;
  } finally {
    R.model = prev;
  }
}

// ---------------------------------------------------------------- escrita

const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const cm = (tw) => (Number(tw) / TWIP_PER_CM).toFixed(3) + 'cm';
const pt = (tw) => (Number(tw) / 20).toFixed(2) + 'pt';

class OdtWriter {
  constructor(model) {
    this.model = model;
    this.styles = model.styles;
    this.auto = new Map();
    this.autoXml = [];
    this.named = new Map();
    this.namedXml = [];
    this.listStyles = new Map();
    this.listXml = [];
    this.pics = [];
    this.picN = 0;
    this.tableN = 0;
    this.pendingBreak = false;
    this.fonts = new Set();
  }

  pProps(pPr, extra = {}) {
    const a = [];
    if (pPr.jc) a.push('fo:text-align="' + ({ left: 'start', center: 'center', right: 'end', both: 'justify', distribute: 'justify' }[pPr.jc] || 'start') + '"');
    const ind = pPr.ind || {};
    if (ind.left != null) a.push('fo:margin-left="' + cm(ind.left) + '"');
    if (ind.right != null) a.push('fo:margin-right="' + cm(ind.right) + '"');
    if (ind.hanging != null) a.push('fo:text-indent="-' + cm(ind.hanging) + '"');
    else if (ind.firstLine != null) a.push('fo:text-indent="' + cm(ind.firstLine) + '"');
    const sp = pPr.spacing || {};
    if (sp.before != null) a.push('fo:margin-top="' + cm(sp.before) + '"');
    if (sp.after != null) a.push('fo:margin-bottom="' + cm(sp.after) + '"');
    if (sp.line != null) {
      if (!sp.lineRule || sp.lineRule === 'auto') a.push('fo:line-height="' + Math.round((Number(sp.line) / 240) * 100) + '%"');
      else if (sp.lineRule === 'exact') a.push('fo:line-height="' + pt(sp.line) + '"');
      else a.push('style:line-height-at-least="' + pt(sp.line) + '"');
    }
    const bg = pPr.shd ? shdColor(pPr.shd) : null;
    if (bg) a.push('fo:background-color="' + bg + '"');
    if (pPr.pBdr) {
      for (const side of ['top', 'bottom', 'left', 'right']) {
        const b = pPr.pBdr[side];
        if (b && b.val !== 'nil' && b.val !== 'none') {
          const css = borderCss(b).split(' ');
          a.push('fo:border-' + side + '="' + css[0] + ' solid ' + css[2] + '"');
        }
      }
    }
    if (pPr.keepNext) a.push('fo:keep-with-next="always"');
    if (extra.breakBefore || pPr.pageBreakBefore) a.push('fo:break-before="page"');
    return a.length ? '<style:paragraph-properties ' + a.join(' ') + '/>' : '';
  }

  tProps(r) {
    const a = [];
    if (r.rFonts) {
      const f = fontName(r.rFonts.a);
      if (f) { this.fonts.add(f); a.push('style:font-name="' + esc(f) + '"'); }
    }
    if (r.sz) a.push('fo:font-size="' + r.sz.val / 2 + 'pt"');
    if (r.b) a.push('fo:font-weight="' + (r.b.val ? 'bold' : 'normal') + '"');
    if (r.i) a.push('fo:font-style="' + (r.i.val ? 'italic' : 'normal') + '"');
    if (r.u && r.u.a) {
      const v = r.u.a.val;
      if (v === 'none') a.push('style:text-underline-style="none"');
      else a.push('style:text-underline-style="' + (v === 'dotted' ? 'dotted' : v === 'dash' ? 'dash' : v === 'wave' ? 'wave' : 'solid') + '" style:text-underline-width="auto" style:text-underline-color="font-color"' + (v === 'double' ? ' style:text-underline-type="double"' : ''));
    }
    if ((r.strike && r.strike.val) || (r.dstrike && r.dstrike.val)) a.push('style:text-line-through-style="solid"');
    if (r.color) { const c = colorOf(r.color.a); if (c && c !== 'auto') a.push('fo:color="' + c + '"'); }
    if (r.highlight && r.highlight.val !== 'none') {
      const map = { yellow: '#FFFF00', green: '#00FF00', cyan: '#00FFFF', magenta: '#FF00FF', blue: '#0000FF', red: '#FF0000', darkBlue: '#000080', darkCyan: '#008080', darkGreen: '#008000', darkMagenta: '#800080', darkRed: '#800000', darkYellow: '#808000', darkGray: '#808080', lightGray: '#C0C0C0', black: '#000000' };
      a.push('fo:background-color="' + (map[r.highlight.val] || '#FFFF00') + '"');
    } else if (r.shd) { const c = shdColor(r.shd.a); if (c) a.push('fo:background-color="' + c + '"'); }
    if (r.vertAlign && r.vertAlign.val === 'superscript') a.push('style:text-position="super 58%"');
    if (r.vertAlign && r.vertAlign.val === 'subscript') a.push('style:text-position="sub 58%"');
    if (r.smallCaps && r.smallCaps.val) a.push('fo:font-variant="small-caps"');
    if (r.caps && r.caps.val) a.push('fo:text-transform="uppercase"');
    if (r.spacing && r.spacing.val) a.push('fo:letter-spacing="' + pt(r.spacing.val) + '"');
    return a.length ? '<style:text-properties ' + a.join(' ') + '/>' : '';
  }

  namedStyle(styleId) {
    const st = this.styles.get(styleId) || this.styles.get(this.styles.defaultPara);
    const id = st ? st.id : 'Normal';
    if (this.named.has(id)) return this.named.get(id);
    let name;
    const level = this.styles.headingLevel(id);
    if (id === this.styles.defaultPara) name = 'Standard';
    else if (level >= 1 && level <= 9 && /^heading/i.test(st.name || '')) name = 'Heading_20_' + level;
    else name = 'W_' + id.replace(/[^A-Za-z0-9_]/g, '_');
    this.named.set(id, name);
    const res = this.styles.resolvePara(id);
    const display = this.styles.displayName(st);
    const parent = id === this.styles.defaultPara ? '' : ' style:parent-style-name="Standard"';
    const outline = level >= 1 && /^heading/i.test(st.name || '') ? ' style:default-outline-level="' + level + '"' : '';
    this.namedXml.push('<style:style style:name="' + name + '" style:display-name="' + esc(display) + '" style:family="paragraph"' + parent + outline + ' style:class="text">' + this.pProps(res.pPr) + this.tProps(res.r) + '</style:style>');
    return name;
  }

  autoStyle(family, parent, inner) {
    if (!inner) return parent;
    const key = family + '|' + parent + '|' + inner;
    if (this.auto.has(key)) return this.auto.get(key);
    const prefix = family === 'paragraph' ? 'P' : family === 'text' ? 'T' : family === 'table-cell' ? 'C' : family === 'graphic' ? 'fr' : 'S';
    const name = prefix + (this.auto.size + 1);
    this.auto.set(key, name);
    this.autoXml.push('<style:style style:name="' + name + '" style:family="' + family + '"' + (parent ? ' style:parent-style-name="' + parent + '"' : '') + '>' + inner + '</style:style>');
    return name;
  }

  listStyle(numId) {
    if (this.listStyles.has(numId)) return this.listStyles.get(numId);
    const name = 'L' + (this.listStyles.size + 1);
    this.listStyles.set(numId, name);
    let x = '<text:list-style style:name="' + name + '">';
    for (let i = 0; i < 9; i++) {
      const l = this.model.numbering.level(numId, i) || { numFmt: 'bullet', lvlText: '•' };
      const ind = (l.pPr && l.pPr.ind) || { left: String(720 * (i + 1)), hanging: '360' };
      const props = '<style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="' + cm(ind.left || 720 * (i + 1)) + '" fo:text-indent="-' + cm(ind.hanging || 360) + '" fo:margin-left="' + cm(ind.left || 720 * (i + 1)) + '"/></style:list-level-properties>';
      if (l.numFmt === 'bullet') {
        const ch = { '': '•', o: '◦', '': '▪' }[l.lvlText] || (l.lvlText && l.lvlText.charCodeAt(0) < 0xF000 ? l.lvlText : '•');
        x += '<text:list-level-style-bullet text:level="' + (i + 1) + '" text:bullet-char="' + esc(ch) + '">' + props + '</text:list-level-style-bullet>';
      } else {
        const fmt = { decimal: '1', lowerLetter: 'a', upperLetter: 'A', lowerRoman: 'i', upperRoman: 'I' }[l.numFmt] || '1';
        const m = /^(.*?)%\d(.*)$/.exec(l.lvlText || '%1.');
        x += '<text:list-level-style-number text:level="' + (i + 1) + '" style:num-format="' + fmt + '"' + (m && m[1] ? ' style:num-prefix="' + esc(m[1]) + '"' : '') + ' style:num-suffix="' + esc(m ? m[2] : '.') + '" text:start-value="' + (l.start ?? 1) + '">' + props + '</text:list-level-style-number>';
      }
    }
    this.listXml.push(x + '</text:list-style>');
    return name;
  }

  async run(child, para) {
    const r = {};
    for (const m of child.marks) if (m.type.name !== 'hyperlink' && m.type.name !== 'sdt' && m.type.name !== 'ins' && m.type.name !== 'rprExtra') r[m.type.name] = m.attrs;
    if (r.rStyle) {
      const cs = this.styles.resolveChar(r.rStyle.val).r;
      for (const [k, v] of Object.entries(cs)) if (!r[k]) r[k] = v;
      delete r.rStyle;
    }
    const style = this.tProps(r);
    const name = style ? this.autoStyle('text', '', style) : null;
    let inner = '';
    if (child.isText) inner = esc(child.text).replace(/ {2,}/g, (m) => ' <text:s text:c="' + (m.length - 1) + '"/>').replace(/\t/g, '<text:tab/>');
    else if (child.type === nodes.tab) inner = '<text:tab/>';
    else if (child.type === nodes.hard_break) {
      if (child.attrs.type === 'page') { this.pendingBreak = true; return ''; }
      inner = '<text:line-break/>';
    } else if (child.type === nodes.drawing) inner = await this.drawing(child);
    else if (child.type === nodes.sym) inner = esc(child.textContent);
    else if (child.type === nodes.footnote_ref) inner = await this.note(child);
    else if (child.type === nodes.opaque_inline && child.attrs.kind === 'math') inner = esc(child.attrs.text || '');
    else return '';
    return name ? '<text:span text:style-name="' + name + '">' + inner + '</text:span>' : inner;
  }

  async note(ref) {
    const map = ref.attrs.kind === 'endnote' ? this.model.endnotes : this.model.footnotes;
    const fn = map.get(String(ref.attrs.id));
    this.noteN = (this.noteN || 0) + 1;
    let body = '';
    if (fn) body = await this.blocks(fn.doc);
    return '<text:note text:id="ftn' + this.noteN + '" text:note-class="' + (ref.attrs.kind === 'endnote' ? 'endnote' : 'footnote') + '"><text:note-citation>' + this.noteN + '</text:note-citation><text:note-body>' + body + '</text:note-body></text:note>';
  }

  async drawing(node) {
    const a = node.attrs;
    const w = ((a.cx || 914400) / 360000).toFixed(3) + 'cm';
    const h = ((a.cy || 685800) / 360000).toFixed(3) + 'cm';
    if (a.content && a.content.length) {
      const blocks = a.content.map((j) => PMNode.fromJSON(schema, j));
      const inner = await this.blocks(schema.node('doc', null, blocks));
      return '<draw:frame draw:style-name="' + this.autoStyle('graphic', '', '<style:graphic-properties fo:border="0.5pt solid #000000" fo:padding="0.1cm"/>') + '" text:anchor-type="as-char" svg:width="' + w + '" svg:height="' + h + '"><draw:text-box>' + inner + '</draw:text-box></draw:frame>';
    }
    let bytes = null;
    let ext = 'png';
    if (a.mediaPath && this.model.media.get(a.mediaPath)) { const m = this.model.media.get(a.mediaPath); bytes = m.bytes; ext = m.ext || 'png'; }
    else if (a.src) {
      try { const res = await fetch(a.src); bytes = new Uint8Array(await res.arrayBuffer()); ext = (res.headers.get('content-type') || 'image/png').split('/')[1].replace('jpeg', 'jpg').replace('svg+xml', 'svg'); } catch { bytes = null; }
    }
    if (!bytes) return '';
    this.picN++;
    const path = 'Pictures/imagem' + this.picN + '.' + ext;
    this.pics.push({ path, bytes, type: MIME[ext] || 'image/png' });
    return '<draw:frame draw:style-name="' + this.autoStyle('graphic', '', '<style:graphic-properties style:vertical-pos="top" style:vertical-rel="baseline"/>') + '" draw:name="Imagem' + this.picN + '" text:anchor-type="as-char" svg:width="' + w + '" svg:height="' + h + '" draw:z-index="0"><draw:image xlink:href="' + path + '" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>' + (a.alt ? '<svg:desc>' + esc(a.alt) + '</svg:desc>' : '') + '</draw:frame>';
  }

  async inline(para) {
    let out = '';
    let link = null;
    let linkBuf = '';
    const children = [];
    para.forEach((c) => children.push(c));
    for (const c of children) {
      const lm = c.marks.find((m) => m.type.name === 'hyperlink');
      const same = lm && link && lm.eq(link);
      if (!same && link) { out += '<text:a xlink:type="simple" xlink:href="' + esc(link.attrs.href || '#' + (link.attrs.anchor || '')) + '">' + linkBuf + '</text:a>'; link = null; linkBuf = ''; }
      if (lm && !link) link = lm;
      const x = await this.run(c, para);
      if (link) linkBuf += x; else out += x;
    }
    if (link) out += '<text:a xlink:type="simple" xlink:href="' + esc(link.attrs.href || '#' + (link.attrs.anchor || '')) + '">' + linkBuf + '</text:a>';
    return out;
  }

  async paragraph(n) {
    const a = n.attrs;
    const named = this.namedStyle(a.styleId);
    const direct = {};
    for (const k of ['jc', 'ind', 'spacing', 'shd', 'pBdr', 'keepNext', 'pageBreakBefore']) if (a[k] != null) direct[k] = a[k];
    if (effectiveNumPr(a)) delete direct.ind;
    let mark = '';
    if (a.markRPr && a.markRPr.length) {
      const r = {};
      for (const m of a.markRPr) r[m.type] = m.attrs;
      delete r.rStyle; delete r.rprExtra;
      mark = this.tProps(r);
    }
    const breakBefore = this.pendingBreak;
    this.pendingBreak = false;
    const props = this.pProps(direct, { breakBefore }) + mark;
    const style = this.autoStyle('paragraph', named, props);
    const inner = await this.inline(n);
    const level = this.styles.headingLevel(a.styleId || this.styles.defaultPara);
    if (level >= 1 && level <= 9 && named.startsWith('Heading')) return '<text:h text:style-name="' + style + '" text:outline-level="' + level + '">' + inner + '</text:h>';
    return '<text:p text:style-name="' + style + '">' + inner + '</text:p>';
  }

  async table(t) {
    this.tableN++;
    const name = 'Tabela' + this.tableN;
    const grid = t.attrs.grid && t.attrs.grid.length ? t.attrs.grid : null;
    const cols = [];
    t.firstChild.forEach((cell) => { for (let i = 0; i < (cell.attrs.colspan || 1); i++) cols.push(cell.attrs.colwidth && cell.attrs.colwidth[i] ? cell.attrs.colwidth[i] * 15 : grid ? grid[cols.length] : 1440); });
    const total = cols.reduce((s, x) => s + (x || 0), 0);
    this.autoXml.push('<style:style style:name="' + name + '" style:family="table"><style:table-properties style:width="' + cm(total) + '" table:align="' + (t.attrs.jc === 'center' ? 'center' : t.attrs.jc === 'right' ? 'right' : 'left') + '"/></style:style>');
    let x = '<table:table table:name="' + name + '" table:style-name="' + name + '">';
    cols.forEach((w, i) => {
      const cn = name + '.C' + (i + 1);
      this.autoXml.push('<style:style style:name="' + cn + '" style:family="table-column"><style:table-column-properties style:column-width="' + cm(w || 1440) + '"/></style:style>');
      x += '<table:table-column table:style-name="' + cn + '"/>';
    });
    const R2 = this.styles.resolveTable(t.attrs.styleId || this.styles.defaultTable);
    const tb = { ...(R2.tblPr.tblBorders || {}), ...(t.attrs.tblBorders || {}) };
    const bstr = (b) => (b && b.val !== 'nil' && b.val !== 'none' ? borderCss(b).replace(/ (solid|double|dotted|dashed).*? /, ' $1 ') : 'none');
    const covered = new Map();
    let r = 0;
    const rows = [];
    t.forEach((row) => rows.push(row));
    for (const row of rows) {
      x += '<table:table-row>';
      let c = 0;
      const cells = [];
      row.forEach((cell) => cells.push(cell));
      let ci = 0;
      while (ci < cells.length || covered.has(r + ':' + c)) {
        if (covered.has(r + ':' + c)) { x += '<table:covered-table-cell/>'; c++; continue; }
        const cell = cells[ci++];
        if (!cell) break;
        const cs = cell.attrs.colspan || 1;
        const rs = cell.attrs.rowspan || 1;
        for (let dr = 0; dr < rs; dr++) for (let dc = 0; dc < cs; dc++) if (dr || dc) covered.set((r + dr) + ':' + (c + dc), true);
        const bg = cell.attrs.shd ? shdColor(cell.attrs.shd) : null;
        const b = { ...tb, ...(cell.attrs.tcBorders || {}) };
        const props = '<style:table-cell-properties fo:padding="0.097cm"' + (bg ? ' fo:background-color="' + bg + '"' : '')
          + ' fo:border-top="' + bstr(r === 0 ? b.top : b.insideH || b.top) + '" fo:border-bottom="' + bstr(b.bottom) + '" fo:border-left="' + bstr(c === 0 ? b.left : b.insideV || b.left) + '" fo:border-right="' + bstr(b.right) + '"'
          + (cell.attrs.vAlign === 'center' ? ' style:vertical-align="middle"' : cell.attrs.vAlign === 'bottom' ? ' style:vertical-align="bottom"' : '') + '/>';
        const cstyle = this.autoStyle('table-cell', '', props);
        x += '<table:table-cell table:style-name="' + cstyle + '" office:value-type="string"' + (cs > 1 ? ' table:number-columns-spanned="' + cs + '"' : '') + (rs > 1 ? ' table:number-rows-spanned="' + rs + '"' : '') + '>' + (await this.blocks(cell)) + '</table:table-cell>';
        for (let k = 1; k < cs; k++) { x += '<table:covered-table-cell/>'; covered.delete(r + ':' + (c + k)); }
        c += cs;
      }
      x += '</table:table-row>';
      r++;
    }
    return x + '</table:table>';
  }

  async blocks(parent) {
    let out = '';
    const list = [];
    parent.forEach((n) => list.push(n));
    for (let i = 0; i < list.length; i++) {
      const n = list[i];
      if (n.type === nodes.paragraph && effectiveNumPr(n.attrs)) {
        // agrupa itens de lista consecutivos da mesma lista
        const np = effectiveNumPr(n.attrs);
        const group = [];
        while (i < list.length && list[i].type === nodes.paragraph && effectiveNumPr(list[i].attrs) && effectiveNumPr(list[i].attrs).numId === np.numId) group.push(list[i++]);
        i--;
        out += await this.listXmlFor(group, np.numId);
        continue;
      }
      if (n.type === nodes.paragraph) out += await this.paragraph(n);
      else if (n.type === nodes.table) out += await this.table(n);
      else if (n.type === nodes.sdt_block) out += await this.blocks(n);
      else if (n.type === nodes.opaque_block && n.attrs.text) out += '<text:p>' + esc(n.attrs.text) + '</text:p>';
    }
    return out;
  }

  async listXmlFor(group, numId) {
    const ls = this.listStyle(numId);
    let out = '<text:list text:style-name="' + ls + '">';
    let depth = 0;
    let openItem = false;
    for (const p of group) {
      const lvl = effectiveNumPr(p.attrs).ilvl;
      while (depth < lvl) { out += (openItem ? '' : '<text:list-item>') + '<text:list>'; openItem = false; depth++; }
      while (depth > lvl) { out += (openItem ? '</text:list-item>' : '') + '</text:list></text:list-item>'; openItem = true; depth--; }
      if (openItem) out += '</text:list-item>';
      out += '<text:list-item>' + (await this.paragraph(p));
      openItem = true;
    }
    if (openItem) out += '</text:list-item>';
    while (depth > 0) { out += '</text:list></text:list-item>'; depth--; }
    return out + '</text:list>';
  }
}

export async function writeOdt(doc, model, opts = {}) {
  if (canPreserveOdtSource(model, doc, opts)) return model.extra.odt.bytes.slice();
  const prev = R.model;
  R.model = model;
  try {
    const w = new OdtWriter(model);
    const body = await w.blocks(doc);
    // cabeçalho e rodapé padrão
    const hdr = model.headerFor('header', 1, false) || model.headerFor('header', 0, true);
    const ftr = model.headerFor('footer', 1, false) || model.headerFor('footer', 0, true);
    const hdrXml = hdr ? '<style:header>' + (await w.blocks(hdr.doc)) + '</style:header>' : '';
    const ftrXml = ftr ? '<style:footer>' + (await w.blocks(ftr.doc)) + '</style:footer>' : '';
    const s = model.section;
    const fontDecls = '<office:font-face-decls>' + [...w.fonts].map((f) => '<style:font-face style:name="' + esc(f) + '" svg:font-family="&apos;' + esc(f) + '&apos;"/>').join('') + '</office:font-face-decls>';
    const ns = 'xmlns:office="' + O.office + '" xmlns:style="' + O.style + '" xmlns:text="' + O.text + '" xmlns:table="' + O.table + '" xmlns:draw="' + O.draw + '" xmlns:fo="' + O.fo + '" xmlns:xlink="' + O.xlink + '" xmlns:svg="' + O.svg + '" xmlns:dc="' + O.dc + '" xmlns:meta="' + O.meta + '" office:version="1.3"';
    const def = model.styles.defaults;
    const defFont = def.r.rFonts ? fontName(def.r.rFonts.a) : 'Calibri';
    if (defFont) w.fonts.add(defFont);
    const defSize = def.r.sz ? def.r.sz.val / 2 : 11;
    const content = '<?xml version="1.0" encoding="UTF-8"?>\n<office:document-content ' + ns + '>' + fontDecls
      + '<office:automatic-styles>' + w.listXml.join('') + w.autoXml.join('') + '</office:automatic-styles>'
      + '<office:body><office:text>' + body + '</office:text></office:body></office:document-content>';
    const styles = '<?xml version="1.0" encoding="UTF-8"?>\n<office:document-styles ' + ns + '>' + fontDecls
      + '<office:styles><style:default-style style:family="paragraph"><style:paragraph-properties style:tab-stop-distance="' + cm(model.settings.defaultTabStop || 708) + '"/><style:text-properties style:font-name="' + esc(defFont || 'Calibri') + '" fo:font-size="' + defSize + 'pt" fo:language="pt" fo:country="BR"/></style:default-style>'
      + w.namedXml.join('') + '</office:styles>'
      + '<office:automatic-styles><style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="' + cm(s.pgSz.w) + '" fo:page-height="' + cm(s.pgSz.h) + '" style:print-orientation="' + (s.pgSz.w > s.pgSz.h ? 'landscape' : 'portrait') + '" fo:margin-top="' + cm(hdr ? s.pgMar.header : s.pgMar.top) + '" fo:margin-bottom="' + cm(ftr ? s.pgMar.footer : s.pgMar.bottom) + '" fo:margin-left="' + cm(s.pgMar.left) + '" fo:margin-right="' + cm(s.pgMar.right) + '"/>'
      + (hdr ? '<style:header-style><style:header-footer-properties fo:min-height="0cm" fo:margin-bottom="' + cm(Math.max(0, s.pgMar.top - s.pgMar.header - 300)) + '"/></style:header-style>' : '')
      + (ftr ? '<style:footer-style><style:header-footer-properties fo:min-height="0cm" fo:margin-top="' + cm(Math.max(0, s.pgMar.bottom - s.pgMar.footer - 300)) + '"/></style:footer-style>' : '')
      + '</style:page-layout></office:automatic-styles>'
      + '<office:master-styles><style:master-page style:name="Standard" style:page-layout-name="pm1">' + hdrXml + ftrXml + '</style:master-page></office:master-styles></office:document-styles>';
    const now = new Date().toISOString().replace(/\.\d+Z$/, '');
    const meta = '<?xml version="1.0" encoding="UTF-8"?>\n<office:document-meta ' + ns + '><office:meta><meta:generator>Redator</meta:generator>'
      + (model.properties.title ? '<dc:title>' + esc(model.properties.title) + '</dc:title>' : '')
      + (model.properties.author ? '<meta:initial-creator>' + esc(model.properties.author) + '</meta:initial-creator><dc:creator>' + esc(model.properties.author) + '</dc:creator>' : '')
      + '<dc:date>' + now + '</dc:date></office:meta></office:document-meta>';
    const zip = new JSZip();
    // modelo (.ott) grava com o tipo de modelo, para continuar abrindo como modelo
    const mime = opts.template ? 'application/vnd.oasis.opendocument.text-template' : 'application/vnd.oasis.opendocument.text';
    zip.file('mimetype', mime, { compression: 'STORE' });
    let manifest = '<?xml version="1.0" encoding="UTF-8"?>\n<manifest:manifest xmlns:manifest="' + O.manifest + '" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="' + mime + '"/>'
      + '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>';
    for (const p of w.pics) { zip.file(p.path, p.bytes); manifest += '<manifest:file-entry manifest:full-path="' + p.path + '" manifest:media-type="' + p.type + '"/>'; }
    manifest += '</manifest:manifest>';
    zip.file('content.xml', content);
    zip.file('styles.xml', styles);
    zip.file('meta.xml', meta);
    zip.file('META-INF/manifest.xml', manifest);
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', mimeType: mime });
  } finally {
    R.model = prev;
  }
}
