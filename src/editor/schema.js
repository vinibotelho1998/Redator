// Esquema do documento, espelhando o modelo do WordprocessingML.
import { Schema } from 'prosemirror-model';
import {
  R, paraDirectCss, paraClasses, fontCss, colorCss, underlineCss, shdCss, highlightCss,
  cellDirectCss, rowDirectCss, tableDirectCss, tableClasses, drawingBoxCss, symChar,
  tabInfo, needsTabInfo, isDarkShd, calloutInfo,
} from './render.js';
import { paraAttrsFromHTML, marksFromStyleHTML, cssLengthToTwip } from './htmlparse.js';

const attr = (def = null) => ({ default: def });

export const P_ATTR_NAMES = [
  'styleId', 'jc', 'ind', 'spacing', 'numId', 'ilvl', 'keepNext', 'keepLines', 'pageBreakBefore',
  'widowControl', 'contextualSpacing', 'pBdr', 'shd', 'tabs', 'outlineLvl', 'markRPr', 'sectPr', 'extra', 'mdFence', 'md',
];

const P_ATTRS = Object.fromEntries(P_ATTR_NAMES.map((n) => [n, attr()]));

function readData(dom, name) {
  const raw = dom.getAttribute(name);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch { return null; }
}

function drawingFromClipboard(dom) {
  const a = readData(dom, 'data-da') || {};
  const sourceId = R.model && R.model.clipboardId;
  // Um caminho como word/media/image1.png não identifica uma imagem fora do
  // pacote original. XML e relacionamentos antigos também não são portáveis.
  const foreign = !a.sourceId || a.sourceId !== sourceId;
  return {
    ...a,
    ...(foreign ? { mediaPath: null, rId: null, xml: null, srcPart: null, ocx: null, ocy: null } : {}),
    sourceId: null, src: dom.getAttribute('src'),
  };
}

const nodes = {
  doc: { content: 'block+', attrs: { pageProperties: attr() } },

  paragraph: {
    group: 'block',
    content: 'inline*',
    attrs: P_ATTRS,
    parseDOM: [
      { tag: 'p[data-pa]', priority: 90, getAttrs: (dom) => readData(dom, 'data-pa') || {} },
      { tag: 'p', getAttrs: (dom) => paraAttrsFromHTML(dom) },
      { tag: 'h1', getAttrs: (dom) => ({ ...paraAttrsFromHTML(dom), styleId: R.headingStyleId(1) }) },
      { tag: 'h2', getAttrs: (dom) => ({ ...paraAttrsFromHTML(dom), styleId: R.headingStyleId(2) }) },
      { tag: 'h3', getAttrs: (dom) => ({ ...paraAttrsFromHTML(dom), styleId: R.headingStyleId(3) }) },
      { tag: 'h4', getAttrs: (dom) => ({ ...paraAttrsFromHTML(dom), styleId: R.headingStyleId(4) }) },
      { tag: 'h5', getAttrs: (dom) => ({ ...paraAttrsFromHTML(dom), styleId: R.headingStyleId(5) }) },
      { tag: 'h6', getAttrs: (dom) => ({ ...paraAttrsFromHTML(dom), styleId: R.headingStyleId(6) }) },
      { tag: 'div.p', getAttrs: (dom) => paraAttrsFromHTML(dom) },
    ],
    toDOM(node) {
      const attrs = { class: paraClasses(node) };
      const css = paraDirectCss(node.attrs);
      if (css) attrs.style = css;
      if (needsTabInfo(node)) attrs['data-tab'] = tabInfo(node.attrs);
      const md = node.attrs.md;
      if (md && md.callout && md.head) {
        const ci = calloutInfo(md.callout);
        attrs['data-icon'] = ci.icon;
        if (!node.content.size) attrs['data-label'] = ci.label;
      }
      return ['p', attrs, 0];
    },
  },

  table: {
    group: 'block',
    content: 'table_row+',
    tableRole: 'table',
    isolating: true,
    attrs: {
      styleId: attr(), tblW: attr(), jc: attr(), tblInd: attr(), tblBorders: attr(), shd: attr(),
      tblLayout: attr(), tblCellMar: attr(), tblLook: attr(), grid: attr(), extra: attr(), tblpPr: attr(),
    },
    parseDOM: [
      { tag: 'table[data-ta]', priority: 90, getAttrs: (dom) => readData(dom, 'data-ta') || {} },
      { tag: 'table', getAttrs: () => ({ styleId: R.tableGridStyleId() }) },
    ],
    toDOM(node) {
      const cols = (node.attrs.grid || []).map((w) => ['col', { style: 'width:' + (w / 15).toFixed(2) + 'px' }]);
      return ['table', { class: tableClasses(node), style: tableDirectCss(node.attrs) }, ['colgroup', ...cols], ['tbody', 0]];
    },
  },

  table_row: {
    content: 'table_cell+',
    tableRole: 'row',
    attrs: { height: attr(), hRule: attr(), header: attr(), cantSplit: attr(), extra: attr(), gridBefore: attr(), gridAfter: attr() },
    parseDOM: [
      { tag: 'tr[data-ra]', priority: 90, getAttrs: (dom) => readData(dom, 'data-ra') || {} },
      { tag: 'tr' },
    ],
    toDOM(node) {
      const css = rowDirectCss(node.attrs);
      return ['tr', css ? { style: css } : {}, 0];
    },
  },

  table_cell: {
    content: 'block+',
    tableRole: 'cell',
    isolating: true,
    attrs: {
      colspan: attr(1), rowspan: attr(1), colwidth: attr(),
      tcW: attr(), shd: attr(), tcBorders: attr(), vAlign: attr(), tcMar: attr(), noWrap: attr(),
      textDirection: attr(), extra: attr(), phantom: attr(),
    },
    parseDOM: [
      { tag: 'td[data-ca]', priority: 90, getAttrs: (dom) => cellAttrsFromDOM(dom, readData(dom, 'data-ca')) },
      { tag: 'th[data-ca]', priority: 90, getAttrs: (dom) => cellAttrsFromDOM(dom, readData(dom, 'data-ca')) },
      { tag: 'td', getAttrs: (dom) => cellAttrsFromDOM(dom, null) },
      { tag: 'th', getAttrs: (dom) => cellAttrsFromDOM(dom, null) },
    ],
    toDOM(node) {
      const a = { class: (node.attrs.phantom ? 'tc phantom' : 'tc') + (node.attrs.shd && isDarkShd(node.attrs.shd) ? ' dark-bg' : '') };
      if (node.attrs.colspan > 1) a.colspan = node.attrs.colspan;
      if (node.attrs.rowspan > 1) a.rowspan = node.attrs.rowspan;
      const css = cellDirectCss(node.attrs);
      if (css) a.style = css;
      return ['td', a, 0];
    },
  },

  sdt_block: {
    group: 'block',
    content: 'block+',
    attrs: { pr: attr(), endPr: attr() },
    parseDOM: [{ tag: 'div[data-sdt]', getAttrs: (dom) => readData(dom, 'data-sdt') || {} }],
    toDOM(node) {
      return ['div', { class: 'sdt-block', 'data-sdt': JSON.stringify({ pr: node.attrs.pr, endPr: node.attrs.endPr }) }, 0];
    },
  },

  opaque_block: {
    group: 'block',
    atom: true,
    selectable: true,
    attrs: { xml: attr(), kind: attr('bloco'), text: attr('') },
    parseDOM: [{ tag: 'div[data-ob]', getAttrs: (dom) => readData(dom, 'data-ob') || {} }],
    toDOM(node) {
      return ['div', { class: 'opaque-block' + (node.attrs.kind === 'md' ? ' ob-md' : ''), contenteditable: 'false', 'data-ob': JSON.stringify(node.attrs) },
        node.attrs.text ? node.attrs.text : '[' + (node.attrs.kind || 'conteúdo') + ']'];
    },
  },

  text: { group: 'inline' },

  tab: {
    group: 'inline',
    inline: true,
    atom: true,
    selectable: false,
    attrs: { ptab: attr() },
    parseDOM: [{ tag: 'span.tab' }],
    toDOM: () => ['span', { class: 'tab', contenteditable: 'false' }],
    leafText: () => '\t',
  },

  hard_break: {
    group: 'inline',
    inline: true,
    selectable: false,
    attrs: { type: attr(), clear: attr() },
    parseDOM: [
      { tag: 'span.pgbrk', getAttrs: () => ({ type: 'page' }) },
      { tag: 'br.colbrk', getAttrs: () => ({ type: 'column' }) },
      { tag: 'br', getAttrs: (dom) => (dom.classList.contains('ProseMirror-trailingBreak') ? false : {}) },
    ],
    toDOM(node) {
      if (node.attrs.type === 'page') return ['span', { class: 'pgbrk', contenteditable: 'false' }, ['br']];
      if (node.attrs.type === 'column') return ['br', { class: 'colbrk' }];
      return ['br'];
    },
    leafText: () => '\n',
  },

  drawing: {
    group: 'inline',
    inline: true,
    atom: true,
    draggable: true,
    attrs: {
      kind: attr('image'),     // image | shape | chart | smartart | object | other
      src: attr(),              // url para exibição
      mediaPath: attr(),        // caminho da mídia no pacote
      rId: attr(),
      cx: attr(0), cy: attr(0), // tamanho em EMU
      ocx: attr(), ocy: attr(), // tamanho original para saber se mudou
      alt: attr(''), title: attr(''), name: attr(''),
      anchor: attr(),           // dados de posicionamento quando flutuante
      crop: attr(),
      xml: attr(),              // XML original preservado
      content: attr(),          // blocos de caixa de texto (JSON)
      fill: attr(), line: attr(),
      ext: attr(),              // extensão original da imagem
      srcPart: attr(),          // parte de origem, para validar ids de relacionamento
      sourceId: attr(),         // identidade de sessão da origem do clipboard
    },
    parseDOM: [
      { tag: 'img[data-da]', priority: 90, getAttrs: drawingFromClipboard },
      {
        tag: 'img[src]',
        getAttrs(dom) {
          const w = parseFloat(dom.getAttribute('width')) || dom.naturalWidth || 0;
          const h = parseFloat(dom.getAttribute('height')) || dom.naturalHeight || 0;
          return { src: dom.getAttribute('src'), alt: dom.getAttribute('alt') || '', cx: Math.round(w * 9525), cy: Math.round(h * 9525) };
        },
      },
    ],
    toDOM(node) {
      const a = node.attrs;
      if (a.kind === 'image' || (a.src && a.kind !== 'shape')) {
        return ['img', { src: a.src || '', alt: a.alt || '', style: drawingBoxCss(a), class: 'drw' }];
      }
      return ['span', { class: 'drw-shape', style: drawingBoxCss(a) }, a.alt || ''];
    },
  },

  field_char: {
    group: 'inline',
    inline: true,
    atom: true,
    selectable: false,
    attrs: { type: attr('begin'), xml: attr() },
    parseDOM: [{ tag: 'span.fc', getAttrs: (dom) => readData(dom, 'data-fa') || {} }],
    toDOM: (node) => ['span', { class: 'fc fc-' + node.attrs.type, contenteditable: 'false', 'data-fa': JSON.stringify(node.attrs) }],
    leafText: () => '',
  },

  instr_text: {
    group: 'inline',
    inline: true,
    atom: true,
    selectable: false,
    attrs: { text: attr('') },
    parseDOM: [{ tag: 'span.fi', getAttrs: (dom) => ({ text: dom.getAttribute('data-instr') || '' }) }],
    toDOM: (node) => ['span', { class: 'fi', contenteditable: 'false', 'data-instr': node.attrs.text }],
    leafText: () => '',
  },

  opaque_inline: {
    group: 'inline',
    inline: true,
    atom: true,
    attrs: { xml: attr(), level: attr('p'), kind: attr('item'), text: attr(''), id: attr() },
    parseDOM: [{ tag: 'span.oi', getAttrs: (dom) => readData(dom, 'data-oa') || {} }],
    toDOM(node) {
      const k = node.attrs.kind;
      return ['span', { class: 'oi oi-' + k, contenteditable: 'false', 'data-oa': JSON.stringify(node.attrs) }, node.attrs.text || ''];
    },
    leafText: (node) => (node.attrs.kind === 'del' ? '' : node.attrs.text || ''),
  },

  footnote_ref: {
    group: 'inline',
    inline: true,
    atom: true,
    attrs: { id: attr(), kind: attr('footnote'), custom: attr() },
    parseDOM: [{ tag: 'sup.fnref', getAttrs: (dom) => readData(dom, 'data-fn') || {} }],
    toDOM: (node) => ['sup', { class: 'fnref fnref-' + node.attrs.kind, contenteditable: 'false', 'data-fn': JSON.stringify(node.attrs) }],
    leafText: () => '',
  },

  sym: {
    group: 'inline',
    inline: true,
    atom: true,
    attrs: { font: attr(), char: attr() },
    parseDOM: [{ tag: 'span.sym', getAttrs: (dom) => readData(dom, 'data-sym') || {} }],
    toDOM(node) {
      const { text, font } = symChar(node.attrs.font, node.attrs.char);
      return ['span', { class: 'sym', style: font ? 'font-family:"' + font + '"' : '', 'data-sym': JSON.stringify(node.attrs) }, text];
    },
    leafText: (node) => symChar(node.attrs.font, node.attrs.char).text,
  },
};

function cellAttrsFromDOM(dom, data) {
  const base = data || {};
  const colspan = Number(dom.getAttribute('colspan') || 1);
  const rowspan = Number(dom.getAttribute('rowspan') || 1);
  const out = { ...base, colspan, rowspan };
  if (!data) {
    const w = dom.getAttribute('data-colwidth');
    if (w) out.colwidth = w.split(',').map(Number);
    const bg = dom.style.backgroundColor;
    if (bg) {
      const m = /rgb\((\d+),\s*(\d+),\s*(\d+)\)/.exec(bg);
      if (m) out.shd = { val: 'clear', color: 'auto', fill: [m[1], m[2], m[3]].map((x) => Number(x).toString(16).padStart(2, '0')).join('').toUpperCase() };
    }
    const va = dom.style.verticalAlign;
    if (va === 'middle') out.vAlign = 'center';
    else if (va === 'bottom') out.vAlign = 'bottom';
  }
  return out;
}

const toggle = (name, cssOn, cssOff) => ({
  attrs: { val: attr(true) },
  parseDOM: [{ tag: 'span[data-m="' + name + '"]', priority: 90, getAttrs: (dom) => ({ val: dom.getAttribute('data-v') !== '0' }) }],
  toDOM: (m) => ['span', { style: m.attrs.val ? cssOn : cssOff }, 0],
});

const marks = {
  hyperlink: {
    attrs: { href: attr(), anchor: attr(), a: attr() },
    inclusive: false,
    parseDOM: [{
      tag: 'a[href]',
      getAttrs(dom) {
        const data = readData(dom, 'data-ma');
        if (data) return data;
        const href = dom.getAttribute('href');
        if (href && href.startsWith('#')) return { anchor: href.slice(1) };
        return { href };
      },
    }],
    toDOM(m) {
      const target = m.attrs.href || (m.attrs.anchor ? '#' + m.attrs.anchor : '');
      const tip = (m.attrs.a && m.attrs.a.tooltip) || target;
      return ['a', { href: target, class: 'lnk', title: tip + '\nCtrl+Clique para seguir o link' }, 0];
    },
  },
  sdt: {
    attrs: { pr: attr(), endPr: attr() },
    inclusive: true,
    parseDOM: [{ tag: 'span[data-m="sdt"]', priority: 90, getAttrs: (dom) => readData(dom, 'data-ma') || {} }],
    toDOM: () => ['span', { class: 'sdt' }, 0],
  },
  ins: {
    attrs: { a: attr(), tag: attr('ins') },
    inclusive: false,
    parseDOM: [{ tag: 'ins[data-ma]', getAttrs: (dom) => readData(dom, 'data-ma') || {} }],
    toDOM(m) {
      const a = m.attrs.a || {};
      const who = a.author ? 'Inserido por ' + a.author : 'Inserção';
      return ['ins', { class: 'trk-ins', title: who + (a.date ? ' em ' + new Date(a.date).toLocaleString('pt-BR') : '') }, 0];
    },
  },
  rStyle: {
    attrs: { val: attr() },
    parseDOM: [{ tag: 'span[data-m="rStyle"]', priority: 90, getAttrs: (dom) => ({ val: dom.getAttribute('data-v') }) }],
    toDOM: (m) => ['span', { class: R.charClass(m.attrs.val) }, 0],
  },
  rFonts: {
    attrs: { a: attr() },
    parseDOM: [
      { tag: 'span[data-m="rFonts"]', priority: 90, getAttrs: (dom) => readData(dom, 'data-ma') || {} },
      { style: 'font-family', getAttrs: (v) => marksFromStyleHTML.fontFamily(v) },
      { tag: 'code', getAttrs: () => ({ a: { ascii: 'Consolas', hAnsi: 'Consolas', cs: 'Consolas' } }) },
      { tag: 'kbd', getAttrs: () => ({ a: { ascii: 'Consolas', hAnsi: 'Consolas', cs: 'Consolas' } }) },
    ],
    toDOM: (m) => ['span', { style: fontCss(m.attrs.a) }, 0],
  },
  b: {
    ...toggle('b', 'font-weight:bold', 'font-weight:normal'),
    parseDOM: [
      { tag: 'span[data-m="b"]', priority: 90, getAttrs: (dom) => ({ val: dom.getAttribute('data-v') !== '0' }) },
      { tag: 'strong' },
      { tag: 'b', getAttrs: (dom) => dom.style.fontWeight !== 'normal' && null },
      { style: 'font-weight', getAttrs: (v) => /^(bold(er)?|[6-9]\d\d)$/.test(v) && null },
    ],
  },
  i: {
    ...toggle('i', 'font-style:italic', 'font-style:normal'),
    parseDOM: [
      { tag: 'span[data-m="i"]', priority: 90, getAttrs: (dom) => ({ val: dom.getAttribute('data-v') !== '0' }) },
      { tag: 'i' }, { tag: 'em' }, { tag: 'cite' },
      { style: 'font-style', getAttrs: (v) => (v === 'italic' || v === 'oblique') && null },
    ],
  },
  caps: toggle('caps', 'text-transform:uppercase', 'text-transform:none'),
  smallCaps: {
    ...toggle('smallCaps', 'font-variant:small-caps', 'font-variant:normal'),
    parseDOM: [
      { tag: 'span[data-m="smallCaps"]', priority: 90, getAttrs: (dom) => ({ val: dom.getAttribute('data-v') !== '0' }) },
      { style: 'font-variant', getAttrs: (v) => v === 'small-caps' && null },
    ],
  },
  strike: {
    ...toggle('strike', 'text-decoration-line:line-through', ''),
    parseDOM: [
      { tag: 'span[data-m="strike"]', priority: 90, getAttrs: (dom) => ({ val: dom.getAttribute('data-v') !== '0' }) },
      { tag: 's' }, { tag: 'strike' }, { tag: 'del' },
      { style: 'text-decoration', getAttrs: (v) => /line-through/.test(v) && null },
      { style: 'text-decoration-line', getAttrs: (v) => /line-through/.test(v) && null },
    ],
  },
  dstrike: toggle('dstrike', 'text-decoration-line:line-through;text-decoration-style:double', ''),
  vanish: {
    attrs: { val: attr(true) },
    parseDOM: [{ tag: 'span[data-m="vanish"]', priority: 90, getAttrs: (dom) => ({ val: dom.getAttribute('data-v') !== '0' }) }],
    toDOM: (m) => ['span', { class: m.attrs.val ? 'vanish' : '' }, 0],
  },
  color: {
    attrs: { a: attr() },
    parseDOM: [
      { tag: 'span[data-m="color"]', priority: 90, getAttrs: (dom) => readData(dom, 'data-ma') || {} },
      { style: 'color', getAttrs: (v) => marksFromStyleHTML.color(v) },
      { tag: 'font[color]', getAttrs: (dom) => marksFromStyleHTML.color(dom.getAttribute('color')) },
    ],
    toDOM: (m) => ['span', { style: colorCss(m.attrs.a) }, 0],
  },
  sz: {
    attrs: { val: attr(22) },
    parseDOM: [
      { tag: 'span[data-m="sz"]', priority: 90, getAttrs: (dom) => ({ val: Number(dom.getAttribute('data-v')) || 22 }) },
      { style: 'font-size', getAttrs: (v) => marksFromStyleHTML.fontSize(v) },
    ],
    toDOM: (m) => ['span', { style: 'font-size:' + m.attrs.val / 2 + 'pt' }, 0],
  },
  highlight: {
    attrs: { val: attr('yellow') },
    parseDOM: [
      { tag: 'span[data-m="highlight"]', priority: 90, getAttrs: (dom) => ({ val: dom.getAttribute('data-v') }) },
      { tag: 'mark', getAttrs: () => ({ val: 'yellow' }) },
    ],
    toDOM: (m) => ['span', { style: highlightCss(m.attrs.val) }, 0],
  },
  u: {
    attrs: { a: attr() },
    parseDOM: [
      { tag: 'span[data-m="u"]', priority: 90, getAttrs: (dom) => readData(dom, 'data-ma') || {} },
      { tag: 'u', getAttrs: () => ({ a: { val: 'single' } }) },
      { tag: 'ins', getAttrs: (dom) => (dom.hasAttribute('data-ma') ? false : { a: { val: 'single' } }) },
      { style: 'text-decoration', getAttrs: (v) => /underline/.test(v) && { a: { val: 'single' } } },
      { style: 'text-decoration-line', getAttrs: (v) => /underline/.test(v) && { a: { val: 'single' } } },
    ],
    toDOM: (m) => ['span', { style: underlineCss(m.attrs.a) }, 0],
  },
  shd: {
    attrs: { a: attr() },
    parseDOM: [
      { tag: 'span[data-m="shd"]', priority: 90, getAttrs: (dom) => readData(dom, 'data-ma') || {} },
      { style: 'background-color', getAttrs: (v) => marksFromStyleHTML.background(v) },
    ],
    toDOM: (m) => ['span', { style: shdCss(m.attrs.a) }, 0],
  },
  vertAlign: {
    attrs: { val: attr('superscript') },
    parseDOM: [
      { tag: 'span[data-m="vertAlign"]', priority: 90, getAttrs: (dom) => ({ val: dom.getAttribute('data-v') }) },
      { tag: 'sup', getAttrs: (dom) => (dom.classList.contains('fnref') ? false : { val: 'superscript' }) },
      { tag: 'sub', getAttrs: () => ({ val: 'subscript' }) },
      { style: 'vertical-align', getAttrs: (v) => (v === 'super' ? { val: 'superscript' } : v === 'sub' ? { val: 'subscript' } : false) },
    ],
    toDOM: (m) => ['span', { class: m.attrs.val === 'superscript' ? 'sup' : m.attrs.val === 'subscript' ? 'sub' : '' }, 0],
  },
  spacing: {
    attrs: { val: attr(0) },
    parseDOM: [{ tag: 'span[data-m="spacing"]', priority: 90, getAttrs: (dom) => ({ val: Number(dom.getAttribute('data-v')) || 0 }) },
      { style: 'letter-spacing', getAttrs: (v) => { const t = cssLengthToTwip(v); return t ? { val: t } : false; } }],
    toDOM: (m) => ['span', { style: 'letter-spacing:' + m.attrs.val / 20 + 'pt' }, 0],
  },
  rprExtra: {
    attrs: { items: attr() },
    parseDOM: [{ tag: 'span[data-m="rprExtra"]', priority: 90, getAttrs: (dom) => readData(dom, 'data-ma') || {} }],
    toDOM: () => ['span', 0],
  },
};

export const schema = new Schema({ nodes, marks });

// Marcas de propriedades de execução, na ordem em que aparecem no rPr
export const RUN_MARKS = ['rStyle', 'rFonts', 'b', 'i', 'caps', 'smallCaps', 'strike', 'dstrike', 'vanish', 'color', 'spacing', 'sz', 'highlight', 'u', 'shd', 'vertAlign', 'rprExtra'];
export const WRAPPER_MARKS = ['hyperlink', 'sdt', 'ins'];
export const TOGGLE_MARKS = ['b', 'i', 'caps', 'smallCaps', 'strike', 'dstrike', 'vanish'];
