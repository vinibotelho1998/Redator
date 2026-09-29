// Montagem do editor: estado, plugins e visualização.
import { EditorState, Plugin, PluginKey, TextSelection, NodeSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { DOMSerializer, DOMParser as PMDOMParser, Slice, Fragment } from 'prosemirror-model';
import { history } from 'prosemirror-history';
import { keymap } from 'prosemirror-keymap';
import { baseKeymap } from 'prosemirror-commands';
import { dropCursor } from 'prosemirror-dropcursor';
import { gapCursor } from 'prosemirror-gapcursor';
import { columnResizing, tableEditing, fixTables } from 'prosemirror-tables';
import { schema, P_ATTR_NAMES } from './schema.js';
import { R } from './render.js';
import { wordKeymap } from './keymap.js';
import { decorationsPlugin } from './plugins/decorations.js';
import { findPlugin } from './plugins/find.js';
import { autoFormatPlugin } from './plugins/autoformat.js';
import { layoutPlugin } from './layout.js';
import { DrawingView, TabView, TableView, OpaqueView, mediaSrc } from './nodeviews.js';
import { renderBlocksJSON } from './static.js';
import { normalizeExternalHTML } from './htmlparse.js';
import { linkRangeAt } from './commands.js';

// ---------------------------------------------------------------- área de transferência

function bytesToBase64(bytes) {
  let s = '';
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) s += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
  return btoa(s);
}

function withData(spec, attrs) {
  if (Array.isArray(spec) && spec.length > 1 && spec[1] && typeof spec[1] === 'object' && !Array.isArray(spec[1]) && !(spec[1] instanceof Node)) {
    return [spec[0], { ...spec[1], ...attrs }, ...spec.slice(2)];
  }
  if (Array.isArray(spec)) return [spec[0], attrs, ...spec.slice(1)];
  return spec;
}

let clipSer = null;
export function clipboardSerializer() {
  if (clipSer) return clipSer;
  const base = DOMSerializer.fromSchema(schema);
  const nodes = { ...base.nodes };
  const marks = {};
  nodes.paragraph = (n) => withData(base.nodes.paragraph(n), { 'data-pa': JSON.stringify(n.attrs) });
  nodes.table = (n) => withData(base.nodes.table(n), { 'data-ta': JSON.stringify(n.attrs) });
  nodes.table_row = (n) => withData(base.nodes.table_row(n), { 'data-ra': JSON.stringify(n.attrs) });
  nodes.table_cell = (n) => withData(base.nodes.table_cell(n), { 'data-ca': JSON.stringify(n.attrs) });
  nodes.drawing = (n) => {
    const a = n.attrs;
    let src = mediaSrc(a);
    if (a.mediaPath && R.model) {
      const m = R.model.media.get(a.mediaPath);
      if (m && m.bytes && m.bytes.length < 8 * 1024 * 1024 && !String(m.type).includes('emf') && !String(m.type).includes('wmf')) src = 'data:' + m.type + ';base64,' + bytesToBase64(m.bytes);
    }
    const w = (a.cx || 0) / 9525, h = (a.cy || 0) / 9525;
    const { src: _s, ...rest } = a;
    return ['img', { src, width: Math.round(w), height: Math.round(h), alt: a.alt || '', 'data-da': JSON.stringify(rest) }];
  };
  for (const [name, fn] of Object.entries(base.marks)) {
    marks[name] = (mark, inline) => {
      const spec = fn(mark, inline);
      const a = mark.attrs;
      let data;
      if ('val' in a && Object.keys(a).length === 1) data = { 'data-m': name, 'data-v': typeof a.val === 'boolean' ? (a.val ? '1' : '0') : String(a.val) };
      else data = { 'data-m': name, 'data-ma': JSON.stringify(a) };
      return withData(spec, data);
    };
  }
  clipSer = new DOMSerializer(nodes, marks);
  return clipSer;
}

function transformPastedHTML(html) {
  if (html.includes('data-pm-slice')) return html;
  const doc = new DOMParser().parseFromString(html, 'text/html');
  const styles = R.model && R.model.styles;
  const code = styles && (styles.findByName('Code') || styles.findByName('HTML Preformatted'));
  const quote = styles && (styles.findByName('Quote') || styles.findByName('Block Text'));
  normalizeExternalHTML(doc.body, { codeStyle: code ? code.id : null, quoteStyle: quote ? quote.id : null });
  return doc.body.innerHTML;
}

// Resolve listas vindas de HTML e estilos inexistentes depois de colar
export function resolvePastedLists(tr, model, listMap) {
  const lists = listMap || new Map();
  let changed = false;
  tr.doc.descendants((node, pos) => {
    if (node.type !== schema.nodes.paragraph) return true;
    const a = node.attrs;
    let next = null;
    if (a.numId && String(a.numId).startsWith('__')) {
      const key = String(a.numId).replace(/@\d+$/, '');
      if (!lists.has(key)) {
        const kind = key.slice(2).split(':')[0] === 'number' ? 'number' : 'bullet';
        let id = model.numbering.createList(kind, 0);
        const start = /@(\d+)$/.exec(String(a.numId));
        if (start && Number(start[1]) !== 1) {
          id = model.numbering.restartList(id);
          const num = model.numbering.nums.get(id);
          if (num) { const ov = num.overrides.get(0); if (ov) ov.startOverride = Number(start[1]); }
        }
        lists.set(key, id);
      }
      next = { ...(next || a), numId: lists.get(key) };
      const lp = model.styles.findByName('List Paragraph');
      if (lp && !next.styleId) next.styleId = lp.id;
    }
    if (a.styleId && !model.styles.get(a.styleId)) next = { ...(next || a), styleId: null };
    if (next) { tr.setNodeMarkup(pos, null, next); changed = true; }
    return false;
  });
  return changed;
}

const pasteKey = new PluginKey('pasteFix');

function pasteFixPlugin() {
  return new Plugin({
    key: pasteKey,
    appendTransaction(trs, oldState, newState) {
      if (!trs.some((t) => t.docChanged && (t.getMeta('paste') || t.getMeta('uiEvent') === 'paste' || t.getMeta('uiEvent') === 'drop' || t.getMeta('fixLists')))) return null;
      if (!R.model) return null;
      const tr = newState.tr;
      const changed = resolvePastedLists(tr, R.model);
      return changed ? tr.setMeta('refreshDecorations', true) : null;
    },
  });
}

// ---------------------------------------------------------------- imagens coladas ou arrastadas

function imageFiles(list) {
  return Array.from(list || []).filter((f) => f.type && f.type.startsWith('image/'));
}

// ---------------------------------------------------------------- criação

export function createState(doc, app, opts = {}) {
  const plugins = [
    history({ depth: 500 }),
    autoFormatPlugin(() => app.prefs),
    keymap(wordKeymap()),
    keymap(baseKeymap),
    dropCursor({ color: '#185abd', width: 2 }),
    gapCursor(),
    columnResizing({ View: TableView, cellMinWidth: 16, handleWidth: 5 }),
    tableEditing(),
    decorationsPlugin(),
    findPlugin(),
    pasteFixPlugin(),
  ];
  if (opts.layout) plugins.push(layoutPlugin(opts.layout));
  let state = EditorState.create({ doc, plugins });
  const fix = fixTables(state);
  if (fix) state = state.apply(fix.setMeta('addToHistory', false));
  return state;
}

export function createView(mount, state, app, opts = {}) {
  const view = new EditorView(mount, {
    state,
    attributes: { class: 'doc-body pm-root', spellcheck: app.prefs.spellcheck === false ? 'false' : 'true', lang: 'pt-BR' },
    nodeViews: {
      drawing: (n, v, g) => new DrawingView(n, v, g, renderBlocksJSON),
      tab: (n) => new TabView(n),
      opaque_inline: (n) => new OpaqueView(n),
    },
    dispatchTransaction(tr) {
      const ns = view.state.apply(tr);
      view.updateState(ns);
      app.onTransaction(view, tr);
    },
    clipboardSerializer: clipboardSerializer(),
    transformPastedHTML,
    editable: () => !app.readOnly,
    handlePaste(v, event, slice) {
      const files = imageFiles(event.clipboardData && event.clipboardData.files);
      const html = event.clipboardData && event.clipboardData.getData('text/html');
      if (files.length && !html) {
        app.insertImageFiles(files, v);
        return true;
      }
      // colando em parágrafo vazio: o primeiro parágrafo colado mantém o próprio estilo
      const { $from, empty } = v.state.selection;
      const first = slice && slice.content.firstChild;
      if (empty && slice.openStart > 0 && first && first.type === schema.nodes.paragraph && $from.parent.type === schema.nodes.paragraph
        && $from.parent.content.size === 0 && P_ATTR_NAMES.some((k) => first.attrs[k] != null && k !== 'markRPr')) {
        const tr = v.state.tr.replaceWith($from.before(), $from.after(), slice.content);
        v.dispatch(tr.setMeta('paste', true).setMeta('uiEvent', 'paste').scrollIntoView());
        return true;
      }
      return false;
    },
    handleDrop(v, event, slice, moved) {
      if (moved) return false;
      const files = Array.from((event.dataTransfer && event.dataTransfer.files) || []);
      if (!files.length) return false;
      const imgs = imageFiles(files);
      if (imgs.length) {
        const pos = v.posAtCoords({ left: event.clientX, top: event.clientY });
        if (pos) v.dispatch(v.state.tr.setSelection(TextSelection.near(v.state.doc.resolve(pos.pos))));
        app.insertImageFiles(imgs, v);
        return true;
      }
      return false;
    },
    handleClick(v, pos, event) {
      if (!(event.ctrlKey || event.metaKey)) return false;
      const r = linkRangeAt(v.state, pos);
      if (!r) return false;
      app.followLink(r.mark.attrs);
      return true;
    },
    handleDOMEvents: {
      mouseover(v, event) {
        const a = event.target.closest && event.target.closest('a.lnk');
        v.dom.classList.toggle('link-hover', !!a && event.ctrlKey);
        return false;
      },
      focus() { app.onEditorFocus(view); return false; },
    },
  });
  view.redator = app;
  return view;
}
