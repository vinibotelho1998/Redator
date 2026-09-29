// Redator: controlador principal da janela.
import { TextSelection, NodeSelection } from 'prosemirror-state';
import { Fragment } from 'prosemirror-model';
import { undo, redo, undoDepth, redoDepth } from 'prosemirror-history';
import { selectAll } from 'prosemirror-commands';
import {
  CellSelection, TableMap, isInTable, selectedRect, addColumnAfter, addColumnBefore, addRowAfter, addRowBefore,
  deleteColumn, deleteRow, deleteTable, mergeCells, splitCell,
} from 'prosemirror-tables';
import { schema } from './editor/schema.js';
import { R, effectiveNumPr, fontName } from './editor/render.js';
import * as C from './editor/commands.js';
import { createState, createView } from './editor/setup.js';
import { PageLayout } from './editor/layout.js';
import { setSearch, clearSearch, findNext, gotoMatch, replaceCurrent, replaceAll, findKey } from './editor/plugins/find.js';
import { columnWidths } from './editor/nodeviews.js';
import { readDocx, convertPendingMedia, MIME, NEEDS_CONVERSION } from './docx/reader.js';
import { writeDocx, newHeaderFooter } from './docx/writer.js';
import { blankDocxBytes } from './docx/template.js';
import { readMarkdown, writeMarkdown, bindMarkdownSource } from './formats/markdown.js';
import { readHtml, writeHtml } from './formats/html.js';
import { readText, writeText } from './formats/text.js';
import { readOdt, writeOdt } from './formats/odt.js';
import { readRtf, writeRtf } from './formats/rtf.js';
import { readDoc } from './formats/doc.js';
import { readPdf } from './formats/pdf.js';
import { readCsv, writeCsv } from './formats/csv.js';
import { loadImageSize, fitSize, contentWidthPx, dirName, encodeText } from './formats/common.js';
import { PAPER_SIZES, MARGIN_PRESETS } from './model.js';
import { h } from './ui/dom.js';
import { icon } from './ui/icons.js';
import { Ribbon } from './ui/ribbon.js';
import { StatusBar } from './ui/statusbar.js';
import { Ruler } from './ui/ruler.js';
import { NavPane } from './ui/navpane.js';
import { Backstage } from './ui/backstage.js';
import { openContextMenu } from './ui/contextmenu.js';
import { openMenu, closeAllPopups, tooltip } from './ui/popup.js';
import * as D from './ui/dialogs.js';
import { getHost } from './host.js';
import { runSelfTest, testApi } from './selftest.js';
import { cmToTwip, twipToCm, fmtNumber } from './util/units.js';

const FORMAT_INFO = {
  docx: { label: 'Documento do Word', filters: [{ name: 'Documento do Word', extensions: ['docx'] }] },
  odt: { label: 'Texto OpenDocument', filters: [{ name: 'Texto OpenDocument', extensions: ['odt'] }] },
  md: { label: 'Markdown', filters: [{ name: 'Markdown', extensions: ['md', 'markdown'] }] },
  html: { label: 'Página da Web', filters: [{ name: 'Página da Web', extensions: ['html', 'htm'] }] },
  txt: { label: 'Texto sem Formatação', filters: [{ name: 'Texto sem Formatação', extensions: ['txt'] }] },
  csv: { label: 'CSV (separado por vírgulas)', filters: [{ name: 'CSV', extensions: ['csv'] }] },
  rtf: { label: 'Rich Text Format', filters: [{ name: 'Rich Text Format', extensions: ['rtf'] }] },
  pdf: { label: 'PDF', filters: [{ name: 'PDF', extensions: ['pdf'] }] },
};

const EXT_FORMAT = { docx: 'docx', docm: 'docx', dotx: 'docx', dotm: 'docx', odt: 'odt', md: 'md', markdown: 'md', txt: 'txt', text: 'txt', log: 'txt', csv: 'csv', tsv: 'csv', html: 'html', htm: 'html', xhtml: 'html', rtf: 'rtf', doc: 'doc', dot: 'doc', pdf: 'pdf' };

// arquivos de texto e de código abrem como texto puro, sem mexer nos bytes
export const TEXT_EXTS = ['txt', 'text', 'log', 'json', 'jsonc', 'json5', 'xml', 'xsd', 'xsl', 'svg', 'yaml', 'yml', 'toml', 'ini', 'cfg', 'conf', 'config', 'properties', 'env',
  'sql', 'ps1', 'psm1', 'psd1', 'bat', 'cmd', 'sh', 'bash', 'zsh', 'py', 'js', 'mjs', 'cjs', 'ts', 'tsx', 'jsx', 'vue', 'css', 'scss', 'sass', 'less',
  'java', 'kt', 'kts', 'groovy', 'gradle', 'c', 'h', 'cpp', 'cc', 'hpp', 'cs', 'go', 'rs', 'php', 'rb', 'swift', 'vb', 'vbs', 'lua', 'r', 'pl', 'dart', 'scala',
  'srt', 'vtt', 'sub', 'rst', 'adoc', 'asciidoc', 'org', 'tex', 'bib', 'diff', 'patch', 'nfo', 'reg', 'tsv', 'gitignore', 'gitattributes', 'editorconfig', 'dockerfile'];
for (const e of TEXT_EXTS) if (!EXT_FORMAT[e]) EXT_FORMAT[e] = 'txt';

const FALLBACK_FONTS = ['Arial', 'Arial Black', 'Bahnschrift', 'Calibri', 'Calibri Light', 'Cambria', 'Cambria Math', 'Candara', 'Comic Sans MS', 'Consolas', 'Constantia', 'Corbel', 'Courier New', 'Ebrima', 'Franklin Gothic Medium', 'Gabriola', 'Gadugi', 'Georgia', 'Impact', 'Ink Free', 'Lucida Console', 'Lucida Sans Unicode', 'Malgun Gothic', 'Microsoft Sans Serif', 'MV Boli', 'Palatino Linotype', 'Segoe Print', 'Segoe Script', 'Segoe UI', 'Segoe UI Light', 'Segoe UI Semibold', 'Sitka Text', 'Sylfaen', 'Symbol', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana', 'Webdings', 'Wingdings'];

const LANGS = [['pt-BR', 'Português (Brasil)'], ['pt-PT', 'Português (Portugal)'], ['en-US', 'Inglês (Estados Unidos)'], ['en-GB', 'Inglês (Reino Unido)'], ['es-ES', 'Espanhol (Espanha)'], ['fr-FR', 'Francês (França)'], ['de-DE', 'Alemão (Alemanha)'], ['it-IT', 'Italiano (Itália)']];

function extOf(path) {
  const m = /\.([a-z0-9]+)$/i.exec(path || '');
  return m ? m[1].toLowerCase() : '';
}

function baseName(path) {
  const name = String(path || '').split(/[\\/]/).pop();
  if (/^file:/i.test(String(path))) { try { return decodeURIComponent(name); } catch { return name; } }
  return name;
}

// bytes nulos no começo, fora de UTF-16, indicam arquivo binário
function looksBinary(bytes) {
  const n = Math.min(bytes.length, 8192);
  if (n >= 2 && ((bytes[0] === 0xFF && bytes[1] === 0xFE) || (bytes[0] === 0xFE && bytes[1] === 0xFF))) return false;
  let zeros = 0;
  for (let i = 0; i < n; i++) if (bytes[i] === 0) zeros++;
  return zeros > n * 0.01;
}

function isPdf(bytes) {
  return bytes[0] === 0x25 && bytes[1] === 0x50 && bytes[2] === 0x44 && bytes[3] === 0x46;
}

function stripExt(name) {
  return name.replace(/\.[a-z0-9]+$/i, '');
}

class App {
  constructor(root) {
    this.root = root;
    this.host = getHost();
    this.prefs = {};
    this.model = null;
    this.view = null;
    this.activeView = null;
    this.filePath = null;
    this.format = 'docx';
    this.docName = 'Documento1';
    this.dirty = false;
    this.readOnly = false;
    this.pristine = true;
    this.lastColor = { val: 'FF0000' };
    this.lastHighlight = 'yellow';
    this.lastShading = null;
    this.searchOpts = { matchCase: false, wholeWord: false, regex: false };
    this.lastSearch = '';
    this.lastReplace = '';
    this.searchCount = 0;
    this.showMarks = false;
    this.showMarkup = true;
    this.viewMode = 'print';
    this.painter = null;
    this.hfEdit = null;
    this.uiState = {};
    this.wordCount = 0;
    this.systemFonts = [];
    this.symbols = [];
    this.warnedLossy = new Set();
    // AutoRecuperação: cada alteração muda o carimbo; a cópia guarda o último gravado
    this.editStamp = 0;
    this.recoverStamp = 0;
    this.lastEdit = 0;
    this.lastRecover = 0;
    this.recoverJob = null;
    this.recovered = false;
    this.saving = false;
  }

  // ------------------------------------------------------------ inicialização

  async init() {
    const info = await this.host.init();
    this.userName = info.userName || '';
    this.prefs = {
      spellcheck: true, smartQuotes: true, autoFormat: true, autoRecover: true, ruler: true, navPane: false,
      shortcuts: 'intl', units: 'cm', recentFonts: [], recentSymbols: [], zoom: 1, ...(info.prefs || {}),
    };
    this.symbols = this.prefs.recentSymbols || [];
    this.buildUI();
    this.bindEvents();
    this.host.systemFonts().then((list) => { this.systemFonts = list || []; }).catch(() => {});
    this.tryLocalFonts();
    let restored = false;
    if (info.recover) {
      try { restored = await this.restoreRecovered(info.recover); } catch (err) { console.warn('Recuperação falhou', err); }
    }
    if (restored) {
      // a janela voltou de uma falha com o documento recuperado
    } else if (info.filePath) {
      try { await this.loadPath(info.filePath); }
      catch (err) { await this.newDocument(); this.showError('Não foi possível abrir o arquivo.', err); }
    } else {
      await this.newDocument();
    }
    this.recoverTimer = setInterval(() => this.saveRecovery(), 10000);
    if (info.recovered && info.recovered.length) setTimeout(() => this.offerRecovery(info.recovered), 400);
    if (this.prefs.navPane) this.showNavPane();
    if (this.prefs.testPage != null) {
      this.layout.runNow(8);
      setTimeout(() => this.gotoPage(this.prefs.testPage), 50);
    }
    if (new URLSearchParams(location.search).get('selftest')) {
      this.layout.runNow(8);
      setTimeout(() => runSelfTest(this), 200);
    }
  }

  async tryLocalFonts() {
    try {
      if (!window.queryLocalFonts) return;
      const fonts = await window.queryLocalFonts();
      const set = new Set(this.systemFonts);
      for (const f of fonts) set.add(f.family);
      this.systemFonts = [...set].sort((a, b) => a.localeCompare(b, 'pt-BR'));
    } catch { /* sem permissão */ }
  }

  buildUI() {
    const root = this.root;
    root.className = 'app';
    // barra de título
    const qa = (ic, title, run, shortcut) => {
      const b = h('button', { class: 'qa-btn', type: 'button', html: icon(ic) });
      b.addEventListener('click', run);
      tooltip(b, title, shortcut);
      return b;
    };
    this.undoBtn = qa('undo', 'Desfazer', () => this.undo(), 'Ctrl+Z');
    this.redoBtn = qa('redo', 'Refazer', () => this.redo(), 'Ctrl+Y');
    this.saveBtn = qa('save', 'Salvar', () => this.save(), this.sc('save'));
    this.titleText = h('span', { class: 'tb-title-text' });
    this.titleState = h('span', { class: 'tb-title-state' });
    this.searchBox = h('input', { class: 'tb-search', type: 'search', placeholder: 'Pesquisar', spellcheck: 'false' });
    this.searchBox.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); this.showNavPane('search', this.searchBox.value); }
      if (e.key === 'Escape') { this.searchBox.value = ''; this.focusEditor(); }
    });
    this.titlebar = h('header', { class: 'titlebar' },
      h('div', { class: 'tb-left' }, h('span', { class: 'tb-logo' }, h('img', { src: 'icons/redator-64.png', alt: '', width: '20', height: '20' })), this.saveBtn, this.undoBtn, this.redoBtn),
      h('div', { class: 'tb-center' }, h('div', { class: 'tb-title' }, this.titleText, this.titleState),
        h('div', { class: 'tb-search-wrap' }, h('span', { class: 'tb-search-icon', html: icon('find') }), this.searchBox)),
      h('div', { class: 'tb-right' }, h('span', { class: 'tb-user', text: (this.userName || 'U').slice(0, 1).toUpperCase(), title: this.userName || '' })));
    this.ribbon = new Ribbon(this);
    this.status = new StatusBar(this);
    this.ruler = new Ruler(this);
    this.nav = new NavPane(this);
    this.backstage = new Backstage(this);
    // área do documento
    this.pagesEl = h('div', { class: 'pages-layer' });
    this.hostEl = h('div', { class: 'editor-host' });
    this.zoomEl = h('div', { class: 'doc-zoom' }, this.pagesEl, this.hostEl);
    this.canvasEl = h('div', { class: 'doc-canvas' }, this.zoomEl);
    this.scrollEl = h('div', { class: 'doc-scroll', tabindex: '-1' }, this.canvasEl);
    this.guideEl = h('div', { class: 'ruler-guide', hidden: true });
    this.docArea = h('div', { class: 'docarea' }, this.ruler.el, this.scrollEl, this.guideEl);
    this.workspace = h('div', { class: 'workspace' }, this.nav.el, this.docArea);
    this.styleEl = h('style', { id: 'doc-styles' });
    this.pageStyleEl = h('style', { id: 'page-size' });
    document.head.append(this.styleEl, this.pageStyleEl);
    root.append(this.titlebar, this.ribbon.el, this.workspace, this.status.el, this.backstage.el);
    this.layout = new PageLayout({ scroll: this.scrollEl, canvas: this.canvasEl, zoom: this.zoomEl, pages: this.pagesEl, host: this.hostEl });
    this.layout.onChange(() => this.scheduleUI());
    this.layout.onHeaderDblClick = (kind, page) => this.editHeaderFooter(kind, page);
    this.layout.zoom = this.prefs.zoom || 1;
    this.setRuler(this.prefs.ruler !== false, true);
  }

  bindEvents() {
    window.addEventListener('keydown', (e) => this.onKeyDown(e), true);
    window.addEventListener('resize', () => { this.layout.updateCanvas(); this.ruler.update(); });
    this.scrollEl.addEventListener('scroll', () => { this.scheduleUI(); });
    this.scrollEl.addEventListener('wheel', (e) => {
      if (!e.ctrlKey) return;
      e.preventDefault();
      this.stepZoom(e.deltaY < 0 ? 1 : -1, 0.1);
    }, { passive: false });
    this.scrollEl.addEventListener('dblclick', (e) => this.onDocDblClick(e), true);
    this.scrollEl.addEventListener('mousedown', (e) => this.onDocMouseDown(e), true);
    this.scrollEl.addEventListener('mouseup', () => this.onDocMouseUp());
    document.addEventListener('contextmenu', (e) => {
      if (!e.target.closest('.pm-root') && !e.target.closest('input, textarea')) e.preventDefault();
      if (!this.host.isElectron && e.target.closest('.pm-root')) {
        e.preventDefault();
        this.onContextMenu({ x: e.clientX, y: e.clientY, misspelledWord: '', suggestions: [], isEditable: true, selectionText: this.selectedText() });
      }
    });
    this.host.on('app:context-menu', (info) => this.onContextMenu(info));
    this.host.on('app:request-close', () => this.requestClose());
    document.addEventListener('dragover', (e) => { if (e.dataTransfer && [...e.dataTransfer.types].includes('Files')) e.preventDefault(); });
    document.addEventListener('drop', (e) => {
      const files = Array.from((e.dataTransfer && e.dataTransfer.files) || []);
      if (!files.length) return;
      const docs = files.filter((f) => EXT_FORMAT[extOf(f.name)]);
      if (!docs.length) return;
      e.preventDefault();
      for (const f of docs) {
        const p = this.host.pathForFile(f);
        if (p) this.openPath(p);
      }
    });
  }

  sc(name) {
    const intl = { bold: 'Ctrl+B', italic: 'Ctrl+I', underline: 'Ctrl+U', save: 'Ctrl+S', find: 'Ctrl+F', replace: 'Ctrl+H', selectAll: 'Ctrl+A', alignLeft: 'Ctrl+L', alignCenter: 'Ctrl+E', alignRight: 'Ctrl+R', justify: 'Ctrl+J', open: 'Ctrl+O', newDoc: 'Ctrl+N', print: 'Ctrl+P', font: 'Ctrl+D', redo: 'Ctrl+Y' };
    const pt = { bold: 'Ctrl+N', italic: 'Ctrl+I', underline: 'Ctrl+S', save: 'Ctrl+B', find: 'Ctrl+L', replace: 'Ctrl+U', selectAll: 'Ctrl+T', alignLeft: 'Ctrl+Q', alignCenter: 'Ctrl+E', alignRight: 'Ctrl+G', justify: 'Ctrl+J', open: 'Ctrl+A', newDoc: 'Ctrl+O', print: 'Ctrl+P', font: 'Ctrl+D', redo: 'Ctrl+R' };
    return (this.prefs.shortcuts === 'pt' ? pt : intl)[name] || '';
  }

  // ------------------------------------------------------------ documento

  isPristine() {
    return this.pristine && !this.dirty && !this.filePath;
  }

  async newDocument(kind) {
    if (!this.isPristine() && this.model) { await this.host.newWindow(null); return; }
    const bytes = await blankDocxBytes(this.prefs.author || this.userName || '');
    const { model, doc } = await readDocx(bytes);
    model.format = kind === 'md' ? 'md' : 'docx';
    let finalDoc = doc;
    if (kind === 'abnt') finalDoc = this.applyAbnt(model, doc);
    this.setDocument(model, finalDoc, { path: null, format: kind === 'md' ? 'md' : 'docx', name: 'Documento1' });
    this.pristine = true;
  }

  applyAbnt(model, doc) {
    model.section.setMargins({ top: cmToTwip(3), bottom: cmToTwip(2), left: cmToTwip(3), right: cmToTwip(2), header: cmToTwip(1.25), footer: cmToTwip(1.25) });
    model.styles.setDefaultFont('Arial', 24);
    model.styles.setDefaultParagraphSpacing(0, 0, 360);
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    void W;
    // parágrafo normal justificado com recuo de 1,25 cm
    const normal = model.styles.get('Normal');
    if (normal) { normal.pPr = { ...(normal.pPr || {}), jc: 'both', ind: { firstLine: String(cmToTwip(1.25)) } }; }
    for (const n of [1, 2, 3]) {
      const st = model.styles.get('Heading' + n);
      if (st) {
        st.pPr = { ...(st.pPr || {}), ind: { firstLine: '0' }, jc: 'left', spacing: { before: '240', after: '240', line: '360', lineRule: 'auto' } };
        st.rPr = [{ type: 'rFonts', attrs: { a: { ascii: 'Arial', hAnsi: 'Arial', cs: 'Arial' } } }, { type: 'b', attrs: { val: n <= 2 } }, { type: 'color', attrs: { a: { val: '000000' } } }, { type: 'sz', attrs: { val: 24 } }, ...(n === 1 ? [{ type: 'caps', attrs: { val: true } }] : [])];
      }
    }
    this.rewriteStylesXml(model);
    // número de página no canto superior direito
    const entry = newHeaderFooter(model, 'header');
    model.section.setReference('header', 'default', entry.rId);
    const hs = model.styles.findByName('header');
    entry.doc = schema.node('doc', null, [this.pageNumberParagraph(hs ? hs.id : null, 'right', false)]);
    return doc;
  }

  // regrava estilos alterados no XML para que sejam salvos
  rewriteStylesXml(model) {
    const st = model.styles;
    st.ensureDoc();
    const root = st.doc.documentElement;
    const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    for (const el of Array.from(root.getElementsByTagNameNS(W, 'style'))) {
      const id = el.getAttributeNS(W, 'styleId');
      const s = st.get(id);
      if (!s || !s.touched) continue;
    }
    // os estilos alterados em memória são regravados por completo
    for (const s of st.styles.values()) {
      if (!['Normal', 'Heading1', 'Heading2', 'Heading3'].includes(s.id)) continue;
      const el = Array.from(root.getElementsByTagNameNS(W, 'style')).find((x) => x.getAttributeNS(W, 'styleId') === s.id);
      if (!el) continue;
      for (const c of Array.from(el.children)) if (c.localName === 'pPr' || c.localName === 'rPr') c.remove();
      const frag = new window.DOMParser().parseFromString('<w:root xmlns:w="' + W + '">' + ppr(s.pPr) + rpr(s.rPr) + '</w:root>', 'application/xml').documentElement;
      for (const c of Array.from(frag.children)) el.appendChild(st.doc.importNode(c, true));
    }
    st.cache.clear();
    st.dirty = true;
    function ppr(p) { return p ? pPrXmlSafe(p) : ''; }
    function rpr(r) { return r ? rPrXmlSafe(r) : ''; }
  }

  async openDialog() {
    const paths = await this.host.openDialog({});
    for (const p of paths) await this.openPath(p);
  }

  async openPath(p) {
    if (!p) return;
    if (this.filePath && this.filePath.toLowerCase() === String(p).toLowerCase()) return;
    if (this.isPristine()) {
      try { await this.loadPath(p); }
      catch (err) { this.showError('Não foi possível abrir "' + baseName(p) + '".', err); }
    } else {
      await this.host.newWindow(p);
    }
  }

  async loadPath(p) {
    this.setBusy('Abrindo ' + baseName(p) + '...');
    try {
      const f = await this.host.readFile(p);
      await this.loadBytes(f.data, p, f.readonly);
    } finally {
      this.setBusy(null);
    }
  }

  async loadBytes(bytes, path, readonly) {
    const ext = extOf(path);
    let format = EXT_FORMAT[ext];
    let res;
    if (format === 'csv') {
      res = await readCsv(bytes, path);
      if (!res) format = 'txt';
    }
    if (res) { /* já lido como tabela */ }
    else if (format === 'docx') res = await readDocx(bytes);
    else if (format === 'md') res = await readMarkdown(bytes, path, { host: this.host });
    else if (format === 'html') res = await readHtml(bytes, path);
    else if (format === 'txt') {
      if (looksBinary(bytes)) throw new Error('"' + baseName(path) + '" não é um arquivo de texto.');
      res = await readText(bytes, path);
    }
    else if (format === 'odt') res = await readOdt(bytes, path);
    else if (format === 'rtf') res = await readRtf(bytes);
    else if (format === 'doc') res = await readDoc(bytes);
    else if (format === 'pdf') res = await this.readPdfBytes(bytes, path);
    else {
      // tenta descobrir pelo conteúdo
      if (bytes[0] === 0x50 && bytes[1] === 0x4b) res = await readDocx(bytes);
      else if (bytes[0] === 0xD0 && bytes[1] === 0xCF) res = await readDoc(bytes);
      else if (isPdf(bytes)) res = await this.readPdfBytes(bytes, path);
      else res = await readText(bytes, path);
    }
    // .doc antigo e PDF abrem convertidos: salvar grava um .docx novo ao lado
    const kind = format === 'doc' || format === 'pdf' ? format : !format && bytes[0] === 0xD0 && bytes[1] === 0xCF ? 'doc' : !format && isPdf(bytes) ? 'pdf' : null;
    const compat = !!kind;
    this.setDocument(res.model, res.doc, { path: compat ? null : path, format: compat ? 'docx' : format || 'txt', name: baseName(path) });
    if (compat) { this.importPath = path; this.importKind = kind; this.pristine = false; this.updateTitle(); }
    this.readOnly = !!readonly;
    if (readonly) this.setInfo('Somente leitura');
    this.host.setDocState({ filePath: path, addRecent: true });
    this.host.addRecent(path);
    if ([...res.model.media.values()].some((m) => m.pending)) {
      convertPendingMedia(res.model, this.host).then(() => { this.layout.invalidateStatic(); this.layout.invalidate(); });
    }
  }

  async readPdfBytes(bytes, path) {
    const name = baseName(path);
    return readPdf(bytes, {
      onProgress: (n, total) => { if (total > 3) this.setBusy('Abrindo ' + name + '... página ' + n + ' de ' + total); },
      askPassword: (retry) => D.passwordDialog(name, retry),
      fonts: this.systemFonts,
    });
  }

  setDocument(model, doc, opts) {
    if (this.hfEdit) this.closeHeaderFooter(true);
    if (this.view) { this.view.destroy(); this.view = null; }
    this.model = model;
    R.model = model;
    this.filePath = opts.path || null;
    this.importPath = null;
    this.importKind = null;
    this.format = opts.format || 'docx';
    this.docName = opts.name || 'Documento1';
    this.dirty = false;
    this.pristine = !opts.path;
    this.recovered = false;
    this.recoverStamp = this.editStamp;
    this.applyDocCss();
    this.layout.setModel(model);
    const state = createState(doc, this, { layout: this.layout });
    if (model.extra && model.extra.md && state.doc !== doc) bindMarkdownSource(model, state.doc);
    this.view = createView(this.hostEl, state, this);
    this.activeView = this.view;
    this.view.dom.classList.toggle('show-markup', this.showMarkup);
    this.layout.setZoom(this.prefs.zoom || 1);
    this.layout.updateCanvas();
    this.scrollEl.scrollTop = 0;
    this.updateTitle();
    this.ribbon.renderStyles && this.ribbon.renderStyles();
    this.nav.render();
    this.countWords();
    this.scheduleUI();
    setTimeout(() => this.focusEditor(), 0);
    if (document.fonts && document.fonts.ready) document.fonts.ready.then(() => this.layout.invalidate());
  }

  applyDocCss() {
    this.styleEl.textContent = this.model.styles.css();
    const s = this.model.section;
    this.pageStyleEl.textContent = '@page{size:' + (s.pgSz.w / 1440).toFixed(4) + 'in ' + (s.pgSz.h / 1440).toFixed(4) + 'in;margin:0}';
  }

  refreshStyles() {
    this.model.styles.cache.clear();
    this.applyDocCss();
    // força nova renderização dos nós
    if (this.view) this.view.setProps({ nodeViews: { ...this.view.props.nodeViews } });
    this.layout.invalidateStatic();
    this.layout.invalidate();
    this.ribbon.renderStyles && this.ribbon.renderStyles();
  }

  fileName() {
    return this.docName;
  }

  formatLabel() {
    return (FORMAT_INFO[this.format] || {}).label || this.format;
  }

  updateTitle() {
    const name = this.docName;
    this.titleText.textContent = name + ' - Redator';
    this.titleState.textContent = this.recovered ? '• Recuperado' : this.dirty ? '• Não salvo' : this.filePath ? '• Salvo neste PC' : this.importPath ? (this.importKind === 'pdf' ? '• Convertido de PDF' : '• Modo de compatibilidade') : '';
    this.host.setDocState({ title: (this.dirty ? '* ' : '') + name + ' - Redator', dirty: this.dirty, filePath: this.filePath || this.importPath || null });
  }

  markDirty() {
    this.pristine = false;
    this.editStamp++;
    this.lastEdit = Date.now();
    if (!this.dirty) { this.dirty = true; this.updateTitle(); }
  }

  markClean() {
    this.dirty = false;
    this.recovered = false;
    this.recoverStamp = this.editStamp;
    this.updateTitle();
    if (this.host.clearRecovery) this.host.clearRecovery();
  }

  // ------------------------------------------------------------ AutoRecuperação

  // Grava uma cópia do documento quando há alterações novas, o usuário parou de
  // digitar e já passou um minuto desde a última cópia.
  saveRecovery(force) {
    if (!this.host.saveRecovery || this.prefs.autoRecover === false || !this.view) return null;
    if (!this.dirty || this.saving || this.recoverJob || this.recoverStamp === this.editStamp) return null;
    const now = Date.now();
    if (!force && (now - this.lastEdit < 2500 || now - this.lastRecover < (this.recoverEvery || 60000))) return null;
    const stamp = this.editStamp;
    const job = async () => {
      const doc = this.view.state.doc;
      const src = this.model.extra && this.model.extra.md;
      const textInfo = (this.model.extra && this.model.extra.text) || null;
      let data = null;
      let kind = 'docx';
      // Markdown aberto de um arquivo guarda o próprio texto, para não reescrever o arquivo inteiro ao salvar
      if (this.format === 'md' && src && src.bound) {
        const path = this.filePath || this.importPath || '';
        const { text, files } = writeMarkdown(doc, this.model, { baseDir: dirName(path), baseName: stripExt(baseName(path) || this.docName) });
        if (!files.length) { data = encodeText(text, textInfo || { encoding: 'utf-8', bom: false }); kind = 'md'; }
      } else if (this.format === 'txt') {
        // texto puro é rápido de gravar mesmo num log enorme
        data = writeText(doc, this.model);
        kind = 'txt';
      } else if (this.format === 'csv') {
        data = writeCsv(doc, this.model);
        if (data) kind = 'csv';
      }
      if (!data) data = await writeDocx(this.model, doc, { author: this.prefs.author || this.userName || '' });
      const meta = { name: this.docName, path: this.filePath || this.importPath || null, format: this.format, kind, text: textInfo };
      const id = await this.host.saveRecovery(data, meta);
      if (id) { this.recoverStamp = stamp; this.recoverId = id; }
      // documento pesado de gravar: a cópia fica mais espaçada para não pesar na digitação
      this.recoverEvery = Math.max(60000, (Date.now() - now) * 40);
    };
    this.lastRecover = now;
    this.recoverJob = job().catch((err) => console.warn('AutoRecuperação falhou', err)).finally(() => { this.recoverJob = null; });
    return this.recoverJob;
  }

  async restoreRecovered(id) {
    const r = await this.host.takeRecovery(id);
    if (!r) return false;
    const { meta, data } = r;
    const res = meta.kind === 'md' ? await readMarkdown(data, meta.path || '', { host: this.host })
      : meta.kind === 'txt' ? await readText(data, meta.path || '')
        : meta.kind === 'csv' ? (await readCsv(data, meta.path || '')) || (await readText(data, meta.path || ''))
          : await readDocx(data);
    if (meta.kind === 'docx' && meta.text && ['md', 'txt'].includes(meta.format)) res.model.extra.text = meta.text;
    this.setDocument(res.model, res.doc, { path: null, format: meta.format || 'docx', name: meta.name || 'Documento1' });
    // salvar pede confirmação do local, que já vem com o arquivo original
    this.importPath = meta.path || null;
    this.recovered = true;
    this.pristine = false;
    this.dirty = true;
    this.updateTitle();
    this.setInfo('Documento recuperado');
    return true;
  }

  async offerRecovery(list) {
    const n = list.length;
    const when = (t) => (t ? ', ' + new Date(t).toLocaleString('pt-BR', { dateStyle: 'short', timeStyle: 'short' }) : '');
    const lines = list.slice(0, 6).map((e) => '• ' + e.name + when(e.time));
    if (n > 6) lines.push('• e mais ' + (n - 6));
    const head = n === 1 ? 'O Redator foi fechado sem salvar este documento:' : 'O Redator foi fechado sem salvar estes ' + n + ' documentos:';
    const r = await D.messageBox('Recuperação de Documentos', head + '\n' + lines.join('\n') + '\nDeseja recuperar a última versão guardada?', [
      { label: 'Recuperar', primary: true, value: 'ok' },
      { label: 'Descartar', value: 'discard' },
      { label: 'Agora não', value: 'later' },
    ], { cancelValue: 'later', icon: 'warn' });
    if (r === 'discard') { await this.host.discardRecovery(list.map((e) => e.id)); return; }
    if (r !== 'ok') return;
    let rest = list;
    if (this.isPristine()) {
      try { if (await this.restoreRecovered(list[0].id)) rest = list.slice(1); } catch (err) { this.showError('Não foi possível recuperar "' + list[0].name + '".', err); rest = list.slice(1); }
    }
    for (const e of rest) await this.host.openRecovered(e.id);
  }

  // ------------------------------------------------------------ salvar

  defaultExt() {
    return this.format === 'pdf' ? 'docx' : this.format || 'docx';
  }

  async save() {
    if (!this.filePath) return this.saveAs(this.defaultExt());
    return this.writeTo(this.filePath, this.format);
  }

  async saveAs(ext) {
    const fmt = ext || this.defaultExt();
    // arquivo de código ou configuração continua com a própria extensão
    const origExt = extOf(this.docName);
    const keepExt = fmt === 'txt' && origExt && origExt !== 'txt' && EXT_FORMAT[origExt] === 'txt';
    const name = keepExt ? this.docName : stripExt(this.docName) + '.' + (fmt === 'md' ? 'md' : fmt);
    const filters = keepExt ? [{ name: 'Arquivo .' + origExt, extensions: [origExt] }, ...FORMAT_INFO.txt.filters] : (FORMAT_INFO[fmt] || FORMAT_INFO.docx).filters;
    const dir = this.filePath ? dirName(this.filePath) : this.importPath ? dirName(this.importPath) : null;
    const path = await this.host.saveDialog({ title: 'Salvar como', defaultPath: dir ? dir + '\\' + name : name, filters });
    if (!path) return false;
    if (extOf(path) === 'pdf') { await this.exportPdf(path); return false; }
    const realFmt = EXT_FORMAT[extOf(path)] || fmt;
    return this.writeTo(path, realFmt);
  }

  async confirmLossy(fmt) {
    if (fmt !== 'txt' && fmt !== 'md' && fmt !== 'csv') return true;
    if (this.format === fmt || this.warnedLossy.has(fmt)) return true;
    const text = fmt === 'csv'
      ? 'Ao salvar como CSV, só a primeira tabela do documento é salva, sem formatação.\nDeseja continuar?'
      : fmt === 'txt'
      ? 'Ao salvar como Texto sem Formatação, toda a formatação, imagens e objetos serão perdidos.\nDeseja continuar?'
      : 'Ao salvar como Markdown, títulos, listas, tabelas, links e imagens são mantidos, mas cores, fontes e layout de página são perdidos.\nDeseja continuar?';
    const r = await D.messageBox('Redator', text, [{ label: 'Salvar', primary: true, value: 'ok' }, { label: 'Cancelar', value: 'cancel' }], { cancelValue: 'cancel', icon: 'warn' });
    if (r !== 'ok') return false;
    this.warnedLossy.add(fmt);
    return true;
  }

  async serialize(fmt, path) {
    const doc = this.view.state.doc;
    const author = this.prefs.author || this.userName || '';
    switch (fmt) {
      case 'docx': return { data: await writeDocx(this.model, doc, { author }) };
      case 'md': {
        const { text, files } = writeMarkdown(doc, this.model, { baseDir: dirName(path), baseName: stripExt(baseName(path)) });
        // mantém a codificação e a marca BOM do arquivo original
        const info = this.model.format === 'md' && this.model.extra.text ? this.model.extra.text : { encoding: 'utf-8', bom: false };
        return { data: encodeText(text, info), extras: files };
      }
      case 'html': return { data: writeHtml(doc, this.model, { title: stripExt(baseName(path)) }) };
      case 'txt': return { data: writeText(doc, this.model) };
      case 'csv': {
        const data = writeCsv(doc, this.model);
        if (!data) throw new Error('Para salvar como CSV, o documento precisa ter uma tabela.');
        return { data };
      }
      case 'odt': return { data: await writeOdt(doc, this.model) };
      case 'rtf': return { data: await writeRtf(doc, this.model) };
      default: throw new Error('Formato não suportado: ' + fmt);
    }
  }

  async writeTo(path, fmt) {
    if (!(await this.confirmLossy(fmt))) return false;
    if (this.hfEdit) this.commitHeaderFooter();
    this.setBusy('Salvando...');
    this.saving = true;
    try {
      // a cópia de recuperação usa o mesmo pacote em memória; espera ela terminar
      if (this.recoverJob) await this.recoverJob;
      const { data, extras } = await this.serialize(fmt, path);
      const res = await this.host.writeFile(path, data, extras);
      if (!res || !res.ok) {
        const code = res && res.error;
        const why = code === 'EBUSY' || code === 'EPERM' ? 'O arquivo está aberto em outro programa ou você não tem permissão para gravá-lo.' : code === 'EACCES' ? 'Você não tem permissão para gravar neste local.' : 'Erro: ' + code;
        await D.messageBox('Redator', 'Não foi possível salvar "' + baseName(path) + '".\n' + why, undefined, { icon: 'error' });
        return false;
      }
      this.filePath = path;
      this.importPath = null;
      this.format = fmt;
      this.docName = baseName(path);
      this.model.format = fmt;
      this.markClean();
      this.setInfo('Salvo');
      return true;
    } catch (err) {
      this.showError('Não foi possível salvar o documento.', err);
      return false;
    } finally {
      this.saving = false;
      this.setBusy(null);
    }
  }

  async requestClose() {
    if (!this.dirty) { this.host.closeConfirmed(); return; }
    const r = await D.confirmSave(this.docName);
    if (r === 'save') {
      const ok = await this.save();
      if (ok) this.host.closeConfirmed();
    } else if (r === 'discard') {
      this.host.closeConfirmed();
    }
  }

  closeWindow() {
    this.host.closeWindow();
  }

  // ------------------------------------------------------------ impressão e PDF

  async preparePrint() {
    closeAllPopups();
    if (this.hfEdit) this.closeHeaderFooter();
    const L = this.layout;
    this.printSaved = { zoom: L.zoom, gap: L.gap, padTop: L.padTop, padBottom: L.padBottom, mode: L.mode, scroll: this.scrollEl.scrollTop, marks: this.showMarks };
    if (this.showMarks) this.root.classList.remove('show-marks');
    L.mode = 'print';
    L.gap = 0;
    L.padTop = 0;
    L.padBottom = 0;
    L.zoom = 1;
    L.pageKeys = [];
    document.documentElement.classList.add('printing');
    this.applyDocCss();
    L.updateCanvas();
    L.runNow(10);
    await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
  }

  async restoreAfterPrint() {
    const s = this.printSaved;
    const L = this.layout;
    document.documentElement.classList.remove('printing');
    if (!s) return;
    L.gap = s.gap; L.padTop = s.padTop; L.padBottom = s.padBottom; L.mode = s.mode; L.zoom = s.zoom;
    L.pageKeys = [];
    if (s.marks) this.root.classList.add('show-marks');
    L.updateCanvas();
    L.runNow(10);
    this.scrollEl.scrollTop = s.scroll;
    this.printSaved = null;
  }

  async print() {
    await this.preparePrint();
    try { await this.host.print(); } finally { await this.restoreAfterPrint(); }
  }

  async exportPdf(target) {
    const name = stripExt(this.docName) + '.pdf';
    const dir = this.filePath ? dirName(this.filePath) : this.importPath ? dirName(this.importPath) : null;
    const path = typeof target === 'string' ? target : await this.host.saveDialog({ title: 'Publicar como PDF', defaultPath: dir ? dir + '\\' + name : name, filters: FORMAT_INFO.pdf.filters });
    if (!path) return;
    await this.preparePrint();
    let res;
    try {
      const s = this.model.section;
      res = await this.host.exportPdf(path, { pageSize: { width: s.pgSz.w / 1440, height: s.pgSz.h / 1440 } });
    } finally {
      await this.restoreAfterPrint();
    }
    if (res && res.ok === false) this.showError('Não foi possível criar o PDF.', new Error(res.error));
    else {
      this.setInfo('PDF salvo');
      const r = await D.messageBox('Redator', 'O PDF foi salvo em:\n' + path, [{ label: 'Abrir PDF', value: 'open', primary: true }, { label: 'Fechar', value: 'close' }], { cancelValue: 'close' });
      if (r === 'open') this.host.openPath(path);
    }
  }

  renderPreview(container) {
    const first = this.layout.pageEls[0];
    if (!first) return;
    const g = this.layout.geom;
    const scale = Math.min(0.5, 420 / g.pageW);
    const clone = this.zoomEl.cloneNode(true);
    clone.style.transform = 'scale(' + scale + ')';
    clone.style.left = '0';
    clone.style.top = '0';
    const frame = h('div', { class: 'bs-preview-frame', style: { width: g.pageW * scale + 'px', height: g.pageH * scale + 'px' } }, clone);
    container.appendChild(frame);
  }

  // ------------------------------------------------------------ edição

  focusEditor() {
    const v = this.activeView || this.view;
    if (v && !this.backstage.isOpen && !D.isDialogOpen()) v.focus();
  }

  onEditorFocus(view) {
    if (this.hfEdit && view === this.view) return;
    this.activeView = view;
  }

  edit(action) {
    this.focusEditor();
    if (this.host.edit) this.host.edit(action);
    else document.execCommand(action);
  }

  undo() { const v = this.activeView; if (v) { undo(v.state, v.dispatch); v.focus(); } }
  redo() { const v = this.activeView; if (v) { redo(v.state, v.dispatch); v.focus(); } }

  cmd(name, ...args) {
    const view = this.activeView || this.view;
    if (!view || (this.readOnly && !['selectAll'].includes(name))) return;
    const run = (command) => command(view.state, view.dispatch, view);
    switch (name) {
      case 'toggle': run(C.toggleRunProp(args[0])); break;
      case 'setFont': if (args[0]) { run(C.setFont(args[0])); this.addRecentFont(args[0]); } break;
      case 'setFontSize': run(C.setFontSize(args[0])); break;
      case 'setColor': run(C.setColor(args[0])); break;
      case 'highlight': run(C.setHighlight(args[0])); break;
      case 'setUnderline': run(C.setUnderline(args[0])); break;
      case 'setMark': run(C.setRunMark(args[0], args[1])); break;
      case 'vertAlign': run(C.setVertAlign(args[0])); break;
      case 'growFont': run(C.growFont(args[0], args[1])); break;
      case 'clearAll': run(C.clearAllFormatting); break;
      case 'clearChar': run(C.clearCharFormatting); break;
      case 'changeCase': run(C.changeCase(args[0])); break;
      case 'align': run(C.setAlign(args[0])); break;
      case 'lineSpacing': run(C.setLineSpacing(args[0])); break;
      case 'spaceBefore': run(C.toggleSpaceBefore); break;
      case 'spaceAfter': run(C.toggleSpaceAfter); break;
      case 'spacing': run(C.setSpacing(args[0], args[1])); break;
      case 'indent': run(C.changeIndent(args[0])); break;
      case 'indentValues': run(C.setIndent(args[0])); break;
      case 'paraProps': run(C.setParaProps(args[0])); break;
      case 'paraShading': run(C.setParaShading(args[0])); break;
      case 'paraBorders': run(C.setParaBorders(args[0])); break;
      case 'horizontalLine': this.insertHorizontalLine(); break;
      case 'applyStyle': run(C.applyStyle(args[0])); break;
      case 'list': run(C.toggleList(args[0], args[1])); break;
      case 'removeList': run(C.updateParagraphs((a) => { const st = this.model.styles.resolvePara(a.styleId || this.model.styles.defaultPara).pPr; return { ...a, numId: st.numId ? '0' : null, ilvl: null }; })); break;
      case 'restartNumbering': run(C.restartNumbering); break;
      case 'insertTable': run(C.insertTable(args[0], args[1])); break;
      case 'insertLink': run(C.insertLink(args[0], args[1])); break;
      case 'removeLink': run(C.removeLink); break;
      case 'insertText': run(C.insertText(args[0])); break;
      case 'pageBreak': run(C.pageBreak); break;
      case 'lineBreak': run(C.lineBreak); break;
      case 'selectAll': run(selectAll); break;
      case 'selectParagraph': run(C.selectParagraph); break;
      case 'table': this.tableOp(args[0]); break;
      case 'tableAttrs': run(C.setTableAttrs(args[0])); break;
      case 'cellAttrs': run(C.setCellAttrs(args[0])); break;
      case 'tableBorders': run(C.setCellBorders(args[0])); break;
      default: console.warn('Comando desconhecido', name);
    }
    this.focusEditor();
  }

  tableOp(op) {
    const view = this.activeView;
    const ops = { addRowBefore, addRowAfter, addColumnBefore, addColumnAfter, deleteRow, deleteColumn, deleteTable, mergeCells, splitCell, toggleHeaderRow: C.toggleHeaderRow };
    const fn = ops[op];
    if (!fn) return;
    const before = C.findTableAt(view.state);
    const widthBefore = before ? columnWidths(before.node).reduce((s, w) => s + w, 0) : 0;
    fn(view.state, view.dispatch);
    if ((op === 'addColumnBefore' || op === 'addColumnAfter' || op === 'deleteColumn') && widthBefore) this.normalizeTableWidth(widthBefore);
  }

  // mantém a largura total da tabela ao inserir ou excluir colunas
  normalizeTableWidth(total) {
    const view = this.activeView;
    const t = C.findTableAt(view.state);
    if (!t) return;
    const widths = columnWidths(t.node);
    const sum = widths.reduce((s, w) => s + w, 0);
    if (!sum || Math.abs(sum - total) < 1) return;
    const scaled = widths.map((w) => (w * total) / sum);
    const tr = view.state.tr;
    const map = TableMap.get(t.node);
    t.node.descendants((cell, pos) => {
      if (cell.type !== schema.nodes.table_cell) return true;
      const rect = map.findCell(pos);
      const cw = scaled.slice(rect.left, rect.right).map((w) => Math.round(w * 100) / 100);
      tr.setNodeMarkup(t.pos + 1 + pos, null, { ...cell.attrs, colwidth: cw });
      return false;
    });
    tr.setNodeMarkup(t.pos, null, { ...t.node.attrs, grid: scaled.map((w) => Math.round(w * 15)) });
    view.dispatch(tr);
  }

  insertHorizontalLine() {
    const view = this.activeView;
    const p = C.currentParagraph(view.state);
    if (!p) return;
    const hr = { bottom: { val: 'single', sz: '6', space: '1', color: 'auto' } };
    const tr = view.state.tr;
    if (p.node.content.size === 0) tr.setNodeMarkup(p.pos, null, { ...p.node.attrs, pBdr: hr });
    else {
      const after = p.pos + p.node.nodeSize;
      tr.insert(after, schema.nodes.paragraph.create({ pBdr: hr }));
      tr.insert(after + 2, schema.nodes.paragraph.create());
      tr.setSelection(TextSelection.create(tr.doc, after + 3));
    }
    view.dispatch(tr.scrollIntoView());
  }

  onTransaction(view, tr) {
    if (tr.docChanged) {
      if (view === this.view) {
        this.markDirty();
        this.countWordsSoon();
        this.navSoon();
      } else if (this.hfEdit && view === this.hfEdit.view) {
        this.hfEdit.entry.doc = view.state.doc;
        this.hfEdit.entry.dirty = true;
        this.markDirty();
        this.layout.invalidateStatic();
        this.layout.invalidate();
      }
    }
    this.scheduleUI();
  }

  layoutSoon() { this.layout.invalidate(); }

  get zoom() { return this.layout.zoom; }

  // ------------------------------------------------------------ estado da interface

  scheduleUI() {
    if (this.uiFrame) return;
    this.uiFrame = requestAnimationFrame(() => { this.uiFrame = null; this.updateUI(); });
  }

  currentFontName(r) {
    const run = r || this.effectiveRun();
    const f = run.rFonts ? fontName(run.rFonts.a) : null;
    return f || 'Times New Roman';
  }

  effectiveRun() {
    const v = this.activeView || this.view;
    return v ? C.effectiveRunAt(v.state) : {};
  }

  effectivePara() {
    const v = this.activeView || this.view;
    return v ? C.effectivePara(v.state) : {};
  }

  updateUI() {
    const view = this.activeView || this.view;
    if (!view || !this.model) return;
    const st = view.state;
    const r = C.effectiveRunAt(st);
    const eff = C.effectivePara(st);
    const para = C.currentParagraph(st);
    const inTable = isInTable(st);
    const sel = st.selection;
    const drawing = sel instanceof NodeSelection && sel.node.type === schema.nodes.drawing ? sel.node.attrs : null;
    const jc = eff.jc === 'start' ? 'left' : eff.jc === 'end' ? 'right' : eff.jc === 'distribute' ? 'both' : eff.jc || 'left';
    const sp = eff.spacing || {};
    const ind = eff.ind || {};
    const s = this.model.section;
    const paper = PAPER_SIZES.find((p) => (Math.abs(p.w - s.pgSz.w) < 20 && Math.abs(p.h - s.pgSz.h) < 20) || (Math.abs(p.h - s.pgSz.w) < 20 && Math.abs(p.w - s.pgSz.h) < 20));
    const preset = MARGIN_PRESETS.find((m) => Math.abs(m.top - s.pgMar.top) < 10 && Math.abs(m.bottom - s.pgMar.bottom) < 10 && Math.abs(m.left - s.pgMar.left) < 10 && Math.abs(m.right - s.pgMar.right) < 10);
    let look = null;
    if (inTable) {
      const t = C.findTableAt(st);
      const lk = (t && t.node.attrs.tblLook) || {};
      const on = (k) => lk[k] === '1' || lk[k] === 'true';
      if (lk.firstRow != null) look = { firstRow: on('firstRow'), lastRow: on('lastRow'), firstColumn: on('firstColumn'), lastColumn: on('lastColumn'), noHBand: on('noHBand'), noVBand: on('noVBand') };
      else { const v = parseInt(lk.val || '04A0', 16); look = { firstRow: !!(v & 0x20), lastRow: !!(v & 0x40), firstColumn: !!(v & 0x80), lastColumn: !!(v & 0x100), noHBand: !!(v & 0x200), noVBand: !!(v & 0x400) }; }
    }
    let rowHeader = false;
    if (inTable) {
      const $f = sel.$from;
      for (let d = $f.depth; d > 0; d--) if ($f.node(d).type === schema.nodes.table_row) { rowHeader = !!$f.node(d).attrs.header; break; }
    }
    const ui = {
      bold: C.isRunPropOn(st, 'b'), italic: C.isRunPropOn(st, 'i'), underline: C.isRunPropOn(st, 'u'), strike: C.isRunPropOn(st, 'strike'),
      sub: !!(r.vertAlign && r.vertAlign.val === 'subscript'), sup: !!(r.vertAlign && r.vertAlign.val === 'superscript'),
      font: this.currentFontName(r), size: r.sz ? r.sz.val / 2 : 10,
      align: jc === 'both' ? 'both' : jc, list: C.currentListKind(st),
      styleId: para ? para.node.attrs.styleId || this.model.styles.defaultPara : null,
      lineSpacing: sp.line && (!sp.lineRule || sp.lineRule === 'auto') ? Math.round((Number(sp.line) / 240) * 100) / 100 : null,
      spaceBefore: Number(sp.before) > 0, spaceAfter: Number(sp.after) > 0,
      para: { left: Number(ind.left) || 0, right: Number(ind.right) || 0, firstLine: Number(ind.firstLine) || 0, hanging: Number(ind.hanging) || 0, before: Number(sp.before) || 0, after: Number(sp.after) || 0, tabs: eff.tabs || [] },
      inTable, look, headerRow: rowHeader, drawing,
      headerFooter: !!this.hfEdit, titlePg: s.titlePg, evenOdd: !!this.model.settings.evenAndOdd,
      showMarks: this.showMarks, viewMode: this.viewMode, painter: !!this.painter, spellcheck: this.prefs.spellcheck !== false,
      ruler: this.prefs.ruler !== false, navPane: this.nav.visible,
      landscape: s.pgSz.w > s.pgSz.h, paper: paper ? paper.id : null, marginPreset: preset ? preset.id : null,
      canPageBreak: !inTable && !this.hfEdit,
    };
    this.uiState = ui;
    this.ribbon.update(ui);
    this.ruler.update(ui.para);
    this.undoBtn.disabled = undoDepth(st) === 0;
    this.redoBtn.disabled = redoDepth(st) === 0;
    const page = this.layout.mode === 'print' ? (sel.empty ? this.layout.pageOfPos(sel.head) : this.layout.pageAtScroll()) : 0;
    let selWords = null;
    if (!sel.empty && view === this.view) selWords = this.countIn(st.doc.textBetween(sel.from, sel.to, ' ', ' '));
    this.status.update({
      page, pages: this.layout.mode === 'print' ? this.layout.pageCount : 1, words: this.wordCount, selWords, zoom: this.layout.zoom, viewMode: this.viewMode,
      language: this.languageName(), info: this.infoText || '',
    });
  }

  countIn(text) {
    let n = 0;
    for (const w of text.split(/\s+/)) if (w && /[\p{L}\p{N}]/u.test(w)) n++;
    return n;
  }

  countWords() {
    if (!this.view) return;
    this.wordCount = this.countIn(this.view.state.doc.textBetween(0, this.view.state.doc.content.size, ' ', ' '));
    this.scheduleUI();
  }

  countWordsSoon() {
    clearTimeout(this.wcTimer);
    this.wcTimer = setTimeout(() => this.countWords(), 400);
  }

  navSoon() {
    if (!this.nav.visible) return;
    clearTimeout(this.navTimer);
    this.navTimer = setTimeout(() => this.nav.render(), 600);
  }

  stats(full) {
    const doc = this.view.state.doc;
    const text = doc.textBetween(0, doc.content.size, '\n', ' ');
    let paragraphs = 0;
    doc.descendants((n) => { if (n.type === schema.nodes.paragraph) { if (n.textContent.trim()) paragraphs++; return false; } return true; });
    let lines = 0;
    if (full) {
      for (const p of this.view.dom.querySelectorAll('p')) {
        const rects = p.getClientRects();
        const lh = parseFloat(getComputedStyle(p).lineHeight) || 16;
        if (p.textContent.trim()) lines += Math.max(1, Math.round((p.getBoundingClientRect().height / this.layout.zoom) / lh));
        void rects;
      }
    }
    return {
      pages: this.layout.pageCount, words: this.countIn(text), chars: text.replace(/\n/g, '').length,
      charsNoSpaces: text.replace(/\s/g, '').length, paragraphs, lines,
    };
  }

  setInfo(text) {
    this.infoText = text;
    clearTimeout(this.infoTimer);
    if (text) this.infoTimer = setTimeout(() => { this.infoText = ''; this.scheduleUI(); }, 4000);
    this.scheduleUI();
  }

  setBusy(text) {
    this.root.classList.toggle('busy', !!text);
    if (text) this.setInfo(text);
  }

  showError(msg, err) {
    console.error(err);
    D.messageBox('Redator', msg + (err && err.message ? '\n\n' + err.message : ''), undefined, { icon: 'error' });
  }

  // ------------------------------------------------------------ fontes e símbolos

  fontList() {
    const set = new Set([...(this.systemFonts.length ? this.systemFonts : FALLBACK_FONTS)]);
    if (this.model) {
      for (const f of this.model.fontTable.keys()) set.add(f);
      set.add(this.model.theme.major);
      set.add(this.model.theme.minor);
    }
    return [...set].filter(Boolean).sort((a, b) => a.localeCompare(b, 'pt-BR'));
  }

  recentFonts() {
    return (this.prefs.recentFonts || []).slice(0, 6);
  }

  addRecentFont(f) {
    const list = [f, ...(this.prefs.recentFonts || []).filter((x) => x !== f)].slice(0, 8);
    this.prefs.recentFonts = list;
    this.host.savePrefs({ recentFonts: list });
  }

  recentSymbols() {
    return this.symbols.slice(0, 16);
  }

  insertSymbol(ch) {
    this.cmd('insertText', ch);
    this.symbols = [ch, ...this.symbols.filter((x) => x !== ch)].slice(0, 20);
    this.host.savePrefs({ recentSymbols: this.symbols });
  }

  selectedText() {
    const v = this.activeView || this.view;
    if (!v) return '';
    const { from, to } = v.state.selection;
    return v.state.doc.textBetween(from, to, '\n', ' ');
  }

  // ------------------------------------------------------------ diálogos

  dialog(name, arg) {
    closeAllPopups();
    switch (name) {
      case 'font': return D.fontDialog(this);
      case 'paragraph': return D.paragraphDialog(this, arg || 0);
      case 'pageSetup': return D.pageSetupDialog(this, arg || 0);
      case 'find': return D.findReplaceDialog(this, arg || 0);
      case 'insertTable': return D.insertTableDialog(this);
      case 'link': return D.linkDialog(this);
      case 'wordCount': return D.wordCountDialog(this);
      case 'symbol': return D.symbolDialog(this);
      case 'dateTime': return D.dateTimeDialog(this);
      case 'zoom': return D.zoomDialog(this);
      case 'options': return D.optionsDialog(this);
      case 'tableProperties': return D.tablePropertiesDialog(this);
      case 'imageSize': return D.imageSizeDialog(this);
      default: return null;
    }
  }

  async savePrefs(p) {
    const oldShortcuts = this.prefs.shortcuts;
    this.prefs = { ...this.prefs, ...p };
    await this.host.savePrefs(p);
    if (this.view) this.view.dom.spellcheck = this.prefs.spellcheck !== false;
    if (p.shortcuts && p.shortcuts !== oldShortcuts) this.setInfo('Atalhos atualizados');
    this.scheduleUI();
  }

  stylesPane() {
    const anchor = this.ribbon.controls.get('styles');
    const items = [];
    const st = this.model.styles;
    items.push({ header: 'Estilos de Parágrafo' });
    for (const s of st.allStyles('paragraph').filter((x) => !x.hidden)) items.push({ label: st.displayName(s), checked: this.uiState.styleId === s.id, run: () => this.cmd('applyStyle', s.id) });
    items.push({ header: 'Estilos de Caractere' });
    for (const s of st.allStyles('character').filter((x) => !x.hidden && !x.semiHidden)) items.push({ label: st.displayName(s), run: () => this.cmd('applyStyle', s.id) });
    openMenu(anchor ? anchor.el : this.ribbon.el, items, { className: 'styles-menu' });
  }

  // ------------------------------------------------------------ busca

  search(q, opts) {
    const v = this.view;
    if (!v) return 0;
    this.lastSearch = q;
    this.searchOpts = opts || this.searchOpts;
    this.searchCount = q ? setSearch(v, q, this.searchOpts) : (clearSearch(v), 0);
    if (this.nav.visible) this.nav.render();
    return this.searchCount;
  }

  findNext(dir, fromStart) {
    const v = this.view;
    if (!v) return;
    if (fromStart) gotoMatch(v, 0);
    else findNext(v, dir);
    this.scrollSelectionIntoView();
  }

  gotoMatch(i) {
    gotoMatch(this.view, i);
    this.scrollSelectionIntoView();
    this.view.focus();
  }

  searchResults() {
    return this.view ? findKey.getState(this.view.state).matches : [];
  }

  replaceCurrent(text) { replaceCurrent(this.view, text); }
  replaceAll(text) { return replaceAll(this.view, text); }

  scrollSelectionIntoView() {
    const v = this.view;
    try {
      const c = v.coordsAtPos(v.state.selection.from);
      const r = this.scrollEl.getBoundingClientRect();
      if (c.top < r.top + 40 || c.bottom > r.bottom - 40) this.scrollEl.scrollTop += c.top - r.top - r.height / 3;
    } catch { /* ignora */ }
  }

  // ------------------------------------------------------------ navegação

  headingList() {
    const out = [];
    if (!this.view) return out;
    const st = this.model.styles;
    this.view.state.doc.descendants((n, pos) => {
      if (n.type === schema.nodes.paragraph) {
        let level = st.headingLevel(n.attrs.styleId || st.defaultPara);
        if (n.attrs.outlineLvl != null && n.attrs.outlineLvl < 9) level = n.attrs.outlineLvl + 1;
        if (level >= 1 && level <= 9 && n.textContent.trim()) {
          let bookmark = null;
          n.forEach((c) => {
            if (!bookmark && c.type === schema.nodes.opaque_inline && c.attrs.kind === 'bookmarkStart') {
              const m = /w:name="([^"]+)"/.exec(c.attrs.xml || '');
              if (m && m[1] !== '_GoBack') bookmark = m[1];
            }
          });
          out.push({ level, text: n.textContent.trim().slice(0, 200), pos, bookmark: bookmark || '_Ref' + (pos * 7919 % 1000000000) });
        }
        return false;
      }
      return n.type.name !== 'table' || true;
    });
    return out;
  }

  currentHeadingIndex() {
    const list = this.headingList();
    const from = this.view.state.selection.from;
    let idx = -1;
    list.forEach((hd, i) => { if (hd.pos <= from) idx = i; });
    return idx;
  }

  gotoPos(pos) {
    const v = this.view;
    const sel = TextSelection.near(v.state.doc.resolve(Math.min(pos + 1, v.state.doc.content.size)));
    v.dispatch(v.state.tr.setSelection(sel));
    v.focus();
    try {
      const c = v.coordsAtPos(sel.from);
      const r = this.scrollEl.getBoundingClientRect();
      this.scrollEl.scrollTop += c.top - r.top - 60;
    } catch { /* ignora */ }
  }

  gotoPage(k) {
    const L = this.layout;
    if (L.mode !== 'print') return;
    const page = Math.max(0, Math.min(L.pageCount - 1, k));
    const y = (L.padTop + page * L.geom.pitch) * L.zoom;
    this.scrollEl.scrollTop = y - 10;
    const zr = this.zoomEl.getBoundingClientRect();
    const pos = this.view.posAtCoords({ left: zr.left + (L.geom.mLeft + 5) * L.zoom, top: zr.top + (page * L.geom.pitch + L.cTop(page) + 5) * L.zoom });
    if (pos) this.view.dispatch(this.view.state.tr.setSelection(TextSelection.near(this.view.state.doc.resolve(pos.pos))));
    this.view.focus();
  }

  showNavPane(tab, query) {
    this.nav.el.hidden = false;
    this.prefs.navPane = true;
    this.host.savePrefs({ navPane: true });
    if (tab === 'search') {
      if (query != null) this.nav.input.value = query;
      else if (this.selectedText() && !this.selectedText().includes('\n')) this.nav.input.value = this.selectedText().slice(0, 100);
      this.nav.input.focus();
      this.nav.input.select();
      if (this.nav.input.value) this.nav.search();
    }
    this.nav.render();
    this.layout.updateCanvas();
    this.ruler.update();
    this.scheduleUI();
  }

  hideNavPane() {
    this.nav.el.hidden = true;
    this.prefs.navPane = false;
    this.host.savePrefs({ navPane: false });
    if (this.view) clearSearch(this.view);
    this.layout.updateCanvas();
    this.ruler.update();
    this.scheduleUI();
    this.focusEditor();
  }

  // ------------------------------------------------------------ exibição

  setZoom(z) {
    this.layout.setZoom(z);
    this.prefs.zoom = this.layout.zoom;
    clearTimeout(this.zoomSaveTimer);
    this.zoomSaveTimer = setTimeout(() => this.host.savePrefs({ zoom: this.layout.zoom }), 800);
    if (this.viewMode !== 'print') this.hostEl.style.zoom = String(this.layout.zoom);
    this.ruler.update();
    this.scheduleUI();
  }

  stepZoom(dir, step = 0.1) {
    const z = this.layout.zoom;
    const next = dir > 0 ? Math.floor((z + step + 1e-6) * 10) / 10 : Math.ceil((z - step - 1e-6) * 10) / 10;
    this.setZoom(Math.max(0.1, Math.min(5, next)));
  }

  zoomTo(kind) {
    const g = this.layout.geom;
    const w = this.scrollEl.clientWidth - 40;
    const hgt = this.scrollEl.clientHeight - 30;
    if (kind === 'width') this.setZoom(Math.round((w / g.pageW) * 100) / 100);
    else if (kind === 'text') this.setZoom(Math.round((w / (g.contentW + 20)) * 100) / 100);
    else if (kind === 'page') this.setZoom(Math.round(Math.min(w / g.pageW, hgt / g.pageH) * 100) / 100);
  }

  setViewMode(mode) {
    if (this.hfEdit) this.closeHeaderFooter();
    this.viewMode = mode;
    this.root.classList.toggle('mode-web', mode === 'web' || mode === 'read');
    this.root.classList.toggle('mode-read', mode === 'read');
    this.layout.mode = mode === 'print' ? 'print' : 'web';
    this.hostEl.style.zoom = mode === 'print' ? '' : String(this.layout.zoom);
    if (mode === 'read') { this.readOnlyBefore = this.readOnly; this.readOnly = true; }
    else if (this.readOnlyBefore != null) { this.readOnly = this.readOnlyBefore; this.readOnlyBefore = null; }
    if (this.view) this.view.setProps({});
    this.layout.pageKeys = [];
    this.layout.invalidate();
    this.layout.updateCanvas();
    this.scheduleUI();
    this.focusEditor();
  }

  toggleMarks() {
    this.showMarks = !this.showMarks;
    this.root.classList.toggle('show-marks', this.showMarks);
    this.layout.invalidate();
    this.scheduleUI();
  }

  toggleFocus() {
    this.root.classList.toggle('focus-mode');
    this.onRibbonResize();
  }

  setRuler(on, silent) {
    this.prefs.ruler = on;
    this.ruler.el.hidden = !on;
    this.root.classList.toggle('no-ruler', !on);
    if (!silent) this.host.savePrefs({ ruler: on });
    requestAnimationFrame(() => { this.layout.updateCanvas(); this.ruler.update(); });
    this.scheduleUI();
  }

  onRibbonResize() {
    requestAnimationFrame(() => { this.layout.updateCanvas && this.layout.geom && this.layout.updateCanvas(); this.ruler.update(); });
  }

  showGuide() {
    return (x) => {
      if (x == null) { this.guideEl.hidden = true; return; }
      const r = this.docArea.getBoundingClientRect();
      this.guideEl.hidden = false;
      this.guideEl.style.left = x - r.left + 'px';
    };
  }

  openBackstage() { closeAllPopups(); this.backstage.open('home'); }
  setOverlayColor() {}

  languageName() {
    const code = this.lang || 'pt-BR';
    const f = LANGS.find((l) => l[0] === code);
    return f ? f[1] : code;
  }

  languageMenu() {
    return LANGS.map(([code, name]) => ({ label: name, checked: (this.lang || 'pt-BR') === code, run: () => this.setLanguage(code) }));
  }

  openLanguageMenu(anchor) {
    openMenu(anchor, this.languageMenu());
  }

  async setLanguage(code) {
    const ok = await this.host.setSpellLanguages([code]);
    if (ok === false) { D.messageBox('Redator', 'O dicionário deste idioma não está disponível no sistema.'); return; }
    this.lang = code;
    if (this.view) this.view.dom.setAttribute('lang', code);
    this.scheduleUI();
  }

  toggleSpellcheck() {
    this.savePrefs({ spellcheck: this.prefs.spellcheck === false });
  }

  setShowMarkup(on) {
    this.showMarkup = on;
    if (this.view) this.view.dom.classList.toggle('show-markup', on);
    this.layout.invalidate();
  }

  newWindow() {
    this.host.newWindow(null);
  }

  // ------------------------------------------------------------ teclado e mouse

  onKeyDown(e) {
    if (D.isDialogOpen()) return;
    const mod = e.ctrlKey || e.metaKey;
    const code = e.code;
    const inInput = !!(e.target && e.target.closest && e.target.closest('input, textarea, select'));
    const stop = () => { e.preventDefault(); e.stopPropagation(); };
    if (this.backstage.isOpen) return;
    if (e.key === 'F12') { stop(); this.saveAs(); return; }
    if (e.key === 'F7') { stop(); this.toggleSpellcheck(); return; }
    if (e.key === 'Escape') {
      if (this.hfEdit) { stop(); this.closeHeaderFooter(); return; }
      if (this.painter) { stop(); this.cancelPainter(); return; }
      if (this.viewMode === 'read') { stop(); this.setViewMode('print'); return; }
      if (this.root.classList.contains('focus-mode')) { stop(); this.toggleFocus(); return; }
    }
    if (e.shiftKey && e.key === 'F3' && !mod) { stop(); this.cycleCase(); return; }
    if (!mod) return;
    const pt = this.prefs.shortcuts === 'pt';
    const key = code.replace(/^Key/, '').toLowerCase();
    // atalhos que valem até dentro de campos
    const global = pt
      ? { b: () => this.save(), a: () => this.openDialog(), o: () => this.newDocument(), p: () => this.print(), w: () => this.requestCloseWindow() }
      : { s: () => (e.shiftKey ? this.saveAs() : this.save()), o: () => this.openDialog(), n: () => this.newDocument(), p: () => this.print(), w: () => this.requestCloseWindow() };
    if (!e.altKey && !(pt ? false : e.shiftKey && key !== 's') && global[key]) { stop(); global[key](); return; }
    if (inInput) return;
    if (e.altKey && /^Digit[1-3]$/.test(code)) { stop(); const id = R.headingStyleId(Number(code.slice(5))); if (id) this.cmd('applyStyle', id); return; }
    if (e.shiftKey && code === 'Digit8') { stop(); this.toggleMarks(); return; }
    if (e.shiftKey && key === 'c') { stop(); this.formatPainter(false, true); return; }
    if (e.shiftKey && key === 'v') { stop(); if (this.copiedFormat) this.activeView.dispatch && C.applyFormat(this.copiedFormat)(this.activeView.state, this.activeView.dispatch); return; }
    if (e.shiftKey) return;
    const local = pt
      ? { n: () => this.cmd('toggle', 'b'), s: () => this.cmd('toggle', 'u'), t: () => this.cmd('selectAll'), l: () => this.showNavPane('search'), u: () => this.dialog('find', 1), q: () => this.cmd('align', 'left'), g: () => this.cmd('align', 'right'), r: () => this.redo(), y: () => this.dialog('find', 2), k: () => this.dialog('link'), d: () => this.dialog('font'), f: () => this.showNavPane('search'), h: () => this.dialog('find', 1) }
      : { f: () => this.showNavPane('search'), h: () => this.dialog('find', 1), g: () => this.dialog('find', 2), k: () => this.dialog('link'), d: () => this.dialog('font') };
    if (!e.altKey && local[key]) { stop(); local[key](); }
  }

  requestCloseWindow() {
    this.host.closeWindow();
  }

  cycleCase() {
    const v = this.activeView;
    const text = this.selectedText();
    if (!text) return;
    const mode = text === text.toLocaleUpperCase('pt-BR') ? 'lower' : text === text.toLocaleLowerCase('pt-BR') ? 'title' : 'upper';
    C.changeCase(mode)(v.state, v.dispatch);
  }

  onDocMouseDown(e) {
    if (this.hfEdit && !e.target.closest('.page-header, .page-footer')) {
      // clique fora do cabeçalho fecha a edição, como no Word
      if (e.detail >= 2 && e.target.closest('.editor-host, .page')) this.closeHeaderFooter();
    }
  }

  onDocMouseUp() {
    if (this.painter) {
      const v = this.activeView;
      setTimeout(() => {
        if (!v.state.selection.empty || this.painter.clickApply) {
          C.applyFormat(this.painter.fmt)(v.state, v.dispatch);
          if (!this.painter.sticky) this.cancelPainter();
        }
      }, 0);
    }
  }

  onDocDblClick(e) {
    if (this.layout.mode !== 'print' || this.hfEdit) return;
    const L = this.layout;
    const zr = this.zoomEl.getBoundingClientRect();
    const y = (e.clientY - zr.top) / L.zoom;
    const k = Math.floor(y / L.geom.pitch);
    if (k < 0 || k >= L.pageCount) return;
    const inPage = y - k * L.geom.pitch;
    if (inPage < L.cTop(k) - 2) { e.preventDefault(); e.stopPropagation(); this.editHeaderFooter('header', k); }
    else if (inPage > L.geom.pageH - L.cBottom(k) + 2 && inPage < L.geom.pageH) { e.preventDefault(); e.stopPropagation(); this.editHeaderFooter('footer', k); }
  }

  // palavra sob o ponteiro e sugestões do corretor do sistema
  wordAt(v, x, y) {
    const hit = v.posAtCoords({ left: x, top: y });
    if (!hit) return null;
    const $p = v.state.doc.resolve(hit.pos);
    if (!$p.parent.isTextblock) return null;
    const text = $p.parent.textBetween(0, $p.parent.content.size, null, '￼');
    let a = $p.parentOffset;
    let b = a;
    const isW = (ch) => /[\p{L}\p{M}'’-]/u.test(ch);
    while (a > 0 && isW(text[a - 1])) a--;
    while (b < text.length && isW(text[b])) b++;
    const word = text.slice(a, b).replace(/^['’-]+|['’-]+$/g, '');
    if (!word) return null;
    const start = $p.start() + a + text.slice(a, b).indexOf(word);
    return { word, from: start, to: start + word.length };
  }

  async onContextMenu(info) {
    const v = this.activeView || this.view;
    if (!v) return;
    if (this.hfEdit && v === this.view) return;
    this.spellTarget = null;
    if (!info.misspelledWord && this.host.checkWord && this.prefs.spellcheck !== false) {
      const w = this.wordAt(v, info.x, info.y);
      if (w && w.word.length > 1 && !/\d/.test(w.word)) {
        try {
          const res = await this.host.checkWord(w.word);
          if (res && res.misspelled) {
            info = { ...info, misspelledWord: w.word, suggestions: res.suggestions || [] };
            this.spellTarget = { view: v, ...w };
          }
        } catch { /* sem sugestões */ }
      }
    }
    const st = v.state;
    const link = C.linkRangeAt(st, st.selection.from);
    openContextMenu(this, {
      ...info,
      hasSelection: !st.selection.empty,
      selectionText: this.selectedText(),
      link: link ? link.mark.attrs : null,
      drawing: st.selection instanceof NodeSelection && st.selection.node.type === schema.nodes.drawing,
      inTable: isInTable(st),
      inList: !!C.currentListKind(st),
    });
  }

  replaceMisspelling(word) {
    const t = this.spellTarget;
    if (t && t.view.state.doc.textBetween(t.from, t.to) === t.word) {
      t.view.dispatch(t.view.state.tr.insertText(word, t.from, t.to));
      t.view.focus();
    } else this.host.replaceMisspelling(word);
    this.spellTarget = null;
  }

  formatPainter(sticky, copyOnly) {
    const v = this.activeView;
    const fmt = C.captureFormat(v.state);
    if (copyOnly) { this.copiedFormat = fmt; this.setInfo('Formatação copiada'); return; }
    if (this.painter) { this.cancelPainter(); return; }
    this.painter = { fmt, sticky };
    this.root.classList.add('painter');
    this.scheduleUI();
  }

  cancelPainter() {
    this.painter = null;
    this.root.classList.remove('painter');
    this.scheduleUI();
  }

  sortParagraphs() {
    const v = this.activeView;
    const paras = C.paragraphsInSelection(v.state);
    if (paras.length < 2) return;
    const sorted = [...paras].sort((a, b) => a.node.textContent.localeCompare(b.node.textContent, 'pt-BR', { sensitivity: 'base', numeric: true }));
    const from = paras[0].pos;
    const last = paras[paras.length - 1];
    const to = last.pos + last.node.nodeSize;
    const $a = v.state.doc.resolve(from);
    const $b = v.state.doc.resolve(to);
    if ($a.parent !== $b.parent) return;
    v.dispatch(v.state.tr.replaceWith(from, to, sorted.map((p) => p.node)).scrollIntoView());
  }

  async followLink(attrs) {
    const href = attrs.href || '';
    if (/^wiki:/i.test(href)) { await this.openWikiLink(href.slice(5)); return; }
    if (href) {
      if (/^(https?:|mailto:)/i.test(href)) { this.host.openExternal(href); return; }
      let path = null;
      let anchor = null;
      if (/^file:/i.test(href)) path = decodeURI(href.replace(/^file:\/\/\/?/, '')).replace(/\//g, '\\');
      else if (this.filePath && !/^[a-z]+:/i.test(href)) {
        const [rel, hash] = href.split('#');
        anchor = hash || null;
        let clean = rel;
        try { clean = decodeURIComponent(rel); } catch { /* mantém */ }
        path = dirName(this.filePath) + '\\' + clean.replace(/\//g, '\\');
      }
      if (!path) return;
      if (EXT_FORMAT[extOf(path)]) {
        if (this.filePath && path.toLowerCase() === this.filePath.toLowerCase()) { if (anchor) this.gotoBookmark(anchor); return; }
        if (await this.host.fileExists(path)) { this.openPath(path); return; }
      }
      this.host.openPath(path);
      return;
    }
    if (attrs.anchor) this.gotoBookmark(attrs.anchor);
  }

  // [[Nota]] do Obsidian: procura a nota no cofre e abre
  async openWikiLink(target) {
    const [name, heading] = target.split('#');
    if (!name.trim()) { if (heading) this.gotoBookmark(heading); return; }
    if (!this.filePath || !this.host.findNote) return;
    const found = await this.host.findNote(this.filePath, target);
    if (!found) { this.setInfo('Nota "' + name + '" não encontrada'); return; }
    if (found.toLowerCase() === this.filePath.toLowerCase()) { if (heading) this.gotoBookmark(heading); return; }
    if (EXT_FORMAT[extOf(found)]) this.openPath(found);
    else this.host.openPath(found);
  }

  gotoBookmark(name) {
    let target = null;
    this.view.state.doc.descendants((n, pos) => {
      if (target != null) return false;
      if (n.type === schema.nodes.opaque_inline && n.attrs.kind === 'bookmarkStart' && (n.attrs.xml || '').includes('w:name="' + name + '"')) target = pos;
      return true;
    });
    if (target == null) {
      const hd = this.headingList().find((x) => x.bookmark === name);
      if (hd) target = hd.pos;
    }
    if (target == null) {
      // âncoras de Markdown: #titulo-da-secao, como no GitHub, ou o próprio texto do título
      const slug = (s) => String(s).toLowerCase().trim().replace(/[^\p{L}\p{N}\s-]/gu, '').replace(/\s/g, '-');
      let want = name;
      try { want = decodeURIComponent(name); } catch { /* mantém */ }
      const hd = this.headingList().find((x) => slug(x.text) === slug(want) || x.text.toLowerCase() === want.toLowerCase());
      if (hd) target = hd.pos;
    }
    if (target != null) this.gotoPos(target);
  }

  ensureBookmarkForHeading(name) {
    const hd = this.headingList().find((x) => x.bookmark === name);
    if (!hd) return;
    const v = this.view;
    const node = v.state.doc.nodeAt(hd.pos);
    let has = false;
    node.forEach((c) => { if (c.type === schema.nodes.opaque_inline && (c.attrs.xml || '').includes('w:name="' + name + '"')) has = true; });
    if (has) return;
    const id = String(Math.floor(Math.random() * 1e6));
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const start = schema.nodes.opaque_inline.create({ xml: '<w:bookmarkStart ' + W + ' w:id="' + id + '" w:name="' + name + '"/>', level: 'p', kind: 'bookmarkStart', id });
    const end = schema.nodes.opaque_inline.create({ xml: '<w:bookmarkEnd ' + W + ' w:id="' + id + '"/>', level: 'p', kind: 'bookmarkEnd', id });
    const tr = v.state.tr.insert(hd.pos + node.nodeSize - 1, end).insert(hd.pos + 1, start);
    v.dispatch(tr);
  }

  currentLink() {
    const v = this.activeView;
    const r = C.linkRangeAt(v.state, v.state.selection.from);
    if (!r) return null;
    return { href: r.mark.attrs.href, anchor: r.mark.attrs.anchor, text: v.state.doc.textBetween(r.from, r.to) };
  }

  // ------------------------------------------------------------ imagens

  async insertPictureFromFile() {
    const files = await this.host.imageDialog();
    for (const f of files) await this.insertImageBytes(f.data, f.name, this.activeView);
  }

  async insertImageFiles(files, view) {
    for (const f of files) await this.insertImageBytes(new Uint8Array(await f.arrayBuffer()), f.name || 'imagem.png', view || this.activeView, f.type);
  }

  async insertImageBytes(bytes, name, view, type) {
    let ext = extOf(name) || (type ? type.split('/')[1] : 'png');
    if (ext === 'jpeg') ext = 'jpg';
    if (ext === 'svg+xml') ext = 'svg';
    const mime = MIME[ext] || type || 'image/png';
    let n = 1;
    while (this.model.media.has('word/media/imagem' + n + '.' + ext) || (this.model.zip && this.model.zip.file('word/media/imagem' + n + '.' + ext))) n++;
    const path = 'word/media/imagem' + n + '.' + ext;
    let url = URL.createObjectURL(new Blob([bytes], { type: mime }));
    const entry = { url, bytes, ext, type: mime };
    this.model.media.set(path, entry);
    if (NEEDS_CONVERSION.has(ext)) {
      const png = await this.host.convertImage(bytes, ext, 0, 0);
      if (png) { url = URL.createObjectURL(new Blob([png], { type: 'image/png' })); entry.url = url; }
    }
    const size = await loadImageSize(url);
    // cabe na largura do parágrafo atual, que numa tabela é a largura da célula
    let maxW = contentWidthPx(this.model);
    try {
      const para = C.currentParagraph(view.state);
      const dom = para && view.nodeDOM(para.pos);
      if (dom && dom.clientWidth) maxW = Math.min(maxW, dom.clientWidth - 2);
    } catch { /* usa a largura da página */ }
    const { cx, cy } = size ? fitSize(size.w, size.h, maxW) : { cx: 1905000, cy: 1428750 };
    C.insertDrawing({ kind: 'image', mediaPath: path, cx, cy, ocx: null, ocy: null, alt: stripExt(name || ''), ext })(view.state, view.dispatch);
    view.focus();
  }

  selectedDrawing() {
    const v = this.activeView;
    const sel = v.state.selection;
    if (sel instanceof NodeSelection && sel.node.type === schema.nodes.drawing) return { node: sel.node, pos: sel.from };
    return null;
  }

  updateDrawing(attrs) {
    const d = this.selectedDrawing();
    if (!d) return;
    const v = this.activeView;
    const tr = v.state.tr.setNodeMarkup(d.pos, null, { ...d.node.attrs, ...attrs });
    tr.setSelection(NodeSelection.create(tr.doc, d.pos));
    v.dispatch(tr);
  }

  resizeDrawing(dim, cm) {
    const d = this.selectedDrawing();
    if (!d) return;
    const a = d.node.attrs;
    const ratio = a.cy / a.cx || 1;
    const emu = Math.round(cm * 360000);
    if (dim === 'w') this.updateDrawing({ cx: emu, cy: Math.round(emu * ratio) });
    else this.updateDrawing({ cy: emu, cx: Math.round(emu / ratio) });
  }

  async saveDrawingAsImage() {
    const d = this.selectedDrawing();
    if (!d) return;
    const m = d.node.attrs.mediaPath && this.model.media.get(d.node.attrs.mediaPath);
    if (!m || !m.bytes) return;
    const path = await this.host.saveDialog({ title: 'Salvar como Imagem', defaultPath: 'imagem.' + (m.ext || 'png'), filters: [{ name: 'Imagem', extensions: [m.ext || 'png'] }] });
    if (path) await this.host.writeFile(path, m.bytes);
  }

  // ------------------------------------------------------------ tabelas

  currentTable() { return C.findTableAt(this.activeView.state); }

  currentCell() {
    const $f = this.activeView.state.selection.$from;
    for (let d = $f.depth; d > 0; d--) if ($f.node(d).type === schema.nodes.table_cell) return { node: $f.node(d), pos: $f.before(d) };
    return null;
  }

  currentRow() {
    const $f = this.activeView.state.selection.$from;
    for (let d = $f.depth; d > 0; d--) if ($f.node(d).type === schema.nodes.table_row) return { node: $f.node(d), pos: $f.before(d) };
    return null;
  }

  currentColumnWidthCm() {
    const cell = this.currentCell();
    const t = this.currentTable();
    if (!cell || !t) return null;
    const map = TableMap.get(t.node);
    const rect = map.findCell(cell.pos - t.pos - 1);
    const w = columnWidths(t.node).slice(rect.left, rect.right).reduce((s, x) => s + x, 0);
    return twipToCm(w * 15);
  }

  setColumnWidthCm(cm) {
    const v = this.activeView;
    const cell = this.currentCell();
    const t = this.currentTable();
    if (!cell || !t) return;
    const map = TableMap.get(t.node);
    const rect = map.findCell(cell.pos - t.pos - 1);
    const widths = columnWidths(t.node);
    const newW = (cmToTwip(cm) / 15) / (rect.right - rect.left);
    for (let c = rect.left; c < rect.right; c++) widths[c] = newW;
    const tr = v.state.tr;
    t.node.descendants((c2, pos) => {
      if (c2.type !== schema.nodes.table_cell) return true;
      const r = map.findCell(pos);
      tr.setNodeMarkup(t.pos + 1 + pos, null, { ...c2.attrs, colwidth: widths.slice(r.left, r.right) });
      return false;
    });
    v.dispatch(tr);
  }

  setRowProps(p) {
    const v = this.activeView;
    const row = this.currentRow();
    if (!row) return;
    v.dispatch(v.state.tr.setNodeMarkup(row.pos, null, { ...row.node.attrs, height: p.height, hRule: p.hRule, header: p.header || null }));
  }

  selectTablePart(kind) {
    const v = this.activeView;
    const t = this.currentTable();
    const cell = this.currentCell();
    if (!t || !cell) return;
    const $cell = v.state.doc.resolve(cell.pos);
    let sel;
    if (kind === 'cell') sel = new CellSelection($cell);
    else if (kind === 'row') sel = CellSelection.rowSelection($cell);
    else if (kind === 'column') sel = CellSelection.colSelection($cell);
    else {
      const map = TableMap.get(t.node);
      sel = new CellSelection(v.state.doc.resolve(t.pos + 1 + map.map[0]), v.state.doc.resolve(t.pos + 1 + map.map[map.map.length - 1]));
    }
    v.dispatch(v.state.tr.setSelection(sel));
    v.focus();
  }

  setTableLook(flag, on) {
    const t = this.currentTable();
    if (!t) return;
    const lk = { ...(t.node.attrs.tblLook || { val: '04A0', firstRow: '1', lastRow: '0', firstColumn: '1', lastColumn: '0', noHBand: '0', noVBand: '1' }) };
    if (lk.firstRow == null && lk.val) {
      const v = parseInt(lk.val, 16);
      Object.assign(lk, { firstRow: v & 0x20 ? '1' : '0', lastRow: v & 0x40 ? '1' : '0', firstColumn: v & 0x80 ? '1' : '0', lastColumn: v & 0x100 ? '1' : '0', noHBand: v & 0x200 ? '1' : '0', noVBand: v & 0x400 ? '1' : '0' });
    }
    lk[flag] = on ? '1' : '0';
    let val = 0;
    if (lk.firstRow === '1') val |= 0x20;
    if (lk.lastRow === '1') val |= 0x40;
    if (lk.firstColumn === '1') val |= 0x80;
    if (lk.lastColumn === '1') val |= 0x100;
    if (lk.noHBand === '1') val |= 0x200;
    if (lk.noVBand === '1') val |= 0x400;
    lk.val = val.toString(16).toUpperCase().padStart(4, '0');
    this.cmd('tableAttrs', { tblLook: lk });
  }

  // ------------------------------------------------------------ página

  setPageSetup(p) {
    const s = this.model.section;
    if (p.margins) s.setMargins(p.margins);
    if (p.size) s.setPageSize(p.size.w, p.size.h, p.size.orient);
    if (p.titlePg != null) s.setTitlePg(p.titlePg);
    if (p.evenOdd != null && p.evenOdd !== !!this.model.settings.evenAndOdd) { this.model.settings.evenAndOdd = p.evenOdd; this.model.settings.dirty = true; }
    this.pageChanged();
  }

  setMargins(preset) {
    this.model.section.setMargins({ top: preset.top, bottom: preset.bottom, left: preset.left, right: preset.right });
    this.pageChanged();
  }

  setOrientation(o) {
    const s = this.model.section;
    const w = s.pgSz.w, hh = s.pgSz.h;
    const land = w > hh;
    if ((o === 'landscape') === land) return;
    s.setPageSize(hh, w, o === 'landscape' ? 'landscape' : null);
    const m = s.pgMar;
    s.setMargins(o === 'landscape' ? { top: m.left, bottom: m.right, left: m.bottom, right: m.top } : { top: m.right, bottom: m.left, left: m.top, right: m.bottom });
    this.pageChanged();
  }

  setPaper(p) {
    const s = this.model.section;
    const land = s.pgSz.w > s.pgSz.h;
    s.setPageSize(land ? p.h : p.w, land ? p.w : p.h, land ? 'landscape' : null);
    this.pageChanged();
  }

  setTitlePage(on) { this.model.section.setTitlePg(on); this.pageChanged(); }

  setEvenOdd(on) { this.model.settings.evenAndOdd = on; this.model.settings.dirty = true; this.pageChanged(); }

  setDefaultTabStop(tw) {
    this.model.settings.defaultTabStop = tw;
    this.model.settings.dirty = true;
    this.refreshStyles();
    this.markDirty();
  }

  pageChanged() {
    this.markDirty();
    this.applyDocCss();
    this.layout.invalidateStatic();
    this.layout.computeGeometry();
    this.layout.invalidate();
    this.layout.updateCanvas();
    this.ruler.update();
    this.scheduleUI();
  }

  insertBlankPage() {
    const v = this.activeView;
    if (!C.pageBreak(v.state, v.dispatch)) return;
    C.pageBreak(v.state, v.dispatch);
    const $f = v.state.selection.$from;
    const pos = $f.before();
    v.dispatch(v.state.tr.setSelection(TextSelection.near(v.state.doc.resolve(Math.max(1, pos - 1)), -1)));
  }

  // ------------------------------------------------------------ cabeçalho e rodapé

  editHeaderFooter(kind, page) {
    if (this.viewMode !== 'print') this.setViewMode('print');
    if (this.hfEdit) this.closeHeaderFooter(true);
    const k = page != null ? page : this.layout.pageOfPos(this.view.state.selection.head);
    let entry = this.model.headerFor(kind, k, k === 0);
    if (!entry) {
      entry = newHeaderFooter(this.model, kind);
      const s = this.model.section;
      const type = k === 0 && s.titlePg ? 'first' : this.model.settings.evenAndOdd && (k + 1) % 2 === 0 ? 'even' : 'default';
      s.setReference(kind, type, entry.rId);
      this.markDirty();
    }
    const pageEl = this.layout.pageEls[k];
    if (!pageEl) return;
    const container = pageEl.querySelector(kind === 'header' ? '.page-header' : '.page-footer');
    container.textContent = '';
    container.classList.add('editing');
    const tag = h('div', { class: 'hf-tag', text: kind === 'header' ? 'Cabeçalho' : 'Rodapé' });
    pageEl.appendChild(tag);
    const state = createState(entry.doc, this, {});
    const view = createView(container, state, this);
    view.dom.classList.add('hf-editor');
    this.hfEdit = { kind, page: k, entry, view, container, tag };
    this.layout.headerEditing = { kind, page: k, rId: entry.rId, height: () => container.getBoundingClientRect().height / this.layout.zoom };
    this.root.classList.add('hf-editing');
    this.activeView = view;
    this.ribbon.update({ ...this.uiState, headerFooter: true });
    this.ribbon.select('headerFooter');
    view.focus();
    this.scheduleUI();
  }

  commitHeaderFooter() {
    if (!this.hfEdit) return;
    this.hfEdit.entry.doc = this.hfEdit.view.state.doc;
  }

  closeHeaderFooter(silent) {
    const hf = this.hfEdit;
    if (!hf) return;
    this.commitHeaderFooter();
    hf.view.destroy();
    hf.container.classList.remove('editing');
    hf.tag.remove();
    this.hfEdit = null;
    this.layout.headerEditing = null;
    this.root.classList.remove('hf-editing');
    this.activeView = this.view;
    this.layout.invalidateStatic();
    this.layout.pageKeys = [];
    this.layout.invalidate();
    if (!silent) {
      this.ribbon.select('home');
      this.focusEditor();
    }
    this.scheduleUI();
  }

  switchHeaderFooter(kind) {
    const page = this.hfEdit ? this.hfEdit.page : 0;
    this.closeHeaderFooter(true);
    this.editHeaderFooter(kind, page);
  }

  removeHeaderFooter(kind) {
    const s = this.model.section;
    const refs = kind === 'header' ? s.headers : s.footers;
    const map = kind === 'header' ? this.model.headers : this.model.footers;
    for (const rId of Object.values(refs)) {
      const e = map.get(rId);
      if (e) { e.doc = schema.node('doc', null, [schema.nodes.paragraph.create()]); e.dirty = true; }
    }
    this.markDirty();
    this.layout.invalidateStatic();
    this.layout.invalidate();
  }

  pageNumberParagraph(styleId, align, xOfY) {
    const f = (text) => [
      schema.nodes.field_char.create({ type: 'begin' }), schema.nodes.instr_text.create({ text }),
      schema.nodes.field_char.create({ type: 'separate' }), schema.text('1'), schema.nodes.field_char.create({ type: 'end' }),
    ];
    const content = xOfY ? [schema.text('Página '), ...f(' PAGE '), schema.text(' de '), ...f(' NUMPAGES ')] : f(' PAGE ');
    return schema.nodes.paragraph.create({ styleId, jc: align === 'left' ? null : align }, content);
  }

  insertPageNumber(kind, align, xOfY) {
    if (this.hfEdit) this.closeHeaderFooter(true);
    let entry = this.model.headerFor(kind, 1, false) || this.model.headerFor(kind, 0, true);
    if (!entry) {
      entry = newHeaderFooter(this.model, kind);
      this.model.section.setReference(kind, 'default', entry.rId);
    }
    const st = this.model.styles.findByName(kind);
    const p = this.pageNumberParagraph(st ? st.id : null, align, xOfY);
    const doc = entry.doc;
    const empty = doc.childCount === 1 && doc.firstChild.content.size === 0;
    entry.doc = schema.node('doc', null, empty ? [p] : [...doc.content.content, p]);
    entry.dirty = true;
    this.markDirty();
    this.layout.invalidateStatic();
    this.layout.pageKeys = [];
    this.layout.invalidate();
  }

  async pageNumberStart() {
    const inp = D.textInput(String(this.model.section.pgNumType.start || 1), { width: '80px' });
    const fmt = D.select([['decimal', '1, 2, 3, ...'], ['lowerRoman', 'i, ii, iii, ...'], ['upperRoman', 'I, II, III, ...'], ['lowerLetter', 'a, b, c, ...'], ['upperLetter', 'A, B, C, ...']], this.model.section.pgNumType.fmt || 'decimal');
    const r = await new D.Dialog('Formatar Números de Página', h('div', {}, D.field('Formato do número:', fmt), D.field('Iniciar em:', inp)), [{ label: 'OK', primary: true, value: 'ok' }, { label: 'Cancelar', value: null }], { className: 'dlg-small' }).open();
    if (r !== 'ok') return;
    const n = parseInt(inp.value, 10);
    this.model.section.setPageNumbering(fmt.value === 'decimal' ? null : fmt.value, Number.isFinite(n) ? n : null);
    this.pageChanged();
  }

  // ------------------------------------------------------------ sumário

  tocStyleId(level) {
    const st = this.model.styles;
    const found = st.findByName('toc ' + level);
    if (found) return found.id;
    const s = this.model.section;
    const pos = s.pgSz.w - s.pgMar.left - s.pgMar.right - 10;
    const id = 'TOC' + level;
    st.addStyleXml('<w:style w:type="paragraph" w:styleId="' + id + '"><w:name w:val="toc ' + level + '"/><w:basedOn w:val="' + st.defaultPara + '"/><w:next w:val="' + st.defaultPara + '"/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="' + pos + '"/></w:tabs><w:spacing w:after="100"/>' + (level > 1 ? '<w:ind w:left="' + (level - 1) * 220 + '"/>' : '') + '</w:pPr></w:style>');
    this.applyDocCss();
    return id;
  }

  buildTocBlocks() {
    const heads = this.headingList().filter((hd) => hd.level <= 3);
    for (const hd of heads) this.ensureBookmarkForHeading(hd.bookmark);
    const fresh = this.headingList().filter((hd) => hd.level <= 3);
    const start = this.model.section.pgNumType.start || 1;
    const title = this.model.styles.findByName('TOC Heading');
    const blocks = [schema.nodes.paragraph.create({ styleId: title ? title.id : R.headingStyleId(1) }, [schema.text('Sumário')])];
    const linkStyle = this.model.styles.get('Hyperlink') ? null : null;
    void linkStyle;
    fresh.forEach((hd, i) => {
      const link = schema.marks.hyperlink.create({ anchor: hd.bookmark });
      const page = String(this.layout.pageOfPos(hd.pos) + start);
      const content = [];
      if (i === 0) content.push(schema.nodes.field_char.create({ type: 'begin' }), schema.nodes.instr_text.create({ text: ' TOC \\o "1-3" \\h \\z \\u ' }), schema.nodes.field_char.create({ type: 'separate' }));
      content.push(schema.text(hd.text, [link]), schema.nodes.tab.create(null, null, [link]), schema.text(page, [link]));
      if (i === fresh.length - 1) content.push(schema.nodes.field_char.create({ type: 'end' }));
      blocks.push(schema.nodes.paragraph.create({ styleId: this.tocStyleId(hd.level) }, content));
    });
    if (!fresh.length) blocks.push(schema.nodes.paragraph.create(null, [schema.text('Nenhuma entrada de sumário encontrada. Aplique estilos de Título ao texto para criar o sumário.')]));
    const pr = '<w:sdtPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj></w:sdtPr>';
    return schema.nodes.sdt_block.create({ pr, endPr: null }, blocks);
  }

  findToc() {
    let found = null;
    this.view.state.doc.forEach((n, pos) => { if (!found && n.type === schema.nodes.sdt_block && /Table of Contents/.test(n.attrs.pr || '')) found = { node: n, pos }; });
    return found;
  }

  insertTOC() {
    const v = this.view;
    if (this.findToc()) { this.updateTOC(); return; }
    const toc = this.buildTocBlocks();
    const p = C.currentParagraph(v.state);
    const pos = p ? v.state.doc.resolve(p.pos).before(1) : 0;
    v.dispatch(v.state.tr.insert(Math.max(0, pos), toc).scrollIntoView());
    this.layout.runNow(4);
    this.updateTOC();
  }

  updateTOC() {
    const v = this.view;
    const t = this.findToc();
    if (!t) { this.insertTOC(); return; }
    const fresh = this.buildTocBlocks();
    const t2 = this.findToc();
    v.dispatch(v.state.tr.replaceWith(t2.pos, t2.pos + t2.node.nodeSize, fresh));
    this.setInfo('Sumário atualizado');
  }

  removeTOC() {
    const t = this.findToc();
    if (!t) return;
    const v = this.view;
    v.dispatch(v.state.tr.delete(t.pos, t.pos + t.node.nodeSize));
  }

  // ------------------------------------------------------------ alterações controladas

  acceptAllChanges() {
    const v = this.view;
    const tr = v.state.tr;
    const dels = [];
    v.state.doc.descendants((n, pos) => { if (n.type === schema.nodes.opaque_inline && n.attrs.kind === 'del') dels.push(pos); return true; });
    for (const pos of dels.reverse()) tr.delete(pos, pos + 1);
    tr.removeMark(0, tr.doc.content.size, schema.marks.ins);
    v.dispatch(tr);
    this.setInfo('Alterações aceitas');
  }

  rejectAllChanges() {
    const v = this.view;
    const tr = v.state.tr;
    const ranges = [];
    v.state.doc.descendants((n, pos) => {
      if (n.isInline && n.marks.some((m) => m.type === schema.marks.ins)) ranges.push({ from: pos, to: pos + n.nodeSize, del: false });
      else if (n.type === schema.nodes.opaque_inline && n.attrs.kind === 'del') ranges.push({ from: pos, to: pos + 1, del: true, text: n.attrs.text, marks: n.marks });
      return true;
    });
    for (const r of ranges.reverse()) {
      if (r.del) tr.replaceWith(r.from, r.to, r.text ? schema.text(r.text, r.marks) : Fragment.empty);
      else tr.delete(r.from, r.to);
    }
    v.dispatch(tr);
    this.setInfo('Alterações rejeitadas');
  }
}

// pequenos serializadores usados ao regravar estilos
import { pPrXml, rPrXml } from './docx/props.js';
function pPrXmlSafe(p) { return pPrXml(p); }
function rPrXmlSafe(r) { return rPrXml(r); }

// ---------------------------------------------------------------- partida

const app = new App(document.getElementById('app'));
window.redator = app;
window.__rt = testApi;
app.init().catch((err) => {
  console.error(err);
  document.body.appendChild(Object.assign(document.createElement('pre'), { textContent: 'Falha ao iniciar: ' + (err && err.stack || err) }));
});
