// Visualizações de nós: imagens e formas, tabelas, tabulações.
import { Node as PMNode } from 'prosemirror-model';
import { NodeSelection } from 'prosemirror-state';
import { R, tableWrapCss, tableDirectCss, tableClasses, shdColor } from './render.js';
import { placeholderImage } from '../docx/reader.js';

const px = (emu) => (Number(emu) || 0) / 9525;

export function mediaSrc(a) {
  if (a.mediaPath && R.model) {
    const u = R.model.mediaUrl(a.mediaPath);
    if (u) return u;
  }
  if (a.src) return a.src;
  return placeholderImage(a.ext || a.kind);
}

function colorValue(c) {
  if (!c || c === 'none') return null;
  if (c.startsWith('theme:')) {
    const t = R.model && R.model.theme;
    const map = { tx1: 'dk1', bg1: 'lt1', tx2: 'dk2', bg2: 'lt2' };
    const k = c.slice(6);
    return t ? t.hex(map[k] || k) : null;
  }
  return c;
}

// Posicionamento de objetos flutuantes, aproximado para o fluxo HTML
export function anchorCss(a, w, h) {
  const an = a.anchor;
  if (!an) return { mode: 'inline', css: '' };
  const d = (v) => px(v).toFixed(1) + 'px';
  if (an.wrap === 'topAndBottom') {
    let ml = 0;
    if (an.alignH === 'center') return { mode: 'block', css: 'display:block;margin:' + d(an.distT) + ' auto ' + d(an.distB) + ' auto' };
    if (an.alignH === 'right') return { mode: 'block', css: 'display:block;margin:' + d(an.distT) + ' 0 ' + d(an.distB) + ' auto' };
    if (an.offH) ml = px(an.offH) - (an.relH === 'page' ? pageMarginLeftPx() : 0);
    return { mode: 'block', css: 'display:block;margin:' + d(an.distT) + ' 0 ' + d(an.distB) + ' ' + Math.max(0, ml).toFixed(1) + 'px' };
  }
  if (an.wrap === 'square' || an.wrap === 'tight' || an.wrap === 'through') {
    const contentW = contentWidthPx();
    let side = 'left';
    if (an.alignH === 'right' || an.alignH === 'outside') side = 'right';
    else if (an.alignH === 'center') side = 'left';
    else if (!an.alignH) {
      let x = px(an.offH);
      if (an.relH === 'page') x -= pageMarginLeftPx();
      if (x + w / 2 > contentW / 2) side = 'right';
    }
    const top = an.relV === 'paragraph' || an.relV === 'line' ? Math.max(0, px(an.offV)) : 0;
    const margin = side === 'left'
      ? top.toFixed(1) + 'px ' + d(an.distR) + ' ' + d(an.distB) + ' 0'
      : top.toFixed(1) + 'px 0 ' + d(an.distB) + ' ' + d(an.distL);
    return { mode: 'float', css: 'float:' + side + ';margin:' + margin };
  }
  // sem quebra: posição absoluta relativa ao parágrafo
  let left = px(an.offH);
  const contentW = contentWidthPx();
  if (an.relH === 'page') left -= pageMarginLeftPx();
  if (an.alignH === 'center') left = (an.relH === 'page' ? (contentW + pageMarginLeftPx() * 2) / 2 - pageMarginLeftPx() : contentW / 2) - w / 2;
  else if (an.alignH === 'right') left = (an.relH === 'page' ? contentW + pageMarginLeftPx() : contentW) - w;
  else if (an.alignH === 'left') left = an.relH === 'page' ? -pageMarginLeftPx() : 0;
  let top = px(an.offV);
  if (an.relV === 'page' || an.relV === 'margin' || an.relV === 'topMargin') top = 0;
  const z = an.behind ? -1 : 3;
  return { mode: 'absolute', css: 'position:absolute;left:' + left.toFixed(1) + 'px;top:' + top.toFixed(1) + 'px;z-index:' + z };
}

function pageMarginLeftPx() {
  const s = R.model && R.model.section;
  return s ? s.pgMar.left / 15 : 113.4;
}

function contentWidthPx() {
  const s = R.model && R.model.section;
  return s ? (s.pgSz.w - s.pgMar.left - s.pgMar.right) / 15 : 604.7;
}

// Constrói o DOM de um desenho, usado pelo editor e pela renderização estática
export function drawingDOM(node, opts = {}) {
  const a = node.attrs;
  const w = px(a.cx);
  const h = px(a.cy);
  const wrap = document.createElement('span');
  wrap.className = 'drw-wrap drw-' + (a.kind || 'image');
  wrap.contentEditable = 'false';
  const pos = anchorCss(a, w, h);
  wrap.dataset.mode = pos.mode;
  if (a.kind === 'hr') {
    wrap.className += ' drw-hr';
    wrap.style.cssText = 'display:block;height:' + Math.max(1, h).toFixed(1) + 'px';
    return { dom: wrap, img: null };
  }
  let img = null;
  const box = document.createElement('span');
  box.className = 'drw-box';
  box.style.width = w > 0 ? w.toFixed(2) + 'px' : '';
  box.style.height = h > 0 ? h.toFixed(2) + 'px' : '';
  if (a.kind === 'image' || a.kind === 'object' || (a.mediaPath && !a.content)) {
    img = document.createElement('img');
    img.draggable = false;
    img.src = mediaSrc(a);
    img.alt = a.alt || '';
    if (a.title || a.alt) img.title = a.title || a.alt;
    if (a.crop) {
      const c = a.crop;
      const sw = 1 - (c.l + c.r) / 100000;
      const sh = 1 - (c.t + c.b) / 100000;
      img.style.cssText = 'position:absolute;width:' + (w / sw).toFixed(2) + 'px;height:' + (h / sh).toFixed(2) + 'px;left:' + (-(w / sw) * c.l / 100000).toFixed(2) + 'px;top:' + (-(h / sh) * c.t / 100000).toFixed(2) + 'px;max-width:none';
      box.style.overflow = 'hidden';
      box.style.position = 'relative';
    } else if (!(w > 0 && h > 0)) {
      img.style.maxWidth = contentWidthPx() + 'px';
    }
    box.appendChild(img);
  } else {
    box.classList.add('drw-shapebox');
    const fill = colorValue(a.fill);
    if (fill) box.style.background = fill;
    if (a.line && a.line !== 'none') {
      const lc = colorValue(a.line.color) || '#000';
      box.style.border = Math.max(0.75, (a.line.w || 9525) / 12700).toFixed(2) + 'pt solid ' + lc;
    }
    if (a.content && a.content.length) {
      const inner = document.createElement('div');
      inner.className = 'drw-text';
      try {
        const frag = opts.renderBlocks ? opts.renderBlocks(a.content) : null;
        if (frag) inner.appendChild(frag);
      } catch (err) { console.warn(err); }
      box.appendChild(inner);
      box.style.height = '';
      box.style.minHeight = h > 0 ? h.toFixed(2) + 'px' : '';
    } else if (a.kind === 'chart' || a.kind === 'smartart' || a.kind === 'canvas' || a.kind === 'group' || a.kind === 'other') {
      const lbl = document.createElement('span');
      lbl.className = 'drw-label';
      lbl.textContent = { chart: 'Gráfico', smartart: 'SmartArt', canvas: 'Tela de desenho', group: 'Grupo de formas', other: 'Objeto' }[a.kind] || 'Objeto';
      box.appendChild(lbl);
    }
  }
  wrap.appendChild(box);
  if (pos.css) wrap.style.cssText += ';' + pos.css;
  return { dom: wrap, img, box };
}

// Imagens: redimensionamento com alças, como no Word
export class DrawingView {
  constructor(node, view, getPos, renderBlocks) {
    this.node = node;
    this.view = view;
    this.getPos = getPos;
    this.renderBlocks = renderBlocks;
    this.build();
    this.unsub = R.model ? R.model.onMedia((path, url) => {
      if (this.img && path === this.node.attrs.mediaPath) this.img.src = url;
    }) : null;
  }

  build() {
    const { dom, img, box } = drawingDOM(this.node, { renderBlocks: this.renderBlocks });
    this.dom = dom;
    this.img = img;
    this.box = box;
    if (img) img.addEventListener('load', () => { if (this.view.redator) this.view.redator.layoutSoon(); });
    if (box && (this.node.attrs.kind === 'image' || this.node.attrs.kind === 'object' || img)) this.addHandles();
  }

  addHandles() {
    for (const h of ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w']) {
      const el = document.createElement('span');
      el.className = 'drw-handle drw-h-' + h;
      el.addEventListener('mousedown', (e) => this.startResize(e, h));
      this.box.appendChild(el);
    }
  }

  startResize(e, handle) {
    e.preventDefault();
    e.stopPropagation();
    const zoom = (this.view.redator && this.view.redator.zoom) || 1;
    const startX = e.clientX, startY = e.clientY;
    const w0 = px(this.node.attrs.cx) || this.box.offsetWidth;
    const h0 = px(this.node.attrs.cy) || this.box.offsetHeight;
    const ratio = h0 / w0 || 1;
    const corner = handle.length === 2;
    let w = w0, h = h0;
    const move = (ev) => {
      const dx = (ev.clientX - startX) / zoom;
      const dy = (ev.clientY - startY) / zoom;
      if (handle.includes('e')) w = w0 + dx;
      if (handle.includes('w')) w = w0 - dx;
      if (handle.includes('s')) h = h0 + dy;
      if (handle.includes('n')) h = h0 - dy;
      if (corner && !ev.shiftKey) {
        if (Math.abs(dx) > Math.abs(dy)) h = w * ratio; else w = h / ratio;
      }
      w = Math.max(8, w);
      h = Math.max(8, h);
      this.box.style.width = w.toFixed(1) + 'px';
      this.box.style.height = h.toFixed(1) + 'px';
      if (this.img && this.node.attrs.crop) {
        this.img.style.width = '100%';
        this.img.style.height = '100%';
        this.img.style.left = '0';
        this.img.style.top = '0';
      }
    };
    const up = () => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      const pos = this.getPos();
      if (pos == null) return;
      const tr = this.view.state.tr.setNodeMarkup(pos, null, { ...this.node.attrs, cx: Math.round(w * 9525), cy: Math.round(h * 9525) });
      tr.setSelection(NodeSelection.create(tr.doc, pos));
      this.view.dispatch(tr);
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  }

  selectNode() { this.dom.classList.add('selected'); }
  deselectNode() { this.dom.classList.remove('selected'); }

  update(node) {
    if (node.type !== this.node.type) return false;
    const a = node.attrs, b = this.node.attrs;
    if (a.cx === b.cx && a.cy === b.cy && a.mediaPath === b.mediaPath && a.src === b.src && a.anchor === b.anchor && a.content === b.content && a.crop === b.crop) {
      this.node = node;
      return true;
    }
    return false;
  }

  stopEvent(e) {
    return e.type === 'mousedown' && e.target.classList && e.target.classList.contains('drw-handle');
  }

  ignoreMutation() { return true; }

  destroy() { if (this.unsub) this.unsub(); }
}

// Tabulação: largura calculada pelo layout
export class TabView {
  constructor(node) {
    this.node = node;
    this.dom = document.createElement('span');
    this.dom.className = 'tab';
    this.dom.contentEditable = 'false';
  }
  update(node) {
    if (node.type !== this.node.type) return false;
    this.node = node;
    return true;
  }
  ignoreMutation() { return true; }
}

// Tabela no estilo do Word, compatível com o redimensionamento de colunas
export class TableView {
  constructor(node, cellMinWidth) {
    this.node = node;
    this.cellMinWidth = cellMinWidth || 10;
    this.dom = document.createElement('div');
    this.dom.className = 'tbl-wrap';
    this.table = this.dom.appendChild(document.createElement('table'));
    this.colgroup = this.table.appendChild(document.createElement('colgroup'));
    this.contentDOM = this.table.appendChild(document.createElement('tbody'));
    this.render();
  }

  render() {
    const n = this.node;
    this.dom.style.cssText = tableWrapCss(n.attrs);
    this.table.className = tableClasses(n);
    this.table.style.cssText = tableDirectCss(n.attrs);
    this.updateColumns();
  }

  updateColumns(overrideCol, overrideValue) {
    const widths = columnWidths(this.node);
    let total = 0;
    let next = this.colgroup.firstChild;
    widths.forEach((w, i) => {
      const val = overrideCol === i ? overrideValue : w;
      total += val;
      if (!next) { next = document.createElement('col'); this.colgroup.appendChild(next); }
      next.style.width = val.toFixed(2) + 'px';
      next = next.nextSibling;
    });
    while (next) { const after = next.nextSibling; next.remove(); next = after; }
    this.table.style.width = total.toFixed(2) + 'px';
  }

  update(node) {
    if (node.type !== this.node.type) return false;
    this.node = node;
    this.render();
    return true;
  }

  ignoreMutation(record) {
    return record.type === 'attributes' && (record.target === this.table || record.target === this.dom || this.colgroup.contains(record.target));
  }
}

// Larguras de coluna em px a partir das células (ou da grade do Word)
export function columnWidths(table) {
  const grid = table.attrs.grid || [];
  const widths = [];
  const row = table.firstChild;
  if (row) {
    let col = 0;
    for (let i = 0; i < row.childCount; i++) {
      const cell = row.child(i);
      const span = cell.attrs.colspan || 1;
      for (let j = 0; j < span; j++, col++) {
        const cw = cell.attrs.colwidth && cell.attrs.colwidth[j];
        widths[col] = cw || (grid[col] ? grid[col] / 15 : null);
      }
    }
  }
  const n = Math.max(widths.length, grid.length);
  const known = widths.filter((x) => x);
  const avg = known.length ? known.reduce((s, x) => s + x, 0) / known.length : contentWidthPx() / Math.max(1, n);
  for (let i = 0; i < n; i++) if (!widths[i]) widths[i] = grid[i] ? grid[i] / 15 : avg;
  return widths;
}

// Nó opaco: comentários mostram o texto ao passar o mouse
export class OpaqueView {
  constructor(node) {
    this.node = node;
    const a = node.attrs;
    this.dom = document.createElement('span');
    this.dom.className = 'oi oi-' + a.kind;
    this.dom.contentEditable = 'false';
    if (a.kind === 'commentRef' && R.model) {
      const c = R.model.comments.get(a.id);
      if (c) this.dom.title = (c.author ? c.author + ': ' : '') + c.text;
    } else if (a.kind === 'del') {
      this.dom.textContent = a.text || '';
      this.dom.title = 'Excluído' + (a.id ? ' por ' + a.id : '');
    } else if (a.kind === 'math') {
      this.dom.textContent = a.text || '';
      this.dom.title = 'Equação';
    } else if (a.kind === 'alternate' || a.kind === 'object') {
      this.dom.textContent = a.text || '';
    }
  }
  update(node) {
    if (node.type !== this.node.type || node.attrs !== this.node.attrs) return false;
    this.node = node;
    return true;
  }
  ignoreMutation() { return true; }
}
