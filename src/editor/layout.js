// Layout de páginas: tabulações, paginação com quebras de verdade,
// cabeçalhos, rodapés e notas de rodapé em cada página.
import { Plugin, PluginKey } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import { R } from './render.js';
import { renderStatic, fillFields } from './static.js';

export const layoutKey = new PluginKey('layout');

const EPS = 0.6;

function breakWidget(b) {
  const key = b.kind + ':' + b.h.toFixed(2) + ':' + (b.cols || 0);
  if (b.kind === 'row') {
    return Decoration.widget(b.pos, () => {
      const tr = document.createElement('tr');
      tr.className = 'pb pb-row';
      tr.dataset.h = String(b.h);
      const td = document.createElement('td');
      td.colSpan = b.cols || 1;
      td.style.height = b.h + 'px';
      tr.appendChild(td);
      return tr;
    }, { side: -1, key, pb: b, ignoreSelection: true });
  }
  if (b.kind === 'inline') {
    return Decoration.widget(b.pos, () => {
      const s = document.createElement('span');
      s.className = 'pb pb-inline';
      s.dataset.h = String(b.h);
      s.style.height = b.h + 'px';
      return s;
    }, { side: -1, marks: [], key, pb: b, ignoreSelection: true });
  }
  return Decoration.widget(b.pos, () => {
    const d = document.createElement('div');
    d.className = 'pb pb-block';
    d.dataset.h = String(b.h);
    d.style.height = b.h + 'px';
    d.contentEditable = 'false';
    return d;
  }, { side: -1, key, pb: b, ignoreSelection: true });
}

export function currentBreaks(state) {
  const set = layoutKey.getState(state);
  if (!set) return [];
  return set.find().map((d) => ({ ...d.spec.pb, pos: d.from })).sort((a, b) => a.pos - b.pos);
}

function sameBreaks(a, b) {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) {
    if (a[i].pos !== b[i].pos || a[i].kind !== b[i].kind || Math.abs(a[i].h - b[i].h) > 0.75) return false;
  }
  return true;
}

export function layoutPlugin(layout) {
  return new Plugin({
    key: layoutKey,
    state: {
      init: () => DecorationSet.empty,
      apply(tr, set) {
        const meta = tr.getMeta(layoutKey);
        if (meta) return DecorationSet.create(tr.doc, meta.breaks.map(breakWidget));
        return tr.docChanged ? set.map(tr.mapping, tr.doc) : set;
      },
    },
    props: {
      decorations(state) { return layoutKey.getState(state); },
    },
    view(view) {
      layout.attach(view);
      return {
        update(v, prev) { if (v.state.doc !== prev.doc) layout.schedule(true); },
        destroy() { layout.detach(view); },
      };
    },
  });
}

// ---------------------------------------------------------------- tabulações

const tabCache = new WeakMap();

function tabInfoOf(p) {
  const raw = p.getAttribute('data-tab');
  if (!raw) return null;
  const c = tabCache.get(p);
  if (c && c.raw === raw) return c.info;
  let info = null;
  try { info = JSON.parse(raw); } catch { info = null; }
  tabCache.set(p, { raw, info });
  return info;
}

function nextStop(info, xTw) {
  let best = null;
  for (const [pos, val, leader] of info.s) {
    if (val === 'bar' || val === 'clear') continue;
    if (pos > xTw + 3) { best = { pos, val, leader }; break; }
  }
  if (info.h > 0 && info.l > xTw + 3 && (!best || info.l < best.pos)) best = { pos: info.l, val: 'left', leader: 'none' };
  if (!best) {
    const d = info.d || 708;
    best = { pos: (Math.floor((xTw + 3) / d) + 1) * d, val: 'left', leader: 'none' };
  }
  return best;
}

export function layoutTabs(root, zoom) {
  const z = zoom || 1;
  // rótulos de numeração: largura até a próxima parada
  const labels = root.querySelectorAll('.num-lbl');
  const reads = [];
  for (const lbl of labels) {
    const tab = lbl.lastElementChild && lbl.lastElementChild.classList.contains('num-tab') ? lbl.lastElementChild : null;
    if (!tab) continue;
    const p = lbl.parentElement && lbl.parentElement.closest('p');
    if (!p) continue;
    const info = tabInfoOf(p);
    if (!info) continue;
    const txt = lbl.firstElementChild;
    reads.push({ tab, info, w: txt ? txt.getBoundingClientRect().width / z : 0, right: lbl.classList.contains('num-right') });
  }
  for (const r of reads) {
    const start = r.info.l - r.info.h + (r.info.h ? 0 : r.info.f);
    const x = r.right ? start : start + r.w * 15;
    const stop = nextStop(r.info, x);
    r.tab.style.width = Math.max(0, (stop.pos - x) / 15).toFixed(2) + 'px';
    if (r.right) r.tab.previousElementSibling.style.marginLeft = (-r.w).toFixed(2) + 'px';
  }
  // tabulações do texto, parágrafo a parágrafo
  const tabs = root.querySelectorAll('span.tab');
  const byPara = new Map();
  for (const t of tabs) {
    const p = t.closest('p');
    if (!p) continue;
    if (!byPara.has(p)) byPara.set(p, []);
    byPara.get(p).push(t);
  }
  // Parágrafos só com paradas à esquerda: zera tudo, mede tudo de uma vez e calcula
  // as larguras em seguida. Medir tabulação por tabulação força o navegador a refazer
  // o layout do documento a cada uma, o que trava documentos com milhares delas.
  const simple = [];
  const complex = [];
  for (const [p, list] of byPara) {
    const info = tabInfoOf(p) || { s: [], l: 0, h: 0, f: 0, d: 708 };
    const aligned = info.s.some(([, val]) => val === 'right' || val === 'end' || val === 'center' || val === 'decimal' || val === 'num');
    (aligned ? complex : simple).push([p, list, info]);
  }
  for (const [, list] of simple) for (const t of list) if (t.style.width !== '0px') t.style.width = '0px';
  const measured = simple.map(([p, list, info]) => {
    const cs = getComputedStyle(p);
    const originX = p.getBoundingClientRect().left - (parseFloat(cs.marginLeft) || 0) * z;
    return { list, info, originX, rects: list.map((t) => t.getBoundingClientRect()) };
  });
  for (const m of measured) {
    let lineTop = null;
    let add = 0;
    for (let i = 0; i < m.list.length; i++) {
      const r = m.rects[i];
      const top = Math.round(r.top);
      if (lineTop == null || Math.abs(top - lineTop) > 2) { lineTop = top; add = 0; }
      const xTw = ((r.left - m.originX) / z + add) * 15;
      const stop = nextStop(m.info, xTw);
      const w = Math.max(0, (stop.pos - xTw) / 15);
      const t = m.list[i];
      t.style.width = w.toFixed(2) + 'px';
      add += w;
      const leader = stop.leader && stop.leader !== 'none' ? stop.leader : '';
      if (t.dataset.leader !== leader) t.dataset.leader = leader;
    }
  }
  for (const [p, list, info] of complex) {
    const cs = getComputedStyle(p);
    const originX = p.getBoundingClientRect().left - (parseFloat(cs.marginLeft) || 0) * z;
    for (let i = 0; i < list.length; i++) {
      const t = list[i];
      t.style.width = '0px';
      const tr = t.getBoundingClientRect();
      const xTw = ((tr.left - originX) / z) * 15;
      const stop = nextStop(info, xTw);
      let w = (stop.pos - xTw) / 15;
      if (stop.val === 'right' || stop.val === 'end' || stop.val === 'center' || stop.val === 'decimal' || stop.val === 'num') {
        const seg = segmentWidth(t, list[i + 1] || null, p, stop.val === 'decimal', z);
        if (stop.val === 'center') w -= seg / 2;
        else w -= seg;
      }
      t.style.width = Math.max(0, w).toFixed(2) + 'px';
      const leader = stop.leader && stop.leader !== 'none' ? stop.leader : '';
      if (t.dataset.leader !== leader) t.dataset.leader = leader;
    }
  }
}

function segmentWidth(tab, nextTab, p, decimal, z) {
  const range = document.createRange();
  range.setStartAfter(tab);
  if (nextTab) range.setEndBefore(nextTab);
  else range.setEnd(p, p.childNodes.length);
  const tr = tab.getBoundingClientRect();
  const lineMid = tr.top + tr.height / 2;
  let right = tr.right;
  let found = false;
  if (decimal) {
    // mede só até o separador decimal
    const walker = document.createTreeWalker(range.commonAncestorContainer.nodeType === 1 ? range.commonAncestorContainer : p, NodeFilter.SHOW_TEXT);
    let n;
    while ((n = walker.nextNode())) {
      if (!range.intersectsNode(n)) continue;
      const idx = n.data.search(/[,.]/);
      if (idx >= 0) {
        const r2 = document.createRange();
        r2.setStart(n, 0);
        r2.setEnd(n, idx);
        const rects = r2.getClientRects();
        for (const r of rects) if (Math.abs(r.top + r.height / 2 - lineMid) < tr.height + 4) right = Math.max(right, r.right);
        found = true;
        break;
      }
    }
  }
  if (!found) {
    for (const r of range.getClientRects()) {
      if (r.width <= 0) continue;
      if (Math.abs(r.top + r.height / 2 - lineMid) < Math.max(8, tr.height)) right = Math.max(right, r.right);
    }
  }
  return Math.max(0, (right - tr.right) / z);
}

// topo, na tela, do caractere numa posição do DOM
const topRange = typeof document !== 'undefined' ? document.createRange() : null;
function domTop(node, offset) {
  if (!node) return null;
  if (node.nodeType === 3) {
    if (offset >= node.length) return null;
    topRange.setStart(node, offset);
    topRange.setEnd(node, offset + 1);
    const rects = topRange.getClientRects();
    return rects.length ? rects[0].top : null;
  }
  const child = node.childNodes[offset];
  if (!child) return null;
  if (child.nodeType === 3) return child.length ? domTop(child, 0) : null;
  if (child.nodeType === 1) {
    const r = child.getBoundingClientRect();
    return r.height || r.width ? r.top : null;
  }
  return null;
}

// ---------------------------------------------------------------- linhas de um parágrafo

function lineBoxes(p) {
  const rects = [];
  const widgets = [];
  const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT, {
    acceptNode(n) {
      if (n.nodeType === 3) return n.data.length ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_REJECT;
      const el = n;
      if (el.classList.contains('pb')) { widgets.push(el); return NodeFilter.FILTER_REJECT; }
      if (el.classList.contains('drw-wrap')) {
        const m = el.dataset.mode;
        if (m === 'inline' || m === 'block') { const r = el.getBoundingClientRect(); if (r.height > 0) rects.push({ top: r.top, bottom: r.bottom }); }
        return NodeFilter.FILTER_REJECT;
      }
      if (el.tagName === 'BR') { const r = el.getBoundingClientRect(); if (r.height > 0) rects.push({ top: r.top, bottom: r.bottom }); return NodeFilter.FILTER_REJECT; }
      if (el.classList.contains('fnref')) { const r = el.getBoundingClientRect(); if (r.height > 0) rects.push({ top: r.top, bottom: r.bottom }); return NodeFilter.FILTER_REJECT; }
      return NodeFilter.FILTER_SKIP;
    },
  });
  const range = document.createRange();
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    range.selectNodeContents(n);
    for (const r of range.getClientRects()) if (r.height > 0) rects.push({ top: r.top, bottom: r.bottom });
  }
  rects.sort((a, b) => a.top - b.top || a.bottom - b.bottom);
  const lines = [];
  for (const r of rects) {
    const cur = lines[lines.length - 1];
    if (!cur || r.top >= cur.bottom - 2) lines.push({ top: r.top, bottom: r.bottom, textTop: r.top });
    else { cur.top = Math.min(cur.top, r.top); cur.bottom = Math.max(cur.bottom, r.bottom); cur.textTop = Math.min(cur.textTop, r.top); }
  }
  const wrects = widgets.map((w) => { const r = w.getBoundingClientRect(); return { top: r.top, bottom: r.bottom, h: Number(w.dataset.h) || r.height }; });
  return { lines, widgets: wrects };
}

// ---------------------------------------------------------------- motor de páginas

export class PageLayout {
  constructor(els) {
    this.scrollEl = els.scroll;
    this.canvasEl = els.canvas;
    this.zoomEl = els.zoom;
    this.pagesEl = els.pages;
    this.hostEl = els.host;
    this.view = null;
    this.model = null;
    this.zoom = 1;
    this.gap = 20;
    this.padTop = 20;
    this.padBottom = 36;
    this.mode = 'print';
    this.pageCount = 1;
    this.pageFns = [[]];
    this.pageEls = [];
    this.pageKeys = [];
    this.passes = 0;
    this.frame = null;
    this.timer = null;
    this.lastCost = 0;
    this.listeners = new Set();
    this.staticCache = new Map();
    this.fnCache = new Map();
    this.fnSepH = 18;
    this.headerEditing = null;
    this.measureEl = document.createElement('div');
    this.measureEl.className = 'doc-body static-doc measure';
    this.zoomEl.appendChild(this.measureEl);
  }

  setModel(model) {
    this.model = model;
    this.staticCache.clear();
    this.fnCache.clear();
    this.pageKeys = [];
    this.computeGeometry();
  }

  attach(view) { this.view = view; this.schedule(true); }
  detach(view) { if (this.view === view) this.view = null; }
  onChange(fn) { this.listeners.add(fn); return () => this.listeners.delete(fn); }

  invalidateStatic() {
    this.staticCache.clear();
    this.fnCache.clear();
    this.pageKeys = [];
  }

  // Recalcula tudo na próxima passada (fontes, imagens, página, estilos)
  invalidate() {
    this.full = true;
    this.schedule(false);
  }

  schedule(fromChange) {
    if (fromChange) this.passes = 0;
    else this.full = true;
    if (this.frame) return;
    // documentos grandes esperam a digitação parar, com um limite de espera
    if (this.lastCost > 40 && fromChange) {
      const now = performance.now();
      if (!this.timer) this.firstWait = now;
      clearTimeout(this.timer);
      const delay = now - this.firstWait > 1500 ? 0 : Math.min(450, this.lastCost * 3);
      this.timer = setTimeout(() => { this.timer = null; this.frame = requestAnimationFrame(() => this.run()); }, delay);
      return;
    }
    if (this.timer) return;
    this.frame = requestAnimationFrame(() => this.run());
  }

  runNow(maxPasses = 6) {
    if (this.frame) { cancelAnimationFrame(this.frame); this.frame = null; }
    if (this.timer) { clearTimeout(this.timer); this.timer = null; }
    this.full = true;
    for (let i = 0; i < maxPasses; i++) if (!this.run(true)) break;
  }

  // Atalho: se só mudou texto e as alturas dos blocos ficaram iguais, as quebras continuam válidas
  quickCheck() {
    const view = this.view;
    const prev = this.lastBlocks;
    const doc = view.state.doc;
    if (!prev || prev.length !== doc.childCount) return false;
    const z = this.zoom || 1;
    const breaks = currentBreaks(view.state);
    let pos = 0;
    const touched = [];
    for (let i = 0; i < doc.childCount; i++) {
      const n = doc.child(i);
      const p = prev[i];
      if (n !== p.node) {
        if (n.type !== p.node.type || n.attrs !== p.node.attrs) return false;
        const end = pos + n.nodeSize;
        if (breaks.some((b) => b.pos > pos && b.pos < end)) return false;
        const dom = view.nodeDOM(pos);
        if (!dom || dom.nodeType !== 1) return false;
        if (dom.querySelector('sup.fnref')) return false;
        if (dom.querySelector('span.tab, .num-lbl')) layoutTabs(dom, z);
        const h = dom.getBoundingClientRect().height / z;
        if (Math.abs(h - p.h) > 0.5) return false;
        touched.push([i, { node: n, h }]);
      }
      pos += n.nodeSize;
    }
    for (const [i, v] of touched) prev[i] = v;
    return true;
  }

  // Devolve true se alterou as quebras
  run(sync) {
    this.frame = null;
    const view = this.view;
    if (!view || !this.model) return false;
    const t0 = performance.now();
    if (!this.full && this.mode === 'print' && this.quickCheck()) {
      this.lastQuick = performance.now() - t0;
      for (const fn of this.listeners) fn();
      return false;
    }
    this.full = false;
    this.computeGeometry();
    layoutTabs(view.dom, this.zoom);
    let changed = false;
    if (this.mode === 'print') {
      const breaks = this.paginate();
      if (!sameBreaks(breaks, currentBreaks(view.state))) {
        const tr = view.state.tr.setMeta(layoutKey, { breaks }).setMeta('addToHistory', false);
        view.dispatch(tr);
        changed = true;
        this.passes++;
        if (!sync && this.passes < 5) { this.full = true; this.frame = requestAnimationFrame(() => this.run()); }
      } else {
        this.passes = 0;
      }
    } else if (currentBreaks(view.state).length) {
      view.dispatch(view.state.tr.setMeta(layoutKey, { breaks: [] }).setMeta('addToHistory', false));
      this.pageCount = 1;
      changed = true;
    }
    this.renderPages();
    this.lastCost = performance.now() - t0;
    for (const fn of this.listeners) fn();
    return changed;
  }

  // ------------------------------------------------------------ geometria

  computeGeometry() {
    const s = this.model.section;
    const pageW = s.pgSz.w / 15;
    const pageH = s.pgSz.h / 15;
    const m = s.pgMar;
    const gutter = (m.gutter || 0) / 15;
    const mLeft = m.left / 15 + gutter;
    const mRight = m.right / 15;
    const g = {
      pageW, pageH, mLeft, mRight,
      mTop: Math.abs(m.top) / 15, mBottom: Math.abs(m.bottom) / 15,
      fixedTop: m.top < 0, fixedBottom: m.bottom < 0,
      headerDist: m.header / 15, footerDist: m.footer / 15,
      contentW: Math.max(40, pageW - mLeft - mRight),
      pitch: pageH + this.gap,
    };
    this.geom = g;
    if (this.mode === 'print') {
      this.zoomEl.style.width = pageW + 'px';
      this.hostEl.style.left = mLeft + 'px';
      this.hostEl.style.width = g.contentW + 'px';
      this.hostEl.style.paddingTop = this.cTop(0) + 'px';
      this.measureEl.style.width = g.contentW + 'px';
    }
    return g;
  }

  staticFor(entry, kind) {
    if (!entry) return null;
    const key = kind + ':' + entry.rId;
    let c = this.staticCache.get(key);
    if (c && c.doc === entry.doc) return c;
    const holder = renderStatic(entry.doc);
    holder.classList.add(kind === 'header' ? 'hdr-content' : 'ftr-content');
    this.measureEl.appendChild(holder);
    layoutTabs(holder, 1);
    const height = holder.getBoundingClientRect().height / (this.zoom || 1);
    holder.remove();
    c = { holder, height, doc: entry.doc };
    this.staticCache.set(key, c);
    return c;
  }

  headerInfo(kind, k) {
    if (!this.model) return null;
    if (this.headerEditing && this.headerEditing.kind === kind) {
      const entry = this.model.headerFor(kind, k, k === 0);
      if (entry && entry.rId === this.headerEditing.rId) return { height: this.headerEditing.height() };
    }
    return this.staticFor(this.model.headerFor(kind, k, k === 0), kind);
  }

  cTop(k) {
    const g = this.geom;
    if (g.fixedTop) return g.mTop;
    const h = this.headerInfo('header', k);
    return Math.max(g.mTop, h ? g.headerDist + h.height : 0);
  }

  cBottom(k) {
    const g = this.geom;
    if (g.fixedBottom) return g.mBottom;
    const f = this.headerInfo('footer', k);
    return Math.max(g.mBottom, f ? g.footerDist + f.height : 0);
  }

  footnoteFor(id) {
    const fn = this.model.footnotes.get(String(id));
    if (!fn) return null;
    let c = this.fnCache.get(String(id));
    if (c && c.doc === fn.doc) return c;
    const holder = renderStatic(fn.doc);
    holder.classList.add('fn-content');
    this.measureEl.appendChild(holder);
    const height = holder.getBoundingClientRect().height / (this.zoom || 1);
    holder.remove();
    c = { holder, height, doc: fn.doc };
    this.fnCache.set(String(id), c);
    return c;
  }

  // ------------------------------------------------------------ paginação

  paginate() {
    const view = this.view;
    const g = this.geom;
    const z = this.zoom || 1;
    const originY = this.zoomEl.getBoundingClientRect().top;
    const Y = (v) => (v - originY) / z;
    const old = currentBreaks(view.state);
    const newBreaks = [];
    const self = this;
    let page = 0;
    let shift = 0;
    let oldAbove = 0;
    let oi = 0;
    const fnUsed = [0];
    const pageFns = [[]];
    let pagePlaced = [];
    let forceNext = false;
    const cTopCache = [];
    const cBotCache = [];
    const cTop = (k) => (cTopCache[k] ??= self.cTop(k));
    const cBot = (k) => (cBotCache[k] ??= self.cBottom(k));
    const contentTop = (k) => k * g.pitch + cTop(k);
    const fnArea = (k, extra) => { const u = (fnUsed[k] || 0) + (extra || 0); return u > 0 ? u + self.fnSepH : 0; };
    const limit = (k, extra) => k * g.pitch + g.pageH - cBot(k) - fnArea(k, extra);
    const atPageTop = (y) => y <= contentTop(page) + 2;
    const newPage = () => { page++; fnUsed[page] = 0; pageFns[page] = []; pagePlaced = []; };
    const fnHeight = (ids) => ids.reduce((s, id) => { const f = self.footnoteFor(id); return s + (f ? f.height : 0); }, 0);

    const blocks = [];
    // os elementos de primeiro nível seguem a ordem dos nós, tirando os marcadores de quebra
    const childDoms = (el, count) => {
      if (!el) return null;
      const list = [];
      for (let c = el.firstElementChild; c; c = c.nextElementSibling) if (!c.classList.contains('pb')) list.push(c);
      return list.length === count ? list : null;
    };
    const collect = (parent, base, parentDom) => {
      const doms = childDoms(parentDom, parent.childCount);
      let i = 0;
      parent.forEach((node, offset) => {
        const pos = base + offset;
        const dom = doms ? doms[i] : view.nodeDOM(pos);
        i++;
        if (node.type.name === 'sdt_block') collect(node, pos + 1, dom);
        else blocks.push({ node, pos, dom });
      });
    };
    collect(view.state.doc, 0, view.dom);
    this.topDoms = childDoms(view.dom, view.state.doc.childCount);

    // usa o elemento já conhecido; procurar pela posição percorre o documento inteiro
    const padTopOf = (d) => {
      if (!d || d.nodeType !== 1 || d.tagName !== 'P') return 0;
      return parseFloat(getComputedStyle(d).paddingTop) || 0;
    };

    // quebra antes de um bloco, respeitando "manter com o próximo"
    const breakBefore = (pos, top, soft, dom) => {
      let target = { pos, top, dom };
      if (soft) {
        for (let i = pagePlaced.length - 1; i >= 0; i--) {
          const pb = pagePlaced[i];
          if (!pb.keepNext) break;
          if (i === 0 && atPageTop(pb.top)) { target = { pos, top, dom }; break; }
          target = pb;
        }
      }
      const sup = soft ? padTopOf(target.dom) : 0;
      const h = contentTop(page + 1) - target.top - sup;
      newBreaks.push({ pos: target.pos, kind: 'block', h: Math.max(0, h) });
      shift += Math.max(0, h);
      const moved = target.pos !== pos;
      newPage();
      return { h: Math.max(0, h), moved };
    };

    for (const b of blocks) {
      while (oi < old.length && old[oi].pos <= b.pos) { oldAbove += old[oi].h; oi++; }
      let innerOld = 0;
      let oj = oi;
      const end = b.pos + b.node.nodeSize;
      while (oj < old.length && old[oj].pos < end) { innerOld += old[oj].h; oj++; }
      const dom = b.dom;
      if (!dom || dom.nodeType !== 1) { oldAbove += innerOld; oi = oj; continue; }
      const rect = dom.getBoundingClientRect();
      let top = Y(rect.top) - oldAbove + shift;
      let bottom = Y(rect.bottom) - oldAbove - innerOld + shift;
      const type = b.node.type.name;
      const eff = type === 'paragraph' && this.model.styles ? this.model.styles.effectivePPr(b.node.attrs, this.model.numbering) : {};

      if (forceNext) {
        forceNext = false;
        const h = contentTop(page + 1) - top;
        newBreaks.push({ pos: b.pos, kind: 'block', h: Math.max(0, h) });
        shift += Math.max(0, h); top += h; bottom += h;
        newPage();
      } else if (eff.pageBreakBefore && !atPageTop(top)) {
        const { h } = breakBefore(b.pos, top, false);
        top += h; bottom += h;
      }

      if (type === 'paragraph') {
        const res = this.placeParagraph(b, dom, eff, top, bottom, {
          Y, oldAbove, get shift() { return shift; }, set shift(v) { shift = v; },
          get page() { return page; }, contentTop, limit, atPageTop, newPage, newBreaks, fnUsed, pageFns, fnHeight, breakBefore,
          pagePlaced: () => pagePlaced,
        });
        if (res.forceAfter) forceNext = true;
        pagePlaced.push({ pos: b.pos, top: res.top, keepNext: !!eff.keepNext, dom });
      } else if (type === 'table') {
        const res = this.placeTable(b, dom, top, bottom, {
          Y, oldAbove, get shift() { return shift; }, set shift(v) { shift = v; },
          get page() { return page; }, contentTop, limit, atPageTop, newPage, newBreaks, fnUsed, pageFns, fnHeight, breakBefore,
        });
        pagePlaced.push({ pos: b.pos, top: res.top, keepNext: false, dom });
      } else {
        if (bottom > limit(page) + EPS && !atPageTop(top)) {
          const { h } = breakBefore(b.pos, top, true, dom);
          top += h;
        }
        pagePlaced.push({ pos: b.pos, top, keepNext: false, dom });
      }
      oldAbove += innerOld;
      oi = oj;
    }
    if (forceNext) {
      // quebra no fim do documento: página em branco no final
      newBreaks.push({ pos: view.state.doc.content.size, kind: 'block', h: 0 });
      newPage();
    }
    this.pageCount = page + 1;
    this.pageFns = pageFns;
    this.fnUsed = fnUsed;
    this.cTopCache = cTopCache;
    this.cBotCache = cBotCache;
    // alturas naturais dos blocos, usadas pelo atalho da próxima digitação
    const lb = [];
    let p2 = 0;
    let k = 0;
    const tops = this.topDoms;
    view.state.doc.forEach((n, _off, idx) => {
      const end = p2 + n.nodeSize;
      while (k < old.length && old[k].pos <= p2) k++;
      let inner = 0;
      for (let j = k; j < old.length && old[j].pos < end; j++) inner += old[j].h;
      const dom = tops ? tops[idx] : view.nodeDOM(p2);
      const h = dom && dom.nodeType === 1 ? dom.getBoundingClientRect().height / z - inner : 0;
      lb.push({ node: n, h });
      p2 = end;
    });
    this.lastBlocks = lb;
    return newBreaks;
  }

  placeParagraph(b, dom, eff, top, bottom, S) {
    const view = this.view;
    const node = b.node;
    // atalho: cabe inteiro na página e não tem nota nem quebra manual, então não
    // precisa medir linha por linha
    let special = !!eff.sectPr;
    if (!special) node.forEach((ch) => { if (ch.type.name === 'footnote_ref' || (ch.type.name === 'hard_break' && ch.attrs.type === 'page')) special = true; });
    if (!special && bottom <= S.limit(S.page) + EPS) return { top, forceAfter: false };
    const { lines: rawLines, widgets } = lineBoxes(dom);
    const cs = getComputedStyle(dom);
    const z = this.zoom || 1;
    const pr = dom.getBoundingClientRect();
    const cTopScreen = pr.top + ((parseFloat(cs.paddingTop) || 0) + (parseFloat(cs.borderTopWidth) || 0)) * z;
    const cBotScreen = pr.bottom - ((parseFloat(cs.paddingBottom) || 0) + (parseFloat(cs.borderBottomWidth) || 0)) * z;
    // limites das linhas em coordenadas da tela
    const n = rawLines.length;
    const lines = [];
    for (let i = 0; i < n; i++) {
      const L = rawLines[i];
      let lt = i === 0 ? Math.min(cTopScreen, L.top) : (rawLines[i - 1].bottom + L.top) / 2;
      let lb = i === n - 1 ? Math.max(cBotScreen, L.bottom) : (L.bottom + rawLines[i + 1].top) / 2;
      for (const w of widgets) {
        if (i > 0 && w.top >= rawLines[i - 1].bottom - 1 && w.bottom <= L.top + 1) lt = Math.max(lt, w.bottom);
        if (i < n - 1 && w.top >= L.bottom - 1 && w.bottom <= rawLines[i + 1].top + 1) lb = Math.min(lb, w.top);
      }
      let above = 0;
      for (const w of widgets) if (w.bottom <= L.top + 1) above += w.h;
      lines.push({ top: S.Y(lt) - S.oldAbove - above, bottom: S.Y(lb) - S.oldAbove - above, textTop: L.textTop, fns: [] });
    }
    const shift0 = S.shift;
    for (const L of lines) { L.top += shift0; L.bottom += shift0; }
    if (!lines.length) lines.push({ top, bottom, textTop: pr.top, fns: [] });
    // referências de nota por linha
    if (this.model.footnotes.size) {
      for (const sup of dom.querySelectorAll('sup.fnref-footnote')) {
        let id = null;
        try { id = JSON.parse(sup.getAttribute('data-fn')).id; } catch { id = null; }
        if (id == null) continue;
        const r = sup.getBoundingClientRect();
        let best = 0;
        for (let i = 0; i < rawLines.length; i++) if (r.top + r.height / 2 >= rawLines[i].top - 2) best = i;
        if (lines[best]) lines[best].fns.push(id);
      }
    }
    // quebras de página manuais dentro do parágrafo
    const manual = [];
    let lastContent = -1;
    node.forEach((ch, off) => { if (ch.type.name !== 'opaque_inline' || ch.attrs.kind === 'del') lastContent = off; });
    let forceAfter = false;
    node.forEach((ch, off) => {
      if (ch.type.name === 'hard_break' && ch.attrs.type === 'page') {
        const pos = b.pos + 1 + off;
        if (off >= lastContent) { forceAfter = true; return; }
        let d = null;
        try { const cd = dom.pmViewDesc && dom.pmViewDesc.descAt(off); d = cd ? cd.nodeDOM || cd.dom : null; } catch { d = null; }
        if (!d) d = view.nodeDOM(pos);
        if (!d || d.nodeType !== 1) return;
        const r = d.getBoundingClientRect();
        let li = 0;
        for (let i = 0; i < rawLines.length; i++) if (r.top + Math.min(r.height, 10) / 2 >= rawLines[i].top - 2) li = i;
        manual.push({ line: li, pos: pos + ch.nodeSize });
      }
    });
    if (eff.sectPr) {
      const t = /<w:type\s+w:val="(\w+)"/.exec(eff.sectPr);
      if (!t || t[1] !== 'continuous') forceAfter = true;
    }
    const widow = eff.widowControl !== false;
    const keepLines = !!eff.keepLines;
    let f = 0;
    let lineShift = 0;
    let firstTop = lines[0].top;
    let startedAt = top;
    for (let k = 0; k < lines.length; k++) {
      const lt = lines[k].top + lineShift;
      const lb = lines[k].bottom + lineShift;
      const fnExtra = lines[k].fns.length ? S.fnHeight(lines[k].fns) : 0;
      if (lb > S.limit(S.page, fnExtra) + EPS) {
        let split = k;
        if (k === f) {
          if (k === 0 && !S.atPageTop(startedAt + lineShift)) {
            const { h } = S.breakBefore(b.pos, top + lineShift, true, dom);
            lineShift += h;
            startedAt = top;
            k = -1;
            continue;
          }
          // linha maior que a página: aceita
        } else {
          if (keepLines && f === 0) split = 0;
          else if (widow) {
            const onThis = split - f;
            const remaining = lines.length - split;
            if (f === 0 && onThis === 1 && lines.length > 1) split = 0;
            else if (remaining === 1 && lines.length > 2 && split - 1 > f) {
              split -= 1;
              if (f === 0 && split < 2) split = 0;
            }
          }
          if (split === 0) {
            if (!S.atPageTop(top + lineShift)) {
              const { h } = S.breakBefore(b.pos, top + lineShift, true, dom);
              lineShift += h;
              k = -1;
              continue;
            }
            split = k;
          }
          const pos = this.posAtLineStart(b, lines[split].textTop, split, dom);
          if (pos != null) {
            const h = S.contentTop(S.page + 1) - (lines[split].top + lineShift);
            S.newBreaks.push({ pos, kind: 'inline', h: Math.max(0, h) });
            lineShift += Math.max(0, h);
            S.shift += Math.max(0, h);
            S.newPage();
            f = split;
            k = split - 1;
            continue;
          }
        }
      }
      if (fnExtra) {
        S.fnUsed[S.page] += fnExtra;
        S.pageFns[S.page].push(...lines[k].fns);
      }
      const mb = manual.find((m) => m.line === k);
      if (mb && k < lines.length - 1) {
        const h = S.contentTop(S.page + 1) - lb;
        S.newBreaks.push({ pos: mb.pos, kind: 'inline', h: Math.max(0, h) });
        lineShift += Math.max(0, h);
        S.shift += Math.max(0, h);
        S.newPage();
        f = k + 1;
      }
    }
    return { top: top + (firstTop - lines[0].top), forceAfter };
  }

  // posição do primeiro caractere de uma linha. Mede pelo próprio parágrafo:
  // perguntar ao editor pela posição percorre o documento inteiro a cada chamada.
  posAtLineStart(b, textTopScreen, lineIndex, dom) {
    const view = this.view;
    let lo = b.pos + 1;
    let hi = b.pos + b.node.nodeSize - 1;
    if (hi <= lo) return null;
    const target = textTopScreen - 1;
    const desc = dom && dom.pmViewDesc;
    const topAt = (pos) => {
      if (desc && desc.domFromPos) {
        try {
          const { node, offset } = desc.domFromPos(pos - b.pos - 1, 1);
          const t = domTop(node, offset);
          if (t != null) return t;
        } catch { /* usa o caminho padrão */ }
      }
      try { const c = view.coordsAtPos(pos, 1); return c ? c.top : null; } catch { return null; }
    };
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      const top = topAt(mid);
      if (top == null) { lo = mid + 1; continue; }
      if (top >= target) hi = mid; else lo = mid + 1;
    }
    if (lo <= b.pos + 1 && lineIndex > 0) return null;
    return lo;
  }

  placeTable(b, dom, top, bottom, S) {
    const view = this.view;
    const fnIds = [];
    if (this.model.footnotes.size) {
      for (const sup of dom.querySelectorAll('sup.fnref-footnote')) {
        try { fnIds.push(JSON.parse(sup.getAttribute('data-fn')).id); } catch { /* ignora */ }
      }
    }
    const fnExtra = fnIds.length ? S.fnHeight(fnIds) : 0;
    if (bottom <= S.limit(S.page, fnExtra) + EPS) {
      if (fnExtra) { S.fnUsed[S.page] += fnExtra; S.pageFns[S.page].push(...fnIds); }
      return { top };
    }
    const tbody = dom.querySelector('tbody');
    if (!tbody) return { top };
    const rowsDom = Array.from(tbody.children).filter((r) => r.tagName === 'TR');
    const table = b.node;
    const offsets = [];
    table.forEach((row, off) => offsets.push(off));
    let colCount = 0;
    if (table.firstChild) table.firstChild.forEach((cell) => { colCount += cell.attrs.colspan || 1; });
    colCount = Math.max(1, colCount);
    let rowOld = 0;
    let ri = 0;
    let rowShift = 0;
    let firstOnPage = 0;
    const rows = [];
    for (const r of rowsDom) {
      if (r.classList.contains('pb')) { rowOld += Number(r.dataset.h) || r.getBoundingClientRect().height / (this.zoom || 1); continue; }
      const rect = r.getBoundingClientRect();
      rows.push({ top: S.Y(rect.top) - S.oldAbove - rowOld + S.shift, bottom: S.Y(rect.bottom) - S.oldAbove - rowOld + S.shift, index: ri++ });
    }
    for (let i = 0; i < rows.length; i++) {
      const rt = rows[i].top + rowShift;
      const rb = rows[i].bottom + rowShift;
      if (rb > S.limit(S.page) + EPS) {
        if (i === firstOnPage) {
          if (i === 0 && !S.atPageTop(top + rowShift)) {
            const { h } = S.breakBefore(b.pos, top + rowShift, true, dom);
            rowShift += h;
            i = -1;
            continue;
          }
          continue;
        }
        const pos = b.pos + 1 + offsets[i];
        const h = S.contentTop(S.page + 1) - rt;
        S.newBreaks.push({ pos, kind: 'row', h: Math.max(0, h), cols: colCount });
        rowShift += Math.max(0, h);
        S.shift += Math.max(0, h);
        S.newPage();
        firstOnPage = i;
        i -= 1;
      }
    }
    if (fnExtra) { S.fnUsed[S.page] += fnExtra; S.pageFns[S.page].push(...fnIds); }
    return { top };
  }

  // ------------------------------------------------------------ páginas na tela

  pageTop(k) {
    return k * this.geom.pitch;
  }

  // página de uma posição do documento
  pageOfPos(pos) {
    if (this.mode !== 'print') return 0;
    const view = this.view;
    if (!view) return 0;
    try {
      const c = view.coordsAtPos(Math.min(pos, view.state.doc.content.size));
      const y = (c.top - this.zoomEl.getBoundingClientRect().top) / (this.zoom || 1);
      return Math.max(0, Math.min(this.pageCount - 1, Math.floor(y / this.geom.pitch)));
    } catch { return 0; }
  }

  pageAtScroll() {
    if (this.mode !== 'print' || !this.geom) return 0;
    const rect = this.zoomEl.getBoundingClientRect();
    const sr = this.scrollEl.getBoundingClientRect();
    const y = (sr.top + sr.height / 3 - rect.top) / (this.zoom || 1);
    return Math.max(0, Math.min(this.pageCount - 1, Math.floor(y / this.geom.pitch)));
  }

  renderPages() {
    const g = this.geom;
    const model = this.model;
    if (!g || !model) return;
    const N = this.mode === 'print' ? this.pageCount : 0;
    while (this.pageEls.length < N) {
      const el = document.createElement('div');
      el.className = 'page';
      const hdr = document.createElement('div');
      hdr.className = 'page-header doc-body';
      const ftr = document.createElement('div');
      ftr.className = 'page-footer doc-body';
      const notes = document.createElement('div');
      notes.className = 'page-notes doc-body';
      el.append(hdr, ftr, notes);
      el.addEventListener('dblclick', (e) => this.onPageDblClick(e, el));
      this.pagesEl.appendChild(el);
      this.pageEls.push(el);
    }
    while (this.pageEls.length > N) this.pageEls.pop().remove();
    const pgStart = model.section.pgNumType.start;
    const fmt = model.section.pgNumType.fmt || 'decimal';
    for (let k = 0; k < N; k++) {
      const el = this.pageEls[k];
      el.style.top = k * g.pitch + 'px';
      el.style.width = g.pageW + 'px';
      el.style.height = g.pageH + 'px';
      el.dataset.page = String(k);
      const hEntry = model.headerFor('header', k, k === 0);
      const fEntry = model.headerFor('footer', k, k === 0);
      const fns = (this.pageFns && this.pageFns[k]) || [];
      const pageNum = (pgStart != null ? pgStart : 1) + k;
      const key = [g.pageW, g.pageH, g.mLeft, g.contentW, g.headerDist, g.footerDist, hEntry && hEntry.rId, hEntry && hEntry.doc && hEntry.doc.content.size, fEntry && fEntry.rId,
        fEntry && fEntry.doc && fEntry.doc.content.size, pageNum, N, fmt, fns.join(','), this.cBottom(k), this.staticCache.size, this.headerEditing ? this.headerEditing.kind + this.headerEditing.page : ''].join('|');
      if (this.pageKeys[k] === key) continue;
      this.pageKeys[k] = key;
      const [hdr, ftr, notes] = el.children;
      const editingHere = this.headerEditing && this.headerEditing.page === k;
      if (!(editingHere && this.headerEditing.kind === 'header')) this.fillArea(hdr, hEntry, 'header', { page: pageNum, pages: N, fmt });
      if (!(editingHere && this.headerEditing.kind === 'footer')) this.fillArea(ftr, fEntry, 'footer', { page: pageNum, pages: N, fmt });
      hdr.style.cssText = 'left:' + g.mLeft + 'px;width:' + g.contentW + 'px;top:' + g.headerDist + 'px';
      ftr.style.cssText = 'left:' + g.mLeft + 'px;width:' + g.contentW + 'px;bottom:' + g.footerDist + 'px';
      notes.textContent = '';
      if (fns.length) {
        const area = (this.fnUsed[k] || 0) + this.fnSepH;
        notes.style.cssText = 'left:' + g.mLeft + 'px;width:' + g.contentW + 'px;top:' + (g.pageH - this.cBottom(k) - area) + 'px';
        const sep = document.createElement('div');
        sep.className = 'fn-sep';
        notes.appendChild(sep);
        for (const id of fns) {
          const f = this.footnoteFor(id);
          if (!f) continue;
          const clone = f.holder.cloneNode(true);
          const num = this.footnoteNumber(id);
          for (const mark of clone.querySelectorAll('.oi-noteMark')) mark.textContent = num;
          notes.appendChild(clone);
        }
      } else notes.style.cssText = 'display:none';
    }
    const totalH = this.mode === 'print' ? N * g.pitch - this.gap : 0;
    this.totalH = totalH;
    this.updateCanvas();
  }

  footnoteNumber(id) {
    if (!this.fnOrder || this.fnOrderDoc !== this.view.state.doc) {
      this.fnOrder = new Map();
      let n = 0;
      this.view.state.doc.descendants((node) => {
        if (node.type.name === 'footnote_ref' && node.attrs.kind === 'footnote' && !node.attrs.custom) this.fnOrder.set(String(node.attrs.id), ++n);
        return true;
      });
      this.fnOrderDoc = this.view.state.doc;
    }
    return String(this.fnOrder.get(String(id)) || '');
  }

  fillArea(container, entry, kind, vars) {
    container.textContent = '';
    if (!entry) { container.classList.add('empty'); return; }
    container.classList.remove('empty');
    const st = this.staticFor(entry, kind);
    if (!st) return;
    const clone = st.holder.cloneNode(true);
    fillFields(clone, vars);
    container.appendChild(clone);
    layoutTabs(clone, 1);
  }

  onPageDblClick(e, pageEl) {
    if (!this.onHeaderDblClick) return;
    const k = Number(pageEl.dataset.page);
    const rect = pageEl.getBoundingClientRect();
    const y = (e.clientY - rect.top) / (this.zoom || 1);
    const g = this.geom;
    if (y < this.cTop(k)) this.onHeaderDblClick('header', k);
    else if (y > g.pageH - this.cBottom(k)) this.onHeaderDblClick('footer', k);
  }

  updateCanvas() {
    const g = this.geom;
    const z = this.zoom;
    if (!g) return;
    if (this.mode === 'print') {
      const h = (this.totalH + this.padTop + this.padBottom) * z;
      const w = g.pageW * z + 40;
      this.canvasEl.style.height = h + 'px';
      this.canvasEl.style.minWidth = w + 'px';
      this.zoomEl.style.transform = 'scale(' + z + ')';
      const avail = this.scrollEl.clientWidth;
      const left = Math.max(20, (Math.max(avail, w) - g.pageW * z) / 2);
      this.zoomEl.style.left = left + 'px';
      this.zoomEl.style.top = this.padTop * z + 'px';
    }
  }

  setZoom(z) {
    const old = this.zoom;
    z = Math.max(0.1, Math.min(5, z));
    if (Math.abs(old - z) < 0.001) return;
    const se = this.scrollEl;
    const cy = se.scrollTop + se.clientHeight / 2;
    const cx = se.scrollLeft + se.clientWidth / 2;
    this.zoom = z;
    if (this.view) this.view.redator = this.view.redator || {};
    this.updateCanvas();
    se.scrollTop = (cy / old) * z - se.clientHeight / 2;
    se.scrollLeft = (cx / old) * z - se.clientWidth / 2;
    for (const fn of this.listeners) fn();
  }
}
