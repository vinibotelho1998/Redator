// Régua horizontal com recuos arrastáveis, em centímetros.
import { h } from './dom.js';
import { TWIP_PER_CM, fmtNumber } from '../util/units.js';

export class Ruler {
  constructor(app) {
    this.app = app;
    this.canvas = h('canvas', { class: 'ruler-canvas' });
    this.first = h('div', { class: 'rm rm-first', title: 'Recuo da Primeira Linha' });
    this.hang = h('div', { class: 'rm rm-hang', title: 'Recuo Deslocado' });
    this.left = h('div', { class: 'rm rm-left', title: 'Recuo à Esquerda' });
    this.right = h('div', { class: 'rm rm-right', title: 'Recuo à Direita' });
    this.tabsLayer = h('div', { class: 'ruler-tabs' });
    this.guide = h('div', { class: 'ruler-guide', hidden: true });
    this.el = h('div', { class: 'ruler' }, this.canvas, this.tabsLayer, this.first, this.hang, this.left, this.right);
    this.state = null;
    for (const [el, kind] of [[this.first, 'first'], [this.hang, 'hang'], [this.left, 'left'], [this.right, 'right']]) {
      el.addEventListener('mousedown', (e) => this.drag(e, kind));
    }
    this.el.addEventListener('dblclick', () => app.dialog('paragraph'));
  }

  geometry() {
    const app = this.app;
    const layout = app.layout;
    if (!layout.geom) return null;
    const zr = layout.zoomEl.getBoundingClientRect();
    const rr = this.el.getBoundingClientRect();
    const g = layout.geom;
    const z = layout.zoom;
    const pageLeft = zr.left - rr.left;
    return { z, pageLeft, pageW: g.pageW * z, mLeft: g.mLeft * z, mRight: g.mRight * z, contentW: g.contentW * z, g };
  }

  // x em pixels da régua para uma posição em twips a partir da margem esquerda
  xOf(tw, geo) { return geo.pageLeft + geo.mLeft + (tw / 15) * geo.z; }
  twOf(x, geo) { return ((x - geo.pageLeft - geo.mLeft) / geo.z) * 15; }

  update(para) {
    if (para) this.state = para;
    const geo = this.geometry();
    if (!geo || this.el.hidden) return;
    const dpr = window.devicePixelRatio || 1;
    const w = this.el.clientWidth;
    const hgt = this.el.clientHeight;
    if (this.canvas.width !== Math.round(w * dpr) || this.canvas.height !== Math.round(hgt * dpr)) {
      this.canvas.width = Math.round(w * dpr);
      this.canvas.height = Math.round(hgt * dpr);
      this.canvas.style.width = w + 'px';
      this.canvas.style.height = hgt + 'px';
    }
    const ctx = this.canvas.getContext('2d');
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.clearRect(0, 0, w, hgt);
    const top = 3, bandH = hgt - 7;
    const x0 = geo.pageLeft, x1 = geo.pageLeft + geo.pageW;
    ctx.fillStyle = '#d6d6d6';
    ctx.fillRect(x0, top, x1 - x0, bandH);
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(x0 + geo.mLeft, top, geo.contentW, bandH);
    ctx.strokeStyle = '#b8b8b8';
    ctx.strokeRect(x0 + 0.5, top + 0.5, x1 - x0 - 1, bandH - 1);
    const unit = this.app.prefs.units || 'cm';
    const unitTw = unit === 'in' ? 1440 : unit === 'pt' ? 20 : TWIP_PER_CM;
    const unitPx = (unitTw / 15) * geo.z;
    this.el.setAttribute('aria-label', 'Régua em ' + (unit === 'in' ? 'polegadas' : unit === 'pt' ? 'pontos' : 'centímetros'));
    const origin = x0 + geo.mLeft;
    ctx.fillStyle = '#555';
    ctx.strokeStyle = '#777';
    ctx.font = '10px Segoe UI, sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    const mid = top + bandH / 2;
    const steps = unit === 'pt' ? [1, 6, 12, 18, 36, 72, 144, 288] : unit === 'in' ? [0.125, 0.25, 0.5, 1, 2, 4, 8] : [0.25, 0.5, 1, 2, 5, 10, 20];
    const step = steps.find((v) => v * unitPx >= 7) || steps[steps.length - 1];
    const labelEvery = steps.find((v) => v >= step && v * unitPx >= 30 && Math.abs(v / step - Math.round(v / step)) < 1e-6) || step * 4;
    const from = Math.floor(-geo.mLeft / (unitPx * step)) - 1;
    const to = Math.ceil((geo.pageW - geo.mLeft) / (unitPx * step)) + 1;
    for (let tick = from; tick <= to; tick++) {
      const c = tick * step;
      const x = origin + c * unitPx;
      if (x < x0 || x > x1) continue;
      const whole = Math.abs(c / labelEvery - Math.round(c / labelEvery)) < 1e-6;
      if (whole) {
        const n = Math.abs(c);
        if (n === 0) continue;
        ctx.fillText(fmtNumber(n, 2), x, mid + 0.5);
      } else if (Math.abs(c * 2 - Math.round(c * 2)) < 1e-6) {
        ctx.beginPath(); ctx.moveTo(x + 0.5, mid - 2.5); ctx.lineTo(x + 0.5, mid + 2.5); ctx.stroke();
      } else {
        ctx.beginPath(); ctx.moveTo(x + 0.5, mid - 1); ctx.lineTo(x + 0.5, mid + 1); ctx.stroke();
      }
    }
    // marcadores
    const p = this.state || { left: 0, right: 0, firstLine: 0, hanging: 0, tabs: [] };
    const firstTw = p.left - (p.hanging || 0) + (p.hanging ? 0 : p.firstLine || 0);
    this.place(this.first, this.xOf(firstTw, geo));
    this.place(this.hang, this.xOf(p.left, geo));
    this.place(this.left, this.xOf(p.left, geo));
    this.place(this.right, geo.pageLeft + geo.mLeft + geo.contentW - (p.right / 15) * geo.z);
    this.tabsLayer.textContent = '';
    for (const t of p.tabs || []) {
      if (t.val === 'clear' || t.val === 'bar') continue;
      const m = h('div', { class: 'rt rt-' + t.val, title: 'Tabulação ' + ({ left: 'Esquerda', right: 'Direita', center: 'Centralizada', decimal: 'Decimal' }[t.val] || '') });
      m.style.left = this.xOf(Number(t.pos), geo) + 'px';
      this.tabsLayer.appendChild(m);
    }
  }

  place(el, x) { el.style.left = Math.round(x) + 'px'; }

  drag(e, kind) {
    e.preventDefault();
    const geo = this.geometry();
    if (!geo || !this.state) return;
    const p = { ...this.state };
    const startX = e.clientX;
    const unit = this.app.prefs.units;
    const interval = unit === 'in' ? 180 : unit === 'pt' ? 120 : TWIP_PER_CM / 4;
    const snap = (tw) => (e.altKey ? tw : Math.round(tw / interval) * interval);
    const firstTw = p.left - (p.hanging || 0) + (p.hanging ? 0 : p.firstLine || 0);
    const guide = this.app.showGuide();
    let result = null;
    const move = (ev) => {
      const dxTw = ((ev.clientX - startX) / geo.z) * 15;
      let left = p.left, first = firstTw, right = p.right;
      if (kind === 'first') first = snap(firstTw + dxTw);
      else if (kind === 'hang') left = snap(p.left + dxTw);
      else if (kind === 'left') { left = snap(p.left + dxTw); first = firstTw + (left - p.left); }
      else if (kind === 'right') right = snap(p.right - dxTw);
      const cur = { ...p, left, right, firstLine: first > left ? first - left : 0, hanging: first < left ? left - first : 0 };
      this.state = cur;
      this.update();
      const x = kind === 'right' ? this.right.getBoundingClientRect().left + 5 : kind === 'first' ? this.first.getBoundingClientRect().left + 5 : this.hang.getBoundingClientRect().left + 5;
      guide(x);
      result = cur;
    };
    const up = () => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      guide(null);
      if (!result) return;
      const vals = { left: Math.round(result.left), right: Math.round(result.right) };
      if (result.hanging > 0) vals.hanging = Math.round(result.hanging);
      else vals.firstLine = Math.round(result.firstLine);
      this.app.cmd('indentValues', vals);
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  }
}
