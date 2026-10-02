// Barra de status: páginas, palavras, idioma, modos de exibição e zoom.
import { h } from './dom.js';
import { icon } from './icons.js';
import { tooltip } from './popup.js';

const ZOOM_MIN = 0.1;
const ZOOM_MAX = 5;

// posição do controle deslizante: 100% no meio, como no Word
function zoomToSlider(z) {
  return z <= 1 ? ((z - ZOOM_MIN) / (1 - ZOOM_MIN)) * 50 : 50 + ((z - 1) / (ZOOM_MAX - 1)) * 50;
}
function sliderToZoom(v) {
  return v <= 50 ? ZOOM_MIN + (v / 50) * (1 - ZOOM_MIN) : 1 + ((v - 50) / 50) * (ZOOM_MAX - 1);
}

export class StatusBar {
  constructor(app) {
    this.app = app;
    this.page = h('button', { class: 'sb-item', type: 'button', onclick: () => app.dialog('find', 2) });
    this.words = h('button', { class: 'sb-item', type: 'button', onclick: () => app.dialog('wordCount') });
    this.lang = h('button', { class: 'sb-item', type: 'button', text: 'Português (Brasil)', onclick: (e) => app.openLanguageMenu(e.currentTarget) });
    this.info = h('span', { class: 'sb-item sb-info', role: 'status', 'aria-live': 'polite', 'aria-atomic': 'true' });
    tooltip(this.page, 'Número da página no documento', null, 'Clique para ir a uma página.');
    tooltip(this.words, 'Contar Palavras', null, 'Clique para ver as estatísticas do documento.');
    const mode = (id, ic, label) => {
      const b = h('button', { class: 'sb-view', type: 'button', html: icon(ic), onclick: () => app.setViewMode(id) });
      tooltip(b, label);
      return b;
    };
    this.modes = { read: mode('read', 'readMode16', 'Modo de Leitura'), print: mode('print', 'printLayout16', 'Layout de Impressão'), web: mode('web', 'webLayout16', 'Layout da Web') };
    this.focus = h('button', { class: 'sb-view', type: 'button', html: icon('focus16'), onclick: () => app.toggleFocus() });
    tooltip(this.focus, 'Foco', null, 'Esconda a faixa de opções e mantenha o foco no texto.');
    this.slider = h('input', { type: 'range', class: 'sb-slider', min: '0', max: '100', step: '0.5', 'aria-label': 'Zoom do documento' });
    this.slider.addEventListener('input', () => {
      let z = sliderToZoom(Number(this.slider.value));
      if (Math.abs(z - 1) < 0.04) z = 1;
      app.setZoom(Math.round(z * 100) / 100);
    });
    const minus = h('button', { class: 'sb-zbtn', type: 'button', html: icon('minus'), onclick: () => app.stepZoom(-1) });
    const plus = h('button', { class: 'sb-zbtn', type: 'button', html: icon('plus'), onclick: () => app.stepZoom(1) });
    tooltip(minus, 'Reduzir');
    tooltip(plus, 'Ampliar');
    this.pct = h('button', { class: 'sb-item sb-pct', type: 'button', onclick: () => app.dialog('zoom') });
    tooltip(this.pct, 'Nível de zoom', null, 'Clique para abrir a caixa de diálogo Zoom.');
    this.el = h('footer', { class: 'statusbar' },
      h('div', { class: 'sb-left' }, this.page, this.words, this.lang, this.info),
      h('div', { class: 'sb-right' }, this.focus, this.modes.read, this.modes.print, this.modes.web, h('span', { class: 'sb-gap' }), minus, this.slider, plus, this.pct));
  }

  update(st) {
    this.page.textContent = 'Página ' + (st.page + 1) + ' de ' + st.pages;
    this.words.textContent = st.selWords != null ? st.selWords.toLocaleString('pt-BR') + ' de ' + st.words.toLocaleString('pt-BR') + ' palavras' : st.words.toLocaleString('pt-BR') + (st.words === 1 ? ' palavra' : ' palavras');
    for (const [k, b] of Object.entries(this.modes)) { b.classList.toggle('active', k === st.viewMode); b.setAttribute('aria-pressed', String(k === st.viewMode)); }
    this.slider.value = String(zoomToSlider(st.zoom));
    this.slider.setAttribute('aria-valuetext', Math.round(st.zoom * 100) + '%');
    this.pct.textContent = Math.round(st.zoom * 100) + '%';
    this.lang.textContent = st.language || 'Português (Brasil)';
    this.info.textContent = st.info || '';
  }
}
