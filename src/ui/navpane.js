// Painel de Navegação: títulos do documento e resultados da pesquisa.
import { h, clear, debounce } from './dom.js';
import { icon } from './icons.js';

export class NavPane {
  constructor(app) {
    this.app = app;
    this.tab = 'headings';
    this.input = h('input', { class: 'nav-search', type: 'search', placeholder: 'Pesquisar no documento', 'aria-label': 'Pesquisar no documento', spellcheck: 'false' });
    this.count = h('div', { class: 'nav-count', role: 'status', 'aria-live': 'polite' });
    this.prevBtn = h('button', { class: 'nav-arrow', type: 'button', title: 'Resultado anterior', html: '<svg viewBox="0 0 10 10" width="9" height="9"><path d="M1.5 6.5L5 3l3.5 3.5" fill="none" stroke="currentColor" stroke-width="1.3"/></svg>', onclick: () => this.step(-1) });
    this.nextBtn = h('button', { class: 'nav-arrow', type: 'button', title: 'Próximo resultado', html: icon('chevronDown'), onclick: () => this.step(1) });
    this.tabs = {
      headings: h('button', { class: 'nav-tab active', type: 'button', text: 'Títulos', onclick: () => this.setTab('headings') }),
      results: h('button', { class: 'nav-tab', type: 'button', text: 'Resultados', onclick: () => this.setTab('results') }),
    };
    this.list = h('div', { class: 'nav-list', role: 'region', 'aria-label': 'Títulos e resultados' });
    this.el = h('aside', { class: 'navpane', hidden: true },
      h('div', { class: 'nav-head' }, h('span', { class: 'nav-title', text: 'Navegação' }), h('button', { class: 'nav-close', type: 'button', title: 'Fechar', html: icon('close'), onclick: () => app.hideNavPane() })),
      h('div', { class: 'nav-search-row' }, this.input, h('span', { class: 'nav-search-icon', html: icon('find') })),
      h('div', { class: 'nav-tabs-row' }, h('div', { class: 'nav-tabs' }, this.tabs.headings, this.tabs.results), h('div', { class: 'nav-arrows' }, this.prevBtn, this.nextBtn)),
      this.count, this.list);
    const run = debounce(() => this.search(), 250);
    this.input.addEventListener('input', run);
    this.input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); if (this.app.searchCount) this.step(e.shiftKey ? -1 : 1); else this.search(); }
      if (e.key === 'Escape') { e.preventDefault(); this.app.focusEditor(); }
    });
  }

  get visible() { return !this.el.hidden; }

  setTab(t) {
    this.tab = t;
    for (const [k, b] of Object.entries(this.tabs)) b.classList.toggle('active', k === t);
    this.render();
  }

  search() {
    const q = this.input.value;
    const n = this.app.search(q, this.app.searchOpts);
    this.count.textContent = q ? (n ? n + (n === 1 ? ' resultado' : ' resultados') : 'Nenhum resultado') : '';
    if (q) this.setTab('results');
    else this.setTab('headings');
    if (n) this.app.findNext(1, true);
  }

  step(dir) {
    if (!this.input.value) return;
    this.app.findNext(dir);
  }

  render() {
    if (this.el.hidden) return;
    clear(this.list);
    if (this.tab === 'headings') {
      const heads = this.app.headingList();
      if (!heads.length) {
        this.list.appendChild(h('div', { class: 'nav-empty' }, h('p', { text: 'Crie uma estrutura interativa do documento.' }), h('p', { class: 'muted', text: 'É uma ótima maneira de controlar onde você está ou de mover o conteúdo rapidamente. Para começar, vá para a guia Página Inicial e aplique estilos de Título aos cabeçalhos do documento.' })));
        return;
      }
      const cur = this.app.currentHeadingIndex();
      heads.forEach((hd, i) => {
        const item = h('button', { type: 'button', 'aria-current': i === cur ? 'location' : null, class: 'nav-item' + (i === cur ? ' active' : ''), style: { paddingLeft: 10 + (hd.level - 1) * 14 + 'px' }, title: hd.text, text: hd.text || '(sem texto)' });
        item.addEventListener('click', () => this.app.gotoPos(hd.pos));
        this.list.appendChild(item);
      });
    } else {
      const results = this.app.searchResults();
      const q = this.input.value;
      // o documento pode ter mudado desde a busca: a contagem acompanha os resultados atuais
      this.count.textContent = q ? (results.length ? results.length + (results.length === 1 ? ' resultado' : ' resultados') : 'Nenhum resultado') : '';
      this.app.searchCount = results.length;
      if (!results.length) {
        this.list.appendChild(h('div', { class: 'nav-empty' }, h('p', { class: 'muted', text: this.input.value ? 'Nenhum resultado encontrado.' : 'Digite o texto a pesquisar.' })));
        return;
      }
      results.slice(0, 500).forEach((r, i) => {
        const ctx = r.context || '';
        const start = Math.max(0, ctx.indexOf(r.text) - 30);
        const item = h('button', { type: 'button', class: 'nav-result' });
        const before = ctx.slice(start, ctx.indexOf(r.text));
        item.append(h('span', { text: (start > 0 ? '…' : '') + before.replace(/\s+/g, ' ') }), h('b', { text: r.text }), h('span', { text: ctx.slice(ctx.indexOf(r.text) + r.text.length).slice(0, 60).replace(/\s+/g, ' ') }));
        item.addEventListener('click', () => this.app.gotoMatch(i));
        this.list.appendChild(item);
      });
    }
  }
}
