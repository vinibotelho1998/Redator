// Faixa de opções no estilo do Word.
import { h, clear } from './dom.js';
import { icon } from './icons.js';
import { openMenu, openPanel, tooltip, closeAllPopups, Popup } from './popup.js';
import { colorPanel, highlightPanel } from './colorpicker.js';
import { HIGHLIGHT } from '../util/color.js';
import { BULLET_LIBRARY, NUMBER_LIBRARY } from '../docx/numbering.js';
import { PAPER_SIZES, MARGIN_PRESETS } from '../model.js';
import { pClass, cClass, tClass } from '../editor/render.js';
import { fmtNumber, twipToCm, cmToTwip } from '../util/units.js';
import { spinner } from './dialogs.js';

const SYMBOLS = ['€', '£', '¥', '©', '®', '™', '±', '≠', '≤', '≥', '÷', '×', '∞', 'µ', 'α', 'β', 'π', 'Ω', '∑', '√', '°', '§', '¶', '•', '…', '–', '—', '“', '”', '«', '»', '←', '→', '↑', '↓', '✓', '✗', '★', '☐', '☑', '½', '¼', '¾', 'º', 'ª'];

export class Ribbon {
  constructor(app) {
    this.app = app;
    this.controls = new Map();
    this.tabsEl = h('div', { class: 'tabs', role: 'tablist' });
    this.bodyEl = h('div', { class: 'ribbon', role: 'toolbar' });
    this.el = h('div', { class: 'ribbon-wrap' }, this.tabsEl, this.bodyEl);
    this.tabs = [];
    this.active = 'home';
    this.collapsed = false;
    this.build();
  }

  // ------------------------------------------------------------ construtores

  reg(id, el, extra = {}) {
    if (id) this.controls.set(id, { el, ...extra });
    return el;
  }

  big(id, label, ic, run, opts = {}) {
    const btn = h('button', { class: 'rb-big' + (opts.menu ? ' has-menu' : '') + (opts.split ? ' split' : ''), type: 'button' });
    const top = h('span', { class: 'rb-big-icon', html: icon(ic) });
    const lbl = h('span', { class: 'rb-big-label' }, h('span', { text: label }), opts.menu ? h('span', { class: 'rb-caret', html: icon('chevronDown') }) : null);
    if (opts.split) {
      const main = h('span', { class: 'rb-split-main' }, top);
      const arrow = h('span', { class: 'rb-split-arrow' }, lbl);
      btn.append(main, arrow);
      main.addEventListener('click', (e) => { e.stopPropagation(); if (run) run(); });
      arrow.addEventListener('click', (e) => { e.stopPropagation(); this.openMenuFor(btn, opts.menu); });
    } else {
      btn.append(top, lbl);
      btn.addEventListener('click', () => { if (opts.menu) this.openMenuFor(btn, opts.menu); else if (run) run(); });
    }
    btn.addEventListener('mousedown', (e) => e.preventDefault());
    tooltip(btn, opts.tip || label, opts.shortcut, opts.desc);
    return this.reg(id, btn, { toggle: opts.toggle });
  }

  small(id, ic, run, opts = {}) {
    const btn = h('button', { class: 'rb-small' + (opts.label ? ' with-label' : '') + (opts.menu ? ' has-menu' : '') + (opts.split ? ' split' : ''), type: 'button' });
    const iconEl = h('span', { class: 'rb-icon', html: icon(ic) });
    if (opts.colorBar) {
      const bar = h('span', { class: 'rb-colorbar', style: { background: opts.colorBar } });
      iconEl.appendChild(bar);
    }
    const labelEl = opts.label ? h('span', { class: 'rb-label', text: opts.label }) : null;
    if (opts.split) {
      const main = h('span', { class: 'rb-split-main' }, iconEl, labelEl);
      const arrow = h('span', { class: 'rb-split-arrow', html: icon('chevronDown') });
      btn.append(main, arrow);
      main.addEventListener('click', (e) => { e.stopPropagation(); if (run) run(); });
      arrow.addEventListener('click', (e) => { e.stopPropagation(); this.openMenuFor(btn, opts.menu); });
    } else {
      btn.append(iconEl);
      if (labelEl) btn.append(labelEl);
      if (opts.menu) btn.append(h('span', { class: 'rb-caret', html: icon('chevronDown') }));
      btn.addEventListener('click', () => { if (opts.menu) this.openMenuFor(btn, opts.menu); else if (run) run(); });
    }
    btn.addEventListener('mousedown', (e) => e.preventDefault());
    tooltip(btn, opts.tip || opts.label || '', opts.shortcut, opts.desc);
    return this.reg(id, btn, { toggle: opts.toggle });
  }

  openMenuFor(btn, menu) {
    const m = typeof menu === 'function' ? menu() : menu;
    if (m instanceof Node) openPanel(btn, m);
    else openMenu(btn, m);
  }

  group(label, content, launcher) {
    const g = h('div', { class: 'rb-group' },
      h('div', { class: 'rb-group-body' }, ...content),
      h('div', { class: 'rb-group-label' }, h('span', { text: label }),
        launcher ? h('button', { class: 'rb-launcher', type: 'button', title: 'Mais opções de ' + label, onclick: launcher, onmousedown: (e) => e.preventDefault() }) : null));
    return g;
  }

  col(...items) { return h('div', { class: 'rb-col' }, ...items); }
  row(...items) { return h('div', { class: 'rb-row' }, ...items); }
  sep() { return h('span', { class: 'rb-sep' }); }

  check(id, label, run) {
    const input = h('input', { type: 'checkbox' });
    const el = h('label', { class: 'rb-check' }, input, h('span', { text: label }));
    input.addEventListener('change', () => run(input.checked));
    el.addEventListener('mousedown', (e) => { if (e.target !== input) e.preventDefault(); });
    return this.reg(id, el, { input });
  }

  // ------------------------------------------------------------ combos de fonte

  fontCombo() {
    const app = this.app;
    const input = h('input', { class: 'rb-combo-input font', type: 'text', spellcheck: 'false', 'aria-label': 'Fonte' });
    const arrow = h('button', { class: 'rb-combo-arrow', type: 'button', tabindex: '-1', html: icon('chevronDown') });
    const box = h('div', { class: 'rb-combo font-combo' }, input, arrow);
    const apply = () => { const v = input.value.trim(); if (v) app.cmd('setFont', v); app.focusEditor(); };
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); closeAllPopups(); apply(); }
      if (e.key === 'Escape') { e.preventDefault(); app.focusEditor(); }
    });
    input.addEventListener('focus', () => input.select());
    const openList = () => openPanel(box, this.fontList(input.value, (f) => { input.value = f; apply(); }), { className: 'font-popup' });
    arrow.addEventListener('mousedown', (e) => e.preventDefault());
    arrow.addEventListener('click', openList);
    input.addEventListener('input', () => {
      const q = input.value.toLowerCase();
      const all = app.fontList();
      const found = all.find((f) => f.toLowerCase().startsWith(q));
      if (found && q && input.selectionStart === input.value.length) {
        const typed = input.value.length;
        input.value = found;
        input.setSelectionRange(typed, found.length);
      }
    });
    tooltip(box, 'Fonte', 'Ctrl+Shift+F', 'Escolha uma nova fonte para o texto.');
    return this.reg('font', box, { input });
  }

  fontList(current, onPick) {
    const app = this.app;
    const wrap = h('div', { class: 'font-list' });
    const theme = app.model.theme;
    const add = (name, suffix) => wrap.appendChild(h('div', {
      class: 'font-item' + (name === current ? ' sel' : ''), style: { fontFamily: '"' + name + '", Segoe UI' },
      onclick: () => { closeAllPopups(); onPick(name); },
    }, h('span', { text: name }), suffix ? h('span', { class: 'font-suffix', text: suffix }) : null));
    wrap.appendChild(h('div', { class: 'menu-header', text: 'Fontes do Tema' }));
    add(theme.major, '(Títulos)');
    add(theme.minor, '(Corpo)');
    const recent = app.recentFonts();
    if (recent.length) {
      wrap.appendChild(h('div', { class: 'menu-header', text: 'Fontes Usadas Recentemente' }));
      for (const f of recent) add(f);
    }
    wrap.appendChild(h('div', { class: 'menu-header', text: 'Todas as Fontes' }));
    for (const f of app.fontList()) add(f);
    setTimeout(() => { const s = wrap.querySelector('.font-item.sel:last-of-type') || wrap.querySelector('.font-item.sel'); if (s) s.scrollIntoView({ block: 'center' }); }, 0);
    return wrap;
  }

  sizeCombo() {
    const app = this.app;
    const input = h('input', { class: 'rb-combo-input size', type: 'text', spellcheck: 'false', 'aria-label': 'Tamanho da Fonte' });
    const arrow = h('button', { class: 'rb-combo-arrow', type: 'button', tabindex: '-1', html: icon('chevronDown') });
    const box = h('div', { class: 'rb-combo size-combo' }, input, arrow);
    const apply = () => {
      const v = parseFloat(input.value.replace(',', '.'));
      if (v >= 1 && v <= 1638) app.cmd('setFontSize', Math.round(v * 2) / 2);
      app.focusEditor();
    };
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); closeAllPopups(); apply(); }
      if (e.key === 'Escape') { e.preventDefault(); app.focusEditor(); }
    });
    input.addEventListener('focus', () => input.select());
    arrow.addEventListener('mousedown', (e) => e.preventDefault());
    arrow.addEventListener('click', () => {
      const list = h('div', { class: 'size-list' });
      for (const s of [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72]) {
        list.appendChild(h('div', { class: 'font-item' + (String(s) === input.value ? ' sel' : ''), text: String(s), onclick: () => { closeAllPopups(); input.value = String(s); apply(); } }));
      }
      openPanel(box, list, { className: 'size-popup' });
    });
    tooltip(box, 'Tamanho da Fonte', 'Ctrl+Shift+P', 'Altere o tamanho do texto.');
    return this.reg('size', box, { input });
  }

  // ------------------------------------------------------------ galeria de estilos

  styleGallery() {
    const app = this.app;
    const strip = h('div', { class: 'style-strip' });
    const up = h('button', { class: 'sg-btn', type: 'button', title: 'Linha anterior', html: '<svg viewBox="0 0 10 10" width="8" height="8"><path d="M1.5 6.5L5 3l3.5 3.5" fill="none" stroke="currentColor" stroke-width="1.3"/></svg>' });
    const down = h('button', { class: 'sg-btn', type: 'button', title: 'Próxima linha', html: icon('chevronDown') });
    const more = h('button', { class: 'sg-btn more', type: 'button', title: 'Mais', html: '<svg viewBox="0 0 10 10" width="8" height="8"><path d="M1.5 1.5h7M1.5 4L5 7.5 8.5 4" fill="none" stroke="currentColor" stroke-width="1.2"/></svg>' });
    for (const b of [up, down, more]) b.addEventListener('mousedown', (e) => e.preventDefault());
    const gallery = h('div', { class: 'style-gallery' }, strip, h('div', { class: 'sg-side' }, up, down, more));
    let offset = 0;
    const tile = (st) => {
      const t = h('button', { class: 'style-tile', type: 'button', title: app.model.styles.displayName(st) },
        h('span', { class: 'doc-body st-sample' }, h('span', { class: (st.type === 'character' ? cClass(st.id) : pClass(st.id)) + ' st-inner', text: 'AaBbCcDd' })),
        h('span', { class: 'st-name', text: app.model.styles.displayName(st) }));
      t.dataset.style = st.id;
      t.addEventListener('mousedown', (e) => e.preventDefault());
      t.addEventListener('click', () => { closeAllPopups(); app.cmd('applyStyle', st.id); });
      return t;
    };
    const render = () => {
      if (!app.model) return;
      clear(strip);
      const list = app.model.styles.quickStyles();
      const perRow = Math.max(2, Math.floor((strip.clientWidth || 440) / 74));
      offset = Math.max(0, Math.min(offset, Math.max(0, list.length - perRow)));
      for (const st of list.slice(offset, offset + perRow)) strip.appendChild(tile(st));
      this.updateStyleSelection();
    };
    up.addEventListener('click', () => { offset -= Math.max(1, Math.floor((strip.clientWidth || 440) / 74)); render(); });
    down.addEventListener('click', () => { offset += Math.max(1, Math.floor((strip.clientWidth || 440) / 74)); render(); });
    more.addEventListener('click', () => {
      const grid = h('div', { class: 'style-grid' });
      for (const st of app.model.styles.quickStyles()) grid.appendChild(tile(st));
      const menu = h('div', {}, grid, h('div', { class: 'menu-sep' }),
        h('div', { class: 'menu-item', onclick: () => { closeAllPopups(); app.cmd('clearAll'); } }, h('span', { class: 'mi-icon', html: icon('clearFormat') }), h('span', { class: 'mi-label', text: 'Limpar Formatação' })),
        h('div', { class: 'menu-item', onclick: () => { closeAllPopups(); app.stylesPane(); } }, h('span', { class: 'mi-icon', html: icon('styles') }), h('span', { class: 'mi-label', text: 'Aplicar Estilos...' })));
      openPanel(gallery, menu, { className: 'style-popup' });
      this.updateStyleSelection(grid);
    });
    this.renderStyles = render;
    return this.reg('styles', gallery);
  }

  updateStyleSelection(root) {
    const id = this.state && this.state.styleId;
    for (const t of (root || this.bodyEl).querySelectorAll('.style-tile')) t.classList.toggle('active', t.dataset.style === id);
  }

  tableStyleGallery() {
    const app = this.app;
    const wrap = h('div', { class: 'tstyle-gallery' });
    const render = () => {
      if (!app.model) return;
      clear(wrap);
      const styles = app.model.styles.allStyles('table').filter((s) => !s.hidden && !s.semiHidden).slice(0, 12);
      for (const st of styles) {
        const tbl = h('table', { class: 'tbl mini ' + tClass(st.id) });
        const tb = h('tbody');
        for (let r = 0; r < 4; r++) {
          const tr = h('tr');
          for (let c = 0; c < 4; c++) {
            const cls = [r === 0 ? 'r-first' : '', r === 3 ? 'r-last' : '', c === 0 ? 'c-first' : '', c === 3 ? 'c-last' : '',
              r === 0 ? 'cf-firstRow' : (r % 2 === 1 ? 'cf-band1Horz' : 'cf-band2Horz'), c === 0 ? 'cf-firstCol' : ''].filter(Boolean).join(' ');
            tr.appendChild(h('td', { class: cls }, h('p', { class: 'p' }, ' ')));
          }
          tb.appendChild(tr);
        }
        tbl.appendChild(tb);
        const b = h('button', { class: 'tstyle-tile', type: 'button', title: app.model.styles.displayName(st) }, h('div', { class: 'doc-body' }, tbl));
        b.addEventListener('mousedown', (e) => e.preventDefault());
        b.addEventListener('click', () => app.cmd('tableAttrs', { styleId: st.id }));
        wrap.appendChild(b);
      }
    };
    this.renderTableStyles = render;
    return wrap;
  }

  // ------------------------------------------------------------ menus reutilizáveis

  tablePicker() {
    const app = this.app;
    const cells = [];
    const label = h('div', { class: 'tp-label', text: 'Inserir Tabela' });
    const grid = h('div', { class: 'tp-grid' });
    for (let r = 0; r < 8; r++) for (let c = 0; c < 10; c++) {
      const cell = h('div', { class: 'tp-cell' });
      cell.addEventListener('mouseenter', () => {
        for (const x of cells) x.el.classList.toggle('on', x.r <= r && x.c <= c);
        label.textContent = 'Tabela ' + (c + 1) + 'x' + (r + 1);
      });
      cell.addEventListener('click', () => { closeAllPopups(); app.cmd('insertTable', r + 1, c + 1); });
      cells.push({ el: cell, r, c });
      grid.appendChild(cell);
    }
    return h('div', { class: 'table-picker' }, label, grid, h('div', { class: 'menu-sep' }),
      h('div', { class: 'menu-item', onclick: () => { closeAllPopups(); app.dialog('insertTable'); } }, h('span', { class: 'mi-icon', html: icon('table16') }), h('span', { class: 'mi-label', text: 'Inserir Tabela...' })));
  }

  bulletMenu() {
    const app = this.app;
    const grid = h('div', { class: 'lib-grid' });
    grid.appendChild(h('button', { class: 'lib-cell none', type: 'button', text: 'Nenhum', onclick: () => { closeAllPopups(); app.cmd('removeList'); } }));
    BULLET_LIBRARY.forEach((b, i) => grid.appendChild(h('button', { class: 'lib-cell', type: 'button', text: b.show, onclick: () => { closeAllPopups(); app.cmd('list', 'bullet', i); } })));
    return h('div', {}, h('div', { class: 'menu-header', text: 'Biblioteca de Marcadores' }), grid);
  }

  numberMenu() {
    const app = this.app;
    const grid = h('div', { class: 'lib-grid num' });
    grid.appendChild(h('button', { class: 'lib-cell none', type: 'button', text: 'Nenhum', onclick: () => { closeAllPopups(); app.cmd('removeList'); } }));
    NUMBER_LIBRARY.forEach((n, i) => grid.appendChild(h('button', { class: 'lib-cell', type: 'button', onclick: () => { closeAllPopups(); app.cmd('list', 'number', i); } }, ...n.show.split('  ').map((l) => h('div', { class: 'lib-line' }, h('span', { text: l }), h('span', { class: 'lib-bar' }))))));
    return h('div', {}, h('div', { class: 'menu-header', text: 'Biblioteca de Numeração' }), grid,
      h('div', { class: 'menu-sep' }),
      h('div', { class: 'menu-item', onclick: () => { closeAllPopups(); app.cmd('restartNumbering'); } }, h('span', { class: 'mi-icon' }), h('span', { class: 'mi-label', text: 'Reiniciar em 1' })));
  }

  lineSpacingMenu() {
    const app = this.app;
    const cur = this.state ? this.state.lineSpacing : null;
    return [
      ...[1, 1.15, 1.5, 2, 2.5, 3].map((v) => ({ label: fmtNumber(v, 2), checked: cur === v, run: () => app.cmd('lineSpacing', v) })),
      { separator: true },
      { label: 'Opções de Espaçamento de Linha...', run: () => app.dialog('paragraph') },
      { separator: true },
      { label: this.state && this.state.spaceBefore ? 'Remover Espaço Antes do Parágrafo' : 'Adicionar Espaço Antes do Parágrafo', run: () => app.cmd('spaceBefore') },
      { label: this.state && this.state.spaceAfter ? 'Remover Espaço Depois do Parágrafo' : 'Adicionar Espaço Depois do Parágrafo', run: () => app.cmd('spaceAfter') },
    ];
  }

  bordersMenu() {
    const app = this.app;
    return [
      { label: 'Borda Inferior', icon: 'borders', run: () => app.cmd('paraBorders', 'bottom') },
      { label: 'Borda Superior', icon: 'bordersTop', run: () => app.cmd('paraBorders', 'top') },
      { label: 'Borda Esquerda', icon: 'bordersLeft', run: () => app.cmd('paraBorders', 'left') },
      { label: 'Borda Direita', icon: 'bordersRight', run: () => app.cmd('paraBorders', 'right') },
      { separator: true },
      { label: 'Sem Borda', icon: 'bordersNone', run: () => app.cmd('paraBorders', 'none') },
      { label: 'Todas as Bordas', icon: 'bordersAll', run: () => app.cmd('paraBorders', 'all') },
      { label: 'Bordas Externas', icon: 'bordersOutside', run: () => app.cmd('paraBorders', 'outside') },
      { separator: true },
      { label: 'Linha Horizontal', icon: 'hr16', run: () => app.cmd('horizontalLine') },
    ];
  }

  caseMenu() {
    const app = this.app;
    return [
      { label: 'Primeira letra da frase em maiúscula.', run: () => app.cmd('changeCase', 'sentence') },
      { label: 'minúsculas', run: () => app.cmd('changeCase', 'lower') },
      { label: 'MAIÚSCULAS', run: () => app.cmd('changeCase', 'upper') },
      { label: 'Colocar Cada Palavra Em Maiúscula', run: () => app.cmd('changeCase', 'title') },
      { label: 'aLTERNAR mAIÚSC./mINÚSC.', run: () => app.cmd('changeCase', 'toggle') },
    ];
  }

  underlineMenu() {
    const app = this.app;
    const item = (val, label, css) => ({
      element: h('div', { class: 'menu-item ul-item', onclick: () => { closeAllPopups(); app.cmd('setUnderline', val); } },
        h('span', { class: 'mi-icon' }), h('span', { class: 'ul-sample', style: css }, label)),
    });
    return [
      item('single', 'Sublinhado', { textDecoration: 'underline' }),
      item('double', 'Sublinhado duplo', { textDecoration: 'underline double' }),
      item('thick', 'Sublinhado grosso', { textDecoration: 'underline', textDecorationThickness: '2px' }),
      item('dotted', 'Sublinhado pontilhado', { textDecoration: 'underline dotted' }),
      item('dash', 'Sublinhado tracejado', { textDecoration: 'underline dashed' }),
      item('wave', 'Sublinhado ondulado', { textDecoration: 'underline wavy' }),
      { separator: true },
      { label: 'Mais Sublinhados...', run: () => app.dialog('font') },
    ];
  }

  symbolMenu() {
    const app = this.app;
    const grid = h('div', { class: 'sym-quick' });
    const list = [...new Set([...app.recentSymbols(), ...SYMBOLS])].slice(0, 40);
    for (const s of list) grid.appendChild(h('button', { class: 'sym-cell', type: 'button', text: s, onclick: () => { closeAllPopups(); app.insertSymbol(s); } }));
    return h('div', {}, grid, h('div', { class: 'menu-sep' }),
      h('div', { class: 'menu-item', onclick: () => { closeAllPopups(); app.dialog('symbol'); } }, h('span', { class: 'mi-icon', text: 'Ω' }), h('span', { class: 'mi-label', text: 'Mais Símbolos...' })));
  }

  // ------------------------------------------------------------ abas

  build() {
    const app = this.app;
    const A = (name, ...args) => () => app.cmd(name, ...args);
    const T = (name) => () => app.cmd('toggle', name);

    this.addTab('file', 'Arquivo', null, { file: true });

    this.addTab('home', 'Página Inicial', () => [
      this.group('Área de Transferência', [
        this.big('paste', 'Colar', 'paste', () => app.edit('paste'), {
          split: true, shortcut: 'Ctrl+V', desc: 'Adicione conteúdo da Área de Transferência ao documento.',
          menu: () => [
            { header: 'Opções de Colagem:' },
            { label: 'Manter Formatação Original', icon: 'paste16', run: () => app.edit('paste') },
            { label: 'Manter Somente Texto', run: () => app.edit('pasteAndMatchStyle') },
          ],
        }),
        this.col(
          this.small('cut', 'cut', () => app.edit('cut'), { label: 'Recortar', shortcut: 'Ctrl+X' }),
          this.small('copy', 'copy', () => app.edit('copy'), { label: 'Copiar', shortcut: 'Ctrl+C' }),
          this.small('painter', 'painter', () => app.formatPainter(false), { label: 'Pincel de Formatação', shortcut: 'Ctrl+Shift+C', toggle: true, desc: 'Copie a formatação de um lugar e aplique em outro. Clique duas vezes para aplicar em vários lugares.' }),
        ),
      ]),
      this.group('Fonte', [
        this.col(
          this.row(this.fontCombo(), this.sizeCombo(),
            this.small('grow', 'growFont', A('growFont', 1), { tip: 'Aumentar Tamanho da Fonte', shortcut: 'Ctrl+>' }),
            this.small('shrink', 'shrinkFont', A('growFont', -1), { tip: 'Diminuir Tamanho da Fonte', shortcut: 'Ctrl+<' }),
            this.sep(),
            this.small('case', 'changeCase', null, { tip: 'Maiúsculas e Minúsculas', menu: () => this.caseMenu() }),
            this.small('clearFmt', 'clearFormat', A('clearAll'), { tip: 'Limpar Toda a Formatação' })),
          this.row(
            this.small('b', 'bold', T('b'), { tip: 'Negrito', shortcut: app.sc('bold'), toggle: true }),
            this.small('i', 'italic', T('i'), { tip: 'Itálico', shortcut: app.sc('italic'), toggle: true }),
            this.small('u', 'underline', T('u'), { tip: 'Sublinhado', shortcut: app.sc('underline'), toggle: true, split: true, menu: () => this.underlineMenu() }),
            this.small('strike', 'strike', T('strike'), { tip: 'Tachado', toggle: true }),
            this.small('sub', 'subscript', A('vertAlign', 'subscript'), { tip: 'Subscrito', shortcut: 'Ctrl+=', toggle: true }),
            this.small('sup', 'superscript', A('vertAlign', 'superscript'), { tip: 'Sobrescrito', shortcut: 'Ctrl+Shift++', toggle: true }),
            this.sep(),
            this.small('highlight', 'highlight', () => app.cmd('highlight', app.lastHighlight), { tip: 'Cor do Realce do Texto', split: true, colorBar: HIGHLIGHT[app.lastHighlight] || '#ff0', menu: () => highlightPanel((v) => { if (v) app.lastHighlight = v; this.paintBars(); app.cmd('highlight', v); }) }),
            this.small('color', 'fontColor', () => app.cmd('setColor', app.lastColor), { tip: 'Cor da Fonte', split: true, colorBar: '#' + app.lastColor.val, menu: () => colorPanel(app.model.theme, (c) => { if (c && c !== 'auto') app.lastColor = c; this.paintBars(); app.cmd('setColor', c === 'auto' ? { val: 'auto' } : c); }) }),
          ),
        ),
      ], () => app.dialog('font')),
      this.group('Parágrafo', [
        this.col(
          this.row(
            this.small('bullets', 'bullets', A('list', 'bullet'), { tip: 'Marcadores', toggle: true, split: true, menu: () => this.bulletMenu() }),
            this.small('numbering', 'numbering', A('list', 'number'), { tip: 'Numeração', toggle: true, split: true, menu: () => this.numberMenu() }),
            this.sep(),
            this.small('outdent', 'outdent', A('indent', -1), { tip: 'Diminuir Recuo' }),
            this.small('indent', 'indent', A('indent', 1), { tip: 'Aumentar Recuo' }),
            this.sep(),
            this.small('sort', 'sort', () => app.sortParagraphs(), { tip: 'Classificar', desc: 'Coloque os parágrafos selecionados em ordem alfabética.' }),
            this.small('marks', 'pilcrow', () => app.toggleMarks(), { tip: 'Mostrar Tudo', shortcut: 'Ctrl+*', toggle: true, desc: 'Mostre as marcas de parágrafo e outros símbolos de formatação ocultos.' }),
          ),
          this.row(
            this.small('alignLeft', 'alignLeft', A('align', 'left'), { tip: 'Alinhar à Esquerda', shortcut: app.sc('alignLeft'), toggle: true }),
            this.small('alignCenter', 'alignCenter', A('align', 'center'), { tip: 'Centralizar', shortcut: app.sc('alignCenter'), toggle: true }),
            this.small('alignRight', 'alignRight', A('align', 'right'), { tip: 'Alinhar à Direita', shortcut: app.sc('alignRight'), toggle: true }),
            this.small('justify', 'justify', A('align', 'both'), { tip: 'Justificar', shortcut: app.sc('justify'), toggle: true }),
            this.sep(),
            this.small('lineSpacing', 'lineSpacing', null, { tip: 'Espaçamento de Linha e Parágrafo', menu: () => this.lineSpacingMenu() }),
            this.sep(),
            this.small('shading', 'shading', () => app.cmd('paraShading', app.lastShading), { tip: 'Sombreamento', split: true, colorBar: app.lastShading || '#fff', menu: () => colorPanel(app.model.theme, (c) => { const v = c && c !== 'auto' ? '#' + c.val : null; app.lastShading = v; this.paintBars(); app.cmd('paraShading', v); }, { auto: false, none: true }) }),
            this.small('borders', 'borders', A('paraBorders', 'bottom'), { tip: 'Bordas', split: true, menu: () => this.bordersMenu() }),
          ),
        ),
      ], () => app.dialog('paragraph')),
      this.group('Estilos', [this.styleGallery()], () => app.stylesPane()),
      this.group('Edição', [
        this.col(
          this.small('find', 'find', () => app.showNavPane('search'), { label: 'Localizar', split: true, shortcut: app.sc('find'), menu: () => [
            { label: 'Localizar', icon: 'find', run: () => app.showNavPane('search') },
            { label: 'Localização Avançada...', run: () => app.dialog('find', 0) },
            { label: 'Ir para...', run: () => app.dialog('find', 2) },
          ] }),
          this.small('replace', 'replace', () => app.dialog('find', 1), { label: 'Substituir', shortcut: app.sc('replace') }),
          this.small('selectMenu', 'select', null, { label: 'Selecionar', menu: () => [
            { label: 'Selecionar Tudo', shortcut: app.sc('selectAll'), run: () => app.cmd('selectAll') },
            { label: 'Selecionar Parágrafo', run: () => app.cmd('selectParagraph') },
          ] }),
        ),
      ]),
    ]);

    this.addTab('insert', 'Inserir', () => [
      this.group('Páginas', [
        this.col(
          this.small('blankPage', 'newdoc16', () => app.insertBlankPage(), { label: 'Página em Branco' }),
          this.small('pageBreak', 'pageBreak16', A('pageBreak'), { label: 'Quebra de Página', shortcut: 'Ctrl+Enter' }),
        ),
      ]),
      this.group('Tabelas', [this.big('table', 'Tabela', 'table', null, { menu: () => this.tablePicker(), desc: 'Adicione uma tabela ao documento.' })]),
      this.group('Ilustrações', [this.big('pictures', 'Imagens', 'picture', () => app.insertPictureFromFile(), { desc: 'Insira imagens do computador.' })]),
      this.group('Links', [this.big('link', 'Link', 'link', () => app.dialog('link'), { shortcut: 'Ctrl+K', desc: 'Crie um link para uma página da Web ou para um local no documento.' })]),
      this.group('Cabeçalho e Rodapé', [
        this.big('header', 'Cabeçalho', 'header', null, { menu: () => [
          { label: 'Editar Cabeçalho', run: () => app.editHeaderFooter('header') },
          { label: 'Remover Cabeçalho', run: () => app.removeHeaderFooter('header') },
        ] }),
        this.big('footer', 'Rodapé', 'footer', null, { menu: () => [
          { label: 'Editar Rodapé', run: () => app.editHeaderFooter('footer') },
          { label: 'Remover Rodapé', run: () => app.removeHeaderFooter('footer') },
        ] }),
        this.big('pageNum', 'Número de Página', 'pageNumber', null, { menu: () => this.pageNumberMenu() }),
      ]),
      this.group('Texto', [
        this.col(
          this.small('dateTime', 'newdoc16', () => app.dialog('dateTime'), { label: 'Data e Hora' }),
          this.small('hline', 'hr16', A('horizontalLine'), { label: 'Linha Horizontal' }),
        ),
      ]),
      this.group('Símbolos', [this.big('symbol', 'Símbolo', 'symbol', null, { menu: () => this.symbolMenu() })]),
    ]);

    this.addTab('layout', 'Layout', () => [
      this.group('Configurar Página', [
        this.big('margins', 'Margens', 'margins', null, { menu: () => this.marginsMenu() }),
        this.big('orientation', 'Orientação', 'orientation', null, { menu: () => [
          { label: 'Retrato', icon: 'printLayout16', checked: this.state && !this.state.landscape, run: () => app.setOrientation('portrait') },
          { label: 'Paisagem', icon: 'webLayout16', checked: this.state && this.state.landscape, run: () => app.setOrientation('landscape') },
        ] }),
        this.big('paperSize', 'Tamanho', 'size', null, { menu: () => [
          ...PAPER_SIZES.map((p) => ({ label: p.name + '   ' + p.desc, checked: this.state && this.state.paper === p.id, run: () => app.setPaper(p) })),
          { separator: true },
          { label: 'Mais Tamanhos de Papel...', run: () => app.dialog('pageSetup', 1) },
        ] }),
        this.col(
          this.small('breaks', 'pageBreak16', null, { label: 'Quebras', menu: () => [
            { header: 'Quebras de Página' },
            { label: 'Página', run: A('pageBreak') },
            { label: 'Quebra Automática de Texto', run: A('lineBreak') },
          ] }),
        ),
      ], () => app.dialog('pageSetup')),
      this.group('Parágrafo', [this.paraSpacingControls()], () => app.dialog('paragraph')),
    ]);

    this.addTab('references', 'Referências', () => [
      this.group('Sumário', [
        this.big('toc', 'Sumário', 'toc', null, { menu: () => [
          { label: 'Sumário Automático', run: () => app.insertTOC() },
          { label: 'Remover Sumário', run: () => app.removeTOC() },
        ] }),
        this.col(this.small('updateToc', 'updateToc16', () => app.updateTOC(), { label: 'Atualizar Sumário' })),
      ]),
    ]);

    this.addTab('review', 'Revisão', () => [
      this.group('Revisão de Texto', [
        this.big('spelling', 'Ortografia', 'spelling', () => app.toggleSpellcheck(), { toggle: true, desc: 'Ative ou desative a verificação ortográfica enquanto você digita.' }),
        this.col(this.small('wordCount', 'spell16', () => app.dialog('wordCount'), { label: 'Contar Palavras' })),
      ]),
      this.group('Idioma', [this.big('language', 'Idioma', 'language', null, { menu: () => app.languageMenu() })]),
      this.group('Controle', [
        this.col(
          this.small('markup', 'trackChanges16', null, { label: 'Marcação', menu: () => [
            { label: 'Toda a Marcação', checked: app.showMarkup, run: () => app.setShowMarkup(true) },
            { label: 'Sem Marcação', checked: !app.showMarkup, run: () => app.setShowMarkup(false) },
          ] }),
        ),
      ]),
      this.group('Alterações', [
        this.big('accept', 'Aceitar', 'accept', null, { menu: () => [{ label: 'Aceitar Todas as Alterações', run: () => app.acceptAllChanges() }] }),
        this.big('reject', 'Rejeitar', 'reject', null, { menu: () => [{ label: 'Rejeitar Todas as Alterações', run: () => app.rejectAllChanges() }] }),
      ]),
    ]);

    this.addTab('view', 'Exibir', () => [
      this.group('Modos de Exibição', [
        this.big('readMode', 'Modo de Leitura', 'readMode', () => app.setViewMode('read'), { toggle: true }),
        this.big('printMode', 'Layout de Impressão', 'printLayout', () => app.setViewMode('print'), { toggle: true }),
        this.big('webMode', 'Layout da Web', 'webLayout', () => app.setViewMode('web'), { toggle: true }),
      ]),
      this.group('Mostrar', [
        this.col(
          this.check('showRuler', 'Régua', (on) => app.setRuler(on)),
          this.check('showNav', 'Painel de Navegação', (on) => (on ? app.showNavPane() : app.hideNavPane())),
        ),
      ]),
      this.group('Zoom', [
        this.big('zoomDlg', 'Zoom', 'zoom', () => app.dialog('zoom')),
        this.big('zoom100', '100%', 'zoom100', () => app.setZoom(1)),
        this.col(
          this.small('onePage', 'printLayout16', () => app.zoomTo('page'), { label: 'Uma Página' }),
          this.small('pageWidth', 'webLayout16', () => app.zoomTo('width'), { label: 'Largura da Página' }),
        ),
      ]),
      this.group('Janela', [this.big('newWindow', 'Nova Janela', 'newWindow', () => app.newWindow(true))]),
    ]);

    // abas contextuais
    this.addTab('tableDesign', 'Design da Tabela', () => [
      this.group('Opções de Estilo de Tabela', [
        this.col(this.check('lookFirstRow', 'Linha de Cabeçalho', (on) => app.setTableLook('firstRow', on)), this.check('lookLastRow', 'Linha de Total', (on) => app.setTableLook('lastRow', on)), this.check('lookHBand', 'Linhas em Tiras', (on) => app.setTableLook('noHBand', !on))),
        this.col(this.check('lookFirstCol', 'Primeira Coluna', (on) => app.setTableLook('firstColumn', on)), this.check('lookLastCol', 'Última Coluna', (on) => app.setTableLook('lastColumn', on)), this.check('lookVBand', 'Colunas em Tiras', (on) => app.setTableLook('noVBand', !on))),
      ]),
      this.group('Estilos de Tabela', [this.tableStyleGallery(),
        this.col(this.small('cellShading', 'shading', null, { label: 'Sombreamento', menu: () => colorPanel(app.model.theme, (c) => app.cmd('cellAttrs', { shd: c && c !== 'auto' ? { val: 'clear', color: 'auto', fill: c.val } : null }), { auto: false, none: true }) }))]),
      this.group('Bordas', [
        this.big('tblBorders', 'Bordas', 'tableBordersBig', null, { menu: () => [
          { label: 'Todas as Bordas', icon: 'bordersAll', run: () => app.cmd('tableBorders', 'all') },
          { label: 'Sem Borda', icon: 'bordersNone', run: () => app.cmd('tableBorders', 'none') },
          { label: 'Bordas Externas', icon: 'bordersOutside', run: () => app.cmd('tableBorders', 'outside') },
          { label: 'Bordas Internas', icon: 'bordersInside', run: () => app.cmd('tableBorders', 'inside') },
        ] }),
      ]),
    ], { contextual: true });

    this.addTab('tableLayout', 'Layout', () => [
      this.group('Tabela', [
        this.col(
          this.small('selRow', 'select', null, { label: 'Selecionar', menu: () => [
            { label: 'Selecionar Célula', run: () => app.selectTablePart('cell') },
            { label: 'Selecionar Coluna', run: () => app.selectTablePart('column') },
            { label: 'Selecionar Linha', run: () => app.selectTablePart('row') },
            { label: 'Selecionar Tabela', run: () => app.selectTablePart('table') },
          ] }),
          this.small('tblProps', 'settings16', () => app.dialog('tableProperties'), { label: 'Propriedades' }),
        ),
      ]),
      this.group('Linhas e Colunas', [
        this.big('tblDelete', 'Excluir', 'deleteTableBig', null, { menu: () => [
          { label: 'Excluir Colunas', icon: 'deleteRow', run: A('table', 'deleteColumn') },
          { label: 'Excluir Linhas', icon: 'deleteRow', run: A('table', 'deleteRow') },
          { label: 'Excluir Tabela', icon: 'deleteTable', run: A('table', 'deleteTable') },
        ] }),
        this.big('rowAbove', 'Inserir Acima', 'rowAboveBig', A('table', 'addRowBefore')),
        this.col(
          this.small('rowBelow', 'rowBelow', A('table', 'addRowAfter'), { label: 'Inserir Abaixo' }),
          this.small('colLeft', 'colLeft', A('table', 'addColumnBefore'), { label: 'Inserir à Esquerda' }),
          this.small('colRight', 'colRight', A('table', 'addColumnAfter'), { label: 'Inserir à Direita' }),
        ),
      ]),
      this.group('Mesclar', [
        this.col(
          this.small('merge', 'merge', A('table', 'mergeCells'), { label: 'Mesclar Células' }),
          this.small('split', 'split', A('table', 'splitCell'), { label: 'Dividir Células' }),
        ),
      ]),
      this.group('Alinhamento', [
        this.col(
          this.small('vTop', 'cellTop', () => app.cmd('cellAttrs', { vAlign: null }), { label: 'Superior' }),
          this.small('vMid', 'cellMiddle', () => app.cmd('cellAttrs', { vAlign: 'center' }), { label: 'Centro' }),
          this.small('vBot', 'cellBottom', () => app.cmd('cellAttrs', { vAlign: 'bottom' }), { label: 'Inferior' }),
        ),
        this.col(this.small('headerRow', 'headerRow', A('table', 'toggleHeaderRow'), { label: 'Repetir Cabeçalho', toggle: true })),
      ]),
    ], { contextual: true });

    this.addTab('picture', 'Formato da Imagem', () => [
      this.group('Tamanho', [this.pictureSizeControls()], () => app.dialog('imageSize')),
      this.group('Acessibilidade', [this.big('altText', 'Texto Alt', 'picture', () => app.dialog('imageSize'))]),
    ], { contextual: true });

    this.addTab('headerFooter', 'Cabeçalho e Rodapé', () => [
      this.group('Cabeçalho e Rodapé', [
        this.big('hfPageNum', 'Número de Página', 'pageNumber', null, { menu: () => this.pageNumberMenu() }),
      ]),
      this.group('Inserir', [
        this.big('hfDate', 'Data e Hora', 'dateTime', () => app.dialog('dateTime')),
        this.big('hfPicture', 'Imagens', 'picture', () => app.insertPictureFromFile()),
      ]),
      this.group('Navegação', [
        this.col(
          this.small('gotoHeader', 'header16', () => app.switchHeaderFooter('header'), { label: 'Ir para Cabeçalho' }),
          this.small('gotoFooter', 'footer16', () => app.switchHeaderFooter('footer'), { label: 'Ir para Rodapé' }),
        ),
      ]),
      this.group('Opções', [
        this.col(
          this.check('hfFirst', 'Primeira Página Diferente', (on) => app.setTitlePage(on)),
          this.check('hfEvenOdd', 'Páginas Pares e Ímpares Diferentes', (on) => app.setEvenOdd(on)),
        ),
      ]),
      this.group('Fechar', [this.big('closeHF', 'Fechar Cabeçalho e Rodapé', 'closeHF', () => app.closeHeaderFooter())]),
    ], { contextual: true });

    this.select('home');
  }

  pageNumberMenu() {
    const app = this.app;
    return [
      { header: 'Início da Página' },
      { label: 'Número Simples à Esquerda', run: () => app.insertPageNumber('header', 'left') },
      { label: 'Número Simples ao Centro', run: () => app.insertPageNumber('header', 'center') },
      { label: 'Número Simples à Direita', run: () => app.insertPageNumber('header', 'right') },
      { header: 'Fim da Página' },
      { label: 'Número Simples à Esquerda', run: () => app.insertPageNumber('footer', 'left') },
      { label: 'Número Simples ao Centro', run: () => app.insertPageNumber('footer', 'center') },
      { label: 'Número Simples à Direita', run: () => app.insertPageNumber('footer', 'right') },
      { separator: true },
      { label: '"Página X de Y" no rodapé', run: () => app.insertPageNumber('footer', 'center', true) },
      { label: 'Começar numeração em...', run: () => app.pageNumberStart() },
    ];
  }

  marginsMenu() {
    const app = this.app;
    const cm = (tw) => fmtNumber(twipToCm(tw), 2) + ' cm';
    const items = MARGIN_PRESETS.map((p) => ({
      element: h('div', { class: 'menu-item margin-item' + (this.state && this.state.marginPreset === p.id ? ' checked' : ''), onclick: () => { closeAllPopups(); app.setMargins(p); } },
        h('span', { class: 'mi-icon big', html: icon('margins') }),
        h('span', { class: 'mi-label' }, h('b', { text: p.name }),
          h('span', { class: 'mi-desc', text: 'Superior: ' + cm(p.top) + '   Inferior: ' + cm(p.bottom) }),
          h('span', { class: 'mi-desc', text: 'Esquerda: ' + cm(p.left) + '   Direita: ' + cm(p.right) }))),
    }));
    return [...items, { separator: true }, { label: 'Margens Personalizadas...', run: () => app.dialog('pageSetup', 0) }];
  }

  paraSpacingControls() {
    const app = this.app;
    const mk = (id, unit, apply) => {
      const s = spinner(0, { unit, min: unit === 'pt' ? 0 : -50, width: '84px' });
      const commit = () => { const v = s.get(); if (v != null) apply(v); };
      s.input.addEventListener('change', commit);
      s.input.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); commit(); app.focusEditor(); } });
      s.querySelectorAll('.spin-btn').forEach((b) => b.addEventListener('click', commit));
      this.reg(id, s, { spinner: s });
      return s;
    };
    const left = mk('indLeft', 'cm', (v) => app.cmd('indentValues', { left: cmToTwip(v) }));
    const right = mk('indRight', 'cm', (v) => app.cmd('indentValues', { right: cmToTwip(v) }));
    const before = mk('spBefore', 'pt', (v) => app.cmd('spacing', v * 20, null));
    const after = mk('spAfter', 'pt', (v) => app.cmd('spacing', null, v * 20));
    return h('div', { class: 'rb-para-spacing' },
      h('div', { class: 'rb-ps-col' }, h('div', { class: 'rb-ps-title', text: 'Recuar' }),
        h('label', { class: 'rb-ps-row' }, h('span', { text: 'Esquerda:' }), left),
        h('label', { class: 'rb-ps-row' }, h('span', { text: 'Direita:' }), right)),
      h('div', { class: 'rb-ps-col' }, h('div', { class: 'rb-ps-title', text: 'Espaçamento' }),
        h('label', { class: 'rb-ps-row' }, h('span', { text: 'Antes:' }), before),
        h('label', { class: 'rb-ps-row' }, h('span', { text: 'Depois:' }), after)));
  }

  pictureSizeControls() {
    const app = this.app;
    const mk = (id, dim) => {
      const s = spinner(0, { unit: 'cm', min: 0.1, width: '84px' });
      const commit = () => { const v = s.get(); if (v) app.resizeDrawing(dim, v); };
      s.input.addEventListener('change', commit);
      s.input.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); commit(); } });
      s.querySelectorAll('.spin-btn').forEach((b) => b.addEventListener('click', commit));
      this.reg(id, s, { spinner: s });
      return s;
    };
    return h('div', { class: 'rb-para-spacing' }, h('div', { class: 'rb-ps-col' },
      h('label', { class: 'rb-ps-row' }, h('span', { text: 'Altura:' }), mk('picH', 'h')),
      h('label', { class: 'rb-ps-row' }, h('span', { text: 'Largura:' }), mk('picW', 'w'))));
  }

  addTab(id, label, builder, opts = {}) {
    const btn = h('button', { class: 'tab' + (opts.file ? ' file' : '') + (opts.contextual ? ' contextual' : ''), type: 'button', role: 'tab', text: label });
    btn.addEventListener('mousedown', (e) => e.preventDefault());
    btn.addEventListener('click', () => {
      if (opts.file) { this.app.openBackstage(); return; }
      if (this.active === id && !this.collapsed) { return; }
      this.select(id);
    });
    btn.addEventListener('dblclick', () => { if (!opts.file) this.toggleCollapse(); });
    if (opts.contextual) btn.hidden = true;
    this.tabsEl.appendChild(btn);
    const tab = { id, label, btn, builder, body: null, contextual: !!opts.contextual };
    this.tabs.push(tab);
    return tab;
  }

  select(id) {
    const tab = this.tabs.find((t) => t.id === id);
    if (!tab || !tab.builder) return;
    this.active = id;
    for (const t of this.tabs) t.btn.classList.toggle('active', t.id === id);
    if (!tab.body) {
      tab.body = h('div', { class: 'rb-tab-body' }, ...tab.builder());
      this.bodyEl.appendChild(tab.body);
    }
    for (const t of this.tabs) if (t.body) t.body.hidden = t.id !== id;
    if (this.collapsed) this.bodyEl.classList.add('peek');
    if (id === 'home' && this.renderStyles) requestAnimationFrame(() => this.renderStyles());
    if (id === 'tableDesign' && this.renderTableStyles) this.renderTableStyles();
    if (this.state) this.update(this.state);
    this.app.onRibbonResize();
  }

  toggleCollapse() {
    this.collapsed = !this.collapsed;
    this.el.classList.toggle('collapsed', this.collapsed);
    this.app.onRibbonResize();
  }

  showContextual(ids) {
    for (const t of this.tabs) {
      if (!t.contextual) continue;
      const show = ids.includes(t.id);
      if (t.btn.hidden === show) t.btn.hidden = !show;
      if (!show && this.active === t.id) this.select('home');
    }
  }

  paintBars() {
    const app = this.app;
    const set = (id, color) => { const c = this.controls.get(id); if (c) { const bar = c.el.querySelector('.rb-colorbar'); if (bar) bar.style.background = color; } };
    set('highlight', HIGHLIGHT[app.lastHighlight] || '#ffff00');
    set('color', '#' + (app.lastColor.val || 'FF0000'));
    set('shading', app.lastShading || '#ffffff');
  }

  // ------------------------------------------------------------ estado

  update(st) {
    this.state = st;
    const on = (id, v) => { const c = this.controls.get(id); if (c) c.el.classList.toggle('active', !!v); };
    const dis = (id, v) => { const c = this.controls.get(id); if (c) { c.el.classList.toggle('disabled', !!v); c.el.disabled = !!v; } };
    on('b', st.bold); on('i', st.italic); on('u', st.underline); on('strike', st.strike);
    on('sub', st.sub); on('sup', st.sup);
    on('alignLeft', st.align === 'left'); on('alignCenter', st.align === 'center'); on('alignRight', st.align === 'right'); on('justify', st.align === 'both');
    on('bullets', st.list === 'bullet'); on('numbering', st.list === 'number');
    on('marks', st.showMarks); on('painter', st.painter);
    on('readMode', st.viewMode === 'read'); on('printMode', st.viewMode === 'print'); on('webMode', st.viewMode === 'web');
    on('spelling', st.spellcheck); on('headerRow', st.headerRow);
    const font = this.controls.get('font');
    if (font && document.activeElement !== font.input) font.input.value = st.font || '';
    const size = this.controls.get('size');
    if (size && document.activeElement !== size.input) size.input.value = st.size ? fmtNumber(st.size, 1) : '';
    const setCheck = (id, v) => { const c = this.controls.get(id); if (c && c.input) c.input.checked = !!v; };
    setCheck('showRuler', st.ruler); setCheck('showNav', st.navPane);
    setCheck('hfFirst', st.titlePg); setCheck('hfEvenOdd', st.evenOdd);
    if (st.look) {
      setCheck('lookFirstRow', st.look.firstRow); setCheck('lookLastRow', st.look.lastRow); setCheck('lookHBand', !st.look.noHBand);
      setCheck('lookFirstCol', st.look.firstColumn); setCheck('lookLastCol', st.look.lastColumn); setCheck('lookVBand', !st.look.noVBand);
    }
    const setSpin = (id, v) => { const c = this.controls.get(id); if (c && c.spinner && document.activeElement !== c.spinner.input) c.spinner.set(v); };
    if (st.para) {
      setSpin('indLeft', twipToCm(st.para.left || 0));
      setSpin('indRight', twipToCm(st.para.right || 0));
      setSpin('spBefore', (st.para.before || 0) / 20);
      setSpin('spAfter', (st.para.after || 0) / 20);
    }
    if (st.drawing) { setSpin('picW', st.drawing.cx / 360000); setSpin('picH', st.drawing.cy / 360000); }
    this.updateStyleSelection();
    const ctx = [];
    if (st.inTable) ctx.push('tableDesign', 'tableLayout');
    if (st.drawing) ctx.push('picture');
    if (st.headerFooter) ctx.push('headerFooter');
    this.showContextual(ctx);
    dis('pageBreak', !st.canPageBreak);
  }
}
