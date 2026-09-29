// Seletores de cor no estilo do Office.
import { h } from './dom.js';
import { themePalette, STANDARD_COLORS, HIGHLIGHT, HIGHLIGHT_LIST } from '../util/color.js';
import { closeAllPopups } from './popup.js';

// onPick recebe {val, themeColor, themeTint, themeShade} ou 'auto' ou null
export function colorPanel(theme, onPick, opts = {}) {
  const wrap = h('div', { class: 'color-panel' });
  const pick = (v) => { closeAllPopups(); onPick(v); };
  if (opts.auto !== false) {
    wrap.appendChild(h('div', { class: 'cp-row-btn', onclick: () => pick('auto') },
      h('span', { class: 'cp-swatch', style: { background: opts.autoColor || '#000' } }), h('span', { text: opts.autoLabel || 'Automático' })));
  }
  wrap.appendChild(h('div', { class: 'cp-title', text: 'Cores do Tema' }));
  const pal = themePalette(theme ? theme.colors : {});
  const grid = h('div', { class: 'cp-theme' });
  const baseRow = h('div', { class: 'cp-row' });
  for (const c of pal) {
    baseRow.appendChild(h('button', {
      class: 'cp-cell', title: c.name, style: { background: c.hex },
      onclick: () => pick({ val: c.hex.slice(1), themeColor: c.themeColor }),
    }));
  }
  grid.appendChild(baseRow);
  const variants = h('div', { class: 'cp-variants' });
  for (let r = 0; r < 5; r++) {
    const row = h('div', { class: 'cp-row' });
    for (const c of pal) {
      const v = c.variants[r];
      row.appendChild(h('button', {
        class: 'cp-cell', title: v.label, style: { background: v.hex },
        onclick: () => pick({ val: v.hex.slice(1), themeColor: v.themeColor, ...(v.tint ? { themeTint: v.tint } : {}), ...(v.shade ? { themeShade: v.shade } : {}) }),
      }));
    }
    variants.appendChild(row);
  }
  grid.appendChild(variants);
  wrap.appendChild(grid);
  wrap.appendChild(h('div', { class: 'cp-title', text: 'Cores Padrão' }));
  const std = h('div', { class: 'cp-row' });
  for (const [hex, name] of STANDARD_COLORS) std.appendChild(h('button', { class: 'cp-cell', title: name, style: { background: hex }, onclick: () => pick({ val: hex.slice(1) }) }));
  wrap.appendChild(std);
  if (opts.none) {
    wrap.appendChild(h('div', { class: 'cp-row-btn', onclick: () => pick(null) }, h('span', { class: 'cp-swatch none' }), h('span', { text: opts.noneLabel || 'Sem Cor' })));
  }
  const input = h('input', { type: 'color', class: 'cp-native' });
  input.addEventListener('change', () => pick({ val: input.value.slice(1).toUpperCase() }));
  wrap.appendChild(h('div', { class: 'cp-row-btn', onclick: () => input.click() }, h('span', { class: 'cp-swatch rainbow' }), h('span', { text: 'Mais Cores...' })));
  wrap.appendChild(input);
  return wrap;
}

export function highlightPanel(onPick) {
  const wrap = h('div', { class: 'color-panel hl-panel' });
  const grid = h('div', { class: 'cp-hl-grid' });
  for (const [key, name] of HIGHLIGHT_LIST) {
    grid.appendChild(h('button', { class: 'cp-cell big', title: name, style: { background: HIGHLIGHT[key] }, onclick: () => { closeAllPopups(); onPick(key); } }));
  }
  wrap.appendChild(grid);
  wrap.appendChild(h('div', { class: 'cp-row-btn', onclick: () => { closeAllPopups(); onPick(null); } }, h('span', { class: 'cp-swatch none' }), h('span', { text: 'Sem Cor' })));
  return wrap;
}
