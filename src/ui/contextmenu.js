// Menu de contexto com a minibarra de ferramentas, como no Word.
import { h } from './dom.js';
import { icon } from './icons.js';
import { buildMenu, Popup, closeAllPopups, openPanel } from './popup.js';
import { colorPanel, highlightPanel } from './colorpicker.js';
import { HIGHLIGHT } from '../util/color.js';

export function openContextMenu(app, info) {
  closeAllPopups();
  const items = [];
  if (info.misspelledWord) {
    const sugg = (info.suggestions || []).slice(0, 6);
    if (sugg.length) for (const s of sugg) items.push({ label: s, bold: true, run: () => app.replaceMisspelling(s) });
    else items.push({ label: '(Sem Sugestões de Ortografia)', disabled: true });
    items.push({ label: 'Adicionar ao Dicionário', run: () => { if (app.host.addWord) app.host.addWord(info.misspelledWord); else app.host.addToDictionary(info.misspelledWord); } });
    items.push({ separator: true });
  }
  items.push({ label: 'Recortar', icon: 'cut', shortcut: 'Ctrl+X', disabled: !info.hasSelection, run: () => app.edit('cut') });
  items.push({ label: 'Copiar', icon: 'copy', shortcut: 'Ctrl+C', disabled: !info.hasSelection, run: () => app.edit('copy') });
  items.push({ label: 'Colar', icon: 'paste16', shortcut: 'Ctrl+V', run: () => app.edit('paste') });
  items.push({ label: 'Colar Somente Texto', run: () => app.edit('pasteAndMatchStyle') });
  items.push({ separator: true });
  if (info.link) {
    items.push({ label: 'Editar Hiperlink...', icon: 'link16', run: () => app.dialog('link') });
    items.push({ label: 'Abrir Hiperlink', run: () => app.followLink(info.link) });
    items.push({ label: 'Copiar Hiperlink', run: () => navigator.clipboard.writeText(info.link.href || '#' + (info.link.anchor || '')) });
    items.push({ label: 'Remover Hiperlink', icon: 'unlink16', run: () => app.cmd('removeLink') });
    items.push({ separator: true });
  }
  if (info.drawing) {
    items.push({ label: 'Tamanho e Posição...', icon: 'picture16', run: () => app.dialog('imageSize') });
    items.push({ label: 'Salvar como Imagem...', run: () => app.saveDrawingAsImage() });
    items.push({ separator: true });
  }
  if (info.inTable) {
    items.push({ label: 'Inserir', submenu: [
      { label: 'Inserir Colunas à Esquerda', icon: 'colLeft', run: () => app.cmd('table', 'addColumnBefore') },
      { label: 'Inserir Colunas à Direita', icon: 'colRight', run: () => app.cmd('table', 'addColumnAfter') },
      { label: 'Inserir Linhas Acima', icon: 'rowAbove', run: () => app.cmd('table', 'addRowBefore') },
      { label: 'Inserir Linhas Abaixo', icon: 'rowBelow', run: () => app.cmd('table', 'addRowAfter') },
    ] });
    items.push({ label: 'Excluir', submenu: [
      { label: 'Excluir Colunas', run: () => app.cmd('table', 'deleteColumn') },
      { label: 'Excluir Linhas', run: () => app.cmd('table', 'deleteRow') },
      { label: 'Excluir Tabela', icon: 'deleteTable', run: () => app.cmd('table', 'deleteTable') },
    ] });
    items.push({ label: 'Mesclar Células', icon: 'merge', run: () => app.cmd('table', 'mergeCells') });
    items.push({ label: 'Dividir Células', icon: 'split', run: () => app.cmd('table', 'splitCell') });
    items.push({ label: 'Propriedades da Tabela...', run: () => app.dialog('tableProperties') });
    items.push({ separator: true });
  }
  items.push({ label: 'Fonte...', icon: 'fontColor', shortcut: 'Ctrl+D', run: () => app.dialog('font') });
  items.push({ label: 'Parágrafo...', icon: 'lineSpacing', run: () => app.dialog('paragraph') });
  if (info.inList) items.push({ label: 'Reiniciar em 1', run: () => app.cmd('restartNumbering') });
  items.push({ separator: true });
  if (!info.link) items.push({ label: 'Link', icon: 'link16', shortcut: 'Ctrl+K', run: () => app.dialog('link') });
  if (info.hasSelection) items.push({ label: 'Pesquisar "' + info.selectionText.slice(0, 24) + (info.selectionText.length > 24 ? '…' : '') + '"', icon: 'find', run: () => { app.showNavPane('search', info.selectionText); } });

  const mini = miniToolbar(app);
  const menu = buildMenu(items, () => p && p.close());
  const wrap = h('div', { class: 'ctx-wrap' }, mini, menu);
  const p = new Popup(wrap, { at: { x: info.x, y: info.y }, className: 'ctx-popup' });
  p.open();
  // a minibarra fica acima do menu quando há espaço
  return p;
}

function miniToolbar(app) {
  const st = app.uiState || {};
  const btn = (ic, run, title, active) => {
    const b = h('button', { class: 'mini-btn' + (active ? ' active' : ''), type: 'button', title, html: icon(ic) });
    b.addEventListener('mousedown', (e) => e.preventDefault());
    b.addEventListener('click', run);
    return b;
  };
  const font = h('input', { class: 'mini-font', type: 'text', value: st.font || '', spellcheck: 'false' });
  font.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); closeAllPopups(); app.cmd('setFont', font.value.trim()); app.focusEditor(); } });
  const size = h('input', { class: 'mini-size', type: 'text', value: st.size || '', spellcheck: 'false' });
  size.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); closeAllPopups(); const v = parseFloat(size.value.replace(',', '.')); if (v > 0) app.cmd('setFontSize', v); app.focusEditor(); } });
  const hl = btn('highlight', () => app.cmd('highlight', app.lastHighlight), 'Cor do Realce do Texto');
  hl.querySelector('svg').insertAdjacentHTML('afterend', '<span class="rb-colorbar" style="background:' + (HIGHLIGHT[app.lastHighlight] || '#ff0') + '"></span>');
  const color = btn('fontColor', () => app.cmd('setColor', app.lastColor), 'Cor da Fonte');
  color.querySelector('svg').insertAdjacentHTML('afterend', '<span class="rb-colorbar" style="background:#' + app.lastColor.val + '"></span>');
  const colorMore = h('button', { class: 'mini-btn narrow', type: 'button', html: icon('chevronDown') });
  colorMore.addEventListener('mousedown', (e) => e.preventDefault());
  colorMore.addEventListener('click', () => openPanel(colorMore, colorPanel(app.model.theme, (c) => app.cmd('setColor', c === 'auto' ? { val: 'auto' } : c)), { keepOthers: true }));
  const hlMore = h('button', { class: 'mini-btn narrow', type: 'button', html: icon('chevronDown') });
  hlMore.addEventListener('mousedown', (e) => e.preventDefault());
  hlMore.addEventListener('click', () => openPanel(hlMore, highlightPanel((v) => app.cmd('highlight', v)), { keepOthers: true }));
  return h('div', { class: 'mini-toolbar' },
    h('div', { class: 'mini-row' }, font, size,
      btn('growFont', () => app.cmd('growFont', 1), 'Aumentar Fonte'),
      btn('shrinkFont', () => app.cmd('growFont', -1), 'Diminuir Fonte'),
      btn('painter', () => app.formatPainter(false), 'Pincel de Formatação'),
      btn('styles', () => app.stylesPane(), 'Estilos')),
    h('div', { class: 'mini-row' },
      btn('bold', () => app.cmd('toggle', 'b'), 'Negrito', st.bold),
      btn('italic', () => app.cmd('toggle', 'i'), 'Itálico', st.italic),
      btn('underline', () => app.cmd('toggle', 'u'), 'Sublinhado', st.underline),
      hl, hlMore, color, colorMore,
      btn('bullets', () => app.cmd('list', 'bullet'), 'Marcadores', st.list === 'bullet'),
      btn('numbering', () => app.cmd('list', 'number'), 'Numeração', st.list === 'number')));
}
