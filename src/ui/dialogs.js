// Caixas de diálogo no estilo do Word.
import { h, clear } from './dom.js';
import { icon } from './icons.js';
import { colorPanel } from './colorpicker.js';
import { openPanel } from './popup.js';
import { twipToCm, cmToTwip, fmtNumber, parseLength, TWIP_PER_CM } from '../util/units.js';
import { normHex } from '../util/color.js';
import { PAPER_SIZES } from '../model.js';

let dialogCount = 0;

export class Dialog {
  constructor(title, body, buttons = [], opts = {}) {
    this.opts = opts;
    this.result = null;
    this.overlay = h('div', { class: 'dlg-overlay' + (opts.dim ? ' dim' : '') });
    this.box = h('div', { class: 'dlg ' + (opts.className || ''), role: 'dialog', 'aria-label': title });
    const closeBtn = h('button', { class: 'dlg-x', title: 'Fechar', html: icon('close'), onclick: () => this.cancel() });
    this.titleBar = h('div', { class: 'dlg-title' }, h('span', { text: title }), closeBtn);
    this.body = h('div', { class: 'dlg-body' }, body);
    this.footer = h('div', { class: 'dlg-footer' });
    if (opts.footerLeft) this.footer.appendChild(h('div', { class: 'dlg-footer-left' }, opts.footerLeft));
    this.buttons = buttons.map((b) => {
      const el = h('button', { class: 'btn' + (b.primary ? ' primary' : ''), text: b.label, onclick: () => this.press(b) });
      if (b.disabled) el.disabled = true;
      this.footer.appendChild(el);
      return el;
    });
    this.box.append(this.titleBar, this.body, this.footer);
    this.overlay.appendChild(this.box);
    this.onKey = (e) => {
      if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); this.cancel(); }
      else if (e.key === 'Enter' && !e.shiftKey && e.target.tagName !== 'TEXTAREA' && e.target.tagName !== 'BUTTON' && !e.target.closest('.dlg-list')) {
        const primary = buttons.find((b) => b.primary);
        if (primary) { e.preventDefault(); this.press(primary); }
      }
    };
    this.drag();
    this.promise = new Promise((resolve) => { this.resolve = resolve; });
  }

  drag() {
    let sx, sy, ox, oy;
    const move = (e) => {
      this.box.style.left = ox + e.clientX - sx + 'px';
      this.box.style.top = oy + e.clientY - sy + 'px';
    };
    const up = () => { window.removeEventListener('mousemove', move); window.removeEventListener('mouseup', up); };
    this.titleBar.addEventListener('mousedown', (e) => {
      if (e.target.closest('button')) return;
      const r = this.box.getBoundingClientRect();
      sx = e.clientX; sy = e.clientY; ox = r.left; oy = r.top;
      this.box.style.position = 'fixed';
      this.box.style.margin = '0';
      this.box.style.left = ox + 'px';
      this.box.style.top = oy + 'px';
      window.addEventListener('mousemove', move);
      window.addEventListener('mouseup', up);
    });
  }

  open() {
    dialogCount++;
    document.body.appendChild(this.overlay);
    document.addEventListener('keydown', this.onKey, true);
    setTimeout(() => {
      const first = this.box.querySelector('[autofocus]') || this.box.querySelector('input:not([type=checkbox]):not([type=radio]),select,textarea');
      if (first) { first.focus(); if (first.select) first.select(); }
      else if (this.buttons.length) this.buttons[this.buttons.length - 1].focus();
    }, 0);
    return this.promise;
  }

  press(b) {
    if (b.run) {
      const r = b.run(this);
      if (r === false) return;
    }
    this.close(b.value !== undefined ? b.value : b.label);
  }

  cancel() {
    if (this.opts.onCancel) this.opts.onCancel();
    this.close(this.opts.cancelValue !== undefined ? this.opts.cancelValue : null);
  }

  close(value) {
    if (!this.overlay.isConnected) return;
    dialogCount--;
    this.overlay.remove();
    document.removeEventListener('keydown', this.onKey, true);
    this.resolve(value);
    if (this.opts.onClose) this.opts.onClose(value);
  }
}

export const isDialogOpen = () => dialogCount > 0;

// ---------------------------------------------------------------- componentes

export function field(label, control, opts = {}) {
  return h('label', { class: 'fld' + (opts.inline ? ' inline' : '') + (opts.className ? ' ' + opts.className : '') }, h('span', { class: 'fld-label', text: label }), control);
}

export function textInput(value, opts = {}) {
  const el = h('input', { type: 'text', class: 'inp' + (opts.className ? ' ' + opts.className : ''), value: value ?? '', spellcheck: 'false' });
  if (opts.width) el.style.width = opts.width;
  if (opts.placeholder) el.placeholder = opts.placeholder;
  return el;
}

export function select(options, value, opts = {}) {
  const el = h('select', { class: 'sel' });
  for (const o of options) {
    const [v, label] = Array.isArray(o) ? o : [o, o];
    const opt = h('option', { value: v, text: label });
    if (String(v) === String(value)) opt.selected = true;
    el.appendChild(opt);
  }
  if (opts.width) el.style.width = opts.width;
  return el;
}

export function checkbox(label, checked, opts = {}) {
  const input = h('input', { type: 'checkbox' });
  input.checked = !!checked;
  if (opts.indeterminate) input.indeterminate = true;
  const el = h('label', { class: 'chk' }, input, h('span', { text: label }));
  el.input = input;
  return el;
}

// Campo numérico com setas e unidade, como os do Word
export function spinner(value, opts = {}) {
  const unit = opts.unit || 'cm';
  const step = opts.step ?? (unit === 'pt' ? 6 : unit === 'cm' ? 0.1 : 1);
  const min = opts.min ?? (unit === 'pt' ? 0 : -100);
  const max = opts.max ?? 1000;
  const input = h('input', { type: 'text', class: 'inp spin-inp', spellcheck: 'false' });
  const fmt = (v) => (v == null || Number.isNaN(v) ? '' : fmtNumber(v, unit === 'pt' ? 1 : 2) + (opts.suffix != null ? opts.suffix : ' ' + unit));
  input.value = fmt(value);
  const read = () => {
    const s = input.value.trim().toLowerCase().replace(',', '.');
    if (s === '' || s === 'auto' || s === 'automático') return opts.allowAuto ? null : value;
    const m = /^(-?\d*\.?\d+)/.exec(s);
    if (!m) return value;
    return Math.max(min, Math.min(max, parseFloat(m[1])));
  };
  const bump = (d) => {
    const v = read();
    const next = Math.round(((v ?? 0) + d * step) * 100) / 100;
    input.value = fmt(Math.max(min, Math.min(max, next)));
    input.dispatchEvent(new Event('input', { bubbles: true }));
  };
  const up = h('button', { class: 'spin-btn up', tabindex: '-1', type: 'button', onclick: () => bump(1) });
  const down = h('button', { class: 'spin-btn down', tabindex: '-1', type: 'button', onclick: () => bump(-1) });
  input.addEventListener('keydown', (e) => {
    if (e.key === 'ArrowUp') { e.preventDefault(); bump(1); }
    if (e.key === 'ArrowDown') { e.preventDefault(); bump(-1); }
  });
  input.addEventListener('blur', () => { const v = read(); input.value = v == null ? (opts.allowAuto ? 'Auto' : '') : fmt(v); });
  const el = h('span', { class: 'spin' }, input, h('span', { class: 'spin-btns' }, up, down));
  el.input = input;
  el.get = read;
  el.set = (v) => { input.value = fmt(v); };
  if (opts.width) el.style.width = opts.width;
  return el;
}

function tabs(defs) {
  const head = h('div', { class: 'dlg-tabs' });
  const bodies = h('div', { class: 'dlg-tab-bodies' });
  const set = (i) => {
    Array.from(head.children).forEach((c, k) => c.classList.toggle('active', k === i));
    Array.from(bodies.children).forEach((c, k) => { c.hidden = k !== i; });
  };
  defs.forEach((d, i) => {
    head.appendChild(h('button', { class: 'dlg-tab', type: 'button', text: d.label, onclick: () => set(i) }));
    const b = h('div', { class: 'dlg-tab-body' }, d.body);
    bodies.appendChild(b);
  });
  set(defs.findIndex((d) => d.active) >= 0 ? defs.findIndex((d) => d.active) : 0);
  return { el: h('div', { class: 'dlg-tabbed' }, head, bodies), set };
}

function groupBox(title, ...kids) {
  return h('fieldset', { class: 'grp' }, h('legend', { text: title }), ...kids);
}

// ---------------------------------------------------------------- mensagens

export function messageBox(title, text, buttons = [{ label: 'OK', primary: true }], opts = {}) {
  const body = h('div', { class: 'msg' },
    opts.icon ? h('div', { class: 'msg-icon ' + opts.icon }) : null,
    h('div', { class: 'msg-text' }, ...String(text).split('\n').map((l) => h('p', { text: l }))));
  return new Dialog(title, body, buttons, { className: 'dlg-msg', cancelValue: opts.cancelValue }).open();
}

// senha de um PDF protegido; devolve null quando a pessoa cancela
export function passwordDialog(name, retry) {
  const input = textInput('', { width: '100%' });
  input.type = 'password';
  const body = h('div', {},
    h('p', { class: 'dlg-p', text: (retry ? 'Senha incorreta. ' : '') + '"' + name + '" é protegido por senha.' }),
    field('Senha:', input));
  return new Dialog('Senha do Documento', body, [
    { label: 'OK', primary: true, run: (d) => { d.close(input.value); return false; } },
    { label: 'Cancelar', value: null },
  ], { className: 'dlg-link', cancelValue: null }).open();
}

export function confirmSave(name) {
  return messageBox('Redator', 'Deseja salvar as alterações em "' + name + '"?\nSe você não salvar, as alterações serão perdidas.', [
    { label: 'Salvar', primary: true, value: 'save' },
    { label: 'Não Salvar', value: 'discard' },
    { label: 'Cancelar', value: 'cancel' },
  ], { cancelValue: 'cancel', icon: 'warn' });
}

// ---------------------------------------------------------------- Fonte

export function fontDialog(app) {
  const r = app.effectiveRun();
  const fontName = app.currentFontName();
  const size = r.sz ? r.sz.val / 2 : 11;
  const bold = !!(r.b && r.b.val);
  const italic = !!(r.i && r.i.val);
  const fonts = app.fontList();
  const fName = textInput(fontName);
  const fList = h('div', { class: 'dlg-list' });
  const fill = () => {
    clear(fList);
    const q = fName.value.toLowerCase();
    for (const f of fonts) {
      const item = h('div', { class: 'dlg-list-item' + (f.toLowerCase() === q ? ' sel' : ''), text: f, onclick: () => { fName.value = f; fill(); preview(); } });
      fList.appendChild(item);
    }
    const sel = fList.querySelector('.sel');
    if (sel) sel.scrollIntoView({ block: 'nearest' });
  };
  const styleSel = select([['0', 'Regular'], ['1', 'Itálico'], ['2', 'Negrito'], ['3', 'Negrito Itálico']], String((bold ? 2 : 0) + (italic ? 1 : 0)));
  styleSel.size = 4;
  styleSel.classList.add('sel-list');
  const sizeInp = textInput(fmtNumber(size, 1), { width: '64px' });
  const sizeSel = select([8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72].map((s) => [String(s), String(s)]), String(size));
  sizeSel.size = 4;
  sizeSel.classList.add('sel-list');
  sizeSel.addEventListener('change', () => { sizeInp.value = sizeSel.value; preview(); });
  let color = r.color ? r.color.a : null;
  const colorBtn = h('button', { class: 'btn color-btn', type: 'button' }, h('span', { class: 'cb-swatch' }), h('span', { html: icon('chevronDown') }));
  const paintColor = () => { colorBtn.firstChild.style.background = color && color.val && color.val !== 'auto' ? '#' + color.val : '#000'; };
  paintColor();
  colorBtn.addEventListener('click', () => openPanel(colorBtn, colorPanel(app.model.theme, (c) => { color = c === 'auto' ? { val: 'auto' } : c; paintColor(); preview(); }), { keepOthers: true }));
  const uVal = r.u && r.u.a ? r.u.a.val : 'none';
  const uSel = select([['none', '(nenhum)'], ['single', 'Simples'], ['words', 'Somente palavras'], ['double', 'Duplo'], ['thick', 'Grosso'], ['dotted', 'Pontilhado'], ['dash', 'Tracejado'], ['dotDash', 'Traço ponto'], ['wave', 'Ondulado']], uVal);
  const eff = {
    strike: checkbox('Tachado', r.strike && r.strike.val),
    dstrike: checkbox('Tachado duplo', r.dstrike && r.dstrike.val),
    sup: checkbox('Sobrescrito', r.vertAlign && r.vertAlign.val === 'superscript'),
    sub: checkbox('Subscrito', r.vertAlign && r.vertAlign.val === 'subscript'),
    smallCaps: checkbox('Versalete', r.smallCaps && r.smallCaps.val),
    caps: checkbox('Todas em maiúsculas', r.caps && r.caps.val),
    vanish: checkbox('Oculto', r.vanish && r.vanish.val),
  };
  eff.sup.input.addEventListener('change', () => { if (eff.sup.input.checked) eff.sub.input.checked = false; preview(); });
  eff.sub.input.addEventListener('change', () => { if (eff.sub.input.checked) eff.sup.input.checked = false; preview(); });
  const spacingKind = r.spacing && r.spacing.val ? (r.spacing.val > 0 ? 'exp' : 'cond') : 'normal';
  const spacingSel = select([['normal', 'Normal'], ['exp', 'Expandido'], ['cond', 'Condensado']], spacingKind);
  const spacingBy = spinner(r.spacing ? Math.abs(r.spacing.val / 20) : 0, { unit: 'pt', step: 0.1, min: 0, max: 100 });
  const prev = h('div', { class: 'font-preview' }, h('span', { text: app.selectedText().slice(0, 40) || fontName }));
  const preview = () => {
    const s = prev.firstChild.style;
    s.fontFamily = '"' + fName.value + '"';
    s.fontSize = (parseFloat(sizeInp.value.replace(',', '.')) || size) + 'pt';
    const v = Number(styleSel.value);
    s.fontWeight = v >= 2 ? 'bold' : 'normal';
    s.fontStyle = v % 2 ? 'italic' : 'normal';
    s.color = color && color.val && color.val !== 'auto' ? '#' + color.val : '#000';
    const lines = [];
    if (uSel.value !== 'none') lines.push('underline');
    if (eff.strike.input.checked || eff.dstrike.input.checked) lines.push('line-through');
    s.textDecorationLine = lines.join(' ') || 'none';
    s.textDecorationStyle = uSel.value === 'double' || eff.dstrike.input.checked ? 'double' : uSel.value === 'dotted' ? 'dotted' : uSel.value === 'wave' ? 'wavy' : 'solid';
    s.fontVariant = eff.smallCaps.input.checked ? 'small-caps' : 'normal';
    s.textTransform = eff.caps.input.checked ? 'uppercase' : 'none';
    s.verticalAlign = eff.sup.input.checked ? 'super' : eff.sub.input.checked ? 'sub' : 'baseline';
  };
  for (const el of [fName, sizeInp, styleSel, uSel]) el.addEventListener('input', preview);
  for (const c of Object.values(eff)) c.input.addEventListener('change', preview);
  fName.addEventListener('input', fill);
  const fontTab = h('div', { class: 'font-tab' },
    h('div', { class: 'row3' },
      h('div', { class: 'col' }, h('span', { class: 'fld-label', text: 'Fonte:' }), fName, fList),
      h('div', { class: 'col' }, h('span', { class: 'fld-label', text: 'Estilo da fonte:' }), styleSel),
      h('div', { class: 'col small' }, h('span', { class: 'fld-label', text: 'Tamanho:' }), sizeInp, sizeSel)),
    h('div', { class: 'row3' },
      field('Cor da fonte:', colorBtn),
      field('Estilo de sublinhado:', uSel)),
    groupBox('Efeitos', h('div', { class: 'grid2' }, ...Object.values(eff))),
    groupBox('Visualização', prev));
  const advTab = h('div', {}, groupBox('Espaçamento entre caracteres', h('div', { class: 'row3' }, field('Espaçamento:', spacingSel), field('Por:', spacingBy))));
  const t = tabs([{ label: 'Fonte', body: fontTab }, { label: 'Avançado', body: advTab }]);
  setTimeout(() => { fill(); preview(); }, 0);
  const apply = () => {
    const name = fName.value.trim();
    if (name && name !== fontName) app.cmd('setFont', name);
    const sz = parseFloat(sizeInp.value.replace(',', '.'));
    if (sz > 0 && Math.abs(sz - size) > 0.01) app.cmd('setFontSize', sz);
    const v = Number(styleSel.value);
    if ((v >= 2) !== bold) app.cmd('toggle', 'b');
    if ((v % 2 === 1) !== italic) app.cmd('toggle', 'i');
    const newColor = color;
    if (JSON.stringify(newColor) !== JSON.stringify(r.color ? r.color.a : null)) app.cmd('setColor', newColor);
    if (uSel.value !== uVal) app.cmd('setUnderline', uSel.value === 'none' ? (r.u ? 'none' : null) : uSel.value);
    const toggles = [['strike', eff.strike], ['dstrike', eff.dstrike], ['smallCaps', eff.smallCaps], ['caps', eff.caps], ['vanish', eff.vanish]];
    for (const [name2, c] of toggles) if (c.input.checked !== !!(r[name2] && r[name2].val)) app.cmd('toggle', name2);
    const va = eff.sup.input.checked ? 'superscript' : eff.sub.input.checked ? 'subscript' : 'baseline';
    const curVa = r.vertAlign ? r.vertAlign.val : 'baseline';
    if (va !== curVa) app.cmd('setMark', 'vertAlign', va === 'baseline' ? null : { val: va });
    const by = spacingBy.get() || 0;
    const sp = spacingSel.value === 'normal' ? 0 : Math.round(by * 20) * (spacingSel.value === 'cond' ? -1 : 1);
    if (sp !== (r.spacing ? r.spacing.val : 0)) app.cmd('setMark', 'spacing', sp ? { val: sp } : null);
  };
  return new Dialog('Fonte', t.el, [{ label: 'OK', primary: true, run: apply }, { label: 'Cancelar', value: null }], { className: 'dlg-font' }).open();
}

// ---------------------------------------------------------------- Parágrafo

export function paragraphDialog(app, tabIndex = 0) {
  const eff = app.effectivePara();
  const ind = eff.ind || {};
  const sp = eff.spacing || {};
  const jc = eff.jc || 'left';
  const cm = (tw) => twipToCm(Number(tw) || 0);
  const align = select([['left', 'Esquerda'], ['center', 'Centralizado'], ['right', 'Direita'], ['both', 'Justificado']], jc === 'start' ? 'left' : jc === 'distribute' ? 'both' : jc);
  const lvl = select([['', 'Corpo do Texto'], ...[1, 2, 3, 4, 5, 6, 7, 8, 9].map((n) => [String(n - 1), 'Nível ' + n])], eff.outlineLvl != null ? String(eff.outlineLvl) : '');
  const left = spinner(cm(ind.left), { unit: 'cm' });
  const right = spinner(cm(ind.right), { unit: 'cm' });
  const special = select([['none', '(nenhum)'], ['first', 'Primeira linha'], ['hanging', 'Deslocamento']], ind.hanging ? 'hanging' : Number(ind.firstLine) > 0 ? 'first' : 'none');
  const by = spinner(cm(ind.hanging || ind.firstLine || 0), { unit: 'cm', min: 0 });
  const before = spinner((Number(sp.before) || 0) / 20, { unit: 'pt', min: 0 });
  const after = spinner((Number(sp.after) || 0) / 20, { unit: 'pt', min: 0 });
  const rule = sp.lineRule || 'auto';
  const lineVal = Number(sp.line) || 240;
  let lsKind = 'multiple';
  if (rule === 'exact') lsKind = 'exact';
  else if (rule === 'atLeast') lsKind = 'atLeast';
  else if (lineVal === 240) lsKind = 'single';
  else if (lineVal === 360) lsKind = '1.5';
  else if (lineVal === 480) lsKind = 'double';
  const ls = select([['single', 'Simples'], ['1.5', '1,5 linha'], ['double', 'Duplo'], ['atLeast', 'Pelo menos'], ['exact', 'Exatamente'], ['multiple', 'Múltiplos']], lsKind);
  const at = spinner(rule === 'auto' ? lineVal / 240 : lineVal / 20, { unit: rule === 'auto' ? 'li' : 'pt', suffix: rule === 'auto' ? '' : ' pt', step: rule === 'auto' ? 0.5 : 1, min: 0 });
  ls.addEventListener('change', () => {
    if (ls.value === 'atLeast' || ls.value === 'exact') at.set(12);
    else if (ls.value === 'multiple') at.set(3);
    else at.set(null);
    at.input.disabled = !['atLeast', 'exact', 'multiple'].includes(ls.value);
  });
  at.input.disabled = !['atLeast', 'exact', 'multiple'].includes(ls.value);
  const ctx = checkbox('Não adicionar espaço entre parágrafos do mesmo estilo', eff.contextualSpacing);
  const widow = checkbox('Controle de linhas órfãs/viúvas', eff.widowControl !== false);
  const keepNext = checkbox('Manter com o próximo', eff.keepNext);
  const keepLines = checkbox('Manter linhas juntas', eff.keepLines);
  const pbb = checkbox('Quebrar página antes', eff.pageBreakBefore);
  const tab1 = h('div', {},
    groupBox('Geral', h('div', { class: 'row3' }, field('Alinhamento:', align), field('Nível do tópico:', lvl))),
    groupBox('Recuo', h('div', { class: 'row3' }, field('Esquerdo:', left), field('Especial:', special)), h('div', { class: 'row3' }, field('Direito:', right), field('Em:', by))),
    groupBox('Espaçamento', h('div', { class: 'row3' }, field('Antes:', before), field('Espaçamento entre linhas:', ls)), h('div', { class: 'row3' }, field('Depois:', after), field('Em:', at)), ctx));
  const tab2 = h('div', {}, groupBox('Paginação', widow, keepNext, keepLines, pbb));
  const t = tabs([{ label: 'Recuos e espaçamento', body: tab1, active: tabIndex === 0 }, { label: 'Quebras de linha e de página', body: tab2, active: tabIndex === 1 }]);
  const apply = () => {
    const props = {};
    props.jc = align.value;
    const newInd = {};
    newInd.left = String(cmToTwip(left.get() || 0));
    newInd.right = String(cmToTwip(right.get() || 0));
    const byTw = String(cmToTwip(by.get() || 0));
    if (special.value === 'first') newInd.firstLine = byTw;
    else if (special.value === 'hanging') newInd.hanging = byTw;
    else newInd.firstLine = '0';
    props.ind = newInd;
    const spacing = { before: String(Math.round((before.get() || 0) * 20)), after: String(Math.round((after.get() || 0) * 20)) };
    const v = at.get();
    switch (ls.value) {
      case 'single': spacing.line = '240'; spacing.lineRule = 'auto'; break;
      case '1.5': spacing.line = '360'; spacing.lineRule = 'auto'; break;
      case 'double': spacing.line = '480'; spacing.lineRule = 'auto'; break;
      case 'multiple': spacing.line = String(Math.round((v || 1) * 240)); spacing.lineRule = 'auto'; break;
      case 'atLeast': spacing.line = String(Math.round((v || 12) * 20)); spacing.lineRule = 'atLeast'; break;
      case 'exact': spacing.line = String(Math.round((v || 12) * 20)); spacing.lineRule = 'exact'; break;
      default: break;
    }
    props.spacing = spacing;
    props.contextualSpacing = ctx.input.checked;
    props.widowControl = widow.input.checked;
    props.keepNext = keepNext.input.checked;
    props.keepLines = keepLines.input.checked;
    props.pageBreakBefore = pbb.input.checked;
    props.outlineLvl = lvl.value === '' ? null : Number(lvl.value);
    app.cmd('paraProps', props);
  };
  return new Dialog('Parágrafo', t.el, [
    { label: 'Tabulação...', run: () => { apply(); setTimeout(() => tabsDialog(app), 0); } },
    { label: 'OK', primary: true, run: apply }, { label: 'Cancelar', value: null },
  ], { className: 'dlg-para' }).open();
}

// ---------------------------------------------------------------- Tabulação

export function tabsDialog(app) {
  const eff = app.effectivePara();
  let stops = (eff.tabs || []).filter((t) => t.val !== 'clear').map((t) => ({ pos: Number(t.pos), val: t.val, leader: t.leader || 'none' }));
  const posInp = spinner(null, { unit: 'cm', allowAuto: true, min: -50 });
  const list = h('div', { class: 'dlg-list tabs-list' });
  const alignSel = select([['left', 'Esquerdo'], ['center', 'Centralizado'], ['right', 'Direito'], ['decimal', 'Decimal'], ['bar', 'Barra']], 'left');
  const leaderSel = select([['none', '1 Nenhum'], ['dot', '2 .......'], ['hyphen', '3 -------'], ['underscore', '4 ______']], 'none');
  const defStop = spinner(twipToCm(app.model.settings.defaultTabStop || 708), { unit: 'cm', min: 0.1 });
  let selected = -1;
  const render = () => {
    clear(list);
    stops.sort((a, b) => a.pos - b.pos);
    stops.forEach((s, i) => list.appendChild(h('div', {
      class: 'dlg-list-item' + (i === selected ? ' sel' : ''), text: fmtNumber(twipToCm(s.pos), 2) + ' cm',
      onclick: () => { selected = i; posInp.set(twipToCm(s.pos)); alignSel.value = s.val; leaderSel.value = s.leader; render(); },
    })));
  };
  render();
  const set = () => {
    const v = posInp.get();
    if (v == null) return false;
    const pos = cmToTwip(v);
    stops = stops.filter((s) => Math.abs(s.pos - pos) > 5);
    stops.push({ pos, val: alignSel.value, leader: leaderSel.value });
    selected = -1;
    render();
    return false;
  };
  const body = h('div', { class: 'tabs-dlg' },
    h('div', { class: 'row3' },
      h('div', { class: 'col' }, h('span', { class: 'fld-label', text: 'Posição da parada de tabulação:' }), posInp, list),
      h('div', { class: 'col' }, field('Paradas de tabulação padrão:', defStop), groupBox('Alinhamento', alignSel), groupBox('Preenchimento', leaderSel))));
  const apply = () => {
    const original = (eff.tabs || []).filter((t) => t.val !== 'clear');
    const tabsOut = stops.map((s) => ({ val: s.val, ...(s.leader && s.leader !== 'none' ? { leader: s.leader } : {}), pos: String(s.pos) }));
    // paradas removidas do estilo precisam de "clear"
    for (const o of original) if (!stops.some((s) => Math.abs(s.pos - Number(o.pos)) < 5)) tabsOut.push({ val: 'clear', pos: String(o.pos) });
    app.cmd('paraProps', { tabs: tabsOut });
    const d = cmToTwip(defStop.get() || 1.25);
    if (d !== app.model.settings.defaultTabStop) app.setDefaultTabStop(d);
  };
  return new Dialog('Tabulação', body, [
    { label: 'Definir', run: set },
    { label: 'Limpar', run: () => { if (selected >= 0) stops.splice(selected, 1); selected = -1; render(); return false; } },
    { label: 'Limpar Tudo', run: () => { stops = []; render(); return false; } },
    { label: 'OK', primary: true, run: apply }, { label: 'Cancelar', value: null },
  ], { className: 'dlg-tabs' }).open();
}

// ---------------------------------------------------------------- Configurar Página

export function pageSetupDialog(app, tabIndex = 0) {
  const s = app.model.section;
  const m = s.pgMar;
  const cm = (tw) => twipToCm(tw);
  const top = spinner(cm(m.top), { unit: 'cm' });
  const bottom = spinner(cm(m.bottom), { unit: 'cm' });
  const left = spinner(cm(m.left), { unit: 'cm', min: 0 });
  const right = spinner(cm(m.right), { unit: 'cm', min: 0 });
  const gutter = spinner(cm(m.gutter || 0), { unit: 'cm', min: 0 });
  let orient = s.pgSz.w > s.pgSz.h ? 'landscape' : 'portrait';
  const oPortrait = h('button', { type: 'button', class: 'orient-btn' + (orient === 'portrait' ? ' active' : ''), html: icon('onePage') + '<span>Retrato</span>' });
  const oLand = h('button', { type: 'button', class: 'orient-btn land' + (orient === 'landscape' ? ' active' : ''), html: icon('orientation') + '<span>Paisagem</span>' });
  const setO = (o) => {
    if (o !== orient) { const w = pw.get(), hh = ph.get(); pw.set(hh); ph.set(w); }
    orient = o;
    oPortrait.classList.toggle('active', o === 'portrait');
    oLand.classList.toggle('active', o === 'landscape');
  };
  oPortrait.addEventListener('click', () => setO('portrait'));
  oLand.addEventListener('click', () => setO('landscape'));
  const cur = PAPER_SIZES.find((p) => (Math.abs(p.w - s.pgSz.w) < 20 && Math.abs(p.h - s.pgSz.h) < 20) || (Math.abs(p.h - s.pgSz.w) < 20 && Math.abs(p.w - s.pgSz.h) < 20));
  const paper = select([...PAPER_SIZES.map((p) => [p.id, p.name + ' (' + p.desc + ')']), ['custom', 'Tamanho personalizado']], cur ? cur.id : 'custom');
  const pw = spinner(cm(s.pgSz.w), { unit: 'cm', min: 1 });
  const ph = spinner(cm(s.pgSz.h), { unit: 'cm', min: 1 });
  paper.addEventListener('change', () => {
    const p = PAPER_SIZES.find((x) => x.id === paper.value);
    if (!p) return;
    if (orient === 'landscape') { pw.set(cm(p.h)); ph.set(cm(p.w)); } else { pw.set(cm(p.w)); ph.set(cm(p.h)); }
  });
  const hdr = spinner(cm(m.header), { unit: 'cm', min: 0 });
  const ftr = spinner(cm(m.footer), { unit: 'cm', min: 0 });
  const titlePg = checkbox('Diferente na primeira página', s.titlePg);
  const evenOdd = checkbox('Diferentes em páginas pares e ímpares', app.model.settings.evenAndOdd);
  const t = tabs([
    { label: 'Margens', active: tabIndex === 0, body: h('div', {},
      groupBox('Margens', h('div', { class: 'row3' }, field('Superior:', top), field('Inferior:', bottom)), h('div', { class: 'row3' }, field('Esquerda:', left), field('Direita:', right)), h('div', { class: 'row3' }, field('Medianiz:', gutter))),
      groupBox('Orientação', h('div', { class: 'orient-row' }, oPortrait, oLand))) },
    { label: 'Papel', active: tabIndex === 1, body: h('div', {}, groupBox('Tamanho do papel', paper, h('div', { class: 'row3' }, field('Largura:', pw), field('Altura:', ph)))) },
    { label: 'Layout', active: tabIndex === 2, body: h('div', {}, groupBox('Cabeçalhos e rodapés', titlePg, evenOdd, h('div', { class: 'row3' }, field('Da borda - Cabeçalho:', hdr), field('Da borda - Rodapé:', ftr)))) },
  ]);
  const apply = () => {
    app.setPageSetup({
      margins: { top: cmToTwip(top.get()), bottom: cmToTwip(bottom.get()), left: cmToTwip(left.get()), right: cmToTwip(right.get()), gutter: cmToTwip(gutter.get() || 0), header: cmToTwip(hdr.get()), footer: cmToTwip(ftr.get()) },
      size: { w: cmToTwip(pw.get()), h: cmToTwip(ph.get()), orient },
      titlePg: titlePg.input.checked,
      evenOdd: evenOdd.input.checked,
    });
  };
  return new Dialog('Configurar Página', t.el, [{ label: 'OK', primary: true, run: apply }, { label: 'Cancelar', value: null }], { className: 'dlg-page' }).open();
}

// ---------------------------------------------------------------- Localizar e Substituir

export function findReplaceDialog(app, tabIndex = 0) {
  const sel = app.selectedText();
  const q = textInput(sel && sel.length < 80 && !sel.includes('\n') ? sel : app.lastSearch || '', { width: '100%' });
  const q2 = textInput(app.lastSearch || '', { width: '100%' });
  const rep = textInput(app.lastReplace || '', { width: '100%' });
  const goto = textInput('', { width: '100%' });
  const matchCase = checkbox('Diferenciar maiúsculas e minúsculas', app.searchOpts.matchCase);
  const whole = checkbox('Localizar apenas palavras inteiras', app.searchOpts.wholeWord);
  const regex = checkbox('Usar expressões regulares', app.searchOpts.regex);
  const status = h('div', { class: 'find-status' });
  const opts = () => ({ matchCase: matchCase.input.checked, wholeWord: whole.input.checked, regex: regex.input.checked });
  const more = h('div', { class: 'find-more', hidden: true }, groupBox('Opções de pesquisa', matchCase, whole, regex));
  let current = tabIndex;
  const t = tabs([
    { label: 'Localizar', active: tabIndex === 0, body: h('div', {}, field('Localizar:', q)) },
    { label: 'Substituir', active: tabIndex === 1, body: h('div', {}, field('Localizar:', q2), field('Substituir por:', rep)) },
    { label: 'Ir para', active: tabIndex === 2, body: h('div', {}, field('Digite o número da página:', goto)) },
  ]);
  t.el.querySelectorAll('.dlg-tab').forEach((b, i) => b.addEventListener('click', () => {
    current = i;
    if (i === 1) q2.value = q2.value || q.value;
    if (i === 0) q.value = q.value || q2.value;
    updateButtons();
  }));
  const query = () => (current === 1 ? q2.value : q.value);
  const findNext = () => {
    if (current === 2) {
      const n = parseInt(goto.value, 10);
      if (n > 0) app.gotoPage(n - 1);
      return false;
    }
    const n = app.search(query(), opts());
    status.textContent = n ? '' : 'O Redator não encontrou o item pesquisado.';
    if (n) app.findNext(1);
    return false;
  };
  const replaceOne = () => { app.search(q2.value, opts()); app.replaceCurrent(rep.value); app.lastReplace = rep.value; return false; };
  const replaceAllBtn = () => {
    app.search(q2.value, opts());
    const n = app.replaceAll(rep.value);
    app.lastReplace = rep.value;
    status.textContent = n ? 'Concluído. ' + n + (n === 1 ? ' substituição feita.' : ' substituições feitas.') : 'O Redator não encontrou o item pesquisado.';
    return false;
  };
  const dlg = new Dialog('Localizar e Substituir', h('div', {}, t.el, more, status), [
    { label: 'Mais >>', run: (d) => { more.hidden = !more.hidden; d.buttons[0].textContent = more.hidden ? 'Mais >>' : '<< Menos'; return false; } },
    { label: 'Substituir', run: replaceOne },
    { label: 'Substituir Tudo', run: replaceAllBtn },
    { label: 'Localizar Próxima', primary: true, run: findNext },
    { label: 'Cancelar', value: null },
  ], { className: 'dlg-find', onClose: () => { app.searchOpts = opts(); app.lastSearch = query(); } });
  const updateButtons = () => {
    dlg.buttons[1].hidden = current !== 1;
    dlg.buttons[2].hidden = current !== 1;
    dlg.buttons[3].textContent = current === 2 ? 'Ir para' : 'Localizar Próxima';
  };
  updateButtons();
  return dlg.open();
}

// ---------------------------------------------------------------- Inserir Tabela

export function insertTableDialog(app) {
  const cols = spinner(5, { unit: '', suffix: '', step: 1, min: 1, max: 63 });
  const rows = spinner(2, { unit: '', suffix: '', step: 1, min: 1, max: 500 });
  const body = groupBox('Tamanho da tabela', h('div', { class: 'row3' }, field('Número de colunas:', cols)), h('div', { class: 'row3' }, field('Número de linhas:', rows)));
  return new Dialog('Inserir Tabela', body, [
    { label: 'OK', primary: true, run: () => app.cmd('insertTable', Math.round(rows.get()), Math.round(cols.get())) },
    { label: 'Cancelar', value: null },
  ], { className: 'dlg-small' }).open();
}

// ---------------------------------------------------------------- Link

export function linkDialog(app) {
  const cur = app.currentLink();
  const text = textInput(cur ? cur.text : app.selectedText(), { width: '100%' });
  const addr = textInput(cur ? (cur.href || (cur.anchor ? '#' + cur.anchor : '')) : '', { width: '100%', placeholder: 'https://' });
  const headings = app.headingList();
  const list = h('div', { class: 'dlg-list link-list' });
  if (headings.length) {
    for (const hd of headings) list.appendChild(h('div', { class: 'dlg-list-item', style: { paddingLeft: 6 + hd.level * 12 + 'px' }, text: hd.text, onclick: () => { addr.value = '#' + hd.bookmark; if (!text.value) text.value = hd.text; } }));
  } else list.appendChild(h('div', { class: 'dlg-list-empty', text: 'Nenhum título no documento.' }));
  const body = h('div', {}, field('Texto para exibição:', text), field('Endereço:', addr), h('div', { class: 'fld-label', text: 'Ou um local neste documento:' }), list);
  const buttons = [
    { label: 'OK', primary: true, run: () => {
      let v = addr.value.trim();
      if (!v) return false;
      if (!/^(https?:|mailto:|ftp:|file:|#)/i.test(v)) v = /@/.test(v) && !/\//.test(v) ? 'mailto:' + v : 'https://' + v;
      if (v.startsWith('#')) app.ensureBookmarkForHeading(v.slice(1));
      app.cmd('insertLink', v, text.value);
    } },
    { label: 'Cancelar', value: null },
  ];
  if (cur) buttons.unshift({ label: 'Remover Link', run: () => app.cmd('removeLink') });
  return new Dialog(cur ? 'Editar Hiperlink' : 'Inserir Hiperlink', body, buttons, { className: 'dlg-link' }).open();
}

// ---------------------------------------------------------------- Contar Palavras

export function wordCountDialog(app) {
  const st = app.stats(true);
  const row = (k, v) => h('tr', {}, h('td', { text: k }), h('td', { class: 'num', text: v.toLocaleString('pt-BR') }));
  const body = h('div', {}, h('div', { class: 'fld-label', text: 'Estatísticas:' }),
    h('table', { class: 'stats' }, row('Páginas', st.pages), row('Palavras', st.words), row('Caracteres (sem espaços)', st.charsNoSpaces), row('Caracteres (com espaços)', st.chars), row('Parágrafos', st.paragraphs), row('Linhas', st.lines)));
  return new Dialog('Contar Palavras', body, [{ label: 'Fechar', primary: true }], { className: 'dlg-small' }).open();
}

// ---------------------------------------------------------------- Símbolo

const SYMBOL_BLOCKS = [
  ['Latim básico', 0x20, 0x7e], ['Suplemento Latim-1', 0xa1, 0xff], ['Latim estendido-A', 0x100, 0x17f], ['Grego', 0x391, 0x3c9],
  ['Pontuação geral', 0x2010, 0x205e], ['Símbolos de moeda', 0x20a0, 0x20c0], ['Símbolos tipo letra', 0x2100, 0x214f], ['Setas', 0x2190, 0x21ff],
  ['Operadores matemáticos', 0x2200, 0x22ff], ['Técnicos diversos', 0x2300, 0x23ff], ['Formas geométricas', 0x25a0, 0x25ff], ['Símbolos diversos', 0x2600, 0x26ff], ['Dingbats', 0x2700, 0x27bf],
];

export function symbolDialog(app) {
  const blockSel = select(SYMBOL_BLOCKS.map((b, i) => [String(i), b[0]]), '1');
  const grid = h('div', { class: 'sym-grid' });
  const codeInp = textInput('', { width: '80px' });
  const nameLbl = h('span', { class: 'sym-name' });
  let chosen = null;
  const choose = (ch, cell) => {
    chosen = ch;
    grid.querySelectorAll('.sel').forEach((c) => c.classList.remove('sel'));
    if (cell) cell.classList.add('sel');
    codeInp.value = ch.codePointAt(0).toString(16).toUpperCase().padStart(4, '0');
    nameLbl.textContent = 'Caractere Unicode ' + codeInp.value;
  };
  const render = () => {
    clear(grid);
    const [, a, b] = SYMBOL_BLOCKS[Number(blockSel.value)];
    for (let c = a; c <= b; c++) {
      const ch = String.fromCodePoint(c);
      const cell = h('button', { type: 'button', class: 'sym-cell', text: ch, title: 'U+' + c.toString(16).toUpperCase().padStart(4, '0') });
      cell.addEventListener('click', () => choose(ch, cell));
      cell.addEventListener('dblclick', () => { app.insertSymbol(ch); });
      grid.appendChild(cell);
    }
  };
  blockSel.addEventListener('change', render);
  codeInp.addEventListener('change', () => {
    const c = parseInt(codeInp.value, 16);
    if (c > 31) choose(String.fromCodePoint(c), null);
  });
  render();
  const recent = h('div', { class: 'sym-recent' }, ...app.recentSymbols().map((ch) => h('button', { type: 'button', class: 'sym-cell', text: ch, onclick: () => app.insertSymbol(ch) })));
  const body = h('div', {}, field('Subconjunto:', blockSel), grid, h('div', { class: 'fld-label', text: 'Símbolos usados recentemente:' }), recent, h('div', { class: 'row3' }, field('Código do caractere:', codeInp), nameLbl));
  return new Dialog('Símbolo', body, [
    { label: 'Inserir', primary: true, run: () => { if (chosen) app.insertSymbol(chosen); return false; } },
    { label: 'Fechar', value: null },
  ], { className: 'dlg-symbol' }).open();
}

// ---------------------------------------------------------------- Data e Hora

export function dateTimeDialog(app) {
  const d = new Date();
  const two = (n) => String(n).padStart(2, '0');
  const months = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'];
  const days = ['domingo', 'segunda-feira', 'terça-feira', 'quarta-feira', 'quinta-feira', 'sexta-feira', 'sábado'];
  const formats = [
    two(d.getDate()) + '/' + two(d.getMonth() + 1) + '/' + d.getFullYear(),
    days[d.getDay()] + ', ' + d.getDate() + ' de ' + months[d.getMonth()] + ' de ' + d.getFullYear(),
    d.getDate() + ' de ' + months[d.getMonth()] + ' de ' + d.getFullYear(),
    two(d.getDate()) + '/' + two(d.getMonth() + 1) + '/' + String(d.getFullYear()).slice(2),
    d.getFullYear() + '-' + two(d.getMonth() + 1) + '-' + two(d.getDate()),
    two(d.getDate()) + '.' + two(d.getMonth() + 1) + '.' + d.getFullYear(),
    months[d.getMonth()] + ' de ' + d.getFullYear(),
    months[d.getMonth()].slice(0, 3) + '-' + String(d.getFullYear()).slice(2),
    two(d.getDate()) + '/' + two(d.getMonth() + 1) + '/' + d.getFullYear() + ' ' + two(d.getHours()) + ':' + two(d.getMinutes()),
    two(d.getHours()) + ':' + two(d.getMinutes()),
    two(d.getHours()) + ':' + two(d.getMinutes()) + ':' + two(d.getSeconds()),
  ];
  let chosen = formats[0];
  const list = h('div', { class: 'dlg-list' });
  formats.forEach((f, i) => {
    const it = h('div', { class: 'dlg-list-item' + (i === 0 ? ' sel' : ''), text: f });
    it.addEventListener('click', () => { chosen = f; list.querySelectorAll('.sel').forEach((x) => x.classList.remove('sel')); it.classList.add('sel'); });
    it.addEventListener('dblclick', () => { chosen = f; dlg.press(dlg.primary); });
    list.appendChild(it);
  });
  const dlg = new Dialog('Data e Hora', h('div', {}, h('div', { class: 'fld-label', text: 'Formatos disponíveis:' }), list), [
    { label: 'OK', primary: true, run: () => app.cmd('insertText', chosen) },
    { label: 'Cancelar', value: null },
  ], { className: 'dlg-small' });
  dlg.primary = { label: 'OK', primary: true, run: () => app.cmd('insertText', chosen) };
  return dlg.open();
}

// ---------------------------------------------------------------- Zoom

export function zoomDialog(app) {
  const cur = Math.round(app.layout.zoom * 100);
  const opts = [[200, '200%'], [100, '100%'], [75, '75%'], ['width', 'Largura da página'], ['text', 'Largura do texto'], ['page', 'Página inteira']];
  let choice = null;
  const radios = h('div', { class: 'zoom-radios' });
  for (const [v, label] of opts) {
    const r = h('input', { type: 'radio', name: 'zoomopt' });
    if (v === cur) r.checked = true;
    r.addEventListener('change', () => { choice = v; if (typeof v === 'number') pct.set(v); });
    radios.appendChild(h('label', { class: 'chk' }, r, h('span', { text: label })));
  }
  const pct = spinner(cur, { unit: '%', suffix: '%', step: 10, min: 10, max: 500 });
  const body = h('div', {}, groupBox('Zoom para', radios, field('Porcentagem:', pct)));
  return new Dialog('Zoom', body, [
    { label: 'OK', primary: true, run: () => {
      if (choice === 'width') app.zoomTo('width');
      else if (choice === 'text') app.zoomTo('text');
      else if (choice === 'page') app.zoomTo('page');
      else app.setZoom((pct.get() || 100) / 100);
    } },
    { label: 'Cancelar', value: null },
  ], { className: 'dlg-small' }).open();
}

// ---------------------------------------------------------------- Opções

export function optionsDialog(app) {
  const p = app.prefs;
  const spell = checkbox('Verificar ortografia ao digitar', p.spellcheck !== false);
  const smart = checkbox('Substituir aspas retas por aspas inteligentes', p.smartQuotes !== false);
  const auto = checkbox('Formatar automaticamente listas, travessões e símbolos ao digitar', p.autoFormat !== false);
  const recover = checkbox('Salvar informações de AutoRecuperação a cada minuto', p.autoRecover !== false);
  const units = select([['cm', 'Centímetros'], ['in', 'Polegadas'], ['pt', 'Pontos']], p.units || 'cm');
  const author = textInput(p.author || app.userName || '', { width: '240px' });
  const shortcuts = select([['intl', 'Padrão internacional: Ctrl+B negrito, Ctrl+S salvar'], ['pt', 'Word em português: Ctrl+N negrito, Ctrl+B salvar']], p.shortcuts || 'intl');
  // aplicativo padrão: o Windows só deixa a escolha com quem usa, nas configurações
  const defStatus = h('p', { class: 'dlg-p', text: 'Verificando...' });
  const defBtn = h('button', { class: 'btn', type: 'button', text: 'Definir como padrão...', onclick: () => app.host.openDefaultApps && app.host.openDefaultApps() });
  const defaults = groupBox('Aplicativo padrão',
    h('p', { class: 'dlg-p', text: 'Para abrir documentos com dois cliques no Redator, escolha-o nas configurações do Windows.' }), defStatus, defBtn);
  if (app.host.defaultStatus) {
    app.host.defaultStatus().then((st) => {
      const mine = Object.keys(st || {}).filter((e) => st[e]);
      const rest = Object.keys(st || {}).filter((e) => !st[e]);
      defStatus.textContent = !mine.length ? 'O Redator ainda não é o padrão para nenhum tipo de documento.'
        : !rest.length ? 'O Redator já é o padrão para ' + mine.join(', ') + '.'
          : 'Já é o padrão para ' + mine.join(', ') + '. Falta: ' + rest.join(', ') + '.';
    }).catch(() => { defStatus.textContent = ''; });
  } else {
    defStatus.textContent = '';
    defBtn.disabled = true;
  }
  const body = tabs([
    { label: 'Geral', body: h('div', {}, groupBox('Personalizar sua cópia', field('Nome de usuário:', author)), groupBox('Exibição', field('Mostrar medidas em:', units)), groupBox('Teclado', field('Atalhos de teclado:', shortcuts)), defaults) },
    { label: 'Revisão de Texto', body: h('div', {}, groupBox('Ao corrigir a ortografia', spell), groupBox('AutoFormatação ao digitar', smart, auto)) },
    { label: 'Salvar', body: h('div', {}, groupBox('Salvar documentos', recover)) },
  ]);
  return new Dialog('Opções do Redator', body.el, [
    { label: 'OK', primary: true, run: () => app.savePrefs({ spellcheck: spell.input.checked, smartQuotes: smart.input.checked, autoFormat: auto.input.checked, autoRecover: recover.input.checked, units: units.value, author: author.value.trim(), shortcuts: shortcuts.value }) },
    { label: 'Cancelar', value: null },
  ], { className: 'dlg-options' }).open();
}

// ---------------------------------------------------------------- Propriedades da Tabela

export function tablePropertiesDialog(app) {
  const t = app.currentTable();
  if (!t) return Promise.resolve(null);
  const a = t.node.attrs;
  const alignSel = select([['left', 'Esquerda'], ['center', 'Centralizado'], ['right', 'Direita']], a.jc === 'center' ? 'center' : a.jc === 'right' || a.jc === 'end' ? 'right' : 'left');
  const indent = spinner(twipToCm(a.tblInd && a.tblInd.type !== 'pct' ? Number(a.tblInd.w) || 0 : 0), { unit: 'cm' });
  const cell = app.currentCell();
  const vSel = select([['top', 'Superior'], ['center', 'Centralizado'], ['bottom', 'Inferior']], cell && cell.node.attrs.vAlign ? cell.node.attrs.vAlign : 'top');
  const row = app.currentRow();
  const rowH = spinner(row && row.node.attrs.height ? twipToCm(row.node.attrs.height) : null, { unit: 'cm', allowAuto: true, min: 0 });
  const header = checkbox('Repetir como linha de cabeçalho no início de cada página', row && row.node.attrs.header);
  const colW = spinner(app.currentColumnWidthCm(), { unit: 'cm', min: 0.2 });
  const tb = tabs([
    { label: 'Tabela', body: h('div', {}, groupBox('Alinhamento', alignSel, field('Recuo da esquerda:', indent))) },
    { label: 'Linha', body: h('div', {}, groupBox('Tamanho', field('Altura:', rowH)), groupBox('Opções', header)) },
    { label: 'Coluna', body: h('div', {}, groupBox('Tamanho', field('Largura preferencial:', colW))) },
    { label: 'Célula', body: h('div', {}, groupBox('Alinhamento vertical', vSel)) },
  ]);
  return new Dialog('Propriedades da Tabela', tb.el, [
    { label: 'OK', primary: true, run: () => {
      app.cmd('tableAttrs', { jc: alignSel.value === 'left' ? null : alignSel.value, tblInd: { w: String(cmToTwip(indent.get() || 0)), type: 'dxa' } });
      app.cmd('cellAttrs', { vAlign: vSel.value === 'top' ? null : vSel.value });
      app.setRowProps({ height: rowH.get() ? cmToTwip(rowH.get()) : null, hRule: rowH.get() ? 'atLeast' : null, header: header.input.checked });
      const w = colW.get();
      if (w) app.setColumnWidthCm(w);
    } },
    { label: 'Cancelar', value: null },
  ], { className: 'dlg-table' }).open();
}

// ---------------------------------------------------------------- Imagem

export function imageSizeDialog(app) {
  const img = app.selectedDrawing();
  if (!img) return Promise.resolve(null);
  const a = img.node.attrs;
  const wcm = (a.cx / 360000);
  const hcm = (a.cy / 360000);
  const ratio = hcm / wcm || 1;
  const w = spinner(wcm, { unit: 'cm', min: 0.1 });
  const hh = spinner(hcm, { unit: 'cm', min: 0.1 });
  const lock = checkbox('Bloquear taxa de proporção', true);
  w.input.addEventListener('input', () => { if (lock.input.checked) hh.set((w.get() || wcm) * ratio); });
  hh.input.addEventListener('input', () => { if (lock.input.checked) w.set((hh.get() || hcm) / ratio); });
  const alt = h('textarea', { class: 'inp', rows: '3' });
  alt.value = a.alt || '';
  const body = h('div', {}, groupBox('Tamanho', h('div', { class: 'row3' }, field('Altura:', hh), field('Largura:', w)), lock), field('Texto alternativo:', alt));
  return new Dialog('Layout', body, [
    { label: 'OK', primary: true, run: () => app.updateDrawing({ cx: Math.round((w.get() || wcm) * 360000), cy: Math.round((hh.get() || hcm) * 360000), alt: alt.value }) },
    { label: 'Cancelar', value: null },
  ], { className: 'dlg-small' }).open();
}

export { normHex, TWIP_PER_CM, parseLength };
