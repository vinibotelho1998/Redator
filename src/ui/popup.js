// Menus suspensos, submenus e painéis flutuantes.
import { h } from './dom.js';
import { icon } from './icons.js';

let openStack = [];
let popupId = 0;

export function closeAllPopups() {
  for (const p of openStack.slice().reverse()) p.close();
  openStack = [];
}

export class Popup {
  constructor(content, opts = {}) {
    this.opts = opts;
    this.el = h('div', { class: 'popup ' + (opts.className || ''), role: opts.role || 'menu', id: 'popup-' + (++popupId), tabindex: '-1' });
    this.el.appendChild(content);
    this.onDown = (e) => {
      if (this.el.contains(e.target)) return;
      if (opts.anchor && opts.anchor.contains(e.target)) return;
      for (const p of openStack) if (p !== this && p.el.contains(e.target)) return;
      this.close();
    };
    this.onKey = (e) => {
      if (openStack[openStack.length - 1] !== this) return;
      if (e.key === 'Escape' || (e.key === 'ArrowLeft' && opts.side === 'right')) {
        e.preventDefault(); e.stopImmediatePropagation(); this.close();
        const back = opts.returnFocus || opts.anchor || this.returnFocus;
        if (back && back.isConnected) back.focus();
        return;
      }
      if (e.key === 'Tab') {
        const root = openStack[0];
        const back = root?.opts.returnFocus || root?.opts.anchor || root?.returnFocus;
        closeAllPopups();
        if (back?.isConnected) back.focus();
        return;
      }
      if (e.target.closest('input,textarea,select')) return;
      const items = this.items();
      const at = items.indexOf(document.activeElement);
      if (['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(e.key)) {
        e.preventDefault(); e.stopPropagation();
        const next = e.key === 'Home' ? 0 : e.key === 'End' ? items.length - 1 : at < 0 ? (e.key === 'ArrowUp' ? items.length - 1 : 0) : (at + (e.key === 'ArrowUp' ? -1 : 1) + items.length) % items.length;
        items[next]?.focus();
      } else if (e.key === 'ArrowRight' && document.activeElement._openSub) {
        e.preventDefault(); e.stopPropagation(); document.activeElement._openSub(true);
      }
    };
  }

  items() {
    return [...this.el.querySelectorAll('button, [role="menuitem"], [role="menuitemcheckbox"], input, select')]
      .filter((el) => !el.disabled && el.getAttribute('aria-disabled') !== 'true' && el.getClientRects().length && !el.closest('[hidden]'));
  }

  focusFirst() { (this.items()[0] || this.el).focus(); }

  open() {
    this.returnFocus = document.activeElement;
    if (!this.opts.keepOthers) closeAllPopups();
    document.body.appendChild(this.el);
    this.position();
    openStack.push(this);
    // The triggering event has already passed document capture; subsequent keys
    // must work immediately, including Escape just after ArrowDown.
    document.addEventListener('mousedown', this.onDown, true);
    document.addEventListener('keydown', this.onKey, true);
    // Legacy panel actions also get native keyboard activation semantics.
    for (const item of this.el.querySelectorAll('div.menu-item, div.cp-row-btn, div.font-item')) {
      if (!item.hasAttribute('role')) item.setAttribute('role', 'menuitem');
      item.tabIndex = 0;
      item.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); e.stopPropagation(); item.click(); }
      });
    }
    if (this.opts.anchor) {
      this.opts.anchor.classList.add('open');
      this.opts.anchor.setAttribute('aria-expanded', 'true');
      this.opts.anchor.setAttribute('aria-controls', this.el.id);
    }
    if (this.opts.keyboard || (this.opts.anchor && document.activeElement === this.opts.anchor)) this.focusFirst();
    return this;
  }

  position() {
    const o = this.opts;
    const el = this.el;
    el.style.left = '0px';
    el.style.top = '0px';
    const r = el.getBoundingClientRect();
    let x, y;
    if (o.at) { x = o.at.x; y = o.at.y; }
    else if (o.anchor) {
      const a = o.anchor.getBoundingClientRect();
      if (o.side === 'right') { x = a.right - 2; y = a.top - 4; }
      else { x = o.alignRight ? a.right - r.width : a.left; y = a.bottom + 1; }
    } else { x = 100; y = 100; }
    const vw = window.innerWidth, vh = window.innerHeight;
    if (x + r.width > vw - 4) x = o.side === 'right' && o.anchor ? o.anchor.getBoundingClientRect().left - r.width + 2 : vw - r.width - 4;
    if (y + r.height > vh - 4) {
      if (o.anchor && !o.at && o.side !== 'right') {
        const a = o.anchor.getBoundingClientRect();
        y = a.top - r.height - 1 > 4 ? a.top - r.height - 1 : Math.max(4, vh - r.height - 4);
      } else y = Math.max(4, vh - r.height - 4);
    }
    el.style.left = Math.max(4, x) + 'px';
    el.style.top = Math.max(4, y) + 'px';
    if (r.height > vh - 8) { el.style.maxHeight = vh - 8 + 'px'; el.style.overflowY = 'auto'; }
  }

  close() {
    if (!this.el.isConnected) return;
    // Closing a parent must also remove any submenu attached to it.
    for (const child of openStack.slice().reverse()) {
      if (child !== this && child.opts.anchor && this.el.contains(child.opts.anchor)) child.close();
    }
    this.el.remove();
    document.removeEventListener('mousedown', this.onDown, true);
    document.removeEventListener('keydown', this.onKey, true);
    openStack = openStack.filter((p) => p !== this);
    if (this.opts.anchor) {
      this.opts.anchor.classList.remove('open');
      this.opts.anchor.setAttribute('aria-expanded', 'false');
      this.opts.anchor.removeAttribute('aria-controls');
    }
    if (this.opts.onClose) this.opts.onClose();
  }
}

// items: {label, icon, shortcut, checked, disabled, run, submenu, separator, header, className, html}
export function buildMenu(items, close) {
  const list = h('div', { class: 'menu' });
  for (const it of items) {
    if (!it) continue;
    if (it.separator) { list.appendChild(h('div', { class: 'menu-sep' })); continue; }
    if (it.header) { list.appendChild(h('div', { class: 'menu-header', text: it.header })); continue; }
    if (it.element) { list.appendChild(it.element); continue; }
    const row = h('button', {
      class: 'menu-item' + (it.disabled ? ' disabled' : '') + (it.checked ? ' checked' : '') + (it.className ? ' ' + it.className : '') + (it.bold ? ' bold' : ''),
      type: 'button', role: it.checked !== undefined ? 'menuitemcheckbox' : 'menuitem', tabindex: '0',
      disabled: !!it.disabled, 'aria-disabled': String(!!it.disabled),
      'aria-checked': it.checked !== undefined ? String(!!it.checked) : null,
      'aria-haspopup': it.submenu ? 'menu' : null, 'aria-expanded': it.submenu ? 'false' : null,
    },
    h('span', { class: 'mi-icon', html: it.checked && !it.icon ? icon('check') : it.icon ? icon(it.icon) || '' : '' }),
    h('span', { class: 'mi-label', text: it.label, style: it.labelStyle || null }),
    it.shortcut ? h('span', { class: 'mi-short', text: it.shortcut }) : null,
    it.submenu ? h('span', { class: 'mi-sub', html: icon('chevronRight') }) : null);
    if (it.title) row.title = it.title;
    if (!it.disabled) {
      if (it.submenu) {
        let sub = null;
        const openSub = (keyboard = false) => {
          if (sub && sub.el.isConnected) { if (keyboard) sub.focusFirst(); return; }
          for (const other of list.querySelectorAll('.menu-item.sub-open')) other.classList.remove('sub-open');
          row.classList.add('sub-open');
          const items2 = typeof it.submenu === 'function' ? it.submenu() : it.submenu;
          sub = new Popup(buildMenu(items2, close), { anchor: row, side: 'right', keepOthers: true, className: 'submenu', keyboard }).open();
          list._sub = sub;
        };
        row.addEventListener('mouseenter', () => { if (list._sub && list._sub !== sub) list._sub.close(); openSub(); });
        row._openSub = openSub;
        row.addEventListener('click', (e) => openSub(e.detail === 0));
      } else {
        row.addEventListener('mouseenter', () => { if (list._sub) { list._sub.close(); list._sub = null; } });
        row.addEventListener('click', (e) => {
          e.preventDefault();
          if (close) close();
          closeAllPopups();
          if (it.run) it.run();
        });
      }
    }
    list.appendChild(row);
  }
  return list;
}

export function openMenu(anchor, items, opts = {}) {
  let p;
  const close = () => p && p.close();
  p = new Popup(buildMenu(items, close), { anchor, ...opts });
  return p.open();
}

export function openPanel(anchor, element, opts = {}) {
  return new Popup(element, { anchor, ...opts }).open();
}

export function tooltip(el, text, shortcut, desc) {
  el.setAttribute('aria-label', text);
  if (desc) el.setAttribute('aria-description', desc);
  el.tooltip = { text, shortcut, desc };
  let t = null;
  let tip = null;
  const show = () => {
    if (tip || !el.isConnected) return;
    const r = el.getBoundingClientRect();
    tip = h('div', { class: 'tooltip' },
      h('div', { class: 'tt-title' }, text, el.tooltip.shortcut ? h('span', { class: 'tt-short', text: ' (' + el.tooltip.shortcut + ')' }) : null),
      desc ? h('div', { class: 'tt-desc', text: desc }) : null);
    document.body.appendChild(tip);
    const tr = tip.getBoundingClientRect();
    let x = r.left;
    let y = r.bottom + 6;
    if (x + tr.width > window.innerWidth - 4) x = window.innerWidth - tr.width - 4;
    if (y + tr.height > window.innerHeight - 4) y = r.top - tr.height - 6;
    tip.style.left = x + 'px';
    tip.style.top = y + 'px';
  };
  const hide = () => { clearTimeout(t); if (tip) { tip.remove(); tip = null; } };
  el.addEventListener('mouseenter', () => { t = setTimeout(show, 550); });
  el.addEventListener('focus', () => { t = setTimeout(show, 550); });
  el.addEventListener('blur', hide);
  el.addEventListener('mouseleave', hide);
  el.addEventListener('mousedown', hide);
}
