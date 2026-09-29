// Conversão de HTML colado ou importado para atributos do documento.

export function cssLengthToTwip(v) {
  if (v == null) return null;
  const s = String(v).trim();
  const m = /^(-?\d*\.?\d+)(pt|px|in|cm|mm|em|rem|pc)?$/.exec(s);
  if (!m) return null;
  const n = parseFloat(m[1]);
  switch (m[2] || 'px') {
    case 'pt': return Math.round(n * 20);
    case 'px': return Math.round(n * 15);
    case 'in': return Math.round(n * 1440);
    case 'cm': return Math.round((n * 1440) / 2.54);
    case 'mm': return Math.round((n * 144) / 2.54);
    case 'pc': return Math.round(n * 240);
    case 'em': case 'rem': return Math.round(n * 240);
    default: return null;
  }
}

function rgbToHex(v) {
  if (!v) return null;
  const s = v.trim().toLowerCase();
  if (/^#([0-9a-f]{3}|[0-9a-f]{6})$/.test(s)) {
    const h = s.slice(1);
    return (h.length === 3 ? h.split('').map((c) => c + c).join('') : h).toUpperCase();
  }
  const m = /^rgba?\((\d+)[,\s]+(\d+)[,\s]+(\d+)(?:[,\s/]+([\d.]+))?\)$/.exec(s);
  if (m) {
    if (m[4] != null && parseFloat(m[4]) === 0) return null;
    return [m[1], m[2], m[3]].map((x) => Number(x).toString(16).padStart(2, '0')).join('').toUpperCase();
  }
  const named = { black: '000000', white: 'FFFFFF', red: 'FF0000', blue: '0000FF', green: '008000', yellow: 'FFFF00', gray: '808080', grey: '808080', orange: 'FFA500', purple: '800080', navy: '000080', teal: '008080', maroon: '800000', silver: 'C0C0C0' };
  return named[s] || null;
}

export function paraAttrsFromHTML(dom) {
  const a = {};
  const st = dom.style || {};
  const ta = st.textAlign || dom.getAttribute('align');
  if (ta) {
    const map = { left: 'left', center: 'center', right: 'right', justify: 'both', start: 'left', end: 'right' };
    if (map[ta]) a.jc = map[ta];
  }
  const ind = {};
  const ml = cssLengthToTwip(st.marginLeft || st.paddingLeft);
  if (ml) ind.left = String(ml);
  const ti = cssLengthToTwip(st.textIndent);
  if (ti) ind[ti > 0 ? 'firstLine' : 'hanging'] = String(Math.abs(ti));
  if (Object.keys(ind).length) a.ind = ind;
  const listLevel = dom.getAttribute('data-list-level');
  const listKind = dom.getAttribute('data-list-kind');
  if (listKind) {
    const start = dom.getAttribute('data-list-start');
    a.numId = '__' + listKind + ':' + (dom.getAttribute('data-list-id') || '0') + (start ? '@' + start : '');
    a.ilvl = Number(listLevel) || 0;
  }
  if (dom.getAttribute('data-hr')) a.pBdr = { bottom: { val: 'single', sz: '6', space: '1', color: 'auto' } };
  const style = dom.getAttribute('data-style');
  if (style) a.styleId = style;
  return a;
}

export const marksFromStyleHTML = {
  fontFamily(v) {
    if (!v) return false;
    const first = v.split(',')[0].trim().replace(/^["']|["']$/g, '');
    if (!first || /^(inherit|initial|serif|sans-serif|monospace|system-ui|-apple-system|blinkmacsystemfont|ui-sans-serif|ui-serif|ui-monospace|cursive|fantasy|var\()/i.test(first)) return false;
    if (/^segoe ui$/i.test(first)) return false;
    return { a: { ascii: first, hAnsi: first, cs: first } };
  },
  color(v) {
    if (!v || /inherit|initial|currentcolor|windowtext|var\(/i.test(v)) return false;
    const h = rgbToHex(v);
    if (!h || h === '000000') return false;
    return { a: { val: h } };
  },
  fontSize(v) {
    if (!v) return false;
    const kw = { 'xx-small': 7.5, 'x-small': 7.5, small: 10, medium: 12, large: 13.5, 'x-large': 18, 'xx-large': 24 };
    let pt = kw[v];
    if (pt == null) {
      const t = cssLengthToTwip(v);
      if (!t || /em|%/.test(v)) return false;
      pt = t / 20;
    }
    if (!pt || pt < 1 || pt > 1638) return false;
    return { val: Math.round(pt * 2) };
  },
  background(v) {
    if (!v || /transparent|inherit|initial|var\(/i.test(v)) return false;
    const h = rgbToHex(v);
    if (!h || h === 'FFFFFF') return false;
    return { a: { val: 'clear', color: 'auto', fill: h } };
  },
};

// Prepara HTML externo: listas viram parágrafos numerados, pre vira código,
// blockquote vira citação. Assim o parser do ProseMirror entende a estrutura.
export function normalizeExternalHTML(root, opts = {}) {
  const doc = root.ownerDocument;
  const codeStyle = opts.codeStyle || null;
  const quoteStyle = opts.quoteStyle || null;
  let listCounter = 0;

  // Word e outros: remove comentários condicionais e metadados
  for (const el of Array.from(root.querySelectorAll('style,script,meta,link,title,xml,o\\:p'))) el.remove();

  const convertList = (list, level, id) => {
    const kind = list.tagName === 'OL' ? 'number' : 'bullet';
    const out = [];
    for (const li of Array.from(list.children)) {
      if (li.tagName !== 'LI') continue;
      const nested = [];
      const p = doc.createElement('p');
      p.setAttribute('data-list-kind', kind);
      p.setAttribute('data-list-level', String(level));
      p.setAttribute('data-list-id', id);
      if (list.tagName === 'OL' && list.getAttribute('start')) p.setAttribute('data-list-start', list.getAttribute('start'));
      const extra = [];
      let first = true;
      for (const child of Array.from(li.childNodes)) {
        if (child.nodeType === 1 && (child.tagName === 'UL' || child.tagName === 'OL')) {
          nested.push(...convertList(child, level + 1, id));
        } else if (child.nodeType === 1 && /^(P|DIV|H[1-6])$/.test(child.tagName)) {
          if (first) { while (child.firstChild) p.appendChild(child.firstChild); first = false; }
          else extra.push(child);
        } else {
          p.appendChild(child);
          if (child.nodeType === 1 || (child.nodeType === 3 && child.textContent.trim())) first = false;
        }
      }
      const cb = p.querySelector('input[type=checkbox]');
      if (cb) cb.replaceWith(doc.createTextNode(cb.checked ? '☑ ' : '☐ '));
      out.push(p, ...extra, ...nested);
    }
    return out;
  };

  for (const list of Array.from(root.querySelectorAll('ul,ol'))) {
    if (!list.isConnected || list.closest('li')) continue;
    listCounter += 1;
    const ps = convertList(list, 0, String(listCounter));
    list.replaceWith(...ps);
  }

  for (const pre of Array.from(root.querySelectorAll('pre'))) {
    const text = pre.textContent.replace(/\n$/, '');
    const frag = doc.createDocumentFragment();
    for (const line of text.split('\n')) {
      const p = doc.createElement('p');
      if (codeStyle) p.setAttribute('data-style', codeStyle);
      p.textContent = line;
      frag.appendChild(p);
    }
    pre.replaceWith(frag);
  }

  for (const bq of Array.from(root.querySelectorAll('blockquote'))) {
    const blocks = Array.from(bq.querySelectorAll('p,div,h1,h2,h3,h4,h5,h6'));
    if (!blocks.length) {
      const p = doc.createElement('p');
      while (bq.firstChild) p.appendChild(bq.firstChild);
      bq.appendChild(p);
      blocks.push(p);
    }
    for (const b of blocks) if (quoteStyle) b.setAttribute('data-style', quoteStyle);
    bq.replaceWith(...Array.from(bq.childNodes));
  }

  for (const hr of Array.from(root.querySelectorAll('hr'))) {
    const p = doc.createElement('p');
    p.setAttribute('data-hr', '1');
    hr.replaceWith(p);
  }
  return root;
}
