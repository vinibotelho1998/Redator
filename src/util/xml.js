// Utilitários para ler e escrever XML do Office Open XML.

export const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  wStrict: 'http://purl.oclc.org/ooxml/wordprocessingml/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  rStrict: 'http://purl.oclc.org/ooxml/officeDocument/relationships',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  dgm: 'http://schemas.openxmlformats.org/drawingml/2006/diagram',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',
  v: 'urn:schemas-microsoft-com:vml',
  o: 'urn:schemas-microsoft-com:office:office',
  wps: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
  wpg: 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
  wpc: 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
  w14: 'http://schemas.microsoft.com/office/word/2010/wordml',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
  xmlns: 'http://www.w3.org/2000/xmlns/',
};

export const REL = {
  officeDocument: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
  styles: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
  numbering: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering',
  settings: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings',
  theme: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
  fontTable: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable',
  webSettings: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings',
  header: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',
  footer: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer',
  footnotes: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes',
  endnotes: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes',
  comments: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
  image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  hyperlink: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
  coreProps: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
  extProps: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
};

const serializer = new XMLSerializer();

export function parseXml(text) {
  const doc = new DOMParser().parseFromString(text, 'application/xml');
  const err = doc.getElementsByTagName('parsererror')[0];
  if (err) throw new Error('XML inválido: ' + err.textContent.slice(0, 300));
  return doc;
}

export function serialize(node) {
  return serializer.serializeToString(node);
}

export function isWNs(ns) {
  return ns === NS.w || ns === NS.wStrict;
}

export function isW(el, name) {
  return !!el && el.nodeType === 1 && isWNs(el.namespaceURI) && (name == null || el.localName === name);
}

export function isNs(el, ns, name) {
  return !!el && el.nodeType === 1 && el.namespaceURI === ns && (name == null || el.localName === name);
}

// atributo w:nome
export function wa(el, name) {
  if (!el) return null;
  if (el.hasAttributeNS(NS.w, name)) return el.getAttributeNS(NS.w, name);
  if (el.hasAttributeNS(NS.wStrict, name)) return el.getAttributeNS(NS.wStrict, name);
  if (el.hasAttribute(name)) return el.getAttribute(name);
  return null;
}

// atributo r:nome
export function ra(el, name) {
  if (!el) return null;
  return el.getAttributeNS(NS.r, name) || el.getAttributeNS(NS.rStrict, name) || null;
}

// todos os atributos do elemento, com chaves locais para w: e qualificadas para os demais
export function wAttrs(el) {
  const o = {};
  if (!el) return o;
  for (const at of el.attributes) {
    if (at.namespaceURI === NS.xmlns || at.name === 'xmlns' || at.name.startsWith('xmlns:')) continue;
    if (isWNs(at.namespaceURI)) o[at.localName] = at.value;
    else if (!at.namespaceURI) o['@' + at.name] = at.value;
    else o[(at.prefix || 'x') + ':' + at.localName] = at.value;
  }
  return o;
}

export function attrsToXml(obj) {
  if (!obj) return '';
  let s = '';
  for (const k of Object.keys(obj)) {
    const v = obj[k];
    if (v == null) continue;
    let name;
    if (k.startsWith('@')) name = k.slice(1);
    else if (k.includes(':')) name = k;
    else name = 'w:' + k;
    s += ' ' + name + '="' + escAttr(String(v)) + '"';
  }
  return s;
}

export function children(el, name) {
  const out = [];
  if (!el) return out;
  for (let c = el.firstElementChild; c; c = c.nextElementSibling) {
    if (name == null || isW(c, name)) out.push(c);
  }
  return out;
}

export function child(el, name) {
  if (!el) return null;
  for (let c = el.firstElementChild; c; c = c.nextElementSibling) if (isW(c, name)) return c;
  return null;
}

export function childNs(el, ns, name) {
  if (!el) return null;
  for (let c = el.firstElementChild; c; c = c.nextElementSibling) if (c.namespaceURI === ns && c.localName === name) return c;
  return null;
}

// primeiro descendente com o nome local indicado, em qualquer namespace
export function findDeep(el, localName, ns) {
  if (!el) return null;
  const list = ns ? el.getElementsByTagNameNS(ns, localName) : el.getElementsByTagNameNS('*', localName);
  return list.length ? list[0] : null;
}

export function findAllDeep(el, localName, ns) {
  if (!el) return [];
  return Array.from(ns ? el.getElementsByTagNameNS(ns, localName) : el.getElementsByTagNameNS('*', localName));
}

export function onOff(el) {
  if (!el) return null;
  const v = wa(el, 'val');
  if (v == null) return true;
  return !(v === '0' || v === 'false' || v === 'off' || v === 'none');
}

export function intAttr(el, name, def = null) {
  const v = wa(el, name);
  if (v == null || v === '') return def;
  const n = parseInt(v, 10);
  return Number.isFinite(n) ? n : def;
}

// Remove caracteres inválidos em XML 1.0
const INVALID_XML = /[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g;

export function escText(s) {
  return String(s).replace(INVALID_XML, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

export function escAttr(s) {
  return String(s).replace(INVALID_XML, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

export function qnameOf(xml) {
  const m = /^\s*<([A-Za-z_][\w.:-]*)/.exec(xml);
  return m ? m[1] : '';
}

// Ordena fragmentos XML segundo a ordem do esquema
export function orderFragments(frags, order) {
  const idx = (q) => {
    const local = q.includes(':') ? q.slice(q.indexOf(':') + 1) : q;
    const prefix = q.includes(':') ? q.slice(0, q.indexOf(':')) : 'w';
    if (prefix !== 'w') return 10000;
    const i = order.indexOf(local);
    return i < 0 ? 9000 : i;
  };
  return frags
    .map((f, i) => ({ f, i, k: idx(qnameOf(f)) }))
    .sort((a, b) => a.k - b.k || a.i - b.i)
    .map((x) => x.f);
}

// Resolve um alvo de relacionamento relativo à pasta da parte
export function resolvePartPath(basePart, target) {
  if (!target) return null;
  if (/^[a-z]+:/i.test(target)) return target;
  if (target.startsWith('/')) return target.slice(1);
  const dir = basePart.includes('/') ? basePart.slice(0, basePart.lastIndexOf('/') + 1) : '';
  const parts = (dir + target).split('/');
  const out = [];
  for (const p of parts) {
    if (p === '..') out.pop();
    else if (p !== '.' && p !== '') out.push(p);
  }
  return out.join('/');
}

export function relsPathFor(partPath) {
  const i = partPath.lastIndexOf('/');
  const dir = i >= 0 ? partPath.slice(0, i + 1) : '';
  const name = i >= 0 ? partPath.slice(i + 1) : partPath;
  return dir + '_rels/' + name + '.rels';
}

export function relativeTarget(fromPart, toPart) {
  const fromDir = fromPart.includes('/') ? fromPart.slice(0, fromPart.lastIndexOf('/')).split('/') : [];
  const to = toPart.split('/');
  let i = 0;
  while (i < fromDir.length && i < to.length - 1 && fromDir[i] === to[i]) i++;
  const up = fromDir.length - i;
  return '../'.repeat(up) + to.slice(i).join('/');
}

// Texto de um elemento considerando apenas w:t / w:delText / m:t
export function plainText(el) {
  let s = '';
  const walk = (n) => {
    for (let c = n.firstElementChild; c; c = c.nextElementSibling) {
      const ln = c.localName;
      if (ln === 't' || ln === 'delText' || ln === 'instrText') {
        if (ln !== 'instrText') s += c.textContent;
      } else if (ln === 'tab') s += '\t';
      else if (ln === 'br' || ln === 'cr') s += '\n';
      else if (ln === 'p' && s && !s.endsWith('\n')) { walk(c); s += '\n'; continue; }
      else walk(c);
    }
  };
  walk(el);
  return s;
}
