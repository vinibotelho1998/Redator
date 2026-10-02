// Recursos do clipboard: IDs de estilos/listas só fazem sentido dentro do pacote
// de origem. O HTML continua legível por outros editores, com metadados adicionais.
import { parseXml, serialize, child, children, wa, NS } from '../util/xml.js';
import { Theme } from '../docx/theme.js';
import { PPR_ORDER, RPR_ORDER } from '../docx/props.js';

export const CLIPBOARD_DATA = 'data-redator-clipboard';

export function clipboardResources(model) {
  return {
    version: 1, source: model.clipboardId,
    styles: model.styles.toXml(), numbering: model.numbering.toXml(),
    defaultPara: model.styles.defaultPara, defaultTable: model.styles.defaultTable,
    theme: { ...model.theme, xml: null },
  };
}

function readJSON(el, name) {
  try { return JSON.parse(el.getAttribute(name)); } catch { return null; }
}

// Congela cores/fontes do tema de origem antes de incorporar os estilos.
function freezeTheme(root, theme) {
  for (const el of [root, ...root.getElementsByTagNameNS(NS.w, '*')]) {
    for (const [ref, value, tint, shade] of [
      ['themeColor', el.localName === 'color' ? 'val' : 'color', 'themeTint', 'themeShade'],
      ['themeFill', 'fill', 'themeFillTint', 'themeFillShade'],
    ]) {
      const name = wa(el, ref);
      if (!name) continue;
      const color = theme.hex(name, wa(el, tint), wa(el, shade));
      if (color) el.setAttributeNS(NS.w, 'w:' + value, color);
      for (const key of [ref, tint, shade]) el.removeAttributeNS(NS.w, key);
    }
    if (el.localName === 'rFonts') {
      for (const key of ['ascii', 'hAnsi', 'eastAsia', 'cs']) {
        const ref = wa(el, key + 'Theme');
        if (!ref) continue;
        el.setAttributeNS(NS.w, 'w:' + key, theme.font(ref));
        el.removeAttributeNS(NS.w, key + 'Theme');
      }
    }
  }
}

export function importClipboardResources(body, model) {
  const marker = body.querySelector('[' + CLIPBOARD_DATA + ']');
  const packet = marker && readJSON(marker, CLIPBOARD_DATA);
  if (!packet || packet.version !== 1 || packet.source === model.clipboardId) return false;
  let sourceStyles, sourceNumbers;
  try {
    sourceStyles = packet.styles ? parseXml(packet.styles).documentElement : null;
    sourceNumbers = packet.numbering ? parseXml(packet.numbering).documentElement : null;
  } catch { return false; }
  const styles = new Map(children(sourceStyles, 'style').map((el) => [wa(el, 'styleId'), el]));
  const nums = new Map(children(sourceNumbers, 'num').map((el) => [wa(el, 'numId'), el]));
  const abstracts = new Map(children(sourceNumbers, 'abstractNum').map((el) => [wa(el, 'abstractNumId'), el]));
  const styleIds = new Map(), numIds = new Map(), abstractIds = new Map();
  const theme = Object.assign(new Theme(), packet.theme || {});
  let styleSerial = 0;
  const defaults = child(sourceStyles, 'docDefaults');
  const numbering = model.numbering;

  const remapXML = (el) => {
    freezeTheme(el, theme);
    for (const ref of [...el.getElementsByTagNameNS(NS.w, '*')]) {
      if (['basedOn', 'next', 'link', 'pStyle', 'rStyle', 'tblStyle', 'numStyleLink', 'styleLink'].includes(ref.localName)) {
        const id = importStyle(wa(ref, 'val'));
        if (id) ref.setAttributeNS(NS.w, 'w:val', id); else ref.remove();
      } else if (ref.localName === 'numId') {
        const id = importNumber(wa(ref, 'val'));
        if (id != null) ref.setAttributeNS(NS.w, 'w:val', id); else ref.remove();
      }
    }
  };

  const importStyle = (id) => {
    if (!id) return null;
    if (styleIds.has(id)) return styleIds.get(id);
    const source = styles.get(id);
    if (!source) return null;
    let next;
    do { next = 'Paste' + (++styleSerial) + '_' + id.replace(/[^A-Za-z0-9_]/g, '_'); }
    while (model.styles.get(next) || [...styleIds.values()].includes(next));
    styleIds.set(id, next); // reserva antes de percorrer referências cíclicas
    const el = source.cloneNode(true);
    el.setAttributeNS(NS.w, 'w:styleId', next);
    el.removeAttributeNS(NS.w, 'default');
    // Estilos sem ancestral também precisam dos padrões do documento de origem.
    if (!child(el, 'basedOn') && wa(el, 'type') === 'paragraph') {
      for (const name of ['pPr', 'rPr']) {
        const def = child(child(defaults, name === 'pPr' ? 'pPrDefault' : 'rPrDefault'), name);
        if (!def) continue;
        let props = child(el, name);
        if (!props) {
          props = el.ownerDocument.createElementNS(NS.w, 'w:' + name);
          el.insertBefore(props, name === 'pPr' ? child(el, 'rPr') : null);
        }
        for (const prop of children(def)) if (!child(props, prop.localName)) props.appendChild(prop.cloneNode(true));
        const order = name === 'pPr' ? PPR_ORDER : RPR_ORDER;
        const index = (n) => { const i = order.indexOf(n.localName); return i < 0 ? order.length : i; };
        for (const prop of children(props).sort((a, b) => index(a) - index(b))) props.appendChild(prop);
      }
    }
    remapXML(el);
    model.styles.addStyleXml(serialize(el));
    return next;
  };

  const importAbstract = (id) => {
    if (abstractIds.has(id)) return abstractIds.get(id);
    const source = abstracts.get(id);
    if (!source) return null;
    numbering.ensureDoc();
    const next = String(numbering.nextAbstractId());
    abstractIds.set(id, next);
    numbering.abstracts.set(next, { id: next, levels: [] });
    const el = numbering.doc.importNode(source, true);
    el.setAttributeNS(NS.w, 'w:abstractNumId', next);
    remapXML(el);
    // Marcadores com imagem dependem de outras partes do pacote. Não reutilizar
    // o ID de uma imagem de marcador que por acaso exista no documento destino.
    for (const ref of [...el.getElementsByTagNameNS(NS.w, 'lvlPicBulletId')]) ref.remove();
    const root = numbering.doc.documentElement;
    root.insertBefore(el, children(root, 'num')[0] || child(root, 'numIdMacAtCleanup'));
    numbering.addAbstractEl(el);
    return next;
  };

  const importNumber = (id) => {
    if (id === '0') return id;
    if (numIds.has(id)) return numIds.get(id);
    const source = nums.get(id);
    if (!source) return null;
    numbering.ensureDoc();
    const next = String(numbering.nextNumId());
    numIds.set(id, next);
    numbering.nums.set(next, { id: next, abstractNumId: null, overrides: new Map() });
    const absId = importAbstract(wa(child(source, 'abstractNumId'), 'val'));
    if (absId == null) { numbering.nums.delete(next); numIds.set(id, null); return null; }
    const el = numbering.doc.importNode(source, true);
    el.setAttributeNS(NS.w, 'w:numId', next);
    child(el, 'abstractNumId').setAttributeNS(NS.w, 'w:val', absId);
    remapXML(el);
    const root = numbering.doc.documentElement;
    root.insertBefore(el, child(root, 'numIdMacAtCleanup'));
    numbering.addNumEl(el);
    numbering.dirty = true;
    return next;
  };

  const remapAttrs = (attrs, type) => {
    if (!attrs || typeof attrs !== 'object') return attrs;
    const a = { ...attrs };
    if ('styleId' in a || type === 'paragraph' || type === 'table') {
      const fallback = type === 'paragraph' ? packet.defaultPara : type === 'table' ? packet.defaultTable : null;
      a.styleId = importStyle(a.styleId || fallback);
    }
    if (a.numId != null && !String(a.numId).startsWith('__')) a.numId = importNumber(String(a.numId));
    // Quebras de seção trazem IDs de cabeçalho/rodapé de outro pacote.
    if (a.sectPr) a.sectPr = null;
    return a;
  };

  for (const [tag, attr, type] of [['p', 'data-pa', 'paragraph'], ['table', 'data-ta', 'table']]) {
    for (const el of body.querySelectorAll(tag + '[' + attr + ']')) {
      el.setAttribute(attr, JSON.stringify(remapAttrs(readJSON(el, attr), type)));
    }
  }
  for (const el of body.querySelectorAll('[data-m="rStyle"]')) {
    const id = importStyle(el.getAttribute('data-v'));
    if (id) el.setAttribute('data-v', id); else { el.removeAttribute('data-m'); el.removeAttribute('data-v'); }
  }
  // O ProseMirror pode retirar os nós externos do HTML e guardá-los neste contexto.
  for (const el of body.querySelectorAll('[data-pm-slice]')) {
    const match = /^(\d+ \d+(?: -\d+)? )([\s\S]*)$/.exec(el.getAttribute('data-pm-slice'));
    if (!match) continue;
    try {
      const context = JSON.parse(match[2]);
      for (let i = 0; i < context.length; i += 2) context[i + 1] = remapAttrs(context[i + 1] || {}, context[i]);
      el.setAttribute('data-pm-slice', match[1] + JSON.stringify(context));
    } catch { /* contexto de outro editor */ }
  }
  marker.removeAttribute(CLIPBOARD_DATA);
  return styleIds.size > 0 || numIds.size > 0;
}
