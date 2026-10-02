// Propriedades de página acompanham o documento no histórico do ProseMirror.
// Referências de cabeçalho/rodapé e XML desconhecido permanecem na seção.
import { child, parseXml, serialize } from '../util/xml.js';

const fields = ['pgSz', 'pgMar', 'titlePg', 'pgNumType'];
const XMLNS = 'http://www.w3.org/2000/xmlns/';
const order = ['headerReference', 'footerReference', 'footnotePr', 'endnotePr', 'type', 'pgSz', 'pgMar', 'paperSrc', 'pgBorders', 'lnNumType', 'pgNumType', 'cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg', 'textDirection', 'bidi', 'rtlGutter', 'docGrid', 'printerSettings'];

export function pageProperties(model) {
  const section = model.section;
  return {
    xml: Object.fromEntries(fields.map((name) => { const el = child(section.el, name); return [name, el ? serialize(el) : null]; })),
    namespaces: Object.fromEntries(fields.map((name) => { const el = child(section.el, name); return [name, el ? Array.from(el.attributes).filter((a) => a.namespaceURI === XMLNS).map((a) => a.name) : []]; })),
    evenAndOdd: !!model.settings.evenAndOdd,
    defaultTabStop: model.settings.defaultTabStop,
  };
}

export function restorePageProperties(model, props) {
  if (!props) return;
  const section = model.section;
  for (const name of fields) {
    const current = child(section.el, name);
    const xml = props.xml[name];
    if (!xml) { if (current) current.remove(); continue; }
    const replacement = section.doc.importNode(parseXml(xml).documentElement, true);
    // XMLSerializer adds namespace declarations when serializing a child alone.
    // Keep only declarations that were present before the edit.
    const originalNamespaces = props.namespaces?.[name] || [];
    for (const attr of Array.from(replacement.attributes)) {
      if (attr.namespaceURI === XMLNS && !originalNamespaces.includes(attr.name)) replacement.removeAttributeNode(attr);
    }
    if (current) current.replaceWith(replacement);
    else {
      const at = order.indexOf(name);
      const next = Array.from(section.el.children).find((el) => order.indexOf(el.localName) > at);
      section.el.insertBefore(replacement, next || null);
    }
  }
  section.read();
  model.settings.evenAndOdd = props.evenAndOdd;
  model.settings.defaultTabStop = props.defaultTabStop;
  model.settings.dirty = true;
}
