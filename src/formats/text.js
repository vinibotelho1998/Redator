// Texto sem formatação.
import { schema } from '../editor/schema.js';
import { R, effectiveNumPr } from '../editor/render.js';
import { NumberingCounter } from '../docx/numbering.js';
import { blankModel, decodeText, encodeText } from './common.js';

const PLAIN_STYLE = '<w:style w:type="paragraph" w:customStyle="1" w:styleId="PlainTextNotepad"><w:name w:val="Texto Simples"/><w:basedOn w:val="Normal"/><w:qFormat/>'
  + '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:style>';

export async function readText(bytes, path) {
  const info = decodeText(bytes);
  const { model } = await blankModel();
  model.format = 'txt';
  model.extra.text = { encoding: info.encoding, bom: info.bom, eol: info.eol };
  model.styles.addStyleXml(PLAIN_STYLE);
  const lines = info.text.split(/\r\n|\n|\r/);
  if (lines.length > 1 && lines[lines.length - 1] === '') { lines.pop(); model.extra.text.trailingNewline = true; }
  const paras = lines.map((line) => {
    const content = [];
    const parts = line.split('\t');
    parts.forEach((p, i) => {
      if (i > 0) content.push(schema.nodes.tab.create());
      const clean = p.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
      if (clean) content.push(schema.text(clean));
    });
    return schema.nodes.paragraph.create({ styleId: 'PlainTextNotepad' }, content);
  });
  const doc = schema.node('doc', null, paras.length ? paras : [schema.nodes.paragraph.create({ styleId: 'PlainTextNotepad' })]);
  return { model, doc };
}

export function docToPlainText(doc, model, opts = {}) {
  const eol = opts.eol || '\n';
  const counter = new NumberingCounter(model.numbering);
  const lines = [];
  const prev = R.model;
  R.model = model;
  const para = (p) => {
    let s = '';
    const np = effectiveNumPr(p.attrs);
    if (np && opts.labels !== false) {
      const lbl = counter.next(np.numId, np.ilvl);
      if (lbl) {
        const kind = model.numbering.kindOf(np.numId, np.ilvl);
        s += '  '.repeat(np.ilvl) + (kind === 'bullet' ? '•' : lbl.text) + ' ';
      }
    }
    p.forEach((c) => {
      if (c.isText) s += c.text;
      else if (c.type.name === 'tab') s += '\t';
      else if (c.type.name === 'hard_break') s += c.attrs.type === 'page' ? '\f' : eol;
      else if (c.type.name === 'sym') s += c.textContent;
      else if (c.type.name === 'opaque_inline' && c.attrs.kind === 'math') s += c.attrs.text || '';
      else if (c.type.name === 'drawing' && c.attrs.content && opts.textBoxes) s += '';
    });
    return s;
  };
  const walk = (parent) => {
    parent.forEach((n) => {
      if (n.type.name === 'paragraph') lines.push(para(n));
      else if (n.type.name === 'table') {
        n.forEach((row) => {
          const cells = [];
          row.forEach((cell) => {
            const parts = [];
            cell.forEach((b) => { if (b.type.name === 'paragraph') parts.push(para(b)); else parts.push(b.textContent); });
            cells.push(parts.join(' '));
          });
          lines.push(cells.join('\t'));
        });
      } else if (n.type.name === 'sdt_block') walk(n);
      else if (n.type.name === 'opaque_block') lines.push(n.attrs.text || '');
    });
  };
  try { walk(doc); } finally { R.model = prev; }
  return lines.join(eol);
}

export function writeText(doc, model) {
  const info = model.extra.text || { encoding: 'utf-8', bom: false, eol: '\r\n' };
  let text = docToPlainText(doc, model, { eol: info.eol || '\r\n' });
  if (info.trailingNewline) text += info.eol || '\r\n';
  return encodeText(text, info);
}
