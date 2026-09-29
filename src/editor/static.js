// Renderização estática (cabeçalhos, rodapés, notas, caixas de texto).
import { DOMSerializer, Node as PMNode } from 'prosemirror-model';
import { schema } from './schema.js';
import { drawingDOM } from './nodeviews.js';
import { NumberingCounter, formatNumber } from '../docx/numbering.js';
import { R } from './render.js';
import { labelFor, labelDOM } from './numlabels.js';
import { tableCellClassList } from './plugins/decorations.js';

let ser = null;

function serializer() {
  if (ser) return ser;
  const base = DOMSerializer.fromSchema(schema);
  const nodes = { ...base.nodes };
  nodes.drawing = (node) => drawingDOM(node, { renderBlocks: renderBlocksJSON }).dom;
  nodes.opaque_inline = (node) => {
    const span = document.createElement('span');
    span.className = 'oi oi-' + node.attrs.kind;
    if (node.attrs.kind === 'del' || node.attrs.kind === 'math') span.textContent = node.attrs.text || '';
    return span;
  };
  ser = new DOMSerializer(nodes, base.marks);
  return ser;
}

export function renderBlocksJSON(list) {
  try {
    const nodes = list.map((j) => PMNode.fromJSON(schema, j));
    const doc = schema.node('doc', null, nodes);
    const holder = renderStatic(doc);
    const frag = document.createDocumentFragment();
    while (holder.firstChild) frag.appendChild(holder.firstChild);
    return frag;
  } catch (err) {
    console.warn('Falha ao renderizar caixa de texto', err);
    return null;
  }
}

export function renderStatic(doc) {
  const holder = document.createElement('div');
  holder.className = 'static-doc';
  holder.appendChild(serializer().serializeFragment(doc.content));
  const counter = new NumberingCounter(R.model && R.model.numbering);
  const ps = Array.from(holder.querySelectorAll('p')).filter((p) => !p.closest('.drw-text'));
  const tables = Array.from(holder.querySelectorAll('table')).filter((t) => !t.closest('.drw-text'));
  let pi = 0;
  let ti = 0;
  doc.descendants((node) => {
    if (node.type.name === 'paragraph') {
      const dom = ps[pi++];
      if (dom) {
        const lbl = labelFor(counter, node.attrs);
        if (lbl) dom.insertBefore(labelDOM(lbl), dom.firstChild);
      }
      return false;
    }
    if (node.type.name === 'table') {
      const dom = tables[ti++];
      if (dom) {
        const classes = tableCellClassList(node);
        const rows = dom.tBodies[0] ? Array.from(dom.tBodies[0].rows) : [];
        rows.forEach((tr, r) => {
          Array.from(tr.cells).forEach((td, c) => {
            const cls = classes[r] && classes[r][c];
            if (cls) td.className += ' ' + cls;
          });
        });
        const wrap = document.createElement('div');
        wrap.className = 'tbl-wrap';
        dom.replaceWith(wrap);
        wrap.appendChild(dom);
      }
    }
    return true;
  });
  return holder;
}

// Preenche campos PAGE e NUMPAGES
export function fillFields(root, vars) {
  const begins = Array.from(root.querySelectorAll('span.fc-begin'));
  for (const begin of begins) {
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_TEXT);
    walker.currentNode = begin;
    let instr = '';
    let depth = 0;
    let sep = null;
    const texts = [];
    let n;
    while ((n = walker.nextNode())) {
      if (n.nodeType === 1) {
        if (n.classList.contains('fc-begin')) depth++;
        else if (n.classList.contains('fc-separate') && depth === 0) sep = n;
        else if (n.classList.contains('fc-end')) { if (depth === 0) break; depth--; }
        else if (n.classList.contains('fi') && depth === 0 && !sep) instr += n.getAttribute('data-instr') || '';
      } else if (sep && depth === 0 && n.nodeType === 3 && n.data) {
        texts.push(n);
      }
    }
    const name = instr.trim().split(/\s+/)[0].toUpperCase();
    let value = null;
    if (name === 'PAGE') value = formatNumber(vars.page, vars.fmt || 'decimal');
    else if (name === 'NUMPAGES' || name === 'SECTIONPAGES') value = String(vars.pages);
    else if (name === 'DATE' && !texts.length) value = new Date().toLocaleDateString('pt-BR');
    if (value == null) continue;
    if (texts.length) {
      texts[0].data = value;
      for (let i = 1; i < texts.length; i++) texts[i].data = '';
    } else if (sep) {
      sep.after(document.createTextNode(value));
    }
  }
}
