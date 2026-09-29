// Decorações derivadas do conteúdo: numeração, espaçamento contextual e classes de células.
import { Plugin, PluginKey } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import { TableMap } from 'prosemirror-tables';
import { NumberingCounter } from '../../docx/numbering.js';
import { R, hasVisibleBorder } from '../render.js';
import { labelFor, labelDOM, labelKey } from '../numlabels.js';

export const decoKey = new PluginKey('docDecorations');

function lookFlags(look) {
  const f = { firstRow: false, lastRow: false, firstColumn: false, lastColumn: false, noHBand: false, noVBand: false };
  if (!look) return { ...f, firstRow: true, firstColumn: true, noVBand: true };
  const on = (v) => v === '1' || v === 'true' || v === 'on';
  if (look.val != null && look.firstRow == null) {
    const v = parseInt(look.val, 16) || 0;
    return { firstRow: !!(v & 0x20), lastRow: !!(v & 0x40), firstColumn: !!(v & 0x80), lastColumn: !!(v & 0x100), noHBand: !!(v & 0x200), noVBand: !!(v & 0x400) };
  }
  for (const k of Object.keys(f)) f[k] = on(look[k]);
  return f;
}

// Classes de cada célula: [linha][índice da célula na linha]
export function tableCellClassList(table) {
  let map;
  try { map = TableMap.get(table); } catch { return []; }
  const look = lookFlags(table.attrs.tblLook);
  const out = [];
  let offset = 0;
  const rowStarts = [];
  table.forEach((row, rOff) => { rowStarts.push(rOff); });
  table.forEach((row, rowOffset, r) => {
    const list = [];
    row.forEach((cell, cellOffset) => {
      const pos = rowOffset + 1 + cellOffset;
      let rect;
      try { rect = map.findCell(pos); } catch { rect = { left: 0, right: 1, top: r, bottom: r + 1 }; }
      const cls = [];
      if (rect.top === 0) cls.push('r-first');
      if (rect.bottom === map.height) cls.push('r-last');
      if (rect.left === 0) cls.push('c-first');
      if (rect.right === map.width) cls.push('c-last');
      const isFirstRow = look.firstRow && rect.top === 0;
      const isLastRow = look.lastRow && rect.bottom === map.height;
      const isFirstCol = look.firstColumn && rect.left === 0;
      const isLastCol = look.lastColumn && rect.right === map.width;
      if (!look.noHBand && !isFirstRow && !isLastRow) {
        const idx = rect.top - (look.firstRow ? 1 : 0);
        cls.push(idx % 2 === 0 ? 'cf-band1Horz' : 'cf-band2Horz');
      }
      if (!look.noVBand && !isFirstCol && !isLastCol) {
        const idx = rect.left - (look.firstColumn ? 1 : 0);
        cls.push(idx % 2 === 0 ? 'cf-band1Vert' : 'cf-band2Vert');
      }
      if (isFirstCol) cls.push('cf-firstCol');
      if (isLastCol) cls.push('cf-lastCol');
      if (isFirstRow) cls.push('cf-firstRow');
      if (isLastRow) cls.push('cf-lastRow');
      if (isFirstRow && isFirstCol) cls.push('cf-nwCell');
      if (isFirstRow && isLastCol) cls.push('cf-neCell');
      if (isLastRow && isFirstCol) cls.push('cf-swCell');
      if (isLastRow && isLastCol) cls.push('cf-seCell');
      list.push(cls.join(' '));
    });
    out.push(list);
  });
  offset += 0;
  return out;
}

function effContextual(attrs) {
  if (attrs.contextualSpacing != null) return attrs.contextualSpacing;
  const st = R.model && R.model.styles;
  if (!st) return false;
  return !!st.resolvePara(attrs.styleId || st.defaultPara).pPr.contextualSpacing;
}

function styleOf(attrs) {
  const st = R.model && R.model.styles;
  return attrs.styleId || (st ? st.defaultPara : 'Normal');
}

function isBordered(attrs) {
  if (hasVisibleBorder(attrs.pBdr)) return true;
  const st = R.model && R.model.styles;
  return !!st && hasVisibleBorder(st.resolvePara(attrs.styleId || st.defaultPara).pPr.pBdr);
}

export function buildDecorations(doc) {
  const decos = [];
  const counter = new NumberingCounter(R.model && R.model.numbering);
  const visit = (parent, base) => {
    let prev = null;
    let prevPos = 0;
    parent.forEach((node, offset) => {
      const pos = base + offset;
      if (node.type.name === 'paragraph') {
        const lbl = labelFor(counter, node.attrs);
        if (lbl) {
          decos.push(Decoration.widget(pos + 1, () => labelDOM(lbl), { side: -1, marks: [], key: labelKey(lbl), ignoreSelection: true, numLabel: true }));
        }
        if (prev && prev.type.name === 'paragraph' && styleOf(prev.attrs) === styleOf(node.attrs)) {
          if (effContextual(prev.attrs)) decos.push(Decoration.node(prevPos, prevPos + prev.nodeSize, { class: isBordered(prev.attrs) ? 'ctx-after-m' : 'ctx-after' }));
          if (effContextual(node.attrs)) decos.push(Decoration.node(pos, pos + node.nodeSize, { class: isBordered(node.attrs) ? 'ctx-before-m' : 'ctx-before' }));
        }
      } else if (node.type.name === 'table') {
        const classes = tableCellClassList(node);
        node.forEach((row, rowOffset, r) => {
          row.forEach((cell, cellOffset, c) => {
            const cpos = pos + 1 + rowOffset + 1 + cellOffset;
            const cls = classes[r] && classes[r][c];
            if (cls) decos.push(Decoration.node(cpos, cpos + cell.nodeSize, { class: cls }));
            visit(cell, cpos + 1);
          });
        });
      } else if (node.type.name === 'sdt_block') {
        visit(node, pos + 1);
      }
      prev = node;
      prevPos = pos;
    });
  };
  visit(doc, 0);
  return DecorationSet.create(doc, decos);
}

// A estrutura (tipos, atributos e contagens) não mudou, só o texto dentro dos parágrafos
function sameShape(a, b) {
  if (a.childCount !== b.childCount) return false;
  for (let i = 0; i < a.childCount; i++) {
    const x = a.child(i);
    const y = b.child(i);
    if (x === y) continue;
    if (x.type !== y.type || x.attrs !== y.attrs) return false;
    if (x.type.name === 'paragraph') continue;
    if (x.isTextblock) continue;
    if (x.isLeaf) return false;
    if (!sameShape(x, y)) return false;
  }
  return true;
}

export function decorationsPlugin() {
  return new Plugin({
    key: decoKey,
    state: {
      init: (_, state) => buildDecorations(state.doc),
      apply(tr, old, oldState, newState) {
        if (tr.getMeta('refreshDecorations')) return buildDecorations(newState.doc);
        if (!tr.docChanged) return old;
        if (sameShape(oldState.doc, newState.doc)) return old.map(tr.mapping, newState.doc);
        return buildDecorations(newState.doc);
      },
    },
    props: {
      decorations(state) { return decoKey.getState(state); },
    },
  });
}
