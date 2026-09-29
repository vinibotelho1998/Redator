// Comandos de edição no estilo do Word.
import { TextSelection, NodeSelection, AllSelection } from 'prosemirror-state';
import { Fragment, Slice } from 'prosemirror-model';
import { CellSelection, TableMap, addRowAfter, goToNextCell, isInTable, selectedRect } from 'prosemirror-tables';
import { schema, RUN_MARKS, WRAPPER_MARKS, TOGGLE_MARKS } from './schema.js';
import { R, effectiveNumPr } from './render.js';

const nodes = schema.nodes;
const markTypes = schema.marks;

// ---------------------------------------------------------------- utilidades

export function paragraphsInSelection(state) {
  const out = [];
  const seen = new Set();
  const sel = state.selection;
  const ranges = sel instanceof CellSelection ? [] : sel.ranges;
  if (sel instanceof CellSelection) {
    sel.forEachCell((cell, pos) => {
      cell.descendants((n, p) => {
        if (n.type === nodes.paragraph) { const abs = pos + 1 + p; if (!seen.has(abs)) { seen.add(abs); out.push({ node: n, pos: abs }); } return false; }
        return true;
      });
    });
    return out;
  }
  for (const r of ranges) {
    state.doc.nodesBetween(r.$from.pos, r.$to.pos, (n, pos) => {
      if (n.type === nodes.paragraph) {
        if (!seen.has(pos)) { seen.add(pos); out.push({ node: n, pos }); }
        return false;
      }
      return true;
    });
  }
  if (!out.length) {
    const $f = sel.$from;
    for (let d = $f.depth; d >= 0; d--) {
      if ($f.node(d).type === nodes.paragraph) { out.push({ node: $f.node(d), pos: $f.before(d) }); break; }
    }
  }
  return out;
}

export function currentParagraph(state) {
  const $f = state.selection.$from;
  for (let d = $f.depth; d >= 0; d--) if ($f.node(d).type === nodes.paragraph) return { node: $f.node(d), pos: d > 0 ? $f.before(d) : 0 };
  if (state.selection instanceof NodeSelection && state.selection.node.type === nodes.paragraph) return { node: state.selection.node, pos: state.selection.from };
  return null;
}

export function updateParagraphs(fn) {
  return (state, dispatch) => {
    const paras = paragraphsInSelection(state);
    if (!paras.length) return false;
    if (dispatch) {
      const tr = state.tr;
      for (const { node, pos } of paras) {
        const next = fn({ ...node.attrs }, node);
        if (next) tr.setNodeMarkup(pos, null, next);
      }
      dispatch(tr.scrollIntoView());
    }
    return true;
  };
}

export function marksToJSON(marks) {
  return marks.filter((m) => RUN_MARKS.includes(m.type.name)).map((m) => ({ type: m.type.name, attrs: m.attrs }));
}

export function marksFromJSON(list) {
  return (list || []).map((m) => { try { return schema.mark(m.type, m.attrs); } catch { return null; } }).filter(Boolean);
}

function cursorMarks(state) {
  const $f = state.selection.$from;
  if (state.storedMarks) return state.storedMarks;
  if ($f.parent.type === nodes.paragraph && $f.parent.content.size === 0 && $f.parent.attrs.markRPr) return marksFromJSON($f.parent.attrs.markRPr);
  return $f.marks();
}

function textNodesInSelection(state, fn) {
  for (const r of state.selection.ranges) {
    state.doc.nodesBetween(r.$from.pos, r.$to.pos, (node, pos, parent) => {
      if (!node.isInline) return true;
      const from = Math.max(pos, r.$from.pos);
      const to = Math.min(pos + node.nodeSize, r.$to.pos);
      if (from < to) fn(node, from, to, parent);
      return false;
    });
  }
}

// ---------------------------------------------------------------- propriedades efetivas

export function effectiveRunAt(state) {
  const styles = R.model && R.model.styles;
  if (!styles) return {};
  const $f = state.selection.$from;
  const para = currentParagraph(state);
  const pattrs = para ? para.node.attrs : {};
  if (state.selection.empty) return styles.effectiveRun(pattrs, cursorMarks(state));
  // primeiro trecho de texto da seleção
  let marks = null;
  let pa = pattrs;
  textNodesInSelection(state, (node, from, to, parent) => { if (!marks && node.isText) { marks = node.marks; pa = parent.attrs; } });
  return styles.effectiveRun(pa, marks || $f.marks());
}

function styleValue(paraAttrs, marks, name) {
  const styles = R.model && R.model.styles;
  if (!styles) return false;
  const r = styles.effectiveRun(paraAttrs, marks.filter((m) => m.type.name !== name));
  return !!(r[name] && r[name].val !== false && (name !== 'u' || (r.u.a && r.u.a.val !== 'none')));
}

function effValue(paraAttrs, marks, name) {
  const direct = marks.find((m) => m.type.name === name);
  if (direct) return name === 'u' ? !!(direct.attrs.a && direct.attrs.a.val !== 'none') : direct.attrs.val !== false;
  return styleValue(paraAttrs, marks, name);
}

export function isRunPropOn(state, name) {
  const para = currentParagraph(state);
  if (state.selection.empty) return effValue(para ? para.node.attrs : {}, cursorMarks(state), name);
  let all = true;
  let any = false;
  textNodesInSelection(state, (node, from, to, parent) => {
    if (!node.isText) return;
    any = true;
    if (!effValue(parent.attrs, node.marks, name)) all = false;
  });
  return any ? all : effValue(para ? para.node.attrs : {}, state.selection.$from.marks(), name);
}

// ---------------------------------------------------------------- formatação de caractere

function applyToParagraphMarkIfEmpty(tr, state, updateMarks) {
  const para = currentParagraph(state);
  if (para && para.node.content.size === 0) {
    const cur = marksFromJSON(para.node.attrs.markRPr);
    tr.setNodeMarkup(para.pos, null, { ...para.node.attrs, markRPr: marksToJSON(updateMarks(cur)) });
  }
}

export function toggleRunProp(name) {
  const type = markTypes[name];
  return (state, dispatch) => {
    if (!dispatch) return true;
    const on = isRunPropOn(state, name);
    const target = !on;
    const tr = state.tr;
    if (state.selection.empty) {
      const para = currentParagraph(state);
      const pa = para ? para.node.attrs : {};
      const update = (marks) => {
        const base = styleValue(pa, marks, name);
        let set = type.removeFromSet(marks);
        if (target !== base) set = (name === 'u' ? type.create({ a: { val: target ? 'single' : 'none' } }) : type.create({ val: target })).addToSet(set);
        return set;
      };
      const next = update(cursorMarks(state));
      applyToParagraphMarkIfEmpty(tr, state, update);
      tr.setStoredMarks(next);
    } else {
      textNodesInSelection(state, (node, from, to, parent) => {
        const base = styleValue(parent.attrs, node.marks, name);
        if (target === base) tr.removeMark(from, to, type);
        else tr.addMark(from, to, name === 'u' ? type.create({ a: { val: target ? 'single' : 'none' } }) : type.create({ val: target }));
      });
    }
    dispatch(tr.scrollIntoView());
    return true;
  };
}

export function setRunMark(name, attrs) {
  const type = markTypes[name];
  return (state, dispatch) => {
    if (!dispatch) return true;
    const tr = state.tr;
    const update = (marks) => {
      let set = type.removeFromSet(marks);
      if (attrs) set = type.create(attrs).addToSet(set);
      return set;
    };
    if (state.selection.empty) {
      const next = update(cursorMarks(state));
      applyToParagraphMarkIfEmpty(tr, state, update);
      tr.setStoredMarks(next);
    } else {
      for (const r of state.selection.ranges) {
        if (attrs) tr.addMark(r.$from.pos, r.$to.pos, type.create(attrs));
        else tr.removeMark(r.$from.pos, r.$to.pos, type);
      }
    }
    dispatch(tr.scrollIntoView());
    return true;
  };
}

export const setFont = (name) => setRunMark('rFonts', name ? { a: { ascii: name, hAnsi: name, cs: name, eastAsia: name } } : null);
export const setFontSize = (pt) => setRunMark('sz', pt ? { val: Math.round(pt * 2) } : null);
export const setColor = (c) => setRunMark('color', c ? { a: c } : null);
export const setHighlight = (v) => setRunMark('highlight', v && v !== 'none' ? { val: v } : null);
export const setUnderline = (val) => setRunMark('u', val ? { a: { val } } : null);
export const setCharSpacing = (tw) => setRunMark('spacing', tw ? { val: tw } : null);

export function setVertAlign(val) {
  return (state, dispatch) => {
    const r = effectiveRunAt(state);
    const cur = r.vertAlign ? r.vertAlign.val : 'baseline';
    return setRunMark('vertAlign', cur === val ? null : { val })(state, dispatch);
  };
}

export const FONT_STEPS = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

function stepSize(size, dir, fine) {
  if (fine) return Math.max(1, Math.min(1638, size + dir));
  if (dir > 0) return FONT_STEPS.find((s) => s > size + 0.01) ?? Math.min(1638, Math.ceil((size + 10) / 10) * 10);
  const lower = [...FONT_STEPS].reverse().find((s) => s < size - 0.01);
  if (lower != null) return lower;
  return Math.max(1, size - 1);
}

export function growFont(dir, fine) {
  return (state, dispatch) => {
    if (!dispatch) return true;
    const styles = R.model.styles;
    const tr = state.tr;
    if (state.selection.empty) {
      const para = currentParagraph(state);
      const marks = cursorMarks(state);
      const r = styles.effectiveRun(para ? para.node.attrs : {}, marks);
      const size = r.sz ? r.sz.val / 2 : 11;
      const next = stepSize(size, dir, fine);
      const set = markTypes.sz.create({ val: Math.round(next * 2) }).addToSet(markTypes.sz.removeFromSet(marks));
      tr.setStoredMarks(set);
    } else {
      textNodesInSelection(state, (node, from, to, parent) => {
        const r = styles.effectiveRun(parent.attrs, node.marks);
        const size = r.sz ? r.sz.val / 2 : 11;
        tr.addMark(from, to, markTypes.sz.create({ val: Math.round(stepSize(size, dir, fine) * 2) }));
      });
    }
    dispatch(tr.scrollIntoView());
    return true;
  };
}

// Limpar formatação de caractere (Ctrl+Espaço)
export function clearCharFormatting(state, dispatch) {
  if (!dispatch) return true;
  const tr = state.tr;
  if (state.selection.empty) {
    tr.setStoredMarks(cursorMarks(state).filter((m) => WRAPPER_MARKS.includes(m.type.name)));
  } else {
    for (const name of RUN_MARKS) {
      if (name === 'rprExtra') continue;
      for (const r of state.selection.ranges) tr.removeMark(r.$from.pos, r.$to.pos, markTypes[name]);
    }
  }
  dispatch(tr.scrollIntoView());
  return true;
}

// Limpar toda a formatação: caractere e parágrafo
export function clearAllFormatting(state, dispatch) {
  if (!dispatch) return true;
  const tr = state.tr;
  for (const name of RUN_MARKS) {
    if (name === 'rprExtra') continue;
    for (const r of state.selection.ranges) tr.removeMark(r.$from.pos, r.$to.pos, markTypes[name]);
  }
  for (const { node, pos } of paragraphsInSelection(state)) {
    tr.setNodeMarkup(pos, null, { sectPr: node.attrs.sectPr, extra: null });
  }
  if (state.selection.empty) tr.setStoredMarks([]);
  dispatch(tr.scrollIntoView());
  return true;
}

export function resetParagraphFormatting(state, dispatch) {
  return updateParagraphs((a) => ({ styleId: a.styleId, numId: a.numId, ilvl: a.ilvl, sectPr: a.sectPr, markRPr: a.markRPr, extra: a.extra }))(state, dispatch);
}

// ---------------------------------------------------------------- maiúsculas e minúsculas

function transformCase(text, mode, state) {
  switch (mode) {
    case 'upper': return text.toLocaleUpperCase('pt-BR');
    case 'lower': return text.toLocaleLowerCase('pt-BR');
    case 'title': return text.toLocaleLowerCase('pt-BR').replace(/(^|[\s\-("'“‘])(\p{L})/gu, (m, a, b) => a + b.toLocaleUpperCase('pt-BR'));
    case 'toggle': return Array.from(text).map((c) => (c === c.toLocaleUpperCase('pt-BR') ? c.toLocaleLowerCase('pt-BR') : c.toLocaleUpperCase('pt-BR'))).join('');
    case 'sentence': {
      let start = state.start;
      let out = '';
      for (const c of text.toLocaleLowerCase('pt-BR')) {
        if (start && /\p{L}/u.test(c)) { out += c.toLocaleUpperCase('pt-BR'); start = false; }
        else { out += c; if (/[.!?]/.test(c)) start = true; }
      }
      state.start = start;
      return out;
    }
    default: return text;
  }
}

export function changeCase(mode) {
  return (state, dispatch) => {
    if (state.selection.empty) return false;
    if (!dispatch) return true;
    const tr = state.tr;
    const st = { start: true };
    const items = [];
    textNodesInSelection(state, (node, from, to) => {
      if (!node.isText) return;
      items.push({ from, to, text: state.doc.textBetween(from, to), marks: node.marks });
    });
    for (const it of items) {
      const t = transformCase(it.text, mode, st);
      if (t !== it.text) tr.replaceWith(tr.mapping.map(it.from), tr.mapping.map(it.to), schema.text(t, it.marks));
    }
    const { from, to } = state.selection;
    tr.setSelection(TextSelection.create(tr.doc, tr.mapping.map(from, -1), tr.mapping.map(to)));
    dispatch(tr);
    return true;
  };
}

// ---------------------------------------------------------------- parágrafo

export const setAlign = (jc) => updateParagraphs((a) => ({ ...a, jc }));

export function setLineSpacing(mult) {
  return updateParagraphs((a) => ({ ...a, spacing: { ...(a.spacing || {}), line: String(Math.round(mult * 240)), lineRule: 'auto' } }));
}

export function setLineSpacingRule(rule, twips) {
  return updateParagraphs((a) => ({ ...a, spacing: { ...(a.spacing || {}), line: String(Math.round(twips)), lineRule: rule } }));
}

export function setSpacing(before, after) {
  return updateParagraphs((a) => {
    const sp = { ...(a.spacing || {}) };
    if (before != null) { sp.before = String(Math.round(before)); delete sp.beforeAutospacing; delete sp.beforeLines; }
    if (after != null) { sp.after = String(Math.round(after)); delete sp.afterAutospacing; delete sp.afterLines; }
    return { ...a, spacing: sp };
  });
}

export function effectivePara(state) {
  const p = currentParagraph(state);
  if (!p || !R.model) return {};
  return R.model.styles.effectivePPr(p.node.attrs, R.model.numbering);
}

export function toggleSpaceBefore(state, dispatch) {
  const eff = effectivePara(state);
  const has = eff.spacing && Number(eff.spacing.before) > 0;
  return setSpacing(has ? 0 : 240, null)(state, dispatch);
}

export function toggleSpaceAfter(state, dispatch) {
  const eff = effectivePara(state);
  const has = eff.spacing && Number(eff.spacing.after) > 0;
  return setSpacing(null, has ? 0 : 160)(state, dispatch);
}

function numLevelInd(attrs) {
  const np = effectiveNumPr(attrs);
  if (!np || !R.model) return null;
  const lvl = R.model.numbering.level(np.numId, np.ilvl);
  return lvl && lvl.pPr && lvl.pPr.ind;
}

export function changeIndent(dir) {
  return (state, dispatch) => {
    const paras = paragraphsInSelection(state);
    if (!paras.length) return false;
    if (dispatch) {
      const tr = state.tr;
      const step = (R.model && R.model.settings.defaultTabStop) || 708;
      for (const { node, pos } of paras) {
        const a = node.attrs;
        const np = effectiveNumPr(a);
        if (np) {
          const ilvl = Math.max(0, Math.min(8, np.ilvl + dir));
          tr.setNodeMarkup(pos, null, { ...a, numId: np.numId, ilvl, ind: a.ind ? { ...a.ind, left: undefined } : null });
          continue;
        }
        const eff = R.model.styles.effectivePPr(a, R.model.numbering);
        const left = Number((eff.ind && eff.ind.left) || 0);
        let next = dir > 0 ? (Math.floor(left / step) + 1) * step : (Math.ceil(left / step) - 1) * step;
        next = Math.max(0, next);
        const ind = { ...(a.ind || {}), left: String(next) };
        tr.setNodeMarkup(pos, null, { ...a, ind });
      }
      dispatch(tr.scrollIntoView());
    }
    return true;
  };
}

export function setIndent(values) {
  return updateParagraphs((a) => {
    const ind = { ...(a.ind || {}) };
    for (const [k, v] of Object.entries(values)) {
      if (v == null) delete ind[k];
      else ind[k] = String(Math.round(v));
    }
    if (values.firstLine != null) delete ind.hanging;
    if (values.hanging != null) delete ind.firstLine;
    return { ...a, ind };
  });
}

export function hangingIndent(dir) {
  return updateParagraphs((a) => {
    const eff = R.model.styles.effectivePPr(a, R.model.numbering);
    const step = R.model.settings.defaultTabStop || 708;
    const left = Number((eff.ind && eff.ind.left) || 0);
    const hanging = Number((eff.ind && eff.ind.hanging) || 0);
    if (dir > 0) return { ...a, ind: { ...(a.ind || {}), left: String(left + step), hanging: String(hanging + step), firstLine: undefined } };
    const h = Math.max(0, hanging - step);
    return { ...a, ind: { ...(a.ind || {}), left: String(Math.max(0, left - step)), hanging: String(h), firstLine: undefined } };
  });
}

export function setParaProps(props) {
  return updateParagraphs((a) => {
    const next = { ...a };
    for (const [k, v] of Object.entries(props)) next[k] = v;
    return next;
  });
}

export function setParaShading(fill) {
  return updateParagraphs((a) => ({ ...a, shd: fill ? { val: 'clear', color: 'auto', fill: fill.replace('#', '') } : null }));
}

export function setParaBorders(kind) {
  const line = { val: 'single', sz: '4', space: '1', color: 'auto' };
  return (state, dispatch) => {
    const paras = paragraphsInSelection(state);
    if (!paras.length) return false;
    if (!dispatch) return true;
    const tr = state.tr;
    paras.forEach(({ node, pos }, i) => {
      const cur = { ...(node.attrs.pBdr || {}) };
      let b;
      switch (kind) {
        case 'none': b = null; break;
        case 'all': case 'outside': b = { top: line, left: { ...line, space: '4' }, bottom: line, right: { ...line, space: '4' }, ...(kind === 'all' ? { between: line } : {}) }; break;
        case 'bottom': case 'top': case 'left': case 'right': {
          const side = kind;
          if (cur[side] && cur[side].val !== 'nil' && cur[side].val !== 'none') delete cur[side];
          else cur[side] = side === 'left' || side === 'right' ? { ...line, space: '4' } : line;
          b = Object.keys(cur).length ? cur : null;
          break;
        }
        case 'hr': b = { bottom: { val: 'single', sz: '6', space: '1', color: 'auto' } }; break;
        default: b = cur;
      }
      tr.setNodeMarkup(pos, null, { ...node.attrs, pBdr: b });
    });
    dispatch(tr);
    return true;
  };
}

// Aplicar estilo de parágrafo ou de caractere
export function applyStyle(styleId) {
  return (state, dispatch) => {
    const styles = R.model.styles;
    const st = styles.get(styleId);
    if (!st) return false;
    if (st.type === 'character') {
      if (state.selection.empty) {
        // aplica à palavra sob o cursor, como o Word
        const $f = state.selection.$from;
        const text = $f.parent.textContent;
        const off = $f.parentOffset;
        let a = off, b = off;
        while (a > 0 && /[\p{L}\p{N}_]/u.test(text[a - 1])) a--;
        while (b < text.length && /[\p{L}\p{N}_]/u.test(text[b])) b++;
        if (a === b) return setRunMark('rStyle', { val: styleId })(state, dispatch);
        if (dispatch) {
          const start = $f.start();
          dispatch(state.tr.addMark(start + a, start + b, markTypes.rStyle.create({ val: styleId })));
        }
        return true;
      }
      return setRunMark('rStyle', { val: styleId })(state, dispatch);
    }
    if (st.type !== 'paragraph') return false;
    if (!dispatch) return true;
    const tr = state.tr;
    const paras = paragraphsInSelection(state);
    for (const { node, pos } of paras) {
      const a = { ...node.attrs, styleId: styleId === styles.defaultPara ? null : styleId };
      // o estilo de título tira a numeração direta de lista comum
      const res = styles.resolvePara(styleId);
      if (res.pPr.numId != null) { a.numId = null; a.ilvl = null; }
      tr.setNodeMarkup(pos, null, a);
      // remove formatação direta que cobre a maior parte do parágrafo
      const size = node.content.size;
      if (size > 0) {
        const cover = {};
        node.forEach((ch) => {
          for (const m of ch.marks) if (['rFonts', 'sz', 'color', 'b', 'i', 'caps'].includes(m.type.name)) cover[m.type.name] = (cover[m.type.name] || 0) + ch.nodeSize;
        });
        for (const [name, n] of Object.entries(cover)) if (n >= size / 2) tr.removeMark(pos + 1, pos + 1 + size, markTypes[name]);
      }
    }
    dispatch(tr.scrollIntoView());
    return true;
  };
}

// ---------------------------------------------------------------- listas

function listKindOf(attrs) {
  const np = effectiveNumPr(attrs);
  if (!np) return null;
  return R.model.numbering.kindOf(np.numId, np.ilvl);
}

export function currentListKind(state) {
  const p = currentParagraph(state);
  return p ? listKindOf(p.node.attrs) : null;
}

export function toggleList(kind, variant) {
  return (state, dispatch) => {
    const paras = paragraphsInSelection(state);
    if (!paras.length) return false;
    if (!dispatch) return true;
    const model = R.model;
    const allSame = variant == null && paras.every((p) => listKindOf(p.node.attrs) === kind);
    const tr = state.tr;
    if (allSame) {
      for (const { node, pos } of paras) {
        const styleNum = model.styles.resolvePara(node.attrs.styleId || model.styles.defaultPara).pPr.numId;
        const a = { ...node.attrs, numId: styleNum ? '0' : null, ilvl: null };
        if (a.styleId && /list ?paragraph/i.test((model.styles.get(a.styleId) || {}).name || '')) a.styleId = null;
        tr.setNodeMarkup(pos, null, a);
      }
      dispatch(tr);
      return true;
    }
    // continua a lista do parágrafo anterior quando for do mesmo tipo
    let numId = null;
    const first = paras[0];
    const $p = state.doc.resolve(first.pos);
    const idx = $p.index();
    if (variant == null && idx > 0) {
      const prev = $p.parent.child(idx - 1);
      if (prev.type === nodes.paragraph && listKindOf(prev.attrs) === kind) numId = effectiveNumPr(prev.attrs).numId;
    }
    if (!numId) numId = model.numbering.createList(kind, variant || 0);
    const listStyle = model.styles.findByName('List Paragraph');
    for (const { node, pos } of paras) {
      const cur = effectiveNumPr(node.attrs);
      const a = { ...node.attrs, numId, ilvl: cur ? cur.ilvl : 0 };
      if (a.ind) { const ind = { ...a.ind }; delete ind.left; delete ind.hanging; delete ind.firstLine; a.ind = Object.keys(ind).length ? ind : null; }
      if (listStyle && (!a.styleId || a.styleId === model.styles.defaultPara)) a.styleId = listStyle.id;
      tr.setNodeMarkup(pos, null, a);
    }
    tr.setMeta('refreshDecorations', true);
    dispatch(tr.scrollIntoView());
    return true;
  };
}

export function setListLevel(dir) {
  return (state, dispatch) => {
    const paras = paragraphsInSelection(state).filter((p) => effectiveNumPr(p.node.attrs));
    if (!paras.length) return false;
    if (dispatch) {
      const tr = state.tr;
      for (const { node, pos } of paras) {
        const np = effectiveNumPr(node.attrs);
        tr.setNodeMarkup(pos, null, { ...node.attrs, numId: np.numId, ilvl: Math.max(0, Math.min(8, np.ilvl + dir)) });
      }
      dispatch(tr.scrollIntoView());
    }
    return true;
  };
}

export function restartNumbering(state, dispatch) {
  const p = currentParagraph(state);
  const np = p && effectiveNumPr(p.node.attrs);
  if (!np) return false;
  if (!dispatch) return true;
  const newId = R.model.numbering.restartList(np.numId);
  // aplica a partir deste parágrafo até o fim da lista contígua
  const tr = state.tr;
  const $p = state.doc.resolve(p.pos);
  const parent = $p.parent;
  let pos = p.pos;
  for (let i = $p.index(); i < parent.childCount; i++) {
    const ch = parent.child(i);
    const cnp = ch.type === nodes.paragraph ? effectiveNumPr(ch.attrs) : null;
    if (i > $p.index() && (!cnp || cnp.numId !== np.numId)) {
      if (!cnp && ch.type === nodes.paragraph && ch.content.size === 0) { pos += ch.nodeSize; continue; }
      if (!cnp) break;
    }
    if (cnp && cnp.numId === np.numId) tr.setNodeMarkup(pos, null, { ...ch.attrs, numId: newId, ilvl: cnp.ilvl });
    pos += ch.nodeSize;
  }
  tr.setMeta('refreshDecorations', true);
  dispatch(tr);
  return true;
}

// ---------------------------------------------------------------- Enter, Backspace, Tab

export function splitParagraph(state, dispatch) {
  const { $from, empty } = state.selection;
  if (!(state.selection instanceof TextSelection) && !(state.selection instanceof AllSelection)) {
    if (!(state.selection instanceof NodeSelection)) return false;
  }
  if ($from.parent.type !== nodes.paragraph) return false;
  const para = $from.parent;
  const model = R.model;
  if (empty && para.content.size === 0 && effectiveNumPr(para.attrs) && para.attrs.numId != null) {
    if (dispatch) {
      const pos = $from.before();
      const a = { ...para.attrs };
      const styleNum = model.styles.resolvePara(a.styleId || model.styles.defaultPara).pPr.numId;
      if (a.ilvl > 0) a.ilvl -= 1;
      else { a.numId = styleNum ? '0' : null; a.ilvl = null; a.ind = null; if (a.styleId && /list ?paragraph/i.test((model.styles.get(a.styleId) || {}).name || '')) a.styleId = null; }
      dispatch(state.tr.setNodeMarkup(pos, null, a).setMeta('refreshDecorations', true));
    }
    return true;
  }
  if (!dispatch) return true;
  const tr = state.tr;
  if (!empty) tr.deleteSelection();
  const $pos = tr.selection.$from;
  if ($pos.parent.type !== nodes.paragraph) { dispatch(tr); return true; }
  const cur = $pos.parent;
  const marks = tr.storedMarks || $pos.marks();
  const atEnd = $pos.parentOffset === cur.content.size;
  const st = model.styles.get(cur.attrs.styleId || model.styles.defaultPara);
  let after = { ...cur.attrs };
  if (atEnd && st && st.next && st.next !== st.id) {
    after = { styleId: st.next === model.styles.defaultPara ? null : st.next };
  }
  after.markRPr = marksToJSON(marks).length ? marksToJSON(marks) : null;
  after.pageBreakBefore = null;
  after.mdFence = null;
  // depois do título de um callout vem o corpo dele
  if (after.md && after.md.head) after.md = { callout: after.md.callout };
  after.sectPr = cur.attrs.sectPr;
  const parentPos = $pos.before();
  tr.split($pos.pos, 1, [{ type: nodes.paragraph, attrs: after }]);
  if (cur.attrs.sectPr) tr.setNodeMarkup(parentPos, null, { ...cur.attrs, sectPr: null });
  const wrappers = marks.filter((m) => !WRAPPER_MARKS.includes(m.type.name) || m.type.name === 'sdt');
  tr.ensureMarks(wrappers);
  dispatch(tr.scrollIntoView());
  return true;
}

export function lineBreak(state, dispatch) {
  if (!state.selection.$from.parent.inlineContent) return false;
  if (dispatch) {
    const marks = state.storedMarks || state.selection.$from.marks();
    dispatch(state.tr.replaceSelectionWith(nodes.hard_break.create(null, null, marks)).scrollIntoView());
  }
  return true;
}

export function pageBreak(state, dispatch) {
  const { $from } = state.selection;
  if ($from.parent.type !== nodes.paragraph) return false;
  for (let d = $from.depth; d > 0; d--) if ($from.node(d).type === nodes.table_cell) return false;
  if (!dispatch) return true;
  const tr = state.tr;
  if (!state.selection.empty) tr.deleteSelection();
  const $p = tr.selection.$from;
  const para = $p.parent;
  const br = nodes.hard_break.create({ type: 'page' }, null, $p.marks());
  if ($p.parentOffset === 0 && para.content.size > 0) {
    tr.insert($p.before(), nodes.paragraph.create({ ...para.attrs, sectPr: null, pageBreakBefore: null }, [br]));
  } else {
    const paraPos = $p.before();
    tr.insert($p.pos, br);
    tr.split($p.pos + br.nodeSize, 1, [{ type: nodes.paragraph, attrs: { ...para.attrs, pageBreakBefore: null } }]);
    if (para.attrs.sectPr) tr.setNodeMarkup(paraPos, null, { ...para.attrs, sectPr: null });
  }
  dispatch(tr.scrollIntoView());
  return true;
}

export function backspaceSpecial(state, dispatch) {
  const { $from, empty } = state.selection;
  if (!empty || $from.parentOffset !== 0 || $from.parent.type !== nodes.paragraph) return false;
  const para = $from.parent;
  const model = R.model;
  const np = effectiveNumPr(para.attrs);
  if (np) {
    if (dispatch) {
      const ind = numLevelInd(para.attrs);
      const styleNum = model.styles.resolvePara(para.attrs.styleId || model.styles.defaultPara).pPr.numId;
      const a = { ...para.attrs, numId: styleNum ? '0' : null, ilvl: null };
      if (ind && ind.left != null) a.ind = { ...(para.attrs.ind || {}), left: ind.left, firstLine: '0' };
      if (a.ind) delete a.ind.hanging;
      dispatch(state.tr.setNodeMarkup($from.before(), null, a).setMeta('refreshDecorations', true));
    }
    return true;
  }
  const ind = para.attrs.ind;
  if (ind && Number(ind.firstLine) > 0) {
    if (dispatch) dispatch(state.tr.setNodeMarkup($from.before(), null, { ...para.attrs, ind: { ...ind, firstLine: undefined } }));
    return true;
  }
  return false;
}

export function insertTab(state, dispatch) {
  if (!state.selection.$from.parent.inlineContent) return false;
  if (dispatch) {
    const marks = state.storedMarks || state.selection.$from.marks();
    dispatch(state.tr.replaceSelectionWith(nodes.tab.create(null, null, marks)).scrollIntoView());
  }
  return true;
}

export function tabKey(dir) {
  return (state, dispatch, view) => {
    const { $from, empty } = state.selection;
    const inList = $from.parent.type === nodes.paragraph && effectiveNumPr($from.parent.attrs);
    if (inList && (!empty || $from.parentOffset === 0)) return setListLevel(dir)(state, dispatch);
    if (isInTable(state)) {
      if (goToNextCell(dir)(state, dispatch)) return true;
      if (dir > 0 && dispatch && view) {
        addRowAfter(state, (tr) => {
          dispatch(tr);
          goToNextCell(1)(view.state, view.dispatch);
        });
        return true;
      }
      return dir < 0;
    }
    if (dir > 0) return insertTab(state, dispatch);
    return changeIndent(-1)(state, dispatch);
  };
}

// ---------------------------------------------------------------- inserções

export function insertText(text) {
  return (state, dispatch) => {
    if (dispatch) {
      const marks = state.storedMarks || state.selection.$from.marks();
      dispatch(state.tr.replaceSelectionWith(schema.text(text, marks), false).scrollIntoView());
    }
    return true;
  };
}

export function insertTable(rows, cols) {
  return (state, dispatch) => {
    if (!dispatch) return true;
    const model = R.model;
    const s = model.section;
    const contentTw = s.pgSz.w - s.pgMar.left - s.pgMar.right - (s.pgMar.gutter || 0);
    const colTw = Math.floor(contentTw / cols);
    const grid = new Array(cols).fill(colTw);
    const gridStyle = model.styles.findByName('Table Grid');
    const cells = () => Array.from({ length: cols }, () => nodes.table_cell.create({ colwidth: [colTw / 15], tcW: { w: String(colTw), type: 'dxa' } }, [nodes.paragraph.create()]));
    const table = nodes.table.create({
      styleId: gridStyle ? gridStyle.id : null,
      tblW: { w: '0', type: 'auto' },
      tblLook: { val: '04A0', firstRow: '1', lastRow: '0', firstColumn: '1', lastColumn: '0', noHBand: '0', noVBand: '1' },
      tblBorders: gridStyle ? null : {
        top: { val: 'single', sz: '4', space: '0', color: 'auto' }, left: { val: 'single', sz: '4', space: '0', color: 'auto' },
        bottom: { val: 'single', sz: '4', space: '0', color: 'auto' }, right: { val: 'single', sz: '4', space: '0', color: 'auto' },
        insideH: { val: 'single', sz: '4', space: '0', color: 'auto' }, insideV: { val: 'single', sz: '4', space: '0', color: 'auto' },
      },
      grid,
    }, Array.from({ length: rows }, () => nodes.table_row.create(null, cells())));
    const tr = state.tr;
    const para = currentParagraph(state);
    if (!para) return false;
    if (!state.selection.empty) tr.deleteSelection();
    const $f = tr.selection.$from;
    let insertPos;
    if (para.node.content.size === 0 || $f.parentOffset === 0) insertPos = $f.before();
    else if ($f.parentOffset === $f.parent.content.size) insertPos = $f.after();
    else {
      tr.split($f.pos);
      insertPos = tr.mapping.map($f.pos, -1) + 1;
    }
    tr.insert(insertPos, table);
    const after = insertPos + table.nodeSize;
    const $after = tr.doc.resolve(Math.min(after, tr.doc.content.size));
    if (!$after.nodeAfter || $after.nodeAfter.type !== nodes.paragraph) tr.insert(after, nodes.paragraph.create());
    tr.setSelection(TextSelection.near(tr.doc.resolve(insertPos + 1)));
    dispatch(tr.scrollIntoView());
    return true;
  };
}

export function insertDrawing(attrs) {
  return (state, dispatch) => {
    if (dispatch) {
      const marks = state.selection.$from.marks().filter((m) => WRAPPER_MARKS.includes(m.type.name));
      const node = nodes.drawing.create(attrs, null, marks);
      const tr = state.tr.replaceSelectionWith(node);
      dispatch(tr.scrollIntoView());
    }
    return true;
  };
}

export function insertLink(href, text) {
  return (state, dispatch) => {
    if (!dispatch) return true;
    const isAnchor = href.startsWith('#');
    const mark = markTypes.hyperlink.create(isAnchor ? { anchor: href.slice(1) } : { href });
    const linkStyle = R.model.styles.findByName('Hyperlink');
    const tr = state.tr;
    const { from, to, empty } = state.selection;
    if (empty) {
      const marks = [mark, ...(linkStyle ? [markTypes.rStyle.create({ val: linkStyle.id })] : [])];
      tr.insert(from, schema.text(text || href, marks));
    } else {
      if (text != null && text !== state.doc.textBetween(from, to)) {
        const marks = state.doc.resolve(from).marks();
        tr.replaceWith(from, to, schema.text(text, marks));
        tr.addMark(from, from + text.length, mark);
        if (linkStyle) tr.addMark(from, from + text.length, markTypes.rStyle.create({ val: linkStyle.id }));
      } else {
        tr.addMark(from, to, mark);
        if (linkStyle) tr.addMark(from, to, markTypes.rStyle.create({ val: linkStyle.id }));
      }
    }
    dispatch(tr.scrollIntoView());
    return true;
  };
}

export function linkRangeAt(state, pos) {
  const $pos = state.doc.resolve(pos);
  const parent = $pos.parent;
  const start = $pos.start();
  let found = null;
  parent.forEach((ch, off) => {
    const m = ch.marks.find((x) => x.type === markTypes.hyperlink);
    if (!m) return;
    const a = start + off, b = a + ch.nodeSize;
    if (pos >= a && pos <= b) found = { mark: m, from: a, to: b };
  });
  if (!found) return null;
  // expande pelos vizinhos com a mesma marca
  let from = found.from, to = found.to;
  parent.forEach((ch, off) => {
    const a = start + off, b = a + ch.nodeSize;
    if (ch.marks.some((x) => x.eq(found.mark))) {
      if (b === from) from = a;
      if (a === to) to = b;
    }
  });
  return { mark: found.mark, from, to };
}

export function removeLink(state, dispatch) {
  const r = linkRangeAt(state, state.selection.from);
  if (!r) return false;
  if (dispatch) {
    const tr = state.tr.removeMark(r.from, r.to, markTypes.hyperlink);
    const linkStyle = R.model.styles.findByName('Hyperlink');
    if (linkStyle) tr.removeMark(r.from, r.to, markTypes.rStyle);
    dispatch(tr);
  }
  return true;
}

// ---------------------------------------------------------------- tabela

export function setCellAttrs(attrs) {
  return (state, dispatch) => {
    if (!isInTable(state)) return false;
    if (dispatch) {
      const tr = state.tr;
      const sel = state.selection;
      if (sel instanceof CellSelection) sel.forEachCell((cell, pos) => tr.setNodeMarkup(pos, null, { ...cell.attrs, ...attrs }));
      else {
        const $f = sel.$from;
        for (let d = $f.depth; d > 0; d--) {
          if ($f.node(d).type === nodes.table_cell) { tr.setNodeMarkup($f.before(d), null, { ...$f.node(d).attrs, ...attrs }); break; }
        }
      }
      dispatch(tr);
    }
    return true;
  };
}

export function findTableAt(state) {
  const $f = state.selection.$from;
  for (let d = $f.depth; d > 0; d--) if ($f.node(d).type === nodes.table) return { node: $f.node(d), pos: $f.before(d) };
  return null;
}

export function setTableAttrs(attrs) {
  return (state, dispatch) => {
    const t = findTableAt(state);
    if (!t) return false;
    if (dispatch) dispatch(state.tr.setNodeMarkup(t.pos, null, { ...t.node.attrs, ...attrs }).setMeta('refreshDecorations', true));
    return true;
  };
}

export function toggleHeaderRow(state, dispatch) {
  const $f = state.selection.$from;
  for (let d = $f.depth; d > 0; d--) {
    if ($f.node(d).type === nodes.table_row) {
      if (dispatch) dispatch(state.tr.setNodeMarkup($f.before(d), null, { ...$f.node(d).attrs, header: !$f.node(d).attrs.header }));
      return true;
    }
  }
  return false;
}

export function setCellBorders(kind) {
  const line = { val: 'single', sz: '4', space: '0', color: 'auto' };
  const none = { val: 'nil' };
  return (state, dispatch) => {
    const t = findTableAt(state);
    if (!t) return false;
    if (!dispatch) return true;
    if (kind === 'all' || kind === 'none' || kind === 'outside' || kind === 'inside') {
      const b = kind === 'none'
        ? { top: none, left: none, bottom: none, right: none, insideH: none, insideV: none }
        : kind === 'all'
          ? { top: line, left: line, bottom: line, right: line, insideH: line, insideV: line }
          : kind === 'outside'
            ? { top: line, left: line, bottom: line, right: line, insideH: none, insideV: none }
            : { top: none, left: none, bottom: none, right: none, insideH: line, insideV: line };
      const tr = state.tr.setNodeMarkup(t.pos, null, { ...t.node.attrs, tblBorders: b });
      // limpa bordas de células para valerem as da tabela
      t.node.descendants((n, p) => {
        if (n.type === nodes.table_cell && n.attrs.tcBorders) tr.setNodeMarkup(t.pos + 1 + p, null, { ...n.attrs, tcBorders: null });
        return n.type !== nodes.table_cell;
      });
      dispatch(tr.setMeta('refreshDecorations', true));
      return true;
    }
    return false;
  };
}

// ---------------------------------------------------------------- pincel de formatação

export function captureFormat(state) {
  const p = currentParagraph(state);
  const marks = state.selection.empty ? cursorMarks(state) : (() => {
    let m = null;
    textNodesInSelection(state, (node) => { if (!m && node.isText) m = node.marks; });
    return m || state.selection.$from.marks();
  })();
  return { marks: marks.filter((m) => RUN_MARKS.includes(m.type.name)), para: p ? { ...p.node.attrs } : null };
}

export function applyFormat(fmt) {
  return (state, dispatch) => {
    if (!fmt) return false;
    if (!dispatch) return true;
    const tr = state.tr;
    if (!state.selection.empty) {
      for (const name of RUN_MARKS) for (const r of state.selection.ranges) tr.removeMark(r.$from.pos, r.$to.pos, markTypes[name]);
      for (const m of fmt.marks) for (const r of state.selection.ranges) tr.addMark(r.$from.pos, r.$to.pos, m);
    }
    if (fmt.para) {
      const paras = paragraphsInSelection(state);
      const whole = state.selection.empty || paras.length > 1;
      if (whole) for (const { node, pos } of paras) tr.setNodeMarkup(pos, null, { ...fmt.para, sectPr: node.attrs.sectPr, markRPr: node.attrs.markRPr });
    }
    dispatch(tr);
    return true;
  };
}

// ---------------------------------------------------------------- seleção

export function selectParagraph(state, dispatch) {
  const p = currentParagraph(state);
  if (!p) return false;
  if (dispatch) dispatch(state.tr.setSelection(TextSelection.create(state.doc, p.pos + 1, p.pos + p.node.nodeSize - 1)));
  return true;
}

export function moveParagraph(dir) {
  return (state, dispatch) => {
    const p = currentParagraph(state);
    if (!p) return false;
    const $p = state.doc.resolve(p.pos);
    const idx = $p.index();
    const parent = $p.parent;
    if (dir < 0 && idx === 0) return false;
    if (dir > 0 && idx >= parent.childCount - 1) return false;
    if (!dispatch) return true;
    const other = parent.child(idx + dir);
    const tr = state.tr;
    const offset = state.selection.from - p.pos;
    if (dir < 0) {
      const otherPos = p.pos - other.nodeSize;
      tr.delete(p.pos, p.pos + p.node.nodeSize);
      tr.insert(otherPos, p.node);
      tr.setSelection(TextSelection.create(tr.doc, otherPos + offset));
    } else {
      const otherEnd = p.pos + p.node.nodeSize + other.nodeSize;
      tr.insert(otherEnd, p.node);
      tr.delete(p.pos, p.pos + p.node.nodeSize);
      tr.setSelection(TextSelection.create(tr.doc, p.pos + other.nodeSize + offset));
    }
    dispatch(tr.scrollIntoView());
    return true;
  };
}

export { TOGGLE_MARKS };
