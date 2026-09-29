// Localizar e substituir, com realce dos resultados.
import { Plugin, PluginKey, TextSelection } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';

export const findKey = new PluginKey('find');

function escapeRegExp(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

// Texto de cada bloco. Todo nó em linha que não é texto ocupa uma posição e vira
// um caractere, então o índice no texto soma direto com o início do bloco.
function textIndex(doc) {
  const blocks = [];
  doc.descendants((node, pos) => {
    if (node.isTextblock) {
      let text = '';
      node.forEach((ch) => {
        if (ch.isText) text += ch.text;
        else text += ch.type.name === 'tab' ? '\t' : ch.type.name === 'hard_break' ? '\n' : '￼';
      });
      blocks.push({ text, start: pos + 1 });
      return false;
    }
    return true;
  });
  return blocks;
}

export function buildRegex(query, opts = {}) {
  if (!query) return null;
  let src = opts.regex ? query : escapeRegExp(query);
  if (opts.wholeWord) src = '(?<![\\p{L}\\p{N}_])(?:' + src + ')(?![\\p{L}\\p{N}_])';
  try { return new RegExp(src, (opts.matchCase ? 'g' : 'gi') + 'u'); } catch { return null; }
}

export function findMatches(doc, query, opts = {}) {
  const re = buildRegex(query, opts);
  if (!re) return [];
  const out = [];
  for (const b of textIndex(doc)) {
    re.lastIndex = 0;
    let m;
    while ((m = re.exec(b.text))) {
      if (!m[0].length) { re.lastIndex++; continue; }
      out.push({ from: b.start + m.index, to: b.start + m.index + m[0].length, text: m[0], context: b.text });
      if (out.length >= 10000) return out;
    }
  }
  return out;
}

const empty = { query: '', opts: {}, matches: [], current: -1, deco: DecorationSet.empty };

function decorate(doc, matches, current) {
  if (!matches.length) return DecorationSet.empty;
  return DecorationSet.create(doc, matches.map((m, i) => Decoration.inline(m.from, m.to, { class: i === current ? 'find-hit find-cur' : 'find-hit' })));
}

export function findPlugin() {
  return new Plugin({
    key: findKey,
    state: {
      init: () => empty,
      apply(tr, prev) {
        const meta = tr.getMeta(findKey);
        if (meta) {
          const query = meta.query != null ? meta.query : prev.query;
          const opts = meta.opts || prev.opts;
          if (!query) return empty;
          const matches = findMatches(tr.doc, query, opts);
          const current = meta.current != null ? meta.current : Math.min(prev.current, matches.length - 1);
          return { query, opts, matches, current, deco: decorate(tr.doc, matches, current) };
        }
        if (tr.docChanged && prev.query) {
          // documento grande: só desloca os realces agora e refaz a busca quando a digitação parar
          if (tr.doc.content.size > 200000) {
            const matches = [];
            for (const m of prev.matches) {
              const from = tr.mapping.map(m.from, 1);
              const to = tr.mapping.map(m.to, -1);
              if (to > from) matches.push({ ...m, from, to });
            }
            return { ...prev, matches, current: Math.min(prev.current, matches.length - 1), deco: prev.deco.map(tr.mapping, tr.doc), stale: true };
          }
          const matches = findMatches(tr.doc, prev.query, prev.opts);
          const current = Math.min(prev.current, matches.length - 1);
          return { ...prev, matches, current, deco: decorate(tr.doc, matches, current) };
        }
        return prev;
      },
    },
    props: {
      decorations(state) { return findKey.getState(state).deco; },
    },
    view() {
      let timer = null;
      return {
        update(view) {
          if (!findKey.getState(view.state).stale) return;
          clearTimeout(timer);
          timer = setTimeout(() => { if (findKey.getState(view.state).stale) view.dispatch(view.state.tr.setMeta(findKey, {})); }, 400);
        },
        destroy() { clearTimeout(timer); },
      };
    },
  });
}

export function setSearch(view, query, opts) {
  view.dispatch(view.state.tr.setMeta(findKey, { query, opts, current: -1 }));
  return findKey.getState(view.state).matches.length;
}

export function clearSearch(view) {
  view.dispatch(view.state.tr.setMeta(findKey, { query: '' }));
}

export function gotoMatch(view, index) {
  const st = findKey.getState(view.state);
  if (!st.matches.length) return -1;
  const i = ((index % st.matches.length) + st.matches.length) % st.matches.length;
  const m = st.matches[i];
  const tr = view.state.tr.setMeta(findKey, { current: i }).setSelection(TextSelection.create(view.state.doc, m.from, m.to)).scrollIntoView();
  view.dispatch(tr);
  return i;
}

export function findNext(view, dir = 1) {
  const st = findKey.getState(view.state);
  if (!st.matches.length) return -1;
  const { from, to } = view.state.selection;
  let idx;
  if (dir > 0) {
    idx = st.matches.findIndex((m) => m.from >= to || (m.from >= from && m.from > from));
    if (idx < 0) idx = 0;
  } else {
    idx = -1;
    for (let i = st.matches.length - 1; i >= 0; i--) if (st.matches[i].to <= from) { idx = i; break; }
    if (idx < 0) idx = st.matches.length - 1;
  }
  return gotoMatch(view, idx);
}

export function replaceCurrent(view, replacement) {
  const st = findKey.getState(view.state);
  const { from, to } = view.state.selection;
  const m = st.matches.find((x) => x.from === from && x.to === to);
  if (!m) return findNext(view, 1);
  const marks = view.state.doc.resolve(from + 1).marks();
  const tr = view.state.tr;
  if (replacement) tr.replaceWith(from, to, view.state.schema.text(replacement, marks));
  else tr.delete(from, to);
  tr.setSelection(TextSelection.create(tr.doc, from + replacement.length));
  view.dispatch(tr);
  return findNext(view, 1);
}

export function replaceAll(view, replacement) {
  const st = findKey.getState(view.state);
  if (!st.matches.length) return 0;
  const tr = view.state.tr;
  for (let i = st.matches.length - 1; i >= 0; i--) {
    const m = st.matches[i];
    const marks = view.state.doc.resolve(m.from + 1).marks();
    if (replacement) tr.replaceWith(m.from, m.to, view.state.schema.text(replacement, marks));
    else tr.delete(m.from, m.to);
  }
  view.dispatch(tr.scrollIntoView());
  return st.matches.length;
}
