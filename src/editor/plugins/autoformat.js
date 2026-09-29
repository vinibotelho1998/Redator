// Autoformatação ao digitar: aspas curvas, listas automáticas e autocorreções.
import { Plugin, PluginKey, TextSelection } from 'prosemirror-state';
import { closeHistory } from 'prosemirror-history';
import { effectiveNumPr } from '../render.js';
import { toggleList } from '../commands.js';

export const autoKey = new PluginKey('autoformat');

const REPLACEMENTS = [
  ['(c)', '©'], ['(C)', '©'], ['(r)', '®'], ['(R)', '®'], ['(tm)', '™'], ['(TM)', '™'],
  ['...', '…'], ['-->', '→'], ['<--', '←'], ['->', '→'], ['<-', '←'], ['<=>', '⇔'], ['==>', '⇒'], ['=>', '⇒'],
];

const OPENERS = /[\s([{“‘«—–-]/;

export function autoFormatPlugin(getPrefs) {
  return new Plugin({
    key: autoKey,
    state: {
      init: () => null,
      apply(tr, prev) {
        const m = tr.getMeta(autoKey);
        if (m !== undefined) return m;
        if (tr.docChanged || tr.selectionSet) return null;
        return prev;
      },
    },
    props: {
      handleTextInput(view, from, to, text) {
        const prefs = getPrefs();
        if (!prefs.autoFormat) return false;
        const state = view.state;
        const $from = state.doc.resolve(from);
        const para = $from.parent;
        if (para.type.name !== 'paragraph') return false;
        const before = para.textBetween(0, $from.parentOffset, null, '￼');
        const marks = state.storedMarks || $from.marks();
        // aspas curvas
        if ((text === '"' || text === "'") && prefs.smartQuotes) {
          const prev = before.slice(-1);
          const open = !prev || OPENERS.test(prev);
          const ch = text === '"' ? (open ? '“' : '”') : (open ? '‘' : '’');
          view.dispatch(state.tr.insertText(ch, from, to));
          return true;
        }
        if (text !== ' ' && text !== '\t') return false;
        // lista automática no começo do parágrafo
        if ($from.parentOffset === para.content.size || $from.parentOffset === before.length) {
          const whole = before;
          if (!effectiveNumPr(para.attrs) && whole.length <= 5 && para.textContent === whole) {
            let kind = null;
            let variant = 0;
            if (/^[*•]$/.test(whole)) { kind = 'bullet'; variant = 0; }
            else if (/^-$/.test(whole)) { kind = 'bullet'; variant = 6; }
            else if (/^>$/.test(whole)) { kind = 'bullet'; variant = 4; }
            else {
              const m = /^(1|a|A|i|I)([.)])$/.exec(whole);
              if (m) {
                kind = 'number';
                if (m[1] === '1') variant = m[2] === ')' ? 1 : 0;
                else if (m[1] === 'I') variant = 2;
                else if (m[1] === 'A') variant = 3;
                else if (m[1] === 'a') variant = m[2] === ')' ? 4 : 5;
                else if (m[1] === 'i') variant = 6;
              }
            }
            if (kind) {
              const start = $from.start();
              const tr1 = state.tr.insertText(text, from, to);
              view.dispatch(tr1);
              const s2 = view.state;
              let tr2 = s2.tr.delete(start, start + whole.length + 1);
              tr2.setSelection(TextSelection.create(tr2.doc, start));
              view.dispatch(closeHistory(tr2));
              toggleList(kind, variant)(view.state, view.dispatch);
              view.dispatch(view.state.tr.setMeta(autoKey, { from: start, to: start, original: whole + text, list: true }));
              return true;
            }
          }
        }
        // substituições
        for (const [k, v] of REPLACEMENTS) {
          if (before.endsWith(k)) {
            const tr1 = state.tr.insertText(text, from, to);
            view.dispatch(tr1);
            const start = from - k.length;
            const tr2 = view.state.tr.replaceWith(start, from, view.state.schema.text(v, marks));
            view.dispatch(closeHistory(tr2).setMeta(autoKey, { from: start, to: start + v.length + 1, original: k + text }));
            return true;
          }
        }
        // travessões: "palavra - palavra" e "palavra -- palavra" viram meia-risca, "palavra--palavra" vira travessão
        let m = /\S( --? )[^\s-]+$/.exec(before);
        let dashAt = -1, dashLen = 0, dashChar = '–';
        if (m) { dashAt = m.index + 2; dashLen = m[1].length - 2; }
        else if ((m = /[^\s-](--)[^\s-]+$/.exec(before))) { dashAt = m.index + 1; dashLen = 2; dashChar = '—'; }
        if (dashAt >= 0 && dashLen > 0) {
          view.dispatch(state.tr.insertText(text, from, to));
          const start = $from.start() + dashAt;
          const tr2 = view.state.tr.replaceWith(start, start + dashLen, view.state.schema.text(dashChar, marks));
          view.dispatch(closeHistory(tr2));
          return true;
        }
        return false;
      },
      handleKeyDown(view, event) {
        if (event.key !== 'Backspace') return false;
        const info = autoKey.getState(view.state);
        if (!info || !view.state.selection.empty || info.list) return false;
        if (view.state.selection.from !== info.to) return false;
        const marks = view.state.doc.resolve(info.from).marks();
        const tr = view.state.tr.replaceWith(info.from, info.to, view.state.schema.text(info.original, marks));
        view.dispatch(tr.setMeta(autoKey, null));
        event.preventDefault();
        return true;
      },
    },
  });
}
