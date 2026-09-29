// Atalhos de teclado do Word.
import { undo, redo } from 'prosemirror-history';
import { chainCommands, deleteSelection, joinBackward, selectNodeBackward, joinForward, selectNodeForward, selectAll } from 'prosemirror-commands';
import * as C from './commands.js';
import { R } from './render.js';

const heading = (n) => (state, dispatch) => {
  const id = R.headingStyleId(n);
  return id ? C.applyStyle(id)(state, dispatch) : false;
};

const normal = (state, dispatch) => C.applyStyle(R.model.styles.defaultPara)(state, dispatch);

export function wordKeymap() {
  return {
    'Mod-z': undo,
    'Mod-y': redo,
    'Shift-Mod-z': redo,
    'Mod-b': C.toggleRunProp('b'),
    'Mod-i': C.toggleRunProp('i'),
    'Mod-u': C.toggleRunProp('u'),
    'Mod-Shift-d': C.setUnderline('double'),
    'Mod-Shift-a': C.toggleRunProp('caps'),
    'Mod-Shift-k': C.toggleRunProp('smallCaps'),
    'Mod-Shift-h': C.toggleRunProp('vanish'),
    'Mod-=': C.setVertAlign('subscript'),
    'Mod-+': C.setVertAlign('superscript'),
    'Mod-e': C.setAlign('center'),
    'Mod-l': C.setAlign('left'),
    'Mod-r': C.setAlign('right'),
    'Mod-j': C.setAlign('both'),
    'Enter': C.splitParagraph,
    'Shift-Enter': C.lineBreak,
    'Mod-Enter': C.pageBreak,
    'Backspace': chainCommands(C.backspaceSpecial, deleteSelection, joinBackward, selectNodeBackward),
    'Delete': chainCommands(deleteSelection, joinForward, selectNodeForward),
    'Tab': C.tabKey(1),
    'Shift-Tab': C.tabKey(-1),
    'Mod-]': C.growFont(1, true),
    'Mod-[': C.growFont(-1, true),
    'Mod->': C.growFont(1),
    'Mod-<': C.growFont(-1),
    'Mod- ': C.clearCharFormatting,
    'Mod-q': C.resetParagraphFormatting,
    'Mod-1': C.setLineSpacing(1),
    'Mod-2': C.setLineSpacing(2),
    'Mod-5': C.setLineSpacing(1.5),
    'Mod-0': C.toggleSpaceBefore,
    'Mod-Alt-1': heading(1),
    'Mod-Alt-2': heading(2),
    'Mod-Alt-3': heading(3),
    'Mod-Shift-n': normal,
    'Mod-Shift-l': C.toggleList('bullet'),
    'Mod-t': C.hangingIndent(1),
    'Mod-Shift-t': C.hangingIndent(-1),
    'Mod-m': C.changeIndent(1),
    'Mod-Shift-m': C.changeIndent(-1),
    'Alt-Shift-ArrowUp': C.moveParagraph(-1),
    'Alt-Shift-ArrowDown': C.moveParagraph(1),
    'Alt-Shift-ArrowRight': C.setListLevel(1),
    'Alt-Shift-ArrowLeft': C.setListLevel(-1),
    'Mod-a': selectAll,
  };
}
