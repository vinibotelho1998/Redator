// Rótulos de numeração e marcadores.
import { bulletText } from '../docx/numbering.js';
import { effectiveNumPr, runCss, fontName } from './render.js';
import { runMap, mergeRun } from '../docx/props.js';

export function labelFor(counter, attrs) {
  const np = effectiveNumPr(attrs);
  if (!np) return null;
  const res = counter.next(np.numId, np.ilvl);
  if (!res) return null;
  const def = res.def;
  let text = res.text;
  let r = runMap(attrs.markRPr || []);
  delete r.rStyle;
  delete r.rprExtra;
  delete r.vanish;
  delete r.highlight;
  delete r.shd;
  if (def.rPr) r = mergeRun(r, runMap(def.rPr));
  delete r.rprExtra;
  let font = null;
  if (def.numFmt === 'bullet') {
    const lvlFont = r.rFonts ? fontName(r.rFonts.a) : null;
    const b = bulletText(def.lvlText, lvlFont);
    text = b.text;
    if (!b.font && lvlFont && /symbol|wingdings|webdings/i.test(lvlFont)) delete r.rFonts;
    font = b.font;
  }
  let css = runCss(r);
  if (font) css += ';font-family:"' + font + '"';
  return { text, css, suff: def.suff || 'tab', jc: def.lvlJc || 'left', numId: np.numId, ilvl: np.ilvl };
}

export function labelDOM(lbl) {
  const span = document.createElement('span');
  span.className = 'num-lbl';
  span.contentEditable = 'false';
  const txt = document.createElement('span');
  txt.className = 'num-txt';
  txt.textContent = lbl.text;
  if (lbl.css) txt.style.cssText = lbl.css;
  span.appendChild(txt);
  if (lbl.suff === 'tab') {
    const t = document.createElement('span');
    t.className = 'num-tab';
    span.appendChild(t);
  } else if (lbl.suff === 'space') {
    span.appendChild(document.createTextNode(' '));
  }
  if (lbl.jc === 'right') span.classList.add('num-right');
  else if (lbl.jc === 'center') span.classList.add('num-center');
  return span;
}

export function labelKey(lbl) {
  return 'num|' + lbl.text + '|' + lbl.css + '|' + lbl.suff + '|' + lbl.jc;
}
