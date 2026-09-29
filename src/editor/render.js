// Contexto de renderização: converte propriedades do Word em CSS.
import { normHex, HIGHLIGHT, mix, isDark } from '../util/color.js';
import { mergeInd } from '../docx/props.js';

export const R = {
  model: null,
  headingStyleId(level) {
    const st = R.model && R.model.styles;
    return (st && st.headingId(level)) || 'Heading' + level;
  },
  tableGridStyleId() {
    const st = R.model && R.model.styles;
    const s = st && st.findByName('Table Grid');
    return s ? s.id : null;
  },
  charClass(id) {
    return id ? cClass(id) : '';
  },
};

export function cssId(id) {
  return String(id == null ? '' : id).replace(/[^A-Za-z0-9_-]/g, (c) => '_' + c.charCodeAt(0).toString(16));
}
export const pClass = (id) => 'ps-' + cssId(id);
export const cClass = (id) => 'cs-' + cssId(id);
export const tClass = (id) => 'ts-' + cssId(id);

const r2 = (n) => Math.round(n * 100) / 100;
const tw = (v) => r2(Number(v) / 20);
const isOn = (v) => v === '1' || v === 'true' || v === 'on' || v === true;

// ---------------------------------------------------------------- fontes

const FALLBACKS = {
  'aptos': ['Aptos', 'Segoe UI'],
  'aptos display': ['Aptos Display', 'Segoe UI'],
  'aptos narrow': ['Aptos Narrow', 'Arial Narrow', 'Segoe UI'],
  'helvetica': ['Helvetica', 'Arial'],
  'helvetica neue': ['Helvetica Neue', 'Arial'],
  'times': ['Times', 'Times New Roman'],
  'menlo': ['Menlo', 'Consolas'],
  'monaco': ['Monaco', 'Consolas'],
  'courier': ['Courier', 'Courier New'],
  'liberation serif': ['Liberation Serif', 'Times New Roman'],
  'liberation sans': ['Liberation Sans', 'Arial'],
  'liberation mono': ['Liberation Mono', 'Courier New'],
  'dejavu sans': ['DejaVu Sans', 'Verdana'],
  'noto sans': ['Noto Sans', 'Segoe UI'],
  'noto serif': ['Noto Serif', 'Georgia'],
  'open sans': ['Open Sans', 'Segoe UI'],
  'roboto': ['Roboto', 'Segoe UI'],
  'arial nova': ['Arial Nova', 'Arial'],
  'garamond': ['Garamond', 'Georgia'],
  'book antiqua': ['Book Antiqua', 'Palatino Linotype', 'Georgia'],
  'century gothic': ['Century Gothic', 'Segoe UI'],
  'calibri': ['Calibri', 'Carlito', 'Segoe UI'],
  'cambria': ['Cambria', 'Caladea', 'Georgia'],
  'ms mincho': ['MS Mincho', 'Yu Mincho'],
  'ms gothic': ['MS Gothic', 'Yu Gothic'],
};

const SERIF = /times|georgia|cambria|garamond|book|palatino|roman|serif|minion|baskerville|bodoni|century schoolbook|constantia|caladea|antiqua/i;
const MONO = /courier|consolas|mono|lucida console|menlo|monaco|code|fixed/i;

export function fontStack(name) {
  const key = String(name).toLowerCase();
  const list = FALLBACKS[key] || [name];
  let generic = 'sans-serif';
  const ft = R.model && R.model.fontTable && R.model.fontTable.get(name);
  if (ft && ft.family) generic = { roman: 'serif', swiss: 'sans-serif', modern: 'monospace', script: 'cursive', decorative: 'fantasy' }[ft.family] || generic;
  else if (MONO.test(name)) generic = 'monospace';
  else if (SERIF.test(name)) generic = 'serif';
  return list.map((f) => '"' + String(f).replace(/["\\;{}]/g, '') + '"').join(',') + ',' + generic;
}

const factorCache = new Map();

export function fontFactor(name) {
  if (!name) return 1.22;
  const key = String(name).toLowerCase();
  if (factorCache.has(key)) return factorCache.get(key);
  let f = 1.22;
  if (typeof document !== 'undefined' && document.body) {
    const probe = document.createElement('div');
    probe.style.cssText = 'position:absolute;left:-9999px;top:0;visibility:hidden;white-space:nowrap;line-height:normal;font-size:100px;font-family:' + fontStack(name);
    probe.textContent = 'Hgjé';
    document.body.appendChild(probe);
    const h = probe.getBoundingClientRect().height;
    probe.remove();
    if (h > 40 && h < 400) f = Math.round(h) / 100;
  }
  factorCache.set(key, f);
  return f;
}

export function fontName(a) {
  if (!a) return null;
  const theme = R.model && R.model.theme;
  const t = a.asciiTheme || a.hAnsiTheme;
  if (t && theme) {
    const f = theme.font(t);
    if (f) return f;
  }
  return a.ascii || a.hAnsi || a.cs || a.eastAsia || null;
}

export function fontCss(a) {
  const n = fontName(a);
  if (!n) return '';
  return 'font-family:' + fontStack(n) + ';--lf:' + fontFactor(n);
}

// ---------------------------------------------------------------- cores

export function themeHex(name, tint, shade) {
  const theme = R.model && R.model.theme;
  if (!theme) return null;
  return theme.hex(name, tint, shade);
}

export function colorOf(a) {
  if (!a) return null;
  if (a.val && a.val !== 'auto') {
    const h = normHex(a.val);
    if (h) return h;
  }
  if (a.themeColor) {
    const h = themeHex(a.themeColor, a.themeTint, a.themeShade);
    if (h) return h;
  }
  if (a.val === 'auto') return 'auto';
  return null;
}

export function colorCss(a) {
  const c = colorOf(a);
  if (!c) return '';
  if (c === 'auto') return 'color:var(--auto-color,#000)';
  return 'color:' + c;
}

export function highlightCss(val) {
  if (!val || val === 'none') return 'background-color:transparent';
  return 'background-color:' + (HIGHLIGHT[val] || '#FFFF00');
}

export function shdColor(shd) {
  if (!shd) return null;
  const val = shd.val;
  if (val === 'nil') return null;
  let fill = shd.fill && shd.fill !== 'auto' ? normHex(shd.fill) : null;
  if (!fill && shd.themeFill) fill = themeHex(shd.themeFill, shd.themeFillTint, shd.themeFillShade);
  let color = shd.color && shd.color !== 'auto' ? normHex(shd.color) : null;
  if (!color && shd.themeColor) color = themeHex(shd.themeColor, shd.themeTint, shd.themeShade);
  if (val === 'solid') return color || '#000000';
  const pct = /^pct(\d+)$/.exec(val || '');
  if (pct) return mix(fill || '#FFFFFF', color || '#000000', Number(pct[1]) / 100);
  return fill;
}

export function shdCss(a) {
  const c = shdColor(a);
  return c ? 'background-color:' + c : '';
}

// Sombreamento de parágrafo pinta só as linhas: o espaço antes e depois fica de
// fora, como no Word. Com borda, o fundo vai até a borda.
function paraShdCss(shd, bordered) {
  const css = shdCss(shd);
  return css && !bordered ? css + ';background-clip:content-box' : css;
}

export function isDarkShd(shd) {
  const c = shdColor(shd);
  return !!c && isDark(c);
}

export function underlineCss(a) {
  const v = (a && a.val) || 'single';
  if (v === 'none') return 'text-decoration-line:none';
  let style = 'solid';
  let thick = null;
  if (v === 'double') style = 'double';
  else if (v === 'thick') thick = '0.12em';
  else if (/^dotted/.test(v)) style = 'dotted';
  else if (/dash|dot/i.test(v)) style = 'dashed';
  else if (/wav/i.test(v)) style = 'wavy';
  if (/Heavy$/.test(v)) thick = '0.12em';
  let css = 'text-decoration-line:underline;text-decoration-style:' + style;
  if (thick) css += ';text-decoration-thickness:' + thick;
  const col = a && a.color && a.color !== 'auto' ? normHex(a.color) : (a && a.themeColor ? themeHex(a.themeColor, a.themeTint, a.themeShade) : null);
  if (col) css += ';text-decoration-color:' + col;
  return css;
}

// ---------------------------------------------------------------- bordas

export function borderCss(b) {
  if (!b) return null;
  const v = b.val;
  if (!v || v === 'nil' || v === 'none') return 'none';
  const sz = Math.max(2, Number(b.sz) || 4) / 8;
  let style = 'solid';
  let w = sz;
  if (/double|triple|thinThick|thickThin/i.test(v)) { style = 'double'; w = Math.max(sz * 3, 2.25); }
  else if (/dotDash|dotDotDash|dashDot/i.test(v)) style = 'dashed';
  else if (/^dot/i.test(v)) style = 'dotted';
  else if (/dash/i.test(v)) style = 'dashed';
  else if (v === 'inset') style = 'inset';
  else if (v === 'outset') style = 'outset';
  else if (/Emboss/i.test(v)) style = 'ridge';
  else if (/Engrave/i.test(v)) style = 'groove';
  let color = b.color && b.color !== 'auto' ? normHex(b.color) : null;
  if (!color && b.themeColor) color = themeHex(b.themeColor, b.themeTint, b.themeShade);
  return r2(w) + 'pt ' + style + ' ' + (color || '#000000');
}

// ---------------------------------------------------------------- parágrafo

export function jcCss(jc) {
  switch (jc) {
    case 'center': return 'center';
    case 'right': case 'end': return 'right';
    case 'both': case 'distribute': case 'thaiDistribute': case 'lowKashida': case 'mediumKashida': case 'highKashida': return 'justify';
    default: return 'left';
  }
}

export function indCss(ind) {
  if (!ind) return '';
  const out = [];
  if (ind.left != null) out.push('margin-left:' + tw(ind.left) + 'pt');
  if (ind.right != null) out.push('margin-right:' + tw(ind.right) + 'pt');
  if (ind.hanging != null) out.push('text-indent:' + -tw(ind.hanging) + 'pt');
  else if (ind.firstLine != null) out.push('text-indent:' + tw(ind.firstLine) + 'pt');
  return out.join(';');
}

// Espaçamento antes e depois vira padding, ou margem quando há borda,
// porque no Word o espaço fica do lado de fora da borda.
export function spacingCss(sp, asMargin) {
  if (!sp) return '';
  const out = [];
  const box = asMargin ? 'margin' : 'padding';
  if (isOn(sp.beforeAutospacing)) out.push(box + '-top:14pt');
  else if (sp.before != null) out.push(box + '-top:' + Math.max(0, tw(sp.before)) + 'pt');
  else if (sp.beforeLines != null) out.push(box + '-top:' + r2((Number(sp.beforeLines) / 100) * 12) + 'pt');
  if (isOn(sp.afterAutospacing)) out.push(box + '-bottom:14pt');
  else if (sp.after != null) out.push(box + '-bottom:' + Math.max(0, tw(sp.after)) + 'pt');
  else if (sp.afterLines != null) out.push(box + '-bottom:' + r2((Number(sp.afterLines) / 100) * 12) + 'pt');
  if (sp.line != null && sp.line !== '') {
    const rule = sp.lineRule || 'auto';
    const v = Number(sp.line);
    if (rule === 'auto') out.push('--lh:' + r2(v / 240) + ';--lhx:initial');
    else if (rule === 'exact') out.push('--lhx:' + tw(Math.abs(v)) + 'pt');
    else out.push('--lhx:max(' + tw(v) + 'pt,calc(var(--lf,1.22) * 1em))');
  }
  return out.join(';');
}

export function hasVisibleBorder(b) {
  return !!b && ['top', 'left', 'bottom', 'right'].some((s) => b[s] && b[s].val && b[s].val !== 'nil' && b[s].val !== 'none');
}

export function pBdrCss(b) {
  if (!b) return '';
  const out = [];
  for (const side of ['top', 'left', 'bottom', 'right']) {
    const s = b[side];
    if (!s) continue;
    const css = borderCss(s);
    out.push('border-' + side + ':' + css);
    if (css !== 'none') out.push('padding-' + side + ':' + (Number(s.space) || 0) + 'pt');
  }
  return out.join(';');
}

// CSS de um estilo de parágrafo; effSpacing inclui os padrões do documento
export function paraCss(pPr, effSpacing) {
  if (!pPr) return '';
  const out = [];
  if (pPr.jc) out.push('text-align:' + jcCss(pPr.jc));
  if (pPr.ind) out.push(indCss(pPr.ind));
  const bordered = hasVisibleBorder(pPr.pBdr);
  if (bordered) {
    out.push(spacingCss({ ...(effSpacing || {}), ...(pPr.spacing || {}) }, true));
    out.push('padding-top:0;padding-bottom:0');
  } else if (pPr.spacing) out.push(spacingCss(pPr.spacing));
  if (pPr.shd) out.push(paraShdCss(pPr.shd, bordered));
  if (pPr.pBdr) out.push(pBdrCss(pPr.pBdr));
  return out.filter(Boolean).join(';');
}

export function effectiveNumPr(a) {
  const m = R.model;
  if (!m || !m.styles) return null;
  const st = m.styles.resolvePara(a.styleId || m.styles.defaultPara).pPr;
  let numId = a.numId;
  let direct = true;
  if (numId == null) { numId = st.numId; direct = false; }
  let ilvl = a.ilvl;
  if (ilvl == null) ilvl = st.ilvl ?? 0;
  if (numId == null || numId === '0') return null;
  return { numId: String(numId), ilvl: Number(ilvl) || 0, direct, styleInd: st.ind };
}

function numberingInd(a) {
  const m = R.model;
  if (!m || !m.numbering) return null;
  const np = effectiveNumPr(a);
  if (!np) return null;
  const lvl = m.numbering.level(np.numId, np.ilvl);
  const li = lvl && lvl.pPr && lvl.pPr.ind;
  if (!li) return null;
  if (np.direct) return li;
  const si = np.styleInd || {};
  const r = { ...li };
  for (const k of Object.keys(si)) delete r[k];
  if (si.firstLine != null) delete r.hanging;
  if (si.hanging != null) delete r.firstLine;
  return Object.keys(r).length ? r : null;
}

// Informação de tabulação usada pelo layout, gravada no DOM do parágrafo
export function tabInfo(attrs) {
  const m = R.model;
  if (!m || !m.styles) return null;
  const eff = m.styles.effectivePPr(attrs, m.numbering);
  const stops = (eff.tabs || [])
    .filter((t) => t.val && t.val !== 'clear')
    .map((t) => [Number(t.pos) || 0, t.val, t.leader || 'none'])
    .sort((a, b) => a[0] - b[0]);
  const ind = eff.ind || {};
  return JSON.stringify({
    s: stops,
    l: Number(ind.left) || 0,
    h: ind.hanging != null ? Number(ind.hanging) || 0 : 0,
    f: ind.firstLine != null ? Number(ind.firstLine) || 0 : 0,
    d: (m.settings && m.settings.defaultTabStop) || 708,
  });
}

export function needsTabInfo(node) {
  if (effectiveNumPr(node.attrs)) return true;
  let found = false;
  node.forEach((c) => { if (c.type.name === 'tab') found = true; });
  return found;
}

// callouts do Obsidian: cor, ícone e nome mostrado quando não há título
const CALLOUTS = {
  note: ['blue', '✎', 'Nota'], info: ['blue', 'ⓘ', 'Info'], todo: ['blue', '☐', 'Tarefa'],
  abstract: ['cyan', '☰', 'Resumo'], summary: ['cyan', '☰', 'Resumo'], tldr: ['cyan', '☰', 'Resumo'],
  tip: ['cyan', '✦', 'Dica'], hint: ['cyan', '✦', 'Dica'], important: ['cyan', '✦', 'Importante'],
  success: ['green', '✓', 'Sucesso'], check: ['green', '✓', 'Ok'], done: ['green', '✓', 'Feito'],
  question: ['orange', '?', 'Pergunta'], help: ['orange', '?', 'Ajuda'], faq: ['orange', '?', 'Perguntas'],
  warning: ['orange', '⚠', 'Atenção'], caution: ['orange', '⚠', 'Cuidado'], attention: ['orange', '⚠', 'Atenção'],
  failure: ['red', '✗', 'Falha'], fail: ['red', '✗', 'Falha'], missing: ['red', '✗', 'Faltando'],
  danger: ['red', '⚡', 'Perigo'], error: ['red', '⚡', 'Erro'], bug: ['red', '✱', 'Bug'],
  example: ['purple', '☷', 'Exemplo'], quote: ['gray', '❝', 'Citação'], cite: ['gray', '❝', 'Citação'],
};

// nomes em português que aparecem nas notas
const CALLOUT_ALIASES = {
  nota: 'note', 'observação': 'note', observacao: 'note', obs: 'note', aviso: 'warning', 'atenção': 'warning', atencao: 'warning',
  dica: 'tip', perigo: 'danger', erro: 'error', exemplo: 'example', importante: 'important', sucesso: 'success',
  pergunta: 'question', 'dúvida': 'question', duvida: 'question', 'citação': 'quote', citacao: 'quote', resumo: 'summary',
  tarefa: 'todo', cuidado: 'caution', falha: 'failure',
};

export function calloutInfo(type) {
  const t = String(type || '').toLowerCase();
  const alias = CALLOUT_ALIASES[t];
  if (alias) { const k = CALLOUTS[alias]; return { color: k[0], icon: k[1], label: t[0].toUpperCase() + t.slice(1) }; }
  const k = CALLOUTS[t];
  if (k) return { color: k[0], icon: k[1], label: k[2] };
  return { color: 'blue', icon: 'ⓘ', label: t ? t[0].toUpperCase() + t.slice(1) : 'Nota' };
}

export function paraClasses(node) {
  const a = node.attrs;
  const styles = R.model && R.model.styles;
  let sid = a.styleId;
  if (!sid || !styles || !styles.get(sid)) sid = styles ? styles.defaultPara : 'Normal';
  let cls = 'p ' + pClass(sid);
  if (a.shd && isDarkShd(a.shd)) cls += ' dark-bg';
  if (a.md && a.md.callout) cls += ' callout co-' + calloutInfo(a.md.callout).color + (a.md.head ? ' callout-head' : '');
  return cls;
}

export function paraDirectCss(a) {
  const out = [];
  if (a.jc) {
    out.push('text-align:' + jcCss(a.jc));
    if (a.jc === 'distribute') out.push('text-align-last:justify');
  }
  let ind = a.ind;
  const ni = numberingInd(a);
  if (ni) ind = mergeInd(ni, ind);
  if (ind) out.push(indCss(ind));
  const styles = R.model && R.model.styles;
  const stPPr = styles ? styles.resolvePara(a.styleId || styles.defaultPara).pPr : {};
  const styleBorder = hasVisibleBorder(stPPr.pBdr);
  const directBorder = hasVisibleBorder(a.pBdr);
  if (directBorder && !styleBorder) {
    const eff = { ...((styles && styles.defaults.pPr.spacing) || {}), ...(stPPr.spacing || {}), ...(a.spacing || {}) };
    out.push(spacingCss(eff, true));
    out.push('padding-top:0;padding-bottom:0');
  } else if (a.spacing) out.push(spacingCss(a.spacing, styleBorder));
  if (a.shd) out.push(paraShdCss(a.shd, directBorder || styleBorder));
  if (a.pBdr) out.push(pBdrCss(a.pBdr));
  if (a.markRPr) {
    const sz = a.markRPr.find((m) => m.type === 'sz');
    if (sz) out.push('--pm-sz:' + sz.attrs.val / 2 + 'pt');
  }
  return out.filter(Boolean).join(';');
}

// ---------------------------------------------------------------- execução (estilos)

export function runCss(r) {
  if (!r) return '';
  const out = [];
  if (r.rFonts) out.push(fontCss(r.rFonts.a));
  if (r.sz) out.push('font-size:' + r.sz.val / 2 + 'pt');
  if (r.b) out.push('font-weight:' + (r.b.val ? 'bold' : 'normal'));
  if (r.i) out.push('font-style:' + (r.i.val ? 'italic' : 'normal'));
  if (r.color) out.push(colorCss(r.color.a));
  const lines = [];
  if (r.u && r.u.a && r.u.a.val !== 'none') lines.push('underline');
  if ((r.strike && r.strike.val) || (r.dstrike && r.dstrike.val)) lines.push('line-through');
  if (lines.length) {
    out.push('text-decoration-line:' + lines.join(' '));
    if (r.u && r.u.a && r.u.a.val !== 'none') {
      const u = underlineCss(r.u.a).split(';').filter((d) => !d.startsWith('text-decoration-line'));
      out.push(...u);
    }
  }
  if (r.caps) out.push('text-transform:' + (r.caps.val ? 'uppercase' : 'none'));
  if (r.smallCaps) out.push('font-variant:' + (r.smallCaps.val ? 'small-caps' : 'normal'));
  if (r.highlight) out.push(highlightCss(r.highlight.val));
  else if (r.shd) out.push(shdCss(r.shd.a));
  if (r.spacing) out.push('letter-spacing:' + r.spacing.val / 20 + 'pt');
  return out.filter(Boolean).join(';');
}

// ---------------------------------------------------------------- tabelas

export function marPt(m) {
  if (!m) return null;
  if (m.type === 'nil') return 0;
  const w = Number(m.w);
  if (!Number.isFinite(w)) return null;
  return r2(w / 20);
}

export function tableVarsCss(tblPr) {
  if (!tblPr) return '';
  const out = [];
  if (tblPr.tblBorders) {
    for (const side of ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']) {
      const b = tblPr.tblBorders[side];
      if (b) out.push('--tb-' + side + ':' + borderCss(b));
    }
  }
  if (tblPr.tblCellMar) {
    for (const side of ['top', 'left', 'bottom', 'right']) {
      const v = marPt(tblPr.tblCellMar[side]);
      if (v != null) out.push('--tc-' + side + ':' + v + 'pt');
    }
  }
  return out.join(';');
}

export function tableDirectCss(a) {
  const out = [tableVarsCss(a)];
  if (a.shd) out.push(shdCss(a.shd));
  if (a.grid && a.grid.length) {
    const total = a.grid.reduce((s, w) => s + (Number(w) || 0), 0);
    if (total > 0) out.push('width:' + r2(total / 15) + 'px');
  }
  return out.filter(Boolean).join(';');
}

export function tableClasses(node) {
  const styles = R.model && R.model.styles;
  let sid = node.attrs.styleId;
  if (!sid || !styles || !styles.get(sid)) sid = styles ? styles.defaultTable : null;
  return 'tbl' + (sid ? ' ' + tClass(sid) : '');
}

export function tableWrapCss(a) {
  const out = [];
  const compat = (R.model && R.model.settings && R.model.settings.compat) || 15;
  let ind = a.tblInd && a.tblInd.type !== 'pct' ? Number(a.tblInd.w) || 0 : 0;
  if (compat < 15) {
    // em modos antigos o texto da célula fica alinhado à margem
    const styles = R.model && R.model.styles;
    const st = styles ? styles.resolveTable(a.styleId || styles.defaultTable) : null;
    const mar = (a.tblCellMar && a.tblCellMar.left) || (st && st.tblPr && st.tblPr.tblCellMar && st.tblPr.tblCellMar.left);
    ind -= mar ? Number(mar.w) || 108 : 108;
  }
  if (a.jc === 'center') out.push('display:flex;justify-content:center');
  else if (a.jc === 'right' || a.jc === 'end') out.push('display:flex;justify-content:flex-end');
  else if (ind) out.push('margin-left:' + r2(ind / 20) + 'pt');
  return out.join(';');
}

export function cellDirectCss(a) {
  const out = [];
  if (a.tcBorders) {
    for (const side of ['top', 'bottom', 'left', 'right']) {
      if (a.tcBorders[side]) out.push('border-' + side + ':' + borderCss(a.tcBorders[side]));
    }
  }
  if (a.shd) {
    const c = shdColor(a.shd);
    if (c) out.push('background-color:' + c);
  }
  if (a.vAlign) out.push('vertical-align:' + ({ center: 'middle', both: 'middle', bottom: 'bottom' }[a.vAlign] || 'top'));
  if (a.tcMar) {
    for (const side of ['top', 'right', 'bottom', 'left']) {
      const v = marPt(a.tcMar[side]);
      if (v != null) out.push('padding-' + side + ':' + v + 'pt');
    }
  }
  if (a.textDirection === 'tbRl' || a.textDirection === 'tbRlV') out.push('writing-mode:vertical-rl');
  else if (a.textDirection === 'btLr') out.push('writing-mode:sideways-lr');
  return out.join(';');
}

export function rowDirectCss(a) {
  if (a.height && a.hRule !== 'auto') return 'height:' + tw(a.height) + 'pt';
  return '';
}

// ---------------------------------------------------------------- desenhos e símbolos

export function drawingBoxCss(a) {
  const w = (Number(a.cx) || 0) / 9525;
  const h = (Number(a.cy) || 0) / 9525;
  const out = [];
  if (w > 0) out.push('width:' + r2(w) + 'px');
  if (h > 0) out.push('height:' + r2(h) + 'px');
  return out.join(';');
}

const SYMBOL_UNICODE = {
  0x22: '∀', 0x24: '∃', 0x27: '∋', 0x2A: '∗', 0x2D: '−', 0x40: '≅', 0x5C: '∴', 0x5E: '⊥', 0x60: '‾', 0x7E: '∼',
  0xA2: '′', 0xA3: '≤', 0xA4: '⁄', 0xA5: '∞', 0xA7: '♣', 0xA8: '♦', 0xA9: '♥', 0xAA: '♠', 0xAB: '↔', 0xAC: '←',
  0xAD: '↑', 0xAE: '→', 0xAF: '↓', 0xB0: '°', 0xB1: '±', 0xB2: '″', 0xB3: '≥', 0xB4: '×', 0xB5: '∝', 0xB6: '∂',
  0xB7: '•', 0xB8: '÷', 0xB9: '≠', 0xBA: '≡', 0xBB: '≈', 0xBC: '…', 0xC0: 'ℵ', 0xC4: '⊗', 0xC5: '⊕', 0xC6: '∅',
  0xC7: '∩', 0xC8: '∪', 0xC9: '⊃', 0xCA: '⊇', 0xCB: '⊄', 0xCC: '⊂', 0xCD: '⊆', 0xCE: '∈', 0xCF: '∉', 0xD0: '∠',
  0xD1: '∇', 0xD2: '®', 0xD3: '©', 0xD4: '™', 0xD5: '∏', 0xD6: '√', 0xD7: '⋅', 0xD8: '¬', 0xD9: '∧', 0xDA: '∨',
  0xDB: '⇔', 0xDC: '⇐', 0xDD: '⇑', 0xDE: '⇒', 0xDF: '⇓', 0xE0: '◊', 0xE1: '〈', 0xE5: '∑', 0xF1: '〉', 0xF2: '∫',
};
const GREEK_UP = 'ΑΒΧΔΕΦΓΗΙϑΚΛΜΝΟΠΘΡΣΤΥςΩΞΨΖ';
const GREEK_LO = 'αβχδεφγηιϕκλμνοπθρστυϖωξψζ';

export function symChar(font, char) {
  const code = parseInt(char || '0', 16) || 0;
  const low = code >= 0xF000 ? code - 0xF000 : code;
  const f = (font || '').toLowerCase();
  if (f === 'symbol') {
    if (low >= 0x41 && low <= 0x5A) return { text: GREEK_UP[low - 0x41], font: null };
    if (low >= 0x61 && low <= 0x7A) return { text: GREEK_LO[low - 0x61], font: null };
    if (SYMBOL_UNICODE[low]) return { text: SYMBOL_UNICODE[low], font: null };
  }
  if (f.startsWith('wingdings') || f === 'webdings' || f === 'symbol') return { text: String.fromCharCode(0xF000 + low), font };
  return { text: String.fromCharCode(code), font: font || null };
}
