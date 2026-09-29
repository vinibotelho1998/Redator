// Estilos do documento (styles.xml): herança, resolução e geração de CSS.
import { parseXml, children, child, wa, isW, onOff, NS } from '../util/xml.js';
import { parsePPr, parseRPr, parseTblPr, parseTrPr, parseTcPr, mergePPr, mergeRun, runMap, RPR_ORDER } from './props.js';
import { runCss, paraCss, pClass, cClass, tClass, tableVarsCss, borderCss, shdColor, isDarkShd, marPt, hasVisibleBorder } from '../editor/render.js';

const PT_NAMES = {
  'normal': 'Normal',
  'title': 'Título',
  'subtitle': 'Subtítulo',
  'no spacing': 'Sem Espaçamento',
  'quote': 'Citação',
  'intense quote': 'Citação Intensa',
  'list paragraph': 'Parágrafo da Lista',
  'strong': 'Forte',
  'emphasis': 'Ênfase',
  'subtle emphasis': 'Ênfase Sutil',
  'intense emphasis': 'Ênfase Intensa',
  'subtle reference': 'Referência Sutil',
  'intense reference': 'Referência Intensa',
  'book title': 'Título do Livro',
  'hyperlink': 'Hiperlink',
  'followedhyperlink': 'Hiperlink Visitado',
  'table grid': 'Tabela com grade',
  'normal table': 'Tabela normal',
  'table normal': 'Tabela normal',
  'header': 'Cabeçalho',
  'footer': 'Rodapé',
  'footnote text': 'Texto de nota de rodapé',
  'footnote reference': 'Ref. de nota de rodapé',
  'endnote text': 'Texto de nota de fim',
  'endnote reference': 'Ref. de nota de fim',
  'toc heading': 'Cabeçalho do Sumário',
  'caption': 'Legenda',
  'balloon text': 'Texto de balão',
  'list bullet': 'Com marcadores',
  'list number': 'Numerada',
  'annotation text': 'Texto de comentário',
  'annotation reference': 'Referência de comentário',
  'placeholder text': 'Texto do Espaço Reservado',
  'default paragraph font': 'Fonte parág. padrão',
  'no list': 'Sem lista',
  'normal (web)': 'Normal (Web)',
  'body text': 'Corpo de texto',
  'block text': 'Texto em bloco',
  'plain text': 'Texto sem formatação',
  'code': 'Código',
  'block quote': 'Citação em Bloco',
  'inline code': 'Código Embutido',
  'grid table 1 light': 'Tabela de Grade 1 Clara',
  'grid table 4 accent 1': 'Tabela de Grade 4 - Ênfase 1',
  'grid table 5 dark accent 1': 'Tabela de Grade 5 Escura - Ênfase 1',
  'list table 3 accent 2': 'Tabela de Lista 3 - Ênfase 2',
  'plain table 1': 'Tabela Simples 1',
};

export const COND_ORDER = ['wholeTable', 'band1Vert', 'band2Vert', 'band1Horz', 'band2Horz', 'firstCol', 'lastCol', 'firstRow', 'lastRow', 'neCell', 'nwCell', 'seCell', 'swCell'];

function parseStyleEl(s) {
  const nameEl = child(s, 'name');
  const st = {
    id: wa(s, 'styleId'),
    type: wa(s, 'type') || 'paragraph',
    isDefault: wa(s, 'default') === '1' || wa(s, 'default') === 'true' || wa(s, 'default') === 'on',
    custom: wa(s, 'customStyle') === '1',
    name: nameEl ? wa(nameEl, 'val') : null,
    basedOn: wa(child(s, 'basedOn'), 'val'),
    next: wa(child(s, 'next'), 'val'),
    link: wa(child(s, 'link'), 'val'),
    uiPriority: child(s, 'uiPriority') ? parseInt(wa(child(s, 'uiPriority'), 'val'), 10) : 99,
    qFormat: !!child(s, 'qFormat'),
    semiHidden: !!child(s, 'semiHidden') && onOff(child(s, 'semiHidden')),
    unhideWhenUsed: !!child(s, 'unhideWhenUsed'),
    hidden: !!child(s, 'hidden') && onOff(child(s, 'hidden')),
    pPr: child(s, 'pPr') ? parsePPr(child(s, 'pPr')) : null,
    rPr: child(s, 'rPr') ? parseRPr(child(s, 'rPr')) : null,
    tblPr: child(s, 'tblPr') ? parseTblPr(child(s, 'tblPr')) : null,
    trPr: child(s, 'trPr') ? parseTrPr(child(s, 'trPr')) : null,
    tcPr: child(s, 'tcPr') ? parseTcPr(child(s, 'tcPr')) : null,
    cond: {},
  };
  if (!st.name) st.name = st.id;
  for (const t of children(s, 'tblStylePr')) {
    st.cond[wa(t, 'type')] = {
      pPr: child(t, 'pPr') ? parsePPr(child(t, 'pPr')) : null,
      rPr: child(t, 'rPr') ? parseRPr(child(t, 'rPr')) : null,
      tblPr: child(t, 'tblPr') ? parseTblPr(child(t, 'tblPr')) : null,
      trPr: child(t, 'trPr') ? parseTrPr(child(t, 'trPr')) : null,
      tcPr: child(t, 'tcPr') ? parseTcPr(child(t, 'tcPr')) : null,
    };
  }
  return st;
}

function mergeObj(a, b) {
  if (!a) return b ? { ...b } : null;
  if (!b) return { ...a };
  const r = { ...a };
  for (const [k, v] of Object.entries(b)) {
    if (v == null) continue;
    if (k === 'tblBorders' || k === 'tcBorders' || k === 'tblCellMar' || k === 'tcMar') r[k] = { ...(r[k] || {}), ...v };
    else if (k === 'extra') r[k] = [...(r[k] || []), ...v];
    else r[k] = v;
  }
  return r;
}

export class StyleSheet {
  constructor() {
    this.styles = new Map();
    this.defaults = { pPr: {}, r: {} };
    this.defaultPara = 'Normal';
    this.defaultChar = null;
    this.defaultTable = null;
    this.doc = null;
    this.text = null;
    this.dirty = false;
    this.cache = new Map();
  }

  static parse(text) {
    const s = new StyleSheet();
    if (!text) return s;
    s.text = text;
    try { s.doc = parseXml(text); } catch { return s; }
    s.load();
    return s;
  }

  load() {
    const root = this.doc.documentElement;
    this.styles.clear();
    this.cache.clear();
    const dd = child(root, 'docDefaults');
    if (dd) {
      const rp = child(child(dd, 'rPrDefault'), 'rPr');
      const pp = child(child(dd, 'pPrDefault'), 'pPr');
      this.defaults = { pPr: pp ? parsePPr(pp) : {}, r: rp ? runMap(parseRPr(rp)) : {} };
    }
    for (const el of children(root, 'style')) {
      const st = parseStyleEl(el);
      if (!st.id) continue;
      this.styles.set(st.id, st);
      if (st.isDefault) {
        if (st.type === 'paragraph') this.defaultPara = st.id;
        else if (st.type === 'character') this.defaultChar = st.id;
        else if (st.type === 'table') this.defaultTable = st.id;
      }
    }
    if (!this.styles.has(this.defaultPara)) {
      const n = this.findByName('Normal');
      if (n) this.defaultPara = n.id;
    }
  }

  get(id) {
    return id == null ? null : this.styles.get(id) || null;
  }

  findByName(name) {
    const low = String(name).toLowerCase();
    for (const s of this.styles.values()) if ((s.name || '').toLowerCase() === low) return s;
    return null;
  }

  headingId(level) {
    const s = this.findByName('heading ' + level);
    if (s) return s.id;
    const byId = this.get('Heading' + level) || this.get('Ttulo' + level);
    return byId ? byId.id : null;
  }

  displayName(st) {
    if (!st) return '';
    const n = st.name || st.id;
    const low = n.toLowerCase();
    let m = /^heading (\d)$/.exec(low);
    if (m) return 'Título ' + m[1];
    m = /^toc (\d)$/.exec(low);
    if (m) return 'Sumário ' + m[1];
    return PT_NAMES[low] || n;
  }

  headingLevel(styleId) {
    const pPr = this.resolvePara(styleId).pPr;
    if (pPr.outlineLvl != null && pPr.outlineLvl < 9) return pPr.outlineLvl + 1;
    const st = this.get(styleId);
    if (!st) return 0;
    const m = /^heading (\d)$/i.exec(st.name || '');
    return m ? Number(m[1]) : 0;
  }

  chain(id) {
    const out = [];
    const seen = new Set();
    let cur = this.get(id);
    while (cur && !seen.has(cur.id)) {
      seen.add(cur.id);
      out.unshift(cur);
      cur = cur.basedOn ? this.get(cur.basedOn) : null;
    }
    return out;
  }

  resolvePara(id) {
    const key = 'p:' + id;
    if (this.cache.has(key)) return this.cache.get(key);
    let pPr = {};
    let r = {};
    for (const st of this.chain(id)) {
      if (st.pPr) pPr = mergePPr(pPr, st.pPr);
      if (st.rPr) r = mergeRun(r, runMap(st.rPr));
    }
    delete pPr.styleId;
    const res = { pPr, r };
    this.cache.set(key, res);
    return res;
  }

  resolveChar(id) {
    const key = 'c:' + id;
    if (this.cache.has(key)) return this.cache.get(key);
    let r = {};
    for (const st of this.chain(id)) if (st.rPr) r = mergeRun(r, runMap(st.rPr));
    const res = { r };
    this.cache.set(key, res);
    return res;
  }

  resolveTable(id) {
    const key = 't:' + id;
    if (this.cache.has(key)) return this.cache.get(key);
    let tblPr = null, tcPr = null, trPr = null, pPr = {}, r = {};
    const cond = {};
    for (const st of this.chain(id)) {
      tblPr = mergeObj(tblPr, st.tblPr);
      tcPr = mergeObj(tcPr, st.tcPr);
      trPr = mergeObj(trPr, st.trPr);
      if (st.pPr) pPr = mergePPr(pPr, st.pPr);
      if (st.rPr) r = mergeRun(r, runMap(st.rPr));
      for (const [type, c] of Object.entries(st.cond || {})) {
        const cur = cond[type] || { pPr: null, r: {}, tcPr: null, tblPr: null };
        cond[type] = {
          pPr: c.pPr ? mergePPr(cur.pPr || {}, c.pPr) : cur.pPr,
          r: c.rPr ? mergeRun(cur.r, runMap(c.rPr)) : cur.r,
          tcPr: mergeObj(cur.tcPr, c.tcPr),
          tblPr: mergeObj(cur.tblPr, c.tblPr),
        };
      }
    }
    const res = { tblPr: tblPr || {}, tcPr: tcPr || {}, trPr: trPr || {}, pPr, r, cond };
    this.cache.set(key, res);
    return res;
  }

  // Propriedades efetivas de um parágrafo: padrões, estilo, numeração e formatação direta
  effectivePPr(attrs, numbering) {
    const sid = attrs.styleId && this.get(attrs.styleId) ? attrs.styleId : this.defaultPara;
    const st = this.resolvePara(sid).pPr;
    let eff = mergePPr(this.defaults.pPr, st);
    const numId = attrs.numId != null ? attrs.numId : eff.numId;
    const ilvl = attrs.ilvl != null ? attrs.ilvl : (eff.ilvl || 0);
    if (numbering && numId != null && numId !== '0') {
      const lvl = numbering.level(numId, ilvl);
      if (lvl && lvl.pPr) {
        const fromLevel = {};
        for (const k of ['ind', 'tabs', 'spacing', 'jc', 'keepNext', 'keepLines']) if (lvl.pPr[k] != null) fromLevel[k] = lvl.pPr[k];
        if (attrs.numId != null) eff = mergePPr(eff, fromLevel);
        else eff = mergePPr(mergePPr(this.defaults.pPr, fromLevel), st);
      }
    }
    const direct = {};
    for (const [k, v] of Object.entries(attrs)) if (v != null && k !== 'styleId' && k !== 'markRPr' && k !== 'extra' && k !== 'sectPr') direct[k] = v;
    eff = mergePPr(eff, direct);
    eff.styleId = sid;
    eff.numId = numId;
    eff.ilvl = ilvl;
    return eff;
  }

  // Propriedades de execução efetivas, dados o parágrafo e as marcas diretas
  effectiveRun(paraAttrs, marks) {
    const sid = paraAttrs && paraAttrs.styleId && this.get(paraAttrs.styleId) ? paraAttrs.styleId : this.defaultPara;
    let r = mergeRun({}, this.defaults.r);
    r = mergeRun(r, this.resolvePara(sid).r);
    const rs = marks && marks.find((m) => (m.type.name || m.type) === 'rStyle');
    if (rs && rs.attrs.val) r = mergeRun(r, this.resolveChar(rs.attrs.val).r);
    if (marks) r = mergeRun(r, runMap(marks.filter((m) => (m.type.name || m.type) !== 'rStyle')));
    return r;
  }

  quickStyles() {
    const out = [];
    for (const s of this.styles.values()) {
      if ((s.type !== 'paragraph' && s.type !== 'character') || s.hidden) continue;
      if (!s.qFormat || s.semiHidden) continue;
      out.push(s);
    }
    const prio = (s) => (s.id === this.defaultPara ? -1 : s.uiPriority);
    out.sort((a, b) => prio(a) - prio(b) || this.displayName(a).localeCompare(this.displayName(b), 'pt-BR'));
    return out;
  }

  allStyles(type) {
    const out = [];
    for (const s of this.styles.values()) if (!type || s.type === type) out.push(s);
    out.sort((a, b) => this.displayName(a).localeCompare(this.displayName(b), 'pt-BR'));
    return out;
  }

  // Acrescenta um estilo a partir de XML (w:style)
  addStyleXml(xml) {
    this.ensureDoc();
    const el = parseXml('<w:root xmlns:w="' + NS.w + '">' + xml + '</w:root>').documentElement.firstElementChild;
    const st = parseStyleEl(el);
    if (this.styles.has(st.id)) return this.styles.get(st.id);
    const imported = this.doc.importNode(el, true);
    this.doc.documentElement.appendChild(imported);
    this.styles.set(st.id, st);
    this.cache.clear();
    this.dirty = true;
    return st;
  }

  // Altera a fonte e o tamanho padrão do documento (docDefaults)
  setDefaultFont(font, halfPoints) {
    this.ensureDoc();
    const root = this.doc.documentElement;
    const W = NS.w;
    const make = (parent, name, before) => {
      let el = child(parent, name);
      if (!el) {
        el = this.doc.createElementNS(W, 'w:' + name);
        parent.insertBefore(el, before || null);
      }
      return el;
    };
    const dd = make(root, 'docDefaults', root.firstElementChild);
    const rd = make(dd, 'rPrDefault', dd.firstElementChild);
    const rpr = make(rd, 'rPr');
    if (font) {
      const rf = make(rpr, 'rFonts', rpr.firstElementChild);
      for (const a of Array.from(rf.attributes)) if (a.localName !== 'eastAsia' && a.localName !== 'eastAsiaTheme' && a.localName !== 'hint') rf.removeAttributeNode(a);
      for (const k of ['ascii', 'hAnsi', 'cs']) rf.setAttributeNS(W, 'w:' + k, font);
      this.defaults.r.rFonts = { a: { ascii: font, hAnsi: font, cs: font } };
    }
    if (halfPoints) {
      const rf = child(rpr, 'rFonts');
      const sz = make(rpr, 'sz');
      sz.setAttributeNS(W, 'w:val', String(halfPoints));
      const szCs = make(rpr, 'szCs');
      szCs.setAttributeNS(W, 'w:val', String(halfPoints));
      this.defaults.r.sz = { val: halfPoints };
    }
    // mantém a ordem exigida pelo esquema
    const order = RPR_ORDER;
    const list = Array.from(rpr.children).sort((x, y) => {
      const ix = order.indexOf(x.localName);
      const iy = order.indexOf(y.localName);
      return (ix < 0 ? 999 : ix) - (iy < 0 ? 999 : iy);
    });
    for (const el of list) rpr.appendChild(el);
    this.cache.clear();
    this.dirty = true;
  }

  setDefaultParagraphSpacing(before, after, line) {
    this.ensureDoc();
    const root = this.doc.documentElement;
    const W = NS.w;
    let dd = child(root, 'docDefaults');
    if (!dd) { dd = this.doc.createElementNS(W, 'w:docDefaults'); root.insertBefore(dd, root.firstElementChild); }
    let pd = child(dd, 'pPrDefault');
    if (!pd) { pd = this.doc.createElementNS(W, 'w:pPrDefault'); dd.appendChild(pd); }
    let ppr = child(pd, 'pPr');
    if (!ppr) { ppr = this.doc.createElementNS(W, 'w:pPr'); pd.appendChild(ppr); }
    let sp = child(ppr, 'spacing');
    if (!sp) { sp = this.doc.createElementNS(W, 'w:spacing'); ppr.appendChild(sp); }
    const vals = { before: String(before), after: String(after), line: String(line), lineRule: 'auto' };
    for (const [k, v] of Object.entries(vals)) sp.setAttributeNS(W, 'w:' + k, v);
    this.defaults.pPr = { ...this.defaults.pPr, spacing: { ...(this.defaults.pPr.spacing || {}), ...vals } };
    this.cache.clear();
    this.dirty = true;
  }

  ensureDoc() {
    if (this.doc) return;
    this.doc = parseXml('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="' + NS.w + '"></w:styles>');
  }

  toXml() {
    if (!this.doc) return this.text;
    let s = new XMLSerializer().serializeToString(this.doc);
    if (!s.startsWith('<?xml')) s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + s;
    return s;
  }

  // ------------------------------------------------------------ CSS

  css() {
    let out = '';
    const dr = runCss(this.defaults.r);
    const effSpacing = this.defaults.pPr.spacing || {};
    const dp = paraCss(this.defaults.pPr, {});
    const lf = this.defaults.r.rFonts ? '' : 'font-family:"Times New Roman",serif;--lf:1.15;';
    const sz = this.defaults.r.sz ? '' : 'font-size:10pt;';
    out += '@layer docdefaults{.doc-body{' + lf + sz + dr + '}.doc-body p{' + dp + '}}\n';

    let t = '';
    for (const st of this.styles.values()) if (st.type === 'table') t += this.tableCss(st, effSpacing);
    out += '@layer tablestyle{' + t + '}\n';

    let p = '';
    for (const st of this.styles.values()) {
      if (st.type !== 'paragraph') continue;
      const res = this.resolvePara(st.id);
      const eff = { ...effSpacing, ...(res.pPr.spacing || {}) };
      const css = paraCss(res.pPr, eff) + ';' + runCss(res.r);
      p += '.' + pClass(st.id) + '{' + css + '}';
      if (hasVisibleBorder(res.pPr.pBdr)) p += '.' + pClass(st.id) + '+.' + pClass(st.id) + '{border-top:none}';
    }
    out += '@layer parastyle{' + p + '}\n';

    let c = '';
    for (const st of this.styles.values()) {
      if (st.type !== 'character') continue;
      const res = this.resolveChar(st.id);
      c += '.' + cClass(st.id) + '{' + runCss(res.r) + '}';
    }
    out += '@layer charstyle{' + c + '}\n';
    return out;
  }

  tableCss(st, effSpacing) {
    const res = this.resolveTable(st.id);
    const cls = '.' + tClass(st.id);
    let out = '';
    const vars = tableVarsCss(res.tblPr);
    const tshd = res.tblPr.shd ? shdColor(res.tblPr.shd) : null;
    out += cls + '{' + vars + (tshd ? ';background-color:' + tshd : '') + '}';
    const cell = cls + '>tbody>tr>td';
    const rc = runCss(res.r);
    const cellShd = res.tcPr && res.tcPr.shd ? shdColor(res.tcPr.shd) : null;
    if (rc || cellShd) out += cell + '{' + rc + (cellShd ? ';background-color:' + cellShd : '') + '}';
    const pc = paraCss(res.pPr, effSpacing);
    if (pc) out += cell + '>p{' + pc + '}';
    for (const type of COND_ORDER) {
      const c = res.cond[type];
      if (!c) continue;
      const sel = cell + '.cf-' + type;
      const decl = [];
      const r = runCss(c.r);
      if (r) decl.push(r);
      if (c.tcPr && c.tcPr.shd) {
        const col = shdColor(c.tcPr.shd);
        if (col) decl.push('background-color:' + col);
        if (isDarkShd(c.tcPr.shd)) decl.push('--auto-color:#FFFFFF');
      }
      if (c.tcPr && c.tcPr.tcMar) {
        for (const side of ['top', 'right', 'bottom', 'left']) {
          const v = marPt(c.tcPr.tcMar[side]);
          if (v != null) decl.push('padding-' + side + ':' + v + 'pt');
        }
      }
      if (decl.length) out += sel + '{' + decl.join(';') + '}';
      const b = c.tcPr && c.tcPr.tcBorders;
      if (b) out += this.condBorders(sel, type, b);
      if (c.pPr) {
        const ppc = paraCss(c.pPr, effSpacing);
        if (ppc) out += sel + '>p{' + ppc + '}';
      }
    }
    return out;
  }

  condBorders(sel, type, b) {
    const css = (x) => (x ? borderCss(x) : null);
    const rows = ['firstRow', 'lastRow', 'band1Horz', 'band2Horz', 'wholeTable'].includes(type);
    const cols = ['firstCol', 'lastCol', 'band1Vert', 'band2Vert'].includes(type);
    let out = '';
    const rule = (s, decls) => { const d = decls.filter((x) => x[1] != null).map((x) => x[0] + ':' + x[1]).join(';'); if (d) out += s + '{' + d + '}'; };
    if (rows) {
      rule(sel, [['border-top', css(b.top)], ['border-bottom', css(b.bottom)], ['border-left', css(b.insideV)], ['border-right', css(b.insideV)]]);
      rule(sel + '.c-first', [['border-left', css(b.left)]]);
      rule(sel + '.c-last', [['border-right', css(b.right)]]);
    } else if (cols) {
      rule(sel, [['border-left', css(b.left)], ['border-right', css(b.right)], ['border-top', css(b.insideH)], ['border-bottom', css(b.insideH)]]);
      rule(sel + '.r-first', [['border-top', css(b.top)]]);
      rule(sel + '.r-last', [['border-bottom', css(b.bottom)]]);
    } else {
      rule(sel, [['border-top', css(b.top)], ['border-bottom', css(b.bottom)], ['border-left', css(b.left)], ['border-right', css(b.right)]]);
    }
    return out;
  }
}
