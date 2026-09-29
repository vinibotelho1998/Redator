// Rich Text Format (.rtf): leitura e escrita.
// A leitura entende o que o Word e o WordPad costumam gravar: fontes, cores,
// estilos de título, listas, tabelas, imagens, links, cabeçalho, rodapé e notas.
import { schema } from '../editor/schema.js';
import { R, effectiveNumPr, fontName, colorOf, shdColor } from '../editor/render.js';
import { NumberingCounter } from '../docx/numbering.js';
import { newHeaderFooter } from '../docx/writer.js';
import { MIME } from '../docx/reader.js';
import { blankModel, finalizeDoc, canvasToBytes } from './common.js';
import { HIGHLIGHT } from '../util/color.js';
import { mergeRun, mergePPr } from '../docx/props.js';

const nodes = schema.nodes;
const marks = schema.marks;

// ---------------------------------------------------------------- leitura

const DECODERS = new Map();
function decoderFor(cp) {
  const name = { 1250: 'windows-1250', 1251: 'windows-1251', 1252: 'windows-1252', 1253: 'windows-1253', 1254: 'windows-1254', 1255: 'windows-1255', 1256: 'windows-1256', 1257: 'windows-1257', 874: 'windows-874', 932: 'shift_jis', 936: 'gbk', 949: 'euc-kr', 950: 'big5', 65001: 'utf-8', 10000: 'macintosh' }[cp] || 'windows-1252';
  if (!DECODERS.has(name)) {
    try { DECODERS.set(name, new TextDecoder(name)); } catch { DECODERS.set(name, new TextDecoder('windows-1252')); }
  }
  return DECODERS.get(name);
}
const LATIN1 = new TextDecoder('latin1');

const CHARSET_CP = { 0: 1252, 1: 1252, 77: 10000, 128: 932, 129: 949, 134: 936, 136: 950, 161: 1253, 162: 1254, 177: 1255, 178: 1256, 186: 1257, 204: 1251, 222: 874, 238: 1250 };

// destinos que guardam texto próprio (nomes de fonte, estilos, instrução de campo)
const TEXT_DEST = new Set(['fonttbl', 'stylesheet', 'fldinst', 'pntext']);
// destinos cujo texto vai para o documento
const BODY_DEST = new Set(['rtf', 'hf', 'footnote']);
// destinos marcados com \* que o leitor entende
const KNOWN_STAR = new Set(['shppict', 'listtable', 'listoverridetable', 'fldinst', 'footnote', 'pntext']);
// grupos ignorados por inteiro
const SKIP_WORDS = new Set(['info', 'rsidtbl', 'generator', 'xmlnstbl', 'mmathPr', 'themedata', 'colorschememapping', 'datastore', 'latentstyles', 'pgdsctbl', 'filetbl', 'revtbl', 'userprops', 'docvar', 'template', 'nonshppict', 'object', 'bkmkstart', 'bkmkend', 'xe', 'tc', 'annotation', 'atnid', 'atnauthor', 'atndate', 'ftnsep', 'ftnsepc', 'aftnsep', 'aftnsepc', 'pnseclvl', 'listpicture', 'listtext', 'pn', 'leveltext', 'levelnumbers', 'listname', 'falt', 'panose', 'blipuid', 'wgrffmtfilter', 'defchp', 'defpap', 'protusertbl', 'passwordhash', 'nesttableprops']);

const HEADER_DEST = { header: ['header', 'default'], headerl: ['header', 'default'], headerr: ['header', 'default'], headerf: ['header', 'first'], footer: ['footer', 'default'], footerl: ['footer', 'default'], footerr: ['footer', 'default'], footerf: ['footer', 'first'] };

const hex2 = (n) => n.toString(16).padStart(2, '0').toUpperCase();

// cor da tabela para o nome de realce do Word mais próximo
function nearestHighlight(hex) {
  let best = null;
  let bestD = Infinity;
  const r = parseInt(hex.slice(0, 2), 16), g = parseInt(hex.slice(2, 4), 16), b = parseInt(hex.slice(4, 6), 16);
  for (const [name, val] of Object.entries(HIGHLIGHT)) {
    const v = String(val).replace('#', '');
    if (v.length !== 6) continue;
    const d = (parseInt(v.slice(0, 2), 16) - r) ** 2 + (parseInt(v.slice(2, 4), 16) - g) ** 2 + (parseInt(v.slice(4, 6), 16) - b) ** 2;
    if (d < bestD) { bestD = d; best = name; }
  }
  return bestD < 2000 ? best : null;
}

const defaultChar = () => ({ f: null, fs: 24, b: false, i: false, ul: null, strike: false, cf: 0, hl: 0, cb: 0, va: null, caps: false, scaps: false, v: false, expnd: 0 });
const defaultPara = () => ({ jc: null, li: 0, ri: 0, fi: 0, sb: 0, sa: 0, sl: 0, slmult: 1, keepn: false, pagebb: false, s: 0, ol: null, intbl: false, ls: 0, ilvl: 0, tabs: [], shd: 0 });

class RtfReader {
  constructor(bytes, model) {
    this.src = bytes;
    this.model = model;
    this.fonts = new Map();
    this.colors = [];
    this.curColor = {};
    this.stylesheet = new Map();
    this.lists = new Map();
    this.overrides = new Map();
    this.lsNum = new Map();
    this.cp = 1252;
    this.deff = 0;
    this.blocks = [];
    this.para = [];
    this.text = '';
    this.textKey = null;
    this.textCh = null;
    this.table = null;
    this.row = null;
    this.cell = null;
    this.rowDef = { cells: [], pending: {} };
    this.section = {};
    this.mediaN = 0;
    this.noteN = 0;
    this.headers = [];
    this.link = null;
    this.destText = '';
    this.markCache = new Map();
    this.tabType = null;
    this.tabLeader = null;
  }

  st() { return this.stack[this.stack.length - 1]; }

  parse() {
    const s = this.src;
    const n = s.length;
    this.stack = [{ dest: 'rtf', ch: defaultChar(), pp: defaultPara(), uc: 1, skip: false }];
    let i = 0;
    let star = false;
    let skipAlt = 0;
    const isLetter = (x) => (x >= 0x61 && x <= 0x7a) || (x >= 0x41 && x <= 0x5a);
    while (i < n) {
      const c = s[i];
      if (c === 0x7b) {
        this.flushText();
        const top = this.st();
        this.stack.push({ dest: top.dest, ch: top.ch, pp: { ...top.pp, tabs: top.pp.tabs.slice() }, uc: top.uc, skip: top.skip, grp: null });
        star = false;
        skipAlt = 0;
        i++;
        continue;
      }
      if (c === 0x7d) {
        this.flushText();
        const done = this.stack.pop();
        this.endGroup(done);
        skipAlt = 0;
        if (!this.stack.length) break;
        i++;
        continue;
      }
      if (c === 0x5c) {
        const d = s[i + 1];
        if (d === undefined) break;
        if (isLetter(d)) {
          let j = i + 1;
          while (j < n && isLetter(s[j])) j++;
          const word = LATIN1.decode(s.subarray(i + 1, j));
          let param = null;
          let k = j;
          if (s[k] === 0x2d && s[k + 1] >= 0x30 && s[k + 1] <= 0x39) {
            k++;
            let v = 0;
            while (k < n && s[k] >= 0x30 && s[k] <= 0x39) { v = v * 10 + (s[k] - 0x30); k++; }
            param = -v;
          } else if (s[k] >= 0x30 && s[k] <= 0x39) {
            let v = 0;
            while (k < n && s[k] >= 0x30 && s[k] <= 0x39) { v = v * 10 + (s[k] - 0x30); k++; }
            param = v;
          }
          if (s[k] === 0x20) k++;
          i = k;
          if (word === 'bin') { i += Math.max(0, param || 0); continue; }
          if (skipAlt > 0) { skipAlt--; continue; }
          const res = this.control(word, param, star);
          star = false;
          if (res === 'u') skipAlt = this.st().uc;
          continue;
        }
        if (d === 0x27) {
          const v = parseInt(LATIN1.decode(s.subarray(i + 2, i + 4)), 16);
          i += 4;
          if (skipAlt > 0) { skipAlt--; continue; }
          if (!Number.isNaN(v)) this.addBytes(Uint8Array.of(v));
          continue;
        }
        i += 2;
        if (d === 0x2a) { star = true; continue; }
        if (skipAlt > 0) { skipAlt--; continue; }
        if (d === 0x7b || d === 0x7d || d === 0x5c) this.addChar(String.fromCharCode(d));
        else if (d === 0x7e) this.addChar('\u00a0');
        else if (d === 0x5f) this.addChar('\u2011');
        else if (d === 0x0a || d === 0x0d) this.control('par', null, false);
        else if (d === 0x09) this.control('tab', null, false);
        continue;
      }
      if (c === 0x0d || c === 0x0a) { i++; continue; }
      let j = i;
      while (j < n && s[j] !== 0x7b && s[j] !== 0x7d && s[j] !== 0x5c && s[j] !== 0x0d && s[j] !== 0x0a) j++;
      let chunk = s.subarray(i, j);
      if (skipAlt > 0) {
        const k = Math.min(skipAlt, chunk.length);
        skipAlt -= k;
        chunk = chunk.subarray(k);
      }
      if (chunk.length) this.addBytes(chunk);
      i = j;
    }
    this.flushText();
    this.endPara(true);
    this.closeTable();
  }

  // bytes no código de página da fonte atual
  addBytes(bytes) {
    const st = this.st();
    if (st.skip) return;
    if (st.dest === 'pict') { this.pictHex += LATIN1.decode(bytes); return; }
    if (st.dest === 'colortbl') {
      for (const b of bytes) if (b === 0x3b) this.colorEntry();
      return;
    }
    if (st.skip) return;
    this.addChar(this.decode(bytes));
  }

  decode(bytes) {
    const st = this.st();
    const font = this.fonts.get(st.ch.f != null ? st.ch.f : this.deff);
    if (font && font.symbol) return Array.from(bytes, (b) => String.fromCharCode(0xF000 + b)).join('');
    return decoderFor(font && font.cp ? font.cp : this.cp).decode(bytes);
  }

  addChar(str) {
    const st = this.st();
    if (st.skip) return;
    if (TEXT_DEST.has(st.dest)) {
      this.destText += str;
      if (st.dest === 'fonttbl' || st.dest === 'stylesheet') this.maybeEntry();
      return;
    }
    if (!BODY_DEST.has(st.dest)) return;
    const key = st.ch;
    if (this.textCh !== key || this.textLink !== this.link) this.flushText();
    this.textCh = key;
    this.textLink = this.link;
    this.text += str;
  }

  // uma entrada da tabela de fontes ou de estilos termina no ponto e vírgula
  maybeEntry() {
    const k = this.destText.indexOf(';');
    if (k < 0) return;
    const name = this.destText.slice(0, k).trim();
    this.destText = this.destText.slice(k + 1);
    const g = this.entry;
    if (!g) return;
    if (g.kind === 'font') this.fonts.set(g.id, { name, cp: g.cp, symbol: /^(symbol|wingdings.*|webdings)$/i.test(name) });
    else if (g.kind === 'style') this.stylesheet.set(g.id, name);
    this.entry = null;
  }

  flushText() {
    if (!this.text) return;
    const t = this.text.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
    this.text = '';
    if (t) this.para.push(schema.text(t, this.marksFor(this.textCh, this.textLink)));
  }

  marksFor(ch, link) {
    const key = JSON.stringify(ch) + '|' + (link ? JSON.stringify(link.attrs) : '');
    const cached = this.markCache.get(key);
    if (cached) return cached;
    const out = [];
    if (link) {
      out.push(link);
      if (this.model.styles.get('Hyperlink')) out.push(marks.rStyle.create({ val: 'Hyperlink' }));
    }
    const font = this.fonts.get(ch.f != null ? ch.f : this.deff);
    const deffName = (this.fonts.get(this.deff) || {}).name;
    if (font && font.name && font.name !== deffName) out.push(marks.rFonts.create({ a: { ascii: font.name, hAnsi: font.name, cs: font.name, eastAsia: font.name } }));
    if (ch.b) out.push(marks.b.create({ val: true }));
    if (ch.i) out.push(marks.i.create({ val: true }));
    if (ch.caps) out.push(marks.caps.create({ val: true }));
    if (ch.scaps) out.push(marks.smallCaps.create({ val: true }));
    if (ch.strike) out.push(marks.strike.create({ val: true }));
    if (ch.v) out.push(marks.vanish.create({ val: true }));
    const col = this.colors[ch.cf];
    if (col && !link) out.push(marks.color.create({ a: { val: col } }));
    if (ch.expnd) out.push(marks.spacing.create({ val: ch.expnd }));
    if (ch.fs && ch.fs !== 24) out.push(marks.sz.create({ val: ch.fs }));
    const hl = this.colors[ch.hl];
    const cb = this.colors[ch.cb];
    if (hl) {
      const name = nearestHighlight(hl);
      out.push(name ? marks.highlight.create({ val: name }) : marks.shd.create({ a: { val: 'clear', color: 'auto', fill: hl } }));
    } else if (cb) out.push(marks.shd.create({ a: { val: 'clear', color: 'auto', fill: cb } }));
    if (ch.ul) out.push(marks.u.create({ a: { val: ch.ul } }));
    if (ch.va) out.push(marks.vertAlign.create({ val: ch.va }));
    const set = out.reduce((acc, m) => m.addToSet(acc), []);
    this.markCache.set(key, set);
    return set;
  }

  colorEntry() {
    const c = this.curColor;
    this.colors.push(c.r == null && c.g == null && c.b == null ? null : hex2(c.r || 0) + hex2(c.g || 0) + hex2(c.b || 0));
    this.curColor = {};
  }

  saveTarget() {
    return { blocks: this.blocks, para: this.para, table: this.table, row: this.row, cell: this.cell, rowDef: this.rowDef, text: this.text, textCh: this.textCh, textLink: this.textLink, link: this.link };
  }

  restore(saved) {
    Object.assign(this, saved);
  }

  // ------------------------------------------------------------ palavras de controle
  control(word, param, star) {
    const st = this.st();
    const on = param !== 0;
    if (st.skip) return null;
    if (star && !KNOWN_STAR.has(word) && !(word in HEADER_DEST) && st.dest !== 'listtable' && st.dest !== 'listoverridetable') { st.skip = true; st.dest = 'skip'; return null; }
    if (SKIP_WORDS.has(word)) { st.skip = true; st.dest = word; return null; }
    // tabelas de definição
    switch (word) {
      case 'rtf': case 'shppict': return null;
      case 'ansicpg': this.cp = param || 1252; return null;
      case 'deff': this.deff = param || 0; return null;
      case 'uc': st.uc = param == null ? 1 : param; return null;
      case 'u': {
        let v = param || 0;
        if (v < 0) v += 65536;
        this.addChar(String.fromCharCode(v));
        return 'u';
      }
      case 'fonttbl': st.dest = 'fonttbl'; this.destText = ''; return null;
      case 'colortbl': st.dest = 'colortbl'; this.curColor = {}; return null;
      case 'stylesheet': st.dest = 'stylesheet'; this.destText = ''; return null;
      case 'listtable': st.dest = 'listtable'; return null;
      case 'listoverridetable': st.dest = 'listoverridetable'; return null;
      default: break;
    }
    if (st.dest === 'fonttbl') {
      if (word === 'f') { this.entry = { kind: 'font', id: param, cp: null }; this.destText = ''; }
      else if (word === 'fcharset' && this.entry) this.entry.cp = CHARSET_CP[param] || null;
      else if (word === 'cpg' && this.entry) this.entry.cp = param;
      return null;
    }
    if (st.dest === 'stylesheet') {
      if (word === 's') { this.entry = { kind: 'style', id: param }; this.destText = ''; }
      else if (word === 'cs' || word === 'ts' || word === 'ds') { this.entry = { kind: 'other' }; this.destText = ''; }
      return null;
    }
    if (st.dest === 'colortbl') {
      if (word === 'red') this.curColor.r = param;
      else if (word === 'green') this.curColor.g = param;
      else if (word === 'blue') this.curColor.b = param;
      return null;
    }
    if (st.dest === 'listtable' || st.dest === 'list' || st.dest === 'listlevel') {
      if (word === 'list' && st.dest === 'listtable') { st.dest = 'list'; st.grp = { levels: [], id: null }; }
      else if (word === 'listlevel' && st.dest === 'list') { st.dest = 'listlevel'; st.grp = { nfc: 0, start: 1, owner: this.stack[this.stack.length - 2].grp }; }
      else if ((word === 'levelnfc' || word === 'levelnfcn') && st.dest === 'listlevel') st.grp.nfc = param || 0;
      else if (word === 'levelstartat' && st.dest === 'listlevel') st.grp.start = param == null ? 1 : param;
      else if (word === 'listid' && st.dest === 'list') st.grp.id = param;
      return null;
    }
    if (st.dest === 'listoverridetable' || st.dest === 'listoverride') {
      if (word === 'listoverride') { st.dest = 'listoverride'; st.grp = { listid: null, ls: null }; }
      else if (word === 'listid' && st.grp) st.grp.listid = param;
      else if (word === 'ls' && st.grp) st.grp.ls = param;
      return null;
    }
    if (st.dest === 'pict') {
      const p = this.pict;
      if (!p) return null;
      if (word === 'pngblip') p.type = 'png';
      else if (word === 'jpegblip') p.type = 'jpg';
      else if (word === 'emfblip') p.type = 'emf';
      else if (word === 'wmetafile') p.type = 'wmf';
      else if (word === 'dibitmap' || word === 'wbitmap' || word === 'macpict' || word === 'pmmetafile') p.type = 'raw';
      else if (word === 'picw') p.w = param;
      else if (word === 'pich') p.h = param;
      else if (word === 'picwgoal') p.gw = param;
      else if (word === 'pichgoal') p.gh = param;
      else if (word === 'picscalex') p.sx = param;
      else if (word === 'picscaley') p.sy = param;
      return null;
    }
    if (st.dest === 'pntext' || st.dest === 'fldinst') {
      if (word === 'tab') this.destText += '\t';
      return null;
    }
    // destinos com conteúdo
    if (word === 'pict') { this.flushText(); st.dest = 'pict'; this.pictHex = ''; this.pict = { type: 'png', w: 0, h: 0, gw: 0, gh: 0, sx: 100, sy: 100 }; return null; }
    if (word === 'pntext') { this.flushText(); st.dest = 'pntext'; this.destText = ''; return null; }
    if (word === 'field') { this.flushText(); st.dest = 'field'; st.grp = { instr: '' }; return null; }
    if (word === 'fldinst') { st.dest = 'fldinst'; this.destText = ''; return null; }
    if (word === 'fldrslt') {
      st.dest = 'rtf';
      // o grupo do campo é um dos de fora; o próprio fldrslt herdou o destino dele
      const f = this.stack.slice(0, -1).reverse().find((x) => x.dest === 'field' && x.grp);
      const instr = f ? f.grp.instr : '';
      const m = /HYPERLINK\s+(\\l\s+)?"([^"]*)"(?:\s+\\l\s+"([^"]*)")?/i.exec(instr);
      if (m) {
        this.flushText();
        const target = m[2];
        const anchor = m[1] ? target : m[3] || null;
        this.link = m[1] ? marks.hyperlink.create({ anchor: target.replace(/^#/, '') }) : anchor ? marks.hyperlink.create({ href: target + '#' + anchor }) : marks.hyperlink.create({ href: target });
        st.linkSet = true;
      } else if (/^\s*(PAGE|NUMPAGES|SECTIONPAGES)\b/i.test(instr)) {
        this.flushText();
        this.para.push(nodes.field_char.create({ type: 'begin' }), nodes.instr_text.create({ text: ' ' + instr.trim() + ' ' }), nodes.field_char.create({ type: 'separate' }));
        st.fieldEnd = true;
      }
      return null;
    }
    if (word === 'footnote') {
      this.flushText();
      st.grp = { saved: this.saveTarget() };
      st.dest = 'footnote';
      this.blocks = [];
      this.para = [];
      this.table = null; this.row = null; this.cell = null; this.link = null;
      st.pp = defaultPara();
      return null;
    }
    if (word in HEADER_DEST) {
      this.flushText();
      st.grp = { kind: HEADER_DEST[word][0], type: HEADER_DEST[word][1], saved: this.saveTarget() };
      st.dest = 'hf';
      this.blocks = [];
      this.para = [];
      this.table = null; this.row = null; this.cell = null; this.link = null;
      st.pp = defaultPara();
      return null;
    }
    if (word === 'chftn' || word === 'chpgn') {
      if (word === 'chpgn') { this.flushText(); this.para.push(nodes.field_char.create({ type: 'begin' }), nodes.instr_text.create({ text: ' PAGE ' }), nodes.field_char.create({ type: 'separate' }), schema.text('1'), nodes.field_char.create({ type: 'end' })); }
      return null;
    }
    // propriedades de caractere
    switch (word) {
      case 'plain': this.setChAll(defaultChar()); return null;
      case 'f': this.setCh('f', param); return null;
      case 'fs': this.setCh('fs', param || 24); return null;
      case 'b': this.setCh('b', on); return null;
      case 'i': this.setCh('i', on); return null;
      case 'ul': this.setCh('ul', on ? 'single' : null); return null;
      case 'ulnone': this.setCh('ul', null); return null;
      case 'uldb': this.setCh('ul', on ? 'double' : null); return null;
      case 'ulth': this.setCh('ul', on ? 'thick' : null); return null;
      case 'uld': this.setCh('ul', on ? 'dotted' : null); return null;
      case 'uldash': this.setCh('ul', on ? 'dash' : null); return null;
      case 'ulwave': this.setCh('ul', on ? 'wave' : null); return null;
      case 'ulw': this.setCh('ul', on ? 'words' : null); return null;
      case 'strike': case 'striked': this.setCh('strike', on); return null;
      case 'cf': this.setCh('cf', param || 0); return null;
      case 'highlight': this.setCh('hl', param || 0); return null;
      case 'cb': case 'chcbpat': this.setCh('cb', param || 0); return null;
      case 'super': this.setCh('va', on ? 'superscript' : null); return null;
      case 'sub': this.setCh('va', on ? 'subscript' : null); return null;
      case 'nosupersub': this.setCh('va', null); return null;
      case 'caps': this.setCh('caps', on); return null;
      case 'scaps': this.setCh('scaps', on); return null;
      case 'v': this.setCh('v', on); return null;
      case 'expndtw': this.setCh('expnd', param || 0); return null;
      default: break;
    }
    const pp = st.pp;
    // propriedades de parágrafo
    switch (word) {
      case 'pard': st.pp = defaultPara(); return null;
      case 'ql': pp.jc = null; return null;
      case 'qc': pp.jc = 'center'; return null;
      case 'qr': pp.jc = 'right'; return null;
      case 'qj': pp.jc = 'both'; return null;
      case 'qd': pp.jc = 'distribute'; return null;
      case 'li': case 'lin': pp.li = param || 0; return null;
      case 'ri': case 'rin': pp.ri = param || 0; return null;
      case 'fi': pp.fi = param || 0; return null;
      case 'sb': pp.sb = param || 0; return null;
      case 'sa': pp.sa = param || 0; return null;
      case 'sl': pp.sl = param || 0; return null;
      case 'slmult': pp.slmult = param || 0; return null;
      case 'keepn': pp.keepn = on; return null;
      case 'pagebb': pp.pagebb = on; return null;
      case 's': pp.s = param || 0; return null;
      case 'outlinelevel': pp.ol = param; return null;
      case 'intbl': pp.intbl = true; return null;
      case 'itap': if (param > 0) pp.intbl = true; return null;
      case 'ls': pp.ls = param || 0; return null;
      case 'ilvl': pp.ilvl = param || 0; return null;
      case 'cbpat': pp.shd = param || 0; return null;
      case 'tqr': this.tabType = 'right'; return null;
      case 'tqc': this.tabType = 'center'; return null;
      case 'tqdec': this.tabType = 'decimal'; return null;
      case 'tldot': this.tabLeader = 'dot'; return null;
      case 'tlhyph': this.tabLeader = 'hyphen'; return null;
      case 'tlul': case 'tlth': this.tabLeader = 'underscore'; return null;
      case 'tlmdot': this.tabLeader = 'middleDot'; return null;
      case 'tx': case 'tb':
        pp.tabs.push({ val: this.tabType || 'left', pos: String(param || 0), leader: this.tabLeader || null });
        this.tabType = null; this.tabLeader = null;
        return null;
      default: break;
    }
    // quebras, tabelas e caracteres especiais
    switch (word) {
      case 'par': this.flushText(); this.endPara(false); return null;
      case 'line': this.flushText(); this.para.push(nodes.hard_break.create(null, null, this.marksFor(st.ch, this.link))); return null;
      case 'page': this.flushText(); this.para.push(nodes.hard_break.create({ type: 'page' })); return null;
      case 'tab': this.flushText(); this.para.push(nodes.tab.create(null, null, this.marksFor(st.ch, this.link))); return null;
      case 'emdash': this.addChar('—'); return null;
      case 'endash': this.addChar('–'); return null;
      case 'bullet': this.addChar('•'); return null;
      case 'lquote': this.addChar('‘'); return null;
      case 'rquote': this.addChar('’'); return null;
      case 'ldblquote': this.addChar('“'); return null;
      case 'rdblquote': this.addChar('”'); return null;
      case 'emspace': this.addChar('\u2003'); return null;
      case 'enspace': this.addChar('\u2002'); return null;
      case 'qmspace': this.addChar('\u2005'); return null;
      case 'zwj': this.addChar('\u200d'); return null;
      case 'zwnj': this.addChar('\u200c'); return null;
      case 'cell': this.flushText(); this.endPara(false, true); this.endCell(); return null;
      case 'nestcell': this.flushText(); this.endPara(false, true); return null;
      case 'row': this.flushText(); this.endRow(); return null;
      case 'trowd': this.rowDef = { cells: [], pending: {} }; return null;
      case 'cellx': this.rowDef.cells.push({ x: param || 0, ...this.rowDef.pending }); this.rowDef.pending = {}; return null;
      case 'clvmgf': this.rowDef.pending.vmerge = 'start'; return null;
      case 'clvmrg': this.rowDef.pending.vmerge = 'cont'; return null;
      case 'clcbpat': this.rowDef.pending.shd = param || 0; return null;
      case 'clbrdrt': case 'clbrdrb': case 'clbrdrl': case 'clbrdrr': this.rowDef.pending.border = true; return null;
      case 'brdrcf': if (this.rowDef.pending.border && this.colors[param]) this.rowDef.pending.borderColor = this.colors[param]; return null;
      case 'trhdr': this.rowDef.header = true; return null;
      case 'sect': this.flushText(); this.endPara(false); return null;
      case 'paperw': this.section.w = param; return null;
      case 'paperh': this.section.h = param; return null;
      case 'margl': this.section.left = param; return null;
      case 'margr': this.section.right = param; return null;
      case 'margt': this.section.top = param; return null;
      case 'margb': this.section.bottom = param; return null;
      case 'headery': this.section.header = param; return null;
      case 'footery': this.section.footer = param; return null;
      case 'landscape': this.section.landscape = true; return null;
      case 'titlepg': this.section.titlePg = true; return null;
      default: return null;
    }
  }

  setCh(key, val) {
    const st = this.st();
    if (st.ch[key] === val) return;
    this.flushText();
    st.ch = { ...st.ch, [key]: val };
  }

  setChAll(ch) {
    this.flushText();
    this.st().ch = ch;
  }

  endGroup(g) {
    if (g.dest === 'fonttbl' || g.dest === 'stylesheet') {
      // entrada sem ponto e vírgula no fim
      if (this.entry && this.destText.trim()) { this.destText += ';'; this.maybeEntry(); }
    } else if (g.dest === 'colortbl') {
      if (Object.keys(this.curColor).length) this.colorEntry();
    } else if (g.dest === 'fldinst') {
      const f = this.stack.slice().reverse().find((x) => x.dest === 'field');
      if (f && f.grp) f.grp.instr += this.destText;
      this.destText = '';
    } else if (g.dest === 'listlevel' && g.grp && g.grp.owner) {
      g.grp.owner.levels.push({ nfc: g.grp.nfc, start: g.grp.start });
    } else if (g.dest === 'list' && g.grp && g.grp.id != null) {
      this.lists.set(g.grp.id, g.grp.levels);
    } else if (g.dest === 'listoverride' && g.grp && g.grp.ls != null) {
      this.overrides.set(g.grp.ls, g.grp.listid);
    } else if (g.dest === 'pntext') {
      this.pntext = this.destText;
      this.destText = '';
    } else if (g.dest === 'pict') {
      this.addPicture();
    } else if (g.dest === 'hf' && g.grp) {
      this.flushText();
      this.endPara(true);
      this.closeTable();
      const blocks = this.blocks;
      this.restore(g.grp.saved);
      if (blocks.length) this.headers.push({ kind: g.grp.kind, type: g.grp.type, blocks });
    } else if (g.dest === 'footnote' && g.grp) {
      this.flushText();
      this.endPara(true);
      this.closeTable();
      const blocks = this.blocks;
      this.restore(g.grp.saved);
      this.addFootnote(blocks);
    }
    if (g.linkSet) { this.flushText(); this.link = null; }
    if (g.fieldEnd) { this.flushText(); this.para.push(nodes.field_char.create({ type: 'end' })); }
  }

  addFootnote(blocks) {
    this.noteN++;
    const id = String(this.noteN);
    if (!blocks.length) blocks.push(nodes.paragraph.create());
    const refStyle = this.model.styles.get('FootnoteReference') ? [marks.rStyle.create({ val: 'FootnoteReference' })] : [];
    const mark = nodes.opaque_inline.create({ xml: '<w:footnoteRef xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>', level: 'r', kind: 'noteMark' }, null, refStyle);
    const fp = blocks[0];
    const textStyle = this.model.styles.get('FootnoteText') ? 'FootnoteText' : null;
    const content = [];
    let lead = true;
    fp.forEach((c) => {
      if (lead && c.isText && !c.text.trim()) return;
      lead = false;
      content.push(c);
    });
    blocks[0] = fp.type.create({ ...fp.attrs, styleId: fp.attrs.styleId || textStyle }, [mark, schema.text(' '), ...content]);
    this.model.footnotes.set(id, { id, doc: schema.node('doc', null, blocks), isNew: true });
    this.para.push(nodes.footnote_ref.create({ id, kind: 'footnote' }, null, refStyle));
  }

  addPicture() {
    const p = this.pict;
    this.pict = null;
    const hex = (this.pictHex || '').replace(/[^0-9a-fA-F]/g, '');
    this.pictHex = '';
    if (!p || hex.length < 16 || p.type === 'raw') return;
    const bytes = new Uint8Array(hex.length >> 1);
    for (let k = 0; k < bytes.length; k++) bytes[k] = parseInt(hex.substr(k * 2, 2), 16);
    const ext = p.type;
    const type = MIME[ext] || 'image/png';
    this.mediaN++;
    const path = 'word/media/rtf' + this.mediaN + '.' + ext;
    this.model.media.set(path, { url: URL.createObjectURL(new Blob([bytes], { type })), bytes, ext, type, pending: ext === 'emf' || ext === 'wmf' });
    // tamanho em twips: o "goal" quando existe, senão a medida da própria imagem
    const wtw = ((p.gw || (p.w ? p.w * 15 : 4320)) * (p.sx || 100)) / 100;
    const htw = ((p.gh || (p.h ? p.h * 15 : 2880)) * (p.sy || 100)) / 100;
    const cx = Math.round(wtw * 635);
    const cy = Math.round(htw * 635);
    this.flushText();
    this.para.push(nodes.drawing.create({ kind: 'image', mediaPath: path, cx, cy, ocx: cx, ocy: cy, ext }));
  }

  // ------------------------------------------------------------ parágrafos e tabelas
  paraAttrs(pp) {
    const a = {};
    if (pp.jc) a.jc = pp.jc;
    const ind = {};
    if (pp.li) ind.left = String(pp.li);
    if (pp.ri) ind.right = String(pp.ri);
    if (pp.fi > 0) ind.firstLine = String(pp.fi);
    else if (pp.fi < 0) ind.hanging = String(-pp.fi);
    if (Object.keys(ind).length) a.ind = ind;
    const sp = {};
    if (pp.sb) sp.before = String(pp.sb);
    if (pp.sa) sp.after = String(pp.sa);
    if (pp.sl) {
      if (pp.sl < 0) { sp.line = String(-pp.sl); sp.lineRule = 'exact'; }
      else if (pp.slmult) { sp.line = String(pp.sl); sp.lineRule = 'auto'; }
      else { sp.line = String(pp.sl); sp.lineRule = 'atLeast'; }
    }
    if (Object.keys(sp).length) a.spacing = sp;
    if (pp.keepn) a.keepNext = true;
    if (pp.pagebb) a.pageBreakBefore = true;
    if (pp.tabs.length) a.tabs = pp.tabs;
    const shd = this.colors[pp.shd];
    if (shd) a.shd = { val: 'clear', color: 'auto', fill: shd };
    // títulos viram nível de tópico, e a formatação continua a do próprio RTF,
    // que já vem explícita em cada trecho
    const sname = (this.stylesheet.get(pp.s) || '').toLowerCase();
    const hm = /^(heading|título|titulo)\s*(\d)$/.exec(sname);
    const level = hm ? Number(hm[2]) : pp.ol != null && pp.ol < 9 ? pp.ol + 1 : 0;
    if (level >= 1 && level <= 9) a.outlineLvl = level - 1;
    if (pp.ls) {
      const numId = this.numIdFor(pp.ls);
      if (numId) { a.numId = numId; a.ilvl = Math.max(0, Math.min(8, pp.ilvl || 0)); delete a.ind; }
    }
    return a;
  }

  numIdFor(ls) {
    if (this.lsNum.has(ls)) return this.lsNum.get(ls);
    const listid = this.overrides.get(ls);
    const levels = listid != null ? this.lists.get(listid) : null;
    const lvl0 = levels && levels[0];
    const bullet = !lvl0 || lvl0.nfc === 23 || lvl0.nfc === 255;
    let id = this.model.numbering.createList(bullet ? 'bullet' : 'number', 0);
    if (lvl0 && !bullet && lvl0.start != null && lvl0.start !== 1) {
      id = this.model.numbering.restartList(id);
      const num = this.model.numbering.nums.get(id);
      if (num) { const ov = num.overrides.get(0); if (ov) ov.startOverride = lvl0.start; }
    }
    this.lsNum.set(ls, id);
    return id;
  }

  endPara(final, inCell) {
    const st = this.st() || { pp: defaultPara() };
    const pp = st.pp;
    let content = this.para;
    this.para = [];
    if (final && !content.length) return;
    // numeração antiga (\pntext) sem lista de verdade: o rótulo fica como texto
    if (this.pntext != null && !pp.ls) {
      const lbl = this.pntext.replace(/\t+$/, '');
      if (lbl) content = [schema.text(lbl), nodes.tab.create(), ...content];
    }
    this.pntext = null;
    const node = nodes.paragraph.create(this.paraAttrs(pp), content);
    if (pp.intbl || inCell) {
      if (!this.cell) this.cell = { paras: [] };
      this.cell.paras.push(node);
      return;
    }
    this.closeTable();
    this.blocks.push(node);
  }

  endCell() {
    if (!this.row) this.row = { cells: [] };
    this.row.cells.push(this.cell || { paras: [] });
    this.cell = null;
  }

  endRow() {
    if (this.cell && this.cell.paras.length) this.endCell();
    const row = this.row || { cells: [] };
    this.row = null;
    const def = this.rowDef || { cells: [] };
    if (!this.table) this.table = { rows: [] };
    this.table.rows.push({ cells: row.cells, def: def.cells.slice(), header: !!def.header });
  }

  closeTable() {
    const t = this.table;
    if (!t) return;
    this.table = null;
    if (!t.rows.length) return;
    const cols = Math.max(1, ...t.rows.map((r) => Math.max(r.cells.length, r.def.length)));
    // larguras pelas bordas direitas das células da primeira linha completa
    const ref = t.rows.find((r) => r.def.length === cols) || t.rows[0];
    const widths = [];
    let prev = 0;
    for (let c = 0; c < cols; c++) {
      const x = ref.def[c] ? ref.def[c].x : prev + 1440;
      widths.push(Math.max(200, x - prev));
      prev = x;
    }
    const hasBorders = t.rows.some((r) => r.def.some((d) => d.border));
    const borderColor = (t.rows.flatMap((r) => r.def).find((d) => d.borderColor) || {}).borderColor || 'auto';
    const open = [];
    const rows = t.rows.map((r) => {
      const cells = [];
      for (let c = 0; c < cols; c++) {
        const d = r.def[c] || {};
        const cell = r.cells[c] || { paras: [] };
        if (d.vmerge === 'cont' && open[c]) { open[c].rowspan++; continue; }
        const attrs = { colwidth: [widths[c] / 15], tcW: { w: String(widths[c]), type: 'dxa' } };
        const fill = this.colors[d.shd];
        if (fill) attrs.shd = { val: 'clear', color: 'auto', fill };
        const entry = { attrs, paras: cell.paras.length ? cell.paras : [nodes.paragraph.create()], rowspan: 1 };
        open[c] = d.vmerge === 'start' ? entry : null;
        cells.push(entry);
      }
      return { cells, header: r.header };
    }).filter((r) => r.cells.length);
    const pmRows = rows.map((r) => nodes.table_row.create(r.header ? { header: true } : null, r.cells.map((c) => nodes.table_cell.create({ ...c.attrs, rowspan: c.rowspan }, c.paras))));
    const border = { val: 'single', sz: '4', space: '0', color: borderColor };
    this.blocks.push(nodes.table.create({
      grid: widths, tblW: { w: '0', type: 'auto' },
      tblBorders: hasBorders ? { top: border, left: border, bottom: border, right: border, insideH: border, insideV: border } : null,
    }, pmRows.length ? pmRows : [nodes.table_row.create(null, [nodes.table_cell.create(null, [nodes.paragraph.create()])])]));
  }
}

export async function readRtf(bytes) {
  const { model } = await blankModel();
  model.format = 'rtf';
  if (LATIN1.decode(bytes.subarray(0, 5)) !== '{\\rtf') throw new Error('O arquivo não parece ser um RTF válido.');
  const rd = new RtfReader(bytes, model);
  const prev = R.model;
  R.model = model;
  try {
    rd.parse();
  } finally {
    R.model = prev;
  }
  // padrões do RTF: 12 pt na fonte \deff e parágrafos sem espaço extra
  const deffFont = rd.fonts.get(rd.deff);
  model.styles.setDefaultFont(deffFont && deffFont.name ? deffFont.name : 'Times New Roman', 24);
  model.styles.setDefaultParagraphSpacing(0, 0, 240);
  const s = rd.section;
  if (s.w && s.h) model.section.setPageSize(s.w, s.h, s.landscape || s.w > s.h ? 'landscape' : null);
  if (s.left != null || s.top != null || s.right != null || s.bottom != null) {
    const m = model.section.pgMar;
    model.section.setMargins({ top: s.top ?? m.top, bottom: s.bottom ?? m.bottom, left: s.left ?? m.left, right: s.right ?? m.right, header: s.header ?? m.header, footer: s.footer ?? m.footer });
  }
  if (s.titlePg) model.section.setTitlePg(true);
  for (const hf of rd.headers) {
    const entry = newHeaderFooter(model, hf.kind);
    entry.doc = schema.node('doc', null, hf.blocks);
    entry.dirty = true;
    model.section.setReference(hf.kind, hf.type, entry.rId);
  }
  const blocks = rd.blocks.length ? rd.blocks : [nodes.paragraph.create()];
  return { model, doc: finalizeDoc(schema.node('doc', null, blocks), model) };
}

// ---------------------------------------------------------------- escrita

function rtfEscape(s) {
  let out = '';
  for (const ch of s) {
    const c = ch.codePointAt(0);
    if (ch === '\\') out += '\\\\';
    else if (ch === '{') out += '\\{';
    else if (ch === '}') out += '\\}';
    else if (ch === '\t') out += '\\tab ';
    else if (ch === '\u00a0') out += '\\~';
    else if (c < 0x20) continue;
    else if (c < 0x80) out += ch;
    else if (c <= 0xffff) out += '\\u' + (c > 32767 ? c - 65536 : c) + '?';
    else {
      const hi = Math.floor((c - 0x10000) / 0x400) + 0xd800;
      const lo = ((c - 0x10000) % 0x400) + 0xdc00;
      out += '\\u' + (hi - 65536) + '?\\u' + (lo - 65536) + '?';
    }
  }
  return out;
}

const NFC = { decimal: 0, upperRoman: 1, lowerRoman: 2, upperLetter: 3, lowerLetter: 4, ordinal: 5, bullet: 23 };
const hlRgb = (name) => String(HIGHLIGHT[name] || '#FFFF00').replace('#', '');

class RtfWriter {
  constructor(model) {
    this.model = model;
    this.styles = model.styles;
    this.fonts = new Map();
    this.colors = new Map();
    this.lists = new Map();
    this.counter = new NumberingCounter(model.numbering);
    const def = this.styles.defaults.r;
    this.font((def.rFonts && fontName(def.rFonts.a)) || 'Calibri');
  }

  font(name) {
    if (!this.fonts.has(name)) this.fonts.set(name, this.fonts.size);
    return this.fonts.get(name);
  }

  color(hex) {
    const h = String(hex).replace('#', '').toUpperCase();
    if (!this.colors.has(h)) this.colors.set(h, this.colors.size + 1);
    return this.colors.get(h);
  }

  listIndex(numId) {
    if (!this.lists.has(numId)) this.lists.set(numId, this.lists.size + 1);
    return this.lists.get(numId);
  }

  runProps(r) {
    let s = '';
    const f = r.rFonts ? fontName(r.rFonts.a) : null;
    if (f) s += '\\f' + this.font(f);
    s += '\\fs' + (r.sz ? r.sz.val : 22);
    if (r.b && r.b.val) s += '\\b';
    if (r.i && r.i.val) s += '\\i';
    if (r.u && r.u.a && r.u.a.val && r.u.a.val !== 'none') s += { double: '\\uldb', thick: '\\ulth', dotted: '\\uld', dash: '\\uldash', wave: '\\ulwave', words: '\\ulw' }[r.u.a.val] || '\\ul';
    if ((r.strike && r.strike.val) || (r.dstrike && r.dstrike.val)) s += '\\strike';
    if (r.caps && r.caps.val) s += '\\caps';
    if (r.smallCaps && r.smallCaps.val) s += '\\scaps';
    if (r.vanish && r.vanish.val) s += '\\v';
    if (r.color) { const c = colorOf(r.color.a); if (c && c !== 'auto') s += '\\cf' + this.color(c); }
    if (r.highlight && r.highlight.val && r.highlight.val !== 'none') s += '\\highlight' + this.color(hlRgb(r.highlight.val));
    else if (r.shd) { const c = shdColor(r.shd.a); if (c) s += '\\chcbpat' + this.color(c); }
    if (r.vertAlign && r.vertAlign.val === 'superscript') s += '\\super';
    if (r.vertAlign && r.vertAlign.val === 'subscript') s += '\\sub';
    if (r.spacing && r.spacing.val) s += '\\expndtw' + r.spacing.val;
    return s;
  }

  paraProps(attrs, extra = {}) {
    let eff = this.styles.effectivePPr(attrs, this.model.numbering);
    // numa tabela, o estilo da tabela fica entre o padrão do documento e o estilo do parágrafo
    if (extra.table && extra.table.pPr) {
      const stp = this.styles.resolvePara(attrs.styleId && this.styles.get(attrs.styleId) ? attrs.styleId : this.styles.defaultPara).pPr;
      const tp = extra.table.pPr;
      for (const k of Object.keys(tp)) {
        if (k === 'spacing' || k === 'ind') eff = { ...eff, [k]: { ...(this.styles.defaults.pPr[k] || {}), ...(tp[k] || {}), ...(stp[k] || {}), ...(attrs[k] || {}) } };
        else if (stp[k] == null && attrs[k] == null) eff = { ...eff, [k]: tp[k] };
      }
    }
    let s = '\\pard\\plain';
    if (extra.intbl) s += '\\intbl';
    const level = this.styles.headingLevel(attrs.styleId || this.styles.defaultPara);
    if (level >= 1 && level <= 6) s += '\\s' + level + '\\outlinelevel' + (level - 1);
    else if (attrs.outlineLvl != null && attrs.outlineLvl < 9) s += '\\outlinelevel' + attrs.outlineLvl;
    const jc = eff.jc;
    s += jc === 'center' ? '\\qc' : jc === 'right' || jc === 'end' ? '\\qr' : jc === 'both' || jc === 'distribute' ? '\\qj' : '\\ql';
    const ind = eff.ind || {};
    if (Number(ind.left)) s += '\\li' + Math.round(Number(ind.left)) + '\\lin' + Math.round(Number(ind.left));
    if (Number(ind.right)) s += '\\ri' + Math.round(Number(ind.right)) + '\\rin' + Math.round(Number(ind.right));
    if (Number(ind.hanging)) s += '\\fi-' + Math.round(Number(ind.hanging));
    else if (Number(ind.firstLine)) s += '\\fi' + Math.round(Number(ind.firstLine));
    const sp = eff.spacing || {};
    if (Number(sp.before)) s += '\\sb' + Math.round(Number(sp.before));
    if (Number(sp.after)) s += '\\sa' + Math.round(Number(sp.after));
    if (sp.line) {
      const v = Math.round(Number(sp.line));
      if (!sp.lineRule || sp.lineRule === 'auto') s += '\\sl' + v + '\\slmult1';
      else if (sp.lineRule === 'exact') s += '\\sl-' + Math.abs(v) + '\\slmult0';
      else s += '\\sl' + v + '\\slmult0';
    }
    if (eff.keepNext) s += '\\keepn';
    if (eff.pageBreakBefore) s += '\\pagebb';
    if (eff.shd) { const c = shdColor(eff.shd); if (c) s += '\\cbpat' + this.color(c); }
    for (const t of eff.tabs || []) {
      if (!t.val || t.val === 'clear' || t.val === 'bar') continue;
      s += { right: '\\tqr', end: '\\tqr', center: '\\tqc', decimal: '\\tqdec' }[t.val] || '';
      s += { dot: '\\tldot', hyphen: '\\tlhyph', underscore: '\\tlul', middleDot: '\\tlmdot' }[t.leader] || '';
      s += '\\tx' + Math.round(Number(t.pos) || 0);
    }
    return s;
  }

  // propriedades de caractere do estilo da tabela, quando o parágrafo e o trecho não definem
  withTable(r, attrs, marksList, tableR) {
    if (!tableR) return r;
    const psr = this.styles.resolvePara(attrs.styleId && this.styles.get(attrs.styleId) ? attrs.styleId : this.styles.defaultPara).r;
    const rs = marksList.find((m) => m.type.name === 'rStyle');
    const csr = rs ? this.styles.resolveChar(rs.attrs.val).r : {};
    const direct = new Set(marksList.map((m) => m.type.name));
    const out = { ...r };
    for (const k of Object.keys(tableR)) if (psr[k] == null && csr[k] == null && !direct.has(k)) out[k] = tableR[k];
    return out;
  }

  async inline(para, tableR) {
    let out = '';
    let link = null;
    let field = null;
    const children = [];
    para.forEach((c) => children.push(c));
    for (const c of children) {
      const lm = c.marks.find((m) => m.type.name === 'hyperlink');
      if (link && !(lm && lm.eq(link))) { out += '}}}'; link = null; }
      if (lm && !link && !field) {
        link = lm;
        const href = lm.attrs.href ? lm.attrs.href.replace(/^wiki:/, '') : null;
        const instr = href ? 'HYPERLINK "' + href.replace(/"/g, '') + '"' : 'HYPERLINK \\l "' + String(lm.attrs.anchor || '').replace(/"/g, '') + '"';
        out += '{\\field{\\*\\fldinst{' + instr.replace(/[{}]/g, '') + '}}{\\fldrslt{';
      }
      if (c.type === nodes.field_char) {
        if (c.attrs.type === 'begin') field = { instr: '', stage: 'instr' };
        else if (c.attrs.type === 'separate' && field) { out += '{\\field{\\*\\fldinst{' + rtfEscape(field.instr) + '}}{\\fldrslt{'; field.stage = 'result'; }
        else if (c.attrs.type === 'end' && field) { if (field.stage === 'result') out += '}}}'; field = null; }
        continue;
      }
      if (c.type === nodes.instr_text) { if (field) field.instr += c.attrs.text; continue; }
      if (field && field.stage === 'instr') continue;
      const r = this.withTable(this.styles.effectiveRun(para.attrs, c.marks), para.attrs, c.marks, tableR);
      const props = this.runProps(r);
      if (c.isText) out += '{' + props + ' ' + rtfEscape(c.text) + '}';
      else if (c.type === nodes.tab) out += '{' + props + '\\tab}';
      else if (c.type === nodes.hard_break) out += c.attrs.type === 'page' ? '\\page ' : '{' + props + '\\line}';
      else if (c.type === nodes.sym) out += '{' + props + ' ' + rtfEscape(c.textContent) + '}';
      else if (c.type === nodes.drawing) out += await this.picture(c);
      else if (c.type === nodes.footnote_ref) out += await this.footnote(c, props);
      else if (c.type === nodes.opaque_inline && c.attrs.kind === 'math') out += '{' + props + ' ' + rtfEscape(c.attrs.text || '') + '}';
    }
    if (field && field.stage === 'result') out += '}}}';
    if (link) out += '}}}';
    return out;
  }

  async footnote(ref, props) {
    const map = ref.attrs.kind === 'endnote' ? this.model.endnotes : this.model.footnotes;
    const fn = map.get(String(ref.attrs.id));
    let body = '';
    if (fn) {
      const inner = [];
      fn.doc.forEach((b) => inner.push(b));
      let first = true;
      for (const b of inner) {
        if (b.type !== nodes.paragraph) continue;
        const copy = [];
        b.forEach((c) => { if (!(c.type === nodes.opaque_inline && c.attrs.kind === 'noteMark')) copy.push(c); });
        const p = b.type.create(b.attrs, copy);
        body += this.paraProps(p.attrs) + (first ? '{\\super\\chftn}' : '') + (await this.inline(p)) + '\\par';
        first = false;
      }
    }
    return '{' + props + '\\super\\chftn}{\\footnote' + (ref.attrs.kind === 'endnote' ? '\\ftnalt' : '') + body + '}';
  }

  async picture(node) {
    const a = node.attrs;
    const m = a.mediaPath ? this.model.media.get(a.mediaPath) : null;
    let bytes = m && m.bytes;
    let ext = (m && m.ext) || 'png';
    if (!bytes && a.src) {
      try { const res = await fetch(a.src); bytes = new Uint8Array(await res.arrayBuffer()); ext = (res.headers.get('content-type') || '').includes('jpeg') ? 'jpg' : 'png'; } catch { bytes = null; }
    }
    if (!bytes) return '';
    let blip = { png: '\\pngblip', jpg: '\\jpegblip', jpeg: '\\jpegblip', emf: '\\emfblip', wmf: '\\wmetafile8' }[ext];
    if (!blip) {
      // gif, webp, bmp e outros viram PNG
      try {
        const bmp = await createImageBitmap(new Blob([bytes]));
        const cv = document.createElement('canvas');
        cv.width = bmp.width; cv.height = bmp.height;
        cv.getContext('2d').drawImage(bmp, 0, 0);
        bytes = canvasToBytes(cv, 'image/png');
        blip = '\\pngblip';
      } catch { return ''; }
    }
    const wtw = Math.round((a.cx || 1905000) / 635);
    const htw = Math.round((a.cy || 1428750) / 635);
    const lines = [];
    for (let i = 0; i < bytes.length; i += 64) {
      let line = '';
      for (let k = i; k < Math.min(bytes.length, i + 64); k++) line += hex2(bytes[k]);
      lines.push(line.toLowerCase());
    }
    return '{\\*\\shppict{\\pict' + blip + '\\picw' + Math.round(wtw / 15) + '\\pich' + Math.round(htw / 15) + '\\picwgoal' + wtw + '\\pichgoal' + htw + '\r\n' + lines.join('\r\n') + '}}';
  }

  async paragraph(n, extra = {}) {
    let s = this.paraProps(n.attrs, extra);
    const np = effectiveNumPr(n.attrs);
    if (np && n.content.size) {
      const lbl = this.counter.next(np.numId, np.ilvl);
      s += '\\ls' + this.listIndex(np.numId) + '\\ilvl' + np.ilvl;
      const kind = this.model.numbering.kindOf(np.numId, np.ilvl);
      s += '{\\listtext\\pard\\plain ' + rtfEscape(kind === 'bullet' ? '•' : (lbl && lbl.text) || '') + '\\tab}';
    }
    return s + ' ' + (await this.inline(n, extra.table && extra.table.r)) + (extra.lastInCell ? '\\cell' : '\\par') + '\r\n';
  }

  tableLook(lk) {
    lk = lk || {};
    const on = (k) => lk[k] === '1' || lk[k] === 'true' || lk[k] === true;
    if (lk.firstRow != null) return { firstRow: on('firstRow'), lastRow: on('lastRow'), firstColumn: on('firstColumn'), lastColumn: on('lastColumn'), noHBand: on('noHBand'), noVBand: on('noVBand') };
    const v = parseInt(lk.val || '04A0', 16);
    return { firstRow: !!(v & 0x20), lastRow: !!(v & 0x40), firstColumn: !!(v & 0x80), lastColumn: !!(v & 0x100), noHBand: !!(v & 0x200), noVBand: !!(v & 0x400) };
  }

  // formatação que o estilo da tabela dá a uma célula: cabeçalho, faixas e colunas
  cellStyle(res, look, ri, ci, nRows, nCols) {
    if (!res) return { shd: null, r: null, pPr: null };
    const types = [];
    if (!look.noVBand) {
      const v = ci - (look.firstColumn ? 1 : 0);
      if (v >= 0 && !(look.lastColumn && ci === nCols - 1)) types.push(v % 2 === 0 ? 'band1Vert' : 'band2Vert');
    }
    if (!look.noHBand) {
      const hh = ri - (look.firstRow ? 1 : 0);
      if (hh >= 0 && !(look.lastRow && ri === nRows - 1)) types.push(hh % 2 === 0 ? 'band1Horz' : 'band2Horz');
    }
    if (look.firstColumn && ci === 0) types.push('firstCol');
    if (look.lastColumn && ci === nCols - 1) types.push('lastCol');
    if (look.firstRow && ri === 0) types.push('firstRow');
    if (look.lastRow && ri === nRows - 1) types.push('lastRow');
    if (ri === 0 && ci === 0 && look.firstRow && look.firstColumn) types.push('nwCell');
    if (ri === 0 && ci === nCols - 1 && look.firstRow && look.lastColumn) types.push('neCell');
    if (ri === nRows - 1 && ci === 0 && look.lastRow && look.firstColumn) types.push('swCell');
    if (ri === nRows - 1 && ci === nCols - 1 && look.lastRow && look.lastColumn) types.push('seCell');
    let shd = res.tcPr && res.tcPr.shd ? res.tcPr.shd : null;
    let r = { ...(res.r || {}) };
    let pPr = { ...(res.pPr || {}) };
    for (const t of types) {
      const c = res.cond && res.cond[t];
      if (!c) continue;
      if (c.tcPr && c.tcPr.shd) shd = c.tcPr.shd;
      if (c.r) r = mergeRun(r, c.r);
      if (c.pPr) pPr = mergePPr(pPr, c.pPr);
    }
    return { shd, r, pPr };
  }

  async table(t) {
    let out = '';
    const rows = [];
    t.forEach((r) => rows.push(r));
    const grid = t.attrs.grid && t.attrs.grid.length ? t.attrs.grid.map((w) => Number(w) || 1440) : null;
    const resolved = this.styles.resolveTable(t.attrs.styleId || this.styles.defaultTable);
    const tb = { ...((resolved && resolved.tblPr && resolved.tblPr.tblBorders) || {}), ...(t.attrs.tblBorders || {}) };
    const visible = (b) => b && b.val && b.val !== 'nil' && b.val !== 'none';
    const border = (side, b) => {
      if (!visible(b)) return '';
      const col = b.color && b.color !== 'auto' && /^[0-9a-f]{6}$/i.test(b.color) ? '\\brdrcf' + this.color(b.color) : '';
      return '\\clbrdr' + side + (b.val === 'double' ? '\\brdrdb' : b.val === 'dotted' ? '\\brdrdot' : b.val === 'dashed' ? '\\brdrdash' : '\\brdrs') + '\\brdrw' + Math.max(1, Math.round((Number(b.sz) || 4) * 2.5)) + col;
    };
    const occupied = [];
    const totalCols = grid ? grid.length : 0;
    const look = this.tableLook(t.attrs.tblLook);
    for (let ri = 0; ri < rows.length; ri++) {
      const row = rows[ri];
      const cells = [];
      row.forEach((c) => cells.push(c));
      const defs = [];
      let x = 0;
      let col = 0;
      let ci = 0;
      const last = ri === rows.length - 1;
      while (ci < cells.length || occupied[col] > 0) {
        if (occupied[col] > 0) {
          x += grid ? grid[col] || 1440 : 1440;
          defs.push({ def: '\\clvmrg' + border('t', tb.insideH) + border('l', col === 0 ? tb.left : tb.insideV) + border('b', last ? tb.bottom : tb.insideH) + border('r', col + 1 >= totalCols ? tb.right : tb.insideV) + '\\cellx' + Math.round(x), empty: true });
          occupied[col]--;
          col++;
          continue;
        }
        const c = cells[ci++];
        if (!c) break;
        const span = c.attrs.colspan || 1;
        let w = 0;
        for (let k = 0; k < span; k++) w += c.attrs.colwidth && c.attrs.colwidth[k] ? c.attrs.colwidth[k] * 15 : grid ? grid[col + k] || 1440 : 1440;
        x += w;
        let d = (c.attrs.rowspan || 1) > 1 ? '\\clvmgf' : '';
        d += border('t', ri === 0 ? tb.top : tb.insideH) + border('l', col === 0 ? tb.left : tb.insideV) + border('b', last ? tb.bottom : tb.insideH) + border('r', col + span >= totalCols ? tb.right : tb.insideV);
        const fmt = this.cellStyle(resolved, look, ri, col, rows.length, totalCols || cells.length);
        const fill = c.attrs.shd ? shdColor(c.attrs.shd) : fmt.shd ? shdColor(fmt.shd) : null;
        if (fill) d += '\\clcbpat' + this.color(fill);
        if (c.attrs.vAlign === 'center') d += '\\clvertalc';
        else if (c.attrs.vAlign === 'bottom') d += '\\clvertalb';
        defs.push({ def: d + '\\cellx' + Math.round(x), cell: c, fmt });
        if ((c.attrs.rowspan || 1) > 1) for (let k = 0; k < span; k++) occupied[col + k] = (c.attrs.rowspan || 1) - 1;
        col += span;
      }
      const rowDef = '\\trowd\\trgaph108\\trleft0' + (row.attrs.header ? '\\trhdr' : '') + defs.map((d) => d.def).join('');
      out += rowDef + '\r\n';
      for (const d of defs) {
        if (d.empty) { out += '\\pard\\plain\\intbl \\cell\r\n'; continue; }
        const paras = [];
        d.cell.forEach((b) => paras.push(b));
        for (let k = 0; k < paras.length; k++) {
          const b = paras[k];
          const lastInCell = k === paras.length - 1;
          if (b.type === nodes.paragraph) out += await this.paragraph(b, { intbl: true, lastInCell, table: d.fmt });
          else out += '\\pard\\plain\\intbl ' + rtfEscape(b.textContent) + (lastInCell ? '\\cell' : '\\par') + '\r\n';
        }
      }
      out += rowDef + '\\row\r\n';
    }
    return out;
  }

  async blocks(parent) {
    let out = '';
    const list = [];
    parent.forEach((n) => list.push(n));
    for (const n of list) {
      if (n.type === nodes.paragraph) out += await this.paragraph(n);
      else if (n.type === nodes.table) out += (await this.table(n)) + '\\pard\\plain\r\n';
      else if (n.type === nodes.sdt_block) out += await this.blocks(n);
      else if (n.type === nodes.opaque_block && n.attrs.text && n.attrs.kind !== 'md') out += '\\pard\\plain ' + rtfEscape(n.attrs.text) + '\\par\r\n';
    }
    return out;
  }

  listTable() {
    if (!this.lists.size) return '';
    let lt = '{\\*\\listtable';
    let lo = '{\\*\\listoverridetable';
    for (const [numId, idx] of this.lists) {
      lt += '{\\list\\listtemplateid' + (1000 + idx) + '\\listhybrid';
      for (let lvl = 0; lvl < 9; lvl++) {
        const l = this.model.numbering.level(numId, lvl) || { numFmt: 'bullet', lvlText: '•', start: 1 };
        const nfc = NFC[l.numFmt] ?? 0;
        let text;
        let numbers = '';
        if (nfc === 23) {
          const ch = l.lvlText && l.lvlText.charCodeAt(0) < 0xF000 ? l.lvlText : ['•', 'o', '▪'][lvl % 3];
          text = "\\'01" + rtfEscape(ch);
        } else {
          const raw = l.lvlText || '%' + (lvl + 1) + '.';
          let body = '';
          let count = 0;
          for (let k = 0; k < raw.length; k++) {
            if (raw[k] === '%' && /\d/.test(raw[k + 1] || '')) {
              body += "\\'" + hex2(Number(raw[k + 1]) - 1).toLowerCase();
              count++;
              numbers += "\\'" + hex2(count).toLowerCase();
              k++;
            } else { body += rtfEscape(raw[k]); count++; }
          }
          text = "\\'" + hex2(count).toLowerCase() + body;
        }
        const ind = (l.pPr && l.pPr.ind) || {};
        const left = Number(ind.left) || 720 * (lvl + 1);
        const hang = Number(ind.hanging) || 360;
        lt += '{\\listlevel\\levelnfc' + nfc + '\\levelnfcn' + nfc + '\\leveljc0\\leveljcn0\\levelfollow0\\levelstartat' + (l.start ?? 1) + '\\levelspace0\\levelindent0{\\leveltext' + text + ';}{\\levelnumbers' + numbers + ';}\\fi-' + hang + '\\li' + left + '\\lin' + left + '}';
      }
      lt += '{\\listname ;}\\listid' + (2000 + idx) + '}';
      lo += '{\\listoverride\\listid' + (2000 + idx) + '\\listoverridecount0\\ls' + idx + '}';
    }
    return lt + '}\r\n' + lo + '}\r\n';
  }
}

export async function writeRtf(doc, model) {
  const prev = R.model;
  R.model = model;
  try {
    const w = new RtfWriter(model);
    const body = await w.blocks(doc);
    let hf = '';
    for (const kind of ['header', 'footer']) {
      const e = model.headerFor(kind, 1, false) || model.headerFor(kind, 0, true);
      if (e && e.doc) hf += '{\\' + kind + ' ' + (await w.blocks(e.doc)) + '}\r\n';
    }
    const s = model.section;
    let styles = '{\\stylesheet{\\s0\\ql Normal;}';
    for (let lvl = 1; lvl <= 6; lvl++) {
      const id = model.styles.headingId(lvl);
      if (!id) continue;
      const res = model.styles.resolvePara(id);
      styles += '{\\s' + lvl + '\\outlinelevel' + (lvl - 1) + w.runProps(res.r) + '\\sbasedon0\\snext0 heading ' + lvl + ';}';
    }
    styles += '}';
    const lists = w.listTable();
    const fonttbl = '{\\fonttbl' + [...w.fonts].map(([name, i]) => '{\\f' + i + '\\fnil\\fcharset0 ' + rtfEscape(name) + ';}').join('') + '}';
    const colortbl = '{\\colortbl ;' + [...w.colors.keys()].map((h) => '\\red' + parseInt(h.slice(0, 2), 16) + '\\green' + parseInt(h.slice(2, 4), 16) + '\\blue' + parseInt(h.slice(4, 6), 16) + ';').join('') + '}';
    const page = '\\paperw' + s.pgSz.w + '\\paperh' + s.pgSz.h + '\\margl' + s.pgMar.left + '\\margr' + s.pgMar.right + '\\margt' + Math.abs(s.pgMar.top) + '\\margb' + Math.abs(s.pgMar.bottom)
      + '\\headery' + s.pgMar.header + '\\footery' + s.pgMar.footer + (s.pgSz.w > s.pgSz.h ? '\\landscape' : '') + (s.titlePg ? '\\titlepg' : '');
    const title = model.properties && model.properties.title ? '{\\info{\\title ' + rtfEscape(model.properties.title) + '}}' : '';
    return '{\\rtf1\\ansi\\ansicpg1252\\deff0\\nouicompat\\deflang1046' + fonttbl + '\r\n' + colortbl + '\r\n' + styles + '\r\n' + lists + title + '\r\n'
      + '{\\*\\generator Redator;}\\viewkind4\\uc1' + page + '\r\n' + hf + body + '}\r\n';
  } finally {
    R.model = prev;
  }
}
