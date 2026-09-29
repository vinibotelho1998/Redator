// Documento do Word 97-2003 (.doc): leitura.
// O arquivo é um contêiner binário (CFB) com o texto dividido em peças e a
// formatação guardada em listas de "sprms". O leitor monta parágrafos, trechos,
// listas, tabelas, imagens, links, notas de rodapé, cabeçalho e rodapé.
import { schema, RUN_MARKS } from '../editor/schema.js';
import { R } from '../editor/render.js';
import { newHeaderFooter } from '../docx/writer.js';
import { MIME } from '../docx/reader.js';
import { blankModel, finalizeDoc } from './common.js';

const nodes = schema.nodes;
const marks = schema.marks;
const END = 0xFFFFFFFE;

// ---------------------------------------------------------------- contêiner CFB

class Cfb {
  constructor(bytes) {
    this.b = bytes;
    this.dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const sig = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
    if (!sig.every((v, i) => bytes[i] === v)) throw new Error('O arquivo não é um documento do Word 97-2003.');
    const u16 = (o) => this.dv.getUint16(o, true);
    const u32 = (o) => this.dv.getUint32(o, true);
    this.secSize = 1 << u16(30);
    this.miniSize = 1 << u16(32);
    const numFat = u32(44);
    const dirStart = u32(48);
    this.miniCutoff = u32(56);
    const miniFatStart = u32(60);
    const numMiniFat = u32(64);
    let difatNext = u32(68);
    const numDifat = u32(72);
    const difat = [];
    for (let i = 0; i < 109; i++) { const v = u32(76 + i * 4); if (v < 0xFFFFFFFA) difat.push(v); }
    for (let k = 0; k < numDifat && difatNext < 0xFFFFFFFA; k++) {
      const off = this.off(difatNext);
      const per = this.secSize / 4 - 1;
      for (let i = 0; i < per; i++) { const v = u32(off + i * 4); if (v < 0xFFFFFFFA) difat.push(v); }
      difatNext = u32(off + per * 4);
    }
    const fat = new Uint32Array(Math.max(1, numFat) * (this.secSize / 4));
    let fi = 0;
    for (const s of difat.slice(0, numFat)) {
      const off = this.off(s);
      for (let i = 0; i < this.secSize / 4 && off + i * 4 + 4 <= bytes.length; i++) fat[fi++] = u32(off + i * 4);
    }
    this.fat = fat;
    const dir = this.readChain(dirStart, Infinity);
    this.entries = [];
    for (let o = 0; o + 128 <= dir.length; o += 128) {
      const nameLen = dir[o + 64] | (dir[o + 65] << 8);
      let name = '';
      for (let k = 0; k + 1 < Math.max(0, nameLen - 2); k += 2) name += String.fromCharCode(dir[o + k] | (dir[o + k + 1] << 8));
      const dv = new DataView(dir.buffer, dir.byteOffset + o, 128);
      this.entries.push({ name, type: dir[o + 66], start: dv.getUint32(116, true), size: dv.getUint32(120, true) });
    }
    const root = this.entries[0];
    this.mini = root ? this.readChain(root.start, root.size) : new Uint8Array(0);
    this.miniFat = numMiniFat ? new DataView(this.readChain(miniFatStart, numMiniFat * this.secSize).buffer) : null;
  }

  off(sector) { return (sector + 1) * this.secSize; }

  readChain(start, size) {
    const parts = [];
    let total = 0;
    let s = start;
    let guard = 0;
    while (s < 0xFFFFFFFA && total < size && guard++ < 1e7) {
      const off = this.off(s);
      parts.push(this.b.subarray(off, Math.min(this.b.length, off + this.secSize)));
      total += this.secSize;
      s = this.fat[s];
    }
    const out = new Uint8Array(Math.min(total, size === Infinity ? total : size));
    let p = 0;
    for (const part of parts) {
      const take = Math.min(part.length, out.length - p);
      out.set(part.subarray(0, take), p);
      p += take;
      if (p >= out.length) break;
    }
    return out;
  }

  stream(name) {
    const e = this.entries.find((x) => x.type === 2 && x.name === name);
    if (!e) return null;
    if (e.size < this.miniCutoff && this.miniFat) {
      const out = new Uint8Array(e.size);
      let s = e.start;
      let p = 0;
      let guard = 0;
      while (s < 0xFFFFFFFA && p < e.size && guard++ < 1e7) {
        const off = s * this.miniSize;
        const take = Math.min(this.miniSize, e.size - p);
        out.set(this.mini.subarray(off, off + take), p);
        p += take;
        s = this.miniFat.getUint32(s * 4, true);
      }
      return out;
    }
    return this.readChain(e.start, e.size);
  }
}

// ---------------------------------------------------------------- utilidades binárias

const view = (b) => new DataView(b.buffer, b.byteOffset, b.byteLength);
const CP1252 = new TextDecoder('windows-1252');
const UTF16 = new TextDecoder('utf-16le');

// tamanho do operando de um sprm pelo campo spra
function operandSize(b, o, op) {
  switch (op >> 13) {
    case 0: case 1: return 1;
    case 2: case 4: case 5: return 2;
    case 3: return 4;
    case 7: return 3;
    default:
      if (op === 0xD608 || op === 0xD606) return ((b[o] | (b[o + 1] << 8)) + 1);
      if (op === 0xC615 && b[o] === 255) {
        const del = b[o + 1];
        const add = b[o + 2 + del * 4];
        return 1 + 1 + del * 4 + 1 + add * 3;
      }
      return b[o] + 1;
  }
}

function eachSprm(b, start, end, fn) {
  let o = start;
  while (o + 2 <= end) {
    const op = b[o] | (b[o + 1] << 8);
    o += 2;
    const size = operandSize(b, o, op);
    if (o + size > end + 1) break;
    fn(op, o, size);
    o += size;
  }
}

const ICO = [null, '000000', '0000FF', '00FFFF', '00FF00', 'FF00FF', 'FF0000', 'FFFF00', 'FFFFFF', '000080', '008080', '008000', '800080', '800000', '808000', '808080', 'C0C0C0'];
const HL = [null, 'black', 'blue', 'cyan', 'green', 'magenta', 'red', 'yellow', 'white', 'darkBlue', 'darkCyan', 'darkGreen', 'darkMagenta', 'darkRed', 'darkYellow', 'darkGray', 'lightGray'];
const KUL = { 1: 'single', 2: 'words', 3: 'double', 4: 'dotted', 6: 'thick', 7: 'dash', 9: 'dotDash', 10: 'dotDotDash', 11: 'wave', 20: 'dottedHeavy', 23: 'dashedHeavy', 39: 'dashLong', 43: 'wavyDouble' };
const NFC = { 0: 'decimal', 1: 'upperRoman', 2: 'lowerRoman', 3: 'upperLetter', 4: 'lowerLetter', 5: 'ordinal', 22: 'decimalZero', 23: 'bullet', 255: 'none' };
const JC = { 0: null, 1: 'center', 2: 'right', 3: 'both', 4: 'distribute', 5: 'center', 7: 'right' };

function hexColor(dv, o) {
  const r = dv.getUint8(o), g = dv.getUint8(o + 1), bb = dv.getUint8(o + 2), auto = dv.getUint8(o + 3);
  if (auto === 0xFF) return null;
  return [r, g, bb].map((x) => x.toString(16).padStart(2, '0')).join('').toUpperCase();
}

// ---------------------------------------------------------------- propriedades

const baseChp = () => ({ b: false, i: false, strike: false, dstrike: false, caps: false, smallCaps: false, vanish: false, ul: 0, color: null, hl: 0, hps: 20, ftc: 0, iss: 0, space: 0, spec: false, pic: null, symbol: null, obj: false });
const basePap = () => ({ jc: null, left: 0, right: 0, first: 0, before: 0, after: 0, line: 240, lineMult: true, keepNext: false, keepLines: false, pbb: false, outline: 9, ilvl: 0, ilfo: 0, inTable: false, ttp: false, itap: 0, tabs: [], shd: null, brcBottom: null, brcTop: null });

function applyChp(chp, b, start, end, styleChp, reader) {
  const dv = view(b);
  const tog = (key, v) => {
    if (v === 0x80) chp[key] = styleChp ? styleChp[key] : false;
    else if (v === 0x81) chp[key] = styleChp ? !styleChp[key] : true;
    else chp[key] = !!v;
  };
  eachSprm(b, start, end, (op, o) => {
    switch (op) {
      case 0x0835: tog('b', b[o]); break;
      case 0x0836: tog('i', b[o]); break;
      case 0x0837: tog('strike', b[o]); break;
      case 0x2A53: tog('dstrike', b[o]); break;
      case 0x083A: tog('smallCaps', b[o]); break;
      case 0x083B: tog('caps', b[o]); break;
      case 0x083C: tog('vanish', b[o]); break;
      case 0x2A3E: chp.ul = b[o]; break;
      case 0x2A42: chp.color = ICO[b[o]] || null; break;
      case 0x6870: chp.color = hexColor(dv, o); break;
      case 0x2A0C: chp.hl = b[o]; break;
      case 0x4A43: chp.hps = dv.getUint16(o, true); break;
      case 0x2A48: chp.iss = b[o]; break;
      case 0x4A4F: case 0x4A3D: chp.ftc = dv.getUint16(o, true); break;
      case 0x8840: chp.space = dv.getInt16(o, true); break;
      case 0x0855: chp.spec = !!b[o]; break;
      case 0x6A03: chp.pic = dv.getUint32(o, true); break;
      case 0x0806: chp.data = !!b[o]; break;
      case 0x080A: chp.obj = !!b[o]; break;
      case 0x6A09: chp.symbol = { ftc: dv.getUint16(o, true), ch: dv.getUint16(o + 2, true) }; break;
      case 0x4A30: {
        // estilo de caractere: aplica o estilo antes do resto
        const cs = reader && reader.styleChp(dv.getUint16(o, true), true);
        if (cs) Object.assign(chp, cs.delta || {});
        break;
      }
      default: break;
    }
  });
  return chp;
}

function applyPap(pap, b, start, end) {
  const dv = view(b);
  eachSprm(b, start, end, (op, o, size) => {
    switch (op) {
      case 0x2403: case 0x2461: pap.jc = JC[b[o]] ?? null; break;
      case 0x840F: case 0x845E: pap.left = dv.getInt16(o, true); break;
      case 0x840E: case 0x845D: pap.right = dv.getInt16(o, true); break;
      case 0x8411: case 0x8460: pap.first = dv.getInt16(o, true); break;
      case 0xA413: pap.before = dv.getUint16(o, true); break;
      case 0xA414: pap.after = dv.getUint16(o, true); break;
      case 0x6412: pap.line = dv.getInt16(o, true); pap.lineMult = !!dv.getInt16(o + 2, true); break;
      case 0x2406: pap.keepNext = !!b[o]; break;
      case 0x2405: pap.keepLines = !!b[o]; break;
      case 0x2407: pap.pbb = !!b[o]; break;
      case 0x2640: pap.outline = b[o]; break;
      case 0x260A: pap.ilvl = b[o]; break;
      case 0x460B: pap.ilfo = dv.getInt16(o, true); break;
      case 0x2416: pap.inTable = !!b[o]; break;
      case 0x2417: pap.ttp = !!b[o]; break;
      case 0x6649: pap.itap = dv.getInt32(o, true); if (pap.itap > 0) pap.inTable = true; break;
      case 0x244B: if (b[o]) pap.inTable = true; break;
      case 0x244C: if (b[o]) { pap.ttp = true; pap.inner = true; } break;
      case 0x442D: {
        const v = dv.getUint16(o, true);
        const back = ICO[(v >> 5) & 0x1F];
        pap.shd = back && back !== 'FFFFFF' ? back : null;
        break;
      }
      case 0xC64D: {
        const back = hexColor(dv, o + 5);
        pap.shd = back && back !== 'FFFFFF' ? back : null;
        break;
      }
      case 0x6426: pap.brcBottom = b[o + 1] ? { w: b[o], type: b[o + 1], color: ICO[b[o + 2]] } : null; break;
      case 0x6424: pap.brcTop = b[o + 1] ? { w: b[o], type: b[o + 1], color: ICO[b[o + 2]] } : null; break;
      case 0xC650: pap.brcBottom = b[o + 1 + 5] ? { w: b[o + 1 + 4], type: b[o + 1 + 5], color: hexColor(dv, o + 1) } : null; break;
      case 0xC64E: pap.brcTop = b[o + 1 + 5] ? { w: b[o + 1 + 4], type: b[o + 1 + 5], color: hexColor(dv, o + 1) } : null; break;
      case 0xC60D: case 0xC615: {
        // tabulações: remove e acrescenta paradas
        let p = o + 1;
        const del = b[p]; p++;
        const dels = [];
        for (let k = 0; k < del; k++) dels.push(dv.getInt16(p + k * 2, true));
        p += del * 2;
        if (op === 0xC615 && b[o] === 255) p += del * 2;
        const add = b[p]; p++;
        const adds = [];
        for (let k = 0; k < add; k++) adds.push(dv.getInt16(p + k * 2, true));
        p += add * 2;
        const tbd = [];
        for (let k = 0; k < add; k++) tbd.push(b[p + k]);
        let tabs = pap.tabs.filter((t) => !dels.includes(t.pos));
        adds.forEach((pos, k) => {
          const jc = tbd[k] & 7;
          const tlc = (tbd[k] >> 3) & 7;
          tabs = tabs.filter((t) => t.pos !== pos);
          if (jc === 4) return;
          tabs.push({ pos, val: ['left', 'center', 'right', 'decimal', 'bar', 'left', 'left', 'left'][jc], leader: [null, 'dot', 'hyphen', 'underscore', 'heavy', 'middleDot', null, null][tlc] });
        });
        pap.tabs = tabs.sort((a, c) => a.pos - c.pos);
        void size;
        break;
      }
      default: break;
    }
  });
  return pap;
}

// ---------------------------------------------------------------- leitor

class DocReader {
  constructor(bytes, model) {
    this.model = model;
    this.cfb = new Cfb(bytes);
    this.wd = this.cfb.stream('WordDocument');
    if (!this.wd) throw new Error('O arquivo não tem o fluxo WordDocument.');
    const wv = view(this.wd);
    if (wv.getUint16(0, true) !== 0xA5EC) throw new Error('O arquivo não é um documento do Word 97-2003.');
    const nFib = wv.getUint16(2, true);
    const flags = wv.getUint16(10, true);
    if (flags & 0x0100) throw new Error('O documento está protegido por senha.');
    if (nFib < 0xC1 && nFib !== 0xC0) throw new Error('Este .doc é do Word 6 ou 95, anterior ao Word 97, e ainda não é suportado.');
    this.table = this.cfb.stream(flags & 0x0200 ? '1Table' : '0Table');
    if (!this.table) throw new Error('O arquivo não tem a tabela de formatação.');
    this.data = this.cfb.stream('Data');
    let o = 32;
    const csw = wv.getUint16(o, true); o += 2 + csw * 2;
    const cslw = wv.getUint16(o, true); o += 2;
    this.lw = [];
    for (let i = 0; i < cslw; i++) this.lw.push(wv.getInt32(o + i * 4, true));
    o += cslw * 4;
    const cb = wv.getUint16(o, true); o += 2;
    this.fcLcb = [];
    for (let i = 0; i < cb; i++) this.fcLcb.push([wv.getUint32(o + i * 8, true), wv.getUint32(o + i * 8 + 4, true)]);
    this.ccpText = this.lw[3] || 0;
    this.ccpFtn = this.lw[4] || 0;
    this.ccpHdd = this.lw[5] || 0;
    this.ccpAtn = this.lw[7] || 0;
    this.ccpEdn = this.lw[8] || 0;
    this.mediaN = 0;
    this.noteN = 0;
    this.markCache = new Map();
    this.chpCache = new Map();
  }

  fc(i) { return this.fcLcb[i] || [0, 0]; }

  tableBytes(i) {
    const [fc, lcb] = this.fc(i);
    if (!lcb || fc + lcb > this.table.length) return null;
    return this.table.subarray(fc, fc + lcb);
  }

  // ------------------------------------------------------------ texto por peças
  parsePieces() {
    const clx = this.tableBytes(33);
    this.pieces = [];
    if (!clx) throw new Error('O arquivo não tem a tabela de peças de texto.');
    const dv = view(clx);
    let p = 0;
    this.prc = [];
    while (p < clx.length && clx[p] === 1) {
      const cbGrpprl = dv.getInt16(p + 1, true);
      this.prc.push(clx.subarray(p + 3, p + 3 + cbGrpprl));
      p += 3 + cbGrpprl;
    }
    if (clx[p] !== 2) throw new Error('Tabela de peças inválida.');
    const lcb = dv.getUint32(p + 1, true);
    const plc = clx.subarray(p + 5, p + 5 + lcb);
    const pv = view(plc);
    const n = Math.floor((lcb - 4) / 12);
    for (let i = 0; i < n; i++) {
      const cpStart = pv.getUint32(i * 4, true);
      const cpEnd = pv.getUint32((i + 1) * 4, true);
      const pcd = (n + 1) * 4 + i * 8;
      const raw = pv.getUint32(pcd + 2, true);
      const compressed = !!(raw & 0x40000000);
      const fc = compressed ? (raw & 0x3FFFFFFF) / 2 : raw & 0x3FFFFFFF;
      const prm = pv.getUint16(pcd + 6, true);
      this.pieces.push({ cpStart, cpEnd, fc, compressed, prm });
    }
    const total = this.ccpText + this.ccpFtn + this.ccpHdd + this.ccpAtn + this.ccpEdn + (this.lw[9] || 0) + (this.lw[10] || 0) + 1;
    this.fullText = this.textRange(0, Math.min(total, this.pieces.length ? this.pieces[this.pieces.length - 1].cpEnd : 0));
  }

  pieceAt(cp) {
    let lo = 0, hi = this.pieces.length - 1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      const p = this.pieces[mid];
      if (cp < p.cpStart) hi = mid - 1;
      else if (cp >= p.cpEnd) lo = mid + 1;
      else return p;
    }
    return null;
  }

  textRange(cpStart, cpEnd) {
    let out = '';
    for (const p of this.pieces) {
      const a = Math.max(cpStart, p.cpStart);
      const b = Math.min(cpEnd, p.cpEnd);
      if (a >= b) continue;
      if (p.compressed) {
        const s = p.fc + (a - p.cpStart);
        out += CP1252.decode(this.wd.subarray(s, s + (b - a)));
      } else {
        const s = p.fc + (a - p.cpStart) * 2;
        out += UTF16.decode(this.wd.subarray(s, s + (b - a) * 2));
      }
    }
    return out;
  }

  fcAt(cp) {
    const p = this.pieceAt(cp);
    if (!p) return -1;
    return p.fc + (cp - p.cpStart) * (p.compressed ? 1 : 2);
  }

  // ------------------------------------------------------------ fontes
  parseFonts() {
    this.fonts = [];
    const b = this.tableBytes(15);
    if (!b) return;
    const dv = view(b);
    const count = dv.getUint16(0, true);
    let p = 4;
    for (let i = 0; i < count && p < b.length; i++) {
      const len = b[p];
      const ffn = b.subarray(p + 1, p + 1 + len);
      let name = '';
      for (let k = 39; k + 1 < ffn.length; k += 2) {
        const c = ffn[k] | (ffn[k + 1] << 8);
        if (!c) break;
        name += String.fromCharCode(c);
      }
      this.fonts.push(name || 'Times New Roman');
      p += 1 + len;
    }
  }

  fontName(ftc) { return this.fonts[ftc] || null; }

  // ------------------------------------------------------------ estilos
  parseStyles() {
    this.styles = [];
    const b = this.tableBytes(1);
    if (!b) return;
    const dv = view(b);
    const cbStshi = dv.getUint16(0, true);
    const cstd = dv.getUint16(2, true);
    const cbBase = dv.getUint16(4, true);
    // STSHI: a fonte padrão fica em rgftcStandardChpStsh, 12 bytes depois do início
    this.defaultFtc = cbStshi >= 14 ? dv.getUint16(2 + 12, true) : 0;
    let p = 2 + cbStshi;
    for (let istd = 0; istd < cstd && p + 2 <= b.length; istd++) {
      const cbStd = dv.getUint16(p, true);
      p += 2;
      if (!cbStd) { this.styles.push(null); continue; }
      const std = b.subarray(p, p + cbStd);
      p += cbStd;
      try {
        const sv = view(std);
        const w0 = sv.getUint16(0, true);
        const w1 = sv.getUint16(2, true);
        const w2 = sv.getUint16(4, true);
        const sti = w0 & 0x0FFF;
        const stk = w1 & 0x000F;
        const base = w1 >> 4;
        const cupx = w2 & 0x000F;
        let q = cbBase;
        const cch = sv.getUint16(q, true);
        let name = '';
        for (let k = 0; k < cch; k++) name += String.fromCharCode(sv.getUint16(q + 2 + k * 2, true));
        q += 2 + cch * 2 + 2;
        const upx = [];
        for (let k = 0; k < cupx && q + 2 <= std.length; k++) {
          const cbUpx = sv.getUint16(q, true);
          upx.push(std.subarray(q + 2, q + 2 + cbUpx));
          q += 2 + cbUpx;
          if (cbUpx % 2) q++;
        }
        let papx = null;
        let chpx = null;
        if (stk === 1) { papx = upx[0] || null; chpx = upx[1] || null; }
        else if (stk === 2) chpx = upx[0] || null;
        this.styles.push({ istd, sti, stk, base: base === 0x0FFF ? null : base, name, papx, chpx });
      } catch { this.styles.push(null); }
    }
  }

  // formatação de caractere resolvida de um estilo, pela cadeia de estilos base
  styleChp(istd, deltaOnly) {
    const key = istd + (deltaOnly ? 'd' : '');
    if (this.chpCache.has(key)) return this.chpCache.get(key);
    const st = this.styles[istd];
    let chp;
    if (!st) chp = { ...baseChp(), ftc: this.defaultFtc || 0 };
    else {
      const base = st.base != null && st.base !== istd ? this.styleChp(st.base) : { ...baseChp(), ftc: this.defaultFtc || 0 };
      chp = { ...base };
      if (st.chpx) applyChp(chp, st.chpx, 0, st.chpx.length, base, null);
    }
    if (deltaOnly) {
      // para estilos de caractere, só o que o estilo muda
      const zero = { ...baseChp(), ftc: this.defaultFtc || 0 };
      const delta = {};
      if (st && st.chpx) applyChp(delta, st.chpx, 0, st.chpx.length, zero, null);
      const res = { delta };
      this.chpCache.set(key, res);
      return res;
    }
    this.chpCache.set(key, chp);
    return chp;
  }

  stylePap(istd) {
    const key = 'p' + istd;
    if (this.chpCache.has(key)) return this.chpCache.get(key);
    const st = this.styles[istd];
    let pap;
    if (!st) pap = basePap();
    else {
      const base = st.base != null && st.base !== istd ? this.stylePap(st.base) : basePap();
      pap = { ...base, tabs: base.tabs.slice() };
      // UpxPapx começa com o istd (2 bytes)
      if (st.papx && st.papx.length > 2) applyPap(pap, st.papx, 2, st.papx.length);
    }
    this.chpCache.set(key, pap);
    return pap;
  }

  // ------------------------------------------------------------ tabelas de formatação (FKP)
  parseBte(index, isPap) {
    const plc = this.tableBytes(index);
    const runs = [];
    if (!plc) return runs;
    const dv = view(plc);
    const n = Math.floor((plc.length - 4) / 8);
    for (let i = 0; i < n; i++) {
      const pn = dv.getUint32((n + 1) * 4 + i * 4, true) & 0x3FFFFF;
      const off = pn * 512;
      if (off + 512 > this.wd.length) continue;
      const page = this.wd.subarray(off, off + 512);
      const pv = view(page);
      const crun = page[511];
      for (let k = 0; k < crun; k++) {
        const fcStart = pv.getUint32(k * 4, true);
        const fcEnd = pv.getUint32((k + 1) * 4, true);
        if (isPap) {
          const bOff = page[(crun + 1) * 4 + k * 13] * 2;
          if (!bOff) { runs.push({ fcStart, fcEnd, istd: 0, g: null }); continue; }
          let cb = page[bOff];
          let start;
          if (cb === 0) { cb = page[bOff + 1] * 2; start = bOff + 2; } else { cb = cb * 2 - 1; start = bOff + 1; }
          const grp = page.subarray(start, start + cb);
          runs.push({ fcStart, fcEnd, istd: grp.length >= 2 ? grp[0] | (grp[1] << 8) : 0, g: grp.subarray(2) });
        } else {
          const bOff = page[(crun + 1) * 4 + k] * 2;
          if (!bOff) { runs.push({ fcStart, fcEnd, g: null }); continue; }
          const cb = page[bOff];
          runs.push({ fcStart, fcEnd, g: page.subarray(bOff + 1, bOff + 1 + cb) });
        }
      }
    }
    runs.sort((a, b) => a.fcStart - b.fcStart);
    return runs;
  }

  runAt(runs, fc) {
    let lo = 0, hi = runs.length - 1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      const r = runs[mid];
      if (fc < r.fcStart) hi = mid - 1;
      else if (fc >= r.fcEnd) lo = mid + 1;
      else return r;
    }
    return null;
  }

  // ------------------------------------------------------------ listas
  parseLists() {
    this.lstf = new Map();
    this.lfo = [];
    this.lfoNum = new Map();
    const [fcLst, lcbLst] = this.fc(73);
    // os níveis (LVL) vêm logo depois da PlfLst, fora do tamanho declarado por ela
    const lst = lcbLst ? this.table.subarray(fcLst) : null;
    if (lst) {
      try {
        const dv = view(lst);
        const cLst = dv.getInt16(0, true);
        let p = 2;
        const defs = [];
        for (let i = 0; i < cLst; i++) {
          const lsid = dv.getInt32(p, true);
          const simple = !!(lst[p + 26] & 1);
          defs.push({ lsid, simple });
          p += 28;
        }
        for (const d of defs) {
          const count = d.simple ? 1 : 9;
          const levels = [];
          for (let k = 0; k < count; k++) {
            const start = dv.getInt32(p, true);
            const nfc = lst[p + 4];
            const jcByte = lst[p + 5];
            const rgbxch = Array.from(lst.subarray(p + 6, p + 15));
            const follow = lst[p + 15];
            const cbChpx = lst[p + 24];
            const cbPapx = lst[p + 25];
            p += 28;
            const papx = lst.subarray(p, p + cbPapx);
            p += cbPapx;
            const chpx = lst.subarray(p, p + cbChpx);
            p += cbChpx;
            const cch = dv.getUint16(p, true);
            let xst = '';
            for (let c = 0; c < cch; c++) xst += String.fromCharCode(dv.getUint16(p + 2 + c * 2, true));
            p += 2 + cch * 2;
            const pap = applyPap(basePap(), papx, 0, papx.length);
            const chp = applyChp(baseChp(), chpx, 0, chpx.length, null, null);
            levels.push({ start, nfc, jc: jcByte & 3, rgbxch, follow, xst, pap, chp });
          }
          this.lstf.set(d.lsid, levels);
        }
      } catch { /* listas ilegíveis ficam sem numeração */ }
    }
    const lfo = this.tableBytes(74);
    if (lfo) {
      try {
        const dv = view(lfo);
        const n = dv.getInt32(0, true);
        for (let i = 0; i < n; i++) this.lfo.push({ lsid: dv.getInt32(4 + i * 16, true) });
      } catch { /* sem sobreposições */ }
    }
  }

  numIdFor(ilfo) {
    if (this.lfoNum.has(ilfo)) return this.lfoNum.get(ilfo);
    const lfo = this.lfo[ilfo - 1];
    const levels = lfo ? this.lstf.get(lfo.lsid) : null;
    let numId = null;
    if (levels && levels.length) {
      const defs = levels.map((l, i) => {
        const numFmt = NFC[l.nfc] || 'decimal';
        let lvlText;
        let font = null;
        if (numFmt === 'bullet') {
          lvlText = l.xst || '•';
          font = this.fontName(l.chp.ftc);
          const c = lvlText.charCodeAt(0);
          // Symbol e Wingdings guardam o caractere na área privada
          if (font && /symbol|wingdings/i.test(font) && c < 0x100) lvlText = String.fromCharCode(0xF000 + c);
        } else {
          lvlText = '';
          for (const ch of l.xst) {
            const c = ch.charCodeAt(0);
            lvlText += c < 9 ? '%' + (c + 1) : ch;
          }
        }
        return {
          numFmt, lvlText, start: l.start, font: numFmt === 'bullet' ? font : null,
          left: l.pap.left || 720 * (i + 1), hanging: l.pap.first < 0 ? -l.pap.first : 360,
          jc: ['left', 'center', 'right'][l.jc] || 'left', suff: l.follow === 1 ? 'space' : l.follow === 2 ? 'nothing' : null,
        };
      });
      numId = this.model.numbering.createFromLevels(defs);
    }
    this.lfoNum.set(ilfo, numId);
    return numId;
  }

  // ------------------------------------------------------------ seções, notas e cabeçalhos
  parseSections() {
    this.section = {};
    const b = this.tableBytes(6);
    if (!b) return;
    const dv = view(b);
    const n = Math.floor((b.length - 4) / 16);
    this.sectionEnds = [];
    for (let i = 0; i < n; i++) this.sectionEnds.push(dv.getUint32((i + 1) * 4, true));
    if (!n) return;
    const fcSepx = dv.getUint32((n + 1) * 4 + 2, true);
    if (fcSepx === 0xFFFFFFFF || fcSepx + 2 > this.wd.length) return;
    const wv = view(this.wd);
    const cb = wv.getUint16(fcSepx, true);
    const s = this.section;
    eachSprm(this.wd, fcSepx + 2, fcSepx + 2 + cb, (op, o) => {
      switch (op) {
        case 0xB01F: s.w = wv.getUint16(o, true); break;
        case 0xB020: s.h = wv.getUint16(o, true); break;
        case 0xB021: s.left = wv.getUint16(o, true); break;
        case 0xB022: s.right = wv.getUint16(o, true); break;
        case 0x9023: s.top = wv.getInt16(o, true); break;
        case 0x9024: s.bottom = wv.getInt16(o, true); break;
        case 0xB017: s.header = wv.getUint16(o, true); break;
        case 0xB018: s.footer = wv.getUint16(o, true); break;
        case 0x301D: s.landscape = this.wd[o] === 2; break;
        case 0x300A: s.titlePg = !!this.wd[o]; break;
        default: break;
      }
    });
  }

  plcCps(index) {
    const b = this.tableBytes(index);
    if (!b) return null;
    return b;
  }

  // ------------------------------------------------------------ montagem
  build() {
    this.parsePieces();
    this.parseFonts();
    this.parseStyles();
    this.chpx = this.parseBte(12, false);
    this.papx = this.parseBte(13, true);
    this.parseLists();
    this.parseSections();
    this.parseBookmarks();
    const normal = this.styleChp(0);
    this.normalChp = normal;
    this.normalPap = this.stylePap(0);
    this.normalFont = this.fontName(normal.ftc) || 'Times New Roman';
    // notas de rodapé: onde ficam as referências e o texto de cada uma
    this.footRefs = new Map();
    const refs = this.tableBytes(2);
    const txts = this.tableBytes(3);
    if (refs && txts && this.ccpFtn) {
      const rv = view(refs);
      const tv = view(txts);
      const n = Math.floor((refs.length - 4) / 6);
      for (let i = 0; i < n; i++) {
        const cp = rv.getUint32(i * 4, true);
        const a = tv.getUint32(i * 4, true);
        const b = tv.getUint32((i + 1) * 4, true);
        this.footRefs.set(cp, { a: this.ccpText + a, b: this.ccpText + b });
      }
    }
    const body = this.story(0, this.ccpText);
    // cabeçalhos e rodapés da primeira seção
    this.headers = [];
    const hdd = this.tableBytes(11);
    if (hdd && this.ccpHdd) {
      const hv = view(hdd);
      const n = Math.floor(hdd.length / 4) - 1;
      const base = this.ccpText + this.ccpFtn;
      const at = (k) => (k <= n ? hv.getUint32(k * 4, true) : null);
      const want = [[7, 'header', 'default'], [9, 'footer', 'default'], [10, 'header', 'first'], [11, 'footer', 'first']];
      for (const [k, kind, type] of want) {
        const a = at(k);
        const b = at(k + 1);
        if (a == null || b == null || b <= a) continue;
        const blocks = this.story(base + a, base + b);
        if (blocks.length && blocks.some((x) => x.textContent.trim() || x.type !== nodes.paragraph)) this.headers.push({ kind, type, blocks });
      }
    }
    return body;
  }

  // Monta os blocos de um trecho do texto: parágrafos e tabelas
  story(cpStart, cpEnd) {
    const text = this.fullText.slice(cpStart, cpEnd);
    const blocks = [];
    let table = null;
    let row = null;
    let cell = null;
    const closeTable = () => {
      if (!table) return;
      if (row && row.cells.length) table.rows.push(row);
      const node = this.tableNode(table);
      if (node) blocks.push(node);
      table = null;
      row = null;
      cell = null;
    };
    this.fields = [];
    let s = 0;
    while (s < text.length) {
      let e = s;
      while (e < text.length && text.charCodeAt(e) !== 13 && text.charCodeAt(e) !== 7) e++;
      const markCp = cpStart + Math.min(e, text.length - 1);
      const isCellMark = text.charCodeAt(e) === 7;
      const fc = this.fcAt(markCp);
      const run = fc >= 0 ? this.runAt(this.papx, fc) : null;
      const istd = run ? run.istd : 0;
      const pap = { ...this.stylePap(istd), tabs: this.stylePap(istd).tabs.slice() };
      if (run && run.g) applyPap(pap, run.g, 0, run.g.length);
      const piece = this.pieceAt(markCp);
      if (piece && piece.prm & 1 && this.prc[piece.prm >> 1]) { const g = this.prc[piece.prm >> 1]; applyPap(pap, g, 0, g.length); }
      if (isCellMark && pap.ttp) {
        // fim de linha da tabela: a definição das células vem neste parágrafo
        if (!table) table = { rows: [] };
        if (!row) row = { cells: [] };
        if (cell) { row.cells.push(cell); cell = null; }
        row.def = run && run.g ? this.tableDef(run.g) : null;
        table.rows.push(row);
        row = null;
        s = e + 1;
        continue;
      }
      const para = this.paragraph(cpStart + s, cpStart + e, istd, pap);
      if (pap.inTable || isCellMark) {
        if (!table) table = { rows: [] };
        if (!row) row = { cells: [] };
        if (!cell) cell = { paras: [] };
        cell.paras.push(para);
        if (isCellMark) { row.cells.push(cell); cell = null; }
      } else {
        closeTable();
        blocks.push(para);
      }
      s = e + 1;
    }
    closeTable();
    return blocks;
  }

  tableDef(g) {
    const dv = view(g);
    let def = null;
    const shd = [];
    let borders = null;
    eachSprm(g, 0, g.length, (op, o) => {
      if (op === 0xD608) {
        const itcMac = g[o + 2];
        const x = [];
        for (let k = 0; k <= itcMac; k++) x.push(dv.getInt16(o + 3 + k * 2, true));
        const tcs = [];
        const tcBase = o + 3 + (itcMac + 1) * 2;
        const cb = dv.getUint16(o, true);
        for (let k = 0; k < itcMac; k++) {
          const t = tcBase + k * 20;
          if (t + 20 > o + cb + 1) { tcs.push({ grf: 0 }); continue; }
          const grf = dv.getUint16(t, true);
          const brc = (q) => ({ w: g[q], type: g[q + 1], color: ICO[g[q + 2]] });
          tcs.push({ grf, top: brc(t + 4), left: brc(t + 8), bottom: brc(t + 12), right: brc(t + 16) });
        }
        def = { x, tcs };
      } else if (op === 0xD609) {
        const cb = g[o];
        for (let k = 0; k + 1 < cb; k += 2) { const v = dv.getUint16(o + 1 + k, true); const back = ICO[(v >> 5) & 0x1F]; shd.push(back && back !== 'FFFFFF' ? back : null); }
      } else if (op === 0xD612) {
        const cb = g[o];
        for (let k = 0; k + 9 < cb; k += 10) { const back = hexColor(dv, o + 1 + k + 4); shd.push(back && back !== 'FFFFFF' ? back : null); }
      } else if (op === 0xD605) {
        const brc = (q) => ({ w: g[q], type: g[q + 1], color: ICO[g[q + 2]] });
        borders = { top: brc(o + 1), left: brc(o + 5), bottom: brc(o + 9), right: brc(o + 13), insideH: brc(o + 17), insideV: brc(o + 21) };
      } else if (op === 0xD613) {
        const brc = (q) => ({ w: g[q + 4], type: g[q + 5], color: hexColor(dv, q) });
        borders = { top: brc(o + 1), left: brc(o + 9), bottom: brc(o + 17), right: brc(o + 25), insideH: brc(o + 33), insideV: brc(o + 41) };
      }
    });
    if (def) { def.shd = shd; def.borders = borders; }
    return def;
  }

  // A grade da tabela junta as bordas de célula de todas as linhas: no .doc cada
  // linha define as próprias células, e uma célula larga cobre várias colunas
  tableNode(t) {
    const rows = t.rows.filter((r) => r.cells.length);
    if (!rows.length) return null;
    const xs = new Set();
    for (const r of rows) if (r.def) for (const x of r.def.x) xs.add(x);
    let grid = [...xs].sort((a, b) => a - b);
    // bordas muito próximas viram uma só
    grid = grid.filter((x, i) => i === 0 || x - grid[i - 1] > 30);
    if (grid.length < 2) {
      const cols = Math.max(...rows.map((r) => r.cells.length));
      grid = Array.from({ length: cols + 1 }, (_, i) => i * 1440);
    }
    const widths = [];
    for (let k = 0; k + 1 < grid.length; k++) widths.push(grid[k + 1] - grid[k]);
    const colOf = (x) => {
      let best = 0;
      let d = Infinity;
      grid.forEach((g, i) => { const dd = Math.abs(g - x); if (dd < d) { d = dd; best = i; } });
      return best;
    };
    const brd = (b) => (b && b.type && b.type !== 0xFF ? { val: b.type === 3 ? 'double' : b.type === 6 ? 'dotted' : b.type === 7 ? 'dashed' : 'single', sz: String(Math.max(2, b.w || 4)), space: '0', color: b.color || 'auto' } : { val: 'nil' });
    const ref = rows.find((r) => r.def && r.def.borders);
    const tb = ref ? ref.def.borders : null;
    const vOpen = new Map();
    const pmRows = rows.map((r) => {
      const cells = [];
      const def = r.def;
      r.cells.forEach((c, k) => {
        const tc = def && def.tcs[k] ? def.tcs[k] : { grf: 0 };
        const x0 = def && def.x[k] != null ? def.x[k] : grid[0] + k * 1440;
        const x1 = def && def.x[k + 1] != null ? def.x[k + 1] : x0 + 1440;
        const c0 = colOf(x0);
        const c1 = Math.max(c0 + 1, colOf(x1));
        if (tc.grf & 0x0002) { if (cells.length) cells[cells.length - 1].c1 = Math.max(cells[cells.length - 1].c1, c1); return; }
        if ((tc.grf & 0x0020) && !(tc.grf & 0x0040) && vOpen.get(c0)) { vOpen.get(c0).rowspan++; return; }
        const attrs = { rowspan: 1 };
        const fill = def && def.shd ? def.shd[k] : null;
        if (fill) attrs.shd = { val: 'clear', color: 'auto', fill };
        const cb = {};
        for (const side of ['top', 'left', 'bottom', 'right']) if (tc[side] && tc[side].type) cb[side] = brd(tc[side]);
        if (Object.keys(cb).length) attrs.tcBorders = cb;
        const vAlign = (tc.grf >> 7) & 3;
        if (vAlign === 1) attrs.vAlign = 'center';
        else if (vAlign === 2) attrs.vAlign = 'bottom';
        if (tc.grf & 0x0040) vOpen.set(c0, attrs); else vOpen.delete(c0);
        cells.push({ attrs, c0, c1, paras: c.paras.length ? c.paras : [nodes.paragraph.create()] });
      });
      return nodes.table_row.create(null, cells.map((e) => {
        const cw = widths.slice(e.c0, e.c1);
        const w = cw.reduce((s, x) => s + x, 0);
        return nodes.table_cell.create({ ...e.attrs, colspan: e.c1 - e.c0, colwidth: cw.map((x) => x / 15), tcW: { w: String(Math.round(w)), type: 'dxa' } }, e.paras);
      }));
    });
    const tblBorders = tb ? { top: brd(tb.top), left: brd(tb.left), bottom: brd(tb.bottom), right: brd(tb.right), insideH: brd(tb.insideH), insideV: brd(tb.insideV) } : null;
    // a primeira borda costuma ficar 108 twips antes da margem, como o Word antigo desenhava
    return nodes.table.create({ grid: widths, tblW: { w: '0', type: 'auto' }, tblBorders, tblInd: grid[0] ? { w: String(grid[0]), type: 'dxa' } : null }, pmRows);
  }

  // ------------------------------------------------------------ parágrafos e trechos
  paragraph(cpA, cpB, istd, pap) {
    const content = [];
    let text = '';
    let textKey = null;
    let textMarks = null;
    const flush = () => { if (text) { content.push(schema.text(text, textMarks)); text = ''; } };
    let cp = cpA;
    while (cp < cpB) {
      const fc = this.fcAt(cp);
      const piece = this.pieceAt(cp);
      if (!piece || fc < 0) { cp++; continue; }
      const run = this.runAt(this.chpx, fc);
      const step = piece.compressed ? 1 : 2;
      let limit = Math.min(cpB, piece.cpEnd);
      if (run) limit = Math.min(limit, cp + Math.max(1, Math.ceil((run.fcEnd - fc) / step)));
      const chp = this.chpFor(istd, run, piece);
      const chunk = this.fullText.slice(cp, limit);
      for (let k = 0; k < chunk.length; k++) {
        let ch = chunk[k];
        const c = chunk.charCodeAt(k);
        const bms = this.bookmarks && this.bookmarks.get(cp + k);
        if (bms) { flush(); textKey = null; content.push(...bms); }
        const field = this.fields[this.fields.length - 1];
        if (c === 0x13) { flush(); textKey = null; this.fields.push({ instr: '', stage: 'instr', link: null }); continue; }
        if (c === 0x14 && field) { flush(); textKey = null; this.fieldSeparator(field, content); continue; }
        if (c === 0x15 && field) { flush(); textKey = null; this.fieldEnd(field, content); continue; }
        if (field && field.stage === 'instr') { if (c >= 0x20) field.instr += ch; continue; }
        if (c === 0x1E) ch = '‑';
        else if (c < 0x20) {
          if (c === 0x09) { flush(); content.push(nodes.tab.create(null, null, this.marksFor(chp))); }
          else if (c === 0x0B) { flush(); content.push(nodes.hard_break.create(null, null, this.marksFor(chp))); }
          else if (c === 0x0C) { flush(); content.push(nodes.hard_break.create({ type: 'page' })); }
          else if (c === 0x01 && chp.spec && chp.pic != null) { flush(); const d = this.picture(chp.pic, chp); if (d) content.push(d); }
          else if (c === 0x02 && chp.spec) { flush(); const fn = this.footnote(cp + k); if (fn) content.push(fn); }
          continue;
        }
        if (c === 0x28 && chp.spec && chp.symbol) {
          flush();
          const font = this.fontName(chp.symbol.ftc) || 'Symbol';
          content.push(nodes.sym.create({ font, char: chp.symbol.ch.toString(16).toUpperCase().padStart(4, '0') }, null, this.marksFor(chp)));
          continue;
        }
        if (textKey !== chp) {
          flush();
          textKey = chp;
          textMarks = this.marksFor(chp, this.currentLink());
        }
        text += ch;
      }
      flush();
      textKey = null;
      textMarks = null;
      cp = limit;
    }
    flush();
    const attrs = this.paraAttrs(pap, istd);
    // a formatação da marca de parágrafo vale para o número da lista, como no Word
    if (attrs.numId) {
      const mfc = this.fcAt(cpB);
      const mpiece = this.pieceAt(cpB);
      if (mfc >= 0 && mpiece) {
        const mchp = this.chpFor(istd, this.runAt(this.chpx, mfc), mpiece);
        const json = this.marksFor(mchp).filter((m) => RUN_MARKS.includes(m.type.name)).map((m) => ({ type: m.type.name, attrs: m.attrs }));
        if (json.length) attrs.markRPr = json;
      }
    }
    return nodes.paragraph.create(attrs, content);
  }

  chpFor(istd, run, piece) {
    const key = 'r' + istd + ':' + (run ? run.fcStart : -1) + ':' + ((piece && piece.prm) || 0);
    let chp = this.chpCache.get(key);
    if (chp) return chp;
    const sc = this.styleChp(istd);
    chp = { ...sc };
    if (run && run.g) applyChp(chp, run.g, 0, run.g.length, sc, this);
    if (piece && piece.prm & 1 && this.prc[piece.prm >> 1]) { const g = this.prc[piece.prm >> 1]; applyChp(chp, g, 0, g.length, sc, this); }
    this.chpCache.set(key, chp);
    return chp;
  }

  // marcadores do documento, usados pelos links do sumário
  parseBookmarks() {
    this.bookmarks = new Map();
    const names = this.tableBytes(21);
    const bkf = this.tableBytes(22);
    const bkl = this.tableBytes(23);
    if (!names || !bkf || !bkl) return;
    try {
      const nv = view(names);
      const extended = nv.getUint16(0, true) === 0xFFFF;
      let p = extended ? 2 : 0;
      const count = nv.getUint16(p, true);
      const cbExtra = nv.getUint16(p + 2, true);
      p += 4;
      const list = [];
      for (let i = 0; i < count && p < names.length; i++) {
        const cch = extended ? nv.getUint16(p, true) : names[p];
        p += extended ? 2 : 1;
        let s = '';
        for (let k = 0; k < cch; k++) s += String.fromCharCode(extended ? nv.getUint16(p + k * 2, true) : names[p + k]);
        p += cch * (extended ? 2 : 1) + cbExtra;
        list.push(s);
      }
      const fv = view(bkf);
      const lv = view(bkl);
      const n = Math.floor((bkf.length - 4) / 8);
      const nl = Math.floor((bkl.length - 4) / 4);
      const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
      const add = (cp, node) => { if (!this.bookmarks.has(cp)) this.bookmarks.set(cp, []); this.bookmarks.get(cp).push(node); };
      for (let i = 0; i < n && i < list.length; i++) {
        const name = list[i];
        if (!name) continue;
        const start = fv.getUint32(i * 4, true);
        const ibkl = fv.getInt16((n + 1) * 4 + i * 4, true);
        const end = ibkl >= 0 && ibkl < nl ? lv.getUint32(ibkl * 4, true) : start;
        const id = String(i + 1);
        add(start, nodes.opaque_inline.create({ xml: '<w:bookmarkStart ' + W + ' w:id="' + id + '" w:name="' + name.replace(/[<>&"]/g, '') + '"/>', level: 'p', kind: 'bookmarkStart', id }));
        add(Math.max(start, end), nodes.opaque_inline.create({ xml: '<w:bookmarkEnd ' + W + ' w:id="' + id + '"/>', level: 'p', kind: 'bookmarkEnd', id }));
      }
    } catch { this.bookmarks = new Map(); }
  }

  currentLink() {
    for (let i = this.fields.length - 1; i >= 0; i--) if (this.fields[i].link && this.fields[i].stage === 'result') return this.fields[i].link;
    return null;
  }

  fieldSeparator(field, content) {
    field.stage = 'result';
    const instr = field.instr.trim();
    const m = /^HYPERLINK\s+(?:(\\l)\s+)?"([^"]*)"(?:.*?\\l\s+"([^"]*)")?/i.exec(instr);
    if (m) {
      const target = m[2];
      field.link = m[1] ? marks.hyperlink.create({ anchor: target }) : marks.hyperlink.create({ href: target + (m[3] ? '#' + m[3] : '') });
    } else if (/^(PAGE|NUMPAGES|SECTIONPAGES)\b/i.test(instr)) {
      content.push(nodes.field_char.create({ type: 'begin' }), nodes.instr_text.create({ text: ' ' + instr + ' ' }), nodes.field_char.create({ type: 'separate' }));
      field.pageField = true;
    }
  }

  fieldEnd(field, content) {
    this.fields.pop();
    if (field.pageField) content.push(nodes.field_char.create({ type: 'end' }));
  }

  marksFor(chp, link) {
    const key = JSON.stringify(chp) + '|' + (link ? JSON.stringify(link.attrs) : '');
    const cached = this.markCache.get(key);
    if (cached) return cached;
    const n = this.normalChp;
    const out = [];
    if (link) {
      out.push(link);
      if (this.model.styles.get('Hyperlink')) out.push(marks.rStyle.create({ val: 'Hyperlink' }));
    }
    const font = this.fontName(chp.ftc);
    if (font && font !== this.normalFont) out.push(marks.rFonts.create({ a: { ascii: font, hAnsi: font, cs: font, eastAsia: font } }));
    if (chp.b !== n.b) out.push(marks.b.create({ val: chp.b }));
    if (chp.i !== n.i) out.push(marks.i.create({ val: chp.i }));
    if (chp.caps) out.push(marks.caps.create({ val: true }));
    if (chp.smallCaps) out.push(marks.smallCaps.create({ val: true }));
    if (chp.strike) out.push(marks.strike.create({ val: true }));
    if (chp.dstrike) out.push(marks.dstrike.create({ val: true }));
    if (chp.vanish) out.push(marks.vanish.create({ val: true }));
    if (chp.color && !link) out.push(marks.color.create({ a: { val: chp.color } }));
    if (chp.space) out.push(marks.spacing.create({ val: chp.space }));
    if (chp.hps !== n.hps) out.push(marks.sz.create({ val: chp.hps }));
    if (chp.hl && HL[chp.hl]) out.push(marks.highlight.create({ val: HL[chp.hl] }));
    if (chp.ul && KUL[chp.ul]) out.push(marks.u.create({ a: { val: KUL[chp.ul] } }));
    if (chp.iss === 1) out.push(marks.vertAlign.create({ val: 'superscript' }));
    else if (chp.iss === 2) out.push(marks.vertAlign.create({ val: 'subscript' }));
    const set = out.reduce((acc, m) => m.addToSet(acc), []);
    this.markCache.set(key, set);
    return set;
  }

  paraAttrs(pap, istd) {
    const n = this.normalPap;
    const a = {};
    if (pap.jc) a.jc = pap.jc;
    const ind = {};
    if (pap.left) ind.left = String(pap.left);
    if (pap.right) ind.right = String(pap.right);
    if (pap.first > 0) ind.firstLine = String(pap.first);
    else if (pap.first < 0) ind.hanging = String(-pap.first);
    const sp = {};
    if (pap.before !== n.before) sp.before = String(pap.before);
    if (pap.after !== n.after) sp.after = String(pap.after);
    if (pap.line !== n.line || pap.lineMult !== n.lineMult) {
      if (pap.lineMult) { sp.line = String(pap.line); sp.lineRule = 'auto'; }
      else if (pap.line < 0) { sp.line = String(-pap.line); sp.lineRule = 'exact'; }
      else { sp.line = String(pap.line); sp.lineRule = 'atLeast'; }
    }
    if (Object.keys(sp).length) a.spacing = sp;
    if (pap.keepNext) a.keepNext = true;
    if (pap.keepLines) a.keepLines = true;
    if (pap.pbb) a.pageBreakBefore = true;
    if (pap.tabs.length) a.tabs = pap.tabs.map((t) => ({ val: t.val, pos: String(t.pos), leader: t.leader }));
    if (pap.shd) a.shd = { val: 'clear', color: 'auto', fill: pap.shd };
    const bdr = (b) => ({ val: b.type === 3 ? 'double' : 'single', sz: String(Math.max(2, b.w || 4)), space: '1', color: b.color || 'auto' });
    if (pap.brcBottom || pap.brcTop) a.pBdr = { ...(pap.brcTop ? { top: bdr(pap.brcTop) } : {}), ...(pap.brcBottom ? { bottom: bdr(pap.brcBottom) } : {}) };
    // títulos: pelo estilo de título do Word (sti 1 a 9) ou pelo nível de tópico
    const st = this.styles[istd];
    let level = st && st.sti >= 1 && st.sti <= 9 ? st.sti : pap.outline < 9 ? pap.outline + 1 : 0;
    if (level >= 1 && level <= 9) a.outlineLvl = level - 1;
    if (Object.keys(ind).length) a.ind = ind;
    if (pap.ilfo > 0 && pap.ilfo < 2047) {
      const numId = this.numIdFor(pap.ilfo);
      if (numId) { a.numId = numId; a.ilvl = Math.max(0, Math.min(8, pap.ilvl || 0)); }
    }
    return a;
  }

  footnote(cp) {
    const ref = this.footRefs.get(cp);
    if (!ref) return null;
    const saved = this.fields;
    const blocks = this.story(ref.a, ref.b);
    this.fields = saved;
    this.noteN++;
    const id = String(this.noteN);
    if (!blocks.length) blocks.push(nodes.paragraph.create());
    const refStyle = this.model.styles.get('FootnoteReference') ? [marks.rStyle.create({ val: 'FootnoteReference' })] : [];
    const mark = nodes.opaque_inline.create({ xml: '<w:footnoteRef xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>', level: 'r', kind: 'noteMark' }, null, refStyle);
    const fp = blocks[0];
    const rest = [];
    let lead = true;
    fp.forEach((c) => { if (lead && c.isText && !c.text.trim()) return; lead = false; rest.push(c); });
    const textStyle = this.model.styles.get('FootnoteText') ? 'FootnoteText' : null;
    blocks[0] = fp.type.create({ ...fp.attrs, styleId: fp.attrs.styleId || textStyle }, [mark, schema.text(' '), ...rest]);
    this.model.footnotes.set(id, { id, doc: schema.node('doc', null, blocks), isNew: true });
    return nodes.footnote_ref.create({ id, kind: 'footnote' }, null, refStyle);
  }

  // ------------------------------------------------------------ imagens
  picture(offset, chp) {
    const d = this.data;
    if (!d || offset + 68 > d.length) return null;
    try {
      const dv = view(d);
      const lcb = dv.getUint32(offset, true);
      const cbHeader = dv.getUint16(offset + 4, true);
      const mm = dv.getInt16(offset + 6, true);
      const dxaGoal = dv.getInt16(offset + 28, true);
      const dyaGoal = dv.getInt16(offset + 30, true);
      const mx = dv.getUint16(offset + 32, true);
      const my = dv.getUint16(offset + 34, true);
      let p = offset + cbHeader;
      if (mm === 0x66) p += 1 + d[p];
      const end = Math.min(d.length, offset + lcb);
      const blip = this.findBlip(d, p, end);
      if (!blip) return null;
      const w = (dxaGoal * (mx || 1000)) / 1000;
      const h = (dyaGoal * (my || 1000)) / 1000;
      this.mediaN++;
      const path = 'word/media/doc' + this.mediaN + '.' + blip.ext;
      const type = MIME[blip.ext] || 'image/png';
      this.model.media.set(path, { url: URL.createObjectURL(new Blob([blip.bytes], { type })), bytes: blip.bytes, ext: blip.ext, type, pending: blip.pending || false, inflate: blip.inflate || null });
      const cx = Math.round(Math.max(1, w) * 635);
      const cy = Math.round(Math.max(1, h) * 635);
      return nodes.drawing.create({ kind: 'image', mediaPath: path, cx, cy, ocx: cx, ocy: cy, ext: blip.ext }, null, this.marksFor(chp));
    } catch {
      return null;
    }
  }

  // procura o registro de imagem dentro dos registros OfficeArt
  findBlip(d, start, end) {
    const dv = view(d);
    let p = start;
    let guard = 0;
    while (p + 8 <= end && guard++ < 10000) {
      const verInst = dv.getUint16(p, true);
      const type = dv.getUint16(p + 2, true);
      const len = dv.getUint32(p + 4, true);
      const ver = verInst & 0xF;
      const inst = verInst >> 4;
      const body = p + 8;
      if (ver === 0xF && type !== 0xF007) { p = body; continue; }
      if (type === 0xF007) {
        // FBSE: o BLIP vem depois do cabeçalho e do nome
        const cbName = d[body + 33];
        const inner = this.findBlip(d, body + 36 + cbName, Math.min(end, body + len));
        if (inner) return inner;
        p = body + len;
        continue;
      }
      if (type >= 0xF018 && type <= 0xF117) {
        const bitmaps = { 0xF01D: 'jpg', 0xF01E: 'png', 0xF01F: 'bmp', 0xF029: 'tif', 0xF02A: 'jpg' };
        const metas = { 0xF01A: 'emf', 0xF01B: 'wmf' };
        const twoUid = { 0xF01D: 0x46B, 0xF01E: 0x6E1, 0xF01F: 0x7A9, 0xF029: 0x6E5, 0xF02A: 0x6E3, 0xF01A: 0x3D5, 0xF01B: 0x217 };
        const uidBytes = inst === twoUid[type] ? 32 : 16;
        if (bitmaps[type]) {
          let bytes = d.subarray(body + uidBytes + 1, Math.min(end, body + len));
          let ext = bitmaps[type];
          if (ext === 'bmp') {
            // DIB sem cabeçalho de arquivo: acrescenta os 14 bytes do BMP
            const head = new Uint8Array(14);
            const hv = new DataView(head.buffer);
            head[0] = 0x42; head[1] = 0x4D;
            hv.setUint32(2, bytes.length + 14, true);
            const infoSize = bytes[0] | (bytes[1] << 8) | (bytes[2] << 16) | (bytes[3] << 24);
            const bpp = bytes[14] | (bytes[15] << 8);
            const clr = bytes[32] | (bytes[33] << 8) | (bytes[34] << 16) | (bytes[35] << 24);
            const pal = bpp <= 8 ? (clr || (1 << bpp)) * 4 : 0;
            hv.setUint32(10, 14 + infoSize + pal, true);
            const full = new Uint8Array(bytes.length + 14);
            full.set(head);
            full.set(bytes, 14);
            bytes = full;
          }
          return { bytes: bytes.slice(), ext, pending: ext === 'tif' };
        }
        if (metas[type]) {
          const h = body + uidBytes;
          const compressed = d[h + 32] === 0;
          const bytes = d.subarray(h + 34, Math.min(end, body + len)).slice();
          return { bytes, ext: metas[type], pending: true, inflate: compressed };
        }
        p = body + len;
        continue;
      }
      p = body + len;
    }
    return null;
  }
}

// metafiles vêm comprimidos com deflate dentro do .doc
async function inflateMedia(model) {
  for (const [path, m] of model.media) {
    if (!m.inflate) continue;
    try {
      const ds = new DecompressionStream('deflate');
      const out = await new Response(new Blob([m.bytes]).stream().pipeThrough(ds)).arrayBuffer();
      m.bytes = new Uint8Array(out);
      model.setMediaUrl(path, URL.createObjectURL(new Blob([m.bytes], { type: m.type })));
    } catch { /* fica a imagem original */ }
    m.inflate = null;
  }
}

export async function readDoc(bytes) {
  const { model } = await blankModel();
  model.format = 'doc';
  const prev = R.model;
  R.model = model;
  let rd;
  let body;
  try {
    rd = new DocReader(bytes, model);
    body = rd.build();
  } finally {
    R.model = prev;
  }
  await inflateMedia(model);
  // padrão do documento: a fonte e o tamanho do estilo Normal do .doc, sem espaço extra
  model.styles.setDefaultFont(rd.normalFont, rd.normalChp.hps || 20);
  const np = rd.normalPap;
  model.styles.setDefaultParagraphSpacing(np.before || 0, np.after || 0, np.lineMult ? np.line || 240 : 240);
  const s = rd.section;
  if (s.w && s.h) model.section.setPageSize(s.w, s.h, s.landscape || s.w > s.h ? 'landscape' : null);
  if (s.left != null || s.top != null || s.right != null || s.bottom != null) {
    const m = model.section.pgMar;
    model.section.setMargins({ top: s.top ?? m.top, bottom: s.bottom ?? m.bottom, left: s.left ?? m.left, right: s.right ?? m.right, header: s.header ?? m.header, footer: s.footer ?? m.footer });
  }
  if (s.titlePg) model.section.setTitlePg(true);
  for (const hf of rd.headers) {
    if (hf.type === 'first' && !s.titlePg) continue;
    const entry = newHeaderFooter(model, hf.kind);
    entry.doc = schema.node('doc', null, hf.blocks);
    entry.dirty = true;
    model.section.setReference(hf.kind, hf.type, entry.rId);
  }
  const blocks = body.length ? body : [nodes.paragraph.create()];
  return { model, doc: finalizeDoc(schema.node('doc', null, blocks), model) };
}
