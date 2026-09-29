// CSV: abre como tabela e volta igual. Cada linha lida guarda o texto
// original; ao salvar, a linha que não mudou é gravada como estava, com as
// mesmas aspas, e só as linhas editadas são montadas de novo.
import { schema } from '../editor/schema.js';
import { R } from '../editor/render.js';
import { blankModel, decodeText, encodeText, contentWidthPx } from './common.js';

const nodes = schema.nodes;
// acima disso a tabela pesa demais: o arquivo abre como texto
const MAX_ROWS = 3000;
const MAX_COLS = 40;

// Lê registros no padrão RFC 4180, guardando o texto de cada um
export function parseCsv(text, delim, limit = Infinity) {
  const records = [];
  let i = 0;
  let eol = null;
  const n = text.length;
  while (i < n && records.length < limit) {
    const start = i;
    const cells = [];
    const quoted = [];
    let cell = '';
    let q = false;
    let atStart = true;
    let ended = false;
    while (i < n) {
      const ch = text[i];
      if (q) {
        if (ch === '"') {
          if (text[i + 1] === '"') { cell += '"'; i += 2; continue; }
          q = false;
          i++;
          continue;
        }
        cell += ch;
        i++;
        continue;
      }
      if (ch === '"' && atStart) { q = true; quoted[cells.length] = true; atStart = false; i++; continue; }
      if (ch === delim) { cells.push(cell); cell = ''; atStart = true; i++; continue; }
      if (ch === '\r' || ch === '\n') {
        const len = ch === '\r' && text[i + 1] === '\n' ? 2 : 1;
        if (!eol) eol = text.substr(i, len);
        records.push({ raw: text.slice(start, i), cells: [...cells, cell], quoted: fill(quoted, cells.length + 1) });
        i += len;
        ended = true;
        break;
      }
      cell += ch;
      atStart = false;
      i++;
    }
    if (!ended) records.push({ raw: text.slice(start, i), cells: [...cells, cell], quoted: fill(quoted, cells.length + 1) });
  }
  const trailing = n > 0 && (text.endsWith('\n') || text.endsWith('\r'));
  return { records, eol, trailing };
}

function fill(arr, len) {
  const out = new Array(len).fill(false);
  arr.forEach((v, k) => { if (v) out[k] = true; });
  return out;
}

// Separador mais consistente entre vírgula, ponto e vírgula, tabulação e barra
export function detectDelimiter(text) {
  const sample = text.slice(0, 30000);
  let best = null;
  let bestScore = 0;
  for (const d of [',', ';', '\t', '|']) {
    if (!sample.includes(d)) continue;
    const { records } = parseCsv(sample, d, 40);
    const rows = records.filter((r) => r.raw.trim());
    if (rows.length < 1) continue;
    const counts = new Map();
    for (const r of rows) counts.set(r.cells.length, (counts.get(r.cells.length) || 0) + 1);
    const [mode, freq] = [...counts.entries()].sort((a, b) => b[1] - a[1])[0];
    if (mode < 2) continue;
    const score = (freq / rows.length) * Math.min(mode, 20);
    if (score > bestScore) { best = d; bestScore = score; }
  }
  return best;
}

const NUMERIC = /^[-+(]?\s*(R\$|US\$|\$|€)?\s*[\d.,]+\s*%?\)?$/;

export async function readCsv(bytes, path) {
  const { text, encoding, bom, eol } = decodeText(bytes);
  const tsv = /\.tsv$/i.test(path || '');
  const delim = tsv ? '\t' : detectDelimiter(text);
  if (!delim) return null;
  const parsed = parseCsv(text, delim);
  const records = parsed.records;
  // a última linha vazia é só o fim do arquivo
  const rows = records.length && records[records.length - 1].raw === '' && records[records.length - 1].cells.length === 1 ? records.slice(0, -1) : records;
  const cols = Math.max(0, ...rows.map((r) => r.cells.length));
  if (!rows.length || cols < 2 || rows.length > MAX_ROWS || cols > MAX_COLS) return null;
  const { model } = await blankModel();
  model.format = 'csv';
  model.extra.text = { encoding, bom, eol: parsed.eol || eol || '\r\n' };
  model.extra.csv = { delim, records: rows, trailing: parsed.trailing };
  model.styles.setDefaultParagraphSpacing(0, 0, 240);
  if (cols > 7) {
    // muitas colunas: página deitada
    const s = model.section;
    if (s.pgSz.w < s.pgSz.h) s.setPageSize(s.pgSz.h, s.pgSz.w, 'landscape');
  }
  const prev = R.model;
  R.model = model;
  try {
    // largura das colunas pelo tamanho do texto
    const widthPx = contentWidthPx(model);
    const len = new Array(cols).fill(3);
    for (const r of rows.slice(0, 300)) r.cells.forEach((c, k) => { len[k] = Math.max(len[k], Math.min(40, c.length)); });
    const total = len.reduce((a, b) => a + b, 0);
    const grid = len.map((l) => Math.max(500, Math.round((l / total) * widthPx * 15)));
    // primeira linha é cabeçalho quando é só texto e a de baixo tem números
    const isNum = (v) => NUMERIC.test(v.trim()) && /\d/.test(v);
    const header = rows.length > 1 && !rows[0].cells.some(isNum) && rows[1].cells.some(isNum);
    const brd = { val: 'single', sz: '4', space: '0', color: 'auto' };
    const pmRows = rows.map((r, ri) => {
      const cells = [];
      for (let k = 0; k < cols; k++) {
        const v = r.cells[k] ?? '';
        const content = [];
        v.split(/\r\n|\r|\n/).forEach((part, j) => {
          if (j) content.push(nodes.hard_break.create());
          if (part) content.push(schema.text(part, header && ri === 0 ? [schema.marks.b.create()] : []));
        });
        const para = nodes.paragraph.create(isNum(v) && !(header && ri === 0) ? { jc: 'right' } : null, content);
        cells.push(nodes.table_cell.create({ colwidth: [grid[k] / 15], tcW: { w: String(grid[k]), type: 'dxa' }, shd: header && ri === 0 ? { val: 'clear', color: 'auto', fill: 'D9E2F3' } : null }, [para]));
      }
      return nodes.table_row.create(header && ri === 0 ? { header: true } : null, cells);
    });
    const table = nodes.table.create({ grid, tblW: { w: '0', type: 'auto' }, tblBorders: { top: brd, left: brd, bottom: brd, right: brd, insideH: brd, insideV: brd } }, pmRows);
    const doc = schema.node('doc', null, [table, nodes.paragraph.create()]);
    return { model, doc };
  } finally {
    R.model = prev;
  }
}

function cellText(cell) {
  return cell.textBetween(0, cell.content.size, '\n', (leaf) => (leaf.type.spec.leafText ? leaf.type.spec.leafText(leaf) : ''));
}

function quoteField(v, delim, wasQuoted) {
  if (wasQuoted || v.includes(delim) || v.includes('"') || /[\r\n]/.test(v) || /^\s|\s$/.test(v)) return '"' + v.replace(/"/g, '""') + '"';
  return v;
}

// mesmas células, ignorando as colunas vazias que a tabela acrescentou
function sameCells(rec, cells) {
  for (let k = 0; k < cells.length; k++) {
    const a = k < rec.cells.length ? rec.cells[k] : '';
    if (a.replace(/\r\n?/g, '\n') !== cells[k]) return false;
  }
  return rec.cells.length <= cells.length || rec.cells.slice(cells.length).every((c) => c === '');
}

export function writeCsv(doc, model) {
  let table = null;
  doc.forEach((n) => { if (!table && n.type.name === 'table') table = n; });
  if (!table) return null;
  const src = model.extra && model.extra.csv;
  const info = (model.extra && model.extra.text) || { encoding: 'utf-8', bom: false, eol: '\r\n' };
  const eol = info.eol || '\r\n';
  const delim = src ? src.delim : ';';
  const recs = src ? src.records : [];
  const out = [];
  let j = 0;
  table.forEach((row) => {
    const cells = [];
    row.forEach((cell) => cells.push(cellText(cell)));
    // procura a linha original um pouco à frente, para aguentar linhas apagadas ou inseridas
    let hit = -1;
    for (let k = j; k < Math.min(recs.length, j + 50); k++) if (sameCells(recs[k], cells)) { hit = k; break; }
    if (hit >= 0) {
      out.push(recs[hit].raw);
      j = hit + 1;
      return;
    }
    const ref = recs[j];
    let last = cells.length;
    // colunas vazias no fim que a linha original não tinha ficam de fora
    if (ref) while (last > ref.cells.length && cells[last - 1] === '') last--;
    out.push(cells.slice(0, last).map((v, k) => quoteField(v, delim, ref && ref.quoted[k])).join(delim));
    if (ref && !sameCells(ref, cells)) j++;
  });
  const trailing = src ? src.trailing : true;
  return encodeText(out.join(eol) + (trailing ? eol : ''), info);
}
