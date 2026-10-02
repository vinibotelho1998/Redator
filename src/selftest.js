// Autoteste interno, acionado com ?selftest=1. Não faz parte do uso normal.
import { TextSelection, AllSelection } from 'prosemirror-state';
import { readDocx } from './docx/reader.js';
import { writeDocx } from './docx/writer.js';
import { writeMarkdown, readMarkdown } from './formats/markdown.js';
import { writeHtml, readHtml } from './formats/html.js';
import { writeText, readText } from './formats/text.js';
import { writeOdt, readOdt } from './formats/odt.js';
import { writeRtf, readRtf } from './formats/rtf.js';
import { lossesForFormat } from './formats/compatibility.js';
import { R } from './editor/render.js';
import * as C from './editor/commands.js';
import JSZip from 'jszip';
import { parseXml } from './util/xml.js';

const log = (...a) => console.log('SELFTEST', ...a);

function normalize(json) {
  json = JSON.parse(JSON.stringify(json));
  // remove atributos que mudam legitimamente numa regravação
  const walk = (n) => {
    if (n.attrs) {
      delete n.attrs.pageProperties; // histórico local; propriedades de página são conferidas no modelo
      delete n.attrs.xml; delete n.attrs.srcPart; delete n.attrs.rId; delete n.attrs.ocx; delete n.attrs.ocy;
      if (n.type === 'drawing') { delete n.attrs.mediaPath; delete n.attrs.name; }
      if (n.type === 'opaque_inline' && n.attrs.xml) delete n.attrs.xml;
      if (n.type === 'table') delete n.attrs.grid;
      if (n.type === 'table_cell') { delete n.attrs.colwidth; delete n.attrs.tcW; }
      if (n.type === 'table_row') { delete n.attrs.gridBefore; delete n.attrs.gridAfter; }
    }
    if (n.marks) for (const m of n.marks) if (m.type === 'hyperlink' && m.attrs.a) delete m.attrs.a.history;
    if (n.content) n.content.forEach(walk);
  };
  walk(json);
  return json;
}

function firstDiff(a, b, path = '') {
  if (typeof a !== typeof b) return path + ' tipo ' + typeof a + ' vs ' + typeof b;
  if (a && typeof a === 'object') {
    if (Array.isArray(a) !== Array.isArray(b)) return path + ' array';
    const keys = new Set([...Object.keys(a || {}), ...Object.keys(b || {})]);
    for (const k of keys) {
      const empty = (v) => v == null || (Array.isArray(v) && v.length === 0);
      if (empty(a[k]) && empty(b[k])) continue;
      const d = firstDiff(a[k], b[k], path + '/' + k);
      if (d) return d;
    }
    return null;
  }
  return a === b ? null : path + ': ' + JSON.stringify(a).slice(0, 120) + ' vs ' + JSON.stringify(b).slice(0, 120);
}

async function validateZip(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const bad = [];
  for (const [name, f] of Object.entries(zip.files)) {
    if (f.dir || !/\.(xml|rels)$/.test(name)) continue;
    try { parseXml(await f.async('string')); } catch (e) { bad.push(name + ': ' + e.message.slice(0, 200)); }
  }
  return bad;
}

// ganchos usados pelos roteiros de teste locais
export const testApi = {
  readDocx, writeDocx, readMarkdown, writeMarkdown, readHtml, writeHtml, readText, writeText, readOdt, writeOdt, readRtf, writeRtf,
  normalize, firstDiff, validateZip, JSZip, TextSelection, lossesForFormat,
};

export async function runSelfTest(app) {
  const t0 = performance.now();
  const view = app.view;
  const model = app.model;
  const doc0 = view.state.doc;
  log('páginas', app.layout.pageCount, 'parágrafos', doc0.content.childCount, 'palavras', app.wordCount);
  try {
    const bytes = await writeDocx(model, doc0, { author: 'Teste' });
    const bad = await validateZip(bytes);
    log('docx gravado', bytes.length, 'bytes; XML inválido:', bad.length ? bad.join(' | ') : 'nenhum');
    const again = await readDocx(bytes);
    const a = normalize(doc0.toJSON());
    const b = normalize(again.doc.toJSON());
    const d = firstDiff(a, b);
    log('ida e volta docx:', d ? 'DIFERENÇA ' + d : 'idêntico');
    R.model = model;
  } catch (e) { log('ERRO docx', e && e.stack); R.model = model; }
  for (const [name, w, r] of [
    ['md', (d) => writeMarkdown(d, model, {}).text, (bytes) => readMarkdown(bytes, '')],
    ['html', (d) => writeHtml(d, model, {}), (bytes) => readHtml(bytes, '')],
    ['txt', (d) => writeText(d, model), (bytes) => readText(bytes, '')],
    ['odt', (d) => writeOdt(d, model), (bytes) => readOdt(bytes)],
  ]) {
    try {
      const out = await w(doc0);
      const bytes = typeof out === 'string' ? new TextEncoder().encode(out) : out;
      const back = await r(bytes);
      R.model = model;
      log('formato', name, 'ok', bytes.length, 'bytes, reaberto com', back.doc.content.childCount, 'blocos');
      if (name === 'md') log('md início:', JSON.stringify(out.slice(0, 400)));
    } catch (e) { R.model = model; log('ERRO formato', name, e && e.stack); }
  }
  // edição
  try {
    const st = view.state;
    const para = [];
    st.doc.descendants((n, pos) => { if (n.type.name === 'paragraph' && n.textContent.length > 20 && para.length < 3) para.push(pos); return para.length < 3; });
    if (para.length) {
      const p = para[0];
      view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, p + 1, p + 10)));
      C.toggleRunProp('b')(view.state, view.dispatch);
      C.toggleRunProp('i')(view.state, view.dispatch);
      C.setFontSize(16)(view.state, view.dispatch);
      C.setColor({ val: 'C00000' })(view.state, view.dispatch);
      C.setAlign('center')(view.state, view.dispatch);
      C.toggleList('bullet')(view.state, view.dispatch);
      C.toggleList('number')(view.state, view.dispatch);
      C.changeIndent(1)(view.state, view.dispatch);
      view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, p + 5)));
      C.splitParagraph(view.state, view.dispatch);
      C.insertTable(3, 4)(view.state, view.dispatch);
      C.pageBreak(view.state, view.dispatch);
      view.dispatch(view.state.tr.insertText('Texto de teste com acentuação: ação, pão, você.'));
      log('edição ok; parágrafos agora', view.state.doc.content.childCount);
      app.layout.runNow(6);
      log('layout após edição: páginas', app.layout.pageCount);
      const bytes = await writeDocx(model, view.state.doc, {});
      const bad = await validateZip(bytes);
      log('docx após edição', bytes.length, 'bytes; XML inválido:', bad.length ? bad.join(' | ') : 'nenhum');
      const back = await readDocx(bytes);
      R.model = model;
      const d = firstDiff(normalize(view.state.doc.toJSON()), normalize(back.doc.toJSON()));
      log('ida e volta após edição:', d ? 'DIFERENÇA ' + d : 'idêntico');
      if (d) {
        const m = /^\/content\/(\d+)/.exec(d);
        if (m) {
          const i = Number(m[1]);
          log('antes:', JSON.stringify(view.state.doc.child(i).toJSON()).slice(0, 900));
          log('depois:', JSON.stringify(back.doc.child(i).toJSON()).slice(0, 900));
          const x = await JSZip.loadAsync(bytes);
          const xml = await x.file('word/document.xml').async('string');
          const k = xml.indexOf('<w:body>');
          log('xml:', xml.slice(k, k + 1500));
        }
      }
    }
  } catch (e) { R.model = model; log('ERRO edição', e && e.stack); }
  log('tempo total', Math.round(performance.now() - t0), 'ms');
  log('FIM');
}
