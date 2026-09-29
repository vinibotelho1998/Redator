// HTML: importação e exportação como página independente.
import { DOMParser as PMDOMParser } from 'prosemirror-model';
import { schema } from '../editor/schema.js';
import { R } from '../editor/render.js';
import { normalizeExternalHTML } from '../editor/htmlparse.js';
import { renderStatic } from '../editor/static.js';
import { layoutTabs } from '../editor/layout.js';
import { blankModel, finalizeDoc, decodeText, dirName, resolveRelative, loadImageSize, fitSize, contentWidthPx } from './common.js';

export async function readHtml(bytes, path) {
  const { text, encoding, bom, eol } = decodeText(bytes);
  const { model } = await blankModel();
  model.format = 'html';
  model.extra.text = { encoding, bom, eol };
  const dom = new DOMParser().parseFromString(text, 'text/html');
  model.properties.title = (dom.querySelector('title') || {}).textContent || '';
  const baseDir = dirName(path);
  const maxW = contentWidthPx(model);
  const imgs = Array.from(dom.body.querySelectorAll('img'));
  await Promise.all(imgs.map(async (img) => {
    const src = resolveRelative(baseDir, img.getAttribute('src'));
    img.setAttribute('src', src);
    if (!img.getAttribute('width')) {
      const size = await loadImageSize(src, 3000);
      if (size) {
        const { cx, cy } = fitSize(size.w, size.h, maxW);
        img.setAttribute('width', String(Math.round(cx / 9525)));
        img.setAttribute('height', String(Math.round(cy / 9525)));
      }
    }
  }));
  const prev = R.model;
  R.model = model;
  try {
    normalizeExternalHTML(dom.body, { codeStyle: 'Code', quoteStyle: 'BlockQuote' });
    const doc = PMDOMParser.fromSchema(schema).parse(dom.body);
    return { model, doc: finalizeDoc(doc, model) };
  } finally {
    R.model = prev;
  }
}

function bytesToBase64(bytes) {
  let s = '';
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return btoa(s);
}

const BASE_CSS = `
body{margin:0;background:#f3f3f3}
.page{max-width:{W}px;margin:24px auto;background:#fff;padding:{PT}px {PR}px {PB}px {PL}px;box-shadow:0 1px 4px rgba(0,0,0,.15)}
.doc-body p{margin:0;white-space:pre-wrap;position:relative}
.doc-body p,.doc-body p span{line-height:var(--lhx,calc(var(--lf,1.22) * var(--lh,1)))}
.doc-body table{border-collapse:collapse;table-layout:fixed}
.doc-body td{vertical-align:top;border-top:var(--tb-insideH,none);border-bottom:var(--tb-insideH,none);border-left:var(--tb-insideV,none);border-right:var(--tb-insideV,none);padding:var(--tc-top,0) var(--tc-right,5.4pt) var(--tc-bottom,0) var(--tc-left,5.4pt)}
.doc-body td.r-first{border-top:var(--tb-top,none)}.doc-body td.r-last{border-bottom:var(--tb-bottom,none)}
.doc-body td.c-first{border-left:var(--tb-left,none)}.doc-body td.c-last{border-right:var(--tb-right,none)}
.doc-body table.tbl{--tb-top:none;--tb-left:none;--tb-bottom:none;--tb-right:none;--tb-insideH:none;--tb-insideV:none}
.doc-body .tab{display:inline-block;height:1em}
.doc-body .num-lbl{white-space:pre}.doc-body .num-tab{display:inline-block}
.doc-body .sup{position:relative;top:-.5em;font-size:65%}.doc-body .sub{position:relative;top:.25em;font-size:65%}
.doc-body a.lnk{color:inherit;text-decoration:inherit}
.doc-body .drw-box{display:inline-block;position:relative}.doc-body .drw-box img{width:100%;height:100%;display:block}
.doc-body .fc,.doc-body .fi,.doc-body .oi{display:none}
.doc-body .vanish{display:none}
@media print{body{background:#fff}.page{box-shadow:none;margin:0;max-width:none}}
`;

export function writeHtml(doc, model, opts = {}) {
  const prev = R.model;
  R.model = model;
  try {
    const holder = renderStatic(doc);
    // mede as tabulações num contêiner invisível
    const s = model.section;
    const measure = document.createElement('div');
    measure.className = 'doc-body static-doc';
    measure.style.cssText = 'position:absolute;left:-99999px;top:0;visibility:hidden;width:' + contentWidthPx(model) + 'px';
    measure.appendChild(holder);
    document.body.appendChild(measure);
    layoutTabs(holder, 1);
    for (const t of holder.querySelectorAll('.tab,.num-tab')) t.style.width = t.style.width || '0px';
    measure.remove();
    // imagens incorporadas
    for (const img of holder.querySelectorAll('img')) {
      const src = img.getAttribute('src') || '';
      if (src.startsWith('blob:')) {
        for (const m of model.media.values()) {
          if (m.url === src && m.bytes && !/emf|wmf|tif/.test(m.ext)) { img.setAttribute('src', 'data:' + m.type + ';base64,' + bytesToBase64(m.bytes)); break; }
        }
      }
    }
    for (const el of holder.querySelectorAll('[data-pa],[data-ta],[data-ra],[data-ca],[data-fa],[data-oa],[data-fn],[data-tab],[data-sym],[data-ob]')) {
      for (const a of ['data-pa', 'data-ta', 'data-ra', 'data-ca', 'data-fa', 'data-oa', 'data-fn', 'data-tab', 'data-sym', 'data-ob']) el.removeAttribute(a);
    }
    const css = BASE_CSS.replace('{W}', String(Math.round(s.pgSz.w / 15))).replace('{PT}', String(Math.round(s.pgMar.top / 15)))
      .replace('{PR}', String(Math.round(s.pgMar.right / 15))).replace('{PB}', String(Math.round(s.pgMar.bottom / 15))).replace('{PL}', String(Math.round(s.pgMar.left / 15)));
    const title = (model.properties && model.properties.title) || opts.title || 'Documento';
    const esc = (t) => String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;');
    const html = '<!DOCTYPE html>\n<html lang="pt-BR">\n<head>\n<meta charset="utf-8">\n<meta name="viewport" content="width=device-width, initial-scale=1">\n<meta name="generator" content="Redator">\n<title>' + esc(title) + '</title>\n<style>\n@layer base, docdefaults, tablestyle, parastyle, charstyle;\n'
      + model.styles.css() + '\n' + css + '</style>\n</head>\n<body>\n<div class="page doc-body">\n' + holder.innerHTML + '\n</div>\n</body>\n</html>\n';
    return html;
  } finally {
    R.model = prev;
  }
}
