// Gravação de .docx a partir do documento do editor, preservando o que não foi alterado.
import JSZip from 'jszip';
import { TableMap } from 'prosemirror-tables';
import { schema, RUN_MARKS, WRAPPER_MARKS } from '../editor/schema.js';
import { pPrXml, rPrXml, tblPrXml, trPrXml, tcPrXml } from './props.js';
import { REL, escText, escAttr, parseXml, relsPathFor, relativeTarget, resolvePartPath, NS } from '../util/xml.js';
import { NS_DECL, blankDocxBytes } from './template.js';
import { MIME, extOf, parseRels } from './reader.js';
import { columnWidths } from '../editor/nodeviews.js';

const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

const CT = {
  header: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
  footer: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
  numbering: 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml',
  styles: 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml',
};

// ---------------------------------------------------------------- relacionamentos

class Rels {
  constructor(map, partPath) {
    this.map = new Map(map || []);
    this.part = partPath;
    this.dirty = false;
  }

  nextId() {
    let max = 0;
    for (const id of this.map.keys()) {
      const m = /^rId(\d+)$/.exec(id);
      if (m) max = Math.max(max, Number(m[1]));
    }
    return 'rId' + (max + 1);
  }

  find(type, target, external) {
    for (const r of this.map.values()) {
      if (r.type !== type) continue;
      if (external ? r.target === target : r.path === target) return r.id;
    }
    return null;
  }

  add(type, targetPath, external) {
    const found = this.find(type, targetPath, external);
    if (found) return found;
    const id = this.nextId();
    const target = external ? targetPath : relativeTarget(this.part, targetPath);
    this.map.set(id, { id, type, target, mode: external ? 'External' : null, path: external ? targetPath : targetPath });
    this.dirty = true;
    return id;
  }

  toXml() {
    let s = XML_HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
    for (const r of this.map.values()) {
      s += '<Relationship Id="' + escAttr(r.id) + '" Type="' + escAttr(r.type) + '" Target="' + escAttr(r.target) + '"' + (r.mode ? ' TargetMode="' + r.mode + '"' : '') + '/>';
    }
    return s + '</Relationships>';
  }
}

// ---------------------------------------------------------------- contexto

class WCtx {
  constructor(writer, part, rels) {
    this.w = writer;
    this.part = part;
    this.rels = rels;
  }
}

class Writer {
  constructor(model, zip) {
    this.model = model;
    this.zip = zip;
    this.docPrId = 1;
    this.contentTypes = null;
    this.newDefaults = new Map();
    this.newOverrides = new Map();
    this.mediaCounter = 0;
    this.pending = [];
  }

  // Garante uma mídia no pacote e devolve o caminho
  async ensureMedia(attrs) {
    const model = this.model;
    if (attrs.mediaPath && model.media.has(attrs.mediaPath)) {
      const m = model.media.get(attrs.mediaPath);
      if (!this.zip.file(attrs.mediaPath) && m.bytes) this.zip.file(attrs.mediaPath, m.bytes);
      this.ensureDefault(extOf(attrs.mediaPath), m.type);
      return attrs.mediaPath;
    }
    let bytes = null;
    let type = 'image/png';
    const src = attrs.src || '';
    // imagem de link ou colada já incorporada numa gravação anterior deste documento
    const known = src && model.srcMedia ? model.srcMedia.get(src) : null;
    if (known && this.zip.file(known) && model.media.has(known)) {
      this.ensureDefault(extOf(known), model.media.get(known).type);
      return known;
    }
    try {
      if (src.startsWith('data:')) {
        const m = /^data:([^;,]+)?(;base64)?,(.*)$/s.exec(src);
        if (m) {
          type = m[1] || 'image/png';
          if (m[2]) {
            const bin = atob(m[3]);
            bytes = new Uint8Array(bin.length);
            for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
          } else bytes = new TextEncoder().encode(decodeURIComponent(m[3]));
        }
      } else if (src.startsWith('blob:') || src.startsWith('http') || src.startsWith('file:')) {
        const res = await fetch(src);
        if (res.ok) {
          type = (res.headers.get('content-type') || 'image/png').split(';')[0];
          bytes = new Uint8Array(await res.arrayBuffer());
        }
      }
    } catch (err) {
      console.warn('Imagem não pôde ser incorporada', src.slice(0, 80), err);
    }
    if (!bytes) return null;
    const ext = Object.entries(MIME).find(([, v]) => v === type)?.[0] || 'png';
    let path;
    do { this.mediaCounter++; path = 'word/media/redator' + this.mediaCounter + '.' + ext; } while (this.zip.file(path) || model.media.has(path));
    this.zip.file(path, bytes);
    model.media.set(path, { url: src.startsWith('blob:') ? src : URL.createObjectURL(new Blob([bytes], { type })), bytes, ext, type });
    if (!model.srcMedia) model.srcMedia = new Map();
    model.srcMedia.set(src, path);
    this.ensureDefault(ext, type);
    return path;
  }

  ensureDefault(ext, type) {
    if (!ext) return;
    this.newDefaults.set(ext.toLowerCase(), type || MIME[ext] || 'application/octet-stream');
  }

  nextDocPr() {
    return this.docPrId++;
  }
}

// ---------------------------------------------------------------- texto

function textRuns(text) {
  let out = '';
  let buf = '';
  const flush = () => { if (buf) { out += '<w:t xml:space="preserve">' + escText(buf) + '</w:t>'; buf = ''; } };
  for (const ch of text) {
    if (ch === '\t') { flush(); out += '<w:tab/>'; }
    else if (ch === '\n') { flush(); out += '<w:br/>'; }
    else if (ch === '‑') { flush(); out += '<w:noBreakHyphen/>'; }
    else if (ch === '­') { flush(); out += '<w:softHyphen/>'; }
    else buf += ch;
  }
  flush();
  return out;
}

function wrapperXml(mark, ctx) {
  const t = mark.type.name;
  const a = mark.attrs;
  if (t === 'hyperlink') {
    let attrs = '';
    const extra = a.a || {};
    for (const [k, v] of Object.entries(extra)) if (k !== 'anchor' && k !== 'id' && v != null) attrs += ' w:' + k + '="' + escAttr(v) + '"';
    if (a.href) {
      const id = ctx.rels.add(REL.hyperlink, a.href, true);
      attrs = ' r:id="' + id + '"' + attrs;
    }
    if (a.anchor) attrs += ' w:anchor="' + escAttr(a.anchor) + '"';
    if (!/w:history/.test(attrs) && a.href) attrs += ' w:history="1"';
    return { open: '<w:hyperlink' + attrs + '>', close: '</w:hyperlink>' };
  }
  if (t === 'sdt') {
    return { open: '<w:sdt>' + (a.pr || '<w:sdtPr/>') + (a.endPr || '') + '<w:sdtContent>', close: '</w:sdtContent></w:sdt>' };
  }
  if (t === 'ins') {
    const tag = a.tag === 'moveTo' ? 'w:moveTo' : 'w:ins';
    let attrs = '';
    for (const [k, v] of Object.entries(a.a || {})) if (v != null) attrs += ' w:' + k + '="' + escAttr(v) + '"';
    return { open: '<' + tag + attrs + '>', close: '</' + tag + '>' };
  }
  return { open: '', close: '' };
}

function fixDocPr(xml, ctx) {
  return xml.replace(/(<wp:docPr\b[^>]*?\bid=")(\d+)(")/g, (m, a, id, b) => a + ctx.w.nextDocPr() + b);
}

async function drawingXml(node, ctx) {
  const a = node.attrs;
  let xml = a.xml;
  const resized = xml && (a.cx !== a.ocx || a.cy !== a.ocy);
  if (xml) {
    if (a.kind === 'image' && a.mediaPath) {
      // confere o id da imagem neste pacote
      const rel = a.rId ? ctx.rels.map.get(a.rId) : null;
      if (!rel || rel.path !== a.mediaPath) {
        const path = await ctx.w.ensureMedia(a);
        if (path) {
          const id = ctx.rels.add(REL.image, path, false);
          xml = xml.replace(/(<a:blip\b[^>]*?r:embed=")([^"]*)(")/, (m, p1, _old, p3) => p1 + id + p3).replace(/<a:extLst>[\s\S]*?<\/a:extLst>(?=\s*<\/a:blip>)/, '');
        }
      }
    }
    if (resized) {
      if (xml.startsWith('<mc:AlternateContent')) {
        const m = /<w:drawing[\s\S]*<\/w:drawing>/.exec(xml);
        if (m) xml = m[0];
      }
      xml = xml.replace(/(<wp:extent\b[^>]*?\bcx=")\d+("[^>]*?\bcy=")\d+(")/, (m, p1, p2, p3) => p1 + a.cx + p2 + a.cy + p3);
      xml = xml.replace(/(<a:xfrm\b[^>]*>[\s\S]*?<a:ext\b[^>]*?\bcx=")\d+("[^>]*?\bcy=")\d+(")/, (m, p1, p2, p3) => p1 + a.cx + p2 + a.cy + p3);
      xml = xml.replace(/(<v:shape\b[^>]*?style="[^"]*?width:)[\d.]+(pt;height:)[\d.]+(pt)/, (m, p1, p2, p3) => p1 + (a.cx / 12700).toFixed(2) + p2 + (a.cy / 12700).toFixed(2) + p3);
    }
    return fixDocPr(xml, ctx);
  }
  // imagem nova ou vinda de outro documento
  if (a.kind === 'hr') return '';
  const path = await ctx.w.ensureMedia(a);
  if (!path) {
    if (a.src && /^https?:/.test(a.src)) {
      const id = ctx.rels.add(REL.image, a.src, true);
      return newImageXml(a, ctx, id, true);
    }
    return a.alt ? '<w:t xml:space="preserve">' + escText('[' + a.alt + ']') + '</w:t>' : '';
  }
  const id = ctx.rels.add(REL.image, path, false);
  return newImageXml(a, ctx, id, false);
}

function newImageXml(a, ctx, rId, linked) {
  const cx = Math.max(1, Math.round(a.cx || 914400));
  const cy = Math.max(1, Math.round(a.cy || 685800));
  const id = ctx.w.nextDocPr();
  const name = 'Imagem ' + id;
  const blip = linked ? '<a:blip r:link="' + rId + '"/>' : '<a:blip r:embed="' + rId + '"/>';
  return '<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="' + cx + '" cy="' + cy + '"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
    + '<wp:docPr id="' + id + '" name="' + name + '"' + (a.alt ? ' descr="' + escAttr(a.alt) + '"' : '') + '/>'
    + '<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="' + NS.a + '" noChangeAspect="1"/></wp:cNvGraphicFramePr>'
    + '<a:graphic xmlns:a="' + NS.a + '"><a:graphicData uri="' + NS.pic + '"><pic:pic xmlns:pic="' + NS.pic + '">'
    + '<pic:nvPicPr><pic:cNvPr id="0" name="' + name + '"/><pic:cNvPicPr/></pic:nvPicPr>'
    + '<pic:blipFill>' + blip + '<a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
    + '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
    + '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>';
}

async function inlineXml(parent, ctx) {
  let out = '';
  const stack = [];
  let run = null;
  const flushRun = () => {
    if (run) { out += '<w:r>' + run.rpr + run.content + '</w:r>'; run = null; }
  };
  const syncWrappers = (marks) => {
    const wrappers = marks.filter((m) => WRAPPER_MARKS.includes(m.type.name));
    let keep = 0;
    while (keep < stack.length && keep < wrappers.length && stack[keep].mark.eq(wrappers[keep])) keep++;
    if (keep < stack.length) {
      flushRun();
      while (stack.length > keep) out += stack.pop().close;
    }
    for (let i = keep; i < wrappers.length; i++) {
      flushRun();
      const w = wrapperXml(wrappers[i], ctx);
      out += w.open;
      stack.push({ mark: wrappers[i], close: w.close });
    }
  };
  const children = [];
  parent.forEach((c) => children.push(c));
  for (const child of children) {
    syncWrappers(child.marks);
    const rpr = rPrXml(child.marks.filter((m) => RUN_MARKS.includes(m.type.name)));
    const add = (content) => {
      if (!content) return;
      if (!run || run.rpr !== rpr) { flushRun(); run = { rpr, content: '' }; }
      run.content += content;
    };
    const a = child.attrs;
    switch (child.type.name) {
      case 'text': add(textRuns(child.text)); break;
      case 'tab': add(a.ptab ? '<w:ptab w:relativeTo="margin" w:alignment="' + escAttr(a.ptab) + '" w:leader="none"/>' : '<w:tab/>'); break;
      case 'hard_break':
        if (a.type) add('<w:br w:type="' + a.type + '"/>');
        else if (a.clear) add('<w:br w:clear="' + escAttr(a.clear) + '"/>');
        else add('<w:br/>');
        break;
      case 'drawing': add(await drawingXml(child, ctx)); break;
      case 'field_char': add(a.xml || '<w:fldChar w:fldCharType="' + escAttr(a.type || 'begin') + '"/>'); break;
      case 'instr_text': add('<w:instrText xml:space="preserve">' + escText(a.text || '') + '</w:instrText>'); break;
      case 'footnote_ref': add('<w:' + (a.kind === 'endnote' ? 'endnote' : 'footnote') + 'Reference w:id="' + escAttr(a.id) + '"' + (a.custom ? ' w:customMarkFollows="1"' : '') + '/>'); break;
      case 'sym': add('<w:sym w:font="' + escAttr(a.font || 'Symbol') + '" w:char="' + escAttr(a.char || '') + '"/>'); break;
      case 'opaque_inline':
        if (a.level === 'r') add(a.xml || '');
        else { flushRun(); out += a.xml || ''; }
        break;
      default: break;
    }
  }
  flushRun();
  while (stack.length) out += stack.pop().close;
  return out;
}

async function paragraphXml(node, ctx) {
  const a = { ...node.attrs };
  if (a.numId && String(a.numId).startsWith('__')) { a.numId = null; a.ilvl = null; }
  if (a.styleId && ctx.w.model.styles && !ctx.w.model.styles.get(a.styleId)) a.styleId = null;
  return '<w:p>' + pPrXml(a) + (await inlineXml(node, ctx)) + '</w:p>';
}

async function tableXml(node, ctx) {
  const a = node.attrs;
  let map;
  try { map = TableMap.get(node); } catch { map = null; }
  const widthsPx = columnWidths(node);
  const grid = widthsPx.map((w) => Math.max(1, Math.round(w * 15)));
  const { grid: _g, ...tprops } = a;
  let x = '<w:tbl>' + tblPrXml(tprops) + '<w:tblGrid>' + grid.map((w) => '<w:gridCol w:w="' + w + '"/>').join('') + '</w:tblGrid>';
  const rows = [];
  node.forEach((row) => rows.push(row));
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    const cellsXml = [];
    let gridBefore = 0;
    let gridAfter = 0;
    let seenReal = false;
    const pendingPhantom = [];
    if (map) {
      for (let col = 0; col < map.width;) {
        const cellPos = map.map[r * map.width + col];
        const cell = node.nodeAt(cellPos);
        const rect = map.findCell(cellPos);
        const span = rect.right - rect.left;
        const wTw = grid.slice(rect.left, rect.right).reduce((s, v) => s + v, 0);
        if (rect.top === r) {
          if (cell.attrs.phantom && cell.textContent === '' && cell.childCount <= 1) {
            if (!seenReal) gridBefore += span;
            else pendingPhantom.push(span);
          } else {
            if (pendingPhantom.length) {
              for (const sp of pendingPhantom) cellsXml.push('<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/>' + (sp > 1 ? '<w:gridSpan w:val="' + sp + '"/>' : '') + '<w:tcBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders></w:tcPr><w:p/></w:tc>');
              pendingPhantom.length = 0;
            }
            seenReal = true;
            const vm = rect.bottom - rect.top > 1 ? 'restart' : null;
            const cattrs = { ...cell.attrs, tcW: { w: String(wTw), type: 'dxa' } };
            let content = await blocksXml(cell.content, ctx);
            if (!cell.lastChild || cell.lastChild.type.name !== 'paragraph') content += '<w:p/>';
            cellsXml.push('<w:tc>' + tcPrXml(cattrs, span, vm) + content + '</w:tc>');
          }
        } else {
          // continuação de célula mesclada verticalmente
          seenReal = true;
          const cattrs = { ...cell.attrs, tcW: { w: String(wTw), type: 'dxa' } };
          cellsXml.push('<w:tc>' + tcPrXml({ tcW: cattrs.tcW, tcBorders: cattrs.tcBorders, shd: cattrs.shd, vAlign: cattrs.vAlign }, span, 'continue') + '<w:p/></w:tc>');
        }
        col = rect.right;
      }
      for (const sp of pendingPhantom) gridAfter += sp;
    } else {
      for (let i = 0; i < row.childCount; i++) {
        const cell = row.child(i);
        cellsXml.push('<w:tc>' + tcPrXml(cell.attrs, cell.attrs.colspan, null) + (await blocksXml(cell.content, ctx)) + '</w:tc>');
      }
    }
    const rattrs = { ...row.attrs, gridBefore: gridBefore || null, gridAfter: gridAfter || null };
    if (!cellsXml.length) cellsXml.push('<w:tc><w:p/></w:tc>');
    x += '<w:tr>' + trPrXml(rattrs) + cellsXml.join('') + '</w:tr>';
  }
  return x + '</w:tbl>';
}

async function blocksXml(fragment, ctx) {
  let out = '';
  const list = [];
  fragment.forEach((n) => list.push(n));
  for (const node of list) {
    switch (node.type.name) {
      case 'paragraph': out += await paragraphXml(node, ctx); break;
      case 'table': out += await tableXml(node, ctx); break;
      case 'sdt_block': out += '<w:sdt>' + (node.attrs.pr || '<w:sdtPr/>') + (node.attrs.endPr || '') + '<w:sdtContent>' + (await blocksXml(node.content, ctx)) + '</w:sdtContent></w:sdt>'; break;
      case 'opaque_block': out += node.attrs.xml || ''; break;
      default: break;
    }
  }
  return out;
}

// ---------------------------------------------------------------- tipos de conteúdo

async function updateContentTypes(zip, writer) {
  const f = zip.file('[Content_Types].xml');
  const text = f ? await f.async('string') : XML_HEAD + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>';
  const doc = parseXml(text);
  const root = doc.documentElement;
  const ns = 'http://schemas.openxmlformats.org/package/2006/content-types';
  const defaults = new Set(Array.from(root.getElementsByTagName('Default')).map((d) => (d.getAttribute('Extension') || '').toLowerCase()));
  const overrides = new Set(Array.from(root.getElementsByTagName('Override')).map((o) => o.getAttribute('PartName')));
  let changed = false;
  for (const [ext, type] of writer.newDefaults) {
    if (defaults.has(ext)) continue;
    const el = doc.createElementNS(ns, 'Default');
    el.setAttribute('Extension', ext);
    el.setAttribute('ContentType', type);
    root.insertBefore(el, root.firstChild);
    changed = true;
  }
  for (const [part, type] of writer.newOverrides) {
    if (overrides.has('/' + part)) continue;
    const el = doc.createElementNS(ns, 'Override');
    el.setAttribute('PartName', '/' + part);
    el.setAttribute('ContentType', type);
    root.appendChild(el);
    changed = true;
  }
  if (changed) {
    let s = new XMLSerializer().serializeToString(doc);
    if (!s.startsWith('<?xml')) s = XML_HEAD + s;
    zip.file('[Content_Types].xml', s);
  }
}

const SETTINGS_AFTER_EVENODD = ['bookFoldRevPrinting', 'bookFoldPrinting', 'bookFoldPrintingSheets', 'drawingGridHorizontalSpacing', 'drawingGridVerticalSpacing', 'displayHorizontalDrawingGridEvery', 'displayVerticalDrawingGridEvery', 'doNotUseMarginsForDrawingGridOrigin', 'drawingGridHorizontalOrigin', 'drawingGridVerticalOrigin', 'doNotShadeFormData', 'noPunctuationKerning', 'characterSpacingControl', 'printTwoOnOne', 'strictFirstAndLastChars', 'noLineBreaksAfter', 'noLineBreaksBefore', 'savePreviewPicture', 'doNotValidateAgainstSchema', 'saveInvalidXml', 'ignoreMixedContent', 'alwaysShowPlaceholderText', 'doNotDemarcateInvalidXml', 'saveXmlDataOnly', 'useXSLTWhenSaving', 'saveThroughXslt', 'showXMLTags', 'alwaysMergeEmptyNamespace', 'updateFields', 'hdrShapeDefaults', 'footnotePr', 'endnotePr', 'compat', 'docVars', 'rsids', 'mathPr', 'attachedSchema', 'themeFontLang', 'clrSchemeMapping', 'doNotIncludeSubdocsInStats', 'doNotAutoCompressPictures', 'forceUpgrade', 'captions', 'readModeInkLockDown', 'smartTagType', 'schemaLibrary', 'shapeDefaults', 'doNotEmbedSmartTags', 'decimalSymbol', 'listSeparator'];

async function updateSettings(zip, model, mainRels) {
  let path = null;
  for (const r of mainRels.map.values()) if (r.type === REL.settings) path = r.path;
  if (!path || !zip.file(path)) return;
  const doc = parseXml(await zip.file(path).async('string'));
  const root = doc.documentElement;
  const find = (name) => Array.from(root.children).find((c) => c.localName === name && c.namespaceURI === NS.w);
  const dts = find('defaultTabStop');
  if (dts) dts.setAttributeNS(NS.w, 'w:val', String(model.settings.defaultTabStop || 708));
  const eo = find('evenAndOddHeaders');
  if (model.settings.evenAndOdd && !eo) {
    const el = doc.createElementNS(NS.w, 'w:evenAndOddHeaders');
    const before = Array.from(root.children).find((c) => SETTINGS_AFTER_EVENODD.includes(c.localName));
    root.insertBefore(el, before || null);
  } else if (!model.settings.evenAndOdd && eo) eo.remove();
  let s = new XMLSerializer().serializeToString(doc);
  if (!s.startsWith('<?xml')) s = XML_HEAD + s;
  zip.file(path, s);
  model.settings.dirty = false;
}

async function updateCoreProps(zip, author) {
  const f = zip.file('docProps/core.xml');
  if (!f) return;
  let text = await f.async('string');
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  if (/<dcterms:modified[^>]*>[^<]*<\/dcterms:modified>/.test(text)) text = text.replace(/(<dcterms:modified[^>]*>)[^<]*(<\/dcterms:modified>)/, '$1' + now + '$2');
  if (author) {
    if (/<cp:lastModifiedBy>[^<]*<\/cp:lastModifiedBy>/.test(text)) text = text.replace(/(<cp:lastModifiedBy>)[^<]*(<\/cp:lastModifiedBy>)/, '$1' + escText(author) + '$2');
  }
  zip.file('docProps/core.xml', text);
}

// ---------------------------------------------------------------- entrada principal

export async function writeDocx(model, doc, opts = {}) {
  let zip = model.zip;
  if (!zip) {
    zip = await JSZip.loadAsync(await blankDocxBytes(opts.author || ''));
    model.zip = zip;
  }
  const writer = new Writer(model, zip);
  const mainPath = model.mainPath || 'word/document.xml';
  const mainRels = new Rels(model.rels, mainPath);
  const ctx = new WCtx(writer, mainPath, mainRels);

  // numeração e estilos
  if (model.numbering && (model.numbering.dirty || !model.numbering.text) && model.numbering.nums.size) {
    let path = null;
    for (const r of mainRels.map.values()) if (r.type === REL.numbering) path = r.path;
    if (!path) {
      path = 'word/numbering.xml';
      mainRels.add(REL.numbering, path, false);
      writer.newOverrides.set(path, CT.numbering);
    }
    zip.file(path, model.numbering.toXml());
    model.numbering.dirty = false;
  }
  if (model.styles && model.styles.dirty) {
    let path = null;
    for (const r of mainRels.map.values()) if (r.type === REL.styles) path = r.path;
    if (!path) { path = 'word/styles.xml'; mainRels.add(REL.styles, path, false); writer.newOverrides.set(path, CT.styles); }
    zip.file(path, model.styles.toXml());
    model.styles.dirty = false;
  }

  // cabeçalhos e rodapés alterados
  for (const [kind, map] of [['header', model.headers], ['footer', model.footers]]) {
    for (const entry of map.values()) {
      if (!entry.dirty && !entry.isNew) continue;
      const rels = new Rels(entry.rels || new Map(), entry.path);
      const hctx = new WCtx(writer, entry.path, rels);
      const tag = kind === 'header' ? 'w:hdr' : 'w:ftr';
      let body = await blocksXml(entry.doc.content, hctx);
      if (!body.trim()) body = '<w:p/>';
      zip.file(entry.path, XML_HEAD + '<' + tag + (entry.rootAttrs || ' ' + NS_DECL) + '>' + body + '</' + tag + '>');
      if (rels.dirty || entry.isNew) zip.file(relsPathFor(entry.path), rels.toXml());
      if (entry.isNew) {
        writer.newOverrides.set(entry.path, CT[kind]);
        mainRels.map.set(entry.rId, { id: entry.rId, type: kind === 'header' ? REL.header : REL.footer, target: relativeTarget(mainPath, entry.path), mode: null, path: entry.path });
        mainRels.dirty = true;
        entry.isNew = false;
      }
      entry.rels = rels.map;
      entry.dirty = false;
    }
  }

  // notas criadas por importação (ODT) quando o pacote ainda não tem a parte
  for (const [kind, map, relType, file] of [['footnote', model.footnotes, REL.footnotes, 'footnotes.xml'], ['endnote', model.endnotes, REL.endnotes, 'endnotes.xml']]) {
    if (![...map.values()].some((n) => n.isNew)) continue;
    let path = null;
    for (const r of mainRels.map.values()) if (r.type === relType) path = r.path;
    if (path && zip.file(path)) continue;
    path = 'word/' + file;
    mainRels.add(relType, path, false);
    writer.newOverrides.set(path, 'application/vnd.openxmlformats-officedocument.wordprocessingml.' + kind + 's+xml');
    const rels = new Rels(new Map(), path);
    const nctx = new WCtx(writer, path, rels);
    const sepPara = (tag) => '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:' + tag + '/></w:r></w:p>';
    let xml = XML_HEAD + '<w:' + kind + 's ' + NS_DECL + '>'
      + '<w:' + kind + ' w:type="separator" w:id="-1">' + sepPara('separator') + '</w:' + kind + '>'
      + '<w:' + kind + ' w:type="continuationSeparator" w:id="0">' + sepPara('continuationSeparator') + '</w:' + kind + '>';
    for (const [id, n] of map) xml += '<w:' + kind + ' w:id="' + escAttr(id) + '">' + (await blocksXml(n.doc.content, nctx)) + '</w:' + kind + '>';
    xml += '</w:' + kind + 's>';
    zip.file(path, xml);
    if (rels.dirty) zip.file(relsPathFor(path), rels.toXml());
    for (const n of map.values()) n.isNew = false;
  }

  // corpo
  let body = await blocksXml(doc.content, ctx);
  if (doc.lastChild && doc.lastChild.type.name !== 'paragraph') body += '<w:p/>';
  const sect = model.section ? model.section.toXml() : '';
  const rootAttrs = model.rootAttrs || ' ' + NS_DECL;
  zip.file(mainPath, XML_HEAD + '<w:document' + rootAttrs + '><w:body>' + body + sect + '</w:body></w:document>');

  if (mainRels.dirty) {
    zip.file(relsPathFor(mainPath), mainRels.toXml());
    model.rels = mainRels.map;
  }
  if (model.settings && model.settings.dirty) await updateSettings(zip, model, mainRels);
  await updateContentTypes(zip, writer);
  await updateCoreProps(zip, opts.author);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 }, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
}

// Cria uma parte nova de cabeçalho ou rodapé
export function newHeaderFooter(model, kind) {
  const map = kind === 'header' ? model.headers : model.footers;
  let n = 1;
  const exists = (p) => (model.zip && model.zip.file(p)) || [...model.headers.values(), ...model.footers.values()].some((e) => e.path === p);
  while (exists('word/' + kind + n + '.xml')) n++;
  const path = 'word/' + kind + n + '.xml';
  const rels = new Rels(model.rels, model.mainPath);
  const rId = rels.nextId();
  // reserva o id no mapa de relacionamentos do documento
  model.rels = new Map(model.rels);
  model.rels.set(rId, { id: rId, type: kind === 'header' ? REL.header : REL.footer, target: relativeTarget(model.mainPath, path), mode: null, path });
  const styleName = kind === 'header' ? 'header' : 'footer';
  const st = model.styles.findByName(styleName);
  const doc = schema.node('doc', null, [schema.nodes.paragraph.create({ styleId: st ? st.id : null })]);
  const entry = { rId, kind, path, doc, rootAttrs: ' ' + NS_DECL, rels: new Map(), dirty: true, isNew: true };
  map.set(rId, entry);
  return entry;
}
