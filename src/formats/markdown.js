// Markdown: leitura para o documento e escrita de volta.
// Na leitura, cada bloco guarda o trecho original do arquivo e cada parágrafo sabe
// de que linhas veio. Na gravação, blocos intactos voltam exatamente como estavam,
// parágrafos editados trocam só as próprias linhas e apenas o que mudou de
// estrutura é reescrito. Salvar sem mexer não altera o arquivo.
import MarkdownIt from 'markdown-it';
import { DOMParser as PMDOMParser } from 'prosemirror-model';
import { schema } from '../editor/schema.js';
import { R, effectiveNumPr } from '../editor/render.js';
import { NumberingCounter } from '../docx/numbering.js';
import { normalizeExternalHTML } from '../editor/htmlparse.js';
import { blankModel, finalizeDoc, decodeText, dirName, resolveRelative, fileUrl, loadImageSize, fitSize, contentWidthPx } from './common.js';

const nodes = schema.nodes;
const marks = schema.marks;

export const FRONT_STYLE = 'FrontMatter';
const FRONT_STYLE_XML = '<w:style w:type="paragraph" w:customStyle="1" w:styleId="FrontMatter"><w:name w:val="Propriedades"/><w:basedOn w:val="Code"/><w:uiPriority w:val="99"/><w:semiHidden/>'
  + '<w:pPr><w:shd w:val="clear" w:color="auto" w:fill="F6F6F8"/></w:pPr><w:rPr><w:color w:val="5C5C70"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style>';

const IMAGE_EXT = /\.(png|jpe?g|gif|bmp|webp|svg|avif|ico)$/i;
const CALLOUT_RE = /^\[!(\p{L}[\p{L}\p{N}_-]*)\]([+-]?)[ \t]*/u;
const HTML_TAGS = new Set(('a abbr address area article aside audio b base bdi bdo blockquote body br button canvas caption cite code col colgroup data datalist dd del details dfn dialog div dl dt em embed fieldset figcaption figure font footer form h1 h2 h3 h4 h5 h6 head header hr html i iframe img input ins kbd label legend li link main map mark meta meter nav noscript object ol optgroup option output p param picture pre progress q rp rt ruby s samp script section select small source span strike strong style sub summary sup svg table tbody td template textarea tfoot th thead time title tr track u ul var video wbr center big tt').split(' '));

// ---------------------------------------------------------------- markdown-it

// [[Nota]], [[Nota|apelido]] e ![[anexo.png]] do Obsidian
function wikiLinks(md) {
  md.inline.ruler.before('link', 'wikilink', (state, silent) => {
    const src = state.src;
    let pos = state.pos;
    const embed = src.charCodeAt(pos) === 0x21;
    if (embed) pos++;
    if (src.charCodeAt(pos) !== 0x5b || src.charCodeAt(pos + 1) !== 0x5b) return false;
    const end = src.indexOf(']]', pos + 2);
    if (end < 0) return false;
    const inner = src.slice(pos + 2, end);
    if (!inner.trim() || /[\n[\]]/.test(inner)) return false;
    if (!silent) {
      const parts = inner.split(/\\?\|/);
      const tok = state.push('wikilink', '', 0);
      tok.meta = { target: parts[0], alias: parts.length > 1 ? parts.slice(1).join('|') : null, embed };
      tok.content = src.slice(state.pos, end + 2);
    }
    state.pos = end + 2;
    return true;
  });
}

function createMd() {
  const md = new MarkdownIt({ html: true, linkify: true, typographer: false, breaks: false });
  // mantém os endereços como foram escritos, sem codificar acentos
  md.normalizeLink = (url) => url;
  md.normalizeLinkText = (url) => url;
  md.use(wikiLinks);
  return md;
}

// ---------------------------------------------------------------- leitura

export async function readMarkdown(bytes, path, opts = {}) {
  const { text: raw, encoding, bom, eol } = decodeText(bytes);
  const { model } = await blankModel();
  model.format = 'md';
  model.extra.text = { encoding, bom, eol };
  const text = raw.replace(/\r\n?/g, '\n');
  const fm = /^---[ \t]*\n(?:([\s\S]*?)\n)?(?:---|\.\.\.)[ \t]*(?:\n|$)/.exec(text);
  const front = fm ? fm[1] || '' : null;
  const body = fm ? text.slice(fm[0].length) : text;
  const tokens = createMd().parse(body, {});
  const baseDir = dirName(path);
  // imagens: tamanho e anexos do Obsidian
  const sizes = new Map();
  const embeds = new Map();
  const srcs = new Set();
  const wanted = new Set();
  for (const t of tokens) {
    if (!t.children) continue;
    for (const c of t.children) {
      if (c.type === 'image') srcs.add(resolveRelative(baseDir, c.attrGet('src')));
      else if (c.type === 'wikilink' && c.meta.embed && IMAGE_EXT.test(c.meta.target)) wanted.add(c.meta.target);
    }
  }
  if (wanted.size && opts.host && opts.host.findNote && path) {
    await Promise.all([...wanted].map(async (target) => {
      try {
        const found = await opts.host.findNote(path, target);
        if (found) { const u = fileUrl(found); embeds.set(target, u); srcs.add(u); }
      } catch { /* sem o anexo */ }
    }));
  }
  await Promise.all([...srcs].map(async (s) => sizes.set(s, await loadImageSize(s))));
  if (front != null) model.styles.addStyleXml(FRONT_STYLE_XML);
  const conv = new MdConverter(model, baseDir, sizes, embeds, body);
  conv.convert(tokens);
  const { blocks, chunks } = conv.assemble(front, fm ? fm[0] : '');
  const numMap = new Map();
  const doc = finalizeDoc(schema.node('doc', null, blocks.length ? blocks : [nodes.paragraph.create()]), model, numMap);
  const lists = new Map();
  for (const [key, id] of numMap) if (conv.markers.has(key)) lists.set(String(id), conv.markers.get(key));
  model.extra.md = { chunks, total: blocks.length, lists, bound: false };
  bindMarkdownSource(model, doc);
  return { model, doc };
}

// Liga cada trecho original aos nós do documento. Chamado de novo quando o
// editor ajusta o documento ao abrir (tabelas incompletas, por exemplo).
export function bindMarkdownSource(model, doc) {
  const src = model.extra && model.extra.md;
  if (!src) return;
  if (doc.childCount !== src.total) { model.extra.md = null; return; }
  for (const c of src.chunks) {
    c.nodes = [];
    for (let k = 0; k < c.count; k++) c.nodes.push(doc.child(c.from + k));
    if (c.kind === 'item') {
      const first = c.nodes[0];
      c.numId = first && first.attrs.numId != null ? String(first.attrs.numId) : null;
      const m = /^( *)([-*+]|\d{1,9}[.)])( +|$)/.exec(c.text.slice(c.leadLen || 0));
      c.col = m ? m[0].length : 2;
    }
  }
  src.bound = true;
}

function lineContent(line) {
  const out = [];
  line.split('\t').forEach((part, i) => {
    if (i > 0) out.push(nodes.tab.create());
    const clean = part.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
    if (clean) out.push(schema.text(clean));
  });
  return out;
}

function rawBlock(text) {
  return nodes.opaque_block.create({ kind: 'md', text: text.replace(/\n+$/, '') });
}

function cloneToken(t) {
  return Object.assign(Object.create(Object.getPrototypeOf(t)), t);
}

class MdConverter {
  constructor(model, baseDir, sizes, embeds, body) {
    this.model = model;
    this.baseDir = baseDir;
    this.sizes = sizes;
    this.embeds = embeds || new Map();
    this.lines = body.split('\n');
    this.body = body;
    this.listStack = [];
    this.quotes = [];
    this.pendingHead = null;
    this.listCounter = 0;
    this.styles = model.styles;
    this.maxW = contentWidthPx(model);
    this.groups = [];
    this.out = [];
    this.src = [];
    this.markers = new Map();
  }

  push(node, info) {
    this.out.push(node);
    this.src.push(info || null);
  }

  headingStyle(level) {
    return this.styles.headingId(level) || 'Heading' + level;
  }

  begin(kind, t) {
    this.groups.push({ kind, line: t ? t.map[0] : null, end: t ? t.map[1] : null, from: this.out.length });
  }

  // contexto de citação: tipo de callout e profundidade de lista onde a citação começou
  quoteMd() {
    if (!this.quotes.length) return null;
    const md = { quote: true };
    for (let k = this.quotes.length - 1; k >= 0; k--) if (this.quotes[k].co) { md.callout = this.quotes[k].co.type; break; }
    const ql = this.quotes[0].listDepth;
    if (ql) md.qlist = ql;
    return md;
  }

  paraAttrs() {
    const a = {};
    if (this.listStack.length) {
      const top = this.listStack[this.listStack.length - 1];
      if (top.first) {
        a.numId = top.placeholder;
        a.ilvl = this.listStack.length - 1;
        top.first = false;
      } else {
        a.ind = { left: String(720 * this.listStack.length) };
      }
      a.styleId = 'ListParagraph';
    } else if (this.quotes.length) {
      a.styleId = 'BlockQuote';
    }
    const q = this.quoteMd();
    if (q) a.md = q;
    return a;
  }

  // de que linhas veio um parágrafo, para trocar só elas quando ele for editado
  lineInfo(t, inline, extra) {
    if (!t.map) return null;
    const parts = (inline.content || '').split('\n');
    return { kind: 'lines', a: t.map[0], b: t.map[1], first: parts[0], second: parts.length > 1 ? parts[1] : null, ...(extra || {}) };
  }

  convert(tokens) {
    let i = 0;
    while (i < tokens.length) {
      const t = tokens[i];
      if (t.map && t.nesting !== -1 && (t.level === 0 || (t.type === 'list_item_open' && t.level === 1))) this.begin(t.type === 'list_item_open' ? 'item' : 'block', t);
      switch (t.type) {
        case 'heading_open': {
          const level = Number(t.tag.slice(1));
          const inline = tokens[i + 1];
          const a = { styleId: this.headingStyle(level) };
          const q = this.quoteMd();
          if (q) a.md = q;
          const setext = t.markup === '=' || t.markup === '-';
          const info = this.lineInfo(t, inline, { heading: true });
          if (info && setext) info.b -= 1;
          this.push(nodes.paragraph.create(a, this.inline(inline.children)), info);
          i += 3;
          continue;
        }
        case 'paragraph_open': {
          const inline = tokens[i + 1];
          if (this.pendingHead) {
            this.calloutParagraphs(t, inline);
            i += 3;
            continue;
          }
          const content = this.inline(inline.children);
          const attrs = this.paraAttrs();
          // lista de tarefas
          if (attrs.numId && content.length && content[0].isText && /^\[[ xX]\] /.test(content[0].text)) {
            const checked = content[0].text[1].toLowerCase() === 'x';
            content[0] = schema.text((checked ? '☑ ' : '☐ ') + content[0].text.slice(4), content[0].marks);
          }
          this.push(nodes.paragraph.create(attrs, content), this.lineInfo(t, inline));
          i += 3;
          continue;
        }
        case 'bullet_list_open': case 'ordered_list_open': {
          const kind = t.type === 'bullet_list_open' ? 'bullet' : 'number';
          this.listCounter++;
          const id = String(this.listCounter);
          const start = kind === 'number' ? Number(t.attrGet('start') || 1) : 1;
          const key = '__' + kind + ':' + id;
          this.markers.set(key, t.markup || (kind === 'bullet' ? '-' : '.'));
          this.listStack.push({ kind, id, placeholder: key + (start !== 1 ? '@' + start : ''), first: false, line: t.map ? t.map[0] : null });
          i++;
          continue;
        }
        case 'bullet_list_close': case 'ordered_list_close':
          this.listStack.pop();
          i++;
          continue;
        case 'list_item_open':
          if (this.listStack.length) { const top = this.listStack[this.listStack.length - 1]; top.first = true; top.itemLine = t.map ? t.map[0] : null; }
          i++;
          continue;
        case 'list_item_close': {
          const top = this.listStack[this.listStack.length - 1];
          if (top && top.first) {
            const a = { numId: top.placeholder, ilvl: this.listStack.length - 1, styleId: 'ListParagraph' };
            const q = this.quoteMd();
            if (q) a.md = q;
            const info = top.itemLine != null ? { kind: 'lines', a: top.itemLine, b: top.itemLine + 1, first: '' } : null;
            this.push(nodes.paragraph.create(a), info);
            top.first = false;
          }
          i++;
          continue;
        }
        case 'blockquote_open': {
          const inl = tokens[i + 2];
          const m = tokens[i + 1] && tokens[i + 1].type === 'paragraph_open' && inl && inl.type === 'inline' ? CALLOUT_RE.exec(inl.content) : null;
          const co = m ? { type: m[1], fold: m[2] || '' } : null;
          this.quotes.push({ co, listDepth: this.listStack.length });
          if (co) this.pendingHead = co;
          i++;
          continue;
        }
        case 'blockquote_close':
          this.quotes.pop();
          i++;
          continue;
        case 'fence': case 'code_block': {
          const codeLines = t.content.replace(/\n$/, '').split('\n');
          const lang = (t.info || '').trim();
          const q = this.quoteMd();
          const first = t.map ? (t.type === 'fence' ? t.map[0] + 1 : t.map[0]) : null;
          codeLines.forEach((l, k) => {
            const a = { styleId: 'Code' };
            if (this.listStack.length) a.ind = { left: String(720 * this.listStack.length + 113) };
            if (k === 0) a.mdFence = lang;
            if (q) a.md = q;
            let info = null;
            if (first != null) {
              const ln = first + k;
              const srcLine = this.lines[ln];
              if (srcLine != null && srcLine.endsWith(l)) info = { kind: 'code', a: ln, b: ln + 1, prefix: srcLine.slice(0, srcLine.length - l.length) };
            }
            this.push(nodes.paragraph.create(a, lineContent(l)), info);
          });
          i++;
          continue;
        }
        case 'hr': {
          const a = { pBdr: { bottom: { val: 'single', sz: '6', space: '1', color: 'auto' } } };
          const q = this.quoteMd();
          if (q) a.md = q;
          this.push(nodes.paragraph.create(a));
          i++;
          continue;
        }
        case 'table_open': {
          const { node, next, rows } = this.table(tokens, i);
          this.push(node, t.map && rows.every((r) => r != null) ? { kind: 'table', a: t.map[0], b: t.map[1], rows } : null);
          // parágrafo vazio para poder escrever depois da tabela; não faz parte do arquivo
          if (t.level === 0) {
            this.begin('virtual', null);
            this.push(nodes.paragraph.create());
          }
          i = next;
          continue;
        }
        case 'html_block':
          for (const n of this.html(t.content)) this.push(n);
          i++;
          continue;
        default:
          i++;
      }
    }
  }

  // primeiro parágrafo de um callout: a linha do título vira um parágrafo próprio
  calloutParagraphs(t, inline) {
    const co = this.pendingHead;
    this.pendingHead = null;
    const list = (inline.children || []).slice();
    const cut = list.findIndex((c) => c.type === 'softbreak' || c.type === 'hardbreak');
    const head = cut >= 0 ? list.slice(0, cut) : list;
    const rest = cut >= 0 ? list.slice(cut + 1) : [];
    // tira o marcador [!tipo] do começo, mesmo que venha em mais de um pedaço de texto
    let k = 0;
    let lead = '';
    while (k < head.length && head[k].type === 'text') { lead += head[k].content; k++; }
    if (k > 0) {
      const first = cloneToken(head[0]);
      first.content = lead.replace(CALLOUT_RE, '');
      head.splice(0, k, first);
    }
    const q = this.quoteMd();
    const headInfo = t.map ? { kind: 'head', a: t.map[0], b: t.map[0] + 1 } : null;
    this.push(nodes.paragraph.create({ styleId: 'BlockQuote', md: { ...q, callout: co.type, fold: co.fold || null, head: true } }, this.inline(head)), headInfo);
    if (rest.length) {
      const parts = (inline.content || '').split('\n');
      const info = t.map ? { kind: 'lines', a: t.map[0] + 1, b: t.map[1], first: parts[1] || '', second: parts.length > 2 ? parts[2] : null } : null;
      this.push(nodes.paragraph.create({ styleId: 'BlockQuote', md: { ...q, callout: co.type } }, this.inline(rest)), info);
    }
  }

  html(content) {
    const trimmed = content.replace(/\n+$/, '');
    if (/^\s*<!--[\s\S]*-->\s*$/.test(trimmed)) return [rawBlock(trimmed)];
    try {
      const dom = new DOMParser().parseFromString(content, 'text/html');
      normalizeExternalHTML(dom.body, { codeStyle: 'Code', quoteStyle: 'BlockQuote' });
      for (const img of dom.body.querySelectorAll('img')) img.setAttribute('src', resolveRelative(this.baseDir, img.getAttribute('src')));
      const frag = PMDOMParser.fromSchema(schema).parseSlice(dom.body).content;
      const out = [];
      frag.forEach((n) => { if (n.isBlock) out.push(n); else out.push(nodes.paragraph.create(null, [n])); });
      let visible = false;
      for (const n of out) {
        if (n.textContent.trim() || n.type === nodes.table) visible = true;
        n.descendants((c) => { if (c.type === nodes.drawing) visible = true; return !visible; });
      }
      return visible ? out : [rawBlock(trimmed)];
    } catch {
      return [rawBlock(trimmed)];
    }
  }

  table(tokens, start) {
    const rows = [];
    const rowLines = [];
    let row = null;
    let i = start + 1;
    let header = false;
    while (i < tokens.length && tokens[i].type !== 'table_close') {
      const t = tokens[i];
      if (t.type === 'thead_open') header = true;
      else if (t.type === 'thead_close') header = false;
      else if (t.type === 'tr_open') { row = { header, cells: [] }; rowLines.push(t.map ? t.map[0] : null); }
      else if (t.type === 'tr_close') { rows.push(row); row = null; }
      else if (t.type === 'th_open' || t.type === 'td_open') {
        const style = t.attrGet('style') || '';
        const al = /text-align:\s*(left|center|right)/.exec(style);
        const inline = tokens[i + 1];
        row.cells.push({ align: al ? al[1] : null, content: this.inline(inline.children), header: t.type === 'th_open' });
        i += 3;
        continue;
      }
      i++;
    }
    const cols = Math.max(1, ...rows.map((r) => r.cells.length));
    const totalTw = Math.round(this.maxW * 15);
    const colTw = Math.floor(totalTw / cols);
    const tstyle = this.styles.get('GridTable1Light') ? 'GridTable1Light' : this.styles.get('TableGrid') ? 'TableGrid' : null;
    const pmRows = rows.map((r) => {
      const cells = [];
      for (let c = 0; c < cols; c++) {
        const cell = r.cells[c];
        const p = nodes.paragraph.create(cell && cell.align ? { jc: cell.align } : null, cell ? cell.content : []);
        cells.push(nodes.table_cell.create({ colwidth: [colTw / 15], tcW: { w: String(colTw), type: 'dxa' } }, [p]));
      }
      return nodes.table_row.create(r.header ? { header: true } : null, cells);
    });
    const node = nodes.table.create({
      styleId: tstyle, tblW: { w: '0', type: 'auto' }, grid: new Array(cols).fill(colTw),
      tblLook: { val: '04A0', firstRow: '1', lastRow: '0', firstColumn: '0', lastColumn: '0', noHBand: '1', noVBand: '1' },
    }, pmRows.length ? pmRows : [nodes.table_row.create(null, [nodes.table_cell.create(null, [nodes.paragraph.create()])])]);
    return { node, next: i + 1, rows: pmRows.length ? rowLines : [null] };
  }

  inline(children) {
    const out = [];
    if (!children) return out;
    let active = [];
    const add = (m) => { active = m.addToSet(active); };
    const remove = (type) => { active = type.removeFromSet(active); };
    const linkStyle = this.styles.get('Hyperlink') ? 'Hyperlink' : null;
    const text = (s) => { if (s) out.push(schema.text(s, active)); };
    for (const c of children) {
      switch (c.type) {
        case 'text': text(c.content); break;
        case 'softbreak': text(' '); break;
        case 'hardbreak': out.push(nodes.hard_break.create(null, null, active)); break;
        case 'strong_open': add(marks.b.create({ val: true })); break;
        case 'strong_close': remove(marks.b); break;
        case 'em_open': add(marks.i.create({ val: true })); break;
        case 'em_close': remove(marks.i); break;
        case 's_open': add(marks.strike.create({ val: true })); break;
        case 's_close': remove(marks.strike); break;
        case 'code_inline':
          if (c.content) out.push(schema.text(c.content, marks.rStyle.create({ val: 'InlineCode' }).addToSet(active)));
          break;
        case 'link_open': {
          const href = c.attrGet('href') || '';
          const title = c.attrGet('title');
          const a = {};
          if (c.markup === 'linkify' || c.markup === 'autolink') a.md = 'auto';
          if (title) a.title = title;
          const extra = Object.keys(a).length ? a : null;
          add(href.startsWith('#') ? marks.hyperlink.create({ anchor: href.slice(1), a: extra }) : marks.hyperlink.create({ href, a: extra }));
          if (linkStyle) add(marks.rStyle.create({ val: linkStyle }));
          break;
        }
        case 'link_close': remove(marks.hyperlink); remove(marks.rStyle); break;
        case 'wikilink': {
          const { target, alias, embed } = c.meta;
          if (embed && IMAGE_EXT.test(target)) {
            const src = this.embeds.get(target) || resolveRelative(this.baseDir, target);
            const size = this.sizes.get(src);
            const w = alias && /^\d+(x\d+)?$/.test(alias) ? Number(alias.split('x')[0]) : null;
            const { cx, cy } = size ? fitSize(size.w, size.h, Math.min(this.maxW, w || this.maxW)) : { cx: 0, cy: 0 };
            out.push(nodes.drawing.create({ kind: 'image', src, alt: target, cx, cy, ocx: cx, ocy: cy, name: 'md:' + c.content }, null, active));
            break;
          }
          const link = marks.hyperlink.create({ href: 'wiki:' + target, a: { wiki: embed ? 'embed' : 'link' } });
          let set = link.addToSet(active);
          if (linkStyle) set = marks.rStyle.create({ val: linkStyle }).addToSet(set);
          out.push(schema.text(alias || target, set));
          break;
        }
        case 'image': {
          const raw = c.attrGet('src') || '';
          const src = resolveRelative(this.baseDir, raw);
          const size = this.sizes.get(src);
          const { cx, cy } = size ? fitSize(size.w, size.h, this.maxW) : { cx: 0, cy: 0 };
          out.push(nodes.drawing.create({ kind: 'image', src, alt: c.content || '', title: c.attrGet('title') || '', cx, cy, ocx: cx, ocy: cy, name: 'md:' + raw }, null, active));
          break;
        }
        case 'html_inline': {
          const tag = c.content.toLowerCase().replace(/\s+/g, '');
          const name = (/^<\/?([a-z][a-z0-9-]*)/.exec(tag) || [])[1];
          if (/^<br\/?>$/.test(tag)) out.push(nodes.hard_break.create(null, null, active));
          else if (tag === '<u>' || tag === '<ins>') add(marks.u.create({ a: { val: 'single' } }));
          else if (tag === '</u>' || tag === '</ins>') remove(marks.u);
          else if (tag === '<sub>') add(marks.vertAlign.create({ val: 'subscript' }));
          else if (tag === '<sup>') add(marks.vertAlign.create({ val: 'superscript' }));
          else if (tag === '</sub>' || tag === '</sup>') remove(marks.vertAlign);
          else if (tag === '<mark>') add(marks.highlight.create({ val: 'yellow' }));
          else if (tag === '</mark>') remove(marks.highlight);
          else if (tag === '<b>' || tag === '<strong>') add(marks.b.create({ val: true }));
          else if (tag === '</b>' || tag === '</strong>') remove(marks.b);
          else if (tag === '<i>' || tag === '<em>') add(marks.i.create({ val: true }));
          else if (tag === '</i>' || tag === '</em>') remove(marks.i);
          else if (tag === '<s>' || tag === '<del>' || tag === '<strike>') add(marks.strike.create({ val: true }));
          else if (tag === '</s>' || tag === '</del>' || tag === '</strike>') remove(marks.strike);
          else if (tag === '<kbd>' || tag === '<code>') add(marks.rStyle.create({ val: 'InlineCode' }));
          else if (tag === '</kbd>' || tag === '</code>') remove(marks.rStyle);
          else if (name && HTML_TAGS.has(name)) { /* marcação sem efeito no documento */ }
          else text(c.content);
          break;
        }
        default:
          text(c.content);
      }
    }
    return out;
  }

  // Monta a lista final de blocos e os trechos originais correspondentes
  assemble(front, frontRaw) {
    const lines = this.lines;
    const body = this.body;
    const off = [];
    let acc = 0;
    for (let k = 0; k < lines.length; k++) { off.push(acc); acc += lines[k].length + 1; }
    const slice = (a, b) => (a >= lines.length ? '' : body.slice(off[a], b >= lines.length ? body.length : off[b]));
    const blocks = [];
    const chunks = [];
    const groups = this.groups;
    const out = this.out;
    const firstReal = groups.find((g) => g.line != null);
    const firstLine = firstReal ? firstReal.line : lines.length;
    let lead = '';
    if (front != null) {
      const fl = front === '' ? [''] : front.split('\n');
      const fn = fl.map((l, k) => nodes.paragraph.create({ styleId: FRONT_STYLE, mdFence: k === 0 ? '---' : null }, lineContent(l)));
      chunks.push({ kind: 'front', text: frontRaw + slice(0, firstLine), from: 0, count: fn.length, src: null });
      blocks.push(...fn);
    } else {
      lead = slice(0, firstLine);
    }
    for (let k = 0; k < groups.length; k++) {
      const g = groups[k];
      const to = k + 1 < groups.length ? groups[k + 1].from : out.length;
      const own = out.slice(g.from, to);
      const ownSrc = this.src.slice(g.from, to);
      if (g.line == null) {
        if (own.length) { chunks.push({ kind: 'virtual', text: '', from: blocks.length, count: own.length, src: null }); blocks.push(...own); }
        continue;
      }
      let next = lines.length;
      for (let q = k + 1; q < groups.length; q++) if (groups[q].line != null) { next = groups[q].line; break; }
      // linhas com conteúdo que nenhum bloco cobre, como definições de links
      let gap = -1;
      for (let ln = Math.max(g.end, g.line); ln < next; ln++) if (lines[ln] && lines[ln].trim()) { gap = ln; break; }
      const leadLines = lead ? lead.split('\n').length - 1 : 0;
      const text = lead + slice(g.line, gap >= 0 ? gap : next);
      const leadLen = lead.length;
      lead = '';
      if (!own.length) {
        if (!text.trim()) { lead = text; continue; }
        own.push(rawBlock(text));
        ownSrc.push(null);
      }
      chunks.push({ kind: g.kind, text, from: blocks.length, count: own.length, src: ownSrc, base: g.line - leadLines, leadLen });
      blocks.push(...own);
      if (gap >= 0) {
        const rawText = slice(gap, next);
        chunks.push({ kind: 'raw', text: rawText, from: blocks.length, count: 1, src: null });
        blocks.push(rawBlock(rawText));
      }
    }
    if (lead) {
      if (chunks.length) chunks[chunks.length - 1].text += lead;
      else if (lead.trim()) { chunks.push({ kind: 'raw', text: lead, from: 0, count: 1, src: null }); blocks.push(rawBlock(lead)); }
    }
    return { blocks, chunks };
  }
}

// ---------------------------------------------------------------- escrita

const CODE_FONTS = /consolas|courier|mono|menlo|monaco|lucida console/i;
const PUNCT = /[!-/:-@[-`{-~]/;

function isCodeStyle(styles, id) {
  if (!id) return false;
  const st = styles.get(id);
  const name = (st && st.name) || id;
  if (/^(code|html preformatted|source code|plain text|código)$/i.test(name)) return true;
  const r = styles.resolvePara(id).r;
  return !!(r.rFonts && r.rFonts.a && CODE_FONTS.test(r.rFonts.a.ascii || ''));
}

function isQuoteStyle(styles, id) {
  if (!id) return false;
  const st = styles.get(id);
  return /quote|citação|block text/i.test((st && st.name) || id);
}

// escapa só o que mudaria o sentido do texto
function escapeInline(s, ctx = {}) {
  const bracketRisk = ctx.link || /\[[^\]]*\]\s*[([:]/.test(s) || s.includes('[[') || s.includes('![');
  let out = '';
  for (let i = 0; i < s.length; i++) {
    const c = s[i];
    const prev = i > 0 ? s[i - 1] : ' ';
    const next = i + 1 < s.length ? s[i + 1] : ' ';
    switch (c) {
      case '\\': out += PUNCT.test(next) ? '\\\\' : '\\'; break;
      case '*': out += /\s/.test(prev) && /\s/.test(next) ? '*' : '\\*'; break;
      case '_': {
        const word = /[\p{L}\p{N}]/u;
        out += (word.test(prev) && word.test(next)) || (/\s/.test(prev) && /\s/.test(next)) ? '_' : '\\_';
        break;
      }
      case '`': out += '\\`'; break;
      case '[': case ']': out += bracketRisk ? '\\' + c : c; break;
      case '<': out += /[A-Za-z/!?]/.test(next) ? '\\<' : '<'; break;
      case '|': out += ctx.table ? '\\|' : '|'; break;
      case '~': out += next === '~' || prev === '~' ? '\\~' : '~'; break;
      case '&': out += /^&(#\d+|#x[\da-f]+|[a-z][a-z\d]*);/i.test(s.slice(i, i + 12)) ? '\\&' : '&'; break;
      default: out += c;
    }
  }
  return out;
}

function escapeLineStart(s) {
  return s
    .replace(/^(\s*)(#{1,6})(\s|$)/, '$1\\$2$3')
    .replace(/^(\s*)([>+*-])(\s|$)/, '$1\\$2$3')
    .replace(/^(\s*)(\d+)([.)])(\s|$)/, '$1$2\\$3$4')
    .replace(/^(\s*)(={3,}|-{3,})\s*$/, '$1\\$2');
}

function taskText(t) {
  if (/^☑ /.test(t)) return '[x] ' + t.slice(2);
  if (/^☐ /.test(t)) return '[ ] ' + t.slice(2);
  return null;
}

export function writeMarkdown(doc, model, opts = {}) {
  const prev = R.model;
  R.model = model;
  try {
    const w = new MdWriter(model, opts);
    const src = model.extra.md;
    let text = src && src.bound && !opts.fresh ? w.preserving(doc, src) : w.fresh(doc);
    const eol = (model.extra.text && model.extra.text.eol) || '\n';
    if (eol !== '\n') text = text.replace(/\n/g, eol);
    return { text, files: w.files };
  } finally {
    R.model = prev;
  }
}

function sepBetween(a, b) {
  if (!a || !b) return '\n\n';
  const listy = (p) => p.kind === 'item' || p.kind === 'itemcode';
  if (b.kind === 'item' && listy(a) && (b.ilvl > 0 || a.numId === b.numId || a.kind === 'itemcode')) return '\n';
  if (b.kind === 'itemcode' && (listy(a) || a.kind === 'itemcont')) return '\n';
  return '\n\n';
}

function joinPieces(pieces) {
  let s = '';
  for (let k = 0; k < pieces.length; k++) s += (k ? sepBetween(pieces[k - 1], pieces[k]) : '') + pieces[k].text;
  return s;
}

function inQuoteNode(styles, n) {
  if (n.type !== nodes.paragraph) return false;
  const md = n.attrs.md;
  if (md && (md.quote || md.callout)) return true;
  return !effectiveNumPr(n.attrs) && isQuoteStyle(styles, n.attrs.styleId);
}

class MdWriter {
  constructor(model, opts) {
    this.model = model;
    this.styles = model.styles;
    this.opts = opts;
    this.files = [];
    this.imgCount = 0;
    this.counter = new NumberingCounter(model.numbering);
    this.lists = (model.extra.md && model.extra.md.lists) || new Map();
    this.cols = [];
  }

  fresh(doc) {
    const list = [];
    doc.forEach((n) => list.push(n));
    const body = joinPieces(this.pieces(list));
    return body.replace(/^\n+/, '').replace(/\s+$/, '') + '\n';
  }

  // Alinha os nós atuais com os originais: iguais, editados no lugar ou novos
  align(cur, orig) {
    const at = new Map();
    orig.forEach((o, idx) => at.set(o.node, idx));
    const fate = new Array(cur.length).fill(null);
    const WINDOW = 300;
    let j = 0;
    for (let i = 0; i < cur.length; i++) {
      const n = cur[i];
      const hit = at.get(n);
      if (hit != null && hit >= j) { fate[i] = { o: hit, mod: false }; j = hit + 1; continue; }
      if (j < orig.length && orig[j].node.eq(n)) { fate[i] = { o: j, mod: false }; j++; continue; }
      let found = -1;
      for (let q = j + 1; q < Math.min(orig.length, j + WINDOW); q++) {
        const o = orig[q].node;
        if (o === n || (o.type === n.type && o.content.size === n.content.size && o.eq(n))) { found = q; break; }
      }
      if (found >= 0) { fate[i] = { o: found, mod: false }; j = found + 1; continue; }
      if (j < orig.length) {
        const o = orig[j].node;
        let later = false;
        for (let q = i + 1; q < Math.min(cur.length, i + WINDOW); q++) if (cur[q] === o) { later = true; break; }
        if (!later && o.sameMarkup(n)) { fate[i] = { o: j, mod: true }; j++; continue; }
      }
    }
    return fate;
  }

  preserving(doc, src) {
    const chunks = src.chunks;
    const orig = [];
    chunks.forEach((c, ci) => c.nodes.forEach((node, k) => orig.push({ node, ci, k })));
    const cur = [];
    doc.forEach((n) => cur.push(n));
    const fate = this.align(cur, orig);
    const parts = [];
    let tail = '';
    const emit = (s) => { if (!s) return; parts.push(s); tail = (tail + s).slice(-2); };
    const ensureSep = (sep) => {
      if (!parts.length) return;
      if (sep === '\n') { if (!tail.endsWith('\n')) emit('\n'); return; }
      if (tail.endsWith('\n\n')) return;
      emit(tail.endsWith('\n') ? '\n' : '\n\n');
    };
    let last = null;
    let run = [];
    const chunkPiece = (c) => ({ kind: c.kind === 'item' ? 'item' : 'block', numId: c.numId, ilvl: 0 });
    const lastPiece = () => (last ? (last.chunk ? chunkPiece(last.chunk) : last.piece) : null);
    const flush = () => {
      if (!run.length) return;
      const pieces = this.pieces(run);
      run = [];
      if (!pieces.length) return;
      ensureSep(sepBetween(lastPiece(), pieces[0]));
      emit(joinPieces(pieces));
      last = { piece: pieces[pieces.length - 1] };
    };
    const emitChunk = (ci, text) => {
      const c = chunks[ci];
      if (last && !(last.index === ci - 1)) ensureSep(sepBetween(lastPiece(), chunkPiece(c)));
      emit(text);
      this.observe(c);
      last = { index: ci, chunk: c };
    };
    let i = 0;
    while (i < cur.length) {
      const f = fate[i];
      if (!f) { run.push(cur[i]); i++; continue; }
      const o = orig[f.o];
      const c = chunks[o.ci];
      let e = i;
      let k = o.k;
      let modified = f.mod;
      while (e + 1 < cur.length && fate[e + 1] && orig[fate[e + 1].o].ci === o.ci && orig[fate[e + 1].o].k === k + 1) {
        e++;
        k++;
        if (fate[e].mod) modified = true;
      }
      const whole = o.k === 0 && k === c.count - 1;
      const slice = cur.slice(i, e + 1);
      if (whole && !modified) { flush(); emitChunk(o.ci, c.text); }
      else {
        const patched = whole ? this.patchChunk(c, slice) : null;
        if (patched != null) { flush(); emitChunk(o.ci, patched); }
        else run.push(...slice);
      }
      i = e + 1;
    }
    flush();
    let text = parts.join('');
    if (!(last && last.index === chunks.length - 1)) text = text.replace(/\s+$/, '') + '\n';
    if (!text.trim()) text = '';
    return text;
  }

  // troca só as linhas dos parágrafos editados, mantendo recuos, marcadores e citações
  patchChunk(c, cur) {
    if (!c.src || c.base == null) return null;
    const lines = c.text.split('\n');
    const edits = [];
    for (let k = 0; k < cur.length; k++) {
      const o = c.nodes[k];
      const n = cur[k];
      if (o === n || o.eq(n)) continue;
      const info = c.src[k];
      if (!info || !o.sameMarkup(n)) return null;
      const a = info.a - c.base;
      const b = info.b - c.base;
      if (a < 0 || b > lines.length || a >= b) return null;
      if (info.kind === 'code') {
        edits.push({ from: a, to: b, lines: [info.prefix + n.textContent] });
        continue;
      }
      if (info.kind === 'table') {
        // só as linhas da tabela que mudaram, mantendo o alinhamento das outras
        if (o.childCount !== n.childCount || info.rows.length !== n.childCount) return null;
        for (let r = 0; r < n.childCount; r++) {
          const orow = o.child(r);
          const nrow = n.child(r);
          if (orow === nrow || orow.eq(nrow)) continue;
          if (!orow.sameMarkup(nrow) || orow.childCount !== nrow.childCount) return null;
          for (let q = 0; q < nrow.childCount; q++) if (!orow.child(q).sameMarkup(nrow.child(q))) return null;
          const ln = info.rows[r] - c.base;
          if (ln < 0 || ln >= lines.length) return null;
          const prefix = /^[ \t]*(?:>[ \t]?)*/.exec(lines[ln])[0];
          edits.push({ from: ln, to: ln + 1, lines: [prefix + '| ' + this.rowCells(nrow).join(' | ') + ' |'] });
        }
        continue;
      }
      const src0 = lines[a];
      let t = this.inline(n).trim();
      const task = taskText(t);
      if (task != null) t = task;
      else if (!info.heading && info.kind !== 'head') t = escapeLineStart(t);
      let prefix;
      if (info.kind === 'head') {
        const m = /\[!\p{L}[\p{L}\p{N}_-]*\][+-]?[ \t]*/u.exec(src0);
        if (!m) return null;
        prefix = src0.slice(0, m.index + m[0].length);
        if (t && !/[ \t]$/.test(prefix)) prefix += ' ';
      } else if (!info.first) {
        prefix = src0.replace(/\s+$/, '') + ' ';
      } else {
        const idx = src0.indexOf(info.first);
        if (idx < 0) return null;
        prefix = src0.slice(0, idx);
      }
      let cont = null;
      if (info.second != null && a + 1 < b) {
        const s1 = lines[a + 1];
        const idx2 = info.second ? s1.indexOf(info.second) : -1;
        if (idx2 >= 0) cont = s1.slice(0, idx2);
      }
      if (cont == null) cont = prefix.replace(/([-*+]|\d{1,9}[.)])([ \t]*)$/, (m) => ' '.repeat(m.length)).replace(/\[![^\]]*\][+-]?[ \t]*$/, '');
      const out = t.split('\n').map((l, q) => (q === 0 ? prefix : cont) + l);
      edits.push({ from: a, to: b, lines: out });
    }
    edits.sort((x, y) => y.from - x.from);
    for (const e of edits) lines.splice(e.from, e.to - e.from, ...e.lines);
    return lines.join('\n');
  }

  // acompanha a numeração e a indentação dos trechos mantidos
  observe(chunk) {
    for (const n of chunk.nodes) {
      if (n.type !== nodes.paragraph) continue;
      const np = effectiveNumPr(n.attrs);
      if (np && n.content.size) this.counter.next(np.numId, np.ilvl);
    }
    if (chunk.kind === 'item') { this.cols[0] = chunk.col; this.cols.length = 1; }
  }

  pieces(list, ctx = {}) {
    const out = [];
    for (let i = 0; i < list.length; i++) {
      const n = list[i];
      if (n.type === nodes.paragraph) {
        const a = n.attrs;
        if (a.styleId === FRONT_STYLE) {
          const lines = [];
          while (i < list.length && list[i].type === nodes.paragraph && list[i].attrs.styleId === FRONT_STYLE && (lines.length === 0 || list[i].attrs.mdFence == null)) {
            lines.push(list[i].textContent);
            i++;
          }
          i--;
          out.push({ kind: 'front', text: '---\n' + (lines.length === 1 && lines[0] === '' ? '' : lines.join('\n') + '\n') + '---' });
          continue;
        }
        if (!ctx.inQuote && inQuoteNode(this.styles, n)) {
          const qlist = (a.md && a.md.qlist) || 0;
          const group = [n];
          while (i + 1 < list.length && inQuoteNode(this.styles, list[i + 1]) && ((list[i + 1].attrs.md && list[i + 1].attrs.md.qlist) || 0) === qlist
            && !(list[i + 1].attrs.md && list[i + 1].attrs.md.head)) group.push(list[++i]);
          const q = this.quote(group, qlist);
          if (q) out.push(q);
          continue;
        }
        if (isCodeStyle(this.styles, a.styleId)) {
          const lines = [];
          const lang = a.mdFence || '';
          const sid = a.styleId;
          while (i < list.length && list[i].type === nodes.paragraph && list[i].attrs.styleId === sid && (lines.length === 0 || list[i].attrs.mdFence == null)) {
            lines.push(list[i].textContent);
            i++;
          }
          i--;
          out.push(this.codeBlock(lines, lang, a, ctx));
          continue;
        }
        const piece = this.paragraph(n, ctx);
        if (piece) out.push(piece);
      } else if (n.type === nodes.table) {
        out.push({ kind: 'table', text: this.table(n) });
      } else if (n.type === nodes.sdt_block) {
        const inner = [];
        n.forEach((c) => inner.push(c));
        out.push(...this.pieces(inner, ctx));
      } else if (n.type === nodes.opaque_block) {
        if (n.attrs.kind === 'md') out.push({ kind: 'raw', text: n.attrs.text || '' });
        else if (n.attrs.text) out.push({ kind: 'para', text: escapeInline(n.attrs.text) });
      }
    }
    return out;
  }

  // profundidade de lista de um parágrafo sem numeração, pelo recuo
  depthOf(a, ctx) {
    const left = Number((a.ind && a.ind.left) || 0);
    const d = left > 0 ? Math.max(0, Math.round((left - 113) / 720)) : 0;
    return Math.max(0, d - ((ctx && ctx.qlist) || 0));
  }

  indentFor(depth) {
    let s = 0;
    for (let d = 0; d < depth; d++) s = this.cols[d] != null ? this.cols[d] : s + 3;
    return ' '.repeat(s);
  }

  codeBlock(lines, lang, a, ctx) {
    const depth = a.styleId === 'Code' || /list/i.test(a.styleId || '') ? this.depthOf(a, ctx) : 0;
    const ind = depth ? this.indentFor(depth) : '';
    const fence = lines.some((l) => l.includes('```')) ? '~~~' : '```';
    const body = lines.map((l) => (l ? ind + l : l)).join('\n');
    return { kind: depth ? 'itemcode' : 'code', text: ind + fence + lang + '\n' + (lines.length ? body + '\n' : '') + ind + fence };
  }

  // citação ou callout, com o que houver dentro: parágrafos, listas, código
  quote(group, qlist) {
    const lines = [];
    let body = group;
    const first = group[0];
    if (first.attrs.md && first.attrs.md.head) {
      const md = first.attrs.md;
      const t = this.inline(first).trim();
      lines.push('[!' + md.callout + ']' + (md.fold || '') + (t ? ' ' + t : ''));
      body = group.slice(1);
    }
    if (body.length) {
      const saved = this.cols;
      this.cols = [];
      const inner = joinPieces(this.pieces(body, { inQuote: true, qlist }));
      this.cols = saved;
      if (inner) lines.push(...inner.split('\n'));
    }
    if (!lines.length) return null;
    const ind = qlist ? this.indentFor(qlist) : '';
    return { kind: qlist ? 'itemquote' : 'quote', text: lines.map((l) => ind + (l ? '> ' + l : '>')).join('\n') };
  }

  paragraph(n, ctx = {}) {
    const a = n.attrs;
    const level = this.styles.headingLevel(a.styleId || this.styles.defaultPara);
    const st = this.styles.get(a.styleId);
    const name = ((st && st.name) || '').toLowerCase();
    if (!n.content.size) {
      if (a.pBdr && a.pBdr.bottom && a.pBdr.bottom.val !== 'nil' && a.pBdr.bottom.val !== 'none') return { kind: 'hr', text: '---' };
      return null;
    }
    const text = this.inline(n);
    if (level >= 1 && level <= 6) return { kind: 'heading', text: '#'.repeat(level) + ' ' + text.replace(/\\\n/g, ' ').trim() };
    if (name === 'title') return { kind: 'heading', text: '# ' + text.replace(/\\\n/g, ' ').trim() };
    if (name === 'subtitle') return { kind: 'para', text: '*' + text.trim() + '*' };
    const np = effectiveNumPr(a);
    if (np) {
      const lbl = this.counter.next(np.numId, np.ilvl);
      const kind = this.model.numbering.kindOf(np.numId, np.ilvl);
      const mk = this.lists.get(String(np.numId));
      const marker = kind === 'bullet' ? (mk && /^[-*+]$/.test(mk) ? mk : '-') : ((lbl && lbl.value) || 1) + (mk === ')' ? ')' : '.');
      const ilvl = Math.max(0, np.ilvl - ((ctx && ctx.qlist) || 0));
      const indent = this.indentFor(ilvl);
      this.cols[ilvl] = indent.length + marker.length + 1;
      this.cols.length = ilvl + 1;
      let t = text.trim();
      const task = taskText(t);
      t = task != null ? task : escapeLineStart(t);
      const cont = ' '.repeat(this.cols[ilvl]);
      return { kind: 'item', numId: String(np.numId), ilvl, text: indent + marker + ' ' + t.replace(/\n/g, '\n' + cont) };
    }
    const depth = a.styleId === 'ListParagraph' || /list/i.test(a.styleId || '') ? this.depthOf(a, ctx) : 0;
    if (depth) {
      const ind = this.indentFor(depth);
      return { kind: 'itemcont', text: ind + escapeLineStart(text.trim()).replace(/\n/g, '\n' + ind) };
    }
    const out = escapeLineStart(text.replace(/^\s+|\s+$/g, ''));
    return out ? { kind: 'para', text: out } : null;
  }

  mdMarks(node, parentAttrs) {
    const set = [];
    const r = this.styles.effectiveRun(parentAttrs, node.marks);
    const direct = (nm) => node.marks.find((m) => m.type.name === nm);
    const code = node.marks.some((m) => m.type.name === 'rStyle' && /code/i.test(m.attrs.val || '')) || (direct('rFonts') && CODE_FONTS.test(direct('rFonts').attrs.a.ascii || ''));
    const base = this.styles.resolvePara(parentAttrs.styleId || this.styles.defaultPara).r;
    if (r.b && r.b.val && !(base.b || {}).val) set.push('b');
    if (r.i && r.i.val && !(base.i || {}).val) set.push('i');
    if ((r.strike && r.strike.val) || (r.dstrike && r.dstrike.val)) set.push('s');
    if (direct('u') && direct('u').attrs.a && direct('u').attrs.a.val !== 'none') set.push('u');
    if (r.vertAlign && r.vertAlign.val === 'superscript') set.push('sup');
    if (r.vertAlign && r.vertAlign.val === 'subscript') set.push('sub');
    if (direct('highlight') && direct('highlight').attrs.val !== 'none') set.push('mark');
    // o código fica por dentro das outras marcas
    if (code) set.push('code');
    const link = direct('hyperlink');
    return { set, link: link ? link.attrs : null };
  }

  inline(para, ctx = {}) {
    const open = { b: '**', i: '*', s: '~~', u: '<u>', sup: '<sup>', sub: '<sub>', mark: '<mark>', code: '`' };
    const close = { b: '**', i: '*', s: '~~', u: '</u>', sup: '</sup>', sub: '</sub>', mark: '</mark>', code: '`' };
    let out = '';
    const stack = [];
    let curLink = null;
    let linkText = '';
    let linkPlain = '';
    const closeAll = (keep) => {
      let s = '';
      while (stack.length > keep) s += close[stack.pop()];
      return s;
    };
    const emit = (s) => { if (curLink) linkText += s; else out += s; };
    const finishLink = () => {
      linkText += closeAll(0);
      const a = curLink;
      const extra = a.a || {};
      if (extra.wiki) {
        const target = a.href.replace(/^wiki:/, '');
        const shown = linkPlain;
        const pipe = ctx.table ? '\\|' : '|';
        out += (extra.wiki === 'embed' ? '!' : '') + '[[' + target + (shown && shown !== target ? pipe + shown : '') + ']]';
      } else if (extra.md === 'auto' && a.href && linkPlain && (a.href === linkPlain || a.href.endsWith('//' + linkPlain) || a.href === 'mailto:' + linkPlain)) {
        // endereço solto no texto: continua solto
        out += linkPlain;
      } else {
        let href = a.href || '#' + (a.anchor || '');
        if (/[\s()<>]/.test(href)) href = '<' + href.replace(/[<>]/g, (ch) => encodeURIComponent(ch)) + '>';
        const title = extra.title ? ' "' + String(extra.title).replace(/"/g, '\\"') + '"' : '';
        out += '[' + linkText + '](' + href + title + ')';
      }
      curLink = null;
      linkText = '';
      linkPlain = '';
    };
    const pieces = [];
    para.forEach((child) => pieces.push(child));
    for (const child of pieces) {
      const { set, link } = child.isText || child.type === nodes.drawing ? this.mdMarks(child, para.attrs) : { set: [], link: null };
      const sameLink = link && curLink && link.href === curLink.href && link.anchor === curLink.anchor;
      if (!sameLink) {
        if (curLink) finishLink();
        if (link) { out += closeAll(0); curLink = link; }
      }
      if (curLink && curLink.a && curLink.a.wiki) {
        if (child.isText) linkPlain += child.text;
        continue;
      }
      if (child.isText && !child.text.trim() && !set.includes('code')) {
        emit(closeAll(0) + child.text);
        if (curLink) linkPlain += child.text;
        continue;
      }
      let keep = 0;
      while (keep < stack.length && set.includes(stack[keep])) keep++;
      const closing = closeAll(keep);
      let opening = '';
      for (const m of set) if (!stack.includes(m)) { stack.push(m); opening += open[m]; }
      if (child.isText) {
        const code = set.includes('code');
        if (curLink) linkPlain += child.text;
        let t = code ? child.text : escapeInline(child.text, { table: ctx.table, link: !!curLink });
        if (code && t.includes('`')) {
          // crase dentro do código pede delimitador maior
          stack.pop();
          opening = opening.slice(0, -1);
          emit(closing + opening + '`` ' + t + ' ``');
          continue;
        }
        const lead = code ? '' : /^\s*/.exec(t)[0];
        const trail = code ? '' : /\s*$/.exec(t)[0];
        t = t.slice(lead.length, t.length - trail.length);
        emit(closing + lead + opening + t);
        if (trail) emit(closeAll(0) + trail);
      } else {
        emit(closing + opening);
        if (child.type === nodes.hard_break) emit(child.attrs.type === 'page' ? '' : ctx.table ? '<br>' : '\\\n');
        else if (child.type === nodes.tab) emit(set.includes('code') ? '\t' : ' ');
        else if (child.type === nodes.drawing) emit(this.image(child));
        else if (child.type === nodes.sym) emit(child.textContent);
        else if (child.type === nodes.opaque_inline && child.attrs.kind === 'math') emit('$' + child.attrs.text + '$');
      }
    }
    if (curLink) finishLink();
    else out += closeAll(0);
    return out;
  }

  image(node) {
    const a = node.attrs;
    const name = a.name || '';
    if (name.startsWith('md:')) {
      const raw = name.slice(3);
      if (raw.startsWith('![[') || raw.startsWith('[[')) return raw.startsWith('!') ? raw : '!' + raw;
      const alt = (a.alt || '').replace(/[[\]]/g, '');
      const title = a.title ? ' "' + String(a.title).replace(/"/g, '\\"') + '"' : '';
      return '![' + alt + '](' + (/[\s()]/.test(raw) ? '<' + raw + '>' : raw) + title + ')';
    }
    const alt = (a.alt || '').replace(/[[\]]/g, '');
    if (a.src && /^file:/.test(a.src) && this.opts.baseDir) {
      const target = decodeURI(a.src.replace(/^file:\/\/\//, '')).replace(/\//g, '\\');
      const base = this.opts.baseDir.replace(/\//g, '\\');
      if (target.toLowerCase().startsWith(base.toLowerCase() + '\\')) return '![' + alt + '](' + target.slice(base.length + 1).replace(/\\/g, '/').replace(/ /g, '%20') + ')';
      return '![' + alt + '](' + a.src + ')';
    }
    if (a.src && /^https?:/.test(a.src)) return '![' + alt + '](' + a.src + ')';
    const m = a.mediaPath ? this.model.media.get(a.mediaPath) : null;
    let bytes = m && m.bytes;
    let ext = (m && m.ext) || 'png';
    if (!bytes && a.src && a.src.startsWith('data:')) {
      const mm = /^data:image\/(\w+);base64,(.*)$/.exec(a.src);
      if (mm) { ext = mm[1] === 'jpeg' ? 'jpg' : mm[1]; const bin = atob(mm[2]); bytes = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i); }
    }
    if (!bytes) return alt ? '[' + alt + ']' : '';
    this.imgCount++;
    const folder = (this.opts.baseName || 'documento') + '_arquivos';
    const file = folder + '/imagem' + this.imgCount + '.' + ext;
    this.files.push({ path: file, data: bytes });
    return '![' + alt + '](' + encodeURI(file) + ')';
  }

  rowCells(row) {
    const cells = [];
    row.forEach((cell) => {
      const parts = [];
      cell.forEach((b) => { if (b.type === nodes.paragraph) parts.push(this.inline(b, { table: true })); else parts.push(escapeInline(b.textContent, { table: true })); });
      cells.push(parts.filter((p) => p).join('<br>').replace(/\n/g, ' '));
      for (let k = 1; k < (cell.attrs.colspan || 1); k++) cells.push('');
    });
    return cells;
  }

  table(t) {
    const rows = [];
    t.forEach((row) => rows.push(this.rowCells(row)));
    if (!rows.length) return '';
    const cols = Math.max(...rows.map((r) => r.length));
    const aligns = [];
    const first = t.firstChild;
    if (first) {
      first.forEach((cell) => {
        const p = cell.firstChild;
        const jc = p && p.attrs.jc;
        aligns.push(jc === 'center' ? ':---:' : jc === 'right' || jc === 'end' ? '---:' : jc === 'left' || jc === 'start' ? ':---' : '---');
      });
    }
    while (aligns.length < cols) aligns.push('---');
    const line = (r) => '| ' + Array.from({ length: cols }, (_, i) => r[i] || '').join(' | ') + ' |';
    return [line(rows[0]), '| ' + aligns.slice(0, cols).join(' | ') + ' |', ...rows.slice(1).map(line)].join('\n');
  }
}
