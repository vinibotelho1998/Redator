// EPUB: livro digital aberto como documento. Os capítulos entram na ordem de
// leitura do livro, cada um começando numa página nova, com as imagens de
// dentro do arquivo. Salvar grava um .docx.
import JSZip from 'jszip';
import { schema } from '../editor/schema.js';
import { blankModel, finalizeDoc } from './common.js';
import { htmlToBlocks } from './html.js';

const nodes = schema.nodes;

const IMAGE_TYPES = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', svg: 'image/svg+xml', webp: 'image/webp', bmp: 'image/bmp' };

// junta um caminho relativo ao do arquivo que o cita, dentro do pacote
function joinPath(from, rel) {
  if (!rel) return null;
  if (/^[a-z]+:/i.test(rel)) return null;
  let target = rel.split('#')[0];
  try { target = decodeURIComponent(target); } catch { /* nome com % solto */ }
  if (!target) return null;
  const parts = (from.includes('/') ? from.slice(0, from.lastIndexOf('/') + 1) : '').split('/').filter(Boolean);
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg && seg !== '.') parts.push(seg);
  }
  return parts.join('/');
}

function findFile(zip, path) {
  if (!path) return null;
  return zip.file(path) || zip.file(Object.keys(zip.files).find((k) => k.toLowerCase() === path.toLowerCase()) || '');
}

export async function readEpub(bytes) {
  let zip;
  try { zip = await JSZip.loadAsync(bytes); } catch { throw new Error('O arquivo não parece ser um EPUB válido.'); }
  const container = findFile(zip, 'META-INF/container.xml');
  if (!container) throw new Error('O arquivo não parece ser um EPUB válido.');
  const opfPath = (/full-path\s*=\s*"([^"]+)"/.exec(await container.async('string')) || [])[1];
  const opfFile = findFile(zip, opfPath);
  if (!opfFile) throw new Error('O EPUB está incompleto: falta o índice do livro.');
  // conteúdo cifrado, fora a ofuscação padrão das fontes, é proteção contra cópia
  const enc = findFile(zip, 'META-INF/encryption.xml');
  if (enc) {
    const x = await enc.async('string');
    const locked = [...x.matchAll(/<(?:\w+:)?EncryptedData[\s\S]*?<\/(?:\w+:)?EncryptedData>/g)].some((m) => !/idpf\.org\/2008\/embedding|ns\.adobe\.com\/pdf\/enc#RC/.test(m[0]) && /\.(x?html?|xml)"/i.test(m[0]));
    if (locked) throw new Error('Este EPUB é protegido contra cópia e não pode ser aberto.');
  }
  const opf = new DOMParser().parseFromString(await opfFile.async('string'), 'application/xml');
  const byTag = (name) => Array.from(opf.getElementsByTagNameNS('*', name));
  const manifest = new Map();
  for (const it of byTag('item')) manifest.set(it.getAttribute('id'), { path: joinPath(opfPath, it.getAttribute('href')), type: it.getAttribute('media-type') || '', props: it.getAttribute('properties') || '' });
  let spine = byTag('itemref').filter((r) => r.getAttribute('linear') !== 'no').map((r) => manifest.get(r.getAttribute('idref'))).filter((x) => x && /html|xml/.test(x.type));
  if (!spine.length) spine = [...manifest.values()].filter((x) => /html/.test(x.type));
  const { model } = await blankModel();
  model.format = 'epub';
  const title = byTag('title')[0];
  const creator = byTag('creator')[0];
  model.properties.title = title ? title.textContent.trim() : '';
  if (creator) model.properties.creator = creator.textContent.trim();
  // imagens do livro viram endereços locais da janela
  const urls = new Map();
  const urlFor = async (path) => {
    if (!path) return null;
    if (urls.has(path)) return urls.get(path);
    const f = findFile(zip, path);
    if (!f) return null;
    const ext = (path.split('.').pop() || '').toLowerCase();
    const type = IMAGE_TYPES[ext];
    if (!type) return null;
    const url = URL.createObjectURL(new Blob([await f.async('uint8array')], { type }));
    urls.set(path, url);
    return url;
  };
  const blocks = [];
  for (const item of spine) {
    const f = findFile(zip, item.path);
    if (!f) continue;
    const html = await f.async('string');
    let dom = new DOMParser().parseFromString(html, 'application/xhtml+xml');
    if (!dom.body || dom.getElementsByTagName('parsererror').length) dom = new DOMParser().parseFromString(html, 'text/html');
    const body = dom.body;
    if (!body) continue;
    // capa e ilustrações em SVG com uma imagem dentro viram imagem comum
    for (const svg of Array.from(body.getElementsByTagName('svg'))) {
      const im = svg.getElementsByTagName('image')[0];
      const href = im && (im.getAttribute('href') || im.getAttributeNS('http://www.w3.org/1999/xlink', 'href'));
      if (!href) continue;
      const img = dom.createElement('img');
      img.setAttribute('src', href);
      img.setAttribute('alt', svg.getAttribute('aria-label') || '');
      svg.replaceWith(img);
    }
    const resolved = new Map();
    for (const img of Array.from(body.getElementsByTagName('img'))) {
      const src = img.getAttribute('src') || '';
      if (!resolved.has(src)) resolved.set(src, await urlFor(joinPath(item.path, src)));
    }
    // links para outras partes do livro ficam só como texto
    for (const a of Array.from(body.getElementsByTagName('a'))) {
      if (!/^(https?:|mailto:)/i.test(a.getAttribute('href') || '')) a.removeAttribute('href');
    }
    // XHTML guarda os nomes das tags em minúsculas; o resto da conversão espera HTML comum
    const plain = body.ownerDocument.contentType === 'text/html' ? body : new DOMParser().parseFromString('<!DOCTYPE html><html><body>' + body.innerHTML + '</body></html>', 'text/html').body;
    const part = await htmlToBlocks(plain, model, { resolve: (src) => resolved.get(src) || null });
    const content = part.filter((n) => n.type !== nodes.paragraph || n.content.size || n.attrs.pageBreakBefore);
    if (!content.length) continue;
    // cada capítulo começa numa página nova
    if (blocks.length) {
      if (content[0].type === nodes.paragraph) content[0] = content[0].type.create({ ...content[0].attrs, pageBreakBefore: true }, content[0].content, content[0].marks);
      else blocks.push(nodes.paragraph.create({ pageBreakBefore: true }));
    }
    blocks.push(...content);
  }
  return { model, doc: finalizeDoc(schema.node('doc', null, blocks.length ? blocks : [nodes.paragraph.create()]), model) };
}
