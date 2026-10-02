// Perdas conhecidas de cada escritor. A lista é avaliada antes de substituir
// qualquer arquivo, inclusive quando a extensão de destino é a da abertura.
import { canPreserveOdtSource } from './odt.js';

export function lossesForFormat(model, doc, format, opts = {}) {
  const fmt = format === 'ott' ? 'odt' : format;
  const odtOptions = format === 'ott' ? { ...opts, template: true } : opts;
  if (fmt === 'odt' && canPreserveOdtSource(model, doc, odtOptions)) return [];
  const losses = new Set((model.extra?.odt?.diagnostics || []).map((d) => d.message));
  const add = (message) => losses.add(message);
  if (fmt === 'docx') return [...losses];
  // Imported Markdown uses these properties to display source constructs.
  // A property already represented in the source is not a conversion loss.
  const paragraphKeys = ['jc', 'ind', 'spacing', 'pBdr', 'shd', 'tabs'];
  const paragraphSignature = (attrs) => JSON.stringify(paragraphKeys.map((key) => attrs[key]));
  const originalParagraphs = new Set();
  const originalMarks = new Set();
  const remember = (node) => {
    if (node.type.name === 'paragraph') originalParagraphs.add(paragraphSignature(node.attrs));
    for (const mark of node.marks) originalMarks.add(JSON.stringify(mark.toJSON()));
  };
  if (fmt === 'md' && model.extra?.md?.bound) {
    for (const chunk of model.extra.md.chunks) for (const node of chunk.nodes || []) {
      remember(node); node.descendants(remember);
    }
  }

  const found = {
    images: false, shapes: false, floating: false, tables: 0, outsideTable: false,
    notes: false, comments: model.comments.size > 0, changes: false,
    sections: false, controls: false, objects: false, fields: false,
    formatting: false, markdownFormatting: false,
  };
  const visit = (content) => content.descendants((node) => {
    const name = node.type.name;
    const a = node.attrs;
    if (name === 'drawing') {
      found.images = true;
      found.shapes ||= !!(a.content?.length || a.kind === 'shape');
      found.floating ||= !!a.anchor;
    }
    if (name === 'table') found.tables++;
    if (name === 'footnote_ref') found.notes = true;
    if (name === 'field_char' || name === 'instr_text') found.fields = true;
    if (name === 'sdt_block' || node.marks.some((m) => m.type.name === 'sdt')) found.controls = true;
    if (name === 'opaque_inline' && a.kind === 'del') found.changes = true;
    if (name === 'opaque_inline' && a.kind === 'comment') found.comments = true;
    if ((name === 'opaque_block' && a.kind !== 'md') || (name === 'opaque_inline' && !['noteMark', 'comment', 'del'].includes(a.kind))) found.objects = true;
    if (a.sectPr) found.sections = true;
    if (name === 'paragraph') {
      if (a.styleId && a.styleId !== 'PlainTextNotepad' && a.styleId !== model.styles.defaultPara) found.formatting = true;
      if (paragraphKeys.some((key) => a[key] != null)) {
        found.formatting = true;
        if (!originalParagraphs.has(paragraphSignature(a))) found.markdownFormatting = true;
      }
    }
    if (node.marks.length) found.formatting = true;
    if (node.marks.some((m) => m.type.name === 'ins')) found.changes = true;
    if (node.marks.some((m) => !['b', 'i', 'strike', 'hyperlink', 'ins'].includes(m.type.name) && !originalMarks.has(JSON.stringify(m.toJSON())))) found.markdownFormatting = true;
  });
  visit(doc);
  for (const map of [model.headers, model.footers, model.footnotes, model.endnotes, model.comments]) {
    for (const entry of map.values()) if (entry.doc) visit(entry.doc);
  }
  doc.forEach((node) => { if (node.type.name !== 'table' && (node.textContent.trim() || node.type.name !== 'paragraph')) found.outsideTable = true; });
  const headers = model.headers.size > 0 || model.footers.size > 0;
  const variants = model.settings.evenAndOdd || model.section.titlePg
    || ['first', 'even'].some((type) => model.section.headers[type] || model.section.footers[type]);

  if (fmt === 'txt') {
    if (found.formatting) add('A formatação de texto, fontes, cores e estilos será removida.');
    if (found.images) add('Imagens e caixas de texto não são incluídas no arquivo de texto.');
    if (found.tables) add('Tabelas serão convertidas em texto separado por tabulações.');
    if (doc.attrs.pageProperties) add('As configurações de página não são incluídas no arquivo de texto.');
  } else if (fmt === 'csv') {
    if (found.tables > 1 || found.outsideTable) add('Só a primeira tabela será salva; o restante do documento será removido.');
    if (model.format !== 'csv' || found.images || found.formatting || doc.attrs.pageProperties) add('CSV conserva apenas os valores das células, sem formatação, imagens ou configuração de página.');
  } else if (fmt === 'md') {
    if (found.markdownFormatting || model.format !== 'md' || doc.attrs.pageProperties) add('Fontes, cores, alinhamento e layout de página não são preservados em Markdown.');
    if (found.shapes) add('Caixas de texto e formas não são preservadas em Markdown.');
  }

  if (['txt', 'csv', 'md', 'html'].includes(fmt)) {
    if (headers) add('Cabeçalhos e rodapés não são incluídos neste formato.');
    if (found.notes || model.footnotes.size || model.endnotes.size) add('Notas de rodapé e de fim não são preservadas neste formato.');
  }
  if (['txt', 'csv', 'md', 'html', 'odt', 'rtf'].includes(fmt)) {
    if (found.comments) add('Comentários e suas referências não são preservados neste formato.');
    if (found.changes) add('O histórico de alterações controladas não é preservado neste formato.');
    if (found.controls) add('Controles de conteúdo serão convertidos em conteúdo comum.');
    if (found.objects) add('Equações, objetos incorporados ou elementos avançados podem ser convertidos em texto ou removidos.');
    if (found.sections || model.section.cols > 1) add('Quebras de seção e configurações de múltiplas colunas não são preservadas neste formato.');
  }
  if (['txt', 'csv', 'md', 'html', 'odt'].includes(fmt) && found.fields) add('Campos automáticos, como número da página, serão convertidos em conteúdo estático.');
  if (['odt', 'rtf'].includes(fmt) && variants) add('Cabeçalhos e rodapés diferentes na primeira página ou nas páginas pares não são preservados.');
  if (['odt', 'rtf'].includes(fmt) && found.floating) add('O posicionamento de imagens e caixas flutuantes não é preservado.');
  return [...losses];
}
