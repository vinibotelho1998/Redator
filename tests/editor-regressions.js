// Incluído pelo runner antes de suite.js; exercita o clipboard público da view.
window.runEditorRegressions = async function (app, rt, expect) {
  const schema = app.view.state.schema;
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const bytesEqual = (a, b) => a.length === b.length && a.every((v, i) => v === b[i]);
  const image = (color) => {
    const canvas = document.createElement('canvas'); canvas.width = 2; canvas.height = 2;
    const ctx = canvas.getContext('2d'); ctx.fillStyle = color; ctx.fillRect(0, 0, 2, 2);
    return Uint8Array.from(atob(canvas.toDataURL('image/png').split(',')[1]), (c) => c.charCodeAt(0));
  };
  const setup = async (color, kind, label) => {
    const { model } = await rt.readMarkdown(new TextEncoder().encode(label), '');
    const bytes = image(color);
    const path = 'word/media/image1.png';
    model.media.set(path, { bytes, type: 'image/png', ext: 'png', url: URL.createObjectURL(new Blob([bytes], { type: 'image/png' })) });
    model.styles.addStyleXml('<w:style xmlns:w="' + W + '" w:type="paragraph" w:styleId="SameStyle"><w:name w:val="Mesmo nome"/><w:basedOn w:val="Normal"/><w:rPr><w:color w:val="' + color.slice(1) + '"/></w:rPr></w:style>');
    const numId = model.numbering.createList(kind, 0);
    const doc = schema.node('doc', null, [
      schema.nodes.paragraph.create({ styleId: 'SameStyle', numId, ilvl: 0 }, schema.text(label)),
      schema.nodes.paragraph.create(null, schema.nodes.drawing.create({ mediaPath: path, cx: 19050, cy: 19050, alt: label })),
    ]);
    return { model, doc, bytes, numId, path };
  };
  const show = (item) => app.setDocument(item.model, item.doc, { path: null, name: 'Clipboard', format: 'docx' });
  const source = await setup('#ff0000', 'number', 'Origem vermelha');
  show(source);
  const copied = app.view.serializeForClipboard(app.view.state.doc.slice(0));
  const html = copied.dom.innerHTML;
  expect('clipboard: HTML interoperável inclui imagem incorporada e texto simples', /src="data:image\/png;base64,/.test(html) && copied.text.includes('Origem vermelha'));
  const target = await setup('#0000ff', 'bullet', 'Destino azul');
  show(target);
  source.model.dispose();
  source.model.dispose(); // descarte é idempotente
  app.view.dispatch(app.view.state.tr.setSelection(rt.TextSelection.atEnd(app.view.state.doc)));
  app.view.pasteHTML(html);
  let pastedImage = null, pastedParagraph = null;
  app.view.state.doc.descendants((n) => {
    if (n.type.name === 'drawing' && n.attrs.alt === 'Origem vermelha') pastedImage = n;
    if (n.type.name === 'paragraph' && n.textContent === 'Origem vermelha') pastedParagraph = n;
  });
  expect('clipboard: imagem de outra origem abandona caminho/relacionamento antigo', !!pastedImage && pastedImage.attrs.mediaPath == null && pastedImage.attrs.rId == null && pastedImage.attrs.xml == null && pastedImage.attrs.src.startsWith('data:'));
  expect('clipboard: estilos de mesmo ID têm definições independentes', !!pastedParagraph && pastedParagraph.attrs.styleId !== 'SameStyle' && target.model.styles.resolvePara(pastedParagraph.attrs.styleId).r.color.a.val.toLowerCase() === 'ff0000' && target.model.styles.resolvePara('SameStyle').r.color.a.val.toLowerCase() === '0000ff');
  expect('clipboard: listas de mesmo ID não assumem o formato do destino', !!pastedParagraph && pastedParagraph.attrs.numId !== target.numId && target.model.numbering.kindOf(pastedParagraph.attrs.numId) === 'number' && target.model.numbering.kindOf(target.numId) === 'bullet');
  const bytes = await rt.writeDocx(target.model, app.view.state.doc);
  const back = await rt.readDocx(bytes);
  let backImage = null;
  back.doc.descendants((n) => { if (n.type.name === 'drawing' && n.attrs.alt === 'Origem vermelha') backImage = n; });
  const media = backImage && back.model.media.get(backImage.attrs.mediaPath);
  expect('clipboard: salvar/reabrir mantém os bytes da origem após fechá-la', !!media && bytesEqual(media.bytes, source.bytes));
  expect('clipboard: destino conserva a imagem original com o mesmo caminho', bytesEqual(back.model.media.get(target.path).bytes, target.bytes));
  expect('clipboard: pacote DOCX com recursos remapeados tem XML válido', (await rt.validateZip(bytes)).length === 0);
  back.model.dispose();

  // O mesmo documento mantém a identidade para continuar listas e reutilizar mídias.
  const local = app.view.serializeForClipboard(app.view.state.doc.slice(0)).dom.innerHTML;
  const numCount = target.model.numbering.nums.size;
  const styleCount = target.model.styles.styles.size;
  app.view.pasteHTML(local);
  expect('clipboard: cópia interna não duplica definições de estilo/lista', target.model.numbering.nums.size === numCount && target.model.styles.styles.size === styleCount);

  // A antiga condição <8 MiB deixava o clipboard depender de uma URL blob revogada.
  const large = await setup('#00ff00', 'number', 'Imagem grande');
  const bigBytes = new Uint8Array(8 * 1024 * 1024 + 1);
  bigBytes.set(large.bytes);
  const old = large.model.media.get(large.path);
  URL.revokeObjectURL(old.url);
  old.bytes = bigBytes;
  old.url = URL.createObjectURL(new Blob([bigBytes], { type: 'image/png' }));
  show(large);
  const bigHTML = app.view.serializeForClipboard(app.view.state.doc.slice(0)).dom.innerHTML;
  const receiver = await setup('#0000ff', 'bullet', 'Destino da imagem grande');
  show(receiver);
  large.model.dispose();
  app.view.pasteHTML(bigHTML);
  let largeImage = null;
  app.view.state.doc.descendants((n) => { if (n.type.name === 'drawing' && n.attrs.alt === 'Imagem grande') largeImage = n; });
  const embedded = largeImage && largeImage.attrs.src;
  expect('clipboard: imagem maior que 8 MiB mantém bytes após descarte da origem', !!embedded && embedded.startsWith('data:') && atob(embedded.split(',')[1]).length === bigBytes.length);

  // Depois de reabrir um Markdown, suas imagens têm caminhos preservados. Uma
  // imagem nova deve reservar outro nome antes de publicar os anexos em disco.
  const markdownPath = window.__tmp + 'md-anexos.md';
  const oldImagePath = 'md-anexos_arquivos/imagem1.png';
  const oldImageBytes = image('#ff0000');
  const newImageBytes = image('#0000ff');
  const markdownBytes = new TextEncoder().encode('![Imagem antiga](md-anexos_arquivos/imagem1.png)\n');
  await host.writeFile(markdownPath, markdownBytes, [{ path: oldImagePath, data: oldImageBytes }]);
  const existingMarkdown = await rt.readMarkdown(markdownBytes, markdownPath);
  app.setDocument(existingMarkdown.model, existingMarkdown.doc, { path: markdownPath, name: 'md-anexos.md', format: 'md' });
  const newImageSource = 'data:image/png;base64,' + btoa(String.fromCharCode(...newImageBytes));
  app.view.dispatch(app.view.state.tr.insert(0, schema.nodes.paragraph.create(null, schema.nodes.drawing.create({ src: newImageSource, cx: 19050, cy: 19050, alt: 'Imagem nova' }))));
  const markdownExport = await app.serialize('md', markdownPath);
  expect('md: imagem nova reserva nome diferente da imagem reaberta', markdownExport.extras.length === 1 && markdownExport.extras[0].path !== oldImagePath);
  const markdownSaved = await host.writeFile(markdownPath, markdownExport.data, markdownExport.extras);
  expect('md: publicação dos novos anexos conclui com sucesso', markdownSaved.ok === true, markdownSaved);
  if (!markdownSaved.ok) throw new Error('Gravação dos anexos Markdown falhou: ' + JSON.stringify(markdownSaved));
  const oldImageAfter = await host.readFile(window.__tmp + oldImagePath.replace(/\//g, '\\'));
  const newImageAfter = await host.readFile(window.__tmp + markdownExport.extras[0].path.replace(/\//g, '\\'));
  expect('md: salvar imagem nova conserva os bytes dos dois anexos', markdownSaved.ok && bytesEqual(oldImageAfter.data, oldImageBytes) && bytesEqual(newImageAfter.data, newImageBytes));
  const twice = await rt.readMarkdown(markdownExport.data, markdownPath);
  const reopenedImages = [];
  twice.doc.descendants((n) => { if (n.type.name === 'drawing') reopenedImages.push(n); });
  expect('md: reabrir depois da gravação mantém as duas imagens distintas', reopenedImages.length === 2 && new Set(reopenedImages.map((n) => n.attrs.src)).size === 2);
  twice.model.dispose();

  // O pacote original tem partes que o editor não interpreta. Abrir/salvar deve
  // conservar seus bytes; editar exige aviso antes de substituir o arquivo.
  const seed = await rt.readMarkdown(new TextEncoder().encode('Documento OpenDocument'), '');
  seed.model.properties.author = 'Autor original';
  seed.model.properties.title = 'Título original';
  const seedBytes = await rt.writeOdt(seed.doc, seed.model);
  seed.model.dispose();
  const zip = await rt.JSZip.loadAsync(seedBytes);
  const originalStyles = await zip.file('styles.xml').async('string');
  zip.file('styles.xml', originalStyles.replace('</office:master-styles>', '<style:master-page style:name="Paisagem" style:page-layout-name="pm1"><style:header><text:p>Cabeçalho adicional</text:p></style:header></style:master-page></office:master-styles>'));
  const originalContent = await zip.file('content.xml').async('string');
  zip.file('content.xml', originalContent.replace('</text:p>', '<office:annotation><dc:creator>Revisor</dc:creator><text:p>Comentário original</text:p></office:annotation><draw:frame><draw:object-ole xlink:href="Objeto1"/></draw:frame></text:p>'));
  const originalMeta = await zip.file('meta.xml').async('string');
  zip.file('meta.xml', originalMeta.replace('</office:meta>', '<meta:user-defined meta:name="Cliente">ACME</meta:user-defined></office:meta>'));
  zip.file('Objeto1/content.xml', '<objeto>recurso preservado</objeto>');
  const odtBytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  const odt = await rt.readOdt(odtBytes);
  const odtPath = window.__tmp + 'conservacao-original.odt';
  await host.writeFile(odtPath, odtBytes);
  app.setDocument(odt.model, odt.doc, { path: odtPath, name: 'conservacao-original.odt', format: 'odt' });
  const intact = await app.serialize('odt', odtPath);
  expect('odt: salvar sem edição mantém o pacote byte a byte', bytesEqual(intact.data, odtBytes));
  expect('odt: diagnóstico registra masters, comentários, objetos e metadados', ['page-masters', 'comments', 'objects', 'metadata'].every((code) => odt.model.extra.odt.diagnostics.some((d) => d.code === code)));
  expect('odt: pacote intacto não pede confirmação de perdas', rt.lossesForFormat(odt.model, app.view.state.doc, 'odt').length === 0);
  app.view.dispatch(app.view.state.tr.setDocAttribute('pageProperties', { ...app.initialPageProperties }));
  expect('odt: atributo de histórico não altera os bytes se o modelo está intacto', bytesEqual((await app.serialize('odt', odtPath)).data, odtBytes));
  app.view.dispatch(app.view.state.tr.insertText('EDITADO ', 1));
  const losses = rt.lossesForFormat(odt.model, app.view.state.doc, 'odt');
  expect('odt: edição revela perdas mesmo no formato de origem', losses.some((m) => /estilos de página/.test(m)) && losses.some((m) => /Comentários/.test(m)) && losses.some((m) => /Metadados/.test(m)));
  const save = app.writeTo(odtPath, 'odt');
  let cancel = null;
  for (let i = 0; i < 120 && !cancel; i++) {
    cancel = [...document.querySelectorAll('.dlg-overlay button')].find((b) => b.textContent.trim() === 'Cancelar');
    if (!cancel) await new Promise((resolve) => setTimeout(resolve, 25));
  }
  expect('odt: confirmação aparece antes de substituir original editado', !!cancel);
  if (cancel) cancel.click();
  const saved = await save;
  const unchanged = await host.readFile(odtPath);
  expect('odt: cancelar perdas deixa o arquivo original intacto', saved === false && bytesEqual(unchanged.data, odtBytes));
  const modified = await rt.writeOdt(app.view.state.doc, odt.model);
  const modifiedZip = await rt.JSZip.loadAsync(modified);
  expect('odt: edição mantém o autor importado nos metadados', /<meta:initial-creator>Autor original<\/meta:initial-creator>/.test(await modifiedZip.file('meta.xml').async('string')));
  const baseSection = odt.model.section.toXml();
  const reloaded = await rt.readOdt(odtBytes);
  reloaded.model.section.setMargins({ left: reloaded.model.section.pgMar.left + 200 });
  expect('odt: alteração apenas nas margens invalida preservação do pacote', !bytesEqual(await rt.writeOdt(reloaded.doc, reloaded.model), odtBytes));
  expect('odt: alteração de margem preserva diagnóstico de masters', rt.lossesForFormat(reloaded.model, reloaded.doc, 'odt').some((m) => /estilos de página/.test(m)));
  expect('odt: modelo da abertura mantém a seção enquanto outra origem muda', odt.model.section.toXml() === baseSection);
  reloaded.model.dispose();
  const templateBytes = await rt.writeOdt(odt.doc, odt.model, { template: true });
  const template = await rt.readOdt(templateBytes);
  expect('ott: salvar modelo sem edição conserva bytes', bytesEqual(await rt.writeOdt(template.doc, template.model, { template: true }), templateBytes));
  const convertedTemplate = await rt.JSZip.loadAsync(await rt.writeOdt(template.doc, template.model));
  expect('ott: conversão para odt troca o mimetype do pacote', (await convertedTemplate.file('mimetype').async('string')) === 'application/vnd.oasis.opendocument.text');
  template.model.dispose();

  // A recuperação precisa conservar os recursos editados mesmo quando o formato
  // de origem não os suporta. Os escritores e leitores são os de produção;
  // apenas o armazenamento IPC é simulado para inspecionar a cópia capturada.
  if (app.recoverJob) await app.recoverJob;
  const originalHost = app.host;
  const originalAutoRecover = app.prefs.autoRecover;
  let recoveryPacket = null;
  let adopted = 0;
  app.host = {
    ...originalHost,
    saveRecovery: async (data, meta) => {
      recoveryPacket = { data: data.slice(), meta: JSON.parse(JSON.stringify(meta)) };
      return 'editor-recovery-check';
    },
    takeRecovery: async () => recoveryPacket,
    confirmRecovery: async () => { adopted++; return true; },
    releaseRecovery: async () => true,
  };
  app.prefs.autoRecover = true;
  const recover = async () => {
    recoveryPacket = null;
    app.recoverStamp = -1;
    app.markDirty();
    await app.saveRecovery(true);
    const packet = recoveryPacket;
    if (!packet) throw new Error('A AutoRecuperação não capturou o documento de teste.');
    const restored = await app.restoreRecovered('editor-recovery-check');
    expect('recuperação: cópia capturada é adotada após leitura completa', restored === true);
    return packet;
  };
  try {
    const text = await rt.readText(new TextEncoder().encode('Texto com imagem adicionada'), '');
    app.setDocument(text.model, text.doc, { path: null, name: 'recuperacao-imagem.txt', format: 'txt' });
    const recoveryImageBytes = image('#ff00ff');
    const recoveryImage = 'data:image/png;base64,' + btoa(String.fromCharCode(...recoveryImageBytes));
    app.view.dispatch(app.view.state.tr.insert(app.view.state.doc.content.size, schema.nodes.paragraph.create(null, schema.nodes.drawing.create({ src: recoveryImage, cx: 19050, cy: 19050, alt: 'Imagem da recuperação' }))));
    const textPacket = await recover();
    let restoredDrawing = null;
    app.view.state.doc.descendants((n) => { if (n.type.name === 'drawing' && n.attrs.alt === 'Imagem da recuperação') restoredDrawing = n; });
    const restoredMedia = restoredDrawing && app.model.media.get(restoredDrawing.attrs.mediaPath);
    expect('recuperação: TXT com imagem usa pacote DOCX e conserva os bytes', textPacket.meta.kind === 'docx' && !!restoredMedia && bytesEqual(restoredMedia.bytes, recoveryImageBytes));
    expect('recuperação: TXT mantém formato original e sinaliza alterações recuperadas', app.format === 'txt' && app.recovered && app.dirty);

    await app.loadBytes(new TextEncoder().encode('Nome,Valor\nExemplo,1\n'), window.__tmp + 'recuperacao-conteudo.csv', false);
    app.view.dispatch(app.view.state.tr.insert(app.view.state.doc.content.size, schema.nodes.paragraph.create(null, schema.text('Conteúdo fora da tabela precisa sobreviver.'))));
    const csvPacket = await recover();
    expect('recuperação: CSV com texto fora da tabela usa DOCX e mantém o conteúdo', csvPacket.meta.kind === 'docx' && app.view.state.doc.textContent.includes('Conteúdo fora da tabela precisa sobreviver.') && app.view.state.doc.firstChild.type.name === 'table');

    const recoveryOdt = await rt.readOdt(odtBytes);
    app.setDocument(recoveryOdt.model, recoveryOdt.doc, { path: null, name: 'recuperacao-masters.odt', format: 'odt' });
    app.view.dispatch(app.view.state.tr.insertText('RECUPERADO ', 1));
    const odtPacket = await recover();
    const recoveredLosses = rt.lossesForFormat(app.model, app.view.state.doc, 'odt');
    expect('recuperação: ODT conserva diagnóstico de recursos descartados', odtPacket.meta.odtDiagnostics?.some((d) => d.code === 'page-masters') && recoveredLosses.some((m) => /estilos de página/.test(m)) && recoveredLosses.some((m) => /Comentários/.test(m)));
    expect('recuperação: ODT mantém texto editado e pede novo local para salvar', app.view.state.doc.textContent.startsWith('RECUPERADO ') && app.format === 'odt' && app.filePath === null);
    expect('recuperação: ODT mantém autor e título importados', app.model.properties.author === 'Autor original' && app.model.properties.title === 'Título original', app.model.properties);
    expect('recuperação: todas as cópias são confirmadas somente depois de lidas', adopted === 3);
  } finally {
    if (app.recoverJob) await app.recoverJob;
    app.host = originalHost;
    if (originalAutoRecover === undefined) delete app.prefs.autoRecover;
    else app.prefs.autoRecover = originalAutoRecover;
    recoveryPacket = null;
  }

  const notes = await rt.readMarkdown(new TextEncoder().encode('Notas em tabela'), '');
  notes.model.section.setPageSize(8000, 5000);
  notes.model.section.setMargins({ top: 500, bottom: 500, left: 500, right: 500, header: 0, footer: 0 });
  const rows = [];
  for (let i = 0; i < 18; i++) {
    const content = [schema.text('Linha ' + (i + 1) + ': texto de referência.')];
    if (i % 3 === 0) {
      const id = String(i + 1);
      notes.model.footnotes.set(id, { id, doc: schema.node('doc', null, [schema.nodes.paragraph.create(null, schema.text('Nota da linha ' + id))]) });
      content.push(schema.nodes.footnote_ref.create({ id, kind: 'footnote' }));
    }
    rows.push(schema.nodes.table_row.create({ height: 600, hRule: 'atLeast' }, [schema.nodes.table_cell.create(null, [schema.nodes.paragraph.create(null, content)])]));
  }
  notes.doc = schema.node('doc', null, [schema.nodes.table.create(null, rows)]);
  show(notes);
  app.layout.runNow(12);
  const actual = new Map();
  app.layout.pageFns.forEach((ids, page) => ids.forEach((id) => actual.set(String(id), page)));
  const expected = new Map();
  app.view.state.doc.descendants((n, pos) => { if (n.type.name === 'footnote_ref') expected.set(String(n.attrs.id), app.layout.pageOfPos(pos)); });
  expect('tabela: fixture de notas ocupa várias páginas', app.layout.pageCount > 1 && expected.size === 6);
  expect('tabela: cada nota permanece na página da linha da referência', [...expected].every(([id, page]) => actual.get(id) === page), { esperado: [...expected], recebido: [...actual] });
  expect('tabela: notas não são duplicadas nas passadas de paginação', app.layout.pageFns.flat().length === expected.size);
};
