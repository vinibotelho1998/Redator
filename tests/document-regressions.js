window.runDocumentRegressions = async function ({ app, rt, expect }) {
  const result = await rt.readMarkdown(new TextEncoder().encode('Texto original'), '');
  app.setDocument(result.model, result.doc, { path: null, format: 'docx', name: 'Histórico' });
  app.setViewMode('print');
  const original = app.model.section.toXml();
  const initialTab = app.model.settings.defaultTabStop;
  const initialText = app.view.state.doc.textContent;
  const initialWidth = app.model.section.pgSz.w;
  app.view.dispatch(app.view.state.tr.insertText('EDITADO ', 1));
  app.setMargins({ top: 1234, bottom: 2345, left: 1567, right: 1678 });
  const changedMargins = { ...app.model.section.pgMar };
  app.setOrientation('landscape');
  app.setDefaultTabStop(1440);
  app.undo();
  expect('histórico: desfazer tabulação restaura configuração do documento', app.model.settings.defaultTabStop === initialTab);
  app.undo();
  expect('histórico: desfazer orientação mantém a ação anterior de margens', app.model.section.pgSz.w === initialWidth && app.model.section.pgMar.top === changedMargins.top);
  app.undo();
  expect('histórico: desfazer margens restaura o XML original da seção', app.model.section.toXml() === original, { original, atual: app.model.section.toXml() });
  app.undo();
  expect('histórico: texto e página seguem a mesma ordem de desfazer', app.view.state.doc.textContent === initialText);
  app.redo(); app.redo(); app.redo(); app.redo();
  expect('histórico: refazer recupera texto, orientação e tabulação', app.view.state.doc.textContent.startsWith('EDITADO ') && app.model.section.pgSz.w > app.model.section.pgSz.h && app.model.settings.defaultTabStop === 1440);
  app.setTitlePage(true); app.setEvenOdd(true);
  app.undo();
  expect('histórico: cabeçalhos pares e ímpares podem ser desfeitos', !app.model.settings.evenAndOdd && app.model.section.titlePg);
  app.undo();
  expect('histórico: primeira página diferente pode ser desfeita', !app.model.section.titlePg);
  const packed = await rt.writeDocx(app.model, app.view.state.doc);
  const reopened = await rt.readDocx(packed);
  expect('histórico: configurações refeitas persistem no DOCX', reopened.model.section.pgSz.w > reopened.model.section.pgSz.h && reopened.model.settings.defaultTabStop === 1440 && !reopened.model.settings.evenAndOdd);
  reopened.model.dispose();

  // Uma edição durante a incorporação da imagem precisa entrar no próximo save.
  // O pacote já existe: a pausa acontece dentro da serialização real da parte.
  if (!app.model.zip) throw new Error('A regressão de cabeçalho exige o pacote DOCX já criado.');
  const schema = app.view.state.schema;
  for (const [kind, label] of [['header', 'cabeçalho'], ['footer', 'rodapé']]) {
    app.insertPageNumber(kind, 'left');
    const map = kind === 'header' ? app.model.headers : app.model.footers;
    const entry = [...map.values()][0];
    const canvas = document.createElement('canvas');
    canvas.width = 2; canvas.height = 2;
    const context = canvas.getContext('2d');
    context.fillStyle = '#185abd'; context.fillRect(0, 0, 2, 2);
    const imageBytes = Uint8Array.from(atob(canvas.toDataURL('image/png').split(',')[1]), (character) => character.charCodeAt(0));
    const imageURL = URL.createObjectURL(new Blob([imageBytes], { type: 'image/png' }));
    const drawing = schema.nodes.drawing.create({ src: imageURL, cx: 19050, cy: 19050, alt: 'Imagem da regressão de ' + label });
    const paragraph = (text) => schema.node('doc', null, [schema.nodes.paragraph.create(null, [schema.text(text), drawing])]);
    const oldText = 'Versão anterior do ' + label;
    const newText = 'Nova edição do ' + label;
    entry.doc = paragraph(oldText);
    entry.dirty = true;
    const originalFetch = window.fetch;
    let releaseImage, announceFetch, timer, pendingWrite, firstRead, secondRead;
    const imageGate = new Promise((resolve) => { releaseImage = resolve; });
    const fetchStarted = new Promise((resolve) => { announceFetch = resolve; });
    window.fetch = function (input, init) {
      const requested = typeof input === 'string' ? input : input.url || String(input);
      if (requested !== imageURL) return originalFetch.call(this, input, init);
      announceFetch();
      return imageGate.then(() => originalFetch.call(window, input, init));
    };
    try {
      pendingWrite = rt.writeDocx(app.model, app.view.state.doc);
      await Promise.race([
        fetchStarted,
        new Promise((_, reject) => { timer = setTimeout(() => reject(new Error('A gravação não chegou à imagem do ' + label)), 5000); }),
      ]);
      clearTimeout(timer);
      const latestDoc = paragraph(newText);
      entry.doc = latestDoc;
      entry.dirty = true;
      releaseImage();
      const firstBytes = await pendingWrite;
      firstRead = await rt.readDocx(firstBytes);
      const firstEntry = [...(kind === 'header' ? firstRead.model.headers : firstRead.model.footers).values()].find((item) => item.path === entry.path);
      expect('salvar: edição durante gravação de ' + label + ' conserva a primeira versão no snapshot', firstEntry?.doc.textContent === oldText);
      expect('salvar: edição durante gravação de ' + label + ' permanece pendente', entry.doc === latestDoc && entry.dirty);

      const secondBytes = await rt.writeDocx(app.model, app.view.state.doc);
      secondRead = await rt.readDocx(secondBytes);
      const secondEntry = [...(kind === 'header' ? secondRead.model.headers : secondRead.model.footers).values()].find((item) => item.path === entry.path);
      expect('salvar: próxima gravação inclui a nova edição do ' + label, secondEntry?.doc.textContent === newText && !entry.dirty);
      const invalidFirst = await rt.validateZip(firstBytes), invalidSecond = await rt.validateZip(secondBytes);
      expect('salvar: snapshots de ' + label + ' com imagem têm XML válido', invalidFirst.length === 0 && invalidSecond.length === 0, { primeiro: invalidFirst, segundo: invalidSecond });
    } finally {
      clearTimeout(timer);
      releaseImage();
      if (pendingWrite) await pendingWrite.catch(() => {});
      window.fetch = originalFetch;
      URL.revokeObjectURL(imageURL);
      firstRead?.model.dispose();
      secondRead?.model.dispose();
    }
  }
};
