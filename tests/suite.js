// Bateria de testes que roda dentro do Redator. Quem sobe o programa e lê o
// resultado é o tests/rodar.mjs; este arquivo só usa o que a janela oferece.
(async () => {
  const FIX = window.__fixtures;
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const out = (...a) => console.log('LOG ' + a.map((x) => (typeof x === 'string' ? x : JSON.stringify(x))).join(' '));
  let ok = 0;
  let fail = 0;
  const pass = (name) => { ok++; out('OK', name); };
  const bad = (name, detail) => { fail++; out('FALHA', name + ':', typeof detail === 'string' ? detail : JSON.stringify(detail)); };
  const expect = (name, cond, detail) => (cond ? pass(name) : bad(name, detail || ''));
  const errs = [];
  window.addEventListener('error', (e) => errs.push('error: ' + (e.message || e)));
  window.addEventListener('unhandledrejection', (e) => errs.push('rejection: ' + ((e.reason && (e.reason.stack || e.reason.message)) || e.reason)));
  const origErr = console.error;
  console.error = (...a) => { errs.push('console.error: ' + a.map((x) => (x && x.stack) || String(x)).join(' ').slice(0, 400)); origErr.apply(console, a); };
  const sameBytes = (a, b) => {
    if (a.length !== b.length) return false;
    for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
    return true;
  };
  const diffLines = (a, b) => {
    const A = a.split('\n'), B = b.split('\n');
    const cnt = new Map();
    for (const l of A) cnt.set(l, (cnt.get(l) || 0) + 1);
    const added = [];
    for (const l of B) { const c = cnt.get(l) || 0; if (c > 0) cnt.set(l, c - 1); else added.push(l); }
    const removed = [];
    for (const [l, c] of cnt) for (let k = 0; k < c; k++) removed.push(l);
    return { removed, added };
  };
  const sig = (doc) => {
    const list = [];
    doc.forEach((n) => {
      if (n.type.name === 'paragraph' && !n.content.size && !n.attrs.pBdr) return;
      const md = n.attrs.md || {};
      list.push([n.type.name, n.attrs.styleId || '', n.attrs.numId ? 'L' + (n.attrs.ilvl || 0) : '', md.quote ? 'Q' : '', md.callout || '', n.textContent.replace(/\s+/g, ' ').trim()].join('|'));
    });
    return list;
  };

  try {
    const app = window.redator;
    for (let i = 0; i < 100 && !(app && app.view); i++) await sleep(100);
    const rt = window.__rt;
    const open = async (name) => {
      const p = FIX + name;
      const f = await host.readFile(p);
      app.pristine = true; app.dirty = false; app.filePath = null;
      await app.loadBytes(f.data, p, false);
      app.layout.runNow(8);
      return { p, bytes: f.data };
    };
    const insertAtStart = (view, match, text) => {
      let at = null;
      view.state.doc.descendants((n, pos) => {
        if (at != null) return false;
        if (n.isTextblock && n.textContent.includes(match)) {
          n.forEach((c, off) => { if (at == null && c.isText) at = pos + 1 + off; });
          return false;
        }
        return true;
      });
      if (at == null) throw new Error('trecho não encontrado: ' + match);
      const marks = view.state.doc.resolve(at).marks().filter((m) => m.type.name !== 'hyperlink');
      view.dispatch(view.state.tr.insert(at, view.state.schema.text(text, marks)));
    };
    const insertAtEnd = (view, match, text) => {
      let at = null;
      view.state.doc.descendants((n, pos) => {
        if (at != null) return false;
        if (n.isTextblock && n.textContent.includes(match)) { at = pos + 1 + n.content.size; return false; }
        return true;
      });
      if (at == null) throw new Error('trecho não encontrado: ' + match);
      view.dispatch(view.state.tr.insertText(text, at));
    };

    // 1. Markdown salvo sem mexer volta byte a byte igual
    for (const name of ['nota-obsidian.md', 'readme-simples.md', 'readme-crlf.md']) {
      const { p, bytes } = await open(name);
      const { data } = await app.serialize('md', p);
      expect('md sem edição igual ao original: ' + name, sameBytes(data, bytes), { tamanhoOriginal: bytes.length, tamanhoGravado: data.length, inicio: new TextDecoder().decode(data).slice(0, 200) });
    }

    // 2. Markdown editado troca só as linhas editadas
    {
      const { p, bytes } = await open('nota-obsidian.md');
      const orig = new TextDecoder().decode(bytes);
      const v = app.view;
      insertAtStart(v, 'Parágrafo com', 'EDITADO ');
      insertAtStart(v, 'Corpo do callout', 'X ');
      insertAtEnd(v, 'outro subitem', ' Z');
      insertAtEnd(v, 'return a + b', ' // fim');
      let cellPos = null;
      v.state.doc.descendants((n, pos) => { if (cellPos == null && n.type.name === 'table_cell' && n.textContent === 'a') cellPos = pos; return cellPos == null; });
      v.dispatch(v.state.tr.insertText('2', cellPos + 2 + 1));
      const text = new TextDecoder().decode((await app.serialize('md', p)).data);
      const d = diffLines(orig, text);
      const want = {
        removed: ['Parágrafo com **negrito**, *itálico*, __negrito com sublinhado__, _itálico com sublinhado_,', '~~tachado~~ e `código em linha`. A linha foi quebrada no meio de propósito,', 'como muita gente faz em arquivos de documentação.',
          '> Corpo do callout com `código` e [[link]].', '  * outro subitem', '\treturn a + b; // tabulação no começo', '| a      |    b     |         c |'],
        added: ['EDITADO Parágrafo com **negrito**, *itálico*, **negrito com sublinhado**, *itálico com sublinhado*, ~~tachado~~ e `código em linha`. A linha foi quebrada no meio de propósito, como muita gente faz em arquivos de documentação.',
          '> X Corpo do callout com `código` e [[link]].', '  * outro subitem Z', '\treturn a + b; // tabulação no começo // fim', '| a2 | b | c |'],
      };
      const sameSet = (x, y) => x.length === y.length && [...x].sort().join('\n') === [...y].sort().join('\n');
      expect('md editado muda só as linhas editadas', sameSet(d.removed, want.removed) && sameSet(d.added, want.added), d);
      const back = await rt.readMarkdown(new TextEncoder().encode(text), p);
      const s1 = sig(v.state.doc), s2 = sig(back.doc);
      const first = s1.findIndex((s, i) => s !== s2[i]);
      expect('md editado relido com a mesma estrutura', first < 0 && s1.length === s2.length, { i: first, editado: s1[first], relido: s2[first] });
    }

    // 3. Estrutura do Markdown: callouts, links do Obsidian, anexo e comentário
    {
      await open('nota-obsidian.md');
      const doc = app.view.state.doc;
      let callouts = 0, heads = 0, wiki = 0, img = null, raw = 0, front = 0;
      doc.descendants((n) => {
        if (n.type.name === 'paragraph' && n.attrs.md && n.attrs.md.callout) { callouts++; if (n.attrs.md.head) heads++; }
        if (n.type.name === 'paragraph' && n.attrs.styleId === 'FrontMatter') front++;
        if (n.isText && n.marks.some((m) => m.type.name === 'hyperlink' && /^wiki:/.test(m.attrs.href || ''))) wiki++;
        if (n.type.name === 'drawing') img = n.attrs;
        if (n.type.name === 'opaque_block' && n.attrs.kind === 'md') raw++;
        return true;
      });
      expect('md: propriedades do Obsidian viram bloco próprio', front === 5, { front });
      expect('md: três callouts reconhecidos', heads === 3 && callouts >= 5, { heads, callouts });
      expect('md: links [[...]] reconhecidos', wiki >= 4, { wiki });
      expect('md: anexo ![[...]] encontrado com tamanho', !!img && /figura\.png/.test(img.src || '') && img.cx > 0, img);
      expect('md: comentário e definição de link preservados', raw >= 2, { raw });
      const found = host.findNote ? await host.findNote(FIX + 'nota-obsidian.md', 'figura.png') : null;
      expect('busca de nota no cofre', !!found && /figura\.png$/i.test(found), { found });
    }

    // 4. Texto em Windows-1252 com CRLF e tabulações volta igual
    {
      const { p, bytes } = await open('texto-ansi.txt');
      const { data } = await app.serialize('txt', p);
      expect('txt Windows-1252 sem edição igual ao original', sameBytes(data, bytes), { original: bytes.length, gravado: data.length });
    }

    // 5. docx: gravar, reler e gravar de novo dá o mesmo documento
    {
      await open('nota-obsidian.md');
      const b1 = await rt.writeDocx(app.model, app.view.state.doc, { author: 'Teste' });
      const bad1 = await rt.validateZip(b1);
      const r1 = await rt.readDocx(b1);
      const b2 = await rt.writeDocx(r1.model, r1.doc, { author: 'Teste' });
      const r2 = await rt.readDocx(b2);
      const d = rt.firstDiff(rt.normalize(r1.doc.toJSON()), rt.normalize(r2.doc.toJSON()));
      expect('docx gerado sem XML inválido', bad1.length === 0, bad1);
      expect('docx estável em ida e volta', !d, d);
      app.pristine = true; app.dirty = false; app.filePath = null;
      await app.loadBytes(b2, FIX + 'gerado.docx', false);
      app.layout.runNow(8);
      expect('docx gerado abre e pagina', app.layout.pageCount >= 1 && app.view.state.doc.childCount > 20, { paginas: app.layout.pageCount });
      for (const fmt of ['odt', 'html', 'txt', 'md']) {
        try {
          const { data } = await app.serialize(fmt, FIX + 'saida.' + fmt);
          expect('exporta ' + fmt, data && data.length > 100, { tamanho: data && data.length });
        } catch (e) { bad('exporta ' + fmt, e.stack); }
      }
    }

    // 6. Paginação de um documento longo
    {
      app.pristine = true; app.dirty = false; app.filePath = null;
      await app.newDocument();
      const v = app.view;
      const schema = v.state.schema;
      const paras = [];
      for (let i = 0; i < 1500; i++) paras.push(schema.nodes.paragraph.create(null, schema.text('Linha ' + i + ' com texto suficiente para ocupar espaço na página e quebrar em algum momento.\tfim')));
      v.dispatch(v.state.tr.replaceWith(0, v.state.doc.content.size, paras));
      const t0 = performance.now();
      app.layout.runNow(10);
      const ms = performance.now() - t0;
      expect('pagina 1500 parágrafos rápido', app.layout.pageCount > 20 && ms < (window.__ci ? 15000 : 4000), { paginas: app.layout.pageCount, ms: Math.round(ms) });
      const t1 = performance.now();
      for (let k = 0; k < 20; k++) v.dispatch(v.state.tr.insertText('x', 5));
      app.layout.runNow(4);
      expect('digitação em documento longo', performance.now() - t1 < (window.__ci ? 12000 : 3000), { ms: Math.round(performance.now() - t1) });
    }

    // 7. Faixa de opções e caixas de diálogo não quebram
    {
      const before = errs.length;
      const v = app.view;
      v.dispatch(v.state.tr.setSelection(rt.TextSelection.create(v.state.doc, 2, 10)));
      for (const id of ['b', 'i', 'u', 'strike', 'sub', 'sup', 'bullets', 'numbering', 'indent', 'outdent', 'alignCenter', 'justify', 'alignLeft', 'marks', 'marks', 'clearFmt']) {
        const c = app.ribbon.controls.get(id);
        if (c) (c.el.querySelector('.rb-split-main') || c.el).click();
        await sleep(20);
      }
      for (const [name, arg] of [['font'], ['paragraph'], ['pageSetup'], ['find', 1], ['insertTable'], ['link'], ['wordCount'], ['symbol'], ['dateTime'], ['zoom'], ['options']]) {
        app.dialog(name, arg);
        await sleep(80);
        for (let k = 0; k < 3 && document.querySelector('.dlg-overlay'); k++) { document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })); await sleep(40); }
      }
      for (const tab of ['insert', 'layout', 'references', 'review', 'view', 'home']) { app.ribbon.select(tab); await sleep(40); }
      expect('botões e diálogos sem erro', errs.length === before, errs.slice(before));
    }
    expect('nenhum erro no console durante os testes', errs.length === 0, errs.slice(0, 5));
  } catch (err) {
    bad('execução', err && err.stack);
  }
  out('FIM ok=' + ok + ' falhas=' + fail);
  console.log('TEST:QUIT');
})();
