// Incluído pelo runner antes de suite.js. As ativações usam teclas nativas do Electron.
window.runUiRegressions = async function ({ app, expect, sleep, rt }) {
  const keys = async (...events) => {
    const before = window.__keysDone || 0;
    const input = events.map((event) => {
      const value = typeof event === 'string' ? { key: event } : event;
      // Chromium dispatches the button's Enter activation from its character event.
      return value.key === 'Return' ? { ...value, char: '\r' } : value;
    });
    console.log('TEST:KEYS:' + JSON.stringify(input));
    const until = Date.now() + 5000;
    while ((window.__keysDone || 0) === before && Date.now() < until) await sleep(20);
    if ((window.__keysDone || 0) === before) throw new Error('O harness não terminou o envio das teclas nativas.');
    await frame();
  };
  const frame = () => new Promise((resolve) => {
    const timer = setTimeout(resolve, 1500);
    requestAnimationFrame(() => { clearTimeout(timer); resolve(); });
  });
  const control = (id) => app.ribbon.controls.get(id).el;
  const tab = (id) => app.ribbon.tabs.find((item) => item.id === id).btn;
  const visibleFocusables = (root) => [...root.querySelectorAll('button,input,select,textarea,[tabindex]')]
    .filter((element) => !element.disabled && element.tabIndex >= 0 && element.getClientRects().length && !element.closest('[hidden]'));
  const originalPrefs = { units: app.prefs.units || 'cm', shortcuts: app.prefs.shortcuts || 'intl' };
  const bytes = new TextEncoder().encode('Texto para conferir os controles e a edição.');

  app.setViewMode('print');
  await app.loadBytes(bytes, window.__tmp + 'interface.txt', false);
  await frame();
  app.ribbon.select('home');
  app.setRuler(true, true);

  // Abas e botões compostos precisam reagir à ativação nativa de button.
  tab('home').focus();
  await keys('Right');
  expect('UI: setas alternam abas e atualizam aria-selected', app.ribbon.active === 'insert' && document.activeElement === tab('insert') && tab('insert').getAttribute('aria-selected') === 'true' && tab('home').getAttribute('aria-selected') === 'false');
  const insertPanel = document.getElementById(tab('insert').getAttribute('aria-controls'));
  expect('UI: aba referencia seu painel visível', insertPanel?.getAttribute('role') === 'tabpanel' && !insertPanel.hidden && insertPanel.getAttribute('aria-labelledby') === tab('insert').id);
  await keys('Home');
  expect('UI: Home retorna à primeira aba', app.ribbon.active === 'home' && document.activeElement === tab('home'));

  for (const key of ['Return', 'Space']) {
    app.hideNavPane();
    control('find').focus();
    await keys(key);
    expect('UI: botão dividido Localizar ativa com ' + key, app.nav.visible && document.activeElement === app.nav.input);
  }
  app.hideNavPane();
  app.view.dispatch(app.view.state.tr.setSelection(rt.TextSelection.create(app.view.state.doc, 1, 6)));
  control('u').focus();
  await keys('Return');
  expect('UI: Enter em botão dividido aplica sublinhado uma vez', app.uiState.underline && control('u').getAttribute('aria-pressed') === 'true' && !document.querySelector('.popup'));
  control('u').focus();
  await keys('Space');
  expect('UI: Espaço em botão dividido desativa sublinhado uma vez', !app.uiState.underline && control('u').getAttribute('aria-pressed') === 'false' && !document.querySelector('.popup'));

  const spacing = control('lineSpacing');
  spacing.focus();
  await keys('Down');
  let popup = document.querySelector('.popup');
  let items = [...popup.querySelectorAll('[role^="menuitem"]')];
  expect('UI: ArrowDown abre menu com foco e aria-expanded', spacing.getAttribute('aria-expanded') === 'true' && spacing.getAttribute('aria-controls') === popup.id && document.activeElement === items[0]);
  await keys('Down');
  expect('UI: ArrowDown move foco ao próximo item', document.activeElement === items[1]);
  await keys('Up');
  expect('UI: ArrowUp retorna ao primeiro item', document.activeElement === items[0]);
  await keys('End');
  expect('UI: End foca o último item', document.activeElement === items[items.length - 1]);
  await keys('Home', 'Down', 'Return');
  expect('UI: Enter executa a opção do menu e fecha o popup', !document.querySelector('.popup') && spacing.getAttribute('aria-expanded') === 'false' && app.uiState.lineSpacing === 1.15);
  spacing.focus();
  await keys('Down', 'Escape');
  expect('UI: Escape fecha menu e devolve foco ao botão', !document.querySelector('.popup') && document.activeElement === spacing && spacing.getAttribute('aria-expanded') === 'false');

  // Painéis antigos com itens div também precisam responder ao teclado.
  const size = app.ribbon.controls.get('size').input;
  size.focus();
  await keys('Down');
  popup = document.querySelector('.size-popup');
  items = [...popup.querySelectorAll('.font-item')];
  expect('UI: lista de tamanhos recebe foco pelo teclado', document.activeElement === items[0]);
  await keys('Down', 'Return');
  expect('UI: Enter em item da lista aplica tamanho', !document.querySelector('.popup') && app.uiState.size === 9);
  size.focus();
  await keys('Down', 'Escape');
  expect('UI: Escape na lista devolve foco ao campo', document.activeElement === size && !document.querySelector('.popup'));

  // Modal: navegação segue as regras do navegador e permanece dentro da caixa.
  const trigger = control('replace');
  trigger.focus();
  await keys('Return');
  let modal = document.querySelector('.dlg');
  expect('UI: diálogo declara modalidade e impede foco no fundo', modal?.getAttribute('aria-modal') === 'true' && app.root.inert && modal.contains(document.activeElement));
  let focusables = visibleFocusables(modal);
  const first = focusables[0], last = focusables[focusables.length - 1];
  last.focus();
  await keys('Tab');
  expect('UI: Tab no último controle retorna ao primeiro', document.activeElement === first);
  await keys({ key: 'Tab', mods: ['shift'] });
  expect('UI: Shift+Tab no primeiro controle retorna ao último', document.activeElement === last);
  await keys('Escape');
  expect('UI: Escape fecha modal e retorna foco ao acionador', !document.querySelector('.dlg-overlay') && !app.root.inert && document.activeElement === trigger);

  // Trocar a unidade altera somente a apresentação de medidas existentes.
  const section = () => JSON.stringify({ margins: app.model.section.pgMar, size: app.model.section.pgSz });
  const originalSection = section();
  for (const [unit, label, suffix] of [['in', 'polegadas', 'pol'], ['pt', 'pontos', 'pt'], ['cm', 'centímetros', 'cm']]) {
    await app.savePrefs({ units: unit });
    await frame();
    app.dialog('pageSetup');
    await sleep(70);
    modal = document.querySelector('.dlg');
    const field = modal.querySelector('.spin-inp');
    expect('UI: preferência ' + unit + ' atualiza régua e campos', app.ruler.el.getAttribute('aria-label') === 'Régua em ' + label && field.value.endsWith(' ' + suffix) && field.getAttribute('aria-valuetext') === field.value);
    await keys('Return');
    expect('UI: confirmar medidas em ' + unit + ' conserva twips exatos', section() === originalSection && !document.querySelector('.dlg-overlay'));
  }
  await app.savePrefs({ units: 'in' });
  app.dialog('pageSetup');
  await frame();
  modal = document.querySelector('.dlg');
  const top = modal.querySelector('.spin-inp');
  top.focus(); top.select();
  await keys({ text: '1 pol' }, 'Return');
  expect('UI: valor digitado em polegadas converte para twips', app.model.section.pgMar.top === 1440);
  await app.savePrefs(originalPrefs);

  // Entrar duas vezes em leitura não deve apagar a condição original do arquivo.
  app.setViewMode('read');
  app.setViewMode('read');
  await frame();
  const readDoc = app.view.state.doc;
  app.focusEditor();
  await keys({ text: 'SEM EDIÇÃO' });
  expect('UI: modo leitura repetido bloqueia digitação', app.readOnly && !app.view.editable && app.view.state.doc.eq(readDoc) && app.status.modes.read.getAttribute('aria-pressed') === 'true');
  await keys('Escape');
  expect('UI: sair da leitura restaura edição do documento comum', app.viewMode === 'print' && !app.readOnly && app.view.editable && app.status.modes.print.getAttribute('aria-pressed') === 'true');
  await app.loadBytes(bytes, window.__tmp + 'somente-leitura.txt', true);
  app.setViewMode('read'); app.setViewMode('read'); app.setViewMode('web');
  await frame();
  expect('UI: sair da leitura conserva arquivo somente leitura', app.documentReadOnly && app.readOnly && !app.view.editable && app.root.classList.contains('document-readonly'));
  await app.loadBytes(bytes, window.__tmp + 'editavel.txt', false);
  app.setViewMode('print');
  await frame();
  expect('UI: documento seguinte restaura edição e estado de leitura', !app.documentReadOnly && !app.readOnly && app.view.editable && !app.root.classList.contains('document-readonly'));

  app.setBusy('Operação de teste');
  await frame();
  expect('UI: operação anuncia status e aria-busy', app.workspace.getAttribute('aria-busy') === 'true' && app.status.info.textContent === 'Operação de teste' && app.status.info.getAttribute('role') === 'status');
  app.setBusy(null);
  await frame();
  expect('UI: concluir operação limpa aria-busy', app.workspace.getAttribute('aria-busy') === 'false');
  expect('UI: zoom tem nome e valor acessíveis', app.status.slider.getAttribute('aria-label') === 'Zoom do documento' && app.status.slider.getAttribute('aria-valuetext') === Math.round(app.layout.zoom * 100) + '%');

  // F6 percorre regiões principais; Shift+F6 faz o caminho de volta.
  app.ribbon.select('home');
  tab('home').focus();
  await keys('F6');
  expect('UI: F6 muda da faixa para o editor', document.activeElement === app.view.dom);
  await keys('F6');
  expect('UI: F6 muda do editor para a barra de status', document.activeElement === app.status.page);
  await keys({ key: 'F6', mods: ['shift'] });
  expect('UI: Shift+F6 retorna ao editor', document.activeElement === app.view.dom);
};
