'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

// Exercise the production state transitions with controllable disk/dialog
// boundaries, without launching Electron or writing a user's document.
const source = fs.readFileSync(path.join(__dirname, '../src/app.js'), 'utf8');
const between = (from, to) => source.slice(source.indexOf(from), source.indexOf(to, source.indexOf(from)));
const methods = [
  between('  markClean(', '  // ------------------------------------------------------------ AutoRecuperação'),
  between('  async saveAs(', '  async performSaveAs('),
  between('  async writeTo(', '  closeWindow()'),
].join('\n');
const production = new Function('D', 'baseName', 'class SaveState {' + methods + '} return SaveState.prototype;');
const deferred = () => {
  let resolve;
  const promise = new Promise((r) => { resolve = r; });
  return { promise, resolve };
};
const tick = () => new Promise((r) => setImmediate(r));

function makeApp(dialogs = {}) {
  const app = {
    model: { format: 'md' }, editStamp: 1, dirty: true, text: 'versão 1', format: 'md', docName: 'nota.md',
    filePath: '/nota.md', fileStamp: { path: '/nota.md', hash: 'original' },
    saveJob: null, recoverJob: null, saveAsJob: null,
    cleared: 0, closed: 0, recoveryRequests: 0, errors: [],
    confirmLossy: async () => true,
    updateTitle() {}, setBusy() {}, setInfo() {},
    serialize: async function () { return { data: this.text }; },
    saveRecovery() { this.recoveryRequests++; },
    showError(_message, err) { this.errors.push(err); },
    save() { return this.writeTo(this.filePath, this.format); },
  };
  Object.setPrototypeOf(app, production({ messageBox: async () => 'cancel', confirmSave: async () => 'cancel', ...dialogs }, (p) => path.basename(p)));
  app.host = {
    writeFile: async () => ({ ok: true, stamp: { hash: 'saved' } }),
    clearRecovery: () => app.cleared++,
    closeConfirmed: () => app.closed++,
  };
  return app;
}

test('edição durante a gravação continua suja e conserva a recuperação', async () => {
  const app = makeApp();
  const disk = deferred();
  const writes = [];
  app.host.writeFile = async (_p, data) => { writes.push(data); await disk.promise; return { ok: true, stamp: { hash: 'v1' } }; };
  const saving = app.save();
  await tick();
  app.text = 'versão 2'; app.editStamp++;
  disk.resolve();
  assert.equal(await saving, true);
  assert.deepEqual(writes, ['versão 1']);
  assert.equal(app.dirty, true);
  assert.equal(app.cleared, 0);
  assert.equal(app.recoveryRequests, 1);
});

test('salvamentos concorrentes são serializados e o segundo captura as novas edições', async () => {
  const app = makeApp();
  const firstDisk = deferred();
  const writes = [];
  let active = 0;
  let peak = 0;
  app.host.writeFile = async (_p, data, _extras, options) => {
    active++; peak = Math.max(peak, active);
    writes.push({ data, expected: options.expected.hash });
    if (writes.length === 1) await firstDisk.promise;
    active--;
    return { ok: true, stamp: { hash: data } };
  };
  const first = app.save();
  await tick();
  app.text = 'versão 2'; app.editStamp++;
  const second = app.save();
  await tick();
  assert.equal(writes.length, 1);
  firstDisk.resolve();
  assert.deepEqual(await Promise.all([first, second]), [true, true]);
  assert.equal(peak, 1);
  assert.deepEqual(writes, [
    { data: 'versão 1', expected: 'original' },
    { data: 'versão 2', expected: 'versão 1' },
  ]);
  assert.equal(app.dirty, false);
  assert.equal(app.cleared, 1);
  assert.equal(app.saveJob, null);
});

test('falha de gravação mantém alterações e libera a fila para nova tentativa', async () => {
  const app = makeApp();
  app.host.writeFile = async () => ({ ok: false, error: 'ENOSPC' });
  assert.equal(await app.save(), false);
  assert.equal(app.dirty, true);
  assert.equal(app.cleared, 0);
  assert.equal(app.saving, false);
  app.host.writeFile = async () => ({ ok: true, stamp: { hash: 'saved' } });
  assert.equal(await app.save(), true);
  assert.equal(app.dirty, false);
});

test('falha de rollback mostra onde recuperar a imagem anterior', async () => {
  const messages = [];
  const app = makeApp({ messageBox: async (_title, text) => { messages.push(text); return 'cancel'; } });
  app.host.writeFile = async () => ({ ok: false, error: 'EBUSY', rollbackErrors: [{ path: '/figura.png', backup: '/.figura.backup', error: 'EBUSY' }] });
  assert.equal(await app.save(), false);
  assert.match(messages[0], /versão anterior foi preservada/);
  assert.ok(messages[0].includes('/.figura.backup'));
  assert.equal(app.dirty, true);
  assert.equal(app.cleared, 0);
});

test('fechar aguarda o disco e não descarta edição posterior ao snapshot', async () => {
  let prompts = 0;
  const app = makeApp({ confirmSave: async () => { prompts++; return 'cancel'; } });
  const disk = deferred();
  app.host.writeFile = async () => { await disk.promise; return { ok: true, stamp: { hash: 'v1' } }; };
  const saving = app.save();
  await tick();
  const closing = app.requestClose();
  app.editStamp++;
  disk.resolve();
  await Promise.all([saving, closing]);
  assert.equal(prompts, 1);
  assert.equal(app.closed, 0);
  assert.equal(app.dirty, true);
});

test('Salvar no diálogo de fechar não fecha se houve nova edição enquanto salvava', async () => {
  const app = makeApp({ confirmSave: async () => 'save' });
  const disk = deferred();
  app.host.writeFile = async () => { await disk.promise; return { ok: true, stamp: { hash: 'v1' } }; };
  const closing = app.requestClose();
  await tick();
  app.editStamp++;
  disk.resolve();
  await closing;
  assert.equal(app.closed, 0);
  assert.equal(app.dirty, true);
  assert.equal(app.closingRequested, false);
});

test('fechar após salvar a versão atual fecha uma vez', async () => {
  const app = makeApp({ confirmSave: async () => 'save' });
  await Promise.all([app.requestClose(), app.requestClose()]);
  assert.equal(app.closed, 1);
  assert.equal(app.dirty, false);
});

test('cancelar conflito conserva alterações e não tenta sobrescrever', async () => {
  const app = makeApp();
  let calls = 0;
  app.host.writeFile = async () => { calls++; return { ok: false, error: 'ECONFLICT', current: { hash: 'external' } }; };
  assert.equal(await app.save(), false);
  assert.equal(calls, 1);
  assert.equal(app.dirty, true);
  assert.equal(app.cleared, 0);
  assert.equal(app.fileStamp.hash, 'original');
});

test('substituição explícita usa a versão mostrada no conflito', async () => {
  const app = makeApp({ messageBox: async () => 'overwrite' });
  const expected = [];
  app.host.writeFile = async (_p, _data, _extras, options) => {
    expected.push(options.expected.hash);
    return expected.length === 1 ? { ok: false, error: 'ECONFLICT', current: { hash: 'external' } } : { ok: true, stamp: { hash: 'saved' } };
  };
  assert.equal(await app.save(), true);
  assert.deepEqual(expected, ['original', 'external']);
  assert.equal(app.dirty, false);
});

test('Salvar cópia sai da fila anterior e grava em outro destino', { timeout: 2000 }, async () => {
  const app = makeApp({ messageBox: async () => 'copy' });
  const paths = [];
  app.host.writeFile = async (p) => {
    paths.push(p);
    return p === '/nota.md' ? { ok: false, error: 'ECONFLICT', current: { hash: 'external' } } : { ok: true, stamp: { hash: 'copy' } };
  };
  app.performSaveAs = async (fmt) => app.writeTo('/copia.md', fmt);
  assert.equal(await app.save(), true);
  assert.deepEqual(paths, ['/nota.md', '/copia.md']);
  assert.equal(app.filePath, '/copia.md');
});

test('gravação antiga não troca o documento aberto durante a espera', async () => {
  const app = makeApp();
  const disk = deferred();
  app.host.writeFile = async () => { await disk.promise; return { ok: true, stamp: { hash: 'saved' } }; };
  const saving = app.save();
  await tick();
  app.model = { format: 'txt' }; app.filePath = '/outro.txt'; app.dirty = true;
  disk.resolve();
  assert.equal(await saving, true);
  assert.equal(app.filePath, '/outro.txt');
  assert.equal(app.dirty, true);
  assert.equal(app.cleared, 0);
});

test('comandos repetidos de Salvar como compartilham o mesmo diálogo', async () => {
  const app = makeApp();
  const dialog = deferred();
  let calls = 0;
  app.performSaveAs = async () => { calls++; await dialog.promise; return true; };
  const first = app.saveAs('md');
  const second = app.saveAs('md');
  dialog.resolve();
  assert.deepEqual(await Promise.all([first, second]), [true, true]);
  assert.equal(calls, 1);
});

test('salvar e fechar aguardam a limpeza assíncrona da recuperação', async () => {
  const app = makeApp();
  const cleanup = deferred();
  let finished = false;
  app.host.clearRecovery = () => cleanup.promise;
  const saving = app.save().then((result) => { finished = true; return result; });
  await tick();
  const closing = app.requestClose();
  await tick();
  assert.equal(app.dirty, false);
  assert.equal(finished, false);
  assert.equal(app.closed, 0);
  cleanup.resolve();
  assert.equal(await saving, true);
  await closing;
  assert.equal(app.closed, 1);
});

test('descartar ao fechar aguarda snapshot em andamento', async () => {
  const app = makeApp({ confirmSave: async () => 'discard' });
  const recovery = deferred();
  app.recoverJob = recovery.promise;
  const closing = app.requestClose();
  await tick();
  assert.equal(app.closed, 0);
  recovery.resolve();
  await closing;
  assert.equal(app.closed, 1);
});

const recoveryProduction = new Function('writeDocx', 'writeMarkdown', 'writeText', 'writeCsv', 'encodeText', 'dirName', 'stripExt', 'baseName', 'lossesForFormat',
  'class RecoveryState {' + between('  saveRecovery(force)', '  async restoreRecovered(') + '} return RecoveryState.prototype;');

function makeRecoveryApp(writer) {
  const model = { extra: {} };
  const app = {
    prefs: {}, view: { state: { doc: { text: 'versão inicial' } } }, model,
    filePath: '/nota.docx', docName: 'nota.docx', format: 'docx', dirty: true,
    editStamp: 1, recoverStamp: 0, recoverJob: null, lastRecover: 0, lastEdit: 0,
    host: { saveRecovery: async () => 'recover-1234' },
  };
  Object.setPrototypeOf(app, recoveryProduction(writer, null, null, null, null, path.dirname, null, path.basename, () => []));
  return app;
}

test('AutoRecuperação mantém nome e caminho do snapshot capturado antes da serialização', async () => {
  const serializing = deferred();
  const app = makeRecoveryApp(async () => { await serializing.promise; return Buffer.from('snapshot'); });
  const received = [];
  app.host.saveRecovery = async (data, meta) => { received.push({ data, meta }); return 'recover-1234'; };
  const saving = app.saveRecovery(true);
  app.docName = 'renomeado.docx'; app.filePath = '/renomeado.docx'; app.editStamp++;
  serializing.resolve();
  await saving;
  assert.equal(received.length, 1);
  assert.equal(received[0].meta.name, 'nota.docx');
  assert.equal(received[0].meta.path, '/nota.docx');
  assert.equal(app.recoverStamp, 1);
  assert.equal(app.editStamp, 2);
  assert.equal(app.recoverJob, null);
});

test('AutoRecuperação não publica documento antigo após a troca de modelo', async () => {
  const serializing = deferred();
  const app = makeRecoveryApp(async () => { await serializing.promise; return Buffer.from('snapshot'); });
  let saves = 0;
  app.host.saveRecovery = async () => { saves++; return 'recover-1234'; };
  const saving = app.saveRecovery(true);
  app.model = { extra: {} };
  serializing.resolve();
  await saving;
  assert.equal(saves, 0);
  assert.equal(app.recoverStamp, 0);
  assert.equal(app.recoverJob, null);
});

const restoreProduction = new Function('readDocx', 'readMarkdown', 'readText', 'readCsv',
  'class RestoreState {' + between('  async restoreRecovered(', '  async offerRecovery(') + '} return RestoreState.prototype;');

test('restauração que falha no parser libera claim e preserva a cópia', async () => {
  const released = [];
  let confirmed = 0;
  const app = {
    host: {
      takeRecovery: async () => ({ meta: { kind: 'docx' }, data: new Uint8Array([1]) }),
      confirmRecovery: async () => confirmed++,
      releaseRecovery: async (id) => released.push(id),
    },
  };
  Object.setPrototypeOf(app, restoreProduction(async () => { throw new Error('DOCX ilegível'); }, null, null, null));
  await assert.rejects(app.restoreRecovered('nota-1234'), /DOCX ilegível/);
  assert.deepEqual(released, ['nota-1234']);
  assert.equal(confirmed, 0);
});

test('restauração confirma adoção depois de instalar o documento e manter dirty', async () => {
  const events = [];
  const model = { extra: {} };
  const app = {
    host: {
      takeRecovery: async () => ({ meta: { kind: 'docx', format: 'docx', name: 'nota.docx', path: '/nota.docx' }, data: new Uint8Array([1]) }),
      confirmRecovery: async (id) => { events.push(['confirm', id, app.dirty]); },
      releaseRecovery: async () => { throw new Error('não deve liberar uma adoção concluída'); },
    },
    setDocument: (m) => { app.model = m; events.push(['document']); },
    updateTitle: () => events.push(['title']),
    setInfo: () => {},
  };
  Object.setPrototypeOf(app, restoreProduction(async () => ({ model, doc: {} }), null, null, null));
  assert.equal(await app.restoreRecovered('nota-1234'), true);
  assert.deepEqual(events, [['document'], ['title'], ['confirm', 'nota-1234', true]]);
  assert.equal(app.recovered, true);
  assert.equal(app.importPath, '/nota.docx');
});
