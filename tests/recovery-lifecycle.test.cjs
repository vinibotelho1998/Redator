'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { createFileStore } = require('../services/file-store.cjs');
const { encodeSnapshot, readRecovery, recoveryFiles, RECOVER_ID } = require('../services/recovery-snapshot.cjs');

const source = fs.readFileSync(path.join(__dirname, '../main.js'), 'utf8');
const removeSource = source.slice(source.indexOf('function removeRecovery('), source.indexOf('function scanRecovery('));
const handlerSource = source.slice(source.indexOf("trustedIpc.handle('recover:save'"), source.indexOf("trustedIpc.handle('file:read'"));
const mainRecovery = new Function('trustedIpc', 'windowFor', 'windows', 'recoverFiles', 'fileStore', 'encodeSnapshot', 'readRecovery', 'recoverDir', 'fs', 'RECOVER_ID', 'createWindow', 'HIDDEN_TEST',
  'const recoveryCleanup = new Set();\n' + removeSource + '\n' + handlerSource + '\nreturn { removeRecovery };');
const deferred = () => { let resolve; const promise = new Promise((r) => { resolve = r; }); return { promise, resolve }; };
const meta = { name: 'nota.docx', path: null, format: 'docx', kind: 'docx' };

async function fixture(t, { fileIo = fs.promises, hostIo = fs.promises } = {}) {
  const dir = await fs.promises.mkdtemp(path.join(os.tmpdir(), 'redator-lifecycle-'));
  t.after(() => fs.promises.rm(dir, { recursive: true, force: true }));
  const windows = new Set();
  const handlers = new Map();
  const makeWindow = () => {
    const win = { recoverId: null, recoveryEpoch: 0, recoveryClaim: null, docState: {}, isDestroyed: () => false };
    windows.add(win);
    return win;
  };
  const production = mainRecovery({ handle: (channel, fn) => handlers.set(channel, fn) }, (event) => event.win, windows,
    (id) => recoveryFiles(dir, id), createFileStore(fileIo), encodeSnapshot, readRecovery, dir, { promises: hostIo }, RECOVER_ID, makeWindow, true);
  const invoke = (channel, win, ...args) => handlers.get(channel)({ win }, ...args);
  return { dir, windows, makeWindow, invoke, ...production };
}

test('clear durante preparação do snapshot impede a criação tardia da cópia', async (t) => {
  const gate = deferred();
  const entered = deferred();
  const f = await fixture(t, { hostIo: { ...fs.promises, mkdir: async (...args) => { entered.resolve(); await gate.promise; return fs.promises.mkdir(...args); } } });
  const win = f.makeWindow();
  const saving = f.invoke('recover:save', win, Buffer.from('texto'), meta);
  await entered.promise;
  await f.invoke('recover:clear', win);
  gate.resolve();
  assert.equal(await saving, null);
  assert.deepEqual(await fs.promises.readdir(f.dir), []);
});

test('clear durante publicação aguarda o save e apaga a cópia sem ressuscitá-la', async (t) => {
  const gate = deferred();
  const entered = deferred();
  const f = await fixture(t, { fileIo: { ...fs.promises, rename: async (...args) => { entered.resolve(); await gate.promise; return fs.promises.rename(...args); } } });
  const win = f.makeWindow();
  const saving = f.invoke('recover:save', win, Buffer.from('texto'), meta);
  await entered.promise;
  const clearing = f.invoke('recover:clear', win);
  gate.resolve();
  assert.deepEqual(await Promise.all([saving, clearing]), [null, true]);
  assert.deepEqual(await fs.promises.readdir(f.dir), []);
});

test('duas janelas não adotam a mesma recuperação e release libera uma falha de parse', async (t) => {
  const f = await fixture(t);
  const a = f.makeWindow();
  const b = f.makeWindow();
  await fs.promises.writeFile(recoveryFiles(f.dir, 'nota-1234').snapshot, encodeSnapshot('documento', meta));
  const takingA = f.invoke('recover:take', a, 'nota-1234');
  assert.equal(await f.invoke('recover:take', b, 'nota-1234'), null);
  assert.equal(Buffer.from((await takingA).data).toString(), 'documento');
  assert.equal(a.recoveryClaim.id, 'nota-1234');
  assert.equal(await f.invoke('recover:release', a, 'nota-1234'), true);
  assert.equal(a.recoverId, null);
  assert.equal(a.recoveryClaim, null);
  assert.ok(await f.invoke('recover:take', b, 'nota-1234'));
  assert.equal(await f.invoke('recover:confirm', b, 'nota-1234'), true);
  assert.equal(b.recoveryClaim, null);
  assert.equal(b.recoverId, 'nota-1234');
  assert.equal((await readRecovery(f.dir, 'nota-1234')).data.toString(), 'documento');
});

test('adotar recuperação preserva a anterior até a confirmação do parser', async (t) => {
  const f = await fixture(t);
  const win = f.makeWindow();
  win.recoverId = 'nota-5678';
  for (const id of ['nota-1234', 'nota-5678']) await fs.promises.writeFile(recoveryFiles(f.dir, id).snapshot, encodeSnapshot(id, meta));
  assert.ok(await f.invoke('recover:take', win, 'nota-1234'));
  assert.deepEqual((await fs.promises.readdir(f.dir)).sort(), ['nota-1234.snapshot', 'nota-5678.snapshot']);
  assert.equal(await f.invoke('recover:confirm', win, 'outro-id'), false);
  assert.equal(await f.invoke('recover:confirm', win, 'nota-1234'), true);
  assert.deepEqual(await fs.promises.readdir(f.dir), ['nota-1234.snapshot']);
});

test('take inexistente restaura propriedade anterior e discard respeita claims ativas', async (t) => {
  const f = await fixture(t);
  const win = f.makeWindow();
  win.recoverId = 'nota-5678';
  assert.equal(await f.invoke('recover:take', win, 'nota-1234'), null);
  assert.equal(win.recoverId, 'nota-5678');
  assert.equal(win.recoveryClaim, null);
  await fs.promises.writeFile(recoveryFiles(f.dir, 'nota-1234').snapshot, encodeSnapshot('documento', meta));
  assert.ok(await f.invoke('recover:take', win, 'nota-1234'));
  const other = f.makeWindow();
  assert.equal(await f.invoke('recover:discard', other, ['nota-1234', '../fora']), true);
  assert.equal((await readRecovery(f.dir, 'nota-1234')).data.toString(), 'documento');
});
