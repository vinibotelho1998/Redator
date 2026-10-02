'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { createTrustedIpc } = require('../services/trusted-ipc.cjs');

function fixture() {
  const handles = new Map();
  const listeners = new Map();
  const win = { isDestroyed: () => false };
  const appUrl = 'file:///C:/Redator/app/index.html';
  const frame = { url: appUrl };
  const sender = { mainFrame: frame };
  const event = { sender, senderFrame: frame };
  const windows = new Set([win]);
  const trusted = createTrustedIpc({ handle: (c, f) => handles.set(c, f), on: (c, f) => listeners.set(c, f) }, { windows, getWindow: (e) => e.sender === sender ? win : null, appUrl });
  return { trusted, handles, listeners, win, event, frame, sender, windows };
}

test('autoriza a janela do aplicativo e preserva argumentos e resultado', async () => {
  const f = fixture();
  f.trusted.handle('file:write', (event, p, bytes) => ({ p, bytes, correct: event === f.event }));
  assert.deepEqual(await f.handles.get('file:write')(f.event, '/nota', [1, 2]), { p: '/nota', bytes: [1, 2], correct: true });
});

test('recusa subframes, documentos externos, emissores desconhecidos e janelas encerradas', () => {
  const f = fixture();
  let calls = 0;
  f.trusted.handle('write', () => calls++);
  for (const event of [
    { ...f.event, senderFrame: { url: f.frame.url } },
    { ...f.event, sender: { mainFrame: f.frame } },
    { ...f.event, senderFrame: null },
    null,
  ]) assert.throws(() => f.handles.get('write')(event), /não autorizada/);
  f.frame.url = 'https://example.test/';
  assert.throws(() => f.handles.get('write')(f.event), /não autorizada/);
  f.frame.url = 'file:///C:/Redator/app/index.html';
  f.win.isDestroyed = () => true;
  assert.throws(() => f.handles.get('write')(f.event), /não autorizada/);
  f.win.isDestroyed = () => false;
  f.windows.delete(f.win);
  assert.throws(() => f.handles.get('write')(f.event), /não autorizada/);
  assert.equal(calls, 0);
});

test('eventos não autorizados são ignorados e falhas de listeners são contidas', async (t) => {
  const f = fixture();
  let called = 0;
  const warnings = [];
  t.mock.method(console, 'warn', (...args) => warnings.push(args));
  f.trusted.on('edit', () => called++);
  f.listeners.get('edit')({ ...f.event, senderFrame: { url: f.frame.url } });
  assert.equal(called, 0);
  f.listeners.get('edit')(f.event);
  assert.equal(called, 1);
  f.trusted.on('sync-failure', () => { throw new Error('sincrona'); });
  f.trusted.on('async-failure', async () => { throw new Error('assincrona'); });
  assert.doesNotThrow(() => f.listeners.get('sync-failure')(f.event));
  f.listeners.get('async-failure')(f.event);
  await new Promise((resolve) => setImmediate(resolve));
  assert.equal(warnings.length, 2);
});
