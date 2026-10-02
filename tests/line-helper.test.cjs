'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { EventEmitter } = require('node:events');
const { createLineHelper } = require('../services/line-helper.cjs');

function child() {
  const proc = new EventEmitter();
  proc.stdin = new EventEmitter();
  proc.stdout = new EventEmitter();
  proc.stdout.setEncoding = () => {};
  proc.writes = [];
  proc.stdin.write = (text, done) => { proc.writes.push(text); if (done) done(); };
  proc.killed = false;
  proc.kill = () => { proc.killed = true; };
  return proc;
}

test('monta respostas divididas em chunks e mantém correspondência entre consultas', async (t) => {
  const proc = child();
  const helper = createLineHelper({ spawn: () => proc, timeout: 1000, skipLine: (line) => line.startsWith('init\t') });
  t.after(() => helper.stop());
  const first = helper.query('ação\nextra');
  const second = helper.query('+nome');
  proc.stdout.emit('data', 'init\tok\r\n1\t');
  proc.stdout.emit('data', 'ações\r\nadded\n');
  assert.deepEqual(await Promise.all([first, second]), ['1\tações', 'added']);
  assert.deepEqual(proc.writes, ['ação extra\n', '+nome\n']);
});

test('falhas assíncronas de spawn, stdin e stdout liberam todas as consultas', async () => {
  for (const surface of ['proc', 'stdin', 'stdout']) {
    const proc = child();
    const helper = createLineHelper({ spawn: () => proc, timeout: 1000 });
    const pending = [helper.query('a'), helper.query('b')];
    (surface === 'proc' ? proc : proc[surface]).emit('error', new Error('falha de processo'));
    assert.deepEqual(await Promise.all(pending), ['', '']);
    assert.equal(proc.killed, true);
  }
  assert.equal(await createLineHelper({ spawn: () => { throw new Error('não existe'); }, timeout: 1000 }).query('a'), '');
});

test('timeout reinicia o processo e resposta tardia não contamina consulta nova', async (t) => {
  const processes = [];
  const helper = createLineHelper({ spawn: () => { const proc = child(); processes.push(proc); return proc; }, timeout: 20 });
  t.after(() => helper.stop());
  assert.deepEqual(await Promise.all([helper.query('demorado'), helper.query('na fila')]), ['', '']);
  assert.equal(processes[0].killed, true);
  const next = helper.query('nova consulta');
  processes[0].emit('exit', 0);
  processes[0].stdout.emit('data', 'resposta antiga\n');
  processes[1].stdout.emit('data', 'resposta nova\n');
  assert.equal(await next, 'resposta nova');
  assert.equal(processes[1].killed, false);
});

test('escrita interrompida e saída excessiva encerram o auxiliar sem rejeição solta', async () => {
  for (const mode of ['throw', 'callback', 'overflow']) {
    const proc = child();
    if (mode === 'throw') proc.stdin.write = () => { throw new Error('pipe encerrado'); };
    if (mode === 'callback') proc.stdin.write = (_line, done) => done(new Error('pipe encerrado'));
    const helper = createLineHelper({ spawn: () => proc, timeout: 1000, maxBuffer: 20 });
    const pending = helper.query('palavra');
    if (mode === 'overflow') proc.stdout.emit('data', 'a'.repeat(21));
    assert.equal(await pending, '');
    assert.equal(proc.killed, true);
  }
});
