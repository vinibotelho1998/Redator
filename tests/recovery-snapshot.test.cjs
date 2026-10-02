'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const { encodeSnapshot, decodeSnapshot, scanRecovery, readRecovery, recoveryFiles } = require('../services/recovery-snapshot.cjs');
const { createFileStore } = require('../services/file-store.cjs');

async function fixture(t) {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'redator-recovery-'));
  t.after(() => fs.rm(dir, { recursive: true, force: true }));
  return dir;
}

const meta = { name: 'Nota.md', path: 'C:\\Notas\\Nota.md', format: 'md', kind: 'md', text: { encoding: 'utf-8', bom: true } };

test('snapshot guarda metadados e bytes numa publicação única, incluindo UTF-8 e documento vazio', () => {
  for (const data of [Buffer.from('ação\n漢字'), Buffer.alloc(0), Buffer.from([0, 255, 8])]) {
    const restored = decodeSnapshot(encodeSnapshot(data, meta));
    assert.deepEqual(restored.data, data);
    assert.equal(restored.meta.name, meta.name);
    assert.deepEqual(restored.meta.text, meta.text);
    assert.ok(restored.meta.time > 0);
    assert.match(restored.meta.hash, /^[0-9a-f]{64}$/);
  }
});

test('recusa cabeçalhos inválidos, truncamento e alterações de conteúdo', () => {
  const original = encodeSnapshot(Buffer.from('versão salva'), meta);
  for (const cut of [0, 5, 19, original.length - 1]) assert.throws(() => decodeSnapshot(original.subarray(0, cut)));
  const damaged = Buffer.from(original);
  damaged[damaged.length - 1] ^= 1;
  assert.throws(() => decodeSnapshot(damaged), /corrompida/);
  assert.throws(() => encodeSnapshot('texto', { ...meta, name: { bad: true } }), /Metadados/);
  assert.throws(() => encodeSnapshot('texto', { ...meta, kind: 'exe' }), /Formato/);
  assert.throws(() => encodeSnapshot('texto', { ...meta, name: 'a'.repeat(70000) }), /grandes/);
});

test('descobre e restaura o arquivo snapshot sem depender do par legado', async (t) => {
  const dir = await fixture(t);
  const files = recoveryFiles(dir, 'nota-1234');
  assert.equal((await createFileStore().writeFile(files.snapshot, encodeSnapshot('texto novo', meta))).ok, true);
  const found = scanRecovery(dir);
  assert.equal(found.length, 1);
  assert.equal(found[0].id, 'nota-1234');
  assert.equal((await readRecovery(dir, found[0].id)).data.toString(), 'texto novo');
  assert.deepEqual(await fs.readdir(dir), ['nota-1234.snapshot']);
});

test('lê pares legados e dá preferência ao snapshot válido', async (t) => {
  const dir = await fixture(t);
  const files = recoveryFiles(dir, 'nota-1234');
  await fs.writeFile(files.meta, JSON.stringify({ ...meta, time: 100 }));
  await fs.writeFile(files.data, 'versão antiga');
  assert.equal((await readRecovery(dir, 'nota-1234')).data.toString(), 'versão antiga');
  assert.equal(scanRecovery(dir)[0].time, 100);
  await fs.writeFile(files.snapshot, encodeSnapshot('versão nova', meta));
  assert.equal((await readRecovery(dir, 'nota-1234')).data.toString(), 'versão nova');
  assert.equal(scanRecovery(dir).length, 1);
});

test('snapshot danificado permite recuperar par legado válido sem apagar arquivos', async (t) => {
  const dir = await fixture(t);
  const files = recoveryFiles(dir, 'nota-1234');
  await fs.writeFile(files.snapshot, 'interrompido');
  await fs.writeFile(files.meta, JSON.stringify({ ...meta, time: 100 }));
  await fs.writeFile(files.data, 'versão antiga');
  const names = await fs.readdir(dir);
  assert.equal(scanRecovery(dir).length, 1);
  assert.equal((await readRecovery(dir, 'nota-1234')).data.toString(), 'versão antiga');
  assert.deepEqual(await fs.readdir(dir), names);
});

test('scanning preserva cópias incompletas e ignora IDs/tipos inválidos', async (t) => {
  const dir = await fixture(t);
  await fs.writeFile(path.join(dir, 'nota-1234.snapshot'), 'incompleto');
  await fs.writeFile(path.join(dir, 'nota-5678.dados'), 'órfão');
  await fs.writeFile(path.join(dir, 'nota-9012.json'), JSON.stringify({ name: {}, time: 'ontem' }));
  await fs.writeFile(path.join(dir, 'nota-9012.dados'), 'texto');
  await fs.writeFile(path.join(dir, 'ab.json'), JSON.stringify(meta));
  const names = await fs.readdir(dir);
  assert.deepEqual(scanRecovery(dir), []);
  assert.deepEqual(await fs.readdir(dir), names);
  await assert.rejects(readRecovery(dir, '../escape'), /Identificador/);
  assert.throws(() => recoveryFiles(dir, { toString: () => 'nota-1234' }), /Identificador/);
});

test('falha na publicação conserva o snapshot anterior completo', async (t) => {
  const dir = await fixture(t);
  const files = recoveryFiles(dir, 'nota-1234');
  const original = encodeSnapshot('versão antiga', meta);
  await fs.writeFile(files.snapshot, original);
  const store = createFileStore({ ...fs, rename: async () => { throw Object.assign(new Error('ocupado'), { code: 'EBUSY' }); } });
  assert.equal((await store.writeFile(files.snapshot, encodeSnapshot('nova versão', { ...meta, name: 'Outra.md' }))).ok, false);
  assert.deepEqual(await fs.readFile(files.snapshot), original);
  assert.equal((await readRecovery(dir, 'nota-1234')).meta.name, 'Nota.md');
  assert.deepEqual(await fs.readdir(dir), ['nota-1234.snapshot']);
});
