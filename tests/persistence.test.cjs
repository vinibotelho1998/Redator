'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const { createFileStore } = require('../services/file-store.cjs');

async function fixture(t) {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'redator-persistence-'));
  t.after(() => fs.rm(dir, { recursive: true, force: true }));
  const dest = path.join(dir, 'nota.md');
  await fs.writeFile(dest, 'original');
  return { dir, dest };
}

test('publica bytes completos e devolve a versão gravada', async (t) => {
  const { dir, dest } = await fixture(t);
  const store = createFileStore();
  const result = await store.writeFile(dest, 'versão nova', [], { expected: await store.stamp(dest) });
  assert.equal(result.ok, true);
  assert.equal(await fs.readFile(dest, 'utf8'), 'versão nova');
  assert.deepEqual(result.stamp, await store.stamp(dest));
  assert.deepEqual(await fs.readdir(dir), ['nota.md']);
});

test('rename bloqueado preserva o original e limpa temporário', async (t) => {
  const { dir, dest } = await fixture(t);
  const store = createFileStore({ ...fs, rename: async () => { throw Object.assign(new Error('bloqueado'), { code: 'EPERM' }); } });
  const result = await store.writeFile(dest, 'novo');
  assert.equal(result.error, 'EPERM');
  assert.equal(await fs.readFile(dest, 'utf8'), 'original');
  assert.deepEqual(await fs.readdir(dir), ['nota.md']);
});

test('erro de flush não substitui o original', async (t) => {
  const { dir, dest } = await fixture(t);
  const store = createFileStore({ ...fs, open: async (...args) => {
    const h = await fs.open(...args);
    return { writeFile: h.writeFile.bind(h), close: h.close.bind(h), stat: h.stat.bind(h), sync: async () => { throw Object.assign(new Error('disco'), { code: 'EIO' }); } };
  } });
  assert.equal((await store.writeFile(dest, 'novo')).error, 'EIO');
  assert.equal(await fs.readFile(dest, 'utf8'), 'original');
  assert.deepEqual(await fs.readdir(dir), ['nota.md']);
});

test('salvamentos concorrentes usam temporários distintos e respeitam a versão esperada', async (t) => {
  const { dir, dest } = await fixture(t);
  const staged = [];
  const store = createFileStore({ ...fs, open: async (p, ...args) => { staged.push(p); return fs.open(p, ...args); } });
  const expected = await store.stamp(dest);
  const results = await Promise.all([
    store.writeFile(dest, 'primeiro', [], { expected }),
    store.writeFile(dest, 'segundo', [], { expected }),
  ]);
  assert.equal(results[0].ok, true);
  assert.equal(results[1].error, 'ECONFLICT');
  assert.equal(await fs.readFile(dest, 'utf8'), 'primeiro');
  assert.equal((await store.writeFile(dest, 'terceiro')).ok, true);
  assert.equal(new Set(staged).size, staged.length);
  assert.deepEqual(await fs.readdir(dir), ['nota.md']);
});

test('detecta edição externa mesmo com tamanho e mtime preservados', async (t) => {
  const { dest } = await fixture(t);
  const store = createFileStore();
  const expected = await store.stamp(dest);
  const stat = await fs.stat(dest);
  await fs.writeFile(dest, 'externo!');
  await fs.utimes(dest, stat.atime, stat.mtime);
  const result = await store.writeFile(dest, 'local', [], { expected });
  assert.equal(result.error, 'ECONFLICT');
  assert.equal(result.current.hash, (await store.stamp(dest)).hash);
  assert.equal(await fs.readFile(dest, 'utf8'), 'externo!');
});

test('detecta exclusão externa e permite recriar após decisão explícita', async (t) => {
  const { dest } = await fixture(t);
  const store = createFileStore();
  const expected = await store.stamp(dest);
  await fs.unlink(dest);
  const result = await store.writeFile(dest, 'local', [], { expected });
  assert.equal(result.error, 'ECONFLICT');
  assert.equal(result.current, null);
  assert.equal((await store.writeFile(dest, 'local', [], { expected: null })).ok, true);
});

test('reverte anexos publicados se o documento não puder ser substituído', async (t) => {
  const { dir, dest } = await fixture(t);
  await fs.writeFile(path.join(dir, 'figura.png'), 'imagem antiga');
  const store = createFileStore({ ...fs, rename: async (from, to) => {
    if (to === dest) throw Object.assign(new Error('ocupado'), { code: 'EBUSY' });
    return fs.rename(from, to);
  } });
  const result = await store.writeFile(dest, 'novo', [
    { path: 'figura.png', data: Buffer.from('imagem nova') },
    { path: 'extra.png', data: Buffer.from('imagem extra') },
  ]);
  assert.equal(result.error, 'EBUSY');
  assert.equal(await fs.readFile(dest, 'utf8'), 'original');
  assert.equal(await fs.readFile(path.join(dir, 'figura.png'), 'utf8'), 'imagem antiga');
  assert.deepEqual((await fs.readdir(dir)).sort(), ['figura.png', 'nota.md']);
});

test('erro ao preparar anexo não publica o documento', async (t) => {
  const { dir, dest } = await fixture(t);
  await fs.writeFile(path.join(dir, 'imagens'), 'isto é um arquivo');
  const result = await createFileStore().writeFile(dest, 'novo', [{ path: 'imagens/a.png', data: Buffer.from('imagem') }]);
  assert.equal(result.ok, false);
  assert.equal(await fs.readFile(dest, 'utf8'), 'original');
});

test('recusa anexos fora da pasta, duplicados ou com o caminho do documento', async (t) => {
  const { dir, dest } = await fixture(t);
  const store = createFileStore();
  for (const relative of ['../escape.png', dest, 'nota.md']) {
    assert.equal((await store.writeFile(dest, 'novo', [{ path: relative, data: Buffer.from('imagem') }])).error, 'EATTACHMENTPATH');
  }
  const dup = [{ path: 'figura.png', data: Buffer.from('a') }, { path: './figura.png', data: Buffer.from('b') }];
  assert.equal((await store.writeFile(dest, 'novo', dup)).error, 'EATTACHMENTPATH');
  assert.deepEqual(await fs.readdir(dir), ['nota.md']);
});

test('recusa anexo em junction ou link para fora da pasta', async (t) => {
  const { dir, dest } = await fixture(t);
  const outside = await fs.mkdtemp(path.join(os.tmpdir(), 'redator-outside-'));
  t.after(() => fs.rm(outside, { recursive: true, force: true }));
  await fs.symlink(outside, path.join(dir, 'anexos'), process.platform === 'win32' ? 'junction' : 'dir');
  const result = await createFileStore().writeFile(dest, 'novo', [{ path: 'anexos/a.png', data: Buffer.from('imagem') }]);
  assert.equal(result.error, 'EATTACHMENTPATH');
  assert.deepEqual(await fs.readdir(outside), []);
  assert.equal(await fs.readFile(dest, 'utf8'), 'original');
});

test('falha de rollback preserva backup e informa seu caminho para recuperação', async (t) => {
  const { dir, dest } = await fixture(t);
  const image = path.join(dir, 'figura.png');
  await fs.writeFile(image, 'imagem antiga');
  const store = createFileStore({ ...fs, rename: async (from, to) => {
    if (to === dest || from.endsWith('.backup')) throw Object.assign(new Error('ocupado'), { code: 'EBUSY' });
    return fs.rename(from, to);
  } });
  const result = await store.writeFile(dest, 'novo', [{ path: 'figura.png', data: Buffer.from('imagem nova') }]);
  assert.equal(result.ok, false);
  assert.equal(result.rollbackErrors.length, 1);
  assert.equal(result.rollbackErrors[0].path, image);
  assert.equal(await fs.readFile(result.rollbackErrors[0].backup, 'utf8'), 'imagem antiga');
  assert.equal(await fs.readFile(dest, 'utf8'), 'original');
  assert.equal(await fs.readFile(image, 'utf8'), 'imagem nova');
  assert.equal((await fs.readdir(dir)).filter((name) => name.endsWith('.redator-tmp')).length, 0);
});

test('mudança externa durante preparação de anexos impede publicação de todos os arquivos', async (t) => {
  const { dir, dest } = await fixture(t);
  let changed = false;
  const store = createFileStore({ ...fs, copyFile: async (from, to, flags) => {
    if (!changed) { changed = true; await fs.writeFile(dest, 'externo!'); }
    return fs.copyFile(from, to, flags);
  } });
  const result = await store.writeFile(dest, 'novo', [{ path: 'figura.png', data: Buffer.from('imagem') }]);
  assert.equal(result.error, 'ECONFLICT');
  assert.equal(await fs.readFile(dest, 'utf8'), 'externo!');
  assert.deepEqual(await fs.readdir(dir), ['nota.md']);
});

test('limpeza enfileirada aguarda publicação pendente e não permite ressuscitar snapshot', async (t) => {
  const { dest } = await fixture(t);
  let release;
  const ready = new Promise((resolve) => { release = resolve; });
  let entered;
  const staged = new Promise((resolve) => { entered = resolve; });
  const store = createFileStore({ ...fs, rename: async (...args) => { entered(); await ready; return fs.rename(...args); } });
  const saving = store.writeFile(dest, 'snapshot novo');
  await staged;
  const clearing = store.enqueue([dest], () => fs.unlink(dest));
  release();
  assert.equal((await saving).ok, true);
  await clearing;
  assert.equal(await store.stamp(dest), null);
});
