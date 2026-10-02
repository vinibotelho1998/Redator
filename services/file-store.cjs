'use strict';

const fs = require('node:fs');
const path = require('node:path');
const { randomUUID, createHash } = require('node:crypto');

const hashBytes = (data) => createHash('sha256').update(data).digest('hex');
const error = (code, message) => Object.assign(new Error(message || code), { code });

// One queue for every affected path, shared by all windows. A failed operation
// releases its locks, so the next save can still run.
function createFileStore(io = fs.promises) {
  const pending = new Map();
  const keyFor = (p) => process.platform === 'win32' ? path.resolve(p).toLowerCase() : path.resolve(p);
  const within = (base, p) => {
    const rel = path.relative(base, p);
    return rel !== '..' && !rel.startsWith('..' + path.sep) && !path.isAbsolute(rel);
  };

  function enqueue(paths, action) {
    const keys = [...new Set(paths.map(keyFor))];
    const before = keys.map((key) => pending.get(key)).filter(Boolean);
    const job = Promise.all(before.map((p) => p.catch(() => {}))).then(action);
    for (const key of keys) pending.set(key, job);
    const clear = () => { for (const key of keys) if (pending.get(key) === job) pending.delete(key); };
    job.then(clear, clear);
    return job;
  }

  async function stamp(p) {
    try {
      const bytes = await io.readFile(p);
      const stat = await io.stat(p);
      if (!stat.isFile()) throw error('EINVAL', 'O destino não é um arquivo.');
      return { mtime: stat.mtimeMs, size: stat.size, hash: hashBytes(bytes) };
    } catch (err) {
      if (err.code === 'ENOENT') return null;
      throw err;
    }
  }

  function matches(expected, actual) {
    if (expected === null || actual === null) return expected === actual;
    if (expected.hash) return expected.hash === actual.hash;
    return expected.mtime === actual.mtime && expected.size === actual.size;
  }

  async function checkVersion(p, expected) {
    const current = await stamp(p);
    if (!matches(expected, current)) throw Object.assign(error('ECONFLICT'), { current });
  }

  async function ensureAttachmentDirectory(base, dest) {
    const realBase = await io.realpath(base);
    // Check the nearest existing directory before mkdir, including junctions.
    let parent = path.dirname(dest);
    while (true) {
      try {
        const real = await io.realpath(parent);
        if (!within(realBase, real)) throw error('EATTACHMENTPATH', 'O anexo aponta para fora da pasta do documento.');
        break;
      } catch (err) {
        if (err.code !== 'ENOENT') throw err;
        const next = path.dirname(parent);
        if (next === parent) throw err;
        parent = next;
      }
    }
    await io.mkdir(path.dirname(dest), { recursive: true });
    if (!within(realBase, await io.realpath(path.dirname(dest)))) throw error('EATTACHMENTPATH');
    // Do not replace a link that can refer to another document or directory.
    try { if ((await io.lstat(dest)).isSymbolicLink()) throw error('EATTACHMENTPATH'); }
    catch (err) { if (err.code !== 'ENOENT') throw err; }
  }

  async function stage(dest, data) {
    const tmp = path.join(path.dirname(dest), '.' + path.basename(dest) + '.' + randomUUID() + '.redator-tmp');
    let handle;
    try {
      let mode;
      try { mode = (await io.stat(dest)).mode; } catch (err) { if (err.code !== 'ENOENT') throw err; }
      handle = await io.open(tmp, 'wx', mode);
      await handle.writeFile(data);
      await handle.sync();
      const stat = await handle.stat();
      await handle.close();
      handle = null;
      return { dest, tmp, stamp: { mtime: stat.mtimeMs, size: stat.size, hash: hashBytes(data) } };
    } catch (err) {
      if (handle) await handle.close().catch(() => {});
      await io.unlink(tmp).catch(() => {});
      throw err;
    }
  }

  async function writeFile(p, data, extras, options = {}) {
    try {
      if (typeof p !== 'string' || !path.isAbsolute(p)) throw error('EINVAL', 'O destino precisa ser um caminho absoluto.');
      const dest = path.resolve(p);
      const base = path.dirname(dest);
      const bytes = typeof data === 'string' ? Buffer.from(data, 'utf8') : Buffer.from(data);
      if (extras != null && !Array.isArray(extras)) throw error('EINVAL');
      const seen = new Set([keyFor(dest)]);
      const attachments = (extras || []).map((ex) => {
        if (!ex || typeof ex.path !== 'string' || !ex.path || path.isAbsolute(ex.path)) throw error('EATTACHMENTPATH');
        const full = path.resolve(base, ex.path);
        if (!within(base, full) || seen.has(keyFor(full))) throw error('EATTACHMENTPATH');
        seen.add(keyFor(full));
        return { dest: full, data: Buffer.from(ex.data) };
      });
      return await enqueue([dest, ...attachments.map((a) => a.dest)], async () => {
        const staged = [];
        const backups = [];
        let committed = false;
        try {
          const expected = Object.prototype.hasOwnProperty.call(options, 'expected') ? options.expected : await stamp(dest);
          await checkVersion(dest, expected);
          for (const a of attachments) {
            await ensureAttachmentDirectory(base, a.dest);
            staged.push(await stage(a.dest, a.data));
          }
          const document = await stage(dest, bytes);
          staged.push(document);
          // A copy of each old attachment permits rollback if publication fails.
          for (const a of staged.slice(0, -1)) {
            const backup = a.tmp + '.backup';
            let existed = true;
            try { await io.copyFile(a.dest, backup, fs.constants.COPYFILE_EXCL); }
            catch (err) { if (err.code === 'ENOENT') existed = false; else throw err; }
            backups.push({ dest: a.dest, backup: existed ? backup : null, published: false });
          }
          await checkVersion(dest, expected);
          for (let i = 0; i < backups.length; i++) {
            await io.rename(staged[i].tmp, staged[i].dest);
            backups[i].published = true;
          }
          // Recheck after attachment I/O; never truncate the original on failure.
          await checkVersion(dest, expected);
          await io.rename(document.tmp, dest);
          committed = true;
          return { ok: true, stamp: document.stamp };
        } catch (err) {
          const rollbackErrors = [];
          for (const b of backups.slice().reverse()) {
            if (!b.published) continue;
            try {
              if (b.backup) { await io.rename(b.backup, b.dest); b.backup = null; }
              else await io.unlink(b.dest);
            } catch (rollback) { rollbackErrors.push({ path: b.dest, backup: b.backup, error: rollback.code || rollback.message }); }
          }
          return { ok: false, error: err.code || err.message, ...(err.current !== undefined ? { current: err.current } : {}), ...(rollbackErrors.length ? { rollbackErrors } : {}) };
        } finally {
          for (const a of staged) await io.unlink(a.tmp).catch(() => {});
          for (const b of backups) {
            // Preserve the old file if a rollback failed, for manual recovery.
            if (b.backup && (committed || !b.published)) await io.unlink(b.backup).catch(() => {});
          }
        }
      });
    } catch (err) {
      return { ok: false, error: err.code || err.message };
    }
  }

  return { writeFile, stamp, enqueue };
}

module.exports = { createFileStore, hashBytes };
