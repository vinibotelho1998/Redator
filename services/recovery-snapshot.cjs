'use strict';

const { hashBytes } = require('./file-store.cjs');
const fs = require('node:fs');
const path = require('node:path');
const MAGIC = Buffer.from('REDATOR-RECOVERY-1\n');
const MAX_HEADER = 64 * 1024;
const RECOVER_ID = /^[a-z0-9-]{4,40}$/;

function metadata(value) {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error('Metadados de recuperação inválidos.');
  const meta = { ...value };
  for (const field of ['name', 'path', 'format', 'kind']) {
    if (meta[field] != null && typeof meta[field] !== 'string') throw new Error('Metadados de recuperação inválidos.');
  }
  if (meta.kind && !['docx', 'md', 'txt', 'csv'].includes(meta.kind)) throw new Error('Formato de recuperação inválido.');
  if (meta.time != null && (!Number.isFinite(meta.time) || meta.time < 0)) throw new Error('Data de recuperação inválida.');
  if (meta.text != null && (typeof meta.text !== 'object' || Array.isArray(meta.text))) throw new Error('Codificação de recuperação inválida.');
  return meta;
}

function encodeSnapshot(data, meta) {
  const bytes = Buffer.from(data);
  const header = Buffer.from(JSON.stringify({ ...metadata(meta), time: Date.now(), hash: hashBytes(bytes) }), 'utf8');
  if (header.length > MAX_HEADER) throw new Error('Metadados de recuperação grandes demais.');
  const length = Buffer.alloc(4);
  length.writeUInt32LE(header.length);
  return Buffer.concat([MAGIC, length, header, bytes]);
}

function decodeSnapshot(packet) {
  const bytes = Buffer.from(packet);
  const start = MAGIC.length + 4;
  if (bytes.length < start || !bytes.subarray(0, MAGIC.length).equals(MAGIC)) throw new Error('Cópia de recuperação inválida.');
  const size = bytes.readUInt32LE(MAGIC.length);
  if (!size || size > MAX_HEADER || start + size > bytes.length) throw new Error('Cópia de recuperação incompleta.');
  const meta = metadata(JSON.parse(bytes.subarray(start, start + size).toString('utf8')));
  const data = bytes.subarray(start + size);
  if (!meta || typeof meta !== 'object' || Array.isArray(meta) || meta.hash !== hashBytes(data)) throw new Error('Cópia de recuperação corrompida.');
  return { meta, data };
}

function recoveryFiles(directory, id) {
  if (typeof id !== 'string' || !RECOVER_ID.test(id)) throw new Error('Identificador de recuperação inválido.');
  return { data: path.join(directory, id + '.dados'), meta: path.join(directory, id + '.json'), snapshot: path.join(directory, id + '.snapshot') };
}

// Prefer the indivisible snapshot. A valid pair from an older Redator version
// remains usable if a newer snapshot is missing or damaged.
async function readRecovery(directory, id, io = fs.promises) {
  const files = recoveryFiles(directory, id);
  try { return decodeSnapshot(await io.readFile(files.snapshot)); }
  catch (snapshotError) {
    try {
      const meta = metadata(JSON.parse(await io.readFile(files.meta, 'utf8')));
      const data = await io.readFile(files.data);
      if (!data.length) throw new Error('Cópia de recuperação vazia.');
      return { meta, data };
    } catch { throw snapshotError; }
  }
}

function scanRecovery(directory, io = fs) {
  let names;
  try { names = io.readdirSync(directory); } catch { return []; }
  const ids = new Set(names.map((name) => /^([a-z0-9-]+)\.(snapshot|json)$/.exec(name)).filter((match) => match && RECOVER_ID.test(match[1])).map((match) => match[1]));
  const found = [];
  for (const id of ids) {
    const files = recoveryFiles(directory, id);
    let meta;
    try { meta = decodeSnapshot(io.readFileSync(files.snapshot)).meta; }
    catch {
      try {
        meta = metadata(JSON.parse(io.readFileSync(files.meta, 'utf8')));
        if (!io.statSync(files.data).isFile() || !io.statSync(files.data).size) continue;
      } catch { continue; }
    }
    found.push({ id, name: meta.name || 'Documento', path: meta.path || null, time: meta.time || 0 });
  }
  // Preserve incomplete/corrupt files so scanning cannot destroy user data.
  return found.sort((a, b) => b.time - a.time);
}

module.exports = { encodeSnapshot, decodeSnapshot, readRecovery, scanRecovery, recoveryFiles, RECOVER_ID };
