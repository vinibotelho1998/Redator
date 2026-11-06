'use strict';

const { execFile } = require('node:child_process');
const { promisify } = require('node:util');
const exec = promisify(execFile);

async function listFonts() {
  try {
    const { stdout } = await exec('fc-list', ['--format', '%{family[0]}\n'], { timeout: 5000, maxBuffer: 4000000 });
    return [...new Set(stdout.split('\n').filter(Boolean))].sort((a, b) => a.localeCompare(b, 'pt-BR'));
  } catch { return []; }
}

async function recognizeImage(file) {
  try {
    const langs = (await exec('tesseract', ['--list-langs'], { timeout: 5000 })).stdout.split(/\s+/);
    const lang = ['por', 'eng'].filter((l) => langs.includes(l)).join('+');
    if (!lang) return { err: 'Instale os dados de idioma português (por) do Tesseract.' };
    const { stdout } = await exec('tesseract', [file, 'stdout', '-l', lang, 'tsv'], { timeout: 90000, maxBuffer: 16 * 1024 * 1024 });
    const lines = new Map();
    for (const row of stdout.split(/\r?\n/)) {
      const cols = row.split('\t');
      if (cols[0] !== '5' || !cols[11]?.trim()) continue;
      const key = cols.slice(1, 5).join('.');
      if (!lines.has(key)) lines.set(key, []);
      lines.get(key).push([cols.slice(11).join(' '), ...cols.slice(6, 10).map(Number)]);
    }
    return lines.size ? { lines: [...lines.values()] } : { err: 'Nenhum texto reconhecido nesta página.' };
  } catch (err) {
    return { err: err.code === 'ENOENT' ? 'Instale o Tesseract com os dados de português para reconhecer PDF digitalizado.' : err.message };
  }
}

const mimeTypes = {
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  '.doc': 'application/msword', '.pdf': 'application/pdf', '.md': 'text/markdown',
  '.txt': 'text/plain', '.rtf': 'application/rtf', '.odt': 'application/vnd.oasis.opendocument.text',
  '.epub': 'application/epub+zip', '.csv': 'text/csv',
};

async function defaultStatus() {
  const result = {};
  await Promise.all(Object.entries(mimeTypes).map(async ([ext, mime]) => {
    try {
      const { stdout } = await exec('xdg-mime', ['query', 'default', mime], { timeout: 4000 });
      result[ext] = stdout.trim() === 'br.redator.app.desktop';
    } catch { /* ambiente sem xdg-utils */ }
  }));
  return result;
}

async function convertImage(file, ext, width, height) {
  const format = { tif: 'tiff', tiff: 'tiff', wmf: 'wmf', emf: 'emf' }[ext];
  if (!format) return null;
  const w = Math.min(4000, Math.max(0, Math.round(width) || 0));
  const h = Math.min(4000, Math.max(0, Math.round(height) || 0));
  const size = w || h ? (w || '') + 'x' + (h || '') + (w && h ? '!' : '') : '4000x4000>';
  try {
    const { stdout } = await exec('magick', ['-limit', 'memory', '128MiB', '-limit', 'disk', '256MiB',
      format + ':' + file + '[0]', '-resize', size, 'png:-'], { encoding: 'buffer', timeout: 20000, maxBuffer: 64 * 1024 * 1024 });
    return new Uint8Array(stdout);
  } catch { return null; }
}

module.exports = { listFonts, recognizeImage, defaultStatus, convertImage };
