// Empacota para a plataforma nativa em out/Redator-<plataforma>-<arquitetura>.
// Usa o zip do Electron que já estiver no cache local; sem ele, o packager baixa.
import { packager } from '@electron/packager';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const pkg = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
const electronVersion = JSON.parse(fs.readFileSync(path.join(root, 'node_modules', 'electron', 'package.json'), 'utf8')).version;

const platform = process.platform;
if (!['win32', 'linux'].includes(platform)) throw new Error('Empacotamento disponível em Windows e Linux.');
const arch = platform === 'win32' ? 'x64' : process.arch;
let zipDir = null;
const cacheRoot = process.env.electron_config_cache || (platform === 'win32' ? path.join(process.env.LOCALAPPDATA || '', 'electron', 'Cache') : path.join(process.env.XDG_CACHE_HOME || path.join(process.env.HOME, '.cache'), 'electron'));
try {
  for (const d of fs.readdirSync(cacheRoot)) {
    const full = path.join(cacheRoot, d);
    if (fs.existsSync(path.join(full, 'electron-v' + electronVersion + '-' + platform + '-' + arch + '.zip'))) zipDir = full;
  }
} catch { /* sem cache local */ }

const keep = new Set(['main.js', 'preload.js', 'package.json']);
const paths = await packager({
  dir: root,
  out: path.join(root, 'out'),
  name: 'Redator',
  executableName: 'Redator',
  platform,
  arch,
  electronVersion,
  ...(zipDir ? { electronZipDir: zipDir } : {}),
  icon: path.join(root, 'app', 'icons', platform === 'win32' ? 'redator.ico' : 'redator.png'),
  overwrite: true,
  asar: true,
  prune: false,
  appVersion: pkg.version,
  buildVersion: pkg.version,
  appCopyright: 'Copyright (C) Vinicius Botelho',
  win32metadata: { CompanyName: 'Vinicius Botelho', FileDescription: 'Redator - editor de documentos', ProductName: 'Redator', InternalName: 'Redator', OriginalFilename: 'Redator.exe' },
  ignore: (p) => {
    if (!p) return false;
    const top = p.split('/')[1];
    if (keep.has(top)) return false;
    if (top === 'services') return false;
    if (top === 'app') return false;
    if (top === 'dist') return /\.map$/.test(p);
    return true;
  },
});
// o Chromium traz traduções da interface dele para dezenas de idiomas; o Redator só usa estas
const KEEP_LOCALES = new Set(['en-US.pak', 'pt-BR.pak', 'pt-PT.pak']);
for (const out of paths) {
  const dir = path.join(out, 'locales');
  if (!fs.existsSync(dir)) continue;
  for (const f of fs.readdirSync(dir)) if (!KEEP_LOCALES.has(f)) fs.rmSync(path.join(dir, f), { force: true });
}
console.log('pacote gerado em', paths.join(', '));
