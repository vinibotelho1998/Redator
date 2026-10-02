import * as esbuild from 'esbuild';
import fs from 'node:fs';
import path from 'node:path';
import { createHash } from 'node:crypto';

const watch = process.argv.includes('--watch');
const minify = process.argv.includes('--minify');

const common = {
  bundle: true,
  format: 'iife',
  target: 'chrome130',
  sourcemap: true,
  minify,
  logLevel: 'info',
  legalComments: 'none',
};

const options = { ...common, entryPoints: ['src/app.js'], outfile: 'dist/renderer.js' };
// o pdf.js fica num pacote à parte, carregado só quando um PDF é aberto
const pdfOptions = { ...common, entryPoints: ['src/pdf/pdfjs-entry.js'], outfile: 'dist/pdf.js', minify: true, sourcemap: false };
const pdfPackage = JSON.parse(fs.readFileSync(path.join('node_modules', 'pdfjs-dist', 'package.json'), 'utf8'));
const pdfStamp = pdfPackage.version + ':' + createHash('sha256').update(fs.readFileSync(pdfOptions.entryPoints[0])).digest('hex');
const stampPath = path.join('dist', 'pdfjs.version');
let previousStamp = '';
try { previousStamp = fs.readFileSync(stampPath, 'utf8'); } catch { /* primeiro build */ }
const assetsVersionPath = path.join('dist', 'pdfjs', 'version');
let previousAssetsVersion = '';
try { previousAssetsVersion = fs.readFileSync(assetsVersionPath, 'utf8'); } catch { /* primeiro build */ }

// tabelas de caracteres e fontes padrão que o pdf.js busca ao ler certos PDFs
for (const dir of ['cmaps', 'standard_fonts']) {
  const from = path.join('node_modules', 'pdfjs-dist', dir);
  const to = path.join('dist', 'pdfjs', dir);
  if (fs.existsSync(from) && (previousAssetsVersion !== pdfPackage.version || !fs.existsSync(to))) {
    // O destino é sempre uma das duas pastas de recursos dentro de dist/pdfjs.
    const assetsRoot = path.resolve('dist', 'pdfjs');
    const target = path.resolve(to);
    if (!target.startsWith(assetsRoot + path.sep)) throw new Error('Destino de recursos PDF inválido');
    fs.rmSync(target, { recursive: true, force: true });
    fs.cpSync(from, target, { recursive: true });
  }
}
fs.mkdirSync(path.dirname(assetsVersionPath), { recursive: true });
fs.writeFileSync(assetsVersionPath, pdfPackage.version);

// O bundle PDF também precisa existir no modo watch. O carimbo evita reutilizar
// código de outra versão após reinstalar dependências ou restaurar timestamps.
if (!fs.existsSync(pdfOptions.outfile) || previousStamp !== pdfStamp) {
  await esbuild.build(pdfOptions);
  fs.writeFileSync(stampPath, pdfStamp);
}

if (watch) {
  const ctx = await esbuild.context(options);
  await ctx.watch();
} else {
  await esbuild.build(options);
}
