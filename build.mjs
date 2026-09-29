import * as esbuild from 'esbuild';
import fs from 'node:fs';
import path from 'node:path';

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

// tabelas de caracteres e fontes padrão que o pdf.js busca ao ler certos PDFs
for (const dir of ['cmaps', 'standard_fonts']) {
  const from = path.join('node_modules', 'pdfjs-dist', dir);
  const to = path.join('dist', 'pdfjs', dir);
  if (fs.existsSync(from) && !fs.existsSync(to)) fs.cpSync(from, to, { recursive: true });
}

if (watch) {
  const ctx = await esbuild.context(options);
  await ctx.watch();
} else {
  const pdfOut = path.join('dist', 'pdf.js');
  const pdfSrc = path.join('node_modules', 'pdfjs-dist', 'package.json');
  // o pacote do pdf.js só muda quando a versão muda
  if (!fs.existsSync(pdfOut) || fs.statSync(pdfOut).mtimeMs < fs.statSync(pdfSrc).mtimeMs || fs.statSync(pdfOut).mtimeMs < fs.statSync(pdfOptions.entryPoints[0]).mtimeMs) {
    await esbuild.build(pdfOptions);
  }
  await esbuild.build(options);
}
