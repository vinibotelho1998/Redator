import * as esbuild from 'esbuild';

const watch = process.argv.includes('--watch');

const options = {
  entryPoints: ['src/app.js'],
  bundle: true,
  outfile: 'dist/renderer.js',
  format: 'iife',
  target: 'chrome130',
  sourcemap: true,
  minify: process.argv.includes('--minify'),
  logLevel: 'info',
  legalComments: 'none',
};

if (watch) {
  const ctx = await esbuild.context(options);
  await ctx.watch();
} else {
  await esbuild.build(options);
}
