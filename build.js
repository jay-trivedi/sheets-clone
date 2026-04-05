import * as esbuild from 'esbuild';

const watch = process.argv.includes('--watch');

const buildOpts = {
  entryPoints: ['src/index.js'],
  bundle: true,
  outfile: 'dist/sheets.js',
  format: 'iife',
  globalName: 'SheetsLib',
  sourcemap: true,
  minify: !watch,
  target: ['es2020'],
  banner: {
    js: '/* Sheets - A Google Sheets Clone | MIT License */',
  },
};

if (watch) {
  const ctx = await esbuild.context(buildOpts);
  await ctx.watch();
  console.log('Watching for changes...');
} else {
  await esbuild.build(buildOpts);
  console.log('Built dist/sheets.js');

  // Also build ESM
  await esbuild.build({
    ...buildOpts,
    outfile: 'dist/sheets.esm.js',
    format: 'esm',
    globalName: undefined,
  });
  console.log('Built dist/sheets.esm.js');
}
