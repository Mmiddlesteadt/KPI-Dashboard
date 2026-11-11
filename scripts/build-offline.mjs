import { build } from 'esbuild';
import postcss from 'postcss';
import tailwind from '@tailwindcss/postcss';
import { readFile, writeFile } from 'node:fs/promises';

// Bundle the original client application; no hosted runtime or CDN is needed.
const logo = `data:image/png;base64,${(await readFile('public/wtdc-logo.png')).toString('base64')}`;
const result = await build({
  entryPoints: ['app/offline-entry.tsx'], bundle: true, write: false,
  minify: true, format: 'iife', jsx: 'automatic',
  define: {'process.env.NODE_ENV': '"production"'},
});
const js = result.outputFiles[0].text.replaceAll('/wtdc-logo.png', logo).replaceAll('</script', '<\\/script');
const styles = await postcss([tailwind()]).process(await readFile('app/globals.css', 'utf8'), {from: 'app/globals.css'});
const html = `<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>WTDC Customer Reporting Dashboard</title>
<style>${styles.css}</style></head><body class="antialiased"><div id="root"></div>
<script>${js}</script></body></html>`;
await writeFile('index.html', html);
console.log(`Built standalone index.html (${Math.round(Buffer.byteLength(html)/1024)} KB).`);
