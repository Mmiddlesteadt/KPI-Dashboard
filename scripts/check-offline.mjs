import assert from 'node:assert/strict';
import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { Script } from 'node:vm';
import { build } from 'esbuild';
import { renderToStaticMarkup } from 'react-dom/server';
import { createElement } from 'react';

const html = await readFile('index.html', 'utf8');
assert.ok(!/<script[^>]+src=|<link[^>]+href=/i.test(html), 'No external script or stylesheet');
assert.ok(html.includes('data:image/png;base64,'), 'Logo embedded');
const script = html.match(/<script>([\s\S]*)<\/script>/)[1];
new Script(script);
assert.ok(!html.includes('@import "tailwindcss"'), 'Styles compiled');

// Render the unchanged component without a browser to check every blank section.
await mkdir('output', {recursive:true});
await build({entryPoints:['app/page.tsx'], bundle:true, platform:'node',
  packages:'external', format:'esm', jsx:'automatic', outfile:'output/render-check.mjs'});
const {default: Dashboard} = await import('../output/render-check.mjs');
const markup = renderToStaticMarkup(createElement(Dashboard));
for (const label of ['Customer performance dashboard', 'Choose Excel file',
  'Ending inventory', 'Creation-to-delivery service performance',
  'Inventory older than one year', 'Inbound and outbound cases by month and year',
  'Warehouse movements connected by part number and receipt']) {
  assert.ok(markup.toLowerCase().includes(label.toLowerCase()), label);
}
await writeFile('output/rendered-blank.html', markup);
console.log('PASS: standalone bundle syntax, embedded assets, compiled styles and all reference sections.');
