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
for (const label of ['Customer performance dashboard', 'Choose Excel file', 'Client / Excel format']) {
  assert.ok(markup.includes(label), label);
}
for (const hidden of ['id="overview"', 'id="inventory"', 'id="orders"', 'id="aging"', 'id="volumes"', 'id="details"', 'class="section-nav"', 'class="sharing-guide"']) {
  assert.ok(!markup.includes(hidden), `Blank dashboard hides ${hidden}`);
}
await writeFile('output/rendered-blank.html', markup);
console.log('PASS: standalone bundle syntax, embedded assets, compiled styles and blank upload screen.');

// An exported report must explain calculations without exposing worksheet names.
globalThis.__WTDC_SNAPSHOT__ = {report: {
  overview: {months: ['Sep 2026'], rows: [{label: 'Reported value', values: [42]}]},
  fileName: 'private-source.xlsx', customer: 'Example customer', period: 'September 2026',
  reportDate: '2026-09-30T12:00:00.000Z', generatedDate: '',
  movements: [{part: '123456', description: 'Example item', start: 10, inbound: 5, outbound: 3, end: 12}],
  orders: [], aging: [], details: [], inboundVolumes: [], outboundVolumes: [],
  accidentDays: 8084, accidentAsOf: '9/30/2026', co2: [],
  recognizedSheets: ['In Out Summary', 'KPI', 'Overview'],
}};
const {default: ExportedDashboard} = await import('../output/render-check.mjs?export-check');
const exportedMarkup = renderToStaticMarkup(createElement(ExportedDashboard));
for (const forbidden of ['Local processing', 'Source:', 'Overview worksheet', 'In Out Summary', 'KPI worksheet',
  'Inventory &gt;1 Year worksheet', 'Comparison worksheet', 'Days Accident Free sheet', 'private-source.xlsx']) {
  assert.ok(!exportedMarkup.includes(forbidden), `Export hides ${forbidden}`);
}
assert.ok(!/\bworksheets?\b/i.test(exportedMarkup), 'Export contains no worksheet references');
assert.ok(exportedMarkup.includes('aria-label="How this is calculated"'), 'Calculation icon has an accessible explanation');
assert.ok(exportedMarkup.includes('aria-label="About this information"'), 'Direct values use informational wording');
assert.ok(exportedMarkup.includes('net change is ending minus starting inventory.'), 'Headline calculation explained');
assert.ok(exportedMarkup.includes('years equals days divided by 365.25'), 'Safety conversion explained');
assert.ok(!markup.includes('Source:'), 'Blank dashboard hides report guidance');
delete globalThis.__WTDC_SNAPSHOT__;
console.log('PASS: exported view hides worksheet references and distinguishes calculation help from information.');

for (const [name, view, expected] of [['internal blank', markup, 0], ['exported', exportedMarkup, 9]]) {
  assert.equal((view.match(/class="calculation-help"/g) ?? []).length, expected, `${name}: one icon per report section`);
  for (const section of view.matchAll(/<section\b[^>]*>([\s\S]*?)<\/section>/g)) {
    assert.ok((section[1].match(/class="calculation-help"/g) ?? []).length <= 1, `${name}: no section has duplicate help icons`);
  }
}
console.log('PASS: each report section has at most one information icon.');

for (const view of [markup, exportedMarkup]) {
  for (const match of view.matchAll(/<div class="calculation-body">([\s\S]*?)<\/div>/g)) {
    const explanation = match[1].replace(/<p><strong>Source:<\/strong>[\s\S]*?<\/p>/g, '').replace(/<[^>]*>/g, '');
    const sentences = explanation.split(/(?<=[.!?])\s+(?=[A-Z])/).filter(Boolean);
    assert.ok(sentences.length <= 5, 'Help text stays within five sentences');
    assert.ok(explanation.split(/\s+/).length <= 90, 'Help text stays brief');
  }
}
console.log('PASS: section explanations are brief and contain at most five sentences.');
