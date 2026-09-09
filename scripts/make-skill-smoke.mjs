#!/usr/bin/env node
/**
 * Skill-stack smoke test: generate a small deck with the GLOBALLY installed
 * pptxgenjs (per the official pptx skill), then verify safe-pptx loads,
 * reads, edits and saves it. Also serves as the generator-library resolution
 * pattern for make-torture-deck.mjs.
 */
import { createRequire } from 'node:module';
import { execSync } from 'node:child_process';
import path from 'node:path';
import fs from 'node:fs';
import { fileURLToPath } from 'node:url';

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const outPath = path.join(root, 'fixtures', 'generated', 'skill-smoke.pptx');

function requireGlobal(moduleName) {
  const req = createRequire(import.meta.url);
  try {
    return req(moduleName);
  } catch {
    const globalRoot = execSync('npm root -g').toString().trim();
    return req(path.join(globalRoot, moduleName));
  }
}

const pptxgen = requireGlobal('pptxgenjs');
const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';
pres.author = 'safe-pptx skill smoke';
pres.title = 'Skill smoke deck';

const s1 = pres.addSlide();
s1.background = { color: '0F172A' };
s1.addText('Skill stack smoke', { x: 0.5, y: 2.6, w: 12.33, h: 1.2, fontSize: 44, bold: true, color: 'F8FAFC', fontFace: 'Arial' });
s1.addText('Generated with globally installed pptxgenjs', { x: 0.5, y: 3.8, w: 12.33, h: 0.6, fontSize: 18, color: '94A3B8', fontFace: 'Arial' });
s1.addNotes('Smoke deck notes');

const s2 = pres.addSlide();
s2.addText([{ text: 'Chart slide ', options: { bold: true } }, { text: 'with native bar chart', options: { italic: true } }], { x: 0.5, y: 0.4, w: 12, h: 0.6, fontSize: 24, fontFace: 'Arial' });
s2.addChart(pres.charts.BAR, [{ name: 'Sales', labels: ['Q1', 'Q2', 'Q3'], values: [10, 20, 30] }], {
  x: 0.5, y: 1.2, w: 8, h: 4.5, barDir: 'col', showValue: true, dataLabelPosition: 'outEnd',
});

fs.mkdirSync(path.dirname(outPath), { recursive: true });
await pres.writeFile({ fileName: outPath });
console.log(`generated ${outPath} (${fs.statSync(outPath).size} bytes)`);

// Now: safe-pptx round-trip
const { SessionManager } = await import('../packages/pptx-mcp/dist/session/manager.js');
const { toolCatalog } = await import('../packages/pptx-mcp/dist/tool_registry.js');
const sessions = new SessionManager('skill-smoke');
const ctx = { sessions };
const tool = (n) => toolCatalog.find((t) => t.name === n);

const read = await tool('read_file').handler(ctx, { file_path: outPath });
const hasTitle = read.entries.some((e) => e.text.includes('Skill stack smoke'));
const hasNotes = read.entries.some((e) => e.notes && e.text.includes('Smoke deck notes'));
console.log(`${hasTitle ? 'PASS' : 'FAIL'}  safe-pptx reads pptxgenjs deck (title)`);
console.log(`${hasNotes ? 'PASS' : 'FAIL'}  safe-pptx reads pptxgenjs notes`);

const anchor = read.entries.find((e) => e.anchor && e.text.includes('Skill stack smoke')).anchor;
await tool('replace_text').handler(ctx, {
  file_path: outPath, anchor, old_string: 'smoke', new_string: 'SMOKE', instruction: 'skill smoke edit',
});
await tool('save').handler(ctx, { file_path: outPath });
const reread = await tool('read_file').handler({ sessions: new SessionManager('smoke2') }, { file_path: outPath });
const ok = reread.entries.some((e) => e.text.includes('Skill stack SMOKE'));
console.log(`${ok ? 'PASS' : 'FAIL'}  edit + save + reload round-trip`);
sessions.closeAll();
process.exit(hasTitle && hasNotes && ok ? 0 : 1);
