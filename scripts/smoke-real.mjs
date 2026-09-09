#!/usr/bin/env node
/**
 * Real-deck validation: runs the whole pipeline against a genuine
 * PowerPoint-produced .pptx (copied to fixtures/generated — the user's file
 * is never modified). Validates load, read, anchor edit, save round-trip,
 * and byte-identical untouched parts.
 *
 * Usage: node scripts/smoke-real.mjs [path-to-real.pptx]
 */
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';
import { SessionManager } from '../packages/pptx-mcp/dist/session/manager.js';
import { toolCatalog } from '../packages/pptx-mcp/dist/tool_registry.js';

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const sourceReal = process.argv[2] ?? 'C:\\Users\\Deano\\Downloads\\test-for-mcp.pptx';
if (!fs.existsSync(sourceReal)) {
  console.error(`real deck not found: ${sourceReal}`);
  process.exit(1);
}
const realPath = path.join(root, 'fixtures', 'generated', 'real-copy.pptx');
fs.copyFileSync(sourceReal, realPath);
const savedPath = path.join(root, 'fixtures', 'generated', 'real-saved.pptx');
if (fs.existsSync(savedPath)) fs.rmSync(savedPath);

const sessions = new SessionManager('real-deck-validation');
const ctx = { sessions };
const tool = (name) => {
  const def = toolCatalog.find((t) => t.name === name);
  if (!def) throw new Error(`missing tool ${name}`);
  return def;
};
const checks = [];
function check(name, ok, detail = '') {
  checks.push(ok);
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? ` — ${detail}` : ''}`);
}

const originalBytes = fs.readFileSync(realPath);
const originalZip = await JSZip.loadAsync(originalBytes);

// 1) get_file_status does NOT auto-create sessions (read-only contract)...
const status0 = await tool('get_file_status').handler(ctx, { file_path: realPath });
check('get_file_status reports not-open before first use', status0.open === false);

// 2) Load via outline (auto-creates the session) + outline
const outline = await tool('get_outline').handler(ctx, { file_path: realPath });
check('real deck loads', outline.slides.length > 0, `${outline.slides.length} slides, ${Object.keys(originalZip.files).length} zip entries`);
const status = await tool('get_file_status').handler(ctx, { file_path: realPath });
check('session now open with slide count', status.open === true && status.slide_count === outline.slides.length, `slides=${status.slide_count}`);
console.log('   titles:', JSON.stringify(outline.slides.map((s) => s.title).slice(0, 5)));

// 2) read_file
const read = await tool('read_file').handler(ctx, { file_path: realPath });
check('read_file returns entries', read.entries.length > 0, `${read.entries.length} entries, has_more=${read.has_more}`);

// 3) find a paragraph with real text and edit it (append marker via replace of last word is risky;
//    instead: pick any paragraph and use replace on its exact full text → append " [edited]")
const candidate = read.entries.find((e) => e.anchor && e.text.replace(/<[^>]+>/g, '').trim().length >= 8 && !e.notes);
check('found editable paragraph', Boolean(candidate), candidate ? candidate.text.slice(0, 60) : '');
const plainText = candidate.text.replace(/<[^>]+>/g, '');
const lastWord = plainText.trim().split(/\s+/).pop();
const rep = await tool('replace_text').handler(ctx, {
  file_path: realPath,
  anchor: candidate.anchor,
  old_string: lastWord,
  new_string: `${lastWord}-EDITED`,
  instruction: 'real-deck validation edit',
});
check('replace_text works on real deck', rep.after.includes('-EDITED'), rep.after.slice(0, 80));

// 4) save + verify untouched parts byte-identical
await tool('save').handler(ctx, { file_path: realPath, save_to_local_path: savedPath });
const savedZip = await JSZip.loadAsync(fs.readFileSync(savedPath));
let untouchedIdentical = true;
let touchedDiffer = false;
const editedPart = rep.partPath;
for (const name of Object.keys(originalZip.files)) {
  const a = await originalZip.files[name].async('nodebuffer');
  const bEntry = savedZip.files[name];
  if (!bEntry) {
    untouchedIdentical = false;
    console.log('   MISSING in saved:', name);
    continue;
  }
  const b = await bEntry.async('nodebuffer');
  if (name === editedPart) {
    if (!a.equals(b)) touchedDiffer = true;
  } else if (!a.equals(b)) {
    untouchedIdentical = false;
    console.log('   differs:', name);
  }
}
check('untouched parts byte-identical (real deck)', untouchedIdentical);
check('edited part changed', touchedDiffer);

// 5) reload + verify the edit is present and grep finds it
const readSaved = await tool('read_file').handler({ sessions: new SessionManager('reload') }, { file_path: savedPath });
check('saved real deck contains edit', readSaved.entries.some((e) => e.text.includes('-EDITED')));
const grepSaved = await tool('grep').handler({ sessions: new SessionManager('g') }, { file_path: savedPath, pattern: '-EDITED' });
check('grep finds edit in saved deck', grepSaved.hits.length >= 1);

// 6) compare original vs edited (the redline must show exactly the edit)
const cmp = await tool('compare_decks').handler(ctx, {
  original_path: sourceReal,
  revised_path: savedPath,
});
check('compare shows the single edit', cmp.changes.some((c) => c.kind === 'modified' && c.after.includes('-EDITED')), `${cmp.changes.length} changes`);
check('no phantom changes on edited slide', cmp.changes.filter((c) => c.kind === 'modified').length === 1, JSON.stringify(cmp.changes.map((c) => c.kind)));

sessions.closeAll();
process.exit(checks.every(Boolean) ? 0 : 1);
