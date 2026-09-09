#!/usr/bin/env node
/**
 * Phase-5 smoke test: compare_decks — modified/added/removed paragraphs, slide
 * add/remove detection, markdown report, annotated copy (red runs).
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';
import { SessionManager } from '../packages/pptx-mcp/dist/session/manager.js';
import { toolCatalog } from '../packages/pptx-mcp/dist/tool_registry.js';

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const gen = path.join(root, 'fixtures', 'generated');
const originalPath = path.join(gen, 'cmp-a.pptx');
const revisedPath = path.join(gen, 'cmp-b.pptx');
const reportPath = path.join(gen, 'cmp-report.md');
const annotatedPath = path.join(gen, 'cmp-annotated.pptx');
fs.copyFileSync(path.join(gen, 'sample.pptx'), originalPath);
fs.copyFileSync(path.join(gen, 'sample.pptx'), revisedPath);
for (const stale of [reportPath, annotatedPath]) if (fs.existsSync(stale)) fs.rmSync(stale);

const sessions = new SessionManager('smoke-ai');
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
const anchorOf = async (file, needle) => {
  const read = await tool('read_file').handler(ctx, { file_path: file });
  const e = read.entries.find((e) => e.anchor && e.text.includes(needle));
  return e?.anchor;
};

// Edit the revised deck: modify a paragraph, append one at the end of the shape,
// clear one to empty; add a slide; edit notes.
const a1 = await anchorOf(revisedPath, '12%');
await tool('replace_text').handler(ctx, { file_path: revisedPath, anchor: a1, old_string: '12%', new_string: '18%', instruction: 'update figure' });
const a1c = await anchorOf(revisedPath, 'Churn held steady');
await tool('insert_paragraph').handler(ctx, { file_path: revisedPath, anchor: a1c, position: 'after', text: 'New paragraph added by revision', instruction: 'add detail at end' });
const a2 = await anchorOf(revisedPath, 'Driven by enterprise renewals');
await tool('replace_text').handler(ctx, { file_path: revisedPath, anchor: a2, old_string: 'Driven by enterprise renewals', new_string: '', instruction: 'clear driver line' });
await tool('add_slide').handler(ctx, { file_path: revisedPath, instruction: 'append slide' });
await tool('edit_notes').handler(ctx, { file_path: revisedPath, slide: 2, text: 'Updated notes for comparison', instruction: 'notes change' });
// compare_decks is stateless — persist the session edits to disk first
await tool('save').handler(ctx, { file_path: revisedPath });

// Compare
const cmp = await tool('compare_decks').handler(ctx, {
  original_path: originalPath,
  revised_path: revisedPath,
  report_path: reportPath,
  annotate_copy_path: annotatedPath,
  allow_overwrite: true,
});
check('summary counts slides', cmp.summary.slidesA === 2 && cmp.summary.slidesB === 3, JSON.stringify(cmp.summary));
check('detects modified paragraph', cmp.changes.some((c) => c.kind === 'modified' && c.before.includes('12%') && c.after.includes('18%')));
check('detects added paragraph', cmp.changes.some((c) => c.kind === 'added' && c.after.includes('New paragraph added by revision')));
check('detects cleared paragraph', cmp.changes.some((c) => c.kind === 'modified' && c.before.includes('enterprise renewals') && c.after === ''));
check('detects slide added', cmp.slide_matches.some((m) => m.status === 'added'));
check('word diff present on modified', cmp.changes.find((c) => c.kind === 'modified')?.diff?.some((d) => d.op === 'ins' && d.text === '18%'));
check('detects notes change', cmp.changes.some((c) => c.isNotes && c.kind === 'modified' && c.after === 'Updated notes for comparison'));

// report file
const report = fs.readFileSync(reportPath, 'utf8');
check('report written with summary', report.includes('# Deck comparison') && report.includes('18%'));

// annotated copy: red runs on changed paragraphs
const zip = await JSZip.loadAsync(fs.readFileSync(annotatedPath));
const slide1 = await zip.files['ppt/slides/slide1.xml'].async('string');
check('annotated copy recolors changed paragraph', slide1.includes('18%') && slide1.includes('FF0000'));
const slide2 = await zip.files['ppt/slides/slide2.xml'].async('string');
check('annotated copy leaves untouched slide 2 shapes uncolored', !slide2.includes('FF0000'));

// identical decks
const same = await tool('compare_decks').handler(ctx, { original_path: originalPath, revised_path: originalPath });
check('identical decks report identical', same.summary.identical === true);

sessions.closeAll();
process.exit(checks.every(Boolean) ? 0 : 1);
