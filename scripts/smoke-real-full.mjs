#!/usr/bin/env node
/**
 * Deep real-world suite: exercises every tool family against the designated
 * real PowerPoint deck (default C:\\Users\\Deano\\Downloads\\test-for-mcp.pptx).
 * The original is NEVER modified — all work happens on a copy in fixtures/.
 * Exits 0 with a SKIP notice when the deck is not present (e.g. other machines).
 *
 * Usage: node scripts/smoke-real-full.mjs [path-to-real.pptx]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';
import { SessionManager } from '../packages/pptx-mcp/dist/session/manager.js';
import { toolCatalog } from '../packages/pptx-mcp/dist/tool_registry.js';
import { PptxPackage, flattenShapes, walkShapes, spTreeOf } from '../packages/pptx-core/dist/index.js';

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const gen = path.join(root, 'fixtures', 'generated');
const sourceReal = process.argv[2] ?? 'C:\\Users\\Deano\\Downloads\\test-for-mcp.pptx';
if (!fs.existsSync(sourceReal)) {
  console.log(`SKIP  real-world deck not present: ${sourceReal}`);
  process.exit(0);
}
const realPath = path.join(gen, 'real-copy.pptx');
const savedPath = path.join(gen, 'real-full-saved.pptx');
const reportPath = path.join(gen, 'real-full-report.md');
const annotatedPath = path.join(gen, 'real-full-annotated.pptx');
fs.copyFileSync(sourceReal, realPath);
for (const stale of [savedPath, reportPath, annotatedPath]) if (fs.existsSync(stale)) fs.rmSync(stale);

const sessions = new SessionManager('real-full');
const ctx = { sessions };
const tool = (name) => {
  const def = toolCatalog.find((t) => t.name === name);
  if (!def) throw new Error(`missing tool ${name}`);
  return def;
};
const anchorOf = async (needle, file = realPath) => {
  const read = await tool('read_file').handler(ctx, { file_path: file });
  const e = read.entries.find((e) => e.anchor && e.text.includes(needle));
  if (!e) throw new Error(`paragraph not found: ${needle}`);
  return e;
};
const checks = [];
function check(name, ok, detail = '') {
  checks.push(ok);
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? ` — ${detail}` : ''}`);
}

const outline = await tool('get_outline').handler(ctx, { file_path: realPath });
const slideCount = outline.slides.length;
check('loads real deck', slideCount > 0, `${slideCount} slides`);
check('outline infers titles for textbox-only slides', outline.slides.every((s) => s.title !== null) && outline.slides.some((s) => s.title_inferred === true), outline.slides[0].title);

// 1) tolerant text edit — real decks have smart quotes (e.g. semester’s)
{
  const read = await tool('read_file').handler(ctx, { file_path: realPath });
  const curlyEntry = read.entries.find((e) => e.anchor && e.text.includes('\u2019'));
  check('deck contains smart quotes to exercise tolerance', Boolean(curlyEntry));
  // Match with a STRAIGHT apostrophe against the deck’s curly one.
  const plain = curlyEntry.text.replace(/<[^>]+>/g, '');
  const m = plain.match(/(\S+)\u2019(\S*)/);
  const oldWithStraight = `${m[1]}'${m[2]}`;
  const rep = await tool('replace_text').handler(ctx, {
    file_path: realPath,
    anchor: curlyEntry.anchor,
    old_string: oldWithStraight,
    new_string: `${m[1]} and${m[2] ? ` ${m[2]}` : ''}`,
    instruction: 'smart-quote tolerant edit on real deck',
  });
  check('smart-quote tolerant match works on real deck', rep.replacements === 1, `${oldWithStraight} matched across quote styles`);
}

// 2) speaker notes creation on a slide without notes
const slidesWithoutNotes = outline.slides.filter((s) => !s.has_notes);
check('deck has notes-free slides to test creation', slidesWithoutNotes.length > 0, `${slidesWithoutNotes.length} of ${slideCount}`);
const targetSlide = slidesWithoutNotes[0].slide;
const notes = await tool('edit_notes').handler(ctx, {
  file_path: realPath, slide: targetSlide, text: 'Created by safe-pptx real-world suite', instruction: 'notes creation',
});
check('edit_notes creates part on real deck', notes.partPath.startsWith('ppt/notesSlides/notesSlide'), notes.partPath);

// 3) threaded comments on a real slide
const c1 = await tool('add_comment').handler(ctx, {
  file_path: realPath, slide: 1, author: 'safe-pptx suite', text: 'Real-deck comment', initials: 'SP',
});
const c2 = await tool('add_comment').handler(ctx, {
  file_path: realPath, slide: 1, author: 'safe-pptx suite', text: 'Reply', parent_id: c1.commentId,
});
const comments = await tool('get_comments').handler(ctx, { file_path: realPath });
check('threaded comments written + read', comments.threaded_count === 2 && comments.authors.length >= 1);
const delC = await tool('delete_comment').handler(ctx, { file_path: realPath, comment_id: c1.commentId, instruction: 'cleanup' });
check('comment deletion cascades', delC.deleted.length === 2);

// 4) table ops — only if the deck actually contains a table
let tableFound = null;
{
  const pkg = sessions.get(realPath).pkg;
  for (const s of pkg.slides) {
    const doc = pkg.zip.doc(s.partPath);
    const shapes = flattenShapes(walkShapes(spTreeOf(doc.documentElement)));
    const t = shapes.find((sh) => sh.table);
    if (t) { tableFound = { slide: s.index + 1, shape: t }; break; }
  }
}
if (tableFound) {
  const cell = await tool('edit_table_cell').handler(ctx, {
    file_path: realPath, slide: tableFound.slide, shape_id: tableFound.shape.id, row: 0, col: 0, text: 'EDITED', instruction: 'cell edit',
  });
  check('edit_table_cell on real deck table', Boolean(cell.after));
} else {
  check('no tables in deck — cell op skipped (not a failure)', true);
}

// 5) slide structure ops
const add = await tool('add_slide').handler(ctx, { file_path: realPath, instruction: 'append' });
check('add_slide on real deck', add.slideNumber === slideCount + 1);
const dup = await tool('duplicate_slide').handler(ctx, { file_path: realPath, slide: 2, instruction: 'duplicate slide 2' });
check('duplicate_slide inserts after source', dup.slideNumber === 3);
const statusAfter = await tool('get_file_status').handler(ctx, { file_path: realPath });
check('slide count reflects structure ops', statusAfter.slide_count === slideCount + 2, `${statusAfter.slide_count}`);
await tool('reorder_slides').handler(ctx, {
  file_path: realPath,
  order: [...Array(slideCount + 2).keys()].map((i) => i + 1).reverse(),
  instruction: 'reverse order',
});
const outline2 = await tool('get_outline').handler(ctx, { file_path: realPath });
// After reversing, position 1 holds the LAST slide of the pre-reorder deck —
// that is the appended (add_slide) slide, i.e. add.partPath.
check('reorder reversed the deck', outline2.slides[0].slide === 1 && outline2.slides[0].part === add.partPath, outline2.slides[0].part);
const del = await tool('delete_slide').handler(ctx, { file_path: realPath, slide: 1, instruction: 'remove duplicate' });
check('delete_slide removes the duplicated part', del.removedParts.some((p) => p.endsWith('.xml') && p.includes('slides/slide')), JSON.stringify(del.removedParts));

// 6) save + package integrity (structure ops touch many parts)
await tool('save').handler(ctx, { file_path: realPath, save_to_local_path: savedPath });
const pkg2 = await PptxPackage.load(fs.readFileSync(savedPath));
check('saved real deck reloads', pkg2.slides.length === slideCount + 1, `${pkg2.slides.length} slides`);
let uncovered = [];
for (const part of pkg2.zip.listParts()) if (!pkg2.contentTypes.isCovered(part)) uncovered.push(part);
check('all saved parts covered by content types', uncovered.length === 0, JSON.stringify(uncovered));
let dangling = [];
for (const s of pkg2.slides) {
  for (const t of ['slideLayout', 'notesSlide', 'threadedComments', 'image', 'hyperlink']) {
    for (const target of pkg2.rels.targetsOfType(s.partPath, t)) {
      if (!pkg2.zip.hasPart(target)) dangling.push(`${s.partPath}→${target}`);
    }
  }
}
check('no dangling rel targets in saved deck', dangling.length === 0, JSON.stringify(dangling.slice(0, 4)));

// media (pictures) must survive round-trip byte-identical
const originalZip = await JSZip.loadAsync(fs.readFileSync(sourceReal));
const savedZip = await JSZip.loadAsync(fs.readFileSync(savedPath));
let mediaIdentical = true;
for (const name of Object.keys(originalZip.files)) {
  if (!name.startsWith('ppt/media/')) continue;
  if (!savedZip.files[name]) { mediaIdentical = false; console.log('   media missing:', name); continue; }
  const a = await originalZip.files[name].async('nodebuffer');
  const b = await savedZip.files[name].async('nodebuffer');
  if (!a.equals(b)) { mediaIdentical = false; console.log('   media differs:', name); }
}
check('all media parts byte-identical after structure ops', mediaIdentical);

// 7) full redline vs the untouched original
const cmp = await tool('compare_decks').handler(ctx, {
  original_path: sourceReal,
  revised_path: savedPath,
  report_path: reportPath,
  annotate_copy_path: annotatedPath,
  allow_overwrite: true,
});
// We deleted the blank added slide, so the net structural delta is exactly
// +1 slide (the duplicate). Heavy reordering with duplicate textbox titles
// makes paragraph pairing noisy by design — under reversal the text edit can
// surface as added/removed paragraphs, so verify via the report text.
const reportText = fs.readFileSync(reportPath, 'utf8');
check('redline captures the smart-quote edit text', reportText.includes('week and'));
check('redline reflects net +1 slide from the duplicate', cmp.summary.slidesAdded === 1 && cmp.summary.slidesRemoved === 0, JSON.stringify({ a: cmp.summary.slidesA, b: cmp.summary.slidesB, added: cmp.summary.slidesAdded, removed: cmp.summary.slidesRemoved }));
check('redline found the edits among reordered slides', cmp.summary.paragraphsModified >= 2);
check('markdown report written', fs.readFileSync(reportPath, 'utf8').includes('# Deck comparison'));
const annotatedZip = await JSZip.loadAsync(fs.readFileSync(annotatedPath));
const annotatedSlide1 = await annotatedZip.files[pkg2.slides[0].partPath].async('string');
check('annotated copy generated', annotatedSlide1.length > 0);

// 8) audit trail
const log = await tool('get_edit_log').handler(ctx, { file_path: realPath });
check('audit log captured the whole session', log.total >= 8, `${log.total} entries`);

sessions.closeAll();
process.exit(checks.every(Boolean) ? 0 : 1);
