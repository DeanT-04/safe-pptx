#!/usr/bin/env node
/**
 * Phase-4 smoke test: structure tools — edit_notes (existing + creation path),
 * edit_table_cell, add/delete threaded comments, add/duplicate/reorder/delete
 * slides — each verified for package integrity after save.
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';
import { SessionManager } from '../packages/pptx-mcp/dist/session/manager.js';
import { toolCatalog } from '../packages/pptx-mcp/dist/tool_registry.js';
import { PptxPackage } from '../packages/pptx-core/dist/index.js';

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const fixturePath = path.join(root, 'fixtures', 'generated', 'sample.pptx');
const workPath = path.join(root, 'fixtures', 'generated', 'struct-work.pptx');
const savedPath = path.join(root, 'fixtures', 'generated', 'struct-saved.pptx');
fs.copyFileSync(fixturePath, workPath);
for (const stale of [savedPath]) if (fs.existsSync(stale)) fs.rmSync(stale);

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

// 1) edit_notes on slide 2 (notes part exists)
const notes = await tool('edit_notes').handler(ctx, {
  file_path: workPath, slide: 2, text: 'First note line\nSecond note line', instruction: 'update speaker notes',
});
check('edit_notes rewrites existing notes', notes.before.includes('Emphasize the EMEA beat') && notes.after.includes('Second note line'));

// 2) edit_notes on slide 1 (creation path: notesSlide2 + rels + override)
const notes2 = await tool('edit_notes').handler(ctx, {
  file_path: workPath, slide: 1, text: 'Created notes', instruction: 'create notes for slide 1',
});
check('edit_notes creates notes part', notes2.partPath.includes('notesSlide2.xml'), notes2.partPath);
{
  const pkg = sessions.get(workPath).pkg;
  const created = pkg.slideByIndex(0);
  check('created notes linked from slide 1', pkg.notesSlideForSlide(created.partPath)?.includes('notesSlide2.xml'));
  const ct = pkg.contentTypes.contentTypeFor(notes2.partPath);
  check('created notes has content-type override', Boolean(ct?.includes('notesSlide')));
}

// 3) edit_table_cell
const cell = await tool('edit_table_cell').handler(ctx, {
  file_path: workPath, slide: 2, shape_id: 3, row: 1, col: 1, text: '4.4M', instruction: 'update EMEA bookings',
});
check('edit_table_cell updates cell', cell.before === '4.2M' && cell.after === '4.4M', `${cell.before}→${cell.after}`);
let threw = false;
try {
  await tool('edit_table_cell').handler(ctx, { file_path: workPath, slide: 2, shape_id: 3, row: 9, col: 9, text: 'x', instruction: 'out of range' });
} catch { threw = true; }
check('edit_table_cell rejects out-of-range', threw);

// 4) threaded comments: add + reply + read + delete
const c1 = await tool('add_comment').handler(ctx, {
  file_path: workPath, slide: 1, author: 'Reviewer One', text: 'Check this figure', initials: 'R1', instruction: 'review note',
});
check('add_comment creates threaded comment', c1.commentId.length > 20 && c1.partPath.includes('threadedComment1.xml'), c1.partPath);
const c2 = await tool('add_comment').handler(ctx, {
  file_path: workPath, slide: 1, author: 'Reviewer Two', text: 'Will do', parent_id: c1.commentId, instruction: 'reply',
});
const read = await tool('get_comments').handler(ctx, { file_path: workPath });
check('get_comments sees threaded comments', read.threaded_count === 2 && read.authors.length === 2);
check('reply linked via parentId', read.threaded.some((c) => c.parentId === c1.commentId));

// 5) add_slide + duplicate_slide
const add = await tool('add_slide').handler(ctx, { file_path: workPath, instruction: 'append slide' });
check('add_slide appends slide 3', add.slideNumber === 3 && add.partPath === 'ppt/slides/slide3.xml', add.partPath);
const dup = await tool('duplicate_slide').handler(ctx, { file_path: workPath, slide: 2, instruction: 'duplicate regional slide' });
check('duplicate_slide inserts after source', dup.slideNumber === 3, `new slide at position ${dup.slideNumber}`);

const status = await tool('get_file_status').handler(ctx, { file_path: workPath });
check('slide list reflects 4 slides', status.slide_count === 4, `count=${status.slide_count}`);

// 6) reorder_slides — slides are [s1, s2, s4(dup), s3(new)]; [4,1,2,3] puts s3 first
await tool('reorder_slides').handler(ctx, { file_path: workPath, order: [4, 1, 2, 3], instruction: 'move added slide to front' });
const outline = await tool('get_outline').handler(ctx, { file_path: workPath });
check('reorder moves slide', outline.slides[0].part === 'ppt/slides/slide3.xml', outline.slides[0].part);

let reorderThrew = false;
try { await tool('reorder_slides').handler(ctx, { file_path: workPath, order: [1, 2, 3], instruction: 'bad permutation' }); } catch { reorderThrew = true; }
check('reorder rejects bad permutation', reorderThrew);

// 7) delete_slide (delete the added slide, now position 1)
const del = await tool('delete_slide').handler(ctx, { file_path: workPath, slide: 1, instruction: 'remove added slide' });
check('delete_slide removes parts', del.removedParts.includes('ppt/slides/slide3.xml'), JSON.stringify(del.removedParts));
const status2 = await tool('get_file_status').handler(ctx, { file_path: workPath });
check('back to 3 slides', status2.slide_count === 3);

// 8) save + full package integrity on reload
await tool('save').handler(ctx, { file_path: workPath, save_to_local_path: savedPath });
const pkg2 = await PptxPackage.load(fs.readFileSync(savedPath));
check('saved deck reloads: 3 slides', pkg2.slides.length === 3);
check('saved deck slide order intact', pkg2.slides.map((s) => s.partPath).join(',') === 'ppt/slides/slide1.xml,ppt/slides/slide2.xml,ppt/slides/slide4.xml', pkg2.slides.map((s) => s.partPath).join(','));

const savedZip = await JSZip.loadAsync(fs.readFileSync(savedPath));
check('deleted slide part absent', !savedZip.files['ppt/slides/slide3.xml']);
check('duplicated slide part present', Boolean(savedZip.files['ppt/slides/slide4.xml']));
check('notesSlide2 present after save', Boolean(savedZip.files['ppt/notesSlides/notesSlide2.xml']));
check('threadedComments part present', Object.keys(savedZip.files).some((f) => f.includes('threadedComments/threadedComment1.xml')));
// every part covered by content types
const savedPkg = pkg2;
let uncovered = [];
for (const part of savedPkg.zip.listParts()) {
  if (!savedPkg.contentTypes.isCovered(part)) uncovered.push(part);
}
check('all parts covered by content types', uncovered.length === 0, JSON.stringify(uncovered));
// every slide rel target exists
let dangling = [];
for (const s of savedPkg.slides) {
  for (const relType of ['slideLayout', 'notesSlide', 'threadedComments']) {
    for (const target of savedPkg.rels.targetsOfType(s.partPath, relType)) {
      if (!savedPkg.zip.hasPart(target)) dangling.push(`${s.partPath} → ${target}`);
    }
  }
}
check('no dangling rel targets from slides', dangling.length === 0, JSON.stringify(dangling));

// 9) comment deletion
const del2 = await tool('delete_comment').handler(ctx, { file_path: workPath, comment_id: c1.commentId, instruction: 'cleanup' });
check('delete_comment cascades replies', del2.deleted.length === 2, JSON.stringify(del2.deleted));

sessions.closeAll();
process.exit(checks.every(Boolean) ? 0 : 1);
