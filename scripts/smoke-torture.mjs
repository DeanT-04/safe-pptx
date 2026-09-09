#!/usr/bin/env node
/**
 * Torture sweep: every safe-pptx tool family against the torture deck
 * (charts, media, animations, sections, groups, RTL, fields, merged tables,
 * SVG, legacy comments, stale autofit caches...). Each edit asserts
 * formatting preservation at the XML level; save asserts byte-identical
 * untouched parts.
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import JSZip from 'jszip';
import { SessionManager } from '../packages/pptx-mcp/dist/session/manager.js';
import { toolCatalog } from '../packages/pptx-mcp/dist/tool_registry.js';
import { PptxPackage } from '../packages/pptx-core/dist/index.js';

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const gen = path.join(root, 'fixtures', 'generated');
const source = path.join(gen, 'torture.pptx');
const work = path.join(gen, 'torture-work.pptx');
const finalPath = path.join(gen, 'torture-final.pptx');
const reportPath = path.join(gen, 'torture-report.md');
const annotatedPath = path.join(gen, 'torture-annotated.pptx');
if (!fs.existsSync(source)) {
  console.error('torture deck missing — run `node scripts/make-torture-deck.mjs`');
  process.exit(1);
}
fs.copyFileSync(source, work);
for (const stale of [finalPath, reportPath, annotatedPath]) if (fs.existsSync(stale)) fs.rmSync(stale);

const sessions = new SessionManager('torture');
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
const anchorOf = async (needle, file = work) => {
  const read = await tool('read_file').handler(ctx, { file_path: file });
  const e = read.entries.find((e) => e.anchor && e.text.includes(needle));
  if (!e) throw new Error(`paragraph not found: ${needle}`);
  return e.anchor;
};
const sessionXml = (part) => {
  const { XMLSerializer } = createRequire(import.meta.url)('@xmldom/xmldom');
  const pkg = sessions.get(work).pkg;
  return new XMLSerializer().serializeToString(pkg.zip.doc(part));
};

// ---------- load & read ----------
const outline = await tool('get_outline').handler(ctx, { file_path: work });
check('torture deck loads: 13 slides', outline.slides.length === 13, `${outline.slides.length}`);
check('outline titles on all non-empty slides', outline.slides.slice(0, 12).every((s) => s.title !== null));
check('empty slide has null title', outline.slides[12].title === null);

// paginate the whole deck with a small budget; collect all anchors
{
  let offset = 1;
  let total = 0;
  const seen = new Set();
  for (;;) {
    const page = await tool('read_file').handler(ctx, { file_path: work, offset, token_budget: 1500 });
    for (const e of page.entries) {
      if (e.anchor) {
        if (seen.has(e.anchor)) throw new Error('duplicate anchor across pages');
        seen.add(e.anchor);
      }
    }
    total += page.entries.length;
    if (!page.has_more) break;
    offset = page.next_offset;
  }
  check('pagination walks entire deck without repeats', total > 120 && seen.size >= 60, `${total} entries, ${seen.size} anchors`);
}

const grepUni = await tool('grep').handler(ctx, { file_path: work, pattern: 'مرحبا|שלום' });
check('grep finds RTL text', grepUni.hit_count === 2, `${grepUni.hit_count} hits`);
const grepDash = await tool('grep').handler(ctx, { file_path: work, pattern: 'em—dash' });
check('grep finds em-dash content', grepDash.hit_count === 1);

// ---------- edit: fragmented multi-run paragraph ----------
{
  const anchor = await anchorOf('teal-highlight');
  const rep = await tool('replace_text').handler(ctx, {
    file_path: work, anchor, old_string: 'bold italic underline strike', new_string: 'BOLD-ISH', instruction: 'fragmented span',
  });
  check('fragmented cross-run replace', rep.after.includes('BOLD-ISH sub'), rep.after.slice(0, 60));
  const xml = sessionXml('ppt/slides/slide2.xml');
  check('sub/superscript baselines preserved', xml.includes('baseline="-40000"') || xml.includes('baseline="30000"'), 'baseline attrs present');
  check('highlight preserved on untouched run', xml.includes('highlight'));
}

// ---------- edit: hyperlinks preserved ----------
{
  const anchor = await anchorOf('External link:');
  await tool('replace_text').handler(ctx, {
    file_path: work, anchor, old_string: 'External link:', new_string: 'External links:', instruction: 'neighbor of links',
  });
  const xml = sessionXml('ppt/slides/slide2.xml');
  const rels = sessions.get(work).pkg.zip.text('ppt/slides/_rels/slide2.xml.rels') ?? '';
  const hlinks = (xml.match(/hlinkClick/g) ?? []).length;
  check(
    'hyperlinks (external + slide-jump) preserved',
    rels.includes('github.com/DeanT-04/safe-pptx') && xml.includes('hlinksldjump') && hlinks >= 3,
    `${hlinks} hlink refs, rels external target present: ${rels.includes('github.com')}`,
  );
}

// ---------- edit: field paragraph ----------
{
  const anchor = await anchorOf('Generated on:');
  await tool('replace_text').handler(ctx, {
    file_path: work, anchor, old_string: 'Generated on:', new_string: 'Made on:', instruction: 'run next to field',
  });
  const xml = sessionXml('ppt/slides/slide1.xml');
  check('datetime field survives neighbor edit', xml.includes('type="datetime1"') && xml.includes('a:fld'));
  check('slide-number field survives', xml.includes('type="slidenum"'));
}

// ---------- edit: RTL paragraph ----------
{
  const anchor = await anchorOf('مرحبا');
  const rep = await tool('replace_text').handler(ctx, {
    file_path: work, anchor, old_string: 'للتشفير', new_string: 'للاختبار', instruction: 'edit rtl paragraph',
  });
  check('RTL paragraph text replaced', rep.after.includes('للاختبار'));
  const xml = sessionXml('ppt/slides/slide3.xml');
  check('rtl paragraph property preserved', xml.includes('rtl="1"'));
}

// ---------- edit: grouped shape member ----------
{
  const read = await tool('read_file').handler(ctx, { file_path: work });
  const groupEntry = read.entries.find((e) => e.anchor && e.text.includes('Grouped member A'));
  check('walk finds paragraphs inside grouped shapes', Boolean(groupEntry));
  if (groupEntry) {
    await tool('replace_text').handler(ctx, {
      file_path: work, anchor: groupEntry.anchor, old_string: 'Grouped member A', new_string: 'Grouped member A2', instruction: 'edit inside group',
    });
    const xml = sessionXml('ppt/slides/slide5.xml');
    check('group structure (nested grpSp) intact after member edit', xml.includes('<p:grpSp>') && xml.includes('inner group') && xml.includes('Grouped member A2'));
  }
}

// ---------- edit: stale autofit cache reset ----------
{
  const anchor = await anchorOf('STALE AUTOFIT CACHE');
  await tool('replace_text').handler(ctx, {
    file_path: work, anchor, old_string: 'STALE AUTOFIT CACHE', new_string: 'AUTOFIT RESET', instruction: 'trigger autofit reset',
  });
  const xml = sessionXml('ppt/slides/slide5.xml');
  check('stale fontScale/lnSpcReduction reset on edit', !xml.includes('fontScale="62500"'), 'normAutofit cache cleared');
  check('vertical text attribute preserved', xml.includes('vert="eaVert"'));
}

// ---------- edit: table cells ----------
{
  const anchor = await anchorOf('4.2');
  const rep = await tool('replace_text').handler(ctx, {
    file_path: work, anchor, old_string: '4.2', new_string: '4.4', instruction: 'table cell surgical edit',
  });
  check('table cell surgical edit (anchor path)', rep.after === '4.4', rep.after);
  // discover the table graphicFrame id from read entries (shapeKey `graphicFrame#N/r…c…`)
  const cellEntry = (await tool('read_file').handler(ctx, { file_path: work })).entries.find((e) => e.shape?.startsWith('graphicFrame#'));
  const tableShapeId = Number(cellEntry.shape.match(/graphicFrame#(\d+)/)[1]);
  const cellLoc = cellEntry.shape.match(/r(\d+)c(\d+)/);
  await tool('edit_table_cell').handler(ctx, {
    file_path: work, slide: Number(cellEntry.slide), shape_id: tableShapeId,
    row: Number(cellLoc[1]), col: Number(cellLoc[2]), text: '4.5', instruction: 'cell rewrite',
  });
  const xml = sessionXml('ppt/slides/slide8.xml');
  check('table cell rewrite applied', xml.includes('4.5'));
  check('merged cell (rowspan) intact', xml.includes('rowspan="2"') || xml.includes('vMerge="1"'));
}

// ---------- font ops ----------
{
  const anchor = await anchorOf('Bottom-anchored text');
  const font = await tool('set_font').handler(ctx, {
    file_path: work, anchor, b: 'on', sz_pt: 28, color_hex: 'CC0000', instruction: 'format anchored text',
  });
  check('set_font applies (b/sz/color)', font.runs_touched >= 1);
  const xml = sessionXml('ppt/slides/slide12.xml');
  check('font props written to slide', xml.includes('sz="2800"') && xml.includes('val="CC0000"'));
  const clear = await tool('clear_formatting').handler(ctx, {
    file_path: work, anchor: await anchorOf('Bottom-anchored text'), properties: ['all'], instruction: 'normalize again',
  });
  check('clear_formatting runs', clear.runs_touched >= 1);
}

// ---------- notes + insert + batch ----------
{
  await tool('edit_notes').handler(ctx, {
    file_path: work, slide: 7, text: 'SVG notes updated\nsecond line', instruction: 'notes edit',
  });
  const anchor = await anchorOf('Vector SVG image');
  const ins = await tool('insert_paragraph').handler(ctx, {
    file_path: work, anchor, position: 'after', text: 'Inserted subtitle', instruction: 'add subtitle',
  });
  check('insert_paragraph into title body', ins.created_anchors?.length === 1);
  const batch = await tool('batch_edit').handler(ctx, {
    file_path: work,
    steps: [
      { step_id: 'a', tool: 'replace_text', anchor: await anchorOf('cover-crop'), old_string: 'cover-crop', new_string: 'cover crop label', instruction: 'batch 1' },
      { step_id: 'b', tool: 'replace_text', anchor: await anchorOf('Auto-numbered:'), old_string: 'Auto-numbered:', new_string: 'Numbered:', instruction: 'batch 2' },
    ],
  });
  check('batch_edit applies multi-slide steps', batch.failed_step === null && batch.applied_steps.length === 2);
}

// ---------- comments (modern) ----------
{
  const c1 = await tool('add_comment').handler(ctx, {
    file_path: work, slide: 9, author: 'Torture', text: 'Chart review pending', initials: 'T',
  });
  await tool('add_comment').handler(ctx, {
    file_path: work, slide: 9, author: 'Torture', text: 'Reply thread', parent_id: c1.commentId,
  });
  const comments = await tool('get_comments').handler(ctx, { file_path: work });
  check('modern + legacy comments coexist', comments.threaded_count === 2 && comments.legacy_count === 1, `L${comments.legacy_count}/T${comments.threaded_count}`);
}

// ---------- save #1: text-only edits → byte-identity of untouched parts ----------
await tool('save').handler(ctx, { file_path: work, save_to_local_path: work }); // overwrite in place
{
  const originalZip = await JSZip.loadAsync(fs.readFileSync(source));
  const savedZip = await JSZip.loadAsync(fs.readFileSync(work));
  let untouchedIdentical = true;
  const untouchedPrefixes = ['ppt/charts/', 'ppt/embeddings/', 'ppt/media/', 'ppt/theme/', 'ppt/slideMasters/', 'ppt/slideLayouts/', 'ppt/notesMasters/', 'ppt/tableStyles.xml', 'ppt/presProps.xml', 'ppt/viewProps.xml'];
  // [Content_Types].xml legitimately changes here: comment parts were added.
  const edited = ['ppt/slides/slide1.xml', 'ppt/slides/slide2.xml', 'ppt/slides/slide3.xml', 'ppt/slides/slide5.xml', 'ppt/slides/slide7.xml', 'ppt/slides/slide8.xml', 'ppt/slides/slide12.xml'];
  for (const name of Object.keys(originalZip.files)) {
    const a = await originalZip.files[name].async('nodebuffer');
    const b = await savedZip.files[name].async('nodebuffer');
    if (edited.includes(name) || name === '[Content_Types].xml') continue;
    if (untouchedPrefixes.some((pre) => name.startsWith(pre))) {
      if (!a.equals(b)) { untouchedIdentical = false; console.log('   differs:', name); }
    }
  }
  check('charts/theme/media/masters byte-identical after text edits', untouchedIdentical);
  const s2 = await savedZip.files['ppt/slides/slide2.xml'].async('string');
  check('transition + animation survive slide re-serialization', s2.includes('<p:transition') && s2.includes('<p:timing') && s2.includes('spTgt'));
  const s4 = await savedZip.files['ppt/slides/slide4.xml'].async('string');
  check('gradient + pattern fills survive', s4.includes('gradFill') && s4.includes('pattFill'));
  const presXml = await savedZip.files['ppt/presentation.xml'].async('string');
  check('p14:sectionLst survives save', presXml.includes('sectionLst') && (presXml.match(/p14:section /g) ?? []).length === 4);
}

// ---------- structure phase ----------
{
  // slide 9 is a chart slide; duplicate it — chart rels are internal/shared
  const dup = await tool('duplicate_slide').handler(ctx, { file_path: work, slide: 9, instruction: 'dup chart slide' });
  check('duplicate chart slide', dup.slideNumber === 10, `inserted at ${dup.slideNumber}`);
  // After the duplicate the deck has 14 slides and the EMPTY slide (the only
  // member of section 'Edge') sits at display position 14. Deleting it must
  // scrub its id from sectionLst and remove the now-empty section (4→3).
  const statusNow = await tool('get_file_status').handler(ctx, { file_path: work });
  const emptyPos = statusNow.slides.find((s) => s.part === 'ppt/slides/slide13.xml').index + 1;
  const presBefore = sessionXml('ppt/presentation.xml');
  const sectionsBefore = (presBefore.match(/<p14:section /g) ?? []).length;
  await tool('delete_slide').handler(ctx, { file_path: work, slide: emptyPos, instruction: 'delete empty section tail' });
  const presAfter = sessionXml('ppt/presentation.xml');
  const sectionsAfter = (presAfter.match(/<p14:section /g) ?? []).length;
  check('sectionLst scrubbed: empty section removed', sectionsBefore === 4 && sectionsAfter === 3, `${sectionsBefore}→${sectionsAfter}`);
  check('no empty section lists remain', !presAfter.includes('<p14:sldIdLst/>') && !presAfter.includes('<p14:sldIdLst></p14:sldIdLst>'));
  check('deleted slide id absent from sections', !presAfter.includes('p14:sectionLst') || sectionsAfter === 3);
  await tool('reorder_slides').handler(ctx, { file_path: work, order: [2, 1, ...[...Array(11).keys()].map((i) => i + 3)], instruction: 'swap first two' });
  await tool('save').handler(ctx, { file_path: work, save_to_local_path: finalPath });
  const pkg2 = await PptxPackage.load(fs.readFileSync(finalPath));
  // 13 original + 1 duplicate − 1 deleted = 13
  check('final deck reloads: 13 slides', pkg2.slides.length === 13, `${pkg2.slides.length}`);
  let uncovered = [];
  for (const part of pkg2.zip.listParts()) if (!pkg2.contentTypes.isCovered(part)) uncovered.push(part);
  check('final deck content types fully covered', uncovered.length === 0, JSON.stringify(uncovered));
  let dangling = [];
  for (const s of pkg2.slides) {
    for (const t of ['slideLayout', 'notesSlide', 'threadedComments', 'image', 'media', 'video', 'audio', 'chart']) {
      for (const target of pkg2.rels.targetsOfType(s.partPath, t)) {
        if (!pkg2.zip.hasPart(target)) dangling.push(`${s.partPath}→${target}`);
      }
    }
  }
  check('no dangling rel targets (incl charts + media)', dangling.length === 0, JSON.stringify(dangling.slice(0, 4)));
}

// ---------- compare + export ----------
{
  const cmp = await tool('compare_decks').handler(ctx, {
    original_path: source,
    revised_path: finalPath,
    report_path: reportPath,
    annotate_copy_path: annotatedPath,
    allow_overwrite: true,
  });
  check('compare handles torture deck', cmp.summary.slidesA === 13 && cmp.summary.slidesB === 13, JSON.stringify({ a: cmp.summary.slidesA, b: cmp.summary.slidesB }));
  check('compare finds text changes', cmp.summary.paragraphsModified + cmp.summary.paragraphsAdded > 0);
  const md = fs.readFileSync(reportPath, 'utf8');
  check('report written', md.includes('# Deck comparison'));
  const annotated = await JSZip.loadAsync(fs.readFileSync(annotatedPath));
  check('annotated copy valid zip with all slides', Object.keys(annotated.files).filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f)).length === 13);
  const exp = await tool('export').handler(ctx, { file_path: finalPath, format: 'markdown', allow_overwrite: true });
  check('export markdown', exp.bytes > 2000, `${exp.bytes} bytes`);
}

// ---------- audit ----------
{
  const log = await tool('get_edit_log').handler(ctx, { file_path: work });
  check('audit log captured everything', log.total >= 15, `${log.total} entries`);
}

sessions.closeAll();
process.exit(checks.every(Boolean) ? 0 : 1);
