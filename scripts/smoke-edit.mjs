#!/usr/bin/env node
/**
 * Phase-3 smoke test: the edit core. Drives replace_text (cross-run, tolerant),
 * insert_paragraph, set_font, clear_formatting, batch_edit, get_edit_log and
 * save; then reloads the saved file and verifies byte-identical untouched parts.
 *
 * Anchors are content-derived: an edit changes the anchor, so tests re-read
 * anchors after each edit (the tools also return new_anchor).
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
const workPath = path.join(root, 'fixtures', 'generated', 'work-copy.pptx');
const savedPath = path.join(root, 'fixtures', 'generated', 'saved.pptx');
fs.copyFileSync(fixturePath, workPath);
for (const stale of [savedPath]) {
  if (fs.existsSync(stale)) fs.rmSync(stale);
}

const sessions = new SessionManager('smoke-ai');
const ctx = { sessions };
const tool = (name) => {
  const def = toolCatalog.find((t) => t.name === name);
  if (!def) throw new Error(`missing tool ${name}`);
  return def;
};
const anchorOf = async (needle) => {
  const read = await tool('read_file').handler(ctx, { file_path: workPath });
  const entry = read.entries.find((e) => e.anchor && e.text.includes(needle));
  if (!entry) throw new Error(`paragraph not found: ${needle}`);
  return entry.anchor;
};

const checks = [];
function check(name, ok, detail = '') {
  checks.push(ok);
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? ` — ${detail}` : ''}`);
}

// 1) replace a run fully inside the green bold run → replacement keeps that run's formatting
const a1 = await anchorOf('12%');
const rep1 = await tool('replace_text').handler(ctx, {
  file_path: workPath, anchor: a1, old_string: '12%', new_string: '15%', instruction: 'update figure',
});
check('replace_text inside run', rep1.after === 'Revenue grew 15% quarter over quarter', rep1.after);
check('new_anchor returned', typeof rep1.new_anchor === 'string' && rep1.new_anchor.startsWith('_bk_'), rep1.new_anchor);
{
  const session = sessions.get(workPath);
  const { XMLSerializer } = await import('@xmldom/xmldom');
  const xml = new XMLSerializer().serializeToString(session.pkg.zip.doc('ppt/slides/slide1.xml'));
  check('replacement kept green+bold rPr', xml.includes('<a:t>15%</a:t>') && xml.includes('1F7A33'));
}

// 2) whitespace-tolerant match (pattern has double space; text has single)
const a2 = rep1.new_anchor;
const rep2 = await tool('replace_text').handler(ctx, {
  file_path: workPath, anchor: a2, old_string: 'Revenue grew  15%', new_string: 'Revenue surged 15%', instruction: 'tolerant match',
});
check('whitespace-tolerant match', rep2.after.startsWith('Revenue surged 15%'), rep2.after);

// 3) cross-run replace (spans the plain run tail and the following run)
const a3 = rep2.new_anchor;
const rep3 = await tool('replace_text').handler(ctx, {
  file_path: workPath, anchor: a3, old_string: 'surged 15% quarter', new_string: 'grew 16% quarter', instruction: 'cross-run edit',
});
check('cross-run replace', rep3.after === 'Revenue grew 16% quarter over quarter', rep3.after);

// 4) deletion via empty new_string
const a4 = rep3.new_anchor;
const rep4 = await tool('replace_text').handler(ctx, {
  file_path: workPath, anchor: a4, old_string: ' over', new_string: '', instruction: 'delete phrase',
});
check('deletion via empty replacement', rep4.after === 'Revenue grew 16% quarter over quarter'.replace(' over', ''), rep4.after);

// 5) insert_paragraph before title (style cloned from title paragraph)
const titleAnchor = await anchorOf('Quarterly Review');
const ins = await tool('insert_paragraph').handler(ctx, {
  file_path: workPath, anchor: titleAnchor, position: 'before', text: 'CONFIDENTIAL', instruction: 'add label',
});
check('insert_paragraph creates anchor', ins.created_anchors?.length === 1 && ins.created_anchors[0].startsWith('_bk_'), String(ins.created_anchors?.[0]));

// 6) set_font on a span (split boundary runs)
const a6 = rep4.new_anchor;
const font = await tool('set_font').handler(ctx, {
  file_path: workPath, anchor: a6, u: 'on', sz_pt: 24, span_text: 'Revenue grew', instruction: 'underline lead',
});
check('set_font touched runs', font.runs_touched >= 1, `runs=${font.runs_touched}`);
{
  const session = sessions.get(workPath);
  const { XMLSerializer } = await import('@xmldom/xmldom');
  const xml = new XMLSerializer().serializeToString(session.pkg.zip.doc('ppt/slides/slide1.xml'));
  check('set_font wrote u+sz only on span', xml.includes('u="sng"') && xml.includes('sz="2400"'));
}

// 7) clear_formatting on whole paragraph
const clear = await tool('clear_formatting').handler(ctx, {
  file_path: workPath, anchor: rep4.new_anchor, properties: ['all'], instruction: 'normalize formatting',
});
check('clear_formatting ran', clear.runs_touched >= 1);

// 8) batch_edit — validation failure (unknown anchor) applies nothing
// (title anchor is re-fetched: the insert above changed its prev-neighbor, so the
//  pre-insert anchor is legitimately stale — same contract as safe-docx)
const titleAnchor2 = await anchorOf('Quarterly Review');
const batch = await tool('batch_edit').handler(ctx, {
  file_path: workPath,
  steps: [
    { step_id: 'a', tool: 'replace_text', anchor: titleAnchor2, old_string: 'Quarterly', new_string: 'Monthly', instruction: 'x' },
    { step_id: 'b', tool: 'replace_text', anchor: '_bk_doesnotexist', old_string: 'x', new_string: 'y', instruction: 'fail' },
  ],
});
check('batch_edit fails validation upfront', batch.error?.includes('not found'), String(batch.error).slice(0, 60));
{
  const read = await tool('read_file').handler(ctx, { file_path: workPath });
  check('batch_edit applied nothing on failed validation', !read.entries.some((e) => e.text.includes('Monthly')));
}

// 9) batch_edit — mid-application failure reports exactly what applied
const a9 = await anchorOf('enterprise renewals');     // p2: "Driven by enterprise renewals"
const batch2 = await tool('batch_edit').handler(ctx, {
  file_path: workPath,
  steps: [
    { step_id: 'reword', tool: 'replace_text', anchor: a9, old_string: 'enterprise renewals', new_string: 'enterprise expansions', instruction: 'reword' },
    { step_id: 'wrongpara', tool: 'replace_text', anchor: a9, old_string: 'Churn held steady', new_string: 'Churn stayed flat', instruction: 'targets wrong paragraph' },
  ],
});
check('batch_edit mid-fail reports step', batch2.failed_step === 'wrongpara' && batch2.applied_steps.length === 1, JSON.stringify({ applied: batch2.applied_steps, failed: batch2.failed_step }));
check('batch_edit applied the valid step first', batch2.results?.[0]?.after?.includes('enterprise expansions'));

// the churn step must target its own paragraph
const a9b = await anchorOf('Churn held steady');
const batch3 = await tool('batch_edit').handler(ctx, {
  file_path: workPath,
  steps: [
    { step_id: 'churn', tool: 'replace_text', anchor: a9b, old_string: 'Churn held steady', new_string: 'Churn stayed flat', instruction: 'reword churn' },
  ],
});
check('batch_edit applies to correct paragraph', batch3.failed_step === null && batch3.applied_steps.length === 1);

// 10) audit log
const log = await tool('get_edit_log').handler(ctx, { file_path: workPath });
check('audit log recorded edits', log.total >= 8, `${log.total} entries`);

// 11) save + byte-identical untouched parts
const save = await tool('save').handler(ctx, { file_path: workPath, save_to_local_path: savedPath });
check('save touched only slide1', save.edited_parts.length === 1 && save.edited_parts[0] === 'ppt/slides/slide1.xml', JSON.stringify(save.edited_parts));

const originalZip = await JSZip.loadAsync(fs.readFileSync(fixturePath));
const savedZip = await JSZip.loadAsync(fs.readFileSync(savedPath));
let untouchedIdentical = true;
let touchedDiffer = false;
for (const name of Object.keys(originalZip.files)) {
  const a = await originalZip.files[name].async('nodebuffer');
  const b = await savedZip.files[name].async('nodebuffer');
  if (name === 'ppt/slides/slide1.xml') {
    if (!a.equals(b)) touchedDiffer = true;
  } else if (!a.equals(b)) {
    untouchedIdentical = false;
    console.log('   differs:', name);
  }
}
check('untouched parts byte-identical', untouchedIdentical);
check('edited part actually changed', touchedDiffer);

// 12) reload saved deck
const readSaved = await tool('read_file').handler({ sessions: new SessionManager('r2') }, { file_path: savedPath });
check('saved deck contains final text', Boolean(readSaved.entries.find((e) => e.text.includes('Churn stayed flat'))));
const pkg2 = await PptxPackage.load(fs.readFileSync(savedPath));
check('saved deck reloads with 2 slides', pkg2.slides.length === 2);

sessions.closeAll();
process.exit(checks.every(Boolean) ? 0 : 1);
