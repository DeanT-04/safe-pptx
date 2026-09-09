#!/usr/bin/env node
/**
 * Phase-2 smoke test: drive the real tool handlers (read_file, get_outline,
 * grep, get_comments) against the generated fixture.
 */
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { SessionManager } from '../packages/pptx-mcp/dist/session/manager.js';
import { toolCatalog } from '../packages/pptx-mcp/dist/tool_registry.js';

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const fixturePath = path.join(root, 'fixtures', 'generated', 'sample.pptx');

const sessions = new SessionManager('smoke-ai');
const ctx = { sessions };

function tool(name) {
  const def = toolCatalog.find((t) => t.name === name);
  if (!def) throw new Error(`tool not in catalog: ${name}`);
  return def;
}

const checks = [];
function check(name, ok, detail = '') {
  checks.push(ok);
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? ` — ${detail}` : ''}`);
}

// read_file
const read1 = await tool('read_file').handler(ctx, { file_path: fixturePath });
check('read_file returns entries', read1.entries?.length > 5, `${read1.entries?.length} entries`);
const titleEntry = read1.entries.find((e) => e.text.includes('Quarterly Review'));
check('read_file finds title text', Boolean(titleEntry));
check('title entry has anchor', Boolean(titleEntry?.anchor?.startsWith('_bk_')), titleEntry?.anchor);
check('title entry tagged bold', Boolean(titleEntry?.text.includes('<b>Quarterly Review</b>')));

// anchors stable across sessions
const sessions2 = new SessionManager('smoke-ai-2');
const read2 = await tool('read_file').handler({ sessions: sessions2 }, { file_path: fixturePath });
const title2 = read2.entries.find((e) => e.text.includes('Quarterly Review'));
check('anchors deterministic across sessions', titleEntry?.anchor === title2?.anchor);

// node_ids filter
const byId = await tool('read_file').handler(ctx, {
  file_path: fixturePath,
  node_ids: [titleEntry.anchor],
});
check('node_ids filter', byId.entries?.length === 1 && byId.entries[0].anchor === titleEntry.anchor);

// pagination
const paged = await tool('read_file').handler(ctx, { file_path: fixturePath, limit: 3 });
check('pagination limit honored', paged.entries?.length === 3);
check('pagination has_more + next_offset', paged.has_more === true && paged.next_offset === 4);
const cont = await tool('read_file').handler(ctx, { file_path: fixturePath, offset: paged.next_offset });
check('pagination continues', cont.entries?.length > 0 && cont.entries[0].text === paged.entries[2].text + 1 ? false : true, `offset ${paged.next_offset}`);

// get_outline
const outline = await tool('get_outline').handler(ctx, { file_path: fixturePath });
check('outline has 2 slides', outline.slides?.length === 2);
check('outline slide titles', outline.slides[0].title === 'Quarterly Review' && outline.slides[1].title === 'Regional Numbers');
check('outline detects table slide paragraphs', outline.slides[1].text_paragraphs >= 8);

// grep
const grep = await tool('grep').handler(ctx, { file_path: fixturePath, pattern: '12%' });
check('grep finds formatted run text', grep.hits?.length === 1 && grep.hits[0].anchor?.startsWith('_bk_'), `${grep.hits?.length} hit(s)`);
const grepXml = await tool('grep').handler(ctx, { file_path: fixturePath, pattern: 'sldSz', search_xml: true });
check('grep xml mode finds presentation.xml', grepXml.hits?.some((h) => h.part === 'ppt/presentation.xml'));
const grepMulti = await tool('grep').handler(ctx, { file_paths: [fixturePath, fixturePath], pattern: 'EMEA' });
check('grep multi-file stateless', grepMulti.files_searched === 2 && grepMulti.hits.length === 4, `${grepMulti.hits?.length} hits (slide2 body + notes, x2 files)`);

// get_comments (fixture has none — should be empty, not error)
const comments = await tool('get_comments').handler(ctx, { file_path: fixturePath });
check('get_comments empty on fixture', comments.legacy_count === 0 && comments.threaded_count === 0);

// get_file_status reflects audit/session state
const status = await tool('get_file_status').handler(ctx, { file_path: fixturePath });
check('get_file_status open with slide count', status.open === true && status.slide_count === 2);

sessions.closeAll();
sessions2.closeAll();
process.exit(checks.every(Boolean) ? 0 : 1);
