import { test } from 'node:test';
import assert from 'node:assert/strict';
import { findMatchSpan } from '../edit/text_engine.js';
import { assignAnchors, normalizeForHash } from '../text/anchors.js';
import { buildParagraph, buildTaggedText, type ParagraphInfo } from '../text/model.js';
import { wordDiff } from '../compare/compare.js';
import { resolvePartPath, relsPathFor, nextFreeRelId } from '../package/rels.js';
import { UnsafeArchiveError, assertSafeArchive } from '../zip/archive_guard.js';
import { parseXml } from '../xml/parse.js';
import type { Element } from '../xml/parse.js';

test('findMatchSpan exact match', () => {
  assert.deepEqual(findMatchSpan('hello world', 'world'), { start: 6, end: 11 });
});

test('findMatchSpan occurrence selection', () => {
  const span = findMatchSpan('a b a b', 'a b', 2);
  assert.deepEqual(span, { start: 4, end: 7 });
});

test('findMatchSpan whitespace-tolerant', () => {
  assert.deepEqual(findMatchSpan('hello  world', 'hello world'), { start: 0, end: 12 });
});

test('findMatchSpan smart-quote tolerant', () => {
  assert.deepEqual(findMatchSpan('it\u2019s fine', "it's fine"), { start: 0, end: 9 });
});

test('findMatchSpan returns null when absent', () => {
  assert.equal(findMatchSpan('hello', 'xyz'), null);
});

test('anchors are deterministic and text-sensitive', () => {
  const mk = (text: string): ParagraphInfo => ({ el: null as never, index: 0, runs: [], breaks: 0, level: 0, text });
  const paras = [mk('alpha'), mk('beta'), mk('gamma')];
  const a1 = assignAnchors('ppt/slides/slide1.xml', [['sp#2']], paras);
  const a2 = assignAnchors('ppt/slides/slide1.xml', [['sp#2']], paras);
  assert.deepEqual(a1, a2);
  const a3 = assignAnchors('ppt/slides/slide1.xml', [['sp#2']], [mk('alpha'), mk('BETA'), mk('gamma')]);
  assert.notEqual(a1[1], a3[1]);
});

test('anchors salt duplicates', () => {
  const mk = (text: string): ParagraphInfo => ({ el: null as never, index: 0, runs: [], breaks: 0, level: 0, text });
  const paras = [mk('same'), mk('middle'), mk('same')];
  const ids = assignAnchors('ppt/slides/slide1.xml', [['sp#2']], paras);
  assert.notEqual(ids[0], ids[2]);
});

test('normalizeForHash collapses whitespace and invisibles', () => {
  assert.equal(normalizeForHash('a\u200B  b\n c'), 'a b c');
});

test('wordDiff basic insert/delete', () => {
  const diff = wordDiff('revenue grew 12 percent', 'revenue grew 18 percent');
  assert.deepEqual(
    diff.filter((d) => d.op !== 'same'),
    [
      { op: 'del', text: '12' },
      { op: 'ins', text: '18' },
    ],
  );
});

test('rels path helpers', () => {
  assert.equal(relsPathFor('ppt/slides/slide1.xml'), 'ppt/slides/_rels/slide1.xml.rels');
  assert.equal(relsPathFor('[Content_Types].xml'), '_rels/[Content_Types].xml.rels');
  assert.equal(resolvePartPath('ppt/slides', '../slideLayouts/slideLayout1.xml'), 'ppt/slideLayouts/slideLayout1.xml');
  assert.equal(resolvePartPath('ppt', 'slides/slide1.xml'), 'ppt/slides/slide1.xml');
  const doc = parseXml(
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="t" Target="a"/></Relationships>',
  );
  assert.equal(nextFreeRelId(doc), 'rId2');
});

test('archive guard rejects bombs and traversal', () => {
  assert.throws(() => assertSafeArchive([{ name: 'a', compressedSize: 100, uncompressedSize: 999999999 }]), UnsafeArchiveError);
  assert.throws(() => assertSafeArchive([{ name: '../evil', compressedSize: 1, uncompressedSize: 1 }]), UnsafeArchiveError);
  assert.doesNotThrow(() => assertSafeArchive([{ name: 'ok.xml', compressedSize: 100, uncompressedSize: 100 }]));
});

test('tagged text merges formatting runs', () => {
  const doc = parseXml(
    '<?xml version="1.0"?>' +
      '<p:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
      '<a:p><a:r><a:rPr b="1"/><a:t>bold </a:t></a:r><a:r><a:rPr b="1"/><a:t>text</a:t></a:r>' +
      '<a:r><a:t> plain</a:t></a:r></a:p></p:txBody>',
    'test',
  );
  const p = buildParagraph((doc.documentElement as Element).firstChild as never, 0);
  assert.equal(buildTaggedText(p), '<b>bold text</b> plain');
});
