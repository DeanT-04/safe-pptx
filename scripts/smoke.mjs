#!/usr/bin/env node
/**
 * Phase-1 smoke test: load the generated fixture through pptx-core and
 * verify the structural model (slide order, rels, content types).
 */
import path from 'node:path';
import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
import { PptxPackage } from '../packages/pptx-core/dist/index.js';

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const fixturePath = path.join(root, 'fixtures', 'generated', 'sample.pptx');

if (!fs.existsSync(fixturePath)) {
  console.error('fixture missing — run `npm run fixture` first');
  process.exit(1);
}

const pkg = await PptxPackage.load(fs.readFileSync(fixturePath));

const checks = [
  ['slide count', pkg.slides.length === 2],
  ['slide order by sldIdLst', pkg.slides.map((s) => s.sldId).join(',') === '256,257'],
  ['slide parts resolve', pkg.slides.every((s) => pkg.zip.hasPart(s.partPath))],
  ['layout resolves', pkg.layoutForSlide('ppt/slides/slide1.xml') === 'ppt/slideLayouts/slideLayout1.xml'],
  ['master resolves', pkg.masterForLayout('ppt/slideLayouts/slideLayout1.xml') === 'ppt/slideMasters/slideMaster1.xml'],
  ['content type of slide', pkg.contentTypes.contentTypeFor('ppt/slides/slide1.xml')?.includes('slide+xml')],
  ['theme found', pkg.partRefs.themeParts.includes('ppt/theme/theme1.xml')],
  ['dirty set starts empty', pkg.zip.dirtyParts().length === 0],
  ['untouched doc parse works', pkg.zip.doc('ppt/slides/slide1.xml').documentElement?.localName === 'sld'],
];

let failed = 0;
for (const [name, ok] of checks) {
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}`);
  if (!ok) failed++;
}
console.log(`\nparts: ${pkg.zip.listParts().length}, slides: ${pkg.slides.length}`);
process.exit(failed === 0 ? 0 : 1);
