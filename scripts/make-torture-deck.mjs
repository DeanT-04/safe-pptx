#!/usr/bin/env node
/**
 * Torture-deck generator: a .pptx containing EVERYTHING PowerPoint can show,
 * for stress-testing safe-pptx.
 *
 * Layer 1 (pptxgenjs, per the official pptx skill): rich text, bullets,
 *   hyperlinks (external + slide-jump), fields, shapes, tables with merges,
 *   native charts (each an own part + embedded workbook), speaker notes,
 *   slide numbers, sections-ready structure.
 * Layer 2 (raw OOXML post-process, jszip + @xmldom/xmldom): transitions,
 *   animations (p:timing), p14:sectionLst, gradient/pattern fills, grouped
 *   shapes (nested grpSp), vertical text, stale normAutofit cache, SVG blip,
 *   legacy comments, audio + video media parts (ffmpeg-generated), datetime
 *   field injection.
 *
 * Output: fixtures/generated/torture.pptx (gitignored; regenerate anytime).
 * Run from the repo root: node scripts/make-torture-deck.mjs
 */
import { createRequire } from 'node:module';
import { execSync } from 'node:child_process';
import path from 'node:path';
import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const gen = path.join(root, 'fixtures', 'generated');
const assets = path.join(gen, 'assets');
fs.mkdirSync(assets, { recursive: true });
const outPath = path.join(gen, 'torture.pptx');

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
const sharp = requireGlobal('sharp');

// ---------------------------------------------------------------- assets --
const logoPng = path.join(assets, 'logo.png');
const photoJpg = path.join(assets, 'photo.jpg');
const photoPng = path.join(assets, 'photo.png');
const iconSvgPath = path.join(assets, 'icon.svg');

const svgLogo = `<svg xmlns="http://www.w3.org/2000/svg" width="400" height="400">
  <defs><radialGradient id="g" cx="35%" cy="35%"><stop offset="0%" stop-color="#5EEAD4"/><stop offset="100%" stop-color="#0F766E"/></radialGradient></defs>
  <circle cx="200" cy="200" r="180" fill="url(#g)"/>
  <circle cx="200" cy="200" r="110" fill="#0F172A" fill-opacity="0.85"/>
  <text x="200" y="235" font-family="Arial" font-size="120" font-weight="bold" fill="#F8FAFC" text-anchor="middle">S</text>
</svg>`;
const svgPhoto = `<svg xmlns="http://www.w3.org/2000/svg" width="960" height="540">
  <defs><linearGradient id="sky" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#0C4A6E"/><stop offset="60%" stop-color="#E0F2FE"/><stop offset="100%" stop-color="#FEF9C3"/></linearGradient></defs>
  <rect width="960" height="540" fill="url(#sky)"/>
  <circle cx="760" cy="140" r="70" fill="#FDE68A"/>
  <polygon points="0,540 260,220 480,540" fill="#166534"/>
  <polygon points="300,540 620,160 960,540" fill="#14532D"/>
  <rect y="500" width="960" height="40" fill="#065F46"/>
</svg>`;
await sharp(Buffer.from(svgLogo)).png().toFile(logoPng);
await sharp(Buffer.from(svgPhoto)).jpeg({ quality: 82 }).toFile(photoJpg);
await sharp(Buffer.from(svgPhoto)).png().toFile(photoPng);
const svgIcon = `<svg xmlns="http://www.w3.org/2000/svg" width="240" height="240"><path d="M40 200 L120 40 L200 200 Z" fill="#E8590C" stroke="#7C2D12" stroke-width="6"/><circle cx="120" cy="160" r="24" fill="#FCD34D"/></svg>`;
fs.writeFileSync(iconSvgPath, svgIcon);

// media via ffmpeg
const clipMp4 = path.join(assets, 'clip.mp4');
const toneMp3 = path.join(assets, 'tone.mp3');
for (const f of [clipMp4, toneMp3]) if (fs.existsSync(f)) fs.rmSync(f);
execSync(`ffmpeg -y -loglevel error -f lavfi -i testsrc=duration=1:size=320x180:rate=12 -pix_fmt yuv420p "${clipMp4}"`);
execSync(`ffmpeg -y -loglevel error -f lavfi -i sine=frequency=440:duration=1 -q:a 9 "${toneMp3}"`);

// ------------------------------------------------------------- layer 1 ----
const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';
pres.author = 'safe-pptx torture suite';
pres.title = 'Torture deck';
const W = 13.33;

const notes = (slide, text) => slide.addNotes(text);
const addSlideNum = (slide) => { slide.slideNumber = { x: 12.6, y: 7.05, fontFace: 'Arial', fontSize: 10, color: '64748B' }; };

// Slide 1 — cover
{
  const s = pres.addSlide();
  s.background = { color: '0F172A' };
  s.addText('Torture deck', { x: 0.7, y: 2.3, w: 11.9, h: 1.3, fontSize: 60, bold: true, color: 'F8FAFC', fontFace: 'Arial' });
  s.addText([
    { text: 'Every PowerPoint feature in one file — ', options: { color: '94A3B8' } },
    { text: 'built to break editors', options: { color: '5EEAD4', italic: true } },
  ], { x: 0.7, y: 3.6, w: 11.9, h: 0.7, fontSize: 20, fontFace: 'Arial' });
  s.addText('visit docs', { x: 0.7, y: 6.4, w: 2, h: 0.4, fontSize: 12, color: '94A3B8', hyperlink: { url: 'https://github.com/DeanT-04/safe-pptx' }, fontFace: 'Arial' });
  addSlideNum(s);
  notes(s, 'Cover notes — first line\nSecond line with detail\n(third line)');
}

// Slide 2 — rich text torture (also gets transition + entrance animation)
{
  const s = pres.addSlide();
  s.addText('Rich text torture', { x: 0.5, y: 0.35, w: 12, h: 0.7, fontSize: 32, bold: true, color: '0F172A', fontFace: 'Arial' });
  s.addText([
    { text: 'bold ', options: { bold: true } },
    { text: 'italic ', options: { italic: true } },
    { text: 'underline ', options: { underline: true } },
    { text: 'strike ', options: { strike: true } },
    { text: 'sub X', options: { subscript: true } },
    { text: ' ', options: {} },
    { text: 'sup 2', options: { superscript: true } },
    { text: ' spaced', options: { charSpacing: 6 } },
    { text: ' courier', options: { fontFace: 'Courier New' } },
    { text: ' red', options: { color: 'CC0000' } },
    { text: ' teal-highlight', options: { highlight: '5EEAD4' } },
    { text: ' “smart quotes” — em—dash … ellipsis', options: {} },
  ], { x: 0.5, y: 1.2, w: 12.3, h: 1.6, fontSize: 18, fontFace: 'Arial', color: '1E293B', margin: 0 });
  s.addText([
    { text: 'Auto-numbered:', options: { breakLine: true, bold: true } },
    { text: 'first', options: { bullet: { type: 'number' }, breakLine: true } },
    { text: 'second', options: { bullet: { type: 'number' }, breakLine: true } },
    { text: 'char bullets:', options: { breakLine: true, bold: true } },
    { text: 'one', options: { bullet: { code: '25B8', indent: 12 }, breakLine: true } },
    { text: 'two', options: { bullet: { code: '25B8', indent: 12 } } },
  ], { x: 0.5, y: 3.0, w: 5.8, h: 3.4, fontSize: 16, fontFace: 'Arial', color: '334155', paraSpaceAfter: 6, margin: 0 });
  s.addText([
    { text: 'External link: ', options: {} },
    { text: 'safe-pptx repo', options: { hyperlink: { url: 'https://github.com/DeanT-04/safe-pptx' }, color: '0D9488', underline: true } },
    { text: '   |   internal jump: ', options: {} },
    { text: 'go to slide 9', options: { hyperlink: { slide: 9 }, color: '7C3AED', underline: true } },
  ], { x: 6.6, y: 3.4, w: 6.2, h: 1.2, fontSize: 14, fontFace: 'Arial', color: '1E293B', margin: 0 });
  addSlideNum(s);
  notes(s, 'Slide 2 notes: mixed-run edits must preserve every rPr.');
}

// Slide 3 — international & RTL (also gets the legacy comment)
{
  const s = pres.addSlide();
  s.addText('International & RTL', { x: 0.5, y: 0.35, w: 12, h: 0.7, fontSize: 32, bold: true, fontFace: 'Arial' });
  s.addText([
    { text: 'Arabic (RTL): مرحبا بالعالم — هذا نص عربي للتشفير', options: { rtlMode: true, breakLine: true } },
    { text: 'Hebrew (RTL): שלום עולם בדיקה', options: { rtlMode: true, breakLine: true } },
    { text: 'CJK: 中文测试 — 标点、顿号、省略号……全角', options: { breakLine: true } },
    { text: 'Emoji: 🎉 🚀 ✅ and combining: é ñ ü', options: { breakLine: true } },
    { text: 'Tabs:\tA\tB\tC and trailing spaces   ', options: { breakLine: true } },
    { text: 'Long unbroken: AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA', options: {} },
  ], { x: 0.5, y: 1.3, w: 12.3, h: 4.6, fontSize: 16, fontFace: 'Arial', color: '1E293B', paraSpaceAfter: 10, margin: 0 });
  addSlideNum(s);
  notes(s, 'Slide 3 notes — international');
}

// Slide 4 — shapes gallery
{
  const s = pres.addSlide();
  s.addText('Shapes gallery', { x: 0.5, y: 0.35, w: 12, h: 0.7, fontSize: 32, bold: true, fontFace: 'Arial' });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.3, w: 3, h: 1.6, rectRadius: 0.12, fill: { color: '0D9488' }, shadow: { type: 'outer', color: '000000', blur: 8, offset: 3, angle: 45, opacity: 0.3 } });
  s.addText('rounded + shadow', { x: 0.5, y: 1.75, w: 3, h: 0.6, align: 'center', fontSize: 12, color: 'F0FDFA', fontFace: 'Arial' });
  s.addShape(pres.shapes.OVAL, { x: 4, y: 1.3, w: 1.8, h: 1.8, fill: { color: '7C3AED' }, line: { color: '4C1D95', width: 2 } });
  s.addShape(pres.shapes.CHEVRON, { x: 6.2, y: 1.3, w: 2.2, h: 1.4, fill: { color: 'E8590C' } });
  s.addShape(pres.shapes.CUBE, { x: 8.8, y: 1.3, w: 1.7, h: 1.7, fill: { color: '334155' } });
  s.addShape(pres.shapes.LINE, { x: 0.8, y: 3.6, w: 4, h: 0, line: { color: '0F172A', width: 2.5, endArrowType: 'triangle' } });
  s.addShape(pres.shapes.LINE, { x: 0.8, y: 4.2, w: 4, h: 1.2, line: { color: 'DC2626', width: 1.5, dashType: 'dash', beginArrowType: 'oval', endArrowType: 'stealth' } });
  s.addText('GRADIENT TARGET', { x: 5.4, y: 3.5, w: 3.2, h: 1.2, shape: pres.shapes.RECTANGLE, fill: { color: '0D9488' }, align: 'center', fontSize: 12, color: 'FFFFFF', fontFace: 'Arial' });
  s.addText('PATTERN TARGET', { x: 9, y: 3.5, w: 3.2, h: 1.2, shape: pres.shapes.RECTANGLE, fill: { color: '94A3B8' }, align: 'center', fontSize: 12, color: '0F172A', fontFace: 'Arial' });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 5.6, w: 12.3, h: 1.3, fill: { color: 'E2E8F0' } });
  s.addText('footer band with content: connectors above, presets here', { x: 0.7, y: 6.0, w: 11.9, h: 0.5, fontSize: 12, color: '475569', fontFace: 'Arial' });
  addSlideNum(s);
  notes(s, 'Slide 4 notes — shapes');
}

// Slide 5 — group & text direction (post-process wraps shapes in nested grpSp)
{
  const s = pres.addSlide();
  s.addText('Grouped shapes & text direction', { x: 0.5, y: 0.35, w: 12, h: 0.7, fontSize: 32, bold: true, fontFace: 'Arial' });
  s.addText('Grouped member A', { shape: pres.shapes.RECTANGLE, x: 1, y: 1.5, w: 2.4, h: 1.2, fill: { color: '0D9488' }, fontSize: 12, color: 'FFFFFF', align: 'center', fontFace: 'Arial' });
  s.addText('Grouped member B', { shape: pres.shapes.OVAL, x: 1.6, y: 2.0, w: 2.4, h: 1.2, fill: { color: 'F59E0B' }, fontSize: 12, color: '3B2A00', align: 'center', fontFace: 'Arial' });
  s.addText('Grouped square', { shape: pres.shapes.RECTANGLE, x: 5, y: 1.8, w: 2, h: 2, fill: { color: '7C3AED' }, fontSize: 12, color: 'FFFFFF', align: 'center', fontFace: 'Arial' });
  s.addText('vertical text 一二三四五六七八九十', { x: 8.2, y: 1.4, w: 1.4, h: 4.5, fontSize: 18, fontFace: 'Arial', color: '0F172A', vert: 'eaVert' });
  s.addText('STALE AUTOFIT CACHE — editing this text must reset fontScale', { x: 10, y: 1.4, w: 2.8, h: 4.4, fontSize: 10, fontFace: 'Arial', color: '1E293B' });
  addSlideNum(s);
  notes(s, 'Slide 5 notes — groups');
}

// Slide 6 — raster images
{
  const s = pres.addSlide();
  s.addText('Raster images', { x: 0.5, y: 0.35, w: 12, h: 0.7, fontSize: 32, bold: true, fontFace: 'Arial' });
  s.addImage({ path: logoPng, x: 0.7, y: 1.4, w: 2.4, h: 2.4 });
  s.addImage({ path: photoJpg, x: 3.6, y: 1.4, w: 4.4, h: 2.475 });
  s.addImage({ path: photoPng, x: 8.4, y: 1.4, w: 2.2, h: 1.2375, rotate: 12 });
  s.addImage({ path: logoPng, x: 10.9, y: 1.4, w: 1.6, h: 1.6, rounding: true });
  s.addImage({ path: photoJpg, x: 0.7, y: 4.3, w: 5.5, h: 2.4, sizing: { type: 'cover', w: 5.5, h: 2.4 } });
  s.addImage({ path: logoPng, x: 6.6, y: 4.5, w: 1.2, h: 1.2, transparency: 40 });
  s.addText('cover-crop', { x: 0.9, y: 6.1, w: 2, h: 0.4, fontSize: 11, color: 'FFFFFF', fontFace: 'Arial' });
  addSlideNum(s);
  notes(s, 'Slide 6 notes — images');
}

// Slide 7 — SVG (post-process converts the PNG host to blip+svgBlip)
{
  const s = pres.addSlide();
  s.addText('Vector SVG image', { x: 0.5, y: 0.35, w: 12, h: 0.7, fontSize: 32, bold: true, fontFace: 'Arial' });
  s.addImage({ path: photoPng, x: 4.4, y: 1.6, w: 4.5, h: 2.53 }); // SVGHOST — swapped to svg blip in post-process
  s.addText('The triangle above becomes a real a:svgBlip with PNG fallback after post-processing.', { x: 3.4, y: 4.4, w: 6.5, h: 0.8, align: 'center', fontSize: 13, color: '475569', fontFace: 'Arial' });
  addSlideNum(s);
  notes(s, 'Slide 7 notes — svg');
}

// Slide 8 — tables with merges
{
  const s = pres.addSlide();
  s.addText('Tables with merged cells', { x: 0.5, y: 0.35, w: 12, h: 0.7, fontSize: 32, bold: true, fontFace: 'Arial' });
  const rows = [
    [
      { text: 'Region', options: { bold: true, color: 'FFFFFF', fill: { color: '0F766E' } } },
      { text: 'Q1', options: { bold: true, color: 'FFFFFF', fill: { color: '0F766E' } } },
      { text: 'Q2', options: { bold: true, color: 'FFFFFF', fill: { color: '0F766E' } } },
      { text: 'Total', options: { bold: true, color: 'FFFFFF', fill: { color: '134E4A' } } },
    ],
    [
      { text: 'EMEA', options: { rowspan: 2, fill: { color: 'CCFBF1' } } },
      '4.2', '4.8', '9.0',
    ],
    ['4.1', '4.9', '9.0'],
    [{ text: 'APAC', options: { fill: { color: 'FEF3C7' } } }, '2.9', '3.3', '6.2'],
  ];
  s.addTable(rows, { x: 0.5, y: 1.3, w: 12.3, colW: [3.3, 3, 3, 3], border: { pt: 1, color: '94A3B8' }, fontFace: 'Arial', fontSize: 13, align: 'center', valign: 'middle' });
  addSlideNum(s);
  notes(s, 'Slide 8 notes — merges + spans');
}

// Slides 9-10 — charts
{
  const s = pres.addSlide();
  s.addText('Charts I — bar & line', { x: 0.5, y: 0.35, w: 12, h: 0.7, fontSize: 32, bold: true, fontFace: 'Arial' });
  s.addChart(pres.charts.BAR, [{ name: 'Bookings', labels: ['EMEA', 'AMER', 'APAC'], values: [42, 51, 29] }], {
    x: 0.5, y: 1.2, w: 6, h: 5.2, barDir: 'col', chartColors: ['0D9488'], showValue: true, dataLabelPosition: 'outEnd', showLegend: false,
  });
  s.addChart(pres.charts.BAR, [
    { name: 'New', labels: ['Q1', 'Q2', 'Q3', 'Q4'], values: [10, 22, 18, 30] },
    { name: 'Renewal', labels: ['Q1', 'Q2', 'Q3', 'Q4'], values: [14, 12, 20, 16] },
  ], {
    x: 6.8, y: 1.2, w: 6, h: 2.5, barDir: 'bar', barGrouping: 'stacked', chartColors: ['0D9488', 'F59E0B'], showValue: true, dataLabelPosition: 'ctr', showLegend: true, legendPos: 'b',
  });
  s.addChart(pres.charts.LINE, [
    { name: 'ARR', labels: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'], values: [100, 108, 115, 124, 130, 141] },
    { name: 'Plan', labels: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'], values: [100, 105, 112, 118, 126, 134] },
  ], {
    x: 6.8, y: 3.9, w: 6, h: 2.6, lineSize: 2.5, lineSmooth: true, chartColors: ['0D9488', '94A3B8'], showLegend: true, legendPos: 'b',
  });
  addSlideNum(s);
  notes(s, 'Slide 9 notes — charts have embedded workbooks');
}
{
  const s = pres.addSlide();
  s.addText('Charts II — pie, doughnut, area, radar', { x: 0.5, y: 0.35, w: 12, h: 0.7, fontSize: 32, bold: true, fontFace: 'Arial' });
  s.addChart(pres.charts.PIE, [{ name: 'Share', labels: ['Direct', 'Partner', 'Other'], values: [55, 30, 15] }], {
    x: 0.4, y: 1.2, w: 3.2, h: 2.9, showPercent: true, chartColors: ['0D9488', 'F59E0B', '94A3B8'],
  });
  s.addChart(pres.charts.DOUGHNUT, [{ name: 'Mix', labels: ['Enterprise', 'SMB'], values: [70, 30] }], {
    x: 3.7, y: 1.2, w: 3.2, h: 2.9, holeSize: 55, chartColors: ['7C3AED', 'C4B5FD'],
  });
  s.addChart(pres.charts.AREA, [{ name: 'Users (k)', labels: ['W1', 'W2', 'W3', 'W4'], values: [12, 15, 14, 19] }], {
    x: 0.4, y: 4.2, w: 6.5, h: 2.7, chartColors: ['0EA5E9'],
  });
  s.addChart(pres.charts.RADAR, [
    { name: 'Us', labels: ['Speed', 'Cost', 'Quality', 'Reach'], values: [8, 6, 9, 7] },
    { name: 'Peer', labels: ['Speed', 'Cost', 'Quality', 'Reach'], values: [7, 8, 6, 6] },
  ], {
    x: 7.1, y: 1.2, w: 5.8, h: 5.7, radarStyle: 'marker', chartColors: ['0F766E', 'CBD5E1'],
  });
  addSlideNum(s);
  notes(s, 'Slide 10 notes — four more chart parts');
}

// Slide 11 — media (video + audio injected in post-process)
{
  const s = pres.addSlide();
  s.addText('Media — embedded video & audio', { x: 0.5, y: 0.35, w: 12, h: 0.7, fontSize: 32, bold: true, fontFace: 'Arial' });
  s.addImage({ path: photoPng, x: 0.7, y: 1.4, w: 5.6, h: 3.15 }); // VIDEOPOSTER
  s.addImage({ path: logoPng, x: 7.2, y: 1.6, w: 1.8, h: 1.8 }); // AUDIOPOSTER
  s.addText('video (mp4, ffmpeg testsrc) + audio (mp3 tone)', { x: 7.2, y: 3.7, w: 5.4, h: 0.6, fontSize: 13, color: '475569', fontFace: 'Arial' });
  addSlideNum(s);
  notes(s, 'Slide 11 notes — media');
}

// Slide 12 — autofit & overflow
{
  const s = pres.addSlide();
  s.addText('Autofit & overflow', { x: 0.5, y: 0.35, w: 12, h: 0.7, fontSize: 32, bold: true, fontFace: 'Arial' });
  s.addText(
    'Huge autofit body. '.repeat(40),
    { x: 0.5, y: 1.2, w: 6, h: 5.6, fontSize: 20, fontFace: 'Arial', autoFit: true, color: '1E293B' },
  );
  s.addText('Bottom-anchored text', { x: 7, y: 4.5, w: 5.6, h: 2.2, fontSize: 16, valign: 'bottom', color: '0F766E', fontFace: 'Arial' });
  addSlideNum(s);
  notes(s, 'Slide 12 notes — autofit caches');
}

// Slide 13 — empty slide (no shapes at all beyond the spTree)
{
  const s = pres.addSlide();
  s.background = { color: 'FFFFFF' };
  notes(s, 'Slide 13 notes — the slide itself is empty');
}

// ------------------------------------------------------------- layer 2 ----
await pres.writeFile({ fileName: outPath });
const raw = await JSZip.loadAsync(fs.readFileSync(outPath));

const readText = async (p) => raw.file(p).async('string');
const writeText = (p, s) => raw.file(p, s);
const XMLNS = {
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  p14: 'http://schemas.microsoft.com/office/powerpoint/2010/main',
};

function slidePaths() {
  return Object.keys(raw.files)
    .filter((p) => /^ppt\/slides\/slide\d+\.xml$/.test(p))
    .sort((a, b) => Number(a.match(/(\d+)/)[1]) - Number(b.match(/(\d+)/)[1]));
}

for (const sp of slidePaths()) {
  let xml = await readText(sp);
  // transition on every slide (after clrMapOvr, before timing if any)
  if (!xml.includes('<p:transition')) {
    xml = xml.replace(/<\/p:clrMapOvr>/, '</p:clrMapOvr><p:transition spd="med"><p:fade/></p:transition>');
  }
  writeText(sp, xml);
}

// Entrance animation on slide 2 (fade-in of the first text shape on click).
// Built programmatically so every element nests/closes correctly.
{
  const p = 'ppt/slides/slide2.xml';
  let xml = await readText(p);
  const spid = xml.match(/<p:cNvPr id="(\d+)"/)[1];
  const el = (name, attrs, ...children) =>
    `<${name}${Object.entries(attrs ?? {}).map(([k, v]) => ` ${k}="${v}"`).join('')}>${children.join('')}</${name}>`;
  const ctn = (id, extra, ...children) => el('p:cTn', { id, ...extra }, ...children);
  const stCond = (delay) => el('p:stCondLst', null, el('p:cond', { delay }));
  const tgt = (spid2) => el('p:tgtEl', null, el('p:spTgt', { spid: spid2 }));
  const setVis = el('p:set', null,
    el('p:cBhvr', null,
      ctn(6, { dur: '1', fill: 'hold' }, stCond('0')),
      tgt(spid),
      el('p:attrNameLst', null, el('p:attrName', null, 'style.visibility')),
    ),
    el('p:to', null, el('p:strVal', { val: 'visible' })),
  );
  const fade = el('p:animEffect', { transition: 'in', filter: 'fade' },
    el('p:cBhvr', null, ctn(7, { dur: '500' }), tgt(spid)),
  );
  const timing = el('p:timing', null,
    el('p:tnLst', null,
      el('p:par', null,
        ctn(1, { dur: 'indefinite', restart: 'never', nodeType: 'tmRoot' },
          el('p:childTnLst', null,
            el('p:seq', { concurrent: '1', nextAc: 'seek' },
              ctn(2, { dur: 'indefinite', nodeType: 'mainSeq' },
                el('p:childTnLst', null,
                  el('p:par', null,
                    ctn(3, { fill: 'hold' },
                      stCond('indefinite'),
                      el('p:childTnLst', null,
                        el('p:par', null,
                          ctn(4, { fill: 'hold' },
                            stCond('0'),
                            el('p:childTnLst', null,
                              el('p:par', null,
                                ctn(5, { presetID: '10', presetClass: 'entr', presetSubtype: '0', fill: 'hold', grpId: '0', nodeType: 'clickEffect' },
                                  stCond('0'),
                                  el('p:childTnLst', null, setVis, fade),
                                ),
                              ),
                            ),
                          ),
                        ),
                      ),
                    ),
                  ),
                ),
              ),
              el('p:prevCondLst', null, el('p:cond', { evt: 'onPrev', delay: '0' }, el('p:tgtEl', null, el('p:sldTgt')))),
              el('p:nextCondLst', null, el('p:cond', { evt: 'onNext', delay: '0' }, el('p:tgtEl', null, el('p:sldTgt')))),
            ),
          ),
        ),
      ),
    ),
  );
  xml = xml.replace('</p:sld>', `${timing}</p:sld>`);
  writeText(p, xml);
}

// Gradient + pattern fills on the named shapes (slide 4)
{
  const p = 'ppt/slides/slide4.xml';
  let xml = await readText(p);
  const gradShape = xml.match(/<p:sp>(?:(?!<\/p:sp>).)*?GRADIENT TARGET.*?<\/p:sp>/s)[0];
  const gradFilled = gradShape.replace(
    /<a:solidFill><a:srgbClr val="0D9488"\/><\/a:solidFill>/,
    '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="0D9488"/></a:gs><a:gs pos="100000"><a:srgbClr val="0F172A"/></a:gs></a:gsLst><a:lin ang="2700000" scaled="1"/></a:gradFill>',
  );
  xml = xml.replace(gradShape, gradFilled);
  const pattShape = xml.match(/<p:sp>(?:(?!<\/p:sp>).)*?PATTERN TARGET.*?<\/p:sp>/s)[0];
  const pattFilled = pattShape.replace(
    /<a:solidFill><a:srgbClr val="94A3B8"\/><\/a:solidFill>/,
    '<a:pattFill prst="pct20"><a:fgClr><a:srgbClr val="475569"/></a:fgClr><a:bgClr><a:srgbClr val="E2E8F0"/></a:bgClr></a:pattFill>',
  );
  xml = xml.replace(pattShape, pattFilled);
  writeText(p, xml);
}

// Nested group on slide 5: wrap rect+oval in an inner grpSp, then wrap that
// plus the third shape in an outer grpSp.
{
  const { DOMParser, XMLSerializer } = requireGlobal('@xmldom/xmldom');
  const p = 'ppt/slides/slide5.xml';
  let xml = await readText(p);
  const doc = new DOMParser().parseFromString(xml, 'text/xml');
  const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
  const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const tree = doc.getElementsByTagNameNS(P, 'spTree')[0];
  const els = Array.from(tree.getElementsByTagNameNS(P, 'sp'));
  const byName = (n) => els.find((el) => el.getElementsByTagNameNS(P, 'cNvPr')[0]?.getAttribute('name') === n);
  // pptxgenjs names shapes Text/Autoshape N; find by geometry instead: the two small left shapes + the square
  const sps = Array.from(tree.childNodes).filter((n) => n.nodeType === 1 && n.localName === 'sp');
  const rect = sps.find((el) => el.getElementsByTagNameNS(A, 'solidFill')[0]?.innerHTML?.includes('0D9488') || new XMLSerializer().serializeToString(el).includes('val="0D9488"'));
  const oval = sps.find((el) => new XMLSerializer().serializeToString(el).includes('val="F59E0B"'));
  const square = sps.find((el) => new XMLSerializer().serializeToString(el).includes('val="7C3AED"'));
  if (!rect || !oval || !square) throw new Error('group-slide shapes not found');
  let nextId = 100;
  const mkGroup = (children, name) => {
    const g = doc.createElementNS(P, 'p:grpSp');
    const nv = doc.createElementNS(P, 'p:nvGrpSpPr');
    const cnv = doc.createElementNS(P, 'p:cNvPr');
    cnv.setAttribute('id', String(nextId++));
    cnv.setAttribute('name', name);
    nv.appendChild(cnv);
    const cnvGrp = doc.createElementNS(P, 'p:cNvGrpSpPr');
    nv.appendChild(cnvGrp);
    g.appendChild(nv);
    const grpPr = doc.createElementNS(P, 'p:grpSpPr');
    const xfrm = doc.createElementNS(A, 'a:xfrm');
    const off = doc.createElementNS(A, 'a:off'); off.setAttribute('x', '0'); off.setAttribute('y', '0');
    const ext = doc.createElementNS(A, 'a:ext'); ext.setAttribute('cx', '0'); ext.setAttribute('cy', '0');
    const chOff = doc.createElementNS(A, 'a:chOff'); chOff.setAttribute('x', '0'); chOff.setAttribute('y', '0');
    const chExt = doc.createElementNS(A, 'a:chExt'); chExt.setAttribute('cx', '0'); chExt.setAttribute('cy', '0');
    for (const el of [off, ext, chOff, chExt]) xfrm.appendChild(el);
    grpPr.appendChild(xfrm);
    g.appendChild(grpPr);
    for (const c of children) g.appendChild(c);
    return g;
  };
  const inner = mkGroup([rect, oval], 'inner group');
  const outer = mkGroup([inner, square], 'outer group');
  // mkGroup already detached rect/oval/square from the tree (appendChild moves
  // nodes) — attach the outer group at the end of the shape tree.
  tree.appendChild(outer);
  xml = new XMLSerializer().serializeToString(doc);
  if (xml.startsWith('<?xml')) xml = xml.slice(xml.indexOf('?>') + 2);
  writeText(p, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${xml}`);
}

// Stale autofit cache on slide 5 (vert="eaVert" is already emitted by
// pptxgenjs natively for the vertical text box — nothing to inject there).
{
  const p = 'ppt/slides/slide5.xml';
  let xml = await readText(p);
  const staleShape = xml.match(/<p:sp>(?:(?!<\/p:sp>).)*?STALE AUTOFIT CACHE.*?<\/p:sp>/s)[0];
  if (staleShape.includes('normAutofit')) throw new Error('stale autofit already present');
  xml = xml.replace(staleShape, staleShape.replace(/<a:bodyPr([^>]*)\/>/, '<a:bodyPr$1><a:normAutofit fontScale="62500" lnSpcReduction="1000"/></a:bodyPr>'));
  writeText(p, xml);
}

// Datetime field on slide 1 (injected after the first lstStyle in the deck)
{
  const p = 'ppt/slides/slide1.xml';
  let xml = await readText(p);
  const fld = '<a:fld id="{2E7D1AAB-5A0B-4F4F-9C55-1D3E1F4B7C10}" type="datetime1"><a:rPr lang="en-US"/><a:t>Wednesday, September 9, 2026</a:t></a:fld>';
  const before = xml.length;
  xml = xml.replace('<a:lstStyle/>', `<a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="1200"/><a:t>Generated on: </a:t></a:r>${fld}</a:p>`);
  if (xml.length === before) throw new Error('datetime field injection failed');
  writeText(p, xml);
}

// SVG blip on slide 7 (swap the SVGHOST png blip to png fallback + svgBlip ext)
{
  const p = 'ppt/slides/slide7.xml';
  let xml = await readText(p);
  const relsPath = 'ppt/slides/_rels/slide7.xml.rels';
  let rels = await readText(relsPath);
  // find the rId of the photoPng image (the only image on this slide)
  const m = xml.match(/<a:blip r:embed="(rId\d+)"\s*\/?>/);
  if (!m) throw new Error('SVGHOST blip not found');
  const pngRid = m[1];
  void pngRid;
  const svgRid = 'rIdSvg1';
  // add svg part + rel + content types
  raw.file('ppt/media/icon1.svg', fs.readFileSync(iconSvgPath));
  rels = rels.replace('</Relationships>', `<Relationship Id="${svgRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/icon1.svg"/></Relationships>`);
  writeText(relsPath, rels);
  xml = xml.replace(
    m[0],
    `${m[0]}<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><a:svgBlip r:embed="${svgRid}"/></a:ext></a:extLst>`,
  );
  writeText(p, xml);
}

// Media on slide 11: video (poster + videoFile + p14:media) and audio
{
  const p = 'ppt/slides/slide11.xml';
  const relsPath = 'ppt/slides/_rels/slide11.xml.rels';
  let xml = await readText(p);
  let rels = await readText(relsPath);
  raw.file('ppt/media/clip1.mp4', fs.readFileSync(clipMp4));
  raw.file('ppt/media/tone1.mp3', fs.readFileSync(toneMp3));
  const rel = (id, type, target) => `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`;
  rels = rels.replace('</Relationships>', [
    rel('rIdVid1', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/video', '../media/clip1.mp4'),
    rel('rIdVid2', 'http://schemas.microsoft.com/office/2007/relationships/media', '../media/clip1.mp4'),
    rel('rIdAud1', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio', '../media/tone1.mp3'),
    rel('rIdAud2', 'http://schemas.microsoft.com/office/2007/relationships/media', '../media/tone1.mp3'),
  ].join('') + '</Relationships>');
  writeText(relsPath, rels);
  const blips = [...xml.matchAll(/<a:blip r:embed="(rId\d+)"\s*\/?>/g)].map((m2) => m2[1]);
  const posterRid = blips[0];
  const audioPosterRid = blips[1] ?? blips[0];
  const audioRid = xml.match(/<a:blip r:embed="(rId\d+)"[^/]*\/>/); // second blip (audio poster) — find via lastIndexOf below
  void audioRid;
  const videoPic = `<p:pic><p:nvPicPr><p:cNvPr id="90" name="Torture video"><a:hlinkClick r:id="" action="ppaction://media"/></p:cNvPr><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr><a:videoFile r:link="rIdVid1"/><p:extLst><p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"><p14:media xmlns:p14="${XMLNS.p14}" r:embed="rIdVid2"/></p:ext></p:extLst></p:nvPr></p:nvPicPr><p:blipFill><a:blip r:embed="${posterRid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="640080" y="1828800"/><a:ext cx="5107920" cy="2873375"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
  const audioPic = `<p:pic><p:nvPicPr><p:cNvPr id="91" name="Torture audio"><a:hlinkClick r:id="" action="ppaction://media"/></p:cNvPr><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr><a:audioFile r:link="rIdAud1"/><p:extLst><p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"><p14:media xmlns:p14="${XMLNS.p14}" r:embed="rIdAud2"/></p:ext></p:extLst></p:nvPr></p:nvPicPr><p:blipFill><a:blip r:embed="${audioPosterRid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="6583680" y="1828800"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
  xml = xml.replace('</p:spTree>', `${videoPic}${audioPic}</p:spTree>`);
  writeText(p, xml);
}

// Legacy comments on slide 3 (authors + comment part + rel + content types)
{
  const p = 'ppt/slides/slide3.xml';
  const relsPath = 'ppt/slides/_rels/slide3.xml.rels';
  let rels = await readText(relsPath);
  raw.file('ppt/commentAuthors.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:cmAuthorLst xmlns:a="${XMLNS.a}" xmlns:r="${XMLNS.r}" xmlns:p="${XMLNS.p}"><p:cmAuthor id="0" name="Torture Suite" initials="TS" lastIdx="1" clrIdx="0"/></p:cmAuthorLst>`);
  raw.file('ppt/comments/comment1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:cmLst xmlns:a="${XMLNS.a}" xmlns:r="${XMLNS.r}" xmlns:p="${XMLNS.p}"><p:cm authorId="0" dt="2026-09-09T12:00:00.000" idx="1"><p:pos x="1000" y="1000"/><p:text>Legacy comment — safe-pptx must read this.</p:text></p:cm></p:cmLst>`);
  rels = rels.replace('</Relationships>', '<Relationship Id="rIdCmt1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments/comment1.xml"/></Relationships>');
  writeText(relsPath, rels);
}

// Sections in presentation.xml (p14:sectionLst) — Intro / Visuals / Data / Edge
{
  const p = 'ppt/presentation.xml';
  let xml = await readText(p);
  const ids = [...xml.matchAll(/<p:sldId id="(\d+)"/g)].map((m) => m[1]);
  const sec = (name, list) => `<p14:section name="${name}" id="{A1B2${list.length}C3-D4E5-4F60-9A7B-00000000000${list.length}}"><p14:sldIdLst>${list.map((id) => `<p14:sldId id="${id}"/>`).join('')}</p14:sldIdLst></p14:section>`;
  const sectionLst = `<p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><p14:sectionLst xmlns:p14="${XMLNS.p14}">${sec('Intro', ids.slice(0, 3))}${sec('Visuals', ids.slice(3, 8))}${sec('Data', ids.slice(8, 12))}${sec('Edge', ids.slice(12))}</p14:sectionLst></p:ext></p:extLst>`;
  if (xml.includes('<p:extLst>')) throw new Error('presentation already has extLst — merge logic needed');
  xml = xml.replace('</p:presentation>', `${sectionLst}</p:presentation>`);
  writeText(p, xml);
}

// Content types for the added parts
{
  const ctPath = '[Content_Types].xml';
  let ct = await readText(ctPath);
  const add = (s) => {
    if (!ct.includes(s)) ct = ct.replace('</Types>', `${s}</Types>`);
  };
  add('<Default Extension="mp4" ContentType="video/mp4"/>');
  add('<Default Extension="mp3" ContentType="audio/mpeg"/>');
  add('<Default Extension="svg" ContentType="image/svg+xml"/>');
  add('<Override PartName="/ppt/commentAuthors.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml"/>');
  add('<Override PartName="/ppt/comments/comment1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.comments+xml"/>');
  writeText(ctPath, ct);
}

// Rezip preserving entry order
fs.rmSync(outPath);
const out = new JSZip();
for (const name of Object.keys(raw.files)) {
  if (raw.files[name].dir) continue;
  out.file(name, await raw.files[name].async('nodebuffer'));
}
fs.mkdirSync(path.dirname(outPath), { recursive: true });
const buffer = await out.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 6 } });
fs.writeFileSync(outPath, buffer);

const stats = {
  bytes: buffer.length,
  slides: slidePaths().length,
  charts: Object.keys(raw.files).filter((f) => /^ppt\/charts\/chart\d+\.xml$/.test(f)).length,
  media: Object.keys(raw.files).filter((f) => f.startsWith('ppt/media/')).length,
  notes: Object.keys(raw.files).filter((f) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(f)).length,
};
console.log(`torture deck: ${outPath}`);
console.log(JSON.stringify(stats));
