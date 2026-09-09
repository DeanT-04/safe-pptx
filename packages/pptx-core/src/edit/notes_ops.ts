import { NS } from '../xml/namespaces.js';
import { childrenByName, firstChildByName, firstDescendantByName, makeElement } from '../xml/parse.js';
import type { Document, Element } from '../xml/parse.js';
import { CT } from '../package/content_types.js';
import type { PptxPackage } from '../package/presentation.js';
import { walkShapes, spTreeOf, flattenShapes } from '../package/shapes.js';
import { buildTextBody } from '../text/model.js';
import { addRelToDoc, relsPathFor } from '../package/rels.js';
import { EditError } from './text_engine.js';

/**
 * Notes and table-cell editing. Both rebuild a text body's paragraphs while
 * cloning the existing paragraph properties and run formatting as templates,
 * so new content inherits the deck's styling.
 */

const NOTES_XMLNS = `xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"`;

function maxPartNumber(pkg: PptxPackage, dirPrefix: string, base: string): number {
  let max = 0;
  for (const path of pkg.zip.listParts()) {
    const m = path.match(new RegExp(`^${dirPrefix}/${base}(\\d+)\\.xml$`));
    if (m) max = Math.max(max, Number(m[1]));
  }
  return max;
}

/** Create the notes slide part for a slide (requires an existing notes master) and link it up. */
export function ensureNotesPart(pkg: PptxPackage, slidePartPath: string): string {
  const existing = pkg.notesSlideForSlide(slidePartPath);
  if (existing) return existing;
  const notesMaster = pkg.partRefs.notesMasters[0];
  if (!notesMaster) {
    throw new EditError(
      'deck has no notes master — cannot create speaker notes (open the deck once in PowerPoint and add any note first)',
    );
  }
  const n = maxPartNumber(pkg, 'ppt/notesSlides', 'notesSlide') + 1;
  const partPath = `ppt/notesSlides/notesSlide${n}.xml`;

  const notesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes ${NOTES_XMLNS}>
<p:cSld><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>
<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>
</p:spTree></p:cSld>
<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:notes>`;

  const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="${slidePartPath.replace('ppt/', '../')}"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="${notesMaster.replace('ppt/', '../')}"/>
</Relationships>`;

  pkg.zip.addPart(partPath, Buffer.from(notesXml, 'utf8'));
  pkg.zip.addPart(relsPathFor(partPath), Buffer.from(relsXml, 'utf8'));
  pkg.contentTypes.setOverride(partPath, CT.notesSlide);
  pkg.zip.markDirty('[Content_Types].xml');

  const slideRelsPath = relsPathFor(slidePartPath);
  if (!pkg.zip.hasPart(slideRelsPath)) {
    pkg.zip.addPart(slideRelsPath, Buffer.from(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
      'utf8',
    ));
  }
  const slideRelsDoc = pkg.zip.doc(slideRelsPath);
  addRelToDoc(slideRelsDoc, {
    type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide',
    target: `../notesSlides/notesSlide${n}.xml`,
  });
  pkg.zip.markDirty(slideRelsPath);
  pkg.refreshStructure();
  return partPath;
}

/** Replace the full text of a slide's speaker notes (creating the notes part when needed). */
export function setNotesText(pkg: PptxPackage, slideNumber: number, text: string): { partPath: string; before: string; after: string } {
  const slide = pkg.slideByIndex(slideNumber - 1);
  if (!slide) throw new EditError(`slide ${slideNumber} does not exist`);
  const partPath = ensureNotesPart(pkg, slide.partPath);
  const doc = pkg.zip.doc(partPath);
  const spTree = spTreeOf(doc.documentElement);
  if (!spTree) throw new EditError('notes part is malformed (no spTree)');
  const shapes = flattenShapes(walkShapes(spTree));
  const body = shapes.find((s) => s.phType === 'body');
  if (!body?.txBody) throw new EditError('notes part has no body placeholder');

  // Template formatting: first paragraph pPr and first run rPr before clearing.
  const oldParas = childrenByName(body.txBody, 'a', 'p');
  const before = buildTextBody(body.txBody).map((p) => p.text).join('\n');
  const templateP = oldParas[0] ?? null;
  const pPr = templateP ? firstChildByName(templateP, 'a', 'pPr') : null;
  const templateRun = templateP ? (firstDescendantByName(templateP, 'a', 'r') ?? null) : null;
  const rPr = templateRun ? firstChildByName(templateRun, 'a', 'rPr') : null;

  for (const p of oldParas) body.txBody.removeChild(p);
  for (const line of text.split('\n')) {
    const p = doc.createElementNS(NS.a, 'a:p');
    if (pPr) p.appendChild(pPr.cloneNode(true));
    const r = doc.createElementNS(NS.a, 'a:r');
    if (rPr) r.appendChild(rPr.cloneNode(true));
    r.appendChild(makeElement(doc, 'a', 't', line));
    p.appendChild(r);
    body.txBody.appendChild(p);
  }
  pkg.zip.markDirty(partPath);
  return { partPath, before, after: text };
}

/** Replace the full text of a table cell (shape addressed by its cNvPr id). */
export function setTableCellText(
  pkg: PptxPackage,
  slideNumber: number,
  shapeId: number,
  row: number,
  col: number,
  text: string,
): { before: string; after: string } {
  const slide = pkg.slideByIndex(slideNumber - 1);
  if (!slide) throw new EditError(`slide ${slideNumber} does not exist`);
  const doc = pkg.zip.doc(slide.partPath);
  const spTree = spTreeOf(doc.documentElement);
  if (!spTree) throw new EditError('slide is malformed (no spTree)');
  const shapes = flattenShapes(walkShapes(spTree));
  const shape = shapes.find((s) => s.id === shapeId);
  if (!shape) throw new EditError(`shape id ${shapeId} not found on slide ${slideNumber}`);
  if (!shape.table) throw new EditError(`shape id ${shapeId} is not a table`);
  const cell = shape.table.rows[row]?.[col];
  if (!cell) throw new EditError(`cell r${row}c${col} out of range (table is ${shape.table.rowCount}x${shape.table.colCount})`);
  if (!cell.txBody) throw new EditError(`cell r${row}c${col} has no text body`);

  const oldParas = childrenByName(cell.txBody, 'a', 'p');
  const before = buildTextBody(cell.txBody).map((p) => p.text).join('\n');
  const templateP = oldParas[0] ?? null;
  const pPr = templateP ? firstChildByName(templateP, 'a', 'pPr') : null;
  const templateRun = templateP ? firstDescendantByName(templateP, 'a', 'r') : null;
  const rPr = templateRun ? firstChildByName(templateRun, 'a', 'rPr') : null;

  for (const p of oldParas) cell.txBody.removeChild(p);
  for (const line of text.split('\n')) {
    const p = doc.createElementNS(NS.a, 'a:p');
    if (pPr) p.appendChild(pPr.cloneNode(true));
    const r = doc.createElementNS(NS.a, 'a:r');
    if (rPr) r.appendChild(rPr.cloneNode(true));
    r.appendChild(makeElement(doc, 'a', 't', line));
    p.appendChild(r);
    cell.txBody.appendChild(p);
  }
  pkg.zip.markDirty(slide.partPath);
  return { before, after: text };
}
