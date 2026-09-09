import { NS } from '../xml/namespaces.js';
import { childrenByName, firstChildByName, firstDescendantByName, makeElement, parseXml, serializeXml } from '../xml/parse.js';
import type { Document, Element } from '../xml/parse.js';
import { CT } from '../package/content_types.js';
import { PRESENTATION_PART } from '../package/presentation.js';
import type { PptxPackage } from '../package/presentation.js';
import { walkShapes, spTreeOf } from '../package/shapes.js';
import { addRelToDoc, removeRelFromDoc, relsPathFor } from '../package/rels.js';
import { EditError } from './text_engine.js';

/**
 * Slide-level structural operations. Every operation maintains OPC package
 * integrity: sldIdLst order, presentation relationships, slide rels, and
 * [Content_Types].xml overrides.
 */

const REL_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function maxSlideNumber(pkg: PptxPackage): number {
  let max = 0;
  for (const path of pkg.zip.listParts()) {
    const m = path.match(/^ppt\/slides\/slide(\d+)\.xml$/);
    if (m) max = Math.max(max, Number(m[1]));
  }
  return max;
}

function nextFreeTopLevelId(pkg: PptxPackage): number {
  const presDoc = pkg.presentationDoc();
  let max = 255;
  for (const listName of ['sldIdLst', 'sldMasterIdLst', 'notesMasterIdLst'] as const) {
    const list = childrenByName(presDoc.documentElement, 'p', listName)[0];
    if (!list) continue;
    for (const el of list.childNodes) {
      if (el.nodeType === 1) {
        const v = Number((el as Element).getAttribute('id'));
        if (Number.isFinite(v)) max = Math.max(max, v);
      }
    }
  }
  return max + 1;
}

function updatePresentationRels(pkg: PptxPackage, slidePartPath: string | null, action: 'add' | 'remove', relId?: string): string {
  const relsPath = relsPathFor(PRESENTATION_PART);
  const doc = pkg.zip.doc(relsPath);
  if (action === 'add') {
    const id = addRelToDoc(doc, {
      type: `${REL_BASE}/slide`,
      target: slidePartPath?.replace(/^ppt\//, '') ?? '',
    });
    pkg.zip.markDirty(relsPath);
    return id;
  }
  if (relId) {
    removeRelFromDoc(doc, relId);
    pkg.zip.markDirty(relsPath);
  }
  return relId ?? '';
}

function setOverrideAndMark(pkg: PptxPackage, partPath: string, ct: string | null): void {
  if (ct === null) pkg.contentTypes.removeOverride(partPath);
  else pkg.contentTypes.setOverride(partPath, ct);
  pkg.zip.markDirty('[Content_Types].xml');
}

function refreshSlideList(pkg: PptxPackage): void {
  pkg.zip.markDirty(PRESENTATION_PART);
  pkg.refreshStructure();
}

/**
 * Scrub a deleted slide's id out of p14:sectionLst (presentation extLst).
 * Sections reference slide ids; a dangling entry triggers PowerPoint's
 * repair prompt. Sections left with zero slides are removed, matching
 * PowerPoint's own behavior.
 */
export function scrubSectionsForSlide(presDoc: Document, sldId: number): number {
  let removed = 0;
  const extLst = firstChildByName(presDoc.documentElement, 'p', 'extLst');
  if (!extLst) return 0;
  for (const ext of childrenByName(extLst, 'p', 'ext')) {
    const sectionLst = firstChildByName(ext, 'p14', 'sectionLst');
    if (!sectionLst) continue;
    for (const section of [...childrenByName(sectionLst, 'p14', 'section')]) {
      const idLst = firstChildByName(section, 'p14', 'sldIdLst');
      if (!idLst) continue;
      for (const entry of [...childrenByName(idLst, 'p14', 'sldId')]) {
        if (Number(entry.getAttribute('id')) === sldId) {
          idLst.removeChild(entry);
          removed += 1;
        }
      }
      if (childrenByName(idLst, 'p14', 'sldId').length === 0) {
        sectionLst.removeChild(section);
      }
    }
  }
  return removed;
}

/** Build a new, empty slide document from a layout's placeholder prototypes. */
function buildSlideFromLayout(pkg: PptxPackage, layoutPartPath: string): Document {
  const layoutDoc = pkg.zip.doc(layoutPartPath);
  const layoutTree = spTreeOf(layoutDoc.documentElement);
  const skeleton = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}">
<p:cSld><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
</p:spTree></p:cSld>
<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`;
  const doc = parseXml(skeleton, 'new-slide');
  const tree = spTreeOf(doc.documentElement);
  if (!tree) throw new EditError('internal: slide skeleton malformed');

  let nextId = 2;
  if (layoutTree) {
    for (const layoutSp of walkShapes(layoutTree)) {
      if (layoutSp.kind !== 'sp' || layoutSp.phType === null) continue;
      const clone = doc.importNode(layoutSp.el, true) as Element;
      // Fresh shape id within the new slide.
      const cNvPr = firstDescendantByName(clone, 'p', 'cNvPr');
      cNvPr?.setAttribute('id', String(nextId++));
      // Empty the placeholder text: keep pPr and endParaRPr only.
      const txBody = firstChildByName(clone, 'p', 'txBody');
      if (txBody) {
        for (const p of childrenByName(txBody, 'a', 'p')) {
          for (const child of [...p.childNodes]) {
            if (child.nodeType === 1) {
              const el = child as Element;
              if (el.localName === 'pPr' || el.localName === 'endParaRPr') continue;
              p.removeChild(child);
            }
          }
        }
      }
      tree.appendChild(clone);
    }
  }
  // Copy the layout's color-map override when present.
  const layoutClr = firstChildByName(layoutDoc.documentElement, 'p', 'clrMapOvr');
  const slideClr = firstChildByName(doc.documentElement, 'p', 'clrMapOvr');
  if (layoutClr && slideClr && doc.documentElement) {
    doc.documentElement.replaceChild(doc.importNode(layoutClr, true), slideClr);
  }
  return doc;
}

export interface SlideOpResult {
  slideNumber: number;
  partPath: string;
  relId?: string;
}

/** Append a new empty slide instantiated from `layoutPartPath` (default: the last slide's layout). */
export function addSlide(pkg: PptxPackage, layoutPartPath?: string): SlideOpResult {
  const layout =
    layoutPartPath ??
    (pkg.slides.length > 0 ? pkg.layoutForSlide(pkg.slides[pkg.slides.length - 1].partPath) : null) ??
    pkg.partRefs.slideLayouts[0] ??
    null;
  if (!layout || !pkg.zip.hasPart(layout)) {
    throw new EditError('no slide layout available to instantiate — pass layout_part');
  }
  const n = maxSlideNumber(pkg) + 1;
  const partPath = `ppt/slides/slide${n}.xml`;
  const doc = buildSlideFromLayout(pkg, layout);

  pkg.zip.addPart(partPath, serializePartBytes(doc));
  pkg.zip.addPart(relsPathFor(partPath), serializePartBytes(
    parseXml(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL_BASE}/slideLayout" Target="${layout.replace(/^ppt\//, '../')}"/></Relationships>`,
      'new-slide-rels',
    ),
  ));
  setOverrideAndMark(pkg, partPath, CT.slide);

  const relId = updatePresentationRels(pkg, partPath, 'add');
  const presDoc = pkg.presentationDoc();
  const sldIdLst = childrenByName(presDoc.documentElement, 'p', 'sldIdLst')[0];
  if (!sldIdLst) throw new EditError('presentation.xml has no sldIdLst');
  const sldId = makeElement(presDoc, 'p', 'sldId');
  sldId.setAttribute('id', String(nextFreeTopLevelId(pkg)));
  sldId.setAttributeNS(NS.r, 'r:id', relId);
  sldIdLst.appendChild(sldId);
  refreshSlideList(pkg);

  return { slideNumber: pkg.slides.length, partPath, relId };
}

/** Duplicate a slide (layout/media rels are shared; notes and comments are not copied). */
export function duplicateSlide(pkg: PptxPackage, slideNumber: number): SlideOpResult {
  const source = pkg.slideByIndex(slideNumber - 1);
  if (!source) throw new EditError(`slide ${slideNumber} does not exist`);
  const n = maxSlideNumber(pkg) + 1;
  const partPath = `ppt/slides/slide${n}.xml`;

  const sourceBytes = pkg.zip.originalBytes(source.partPath);
  if (!sourceBytes) throw new EditError(`source slide bytes unavailable: ${source.partPath}`);
  pkg.zip.addPart(partPath, Buffer.from(sourceBytes));

  // Clone slide rels, dropping slide-scoped parts (notes, comments) and rels entries.
  const sourceRelsPath = relsPathFor(source.partPath);
  const SLIDE_SCOPED = ['notesSlide', 'threadedComments', 'comments'];
  if (pkg.zip.hasPart(sourceRelsPath)) {
    const relsDoc = parseXml(pkg.zip.text(sourceRelsPath) ?? '', sourceRelsPath);
    const root = relsDoc.documentElement;
    const kept: string[] = [];
    for (const el of childrenByName(root, 'rel', 'Relationship')) {
      const type = el.getAttribute('Type') ?? '';
      if (SLIDE_SCOPED.some((s) => type.endsWith(`/${s}`))) {
        root?.removeChild(el);
      } else {
        kept.push(el.getAttribute('Id') ?? '');
      }
    }
    pkg.zip.addPart(relsPathFor(partPath), serializePartBytes(relsDoc));
    void kept;
  }

  setOverrideAndMark(pkg, partPath, CT.slide);
  const relId = updatePresentationRels(pkg, partPath, 'add');

  const presDoc = pkg.presentationDoc();
  const sldIdLst = childrenByName(presDoc.documentElement, 'p', 'sldIdLst')[0];
  const sourceSldId = childrenByName(sldIdLst, 'p', 'sldId')[source.index];
  const sldId = makeElement(presDoc, 'p', 'sldId');
  sldId.setAttribute('id', String(nextFreeTopLevelId(pkg)));
  sldId.setAttributeNS(NS.r, 'r:id', relId);
  sourceSldId.parentNode?.insertBefore(sldId, sourceSldId.nextSibling);
  refreshSlideList(pkg);

  return { slideNumber: source.index + 2, partPath, relId };
}

/** Reorder slides; `order` is a permutation of current 1-based slide numbers. */
export function reorderSlides(pkg: PptxPackage, order: number[]): void {
  const n = pkg.slides.length;
  if (order.length !== n) throw new EditError(`order must list all ${n} slides (got ${order.length})`);
  const sorted = [...order].sort((a, b) => a - b);
  if (sorted.some((v, i) => v !== i + 1)) {
    throw new EditError(`order must be a permutation of 1..${n}: got [${order.join(', ')}]`);
  }
  const presDoc = pkg.presentationDoc();
  const sldIdLst = childrenByName(presDoc.documentElement, 'p', 'sldIdLst')[0];
  if (!sldIdLst) throw new EditError('presentation.xml has no sldIdLst');
  const current = childrenByName(sldIdLst, 'p', 'sldId');
  for (const position of order) {
    const el = current[position - 1];
    sldIdLst.appendChild(el); // re-append moves the node to the end in order
  }
  refreshSlideList(pkg);
}

/** Delete a slide and its slide-scoped parts (notes, comments) with full cleanup. */
export function deleteSlide(pkg: PptxPackage, slideNumber: number): { removedParts: string[] } {
  if (pkg.slides.length <= 1) throw new EditError('refusing to delete the last slide in the deck');
  const slide = pkg.slideByIndex(slideNumber - 1);
  if (!slide) throw new EditError(`slide ${slideNumber} does not exist`);
  const removedParts: string[] = [];

  // Remove sldId + presentation rel (+ section membership).
  const presDoc = pkg.presentationDoc();
  const sldIdLst = childrenByName(presDoc.documentElement, 'p', 'sldIdLst')[0];
  const sldIdEl = childrenByName(sldIdLst, 'p', 'sldId')[slide.index];
  const deletedSldId = Number(sldIdEl?.getAttribute('id'));
  if (sldIdEl) sldIdLst.removeChild(sldIdEl);
  if (Number.isFinite(deletedSldId)) scrubSectionsForSlide(presDoc, deletedSldId);
  updatePresentationRels(pkg, null, 'remove', slide.relId);
  refreshSlideList(pkg);

  // Collect slide-scoped parts via the slide's rels before removing them.
  const slideRelsPath = relsPathFor(slide.partPath);
  const scopedTargets: string[] = [];
  if (pkg.zip.hasPart(slideRelsPath)) {
    const relsDoc = parseXml(pkg.zip.text(slideRelsPath) ?? '', slideRelsPath);
    for (const el of childrenByName(relsDoc.documentElement, 'rel', 'Relationship')) {
      const type = el.getAttribute('Type') ?? '';
      const target = el.getAttribute('Target') ?? '';
      if (['notesSlide', 'threadedComments', 'comments'].some((s) => type.endsWith(`/${s}`))) {
        const dir = slide.partPath.slice(0, slide.partPath.lastIndexOf('/'));
        const segments: string[] = [];
        for (const seg of `${dir}/${target}`.split('/')) {
          if (seg === '' || seg === '.') continue;
          if (seg === '..') segments.pop();
          else segments.push(seg);
        }
        const partPath = segments.join('/');
        if (pkg.zip.hasPart(partPath)) scopedTargets.push(partPath);
      }
    }
  }

  // Remove the slide itself.
  pkg.zip.removePart(slide.partPath);
  removedParts.push(slide.partPath);
  if (pkg.zip.hasPart(slideRelsPath)) {
    pkg.zip.removePart(slideRelsPath);
    removedParts.push(slideRelsPath);
  }
  setOverrideAndMark(pkg, slide.partPath, null);

  // Remove slide-scoped parts + their rels + overrides.
  for (const partPath of scopedTargets) {
    const relsPart = relsPathFor(partPath);
    pkg.zip.removePart(partPath);
    removedParts.push(partPath);
    if (pkg.zip.hasPart(relsPart)) {
      pkg.zip.removePart(relsPart);
      removedParts.push(relsPart);
    }
    setOverrideAndMark(pkg, partPath, null);
  }
  return { removedParts };
}

function serializePartBytes(doc: Document): Buffer {
  let xml = serializeXml(doc);
  if (xml.startsWith('<?xml')) {
    xml = xml.slice(xml.indexOf('?>') + 2).replace(/^[\r\n]+/, '');
  }
  return Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n${xml}`, 'utf8');
}
