import { PptxPackage } from '../package/presentation.js';
import { flattenShapes, spTreeOf, walkShapes, type ShapeRef } from '../package/shapes.js';
import { assignAnchors } from '../text/anchors.js';
import { buildTextBody, type ParagraphInfo } from '../text/model.js';

export interface WalkedParagraph {
  /** 0-based slide position (display order). */
  slideIndex: number;
  /** 1-based slide number (what tools take/return). */
  slideNumber: number;
  slidePart: string;
  shape: ShapeRef;
  /** Stable per-slide shape key used in anchor derivation, e.g. `sp#2`. */
  shapeKey: string;
  para: ParagraphInfo;
  anchor: string;
}

/**
 * Flatten a slide to its editable paragraphs (slide shapes including groups
 * and table cells), each with a stable `_bk_*` anchor. Table cells are walked
 * as `tbl#<id>/r<row>c<col>` contexts.
 */
export function walkSlideParagraphs(pkg: PptxPackage, slidePartPath: string): WalkedParagraph[] {
  const doc = pkg.zip.doc(slidePartPath);
  const spTree = spTreeOf(doc.documentElement);
  if (!spTree) return [];
  const shapes = flattenShapes(walkShapes(spTree));
  const allParas: { shapeKey: string; shape: ShapeRef; para: ParagraphInfo }[] = [];
  for (const shape of shapes) {
    if (shape.txBody) {
      const paras = buildTextBody(shape.txBody);
      for (const para of paras) {
        allParas.push({ shapeKey: `${shape.kind}#${shape.id}`, shape, para });
      }
    }
    if (shape.table) {
      for (const row of shape.table.rows) {
        for (const cell of row) {
          if (!cell.txBody) continue;
          const paras = buildTextBody(cell.txBody);
          for (const para of paras) {
            allParas.push({
              shapeKey: `${shape.kind}#${shape.id}/r${cell.row}c${cell.col}`,
              shape,
              para,
            });
          }
        }
      }
    }
  }
  const anchors = assignAnchors(
    slidePartPath,
    allParas.map((p) => p.shapeKey),
    allParas.map((p) => p.para),
  );
  const slideRef = pkg.slideByPartPath(slidePartPath);
  return allParas.map((p, i) => ({
    slideIndex: slideRef?.index ?? -1,
    slideNumber: (slideRef?.index ?? -1) + 1,
    slidePart: slidePartPath,
    shape: p.shape,
    shapeKey: p.shapeKey,
    para: p.para,
    anchor: anchors[i],
  }));
}

/** All slide paragraphs across the deck, in display order. */
export function walkDeckParagraphs(pkg: PptxPackage): WalkedParagraph[] {
  return pkg.slides.flatMap((s) => walkSlideParagraphs(pkg, s.partPath));
}

/** Speaker-notes paragraphs of one slide (empty list if the slide has no notes part). */
export function walkNotesParagraphs(pkg: PptxPackage, slidePartPath: string): {
  part: string;
  items: { anchor: string; para: ParagraphInfo }[];
} {
  const notesPart = pkg.notesSlideForSlide(slidePartPath);
  if (!notesPart || !pkg.zip.hasPart(notesPart)) return { part: '', items: [] };
  const doc = pkg.zip.doc(notesPart);
  const spTree = spTreeOf(doc.documentElement);
  if (!spTree) return { part: notesPart, items: [] };
  const shapes = flattenShapes(walkShapes(spTree));
  const bodyShape = shapes.find((s) => s.phType === 'body') ?? null;
  if (!bodyShape?.txBody) return { part: notesPart, items: [] };
  const paras = buildTextBody(bodyShape.txBody);
  const anchors = assignAnchors(notesPart, [`${bodyShape.kind}#${bodyShape.id}`], paras);
  return { part: notesPart, items: paras.map((para, i) => ({ anchor: anchors[i], para })) };
}
