import { NS } from '../xml/namespaces.js';
import { childrenByName, firstChildByName, firstDescendantByName } from '../xml/parse.js';
import type { Element } from '../xml/parse.js';

/**
 * Shape-tree walking (`p:spTree`): shapes, pictures, tables (graphicFrame),
 * connectors, and recursive groups. Placeholder identity (type + idx) drives
 * the layout/master inheritance chain.
 */

export interface TableCellRef {
  el: Element;
  row: number;
  col: number;
  txBody: Element | null;
}

export interface TableRef {
  el: Element;
  rows: TableCellRef[][];
  rowCount: number;
  colCount: number;
}

export type ShapeKind = 'sp' | 'pic' | 'graphicFrame' | 'grpSp' | 'cxnSp';

export interface ShapeRef {
  kind: ShapeKind;
  el: Element;
  /** `p:cNvPr/@id` — unique within the slide's spTree (animations target it). */
  id: number;
  name: string;
  /** Placeholder type (`p:ph/@type`), null for non-placeholder shapes. */
  phType: string | null;
  /** Placeholder index (`p:ph/@idx`), null when absent. */
  phIdx: number | null;
  txBody: Element | null;
  table: TableRef | null;
  children: ShapeRef[];
  /** True for `p:sp` with `txBox="1"` (plain textbox, no placeholder). */
  isTextBox: boolean;
}

function nvPrOf(el: Element): Element | null {
  for (const nvName of ['nvSpPr', 'nvPicPr', 'nvGraphicFramePr', 'nvGrpSpPr', 'nvCxnSpPr']) {
    const nv = firstChildByName(el, 'p', nvName);
    if (nv) return firstChildByName(nv, 'p', 'cNvPr');
  }
  return null;
}

function phOf(el: Element): { type: string | null; idx: number | null } {
  const nv = ['nvSpPr', 'nvPicPr', 'nvGraphicFramePr', 'nvGrpSpPr', 'nvCxnSpPr']
    .map((name) => firstChildByName(el, 'p', name))
    .find((nv) => nv !== null);
  // Placeholder marker lives at nv*/nvPr/p:ph.
  const nvPr = nv ? firstChildByName(nv, 'p', 'nvPr') : null;
  const ph = nvPr ? firstChildByName(nvPr, 'p', 'ph') : null;
  if (!ph) return { type: null, idx: null };
  const type = ph.getAttribute('type') ?? null;
  const idxAttr = ph.getAttribute('idx');
  return { type, idx: idxAttr === null ? null : Number(idxAttr) };
}

function parseTable(graphicFrame: Element): TableRef | null {
  const tbl = firstDescendantByName(graphicFrame, 'a', 'tbl');
  if (!tbl) return null;
  const rows: TableCellRef[][] = [];
  childrenByName(tbl, 'a', 'tr').forEach((tr, rowIdx) => {
    const cells = childrenByName(tr, 'a', 'tc').map((tc, colIdx) => ({
      el: tc,
      row: rowIdx,
      col: colIdx,
      // Table cell text bodies live in the DrawingML namespace (a:txBody),
      // unlike slide shape bodies which are p:txBody.
      txBody: firstChildByName(tc, 'a', 'txBody'),
    }));
    rows.push(cells);
  });
  return {
    el: tbl,
    rows,
    rowCount: rows.length,
    colCount: rows[0]?.length ?? 0,
  };
}

function shapeRefOf(el: Element): ShapeRef {
  const localName = el.localName as ShapeKind;
  const cNvPr = nvPrOf(el);
  const id = Number(cNvPr?.getAttribute('id') ?? '0');
  const name = cNvPr?.getAttribute('name') ?? '';
  const { type, idx } = phOf(el);
  const txBody = firstChildByName(el, 'p', 'txBody');
  const cNvSpPr = firstChildByName(el, 'p', 'nvSpPr')
    ? firstChildByName(firstChildByName(el, 'p', 'nvSpPr') as Element, 'p', 'cNvSpPr')
    : null;
  const isTextBox = localName === 'sp' && cNvSpPr?.getAttribute('txBox') === '1';
  return {
    kind: localName,
    el,
    id,
    name,
    phType: type,
    phIdx: idx,
    txBody: txBody,
    table: localName === 'graphicFrame' ? parseTable(el) : null,
    children: [],
    isTextBox,
  };
}

const SHAPE_LOCALS = new Set(['sp', 'pic', 'graphicFrame', 'grpSp', 'cxnSp']);

/** Walk direct shape children of a container (`p:spTree` or `p:grpSp`), recursing into groups. */
export function walkShapes(container: Element): ShapeRef[] {
  const out: ShapeRef[] = [];
  let child = container.firstChild;
  while (child) {
    if (child.nodeType === 1) {
      const el = child as Element;
      if (el.namespaceURI === NS.p && SHAPE_LOCALS.has(el.localName ?? '')) {
        const ref = shapeRefOf(el);
        if (ref.kind === 'grpSp') ref.children = walkShapes(el);
        out.push(ref);
      }
    }
    child = child.nextSibling;
  }
  return out;
}

/** Flatten a shape tree (groups included) in document order. */
export function flattenShapes(shapes: ShapeRef[]): ShapeRef[] {
  const out: ShapeRef[] = [];
  for (const shape of shapes) {
    out.push(shape);
    if (shape.children.length > 0) out.push(...flattenShapes(shape.children));
  }
  return out;
}

/** Recursive lookup by `p:cNvPr/@id` within a slide's shape tree. */
export function findShapeById(spTree: Element, id: number): ShapeRef | null {
  return flattenShapes(walkShapes(spTree)).find((s) => s.id === id) ?? null;
}

/** The `p:spTree` of a slide/notesSlide/layout/master document. */
export function spTreeOf(doc: Element | null): Element | null {
  const cSld = firstChildByName(doc, 'p', 'cSld');
  return cSld ? firstChildByName(cSld, 'p', 'spTree') : null;
}
