import { NS } from '../xml/namespaces.js';
import { firstChildByName, makeElement } from '../xml/parse.js';
import type { Document, Element } from '../xml/parse.js';
import type { PptxPackage } from '../package/presentation.js';
import { walkSlideParagraphs, walkNotesParagraphs } from '../read/walk.js';
import type { ParagraphInfo, RunPart } from '../text/model.js';
import { buildParagraph } from '../text/model.js';
import { findMatchSpan, runsInSpan, spliceParagraph, EditError } from './text_engine.js';

/**
 * Paragraph-level edit operations, addressed by stable `_bk_*` anchors.
 * Every operation returns before/after text for the audit log and marks the
 * owning part dirty.
 */

export interface ParagraphLocation {
  partPath: string;
  slideNumber: number | null;
  isNotes: boolean;
  shapeKey: string;
  el: Element;
  info: ParagraphInfo;
}

/** Resolve a paragraph anchor across all slides and speaker notes. */
export function resolveParagraph(pkg: PptxPackage, anchor: string): ParagraphLocation {
  if (!anchor.startsWith('_bk_')) {
    throw new EditError(
      `anchor "${anchor}" is not a paragraph anchor — use read_file/grep to get a _bk_* id`,
    );
  }
  for (const slide of pkg.slides) {
    for (const item of walkSlideParagraphs(pkg, slide.partPath)) {
      if (item.anchor === anchor) {
        return {
          partPath: slide.partPath,
          slideNumber: slide.index + 1,
          isNotes: false,
          shapeKey: item.shapeKey,
          el: item.para.el,
          info: item.para,
        };
      }
    }
  }
  for (const slide of pkg.slides) {
    const notes = walkNotesParagraphs(pkg, slide.partPath);
    for (const item of notes.items) {
      if (item.anchor === anchor) {
        return {
          partPath: notes.part,
          slideNumber: slide.index + 1,
          isNotes: true,
          shapeKey: 'notes',
          el: item.para.el,
          info: item.para,
        };
      }
    }
  }
  throw new EditError(
    `anchor ${anchor} not found — the deck may have changed; re-run read_file or grep`,
  );
}

function txBodyOf(paraEl: Element): Element | null {
  const parent = paraEl.parentNode;
  return parent && parent.nodeType === 1 ? (parent as Element) : null;
}

/** Clear stale autofit scaling after content changes so PowerPoint re-lays out. */
function resetAutofitCache(txBody: Element | null): boolean {
  if (!txBody) return false;
  const bodyPr = firstChildByName(txBody, 'a', 'bodyPr');
  if (!bodyPr) return false;
  let changed = false;
  const fit = ['normAutofit', 'spAutoFit']
    .map((name) => firstChildByName(bodyPr, 'a', name))
    .find((el) => el !== null);
  if (fit && fit.localName === 'normAutofit') {
    for (const attr of ['fontScale', 'lnSpcReduction']) {
      if (fit.getAttribute(attr) !== null) {
        fit.removeAttribute(attr);
        changed = true;
      }
    }
  }
  return changed;
}

function markEdited(pkg: PptxPackage, partPath: string): void {
  pkg.zip.markDirty(partPath);
}

/** Re-derive the content anchor of a paragraph element after an edit changed its text. */
function findAnchorByParagraph(pkg: PptxPackage, partPath: string, el: Element): string | null {
  if (pkg.slideByPartPath(partPath)) {
    for (const item of walkSlideParagraphs(pkg, partPath)) {
      if (item.para.el === el) return item.anchor;
    }
    return null;
  }
  const notes = walkNotesParagraphs(pkg, partPath);
  for (const item of notes.items) {
    if (item.para.el === el) return item.anchor;
  }
  return null;
}

export interface ReplaceResult {
  partPath: string;
  slideNumber: number | null;
  before: string;
  after: string;
  replacements: number;
  /** Anchor of the edited paragraph AFTER the edit (anchors are content-derived). */
  new_anchor: string | null;
}

export function replaceInParagraphByAnchor(
  pkg: PptxPackage,
  anchor: string,
  oldString: string,
  newString: string,
  options: { occurrence?: number; all?: boolean } = {},
): ReplaceResult {
  const loc = resolveParagraph(pkg, anchor);
  const before = loc.info.text;
  let replacements = 0;
  if (options.all === true) {
    // Loop: refresh after each splice; search only text at/after the last edit
    // plus the replacement length so replacements containing old_string don't loop forever.
    let searchFrom = 0;
    for (;;) {
      const current = buildParagraph(loc.el, 0);
      if (searchFrom > current.text.length) break;
      const haystack = current.text.slice(searchFrom);
      const span = findMatchSpan(haystack, oldString, 1);
      if (!span) break;
      spliceParagraph(pkg.zip.doc(loc.partPath), current, searchFrom + span.start, searchFrom + span.end, newString);
      searchFrom += span.end + newString.length;
      replacements += 1;
    }
    if (replacements === 0) {
      throw new EditError(`old_string not found in paragraph: ${JSON.stringify(oldString)}`);
    }
  } else {
    const current = buildParagraph(loc.el, 0);
    const span = findMatchSpan(current.text, oldString, options.occurrence ?? 1);
    if (!span) {
      throw new EditError(`old_string not found in paragraph: ${JSON.stringify(oldString)}`);
    }
    spliceParagraph(pkg.zip.doc(loc.partPath), current, span.start, span.end, newString);
    replacements = 1;
  }
  const after = buildParagraph(loc.el, 0).text;
  resetAutofitCache(txBodyOf(loc.el));
  markEdited(pkg, loc.partPath);
  return {
    partPath: loc.partPath,
    slideNumber: loc.slideNumber,
    before,
    after,
    replacements,
    new_anchor: findAnchorByParagraph(pkg, loc.partPath, loc.el),
  };
}

export interface InsertResult {
  partPath: string;
  slideNumber: number | null;
  /** Anchors of the newly created paragraphs (content-derived). */
  created_anchors: (string | null)[];
  before: string;
  after: string;
}

function firstRunEl(paraEl: Element): Element | null {
  let child = paraEl.firstChild;
  while (child) {
    if (child.nodeType === 1) {
      const el = child as Element;
      if (el.namespaceURI === NS.a && el.localName === 'r') return el;
    }
    child = child.nextSibling;
  }
  return null;
}

export function insertParagraphByAnchor(
  pkg: PptxPackage,
  anchor: string,
  position: 'before' | 'after',
  text: string,
  styleSourceAnchor?: string,
): InsertResult {
  const loc = resolveParagraph(pkg, anchor);
  const doc = pkg.zip.doc(loc.partPath);
  const templateLoc = styleSourceAnchor ? resolveParagraph(pkg, styleSourceAnchor) : null;
  const templatePara = templateLoc ? templateLoc.el : loc.el;

  const paras: string[] = text.split('\n\n');
  const anchorParent = loc.el.parentNode;
  if (!anchorParent) throw new EditError('anchor paragraph is detached');
  let refEl = position === 'before' ? loc.el : loc.el.nextSibling;
  const created: Element[] = [];
  for (const paraText of paras) {
    const newP = doc.createElementNS(NS.a, 'a:p');
    const pPr = firstChildByName(templatePara, 'a', 'pPr');
    if (pPr) newP.appendChild(pPr.cloneNode(true));
    if (paraText.length === 0) {
      const endPr = firstChildByName(templatePara, 'a', 'endParaRPr');
      if (endPr) newP.appendChild(endPr.cloneNode(true));
    } else {
      const templateRun = firstRunEl(templatePara);
      const chunks = paraText.split('\n');
      chunks.forEach((chunk, i) => {
        if (i > 0) {
          const br = doc.createElementNS(NS.a, 'a:br');
          newP.appendChild(br);
        }
        const r = doc.createElementNS(NS.a, 'a:r');
        const rPr = templateRun ? firstChildByName(templateRun, 'a', 'rPr') : null;
        if (rPr) r.appendChild(rPr.cloneNode(true));
        r.appendChild(makeElement(doc, 'a', 't', chunk));
        newP.appendChild(r);
      });
    }
    anchorParent.insertBefore(newP, refEl);
    refEl = newP.nextSibling;
    created.push(newP);
  }
  resetAutofitCache(txBodyOf(loc.el));
  markEdited(pkg, loc.partPath);
  const created_anchors = created.map((el) => findAnchorByParagraph(pkg, loc.partPath, el));
  return {
    partPath: loc.partPath,
    slideNumber: loc.slideNumber,
    created_anchors,
    before: `${paras.length} paragraph(s) inserted ${position} anchor ${anchor}`,
    after: text,
  };
}

export type FontProp = 'b' | 'i' | 'u' | 'sz' | 'color';

export interface FontSpec {
  b?: 'on' | 'off';
  i?: 'on' | 'off';
  u?: 'on' | 'off';
  /** Font size in points (written as hundredths of a point). */
  sz_pt?: number;
  /** Hex RGB, e.g. 'FF0000'. */
  color_hex?: string;
}

function getOrCreateRPr(doc: Document, run: Element): Element {
  let rPr = firstChildByName(run, 'a', 'rPr');
  if (!rPr) {
    rPr = makeElement(doc, 'a', 'rPr');
    run.insertBefore(rPr, run.firstChild);
  }
  return rPr;
}

function applyFontSpec(doc: Document, run: Element, spec: FontSpec): void {
  const rPr = getOrCreateRPr(doc, run);
  if (spec.b) rPr.setAttribute('b', spec.b === 'on' ? '1' : '0');
  if (spec.i) rPr.setAttribute('i', spec.i === 'on' ? '1' : '0');
  if (spec.u) rPr.setAttribute('u', spec.u === 'on' ? 'sng' : 'none');
  if (spec.sz_pt !== undefined) {
    if (!(spec.sz_pt > 0 && spec.sz_pt <= 4000)) throw new EditError(`sz_pt out of range: ${spec.sz_pt}`);
    rPr.setAttribute('sz', String(Math.round(spec.sz_pt * 100)));
  }
  if (spec.color_hex !== undefined) {
    if (!/^[0-9A-Fa-f]{6}$/.test(spec.color_hex)) throw new EditError(`color_hex must be RRGGBB: ${spec.color_hex}`);
    for (const child of [...rPr.childNodes]) {
      if (child.nodeType === 1 && (child as Element).localName === 'solidFill') rPr.removeChild(child);
    }
    const fill = makeElement(doc, 'a', 'solidFill');
    fill.appendChild(makeElement(doc, 'a', 'srgbClr', undefined));
    const clr = fill.firstChild as Element;
    clr.setAttribute('val', spec.color_hex.toUpperCase());
    rPr.insertBefore(fill, rPr.firstChild);
  }
}

function clearProps(doc: Document, run: Element, props: FontProp[]): void {
  const rPr = firstChildByName(run, 'a', 'rPr');
  if (!rPr) return;
  if (props.includes('b')) rPr.removeAttribute('b');
  if (props.includes('i')) rPr.removeAttribute('i');
  if (props.includes('u')) rPr.removeAttribute('u');
  if (props.includes('sz')) rPr.removeAttribute('sz');
  if (props.includes('color')) {
    for (const child of [...rPr.childNodes]) {
      if (child.nodeType === 1 && (child as Element).localName === 'solidFill') rPr.removeChild(child);
    }
  }
}

export interface FontResult {
  partPath: string;
  slideNumber: number | null;
  runsTouched: number;
  before: string;
  after: string;
  new_anchor: string | null;
}

export function applyFontByAnchor(
  pkg: PptxPackage,
  anchor: string,
  spec: FontSpec,
  spanText?: string,
): FontResult {
  const loc = resolveParagraph(pkg, anchor);
  const doc = pkg.zip.doc(loc.partPath);
  const before = loc.info.text;
  let runs: Element[];
  if (spanText !== undefined) {
    const current = buildParagraph(loc.el, 0);
    const span = findMatchSpan(current.text, spanText, 1);
    if (!span) throw new EditError(`span_text not found in paragraph: ${JSON.stringify(spanText)}`);
    runs = runsInSpan(doc, current, span.start, span.end);
  } else {
    runs = loc.info.runs.filter((part: RunPart) => part.kind === 'run').map((part) => part.el);
  }
  if (runs.length === 0) throw new EditError('no text runs to format');
  for (const run of runs) applyFontSpec(doc, run, spec);
  resetAutofitCache(txBodyOf(loc.el));
  markEdited(pkg, loc.partPath);
  return {
    partPath: loc.partPath,
    slideNumber: loc.slideNumber,
    runsTouched: runs.length,
    before,
    after: buildParagraph(loc.el, 0).text,
    new_anchor: findAnchorByParagraph(pkg, loc.partPath, loc.el),
  };
}

export function clearFormattingByAnchor(
  pkg: PptxPackage,
  anchor: string,
  props: FontProp[],
  spanText?: string,
): FontResult {
  const loc = resolveParagraph(pkg, anchor);
  const doc = pkg.zip.doc(loc.partPath);
  const before = loc.info.text;
  let runs: Element[];
  if (spanText !== undefined) {
    const current = buildParagraph(loc.el, 0);
    const span = findMatchSpan(current.text, spanText, 1);
    if (!span) throw new EditError(`span_text not found in paragraph: ${JSON.stringify(spanText)}`);
    runs = runsInSpan(doc, current, span.start, span.end);
  } else {
    runs = loc.info.runs.filter((part) => part.kind === 'run').map((part) => part.el);
  }
  if (runs.length === 0) throw new EditError('no text runs to clear');
  const all = props.includes('b') && props.includes('i') && props.includes('u') && props.includes('sz') && props.includes('color');
  for (const run of runs) {
    if (all) {
      const rPr = firstChildByName(run, 'a', 'rPr');
      if (rPr) run.removeChild(rPr);
    } else {
      clearProps(doc, run, props);
    }
  }
  resetAutofitCache(txBodyOf(loc.el));
  markEdited(pkg, loc.partPath);
  return {
    partPath: loc.partPath,
    slideNumber: loc.slideNumber,
    runsTouched: runs.length,
    before,
    after: buildParagraph(loc.el, 0).text,
    new_anchor: findAnchorByParagraph(pkg, loc.partPath, loc.el),
  };
}
