import { NS } from '../xml/namespaces.js';
import { firstChildByName, makeElement, directText } from '../xml/parse.js';
import type { Document, Element } from '../xml/parse.js';
import { buildParagraph, type ParagraphInfo, type RunPart } from '../text/model.js';

/**
 * The run-aware text engine — the format-preservation core.
 *
 * A paragraph's logical text is spread across `a:r` runs (and `a:fld` fields).
 * This engine finds a match across run boundaries (with whitespace/quote
 * tolerance), splits boundary runs so untouched fragments keep their exact
 * `a:rPr`, and splices replacement runs cloned from the matched span's first
 * run — so replacement text inherits the original formatting.
 */

export class EditError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'EditError';
  }
}

const QUOTES: Record<string, string> = {
  '\u2018': "'", '\u2019': "'", '\u201A': "'", '\u201B': "'",
  '\u201C': '"', '\u201D': '"', '\u201E': '"', '\u201F': '"',
  '\u2013': '-', '\u2014': '-', '\u2212': '-',
  '\u00A0': ' ', '\u2007': ' ', '\u202F': ' ',
};

function normalizeChar(ch: string): string {
  if (QUOTES[ch] !== undefined) return QUOTES[ch];
  return /\s/.test(ch) ? ' ' : ch;
}

export interface Span {
  start: number;
  end: number;
}

interface RawToNorm {
  norm: string;
  /** For each normalized char, its raw index. */
  map: number[];
}

function normalizeWithMap(text: string): RawToNorm {
  let norm = '';
  const map: number[] = [];
  for (let i = 0; i < text.length; i++) {
    const n = normalizeChar(text[i]);
    if (n === ' ' && norm.endsWith(' ')) continue; // collapse whitespace
    norm += n;
    map.push(i);
  }
  return { norm, map };
}

function findOccurrences(haystack: string, needle: string): number[] {
  const out: number[] = [];
  let idx = haystack.indexOf(needle);
  while (idx !== -1) {
    out.push(idx);
    idx = haystack.indexOf(needle, idx + 1);
  }
  return out;
}

/**
 * Find the Nth occurrence of `needle` in the paragraph text, first exactly,
 * then with whitespace/smart-quote tolerance. Returns raw-char span.
 */
export function findMatchSpan(text: string, needle: string, occurrence = 1): Span | null {
  if (needle.length === 0) throw new EditError('old_string must not be empty');
  const exact = findOccurrences(text, needle);
  if (exact.length >= occurrence) {
    const start = exact[occurrence - 1];
    return { start, end: start + needle.length };
  }
  const h = normalizeWithMap(text);
  const n = normalizeWithMap(needle);
  const normHits = findOccurrences(h.norm, n.norm);
  if (normHits.length >= occurrence) {
    const start = h.map[normHits[occurrence - 1]];
    const lastNorm = normHits[occurrence - 1] + n.norm.length - 1;
    const end = h.map[lastNorm] + 1;
    return { start, end };
  }
  return null;
}

interface Seg {
  part: RunPart;
  /** Raw text range of this part within the paragraph. */
  s: number;
  e: number;
}

function segmentParagraph(p: ParagraphInfo): { segs: Seg[]; total: number } {
  const segs: Seg[] = [];
  let pos = 0;
  let child = p.el.firstChild;
  while (child) {
    if (child.nodeType === 1) {
      const el = child as Element;
      if (el.namespaceURI === NS.a && (el.localName === 'r' || el.localName === 'fld')) {
        const t = firstChildByName(el, 'a', 't');
        const text = t ? directText(t) : '';
        const isField = el.localName === 'fld';
        segs.push({ part: { kind: isField ? 'field' : 'run', el, textEl: t, text }, s: pos, e: pos + text.length });
        pos += text.length;
      } else if (el.namespaceURI === NS.a && el.localName === 'br') {
        segs.push({ part: { kind: 'run', el, textEl: null, text: '\n' }, s: pos, e: pos + 1 });
        pos += 1;
      }
    }
    child = child.nextSibling;
  }
  return { segs, total: pos };
}

function setTextContent(el: Element, text: string): void {
  const t = firstChildByName(el, 'a', 't');
  if (!t) return;
  while (t.firstChild) t.removeChild(t.firstChild);
  const ownerDoc = t.ownerDocument;
  if (text.length > 0 && ownerDoc) t.appendChild(ownerDoc.createTextNode(text));
}

function cloneRunWithText(doc: Document, templateRun: Element, text: string): Element {
  const r = doc.createElementNS(NS.a, 'a:r');
  const rPr = firstChildByName(templateRun, 'a', 'rPr');
  if (rPr) r.appendChild(rPr.cloneNode(true));
  const t = makeElement(doc, 'a', 't', text);
  r.appendChild(t);
  return r;
}

/** Split a replacement string on '\n' into runs joined by `a:br`, all sharing the template run's rPr. */
function buildReplacementNodes(doc: Document, templateRun: Element, replacement: string): Element[] {
  const chunks = replacement.split('\n');
  const nodes: Element[] = [];
  chunks.forEach((chunk, i) => {
    if (i > 0) {
      const br = doc.createElementNS(NS.a, 'a:br');
      const brRPr = firstChildByName(templateRun, 'a', 'rPr');
      if (brRPr) br.appendChild(brRPr.cloneNode(true));
      nodes.push(br);
    }
    if (chunk.length > 0) nodes.push(cloneRunWithText(doc, templateRun, chunk));
  });
  if (nodes.length === 0) nodes.push(cloneRunWithText(doc, templateRun, ''));
  return nodes;
}

/**
 * Replace raw-char span [start,end) with `replacement`, splitting boundary
 * runs so untouched fragments keep their formatting. `a:br` segments at span
 * edges are treated as atomic single chars.
 */
export function spliceParagraph(
  doc: Document,
  p: ParagraphInfo,
  start: number,
  end: number,
  replacement: string,
): void {
  const { segs } = segmentParagraph(p);
  const affected = segs.filter((seg) => seg.s < end && seg.e > start);
  if (affected.length === 0) {
    throw new EditError('match span overlaps no text runs');
  }
  // Fields are atomic: refuse partial overlap of an a:fld.
  for (const seg of affected) {
    if (
      seg.part.kind === 'field' &&
      ((seg.s < start && seg.e > start) || (seg.s < end && seg.e > end))
    ) {
      throw new EditError(
        'old_string partially overlaps a field (a:fld) — match the whole field text or none of it',
      );
    }
  }

  const first = affected[0];
  const last = affected[affected.length - 1];
  const leftKeeps = first.s < start;
  const rightKeeps = last.e > end;
  const leftText = leftKeeps ? first.part.text.slice(0, start - first.s) : '';
  const rightText = rightKeeps ? last.part.text.slice(end - last.s) : '';

  const parent = first.part.el.parentNode;
  if (!parent) throw new EditError('paragraph runs are detached');
  // Capture the node following the last affected run before removals — if every
  // affected run is removed, insertion goes there (may be null → append).
  const afterLast = last.part.el.nextSibling;

  // Remove runs fully covered by the span.
  for (const seg of affected) {
    if (seg.s >= start && seg.e <= end) {
      seg.part.el.parentNode?.removeChild(seg.part.el);
    }
  }
  // Trim boundary runs (fragments outside the span stay with original formatting).
  if (leftKeeps) setTextContent(first.part.el, leftText);
  if (rightKeeps) {
    if (last !== first) setTextContent(last.part.el, rightText);
    else if (!leftKeeps) setTextContent(first.part.el, rightText); // match began at this run's start; keep only its tail
  }

  // Template run: leftmost affected run that carries formatting.
  const templateSeg =
    affected.find((seg) => seg.part.kind === 'run' && firstChildByName(seg.part.el, 'a', 'rPr')) ??
    affected[0];

  // Insertion point: after the surviving left fragment, else before the first
  // surviving affected run, else where the last removed run used to be.
  const insertRef = leftKeeps
    ? first.part.el.nextSibling
    : (affected.find((seg) => !(seg.s >= start && seg.e <= end))?.part.el ?? afterLast);

  if (rightKeeps && last === first && leftKeeps) {
    // One run spans the whole match: keep left fragment, insert replacement, then a clone for the right fragment.
    const nodes = buildReplacementNodes(doc, templateSeg.part.el, replacement);
    const rightNode = cloneRunWithText(doc, templateSeg.part.el, rightText);
    for (const node of nodes) parent.insertBefore(node, insertRef);
    parent.insertBefore(rightNode, insertRef);
  } else {
    const nodes = buildReplacementNodes(doc, templateSeg.part.el, replacement);
    for (const node of nodes) parent.insertBefore(node, insertRef);
  }
}

/** Split the run covering raw position `pos` into two runs (cloning rPr), so span edges align with run boundaries. */
function splitAt(doc: Document, p: ParagraphInfo, pos: number): void {
  const { segs } = segmentParagraph(p);
  const seg = segs.find((s) => s.s < pos && s.e > pos);
  if (!seg || seg.part.kind !== 'run') return;
  const parent = seg.part.el.parentNode;
  if (!parent) return;
  const leftText = seg.part.text.slice(0, pos - seg.s);
  const rightText = seg.part.text.slice(pos - seg.s);
  setTextContent(seg.part.el, leftText);
  const clone = cloneRunWithText(doc, seg.part.el, rightText);
  parent.insertBefore(clone, seg.part.el.nextSibling);
}

/** Ensure run boundaries exist at raw positions `start` and `end`; return runs fully inside the span. */
export function runsInSpan(doc: Document, p: ParagraphInfo, start: number, end: number): Element[] {
  splitAt(doc, p, start);
  splitAt(doc, p, end);
  const { segs } = segmentParagraph(buildParagraph(p.el, p.index));
  return segs
    .filter((seg) => seg.s >= start && seg.e <= end && seg.part.kind === 'run')
    .map((seg) => seg.part.el);
}

/** Rebuild paragraph info after DOM mutation. */
export function refreshParagraph(el: Element, index: number): ParagraphInfo {
  return buildParagraph(el, index);
}

/** Paragraph text, re-read from the DOM. */
export function paragraphTextNow(el: Element): string {
  return buildParagraph(el, 0).text;
}

/** True if the element is a DrawingML run (a:r). */
export function isRunEl(el: Element): boolean {
  return el.namespaceURI === NS.a && el.localName === 'r';
}

export { findOccurrences };
