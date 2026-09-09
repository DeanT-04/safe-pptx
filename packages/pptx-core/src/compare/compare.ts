import { NS } from '../xml/namespaces.js';
import { firstChildByName, makeElement } from '../xml/parse.js';
import type { Document, Element } from '../xml/parse.js';
import { PptxPackage } from '../package/presentation.js';
import { minimalSave } from '../save/minimal_save.js';
import { walkSlideParagraphs, walkNotesParagraphs } from '../read/walk.js';

/**
 * Redline-style deck comparison (pptx has no native tracked changes — this is
 * safe-pptx's equivalent of a docx redline). Slides are matched by title
 * first, then positionally; paragraphs within a matched slide pair are paired
 * by shape context + index, and text differences are word-diffed.
 */

export interface WordDiff {
  op: 'same' | 'ins' | 'del';
  text: string;
}

/** Word-level LCS diff, normalized on whitespace. */
export function wordDiff(before: string, after: string): WordDiff[] {
  const a = before.split(/\s+/).filter(Boolean);
  const b = after.split(/\s+/).filter(Boolean);
  const m = a.length;
  const n = b.length;
  const lcs: number[][] = Array.from({ length: m + 1 }, () => new Array<number>(n + 1).fill(0));
  for (let i = m - 1; i >= 0; i--) {
    for (let j = n - 1; j >= 0; j--) {
      lcs[i][j] = a[i] === b[j] ? lcs[i + 1][j + 1] + 1 : Math.max(lcs[i + 1][j], lcs[i][j + 1]);
    }
  }
  const out: WordDiff[] = [];
  const push = (op: WordDiff['op'], text: string): void => {
    const last = out[out.length - 1];
    if (last && last.op === op) last.text += ` ${text}`;
    else out.push({ op, text });
  };
  let i = 0;
  let j = 0;
  while (i < m && j < n) {
    if (a[i] === b[j]) {
      push('same', a[i]);
      i++; j++;
    } else if (lcs[i + 1][j] >= lcs[i][j + 1]) {
      push('del', a[i]); i++;
    } else {
      push('ins', b[j]); j++;
    }
  }
  while (i < m) push('del', a[i++]);
  while (j < n) push('ins', b[j++]);
  return out;
}

export interface SlideMatch {
  /** 1-based slide in deck A, null when added in B. */
  a: number | null;
  /** 1-based slide in deck B, null when removed from A. */
  b: number | null;
  titleA: string | null;
  titleB: string | null;
  status: 'matched' | 'added' | 'removed';
}

export interface ParaChange {
  slideA: number | null;
  slideB: number | null;
  shapeKey: string;
  paraIndex: number;
  kind: 'modified' | 'added' | 'removed';
  isNotes: boolean;
  before: string;
  after: string;
  diff: WordDiff[];
}

export interface DeckDiff {
  slideMatches: SlideMatch[];
  changes: ParaChange[];
  summary: {
    slidesA: number;
    slidesB: number;
    slidesAdded: number;
    slidesRemoved: number;
    paragraphsModified: number;
    paragraphsAdded: number;
    paragraphsRemoved: number;
    identical: boolean;
  };
}

function slideTitle(pkg: PptxPackage, slideNumber: number): string | null {
  const walked = walkSlideParagraphs(pkg, pkg.slideByIndex(slideNumber - 1)?.partPath ?? '');
  const title = walked.find((w) => w.shape.phType === 'title' || w.shape.phType === 'ctrTitle');
  if (title) return title.para.text;
  return walked[0]?.para.text ?? null;
}

interface FlatPara {
  shapeKey: string;
  paraIndex: number;
  text: string;
  isNotes: boolean;
  el?: Element;
}

function flattenSlideText(pkg: PptxPackage, slideNumber: number): FlatPara[] {
  const slide = pkg.slideByIndex(slideNumber - 1);
  if (!slide) return [];
  const out: FlatPara[] = [];
  for (const item of walkSlideParagraphs(pkg, slide.partPath)) {
    out.push({ shapeKey: item.shapeKey, paraIndex: item.para.index, text: item.para.text, isNotes: false, el: item.para.el });
  }
  const notes = walkNotesParagraphs(pkg, slide.partPath);
  for (const item of notes.items) {
    out.push({ shapeKey: 'notes', paraIndex: item.para.index, text: item.para.text, isNotes: true, el: item.para.el });
  }
  return out;
}

/** Match slides: unique exact-title matches first, then positional pairing of the remainders. */
export function matchSlides(pkgA: PptxPackage, pkgB: PptxPackage): SlideMatch[] {
  const nA = pkgA.slides.length;
  const nB = pkgB.slides.length;
  const titlesA = Array.from({ length: nA }, (_, i) => slideTitle(pkgA, i + 1));
  const titlesB = Array.from({ length: nB }, (_, i) => slideTitle(pkgB, i + 1));
  const usedA = new Set<number>();
  const usedB = new Set<number>();
  const result: SlideMatch[] = new Array(nA + nB);
  let write = 0;

  // Exact unique title matches.
  for (let i = 0; i < nA; i++) {
    const hits = titlesB
      .map((t, j) => ({ t, j }))
      .filter(({ t, j }) => j >= 0 && t === titlesA[i] && !usedB.has(j));
    const unique = hits.length === 1 && titlesB.filter((t) => t === titlesA[i]).length === 1;
    if (unique) {
      usedA.add(i);
      usedB.add(hits[0].j);
      result[write++] = { a: i + 1, b: hits[0].j + 1, titleA: titlesA[i], titleB: titlesB[hits[0].j], status: 'matched' };
    }
  }
  // Positional pairing of the leftovers.
  const restA = [...Array(nA).keys()].filter((i) => !usedA.has(i));
  const restB = [...Array(nB).keys()].filter((j) => !usedB.has(j));
  const pairCount = Math.min(restA.length, restB.length);
  for (let k = 0; k < pairCount; k++) {
    const i = restA[k];
    const j = restB[k];
    usedA.add(i);
    usedB.add(j);
    result[write++] = { a: i + 1, b: j + 1, titleA: titlesA[i], titleB: titlesB[j], status: 'matched' };
  }
  for (const j of restB.slice(pairCount)) {
    result[write++] = { a: null, b: j + 1, titleA: null, titleB: titlesB[j], status: 'added' };
  }
  for (const i of restA.slice(pairCount)) {
    result[write++] = { a: i + 1, b: null, titleA: titlesA[i], titleB: null, status: 'removed' };
  }
  return result.slice(0, write);
}

export function comparePackages(pkgA: PptxPackage, pkgB: PptxPackage): DeckDiff {
  const slideMatches = matchSlides(pkgA, pkgB);
  const changes: ParaChange[] = [];
  for (const match of slideMatches) {
    if (match.status === 'added' || match.status === 'removed') continue;
    const aParas = flattenSlideText(pkgA, match.a as number);
    const bParas = flattenSlideText(pkgB, match.b as number);
    const key = (p: FlatPara): string => `${p.shapeKey}#${p.paraIndex}`;
    const mapA = new Map(aParas.map((p) => [key(p), p]));
    const mapB = new Map(bParas.map((p) => [key(p), p]));
    for (const p of bParas) {
      const other = mapA.get(key(p));
      if (!other) {
        changes.push({
          slideA: null, slideB: match.b, shapeKey: p.shapeKey, paraIndex: p.paraIndex,
          kind: 'added', isNotes: p.isNotes, before: '', after: p.text, diff: wordDiff('', p.text),
        });
      } else if (other.text !== p.text) {
        changes.push({
          slideA: match.a, slideB: match.b, shapeKey: p.shapeKey, paraIndex: p.paraIndex,
          kind: 'modified', isNotes: p.isNotes, before: other.text, after: p.text, diff: wordDiff(other.text, p.text),
        });
      }
    }
    for (const p of aParas) {
      if (!mapB.has(key(p))) {
        changes.push({
          slideA: match.a, slideB: null, shapeKey: p.shapeKey, paraIndex: p.paraIndex,
          kind: 'removed', isNotes: p.isNotes, before: p.text, after: '', diff: wordDiff(p.text, ''),
        });
      }
    }
  }
  return {
    slideMatches,
    changes,
    summary: {
      slidesA: pkgA.slides.length,
      slidesB: pkgB.slides.length,
      slidesAdded: slideMatches.filter((m) => m.status === 'added').length,
      slidesRemoved: slideMatches.filter((m) => m.status === 'removed').length,
      paragraphsModified: changes.filter((c) => c.kind === 'modified').length,
      paragraphsAdded: changes.filter((c) => c.kind === 'added').length,
      paragraphsRemoved: changes.filter((c) => c.kind === 'removed').length,
      identical:
        slideMatches.every((m) => m.status === 'matched') &&
        changes.length === 0 &&
        pkgA.slides.length === pkgB.slides.length,
    },
  };
}

/**
 * Build an annotated copy of deck B: every paragraph this diff marks
 * modified/added is re-colored red (reviewer convention, like PowerPoint's
 * Compare import). The original file is untouched — this writes a new buffer.
 */
export async function annotatedCopyBuffer(pkgB: PptxPackage, diff: DeckDiff): Promise<Buffer> {
  const { buffer } = await minimalSave(pkgB);
  const pkgCopy = await PptxPackage.load(buffer);
  const targets = new Set(
    diff.changes
      .filter((c) => (c.kind === 'modified' || c.kind === 'added') && c.slideB !== null)
      .map((c) => `${c.slideB}:${c.shapeKey}#${c.paraIndex}`),
  );
  const recolor = (paraEl: Element): boolean => {
    let changed = false;
    const doc: Document | null = paraEl.ownerDocument;
    if (!doc) return false;
    let child = paraEl.firstChild;
    while (child) {
      if (child.nodeType === 1) {
        const el = child as Element;
        if (el.localName === 'r' && el.namespaceURI === NS.a) {
          let rPr = firstChildByName(el, 'a', 'rPr');
          if (!rPr) {
            rPr = makeElement(doc, 'a', 'rPr');
            el.insertBefore(rPr, el.firstChild);
          }
          for (const fill of [...rPr.childNodes]) {
            if (fill.nodeType === 1 && (fill as Element).localName === 'solidFill') rPr.removeChild(fill);
          }
          const solidFill = makeElement(doc, 'a', 'solidFill');
          const clr = makeElement(doc, 'a', 'srgbClr');
          clr.setAttribute('val', 'FF0000');
          solidFill.appendChild(clr);
          rPr.insertBefore(solidFill, rPr.firstChild);
          changed = true;
        }
      }
      child = child.nextSibling;
    }
    return changed;
  };

  for (const slide of pkgCopy.slides) {
    const slideNumber = slide.index + 1;
    let slideChanged = false;
    let notesChanged = false;
    for (const item of walkSlideParagraphs(pkgCopy, slide.partPath)) {
      if (targets.has(`${slideNumber}:${item.shapeKey}#${item.para.index}`)) {
        slideChanged = recolor(item.para.el) || slideChanged;
      }
    }
    const notes = walkNotesParagraphs(pkgCopy, slide.partPath);
    for (const item of notes.items) {
      if (targets.has(`${slideNumber}:notes#${item.para.index}`)) {
        notesChanged = recolor(item.para.el) || notesChanged;
      }
    }
    if (slideChanged) pkgCopy.zip.markDirty(slide.partPath);
    if (notesChanged && notes.part) pkgCopy.zip.markDirty(notes.part);
  }
  return (await minimalSave(pkgCopy)).buffer;
}
