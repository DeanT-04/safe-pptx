import { NS } from '../xml/namespaces.js';
import { childrenByName, firstChildByName, directText } from '../xml/parse.js';
import type { Element } from '../xml/parse.js';

/**
 * Paragraph/run extraction from DrawingML text bodies (`p:txBody`).
 *
 * The logical text of a paragraph is spread across `a:r` runs (and `a:fld`
 * fields). Formatting lives on each run's `a:rPr`; preserving it is the core
 * "safe" property, so all edits happen at run granularity.
 */

export interface RunPart {
  kind: 'run' | 'field';
  el: Element;
  /** The `a:t` text element (null for empty runs). */
  textEl: Element | null;
  text: string;
}

export interface ParagraphInfo {
  el: Element;
  /** 0-based index within its txBody. */
  index: number;
  runs: RunPart[];
  /** Number of `a:br` soft breaks (contributing '\n' to text). */
  breaks: number;
  /** Concatenated run/field text; `a:br` becomes '\n'. */
  text: string;
  /** Paragraph outline level from `a:pPr/@lvl` (default 0). */
  level: number;
}

export function isTrueAttr(value: string | null | undefined): boolean {
  return value === '1' || value === 'true';
}

export function runTextEl(run: Element): Element | null {
  return firstChildByName(run, 'a', 't');
}

export function textOfElement(el: Element): string {
  const t = runTextEl(el);
  return t ? directText(t) : '';
}

export function buildParagraph(el: Element, index: number): ParagraphInfo {
  const runs: RunPart[] = [];
  let text = '';
  let breaks = 0;
  let child = el.firstChild;
  while (child) {
    if (child.nodeType === 1) {
      const e = child as Element;
      if (e.localName === 'r' && e.namespaceURI === NS.a) {
        const textEl = runTextEl(e);
        const t = textEl ? directText(textEl) : '';
        runs.push({ kind: 'run', el: e, textEl, text: t });
        text += t;
      } else if (e.localName === 'fld' && e.namespaceURI === NS.a) {
        const textEl = runTextEl(e);
        const t = textEl ? directText(textEl) : '';
        runs.push({ kind: 'field', el: e, textEl, text: t });
        text += t;
      } else if (e.localName === 'br' && e.namespaceURI === NS.a) {
        breaks += 1;
        text += '\n';
      }
    }
    child = child.nextSibling;
  }
  const pPr = firstChildByName(el, 'a', 'pPr');
  const level = pPr ? Number(pPr.getAttribute('lvl') ?? '0') || 0 : 0;
  return { el, index, runs, breaks, text, level };
}

export function buildTextBody(txBody: Element): ParagraphInfo[] {
  return childrenByName(txBody, 'a', 'p').map((el, index) => buildParagraph(el, index));
}

/** Raw (untagged) paragraph text in one call. */
export function paragraphText(el: Element): string {
  return buildParagraph(el, 0).text;
}

function runTagSignature(run: Element): string {
  const rPr = firstChildByName(run, 'a', 'rPr');
  if (!rPr) return '';
  let sig = '';
  if (isTrueAttr(rPr.getAttribute('b'))) sig += 'b';
  if (isTrueAttr(rPr.getAttribute('i'))) sig += 'i';
  if ((rPr.getAttribute('u') ?? '') === 'sng') sig += 'u';
  return sig;
}

function wrapTagged(sig: string, text: string): string {
  let out = text;
  for (const tag of [...sig].reverse()) {
    out = `<${tag}>${out}</${tag}>`;
  }
  return out;
}

/**
 * Paragraph text with inline formatting tags (`<b> <i> <u>`) for read views.
 * Adjacent runs with identical formatting are merged; fields render their
 * cached text inside `<fld>` markers.
 */
export function buildTaggedText(p: ParagraphInfo): string {
  let out = '';
  let curSig: string | null = null;
  let curText = '';
  const flush = (): void => {
    if (curSig === null) return;
    out += curSig === 'fld' ? `<fld>${curText}</fld>` : wrapTagged(curSig, curText);
    curText = '';
  };
  for (const part of p.runs) {
    const sig = part.kind === 'field' ? 'fld' : runTagSignature(part.el);
    if (sig !== curSig) {
      flush();
      curSig = sig;
    }
    curText += part.text;
  }
  flush();
  if (p.breaks > 0) return out.replace(/\n/g, '<br>');
  return out;
}
