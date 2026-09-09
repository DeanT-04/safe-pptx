import { createHash } from 'node:crypto';
import type { ParagraphInfo } from './model.js';

/**
 * Stable paragraph anchor ids, ported from safe-docx's `_bk_*` scheme:
 * deterministic across re-opens and machines for the same file content,
 * derived from the paragraph's normalized text and its neighbors. Editing a
 * paragraph's text intentionally changes its id (loud stale-reference
 * failure instead of silent mis-edit). Duplicates get `|salt:N` disambiguation.
 */

export function normalizeForHash(text: string): string {
  return text
    .normalize('NFKC')
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function computeAnchor(parts: (string | number | null | undefined)[], salt: number): string {
  const key = [...parts, salt > 0 ? `|salt:${salt}` : ''].join('\u0000');
  const hex = createHash('sha256').update(key, 'utf8').digest('hex').slice(0, 16);
  return `_bk_${hex}`;
}

/**
 * Assign anchor ids to a slide's flattened paragraph list. `contexts[i]`
 * identifies paragraph i's location (e.g. its shape key) — an anchor depends
 * only on the slide part, its own context, its text, and its immediate
 * neighbors' texts, so unrelated edits elsewhere on the slide don't move it.
 * Duplicates of the same base anchor get salted deterministically.
 */
export function assignAnchors(
  slidePartPath: string,
  contexts: (string | number)[][],
  paras: ParagraphInfo[],
): string[] {
  const texts = paras.map((p) => normalizeForHash(p.text));
  const seen = new Map<string, number>();
  return texts.map((text, i) => {
    const context = [slidePartPath, ...(contexts[i] ?? [])];
    const neighbors = [texts[i - 1] ?? null, texts[i + 1] ?? null];
    const base = computeAnchor([...context, text, ...neighbors], 0);
    const count = seen.get(base) ?? 0;
    seen.set(base, count + 1);
    return count === 0 ? base : computeAnchor([...context, text, ...neighbors], count);
  });
}
