/**
 * Token-budget pagination for read views, mirroring safe-docx's ~14k-token
 * read_file budget (estimated as len/4).
 */
export const DEFAULT_TOKEN_BUDGET = 14_000;
const CHARS_PER_TOKEN = 4;

export interface Paginable {
  /** Estimated character cost of this entry. */
  chars: number;
}

export interface Page<T> {
  items: T[];
  hasMore: boolean;
  /** 1-based offset for the next call, or null when exhausted. */
  nextOffset: number | null;
}

export function paginate<T extends Paginable>(
  entries: T[],
  offset1: number,
  tokenBudget: number,
  maxItems?: number,
): Page<T> {
  const start = Math.max(0, offset1 - 1);
  const budgetChars = tokenBudget * CHARS_PER_TOKEN;
  const items: T[] = [];
  let used = 0;
  let i = start;
  for (; i < entries.length; i++) {
    const entry = entries[i];
    if (items.length > 0 && used + entry.chars > budgetChars) break;
    if (maxItems !== undefined && items.length >= maxItems) break;
    items.push(entry);
    used += entry.chars;
  }
  const hasMore = i < entries.length;
  return {
    items,
    hasMore,
    nextOffset: hasMore ? i + 1 : null,
  };
}
