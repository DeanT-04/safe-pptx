import { childrenByName } from '../xml/parse.js';
import type { Document } from '../xml/parse.js';

export interface Rel {
  id: string;
  type: string;
  target: string;
  targetMode: 'Internal' | 'External';
}

/** Directory containing `partPath`, e.g. `ppt/slides/slide1.xml` → `ppt/slides`. */
export function dirname(partPath: string): string {
  const idx = partPath.lastIndexOf('/');
  return idx === -1 ? '' : partPath.slice(0, idx);
}

/** POSIX-normalized join + resolution of `.`/`..` segments. */
export function resolvePartPath(baseDir: string, target: string): string {
  const segments: string[] = [];
  const combined = target.startsWith('/') ? target : `${baseDir}/${target}`;
  for (const segment of combined.split('/')) {
    if (segment === '' || segment === '.') continue;
    if (segment === '..') segments.pop();
    else segments.push(segment);
  }
  return segments.join('/');
}

export function relsPathFor(partPath: string): string {
  const dir = dirname(partPath);
  const base = partPath.slice(dir.length === 0 ? 0 : dir.length + 1);
  return dir === '' ? `_rels/${base}.rels` : `${dir}/_rels/${base}.rels`;
}

export function parseRels(doc: Document): Rel[] {
  const root = doc.documentElement;
  if (!root || root.localName !== 'Relationships') return [];
  return childrenByName(root, 'rel', 'Relationship').map((el) => ({
    id: el.getAttribute('Id') ?? '',
    type: el.getAttribute('Type') ?? '',
    target: el.getAttribute('Target') ?? '',
    targetMode: (el.getAttribute('TargetMode') as 'Internal' | 'External') || 'Internal',
  }));
}

/**
 * Resolve a relationship target against its source part.
 * External targets are returned untouched.
 */
export function resolveRel(sourcePartPath: string, rel: Rel): string {
  if (rel.targetMode === 'External') return rel.target;
  return resolvePartPath(dirname(sourcePartPath), rel.target);
}

/** Relationship lookup for one source part, with per-rels-part caching. */
export class RelResolver {
  private cache = new Map<string, Rel[]>();

  constructor(
    private readonly load: (relsPartPath: string) => Document | null,
  ) {}

  relsFor(sourcePartPath: string): Rel[] {
    const relsPath = relsPathFor(sourcePartPath);
    let rels = this.cache.get(relsPath);
    if (!rels) {
      const doc = this.load(relsPath);
      rels = doc ? parseRels(doc) : [];
      this.cache.set(relsPath, rels);
    }
    return rels;
  }

  /** Target part path of the relationship with `relId`, or null. */
  targetOf(sourcePartPath: string, relId: string): string | null {
    const rel = this.relsFor(sourcePartPath).find((r) => r.id === relId);
    if (!rel) return null;
    return resolveRel(sourcePartPath, rel);
  }

  /** All internal targets of relationships whose type URI ends with `typeSuffix` (e.g. 'slide'). */
  targetsOfType(sourcePartPath: string, typeSuffix: string): string[] {
    return this.relsFor(sourcePartPath)
      .filter((r) => r.targetMode === 'Internal' && r.type.endsWith(`/${typeSuffix}`))
      .map((r) => resolveRel(sourcePartPath, r));
  }

  /** Invalidate cached rels for one source part (after structural edits). */
  invalidate(sourcePartPath: string): void {
    this.cache.delete(relsPathFor(sourcePartPath));
  }

  invalidateAll(): void {
    this.cache.clear();
  }
}
