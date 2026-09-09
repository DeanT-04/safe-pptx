import { NS } from '../xml/namespaces.js';
import { childrenByName, deepText, firstDescendantByName } from '../xml/parse.js';
import type { Document, Element } from '../xml/parse.js';
import type { PptxPackage } from '../package/presentation.js';

/**
 * Read access to both PowerPoint comment formats:
 * - legacy (2007/2010): `ppt/comments/commentN.xml` (`p:cmLst`/`p:cm`) + `commentAuthors.xml`
 * - modern threaded (2013+): `ppt/threadedComments/threadedCommentN.xml` (p18) with parentId replies
 */

export interface LegacyComment {
  format: 'legacy';
  slidePart: string;
  author?: string;
  dateTime?: string;
  position?: { x: number; y: number };
  text: string;
}

export interface ThreadedComment {
  format: 'threaded';
  slidePart: string;
  id: string;
  parentId?: string;
  author?: string;
  created?: string;
  modified?: string;
  text: string;
  done?: boolean;
}

export interface DeckComments {
  authors: { id: string; name: string; initials?: string }[];
  legacy: LegacyComment[];
  threaded: ThreadedComment[];
}

function parseAuthors(pkg: PptxPackage): { id: string; name: string; initials?: string }[] {
  const out: { id: string; name: string; initials?: string }[] = [];
  const legacyPart = 'ppt/commentAuthors.xml';
  if (pkg.zip.hasPart(legacyPart)) {
    const doc = pkg.zip.doc(legacyPart);
    const list = firstDescendantByName(doc, 'p', 'cmAuthorLst');
    for (const el of list ? childrenByName(list, 'p', 'cmAuthor') : []) {
      out.push({
        id: el.getAttribute('id') ?? '',
        name: el.getAttribute('name') ?? '',
        initials: el.getAttribute('initials') ?? undefined,
      });
    }
  }
  // Modern threaded-comment authors (p18).
  const modernPart = 'ppt/threadedComments/threadedCommentAuthors.xml';
  if (pkg.zip.hasPart(modernPart)) {
    const doc = pkg.zip.doc(modernPart);
    const list = firstDescendantByName(doc, 'p18', 'threadedCommentAuthors');
    for (const el of list ? childrenByName(list, 'p18', 'threadedCommentAuthor') : []) {
      out.push({
        id: el.getAttribute('id') ?? '',
        name: el.getAttribute('name') ?? '',
        initials: el.getAttribute('initials') ?? undefined,
      });
    }
  }
  return out;
}

function posOf(cm: Element): { x: number; y: number } | undefined {
  if (cm.getAttribute('x') === null && cm.getAttribute('y') === null) return undefined;
  return { x: Number(cm.getAttribute('x') ?? 0), y: Number(cm.getAttribute('y') ?? 0) };
}

export function readComments(pkg: PptxPackage): DeckComments {
  const authors = parseAuthors(pkg);
  const authorName = (id: string | undefined): string | undefined =>
    authors.find((a) => a.id === id)?.name;

  const legacy: LegacyComment[] = [];
  const threaded: ThreadedComment[] = [];

  for (const slide of pkg.slides) {
    for (const partPath of pkg.rels.targetsOfType(slide.partPath, 'comments')) {
      if (!pkg.zip.hasPart(partPath)) continue;
      const doc: Document = pkg.zip.doc(partPath);
      const list = firstDescendantByName(doc, 'p', 'cmLst');
      if (!list) continue;
      for (const cm of childrenByName(list, 'p', 'cm')) {
        const textEl = firstDescendantByName(cm, 'p', 'text');
        legacy.push({
          format: 'legacy',
          slidePart: slide.partPath,
          author: authorName(cm.getAttribute('authorId') ?? undefined),
          dateTime: cm.getAttribute('dt') ?? undefined,
          position: posOf(cm),
          text: textEl ? deepText(textEl) : '',
        });
      }
    }
    for (const partPath of pkg.rels.targetsOfType(slide.partPath, 'threadedComments')) {
      if (!pkg.zip.hasPart(partPath)) continue;
      const doc = pkg.zip.doc(partPath);
      const list = firstDescendantByName(doc, 'p18', 'threadedComments');
      if (!list) continue;
      for (const cm of childrenByName(list, 'p18', 'threadedComment')) {
        const textEl = firstDescendantByName(cm, 'p18', 'text');
        threaded.push({
          format: 'threaded',
          slidePart: slide.partPath,
          id: cm.getAttribute('id') ?? '',
          parentId: cm.getAttribute('parentId') ?? undefined,
          author: cm.getAttribute('author') ?? undefined,
          created: cm.getAttribute('created') ?? undefined,
          modified: cm.getAttribute('modified') ?? undefined,
          done: cm.getAttribute('done') === '1' ? true : undefined,
          text: textEl ? deepText(textEl) : '',
        });
      }
    }
  }
  return { authors, legacy, threaded };
}
