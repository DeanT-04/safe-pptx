import { randomUUID } from 'node:crypto';
import { NS } from '../xml/namespaces.js';
import { childrenByName, makeElement } from '../xml/parse.js';
import type { Document } from '../xml/parse.js';
import { CT } from '../package/content_types.js';
import type { PptxPackage } from '../package/presentation.js';
import { relsPathFor, addRelToDoc } from '../package/rels.js';
import { EditError } from './text_engine.js';

/**
 * Modern threaded comments (PowerPoint 2013+, p18 namespace): writes to
 * `ppt/threadedComments/threadedCommentN.xml` + a shared authors part, with
 * rels from the slide and content-type overrides. Replies use parentId.
 */

const TC_REL_TYPE = 'http://schemas.microsoft.com/office/2011/relationships/threadedComments';
const TC_AUTHORS_PART = 'ppt/threadedComments/threadedCommentAuthors.xml';

function isoNow(): string {
  return new Date().toISOString().replace(/\.\d{3}Z$/, 'Z');
}

function ensureAuthorsPart(pkg: PptxPackage): Document {
  if (!pkg.zip.hasPart(TC_AUTHORS_PART)) {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p18:threadedCommentAuthors xmlns:p18="${NS.p18}"/>`;
    pkg.zip.addPart(TC_AUTHORS_PART, Buffer.from(xml, 'utf8'));
    pkg.contentTypes.setOverride(TC_AUTHORS_PART, CT.threadedCommentAuthors);
    pkg.zip.markDirty('[Content_Types].xml');
  }
  return pkg.zip.doc(TC_AUTHORS_PART);
}

function getOrCreateAuthor(pkg: PptxPackage, author: string, initials: string): string {
  const doc = ensureAuthorsPart(pkg);
  const root = doc.documentElement;
  const existing = childrenByName(root, 'p18', 'threadedCommentAuthor').find(
    (el) => (el.getAttribute('name') ?? '') === author,
  );
  if (existing) return existing.getAttribute('id') ?? '';
  const id = randomUUID();
  const el = makeElement(doc, 'p18', 'threadedCommentAuthor');
  el.setAttribute('id', id);
  el.setAttribute('name', author);
  if (initials) el.setAttribute('initials', initials);
  el.setAttribute('latestResponseTime', isoNow());
  root?.appendChild(el);
  pkg.zip.markDirty(TC_AUTHORS_PART);
  return id;
}

function ensureThreadedCommentsPart(pkg: PptxPackage, slidePartPath: string): { partPath: string; doc: Document; created: boolean } {
  const relsPath = relsPathFor(slidePartPath);
  if (pkg.zip.hasPart(relsPath)) {
    const targets = pkg.rels.targetsOfType(slidePartPath, 'threadedComments');
    if (targets[0] && pkg.zip.hasPart(targets[0])) {
      return { partPath: targets[0], doc: pkg.zip.doc(targets[0]), created: false };
    }
  }
  const n = (() => {
    let max = 0;
    for (const path of pkg.zip.listParts()) {
      const m = path.match(/^ppt\/threadedComments\/threadedComment(\d+)\.xml$/);
      if (m) max = Math.max(max, Number(m[1]));
    }
    return max + 1;
  })();
  const partPath = `ppt/threadedComments/threadedComment${n}.xml`;
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p18:threadedComments xmlns:p18="${NS.p18}"/>`;
  pkg.zip.addPart(partPath, Buffer.from(xml, 'utf8'));
  pkg.contentTypes.setOverride(partPath, CT.threadedComments);
  pkg.zip.markDirty('[Content_Types].xml');

  if (!pkg.zip.hasPart(relsPath)) {
    pkg.zip.addPart(relsPath, Buffer.from(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
      'utf8',
    ));
  }
  const relsDoc = pkg.zip.doc(relsPath);
  addRelToDoc(relsDoc, { type: TC_REL_TYPE, target: `../threadedComments/threadedComment${n}.xml` });
  pkg.zip.markDirty(relsPath);
  pkg.refreshStructure();
  return { partPath, doc: pkg.zip.doc(partPath), created: true };
}

export interface AddCommentResult {
  commentId: string;
  partPath: string;
  createdPart: boolean;
}

export function addThreadedComment(
  pkg: PptxPackage,
  slideNumber: number,
  opts: { author: string; text: string; initials?: string; parentId?: string },
): AddCommentResult {
  const slide = pkg.slideByIndex(slideNumber - 1);
  if (!slide) throw new EditError(`slide ${slideNumber} does not exist`);
  if (opts.parentId) {
    // Parent must exist (in any slide's threaded comments).
    let found = false;
    for (const s of pkg.slides) {
      for (const part of pkg.rels.targetsOfType(s.partPath, 'threadedComments')) {
        if (!pkg.zip.hasPart(part)) continue;
        const doc = pkg.zip.doc(part);
        if (childrenByName(doc.documentElement, 'p18', 'threadedComment').some((c) => c.getAttribute('id') === opts.parentId)) {
          found = true;
          break;
        }
      }
    }
    if (!found) throw new EditError(`parent comment ${opts.parentId} not found`);
  }
  const authorId = getOrCreateAuthor(pkg, opts.author, opts.initials ?? '');
  const { partPath, doc, created } = ensureThreadedCommentsPart(pkg, slide.partPath);
  const id = randomUUID();
  const el = makeElement(doc, 'p18', 'threadedComment');
  el.setAttribute('id', id);
  el.setAttribute('created', isoNow());
  el.setAttribute('author', opts.author);
  if (authorId) el.setAttribute('authorId', authorId);
  if (opts.initials) el.setAttribute('initials', opts.initials);
  if (opts.parentId) el.setAttribute('parentId', opts.parentId);
  el.appendChild(makeElement(doc, 'p18', 'text', opts.text));
  doc.documentElement?.appendChild(el);
  pkg.zip.markDirty(partPath);
  return { commentId: id, partPath, createdPart: created };
}

export function deleteThreadedComment(pkg: PptxPackage, commentId: string): { deleted: string[] } {
  const deleted: string[] = [];
  for (const slide of pkg.slides) {
    for (const part of pkg.rels.targetsOfType(slide.partPath, 'threadedComments')) {
      if (!pkg.zip.hasPart(part)) continue;
      const doc = pkg.zip.doc(part);
      const root = doc.documentElement;
      const comments = childrenByName(root, 'p18', 'threadedComment');
      const target = comments.find((c) => c.getAttribute('id') === commentId);
      if (target) {
        root?.removeChild(target);
        deleted.push(commentId);
      }
      // Cascade: replies of this comment.
      for (const c of [...comments]) {
        if (c.getAttribute('parentId') === commentId) {
          root?.removeChild(c);
          deleted.push(c.getAttribute('id') ?? '');
        }
      }
      if (deleted.length > 0) pkg.zip.markDirty(part);
    }
  }
  if (deleted.length === 0) throw new EditError(`comment ${commentId} not found`);
  return { deleted };
}
