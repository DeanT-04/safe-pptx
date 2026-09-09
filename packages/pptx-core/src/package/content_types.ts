import { childrenByName } from '../xml/parse.js';
import type { Document } from '../xml/parse.js';

/**
 * `[Content_Types].xml` model. Every part in an OPC package must be covered by
 * either a `<Default>` (by extension) or an `<Override>` (by part name);
 * missing coverage makes PowerPoint show a repair prompt.
 */
export class ContentTypes {
  private readonly defaults = new Map<string, string>();
  private readonly overrides = new Map<string, string>();

  private constructor() {}

  static parse(doc: Document): ContentTypes {
    const root = doc.documentElement;
    const out = new ContentTypes();
    for (const el of childrenByName(root, 'ct', 'Default')) {
      const ext = el.getAttribute('Extension');
      const ct = el.getAttribute('ContentType');
      if (ext && ct) out.defaults.set(ext.toLowerCase(), ct);
    }
    for (const el of childrenByName(root, 'ct', 'Override')) {
      const name = el.getAttribute('PartName');
      const ct = el.getAttribute('ContentType');
      if (name && ct) out.overrides.set(name.replace(/^\//, ''), ct);
    }
    return out;
  }

  extensionOf(partPath: string): string {
    const base = partPath.slice(partPath.lastIndexOf('/') + 1);
    const dot = base.lastIndexOf('.');
    return dot === -1 ? '' : base.slice(dot + 1).toLowerCase();
  }

  overrideFor(partPath: string): string | null {
    return this.overrides.get(partPath) ?? null;
  }

  defaultFor(partPath: string): string | null {
    return this.defaults.get(this.extensionOf(partPath)) ?? null;
  }

  contentTypeFor(partPath: string): string | null {
    return this.overrideFor(partPath) ?? this.defaultFor(partPath);
  }

  isCovered(partPath: string): boolean {
    return this.contentTypeFor(partPath) !== null;
  }

  setOverride(partPath: string, contentType: string): void {
    this.overrides.set(partPath, contentType);
  }

  removeOverride(partPath: string): void {
    this.overrides.delete(partPath);
  }

  allOverrides(): Map<string, string> {
    return new Map(this.overrides);
  }
}

export const CT = {
  presentation: 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
  slide: 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
  slideLayout: 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml',
  slideMaster: 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
  notesSlide: 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml',
  notesMaster: 'application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml',
  theme: 'application/vnd.openxmlformats-officedocument.theme+xml',
  presProps: 'application/vnd.openxmlformats-officedocument.presentationml.presProps+xml',
  viewProps: 'application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml',
  tableStyles: 'application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml',
  comments: 'application/vnd.openxmlformats-officedocument.presentationml.comments+xml',
  commentAuthors: 'application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml',
  threadedComments: 'http://schemas.microsoft.com/office/powerpoint/2018/threadedComments',
  threadedCommentAuthors: 'http://schemas.microsoft.com/office/powerpoint/2018/threadedCommentAuthors',
  coreProperties: 'application/vnd.openxmlformats-package.core-properties+xml',
  extendedProperties: 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
} as const;
