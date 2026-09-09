import JSZip from 'jszip';
import { assertSafeArchive, normalizeEntryName, UnsafeArchiveError } from './archive_guard.js';
import { parseXml } from '../xml/parse.js';
import type { Document } from '../xml/parse.js';

const XML_SUFFIXES = ['.xml', '.rels'];

export function isXmlPartName(name: string): boolean {
  const normalized = normalizeEntryName(name);
  return XML_SUFFIXES.some((suffix) => normalized.toLowerCase().endsWith(suffix));
}

export interface PptxPart {
  /** Zip path, e.g. `ppt/slides/slide1.xml` — always forward-slashed, no leading slash. */
  path: string;
  bytes: Buffer;
  isXml: boolean;
  /** Decoded UTF-8 text for XML parts (undefined for binary parts). */
  text?: string;
}

/**
 * In-memory OPC container for a .pptx. Every part keeps its original bytes so
 * that save (minimal-restore) can re-emit untouched parts byte-identical.
 * XML parts get a lazily-parsed, cached DOM; editing tools mutate the DOM and
 * mark the part dirty.
 */
export class PptxZip {
  private readonly parts = new Map<string, PptxPart>();
  private readonly docs = new Map<string, Document>();
  private readonly dirty = new Set<string>();
  /** Zip entry order as loaded — preserved on save for maximal fidelity. */
  readonly sourceOrder: string[] = [];

  private constructor() {}

  static async load(buffer: Buffer): Promise<PptxZip> {
    let zip: JSZip;
    try {
      zip = await JSZip.loadAsync(buffer);
    } catch (err) {
      throw new UnsafeArchiveError(
        `not a readable zip/opc container: ${err instanceof Error ? err.message : String(err)}`,
      );
    }
    const entries: { name: string; compressedSize: number; uncompressedSize: number }[] = [];
    zip.forEach((path, file) => {
      const meta = (file as unknown as { _data?: { uncompressedSize?: number; compressedSize?: number } })._data;
      entries.push({
        name: path,
        compressedSize: meta?.compressedSize ?? 0,
        uncompressedSize: Number(meta?.uncompressedSize ?? 0),
      });
    });
    assertSafeArchive(entries);

    const out = new PptxZip();
    const pending: Promise<void>[] = [];
    for (const path of Object.keys(zip.files)) {
      const file = zip.files[path];
      if (!file || file.dir) continue;
      out.sourceOrder.push(path);
      const normalized = normalizeEntryName(path);
      pending.push(
        file.async('nodebuffer').then((bytes) => {
          if (isXmlPartName(normalized)) {
            out.parts.set(normalized, {
              path: normalized,
              bytes,
              isXml: true,
              text: bytes.toString('utf8'),
            });
          } else {
            out.parts.set(normalized, { path: normalized, bytes, isXml: false });
          }
        }),
      );
    }
    await Promise.all(pending);
    if (!out.parts.has('[Content_Types].xml')) {
      throw new UnsafeArchiveError('missing [Content_Types].xml — not an OPC package');
    }
    return out;
  }

  listParts(): string[] {
    return [...this.parts.keys()].sort();
  }

  getPart(path: string): PptxPart | undefined {
    return this.parts.get(path);
  }

  hasPart(path: string): boolean {
    return this.parts.has(path);
  }

  /** Original bytes as loaded — used by minimal-restore save for untouched parts. */
  originalBytes(path: string): Buffer | undefined {
    return this.parts.get(path)?.bytes;
  }

  text(path: string): string | undefined {
    return this.parts.get(path)?.text;
  }

  /** Lazily parsed DOM for an XML part, cached per session. */
  doc(path: string): Document {
    const cached = this.docs.get(path);
    if (cached) return cached;
    const part = this.parts.get(path);
    if (!part || !part.isXml || part.text === undefined) {
      throw new Error(`part is missing or not XML: ${path}`);
    }
    const doc = parseXml(part.text, path);
    this.docs.set(path, doc);
    return doc;
  }

  /** Call after mutating the DOM of an XML part so save re-serializes it. */
  markDirty(path: string): void {
    if (!this.docs.has(path)) {
      throw new Error(`markDirty called for part with no parsed DOM: ${path}`);
    }
    this.dirty.add(path);
  }

  isDirty(path: string): boolean {
    return this.dirty.has(path);
  }

  dirtyParts(): string[] {
    return [...this.dirty].sort();
  }

  /** Add a brand-new part (bytes). XML parts added this way are not parsed until first `doc()` call. */
  addPart(path: string, bytes: Buffer): void {
    if (this.parts.has(path)) throw new Error(`part already exists: ${path}`);
    if (isXmlPartName(path)) {
      this.parts.set(path, { path, bytes, isXml: true, text: bytes.toString('utf8') });
    } else {
      this.parts.set(path, { path, bytes, isXml: false });
    }
    // New parts always serialize from their bytes; nothing to restore from source.
  }

  removePart(path: string): void {
    this.parts.delete(path);
    this.docs.delete(path);
    this.dirty.delete(path);
  }

  /** Parts added after load (not present in the source archive). */
  addedParts(): string[] {
    const source = new Set(this.sourceOrder);
    return [...this.parts.keys()].filter((p) => !source.has(p)).sort();
  }

  /** Source parts deleted during the session. */
  removedParts(): string[] {
    return this.sourceOrder.filter((p) => !this.parts.has(p));
  }

  /** Call after a successful save: dirty parts become the new baseline bytes. */
  markSaved(newBytes: Map<string, Buffer>): void {
    for (const path of this.dirty) {
      const part = this.parts.get(path);
      const bytes = newBytes.get(path);
      if (part && bytes) {
        part.bytes = bytes;
        if (part.isXml) part.text = bytes.toString('utf8');
      }
    }
    this.dirty.clear();
  }
}
