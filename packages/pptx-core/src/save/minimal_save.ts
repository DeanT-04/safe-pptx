import JSZip from 'jszip';
import { serializeXml } from '../xml/parse.js';
import type { PptxPackage } from '../package/presentation.js';

/**
 * Minimal-restore save — the core "safe" guarantee, ported from safe-docx:
 * untouched parts are re-emitted byte-identical from the source archive;
 * only parts edited during the session are re-serialized (keeping the
 * original XML declaration). Source entry order is preserved.
 */
export interface SaveReport {
  editedParts: string[];
  addedParts: string[];
  removedParts: string[];
  bytes: number;
}

const DEFAULT_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

function serializeDirtyPart(pkg: PptxPackage, partPath: string): Buffer {
  const doc = pkg.zip.doc(partPath);
  const originalText = pkg.zip.text(partPath) ?? '';
  const declMatch = originalText.match(/^<\?xml[^>]*\?>/);
  const decl = declMatch ? declMatch[0] : DEFAULT_DECL;
  // xmldom's serializer prepends its own XML declaration when given a
  // Document — strip it so the ORIGINAL declaration is preserved exactly.
  let xml = serializeXml(doc);
  if (xml.startsWith('<?xml')) {
    xml = xml.slice(xml.indexOf('?>') + 2).replace(/^[\r\n]+/, '');
  }
  return Buffer.from(`${decl}\r\n${xml}`, 'utf8');
}

export async function minimalSave(pkg: PptxPackage): Promise<{ buffer: Buffer; report: SaveReport }> {
  const zip = new JSZip();
  const order: string[] = [];
  for (const path of pkg.zip.sourceOrder) {
    if (pkg.zip.hasPart(path)) order.push(path);
  }
  for (const path of pkg.zip.addedParts()) order.push(path);

  const editedParts = pkg.zip.dirtyParts();
  for (const path of order) {
    let bytes: Buffer;
    if (pkg.zip.isDirty(path)) {
      bytes = serializeDirtyPart(pkg, path);
    } else {
      const original = pkg.zip.originalBytes(path);
      if (!original) throw new Error(`missing bytes for part: ${path}`);
      bytes = original;
    }
    zip.file(path, bytes);
  }
  const buffer = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });
  const report: SaveReport = {
    editedParts,
    addedParts: pkg.zip.addedParts(),
    removedParts: pkg.zip.removedParts(),
    bytes: buffer.length,
  };
  return { buffer, report };
}

/** Baseline-rollforward after a successful write so the next save stays minimal. */
export function commitSave(pkg: PptxPackage, report: SaveReport): void {
  const map = new Map<string, Buffer>();
  for (const partPath of report.editedParts) {
    map.set(partPath, serializeDirtyPart(pkg, partPath));
  }
  pkg.zip.markSaved(map);
}
