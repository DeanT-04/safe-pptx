/**
 * Zip-bomb and hostile-archive guards, applied before any entry is materialized.
 * Mirrors the safe-docx archive guard stance: refuse absurd expansion ratios,
 * entry counts, sizes, and unsafe entry names.
 */

export interface ArchiveEntryInfo {
  name: string;
  compressedSize: number;
  uncompressedSize: number;
}

export class UnsafeArchiveError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'UnsafeArchiveError';
  }
}

const MAX_ENTRIES = 20_000;
const MAX_TOTAL_UNCOMPRESSED = 512 * 1024 * 1024; // 512 MB
const MAX_SINGLE_UNCOMPRESSED = 200 * 1024 * 1024; // 200 MB
const MAX_RATIO = 200;
const RATIO_CHECK_MIN_UNCOMPRESSED = 1024 * 1024; // only ratio-check entries ≥ 1 MB

/** Normalize a zip entry name for safety checks (backslashes, drive letters, traversal). */
export function normalizeEntryName(name: string): string {
  return name.replace(/\\/g, '/');
}

export function assertSafeArchive(entries: ArchiveEntryInfo[]): void {
  if (entries.length > MAX_ENTRIES) {
    throw new UnsafeArchiveError(`archive has ${entries.length} entries (max ${MAX_ENTRIES})`);
  }
  let total = 0;
  for (const entry of entries) {
    const name = normalizeEntryName(entry.name);
    if (name.includes('\0')) {
      throw new UnsafeArchiveError(`entry name contains NUL byte: ${JSON.stringify(entry.name)}`);
    }
    const segments = name.split('/');
    if (segments.includes('..')) {
      throw new UnsafeArchiveError(`entry name escapes the archive root: ${name}`);
    }
    if (/^[a-zA-Z]:/.test(name)) {
      throw new UnsafeArchiveError(`entry name is drive-absolute: ${name}`);
    }
    if (entry.uncompressedSize > MAX_SINGLE_UNCOMPRESSED) {
      throw new UnsafeArchiveError(
        `entry ${name} is ${entry.uncompressedSize} bytes (max ${MAX_SINGLE_UNCOMPRESSED})`,
      );
    }
    total += entry.uncompressedSize;
    if (
      entry.compressedSize > 0 &&
      entry.uncompressedSize >= RATIO_CHECK_MIN_UNCOMPRESSED &&
      entry.uncompressedSize / entry.compressedSize > MAX_RATIO
    ) {
      throw new UnsafeArchiveError(
        `entry ${name} has expansion ratio ${Math.round(entry.uncompressedSize / entry.compressedSize)} (max ${MAX_RATIO})`,
      );
    }
  }
  if (total > MAX_TOTAL_UNCOMPRESSED) {
    throw new UnsafeArchiveError(
      `archive expands to ${total} bytes (max ${MAX_TOTAL_UNCOMPRESSED})`,
    );
  }
}
