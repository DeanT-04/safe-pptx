import os from 'node:os';
import path from 'node:path';
import fs from 'node:fs';

/**
 * Path policy mirroring safe-docx: only absolute paths under the current
 * working directory, the user's home, or the OS temp directory are readable
 * or writable. Symlinks are resolved before the check.
 */
export class PathPolicyError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'PathPolicyError';
  }
}

function allowedRoots(): string[] {
  return [process.cwd(), os.homedir(), os.tmpdir()].map((root) =>
    path.resolve(root).toLowerCase(),
  );
}

function isUnderRoot(candidate: string, root: string): boolean {
  const rel = path.relative(root, candidate);
  return rel === '' || (!rel.startsWith('..') && !path.isAbsolute(rel));
}

export function resolveAllowedPath(input: string): string {
  if (!path.isAbsolute(input)) {
    throw new PathPolicyError(
      `path must be absolute (got "${input}"). Use the full path to the .pptx file.`,
    );
  }
  let resolved = path.resolve(input);
  try {
    resolved = fs.realpathSync(resolved);
  } catch {
    // File may not exist yet (save target); fall back to resolving the parent.
    try {
      resolved = path.resolve(fs.realpathSync(path.dirname(input)), path.basename(input));
    } catch {
      throw new PathPolicyError(`path does not exist and parent cannot be resolved: ${input}`);
    }
  }
  const lowered = resolved.toLowerCase();
  const ok = allowedRoots().some((root) => isUnderRoot(lowered, root));
  if (!ok) {
    throw new PathPolicyError(
      `path is outside the allowed roots (workspace, home, temp): ${resolved}`,
    );
  }
  return resolved;
}
