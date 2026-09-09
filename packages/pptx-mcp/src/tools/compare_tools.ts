import fs from 'node:fs';
import { z } from 'zod';
import {
  PptxPackage,
  annotatedCopyBuffer,
  comparePackages,
  type DeckDiff,
} from '@safe-pptx/pptx-core';
import type { ToolDef } from '../tool_catalog.js';
import { resolveAllowedPath } from './path_policy.js';

async function loadForCompare(filePath: string): Promise<PptxPackage> {
  // Compare is stateless; sessions are not required.
  return PptxPackage.load(fs.readFileSync(filePath));
}

function markdownReport(diff: DeckDiff, pathA: string, pathB: string): string {
  const lines: string[] = [];
  lines.push(`# Deck comparison`);
  lines.push('');
  lines.push(`- Original: \`${pathA}\``);
  lines.push(`- Revised: \`${pathB}\``);
  lines.push('');
  const s = diff.summary;
  lines.push(
    `**Summary:** ${s.slidesA} → ${s.slidesB} slides · ` +
      `${s.slidesAdded} added, ${s.slidesRemoved} removed · ` +
      `${s.paragraphsModified} modified, ${s.paragraphsAdded} added, ${s.paragraphsRemoved} removed paragraphs.`,
  );
  if (s.identical) lines.push('', 'Decks are textually identical.');
  lines.push('');
  for (const match of diff.slideMatches) {
    if (match.status === 'added') {
      lines.push(`## Slide ${match.b} (added) — ${match.titleB ?? '(untitled)'}`);
      lines.push('');
      continue;
    }
    if (match.status === 'removed') {
      lines.push(`## Slide ${match.a} (removed) — ${match.titleA ?? '(untitled)'}`);
      lines.push('');
      continue;
    }
    const slideChanges = diff.changes.filter(
      (c) => c.slideB === match.b || (c.slideA === match.a && c.slideB === null),
    );
    if (slideChanges.length === 0) continue;
    lines.push(`## Slide ${match.a}${match.a !== match.b ? ` → ${match.b}` : ''} — ${match.titleB ?? match.titleA ?? '(untitled)'}`);
    lines.push('');
    for (const change of slideChanges) {
      const loc = change.isNotes ? 'speaker notes' : change.shapeKey;
      if (change.kind === 'modified') {
        lines.push(`- **${loc}** (paragraph ${change.paraIndex + 1}):`);
        lines.push(`  - old: ${change.before}`);
        lines.push(`  - new: ${change.after}`);
      } else if (change.kind === 'added') {
        lines.push(`- **${loc}** (paragraph ${change.paraIndex + 1}) added: ${change.after}`);
      } else {
        lines.push(`- **${loc}** (paragraph ${change.paraIndex + 1}) removed: ${change.before}`);
      }
    }
    lines.push('');
  }
  return lines.join('\n');
}

export const compareTools: ToolDef[] = [
  {
    name: 'compare_decks',
    title: 'Compare decks',
    description:
      'Redline-style comparison of two .pptx files (the tracked-changes equivalent — pptx has ' +
      'none natively). Slides matched by title then position; paragraphs paired per shape; ' +
      'word-level diffs. Optionally writes a Markdown report and/or an annotated copy of the ' +
      'revised deck with changed paragraphs recolored red. Stateless — does not need open sessions.',
    schema: {
      original_path: z.string().describe('Absolute path to the original deck (A).'),
      revised_path: z.string().describe('Absolute path to the revised deck (B).'),
      max_changes: z.number().int().optional().describe('Cap on reported changes (default 200).'),
      report_path: z.string().optional().describe('Write a Markdown report to this absolute path.'),
      annotate_copy_path: z
        .string()
        .optional()
        .describe('Write an annotated copy of the revised deck (changed paragraphs in red) to this absolute path.'),
      allow_overwrite: z.boolean().optional().describe('Allow overwriting existing report/annotated outputs.'),
    },
    handler: async (ctx, args) => {
      void ctx;
      const pathA = resolveAllowedPath(String(args.original_path));
      const pathB = resolveAllowedPath(String(args.revised_path));
      const allowOverwrite = args.allow_overwrite === true;
      const maxChanges = typeof args.max_changes === 'number' ? args.max_changes : 200;

      const pkgA = await loadForCompare(pathA);
      const pkgB = await loadForCompare(pathB);
      const diff = comparePackages(pkgA, pkgB);

      const outputs: string[] = [];
      if (typeof args.report_path === 'string') {
        const reportPath = resolveAllowedPath(args.report_path);
        if (fs.existsSync(reportPath) && !allowOverwrite) {
          throw new Error(`report exists (pass allow_overwrite): ${reportPath}`);
        }
        fs.writeFileSync(reportPath, markdownReport(diff, pathA, pathB), 'utf8');
        outputs.push(reportPath);
      }
      if (typeof args.annotate_copy_path === 'string') {
        const annotatePath = resolveAllowedPath(args.annotate_copy_path);
        if (fs.existsSync(annotatePath) && !allowOverwrite) {
          throw new Error(`annotated copy exists (pass allow_overwrite): ${annotatePath}`);
        }
        const buffer = await annotatedCopyBuffer(pkgB, diff);
        fs.writeFileSync(annotatePath, buffer);
        outputs.push(annotatePath);
      }

      const truncated = diff.changes.length > maxChanges;
      return {
        original_path: pathA,
        revised_path: pathB,
        summary: diff.summary,
        slide_matches: diff.slideMatches,
        changes: diff.changes.slice(0, maxChanges),
        truncated,
        outputs_written: outputs,
      };
    },
  },
];
