import fs from 'node:fs';
import path from 'node:path';
import { z } from 'zod';
import {
  minimalSave,
  commitSave,
  applyFontByAnchor,
  clearFormattingByAnchor,
  insertParagraphByAnchor,
  replaceInParagraphByAnchor,
  resolveParagraph,
  type FontProp,
} from '@safe-pptx/pptx-core';
import type { ToolDef } from '../tool_catalog.js';
import type { PptxSession, EditRecord } from '../session/manager.js';
import { resolveAllowedPath } from './path_policy.js';

function requireInstruction(args: Record<string, unknown>): string {
  const instruction = args.instruction;
  if (typeof instruction !== 'string' || instruction.trim().length === 0) {
    throw new Error(
      'instruction is required for edits — describe the intent (e.g. "update FY figure per Q3 report")',
    );
  }
  return instruction;
}

function record(
  session: PptxSession,
  tool: string,
  target: string,
  before: string,
  after: string,
): void {
  const entry: EditRecord = {
    timestamp: new Date().toISOString(),
    tool,
    target,
    before: before.length > 400 ? `${before.slice(0, 400)}…` : before,
    after: after.length > 400 ? `${after.slice(0, 400)}…` : after,
  };
  session.audit.push(entry);
}

const replaceTextTool: ToolDef = {
  name: 'replace_text',
  title: 'Replace text',
  description:
    'Replace text inside one paragraph, addressed by its _bk_* anchor from read_file/grep. ' +
    'Format-preserving: matching spans run boundaries, boundary runs are split so untouched ' +
    'fragments keep their exact formatting, and replacement text inherits the matched text\u2019s ' +
    'formatting. Tolerant of smart quotes and whitespace runs. Use "\\n" in new_string for a ' +
    'soft line break (a:br).',
  schema: {
    file_path: z.string().describe('Absolute path to the .pptx file.'),
    anchor: z.string().describe('Paragraph anchor (_bk_*) from read_file or grep.'),
    old_string: z.string().describe('Text to replace (must appear in the paragraph).'),
    new_string: z.string().describe('Replacement text (may be empty to delete).'),
    instruction: z.string().describe('Short statement of the edit intent — stored in the audit log.'),
    occurrence: z.number().int().optional().describe('1-based occurrence to replace when the text appears multiple times. Default 1.'),
    all_occurrences: z.boolean().optional().describe('Replace every occurrence in the paragraph.'),
  },
  handler: async (ctx, args) => {
    const instruction = requireInstruction(args);
    const filePath = resolveAllowedPath(String(args.file_path));
    const session = await ctx.sessions.getOrCreate(filePath);
    const result = replaceInParagraphByAnchor(session.pkg, String(args.anchor), String(args.old_string), String(args.new_string), {
      occurrence: typeof args.occurrence === 'number' ? args.occurrence : 1,
      all: args.all_occurrences === true,
    });
    record(session, 'replace_text', `slide ${result.slideNumber} ${result.partPath}`, result.before, result.after);
    return { ...result, instruction };
  },
};

const insertParagraphTool: ToolDef = {
  name: 'insert_paragraph',
  title: 'Insert paragraph',
  description:
    'Insert one or more paragraphs before/after an anchor paragraph. Paragraph properties and ' +
    'run formatting are cloned from a style source (the anchor by default). "\\n\\n" splits ' +
    'paragraphs, "\\n" inserts a soft break.',
  schema: {
    file_path: z.string().describe('Absolute path to the .pptx file.'),
    anchor: z.string().describe('Anchor paragraph (_bk_*).'),
    position: z.enum(['before', 'after']).describe('Insertion side of the anchor.'),
    text: z.string().describe('Text for the new paragraph(s); "\\n\\n" splits paragraphs.'),
    instruction: z.string().describe('Short statement of the edit intent.'),
    style_source: z.string().optional().describe('Anchor of the paragraph to clone formatting from (default: the anchor itself).'),
  },
  handler: async (ctx, args) => {
    const instruction = requireInstruction(args);
    const filePath = resolveAllowedPath(String(args.file_path));
    const session = await ctx.sessions.getOrCreate(filePath);
    const result = insertParagraphByAnchor(
      session.pkg,
      String(args.anchor),
      args.position === 'before' ? 'before' : 'after',
      String(args.text),
      typeof args.style_source === 'string' ? args.style_source : undefined,
    );
    record(session, 'insert_paragraph', `slide ${result.slideNumber} ${result.partPath}`, result.before, result.after);
    return { ...result, instruction };
  },
};

const setFontTool: ToolDef = {
  name: 'set_font',
  title: 'Set font properties',
  description:
    'Set run-level formatting (bold/italic/underline, size in points, color as RRGGBB) on a whole ' +
    'paragraph, or only on the runs overlapping span_text (span boundaries are split so ' +
    'neighboring text keeps its formatting).',
  schema: {
    file_path: z.string().describe('Absolute path to the .pptx file.'),
    anchor: z.string().describe('Paragraph anchor (_bk_*).'),
    instruction: z.string().describe('Short statement of the edit intent.'),
    b: z.enum(['on', 'off']).optional().describe('Bold.'),
    i: z.enum(['on', 'off']).optional().describe('Italic.'),
    u: z.enum(['on', 'off']).optional().describe('Underline.'),
    sz_pt: z.number().optional().describe('Font size in points.'),
    color_hex: z.string().optional().describe('Text color as hex RRGGBB.'),
    span_text: z.string().optional().describe('Restrict formatting to the runs overlapping this text.'),
  },
  handler: async (ctx, args) => {
    const instruction = requireInstruction(args);
    const filePath = resolveAllowedPath(String(args.file_path));
    const session = await ctx.sessions.getOrCreate(filePath);
    const spec = {
      b: args.b as 'on' | 'off' | undefined,
      i: args.i as 'on' | 'off' | undefined,
      u: args.u as 'on' | 'off' | undefined,
      sz_pt: typeof args.sz_pt === 'number' ? args.sz_pt : undefined,
      color_hex: typeof args.color_hex === 'string' ? args.color_hex : undefined,
    };
    const result = applyFontByAnchor(
      session.pkg,
      String(args.anchor),
      spec,
      typeof args.span_text === 'string' ? args.span_text : undefined,
    );
    record(session, 'set_font', `slide ${result.slideNumber} ${result.partPath}`, result.before, result.after);
    return { ...result, runs_touched: result.runsTouched, instruction };
  },
};

const clearFormattingTool: ToolDef = {
  name: 'clear_formatting',
  title: 'Clear run formatting',
  description:
    'Clear run-level formatting (b/i/u/sz/color or all) on a whole paragraph or just the runs ' +
    'overlapping span_text. Inherited styling (layout/master/theme) still applies afterwards.',
  schema: {
    file_path: z.string().describe('Absolute path to the .pptx file.'),
    anchor: z.string().describe('Paragraph anchor (_bk_*).'),
    instruction: z.string().describe('Short statement of the edit intent.'),
    properties: z
      .array(z.enum(['b', 'i', 'u', 'sz', 'color', 'all']))
      .describe("Properties to clear, e.g. ['b','i'] or ['all']."),
    span_text: z.string().optional().describe('Restrict clearing to the runs overlapping this text.'),
  },
  handler: async (ctx, args) => {
    const instruction = requireInstruction(args);
    const filePath = resolveAllowedPath(String(args.file_path));
    const session = await ctx.sessions.getOrCreate(filePath);
    const raw = Array.isArray(args.properties) ? (args.properties as string[]) : [];
    const props: FontProp[] = raw.includes('all')
      ? ['b', 'i', 'u', 'sz', 'color']
      : (raw.filter((p): p is FontProp => ['b', 'i', 'u', 'sz', 'color'].includes(p)));
    if (props.length === 0) throw new Error('properties must list at least one of b/i/u/sz/color/all');
    const result = clearFormattingByAnchor(
      session.pkg,
      String(args.anchor),
      props,
      typeof args.span_text === 'string' ? args.span_text : undefined,
    );
    record(session, 'clear_formatting', `slide ${result.slideNumber} ${result.partPath}`, result.before, result.after);
    return { ...result, runs_touched: result.runsTouched, instruction };
  },
};

interface BatchStep {
  step_id?: string;
  tool?: string;
  [key: string]: unknown;
}

const batchEditTool: ToolDef = {
  name: 'batch_edit',
  title: 'Batch edit',
  description:
    'Apply multiple replace_text / insert_paragraph steps in one call. Validates every step first ' +
    '(step ids unique, known tool, anchors resolvable, no duplicate-step conflicts), then applies ' +
    'sequentially. All-or-nothing at validation; a mid-application failure reports exactly which ' +
    'step failed and which steps already applied.',
  schema: {
    file_path: z.string().describe('Absolute path to the .pptx file.'),
    steps: z
      .array(
        z.object({
          step_id: z.string(),
          tool: z.enum(['replace_text', 'insert_paragraph']),
          anchor: z.string(),
          instruction: z.string().optional(),
          old_string: z.string().optional(),
          new_string: z.string().optional(),
          position: z.enum(['before', 'after']).optional(),
          text: z.string().optional(),
          occurrence: z.number().int().optional(),
          all_occurrences: z.boolean().optional(),
          style_source: z.string().optional(),
        }),
      )
      .describe('Edit steps, applied in order.'),
  },
  handler: async (ctx, args) => {
    const steps = (Array.isArray(args.steps) ? args.steps : []) as BatchStep[];
    if (steps.length === 0) throw new Error('steps must be a non-empty array');
    const filePath = resolveAllowedPath(String(args.file_path));
    const session = await ctx.sessions.getOrCreate(filePath);
    const pkg = session.pkg;

    // ---- validation pass (no mutations) ----
    const seen = new Set<string>();
    for (const step of steps) {
      const id = String(step.step_id ?? '');
      if (!id) throw new Error('every step needs step_id');
      if (seen.has(id)) throw new Error(`duplicate step_id: ${id}`);
      seen.add(id);
      if (step.tool !== 'replace_text' && step.tool !== 'insert_paragraph') {
        throw new Error(`step ${id}: unsupported tool ${String(step.tool)}`);
      }
      if (typeof step.anchor !== 'string') throw new Error(`step ${id}: anchor is required`);
    }
    const anchors = new Set(steps.map((s) => String(s.anchor)));
    for (const anchor of anchors) {
      // Validation pass: refuse before mutating anything when an anchor is unresolvable.
      try {
        resolveParagraph(pkg, anchor);
      } catch (err) {
        return {
          validation: 'failed',
          applied_steps: [],
          error: `step ${String(steps.find((s) => String(s.anchor) === anchor)?.step_id)}: ${err instanceof Error ? err.message : String(err)}`,
          note: 'No steps were applied. Fix the anchors (re-run read_file) and retry.',
        };
      }
    }

    // ---- application pass ----
    const results: Record<string, unknown>[] = [];
    const applied: string[] = [];
    for (const step of steps) {
      const id = String(step.step_id);
      try {
        if (step.tool === 'replace_text') {
          const result = replaceInParagraphByAnchor(
            pkg,
            String(step.anchor),
            String(step.old_string ?? ''),
            String(step.new_string ?? ''),
            { occurrence: typeof step.occurrence === 'number' ? step.occurrence : 1, all: step.all_occurrences === true },
          );
          record(session, 'batch_edit:replace_text', `step ${id} slide ${result.slideNumber}`, result.before, result.after);
          results.push({ step_id: id, ok: true, ...result });
        } else {
          const result = insertParagraphByAnchor(
            pkg,
            String(step.anchor),
            step.position === 'before' ? 'before' : 'after',
            String(step.text ?? ''),
            typeof step.style_source === 'string' ? step.style_source : undefined,
          );
          record(session, 'batch_edit:insert_paragraph', `step ${id} slide ${result.slideNumber}`, result.before, result.after);
          results.push({ step_id: id, ok: true, ...result });
        }
        applied.push(id);
      } catch (err) {
        return {
          validation: 'passed',
          applied_steps: applied,
          failed_step: id,
          error: err instanceof Error ? err.message : String(err),
          results,
          note: 'Earlier steps are applied in the session. Save to keep them, or close_file without saving to discard everything.',
        };
      }
    }
    return { validation: 'passed', applied_steps: applied, failed_step: null, results };
  },
};

const getEditLogTool: ToolDef = {
  name: 'get_edit_log',
  title: 'Get edit log',
  description:
    'The session audit log: every edit with tool, target, timestamp, and before/after text. ' +
    'This replaces tracked changes (pptx has none natively).',
  schema: {
    file_path: z.string().describe('Absolute path to the .pptx file.'),
    offset: z.number().int().optional().describe('1-based entry offset. Default 1.'),
    limit: z.number().int().optional().describe('Max entries (default 100).'),
  },
  handler: async (ctx, args) => {
    const filePath = resolveAllowedPath(String(args.file_path));
    const session = ctx.sessions.require(filePath);
    const offset = typeof args.offset === 'number' ? Math.max(1, args.offset) : 1;
    const limit = typeof args.limit === 'number' ? args.limit : 100;
    const slice = session.audit.slice(offset - 1, offset - 1 + limit);
    return {
      file_path: filePath,
      total: session.audit.length,
      offset,
      has_more: offset - 1 + slice.length < session.audit.length,
      entries: slice,
    };
  },
};

const saveTool: ToolDef = {
  name: 'save',
  title: 'Save file',
  description:
    'Write the session to disk with minimal-restore: untouched parts are byte-identical to the ' +
    'source archive; only edited parts are re-serialized. Returns a save report with the edit manifest.',
  schema: {
    file_path: z.string().describe('Absolute path of the open .pptx session.'),
    save_to_local_path: z.string().optional().describe('Save to a different path (default: overwrite the opened file).'),
    allow_overwrite: z.boolean().optional().describe('Required when save_to_local_path already exists.'),
  },
  handler: async (ctx, args) => {
    const filePath = resolveAllowedPath(String(args.file_path));
    const session = ctx.sessions.require(filePath);
    const targetPath =
      typeof args.save_to_local_path === 'string'
        ? resolveAllowedPath(args.save_to_local_path)
        : filePath;
    if (targetPath !== filePath && fs.existsSync(targetPath) && args.allow_overwrite !== true) {
      throw new Error(`target exists (pass allow_overwrite to replace): ${targetPath}`);
    }
    const { buffer, report } = await minimalSave(session.pkg);
    fs.mkdirSync(path.dirname(targetPath), { recursive: true });
    fs.writeFileSync(targetPath, buffer);
    commitSave(session.pkg, report);
    return {
      saved_to: targetPath,
      edited_parts: report.editedParts,
      added_parts: report.addedParts,
      removed_parts: report.removedParts,
      bytes: report.bytes,
      audit_entries: session.audit.length,
      note: 'Untouched parts are byte-identical to the original file.',
    };
  },
};

export const editTools: ToolDef[] = [
  replaceTextTool,
  insertParagraphTool,
  setFontTool,
  clearFormattingTool,
  batchEditTool,
  getEditLogTool,
  saveTool,
];
