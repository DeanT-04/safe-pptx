import { z } from 'zod';
import {
  addSlide,
  addThreadedComment,
  deleteSlide,
  deleteThreadedComment,
  duplicateSlide,
  reorderSlides,
  setNotesText,
  setTableCellText,
} from '@safe-pptx/pptx-core';
import type { ToolDef } from '../tool_catalog.js';
import type { PptxSession, EditRecord } from '../session/manager.js';
import { resolveAllowedPath } from './path_policy.js';

function requireInstruction(args: Record<string, unknown>): string {
  const instruction = args.instruction;
  if (typeof instruction !== 'string' || instruction.trim().length === 0) {
    throw new Error('instruction is required for edits — describe the intent');
  }
  return instruction;
}

function record(session: PptxSession, tool: string, target: string, before: string, after: string): void {
  const entry: EditRecord = {
    timestamp: new Date().toISOString(),
    tool,
    target,
    before: before.length > 400 ? `${before.slice(0, 400)}…` : before,
    after: after.length > 400 ? `${after.slice(0, 400)}…` : after,
  };
  session.audit.push(entry);
}

const editNotesTool: ToolDef = {
  name: 'edit_notes',
  title: 'Edit speaker notes',
  description:
    'Set the full text of a slide\u2019s speaker notes ("\\n" = new paragraph). Creates the notes ' +
    'part with proper rels/content-types when the slide has none (requires the deck to have a notes master, which most do).',
  schema: {
    file_path: z.string().describe('Absolute path to the .pptx file.'),
    slide: z.number().int().describe('Slide number (1-based display order).'),
    text: z.string().describe('New notes text; "\\n" separates paragraphs.'),
    instruction: z.string().describe('Short statement of the edit intent.'),
  },
  handler: async (ctx, args) => {
    const instruction = requireInstruction(args);
    const filePath = resolveAllowedPath(String(args.file_path));
    const session = await ctx.sessions.getOrCreate(filePath);
    const result = setNotesText(session.pkg, Number(args.slide), String(args.text));
    record(session, 'edit_notes', `slide ${args.slide} ${result.partPath}`, result.before, result.after);
    return { ...result, slide: args.slide, instruction };
  },
};

const editTableCellTool: ToolDef = {
  name: 'edit_table_cell',
  title: 'Edit table cell',
  description:
    'Set the text of one table cell, addressed by the table shape\u2019s cNvPr id plus 0-based ' +
    'row/column. Formatting of the cell\u2019s first run is preserved as the template. For surgical ' +
    'in-cell edits use replace_text with the cell paragraph\u2019s anchor instead.',
  schema: {
    file_path: z.string().describe('Absolute path to the .pptx file.'),
    slide: z.number().int().describe('Slide number (1-based).'),
    shape_id: z.number().int().describe('Table shape id (p:cNvPr/@id from read_file).'),
    row: z.number().int().describe('0-based row.'),
    col: z.number().int().describe('0-based column.'),
    text: z.string().describe('New cell text ("\\n" = new paragraph).'),
    instruction: z.string().describe('Short statement of the edit intent.'),
  },
  handler: async (ctx, args) => {
    const instruction = requireInstruction(args);
    const filePath = resolveAllowedPath(String(args.file_path));
    const session = await ctx.sessions.getOrCreate(filePath);
    const result = setTableCellText(
      session.pkg,
      Number(args.slide),
      Number(args.shape_id),
      Number(args.row),
      Number(args.col),
      String(args.text),
    );
    record(session, 'edit_table_cell', `slide ${args.slide} shape ${args.shape_id} r${args.row}c${args.col}`, result.before, result.after);
    return { ...result, slide: args.slide, shape_id: args.shape_id, row: args.row, col: args.col, instruction };
  },
};

const addCommentTool: ToolDef = {
  name: 'add_comment',
  title: 'Add comment',
  description:
    'Add a modern threaded comment to a slide (or a reply to an existing comment via parent_id). ' +
    'Written in the PowerPoint 2013+ threaded-comments format with authors part and rels.',
  schema: {
    file_path: z.string().describe('Absolute path to the .pptx file.'),
    slide: z.number().int().describe('Slide number (1-based).'),
    author: z.string().describe('Comment author display name.'),
    text: z.string().describe('Comment text.'),
    initials: z.string().optional().describe('Author initials.'),
    parent_id: z.string().optional().describe('Comment id to reply to.'),
    instruction: z.string().optional().describe('Why the comment is being added.'),
  },
  handler: async (ctx, args) => {
    const filePath = resolveAllowedPath(String(args.file_path));
    const session = await ctx.sessions.getOrCreate(filePath);
    const result = addThreadedComment(session.pkg, Number(args.slide), {
      author: String(args.author),
      text: String(args.text),
      initials: typeof args.initials === 'string' ? args.initials : undefined,
      parentId: typeof args.parent_id === 'string' ? args.parent_id : undefined,
    });
    record(session, 'add_comment', `slide ${args.slide}`, '', String(args.text));
    return { ...result, slide: args.slide, instruction: args.instruction ?? '' };
  },
};

const deleteCommentTool: ToolDef = {
  name: 'delete_comment',
  title: 'Delete comment',
  description: 'Delete a threaded comment (and its replies) by id from get_comments.',
  schema: {
    file_path: z.string().describe('Absolute path to the .pptx file.'),
    comment_id: z.string().describe('Threaded comment id.'),
    instruction: z.string().optional().describe('Why the comment is being deleted.'),
  },
  handler: async (ctx, args) => {
    const filePath = resolveAllowedPath(String(args.file_path));
    const session = await ctx.sessions.getOrCreate(filePath);
    const result = deleteThreadedComment(session.pkg, String(args.comment_id));
    record(session, 'delete_comment', String(args.comment_id), 'comment present', 'deleted');
    return { ...result, instruction: args.instruction ?? '' };
  },
};

const addSlideTool: ToolDef = {
  name: 'add_slide',
  title: 'Add slide',
  description:
    'Append a new empty slide instantiated from a layout (placeholder prototypes copied, text ' +
    'emptied). Defaults to the last slide\u2019s layout. Package rels and content types maintained.',
  schema: {
    file_path: z.string().describe('Absolute path to the .pptx file.'),
    layout_part: z.string().optional().describe('Layout part path (e.g. ppt/slideLayouts/slideLayout1.xml). Default: last slide\u2019s layout.'),
    instruction: z.string().describe('Short statement of intent.'),
  },
  handler: async (ctx, args) => {
    const filePath = resolveAllowedPath(String(args.file_path));
    const session = await ctx.sessions.getOrCreate(filePath);
    const result = addSlide(session.pkg, typeof args.layout_part === 'string' ? args.layout_part : undefined);
    record(session, 'add_slide', result.partPath, '', `new slide ${result.slideNumber}`);
    return { ...result, instruction: args.instruction ?? '' };
  },
};

const duplicateSlideTool: ToolDef = {
  name: 'duplicate_slide',
  title: 'Duplicate slide',
  description:
    'Duplicate a slide immediately after itself. Layout and media relationships are shared; ' +
    'notes and comments are not copied (slide-scoped parts stay 1:1).',
  schema: {
    file_path: z.string().describe('Absolute path to the .pptx file.'),
    slide: z.number().int().describe('Slide to duplicate (1-based).'),
    instruction: z.string().describe('Short statement of intent.'),
  },
  handler: async (ctx, args) => {
    const filePath = resolveAllowedPath(String(args.file_path));
    const session = await ctx.sessions.getOrCreate(filePath);
    const result = duplicateSlide(session.pkg, Number(args.slide));
    record(session, 'duplicate_slide', `slide ${args.slide}`, '', `duplicate at ${result.partPath}`);
    return { ...result, instruction: args.instruction ?? '' };
  },
};

const reorderSlidesTool: ToolDef = {
  name: 'reorder_slides',
  title: 'Reorder slides',
  description:
    'Reorder slides by rewriting sldIdLst (display order). Files are never renamed. ' +
    'order must be a permutation of the current 1-based slide numbers.',
  schema: {
    file_path: z.string().describe('Absolute path to the .pptx file.'),
    order: z.array(z.number().int()).describe('New order as a permutation of 1..slide_count.'),
    instruction: z.string().describe('Short statement of intent.'),
  },
  handler: async (ctx, args) => {
    const filePath = resolveAllowedPath(String(args.file_path));
    const session = await ctx.sessions.getOrCreate(filePath);
    const order = Array.isArray(args.order) ? (args.order as number[]).map(Number) : [];
    reorderSlides(session.pkg, order);
    record(session, 'reorder_slides', 'presentation.xml', session.pkg.slides.map((s) => s.index + 1).join(','), order.join(','));
    return { order, slide_count: session.pkg.slides.length, instruction: args.instruction ?? '' };
  },
};

const deleteSlideTool: ToolDef = {
  name: 'delete_slide',
  title: 'Delete slide',
  description:
    'Delete a slide with full package cleanup: sldIdLst entry, presentation relationship, slide ' +
    'part, its rels, content-type override, and slide-scoped parts (notes, comments). ' +
    'Refuses to delete the last remaining slide.',
  schema: {
    file_path: z.string().describe('Absolute path to the .pptx file.'),
    slide: z.number().int().describe('Slide to delete (1-based).'),
    instruction: z.string().describe('Short statement of intent.'),
  },
  handler: async (ctx, args) => {
    const filePath = resolveAllowedPath(String(args.file_path));
    const session = await ctx.sessions.getOrCreate(filePath);
    const result = deleteSlide(session.pkg, Number(args.slide));
    record(session, 'delete_slide', `slide ${args.slide}`, `slide ${args.slide}`, 'deleted');
    return { ...result, instruction: args.instruction ?? '' };
  },
};

export const structureTools: ToolDef[] = [
  editNotesTool,
  editTableCellTool,
  addCommentTool,
  deleteCommentTool,
  addSlideTool,
  duplicateSlideTool,
  reorderSlidesTool,
  deleteSlideTool,
];
