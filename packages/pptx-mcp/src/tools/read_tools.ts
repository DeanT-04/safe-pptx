import fs from 'node:fs';
import { z } from 'zod';
import {
  PptxPackage,
  buildTaggedText,
  readComments,
  walkDeckParagraphs,
  walkNotesParagraphs,
  walkSlideParagraphs,
  serializeXml,
  type PptxPart,
} from '@safe-pptx/pptx-core';
import type { ToolDef } from '../tool_catalog.js';
import { resolveAllowedPath } from './path_policy.js';
import { DEFAULT_TOKEN_BUDGET, paginate } from '../pagination.js';

interface ReadEntry {
  chars: number;
  slide?: number;
  part?: string;
  shape?: string;
  n?: number;
  anchor?: string;
  level?: number;
  notes?: boolean;
  text: string;
}

function phLabel(shape: { phType: string | null; phIdx: number | null }): string {
  if (shape.phType === null) return '';
  return shape.phIdx === null ? ` [ph=${shape.phType}]` : ` [ph=${shape.phType} idx=${shape.phIdx}]`;
}

function shapeHeader(shapeKey: string, name: string, ph: string): string {
  return `${shapeKey}${name ? ` "${name}"` : ''}${ph}`;
}

function buildReadEntries(pkg: PptxPackage, slideFilter?: number): ReadEntry[] {
  const entries: ReadEntry[] = [];
  const slides = pkg.slides.filter(
    (s) => slideFilter === undefined || s.index + 1 === slideFilter,
  );
  for (const slide of slides) {
    const walked = walkSlideParagraphs(pkg, slide.partPath);
    const notes = walkNotesParagraphs(pkg, slide.partPath);
    entries.push({
      chars: 40 + slide.partPath.length,
      slide: slide.index + 1,
      part: slide.partPath,
      text: `── Slide ${slide.index + 1} ── ${slide.partPath}`,
    });
    if (walked.length === 0 && notes.items.length === 0) {
      entries.push({ chars: 30, slide: slide.index + 1, text: '(no text)' });
    }
    let lastKey: string | null = null;
    let lastShapeId: number | null = null;
    for (const item of walked) {
      if (item.shapeKey !== lastKey) {
        entries.push({
          chars: 40,
          slide: item.slideNumber,
          part: item.slidePart,
          text: `  ${shapeHeader(item.shapeKey, item.shape.name, phLabel(item.shape))}`,
        });
        lastKey = item.shapeKey;
        lastShapeId = item.shape.id;
      }
      void lastShapeId;
      const text = buildTaggedText(item.para);
      entries.push({
        chars: 60 + text.length,
        slide: item.slideNumber,
        part: item.slidePart,
        shape: item.shapeKey,
        n: item.para.index + 1,
        anchor: item.anchor,
        level: item.para.level,
        text: `    n=${item.para.index + 1} ${item.anchor} ${text}`,
      });
    }
    if (notes.items.length > 0) {
      entries.push({
        chars: 40,
        slide: slide.index + 1,
        part: notes.part,
        notes: true,
        text: `  ── speaker notes ── ${notes.part}`,
      });
      for (const { anchor, para } of notes.items) {
        const text = buildTaggedText(para);
        entries.push({
          chars: 60 + text.length,
          slide: slide.index + 1,
          part: notes.part,
          notes: true,
          n: para.index + 1,
          anchor,
          level: para.level,
          text: `    n=${para.index + 1} ${anchor} ${text}`,
        });
      }
    }
  }
  return entries;
}

const readFileTool: ToolDef = {
  name: 'read_file',
  title: 'Read pptx content',
  description:
    'Read a .pptx as a paginated slide → shape → paragraph view with inline formatting tags ' +
    '(<b> <i> <u> <fld> <br>) and stable paragraph anchors (_bk_*) used to address edits. ' +
    'Token-budgeted (~14k tokens default) with offset/next_offset pagination. ' +
    'Slides are 1-based display order (sldIdLst), not filename order.',
  schema: {
    file_path: z.string().describe('Absolute path to the .pptx file.'),
    slide: z.number().int().optional().describe('Restrict to one slide (1-based display order).'),
    offset: z.number().int().optional().describe('1-based entry offset to continue a previous page.'),
    limit: z.number().int().optional().describe('Max entries to return (otherwise token budget applies).'),
    token_budget: z.number().int().optional().describe('Token budget for this read (default ~14000).'),
    node_ids: z
      .array(z.string())
      .optional()
      .describe('Return only paragraphs with these _bk_* anchors (no pagination).'),
  },
  handler: async (ctx, args) => {
    const filePath = resolveAllowedPath(String(args.file_path));
    const session = await ctx.sessions.getOrCreate(filePath);
    const pkg = session.pkg;
    let entries = buildReadEntries(
      pkg,
      typeof args.slide === 'number' ? args.slide : undefined,
    );
    if (Array.isArray(args.node_ids) && args.node_ids.length > 0) {
      const wanted = new Set(args.node_ids.map(String));
      entries = entries.filter((e) => e.anchor !== undefined && wanted.has(e.anchor));
    }
    const page = paginate(
      entries,
      typeof args.offset === 'number' ? args.offset : 1,
      typeof args.token_budget === 'number' ? args.token_budget : DEFAULT_TOKEN_BUDGET,
      typeof args.limit === 'number' ? args.limit : undefined,
    );
    return {
      file_path: filePath,
      slide_count: pkg.slides.length,
      offset: typeof args.offset === 'number' ? args.offset : 1,
      has_more: page.hasMore,
      next_offset: page.nextOffset,
      entries: page.items,
    };
  },
};

const getOutlineTool: ToolDef = {
  name: 'get_outline',
  title: 'Get deck outline',
  description:
    'Structural outline: per slide — title, part path, layout, shape counts, placeholders, and ' +
    'whether a speaker-notes part exists.',
  schema: {
    file_path: z.string().describe('Absolute path to the .pptx file.'),
  },
  handler: async (ctx, args) => {
    const filePath = resolveAllowedPath(String(args.file_path));
    const session = await ctx.sessions.getOrCreate(filePath);
    const pkg = session.pkg;
    const slides = pkg.slides.map((slide) => {
      const walked = walkSlideParagraphs(pkg, slide.partPath);
      const titlePara = walked.find(
        (w) =>
          !w.shapeKey.includes('/') &&
          (w.shape.phType === 'title' || w.shape.phType === 'ctrTitle'),
      );
      const notes = walkNotesParagraphs(pkg, slide.partPath);
      // Many real decks build slides from plain textboxes with no title
      // placeholder — fall back to the first text paragraph, marked inferred.
      const fallback = walked.find((w) => !w.shapeKey.includes('/') && w.para.text.trim().length > 0);
      const inferred = !titlePara && Boolean(fallback);
      return {
        slide: slide.index + 1,
        part: slide.partPath,
        sldId: slide.sldId,
        title: titlePara ? titlePara.para.text : fallback ? fallback.para.text : null,
        title_inferred: inferred,
        layout: pkg.layoutForSlide(slide.partPath),
        text_paragraphs: walked.length,
        has_notes: notes.part !== '',
      };
    });
    return { file_path: filePath, slide_count: pkg.slides.length, slides };
  },
};

function clipContext(text: string, matchStart: number, matchEnd: number, contextChars: number): string {
  const from = Math.max(0, matchStart - contextChars);
  const to = Math.min(text.length, matchEnd + contextChars);
  const prefix = from > 0 ? '…' : '';
  const suffix = to < text.length ? '…' : '';
  return `${prefix}${text.slice(from, matchStart)}<match>${text.slice(matchStart, matchEnd)}</match>${text.slice(matchEnd, to)}${suffix}`;
}

interface GrepHit {
  file: string;
  slide?: number;
  part?: string;
  shape?: string;
  n?: number;
  anchor?: string;
  context: string;
}

function grepPackage(
  pkg: PptxPackage,
  fileLabel: string,
  regex: RegExp,
  maxResults: number,
  contextChars: number,
): GrepHit[] {
  const hits: GrepHit[] = [];
  for (const slide of pkg.slides) {
    for (const item of walkSlideParagraphs(pkg, slide.partPath)) {
      if (hits.length >= maxResults) return hits;
      const match = regex.exec(item.para.text);
      if (!match) continue;
      hits.push({
        file: fileLabel,
        slide: item.slideNumber,
        part: item.slidePart,
        shape: item.shapeKey,
        n: item.para.index + 1,
        anchor: item.anchor,
        context: clipContext(item.para.text, match.index, match.index + match[0].length, contextChars),
      });
    }
    const notes = walkNotesParagraphs(pkg, slide.partPath);
    for (const { anchor, para } of notes.items) {
      if (hits.length >= maxResults) return hits;
      const match = regex.exec(para.text);
      if (!match) continue;
      hits.push({
        file: fileLabel,
        slide: slide.index + 1,
        part: notes.part,
        shape: 'notes',
        n: para.index + 1,
        anchor,
        context: clipContext(para.text, match.index, match.index + match[0].length, contextChars),
      });
    }
  }
  return hits;
}

function grepPackageXml(
  pkg: PptxPackage,
  fileLabel: string,
  regex: RegExp,
  maxResults: number,
  contextChars: number,
): GrepHit[] {
  const hits: GrepHit[] = [];
  for (const partPath of pkg.zip.listParts()) {
    if (hits.length >= maxResults) return hits;
    const part: PptxPart | undefined = pkg.zip.getPart(partPath);
    if (!part?.isXml || part.text === undefined) continue;
    const xml = pkg.zip.hasPart(partPath) ? serializeXml(pkg.zip.doc(partPath)) : part.text;
    let m: RegExpExecArray | null;
    while ((m = regex.exec(xml)) !== null) {
      hits.push({
        file: fileLabel,
        part: partPath,
        context: clipContext(xml, m.index, m.index + m[0].length, contextChars),
      });
      if (hits.length >= maxResults || m.index === regex.lastIndex) break;
    }
  }
  return hits;
}

const grepTool: ToolDef = {
  name: 'grep',
  title: 'Search pptx text',
  description:
    'Regex search over paragraph text (slides + speaker notes) of one open .pptx, or statelessly ' +
    'across several files via file_paths. With search_xml=true, searches raw serialized XML of ' +
    'every part instead. Returns hits with context and paragraph anchors.',
  schema: {
    file_path: z.string().optional().describe('Absolute path (session mode).'),
    file_paths: z.array(z.string()).optional().describe('Absolute paths for stateless multi-file search.'),
    pattern: z.string().describe('JavaScript regular expression.'),
    case_sensitive: z.boolean().optional().describe('Default false.'),
    max_results: z.number().int().optional().describe('Default 50.'),
    context_chars: z.number().int().optional().describe('Context chars around each match. Default 60.'),
    search_xml: z.boolean().optional().describe('Search raw XML instead of paragraph text. Default false.'),
  },
  handler: async (ctx, args) => {
    const flags = args.case_sensitive === true ? '' : 'i';
    let regex: RegExp;
    try {
      regex = new RegExp(String(args.pattern), flags + 'g');
    } catch (err) {
      throw new Error(`invalid pattern: ${err instanceof Error ? err.message : String(err)}`);
    }
    const maxResults = typeof args.max_results === 'number' ? args.max_results : 50;
    const contextChars = typeof args.context_chars === 'number' ? args.context_chars : 60;
    const searchXml = args.search_xml === true;

    const targets: string[] =
      Array.isArray(args.file_paths) && args.file_paths.length > 0
        ? (args.file_paths as string[]).map((p) => resolveAllowedPath(p))
        : [resolveAllowedPath(String(args.file_path))];

    const hits: GrepHit[] = [];
    for (const target of targets) {
      const session = ctx.sessions.get(target);
      const pkg = session?.pkg ?? (await PptxPackage.load(fs.readFileSync(target)));
      const found = searchXml
        ? grepPackageXml(pkg, target, regex, maxResults - hits.length, contextChars)
        : grepPackage(pkg, target, regex, maxResults - hits.length, contextChars);
      hits.push(...found);
      if (hits.length >= maxResults) break;
    }
    return {
      pattern: String(args.pattern),
      files_searched: targets.length,
      hit_count: hits.length,
      truncated: hits.length >= maxResults,
      hits,
    };
  },
};

const getCommentsTool: ToolDef = {
  name: 'get_comments',
  title: 'Get comments',
  description:
    'Read all comments in the deck: legacy 2007/2010 comments (author, date, position, text) and ' +
    'modern threaded comments (with replies via parentId).',
  schema: {
    file_path: z.string().describe('Absolute path to the .pptx file.'),
  },
  handler: async (ctx, args) => {
    const filePath = resolveAllowedPath(String(args.file_path));
    const session = await ctx.sessions.getOrCreate(filePath);
    const comments = readComments(session.pkg);
    return {
      file_path: filePath,
      authors: comments.authors,
      legacy_count: comments.legacy.length,
      threaded_count: comments.threaded.length,
      legacy: comments.legacy,
      threaded: comments.threaded,
    };
  },
};

const exportTool: ToolDef = {
  name: 'export',
  title: 'Export deck text',
  description:
    'Export deck content to Markdown or plaintext (lossy rendering — no formatting round-trip). ' +
    'Default output is next to the source file with .md/.txt extension.',
  schema: {
    file_path: z.string().describe('Absolute path to the .pptx file.'),
    format: z.enum(['markdown', 'plaintext']).optional().describe('Default markdown.'),
    output_path: z.string().optional().describe('Absolute output path override.'),
    allow_overwrite: z.boolean().optional().describe('Allow overwriting an existing output file.'),
  },
  handler: async (ctx, args) => {
    const filePath = resolveAllowedPath(String(args.file_path));
    const session = await ctx.sessions.getOrCreate(filePath);
    const pkg = session.pkg;
    const format = args.format === 'plaintext' ? 'plaintext' : 'markdown';
    const defaultOut = filePath.replace(/\.pptx$/i, '') + (format === 'markdown' ? '.md' : '.txt');
    const outputPath = resolveAllowedPath(
      typeof args.output_path === 'string' ? args.output_path : defaultOut,
    );
    if (fs.existsSync(outputPath) && args.allow_overwrite !== true) {
      throw new Error(`output exists (pass allow_overwrite to replace): ${outputPath}`);
    }

    const lines: string[] = [];
    const md = format === 'markdown';
    for (const slide of pkg.slides) {
      const walked = walkSlideParagraphs(pkg, slide.partPath);
      const title = walked.find(
        (w) => w.shape.phType === 'title' || w.shape.phType === 'ctrTitle',
      );
      lines.push(md ? `## Slide ${slide.index + 1}${title ? ` — ${title.para.text}` : ''}` : `Slide ${slide.index + 1}${title ? `: ${title.para.text}` : ''}`);
      lines.push('');
      let lastKey: string | null = null;
      for (const item of walked) {
        if (title && item.anchor === title.anchor) continue;
        if (item.shapeKey !== lastKey) {
          if (!md) lines.push(`[${item.shapeKey}]`);
          lastKey = item.shapeKey;
        }
        const indent = '  '.repeat(item.para.level);
        const bullet = md ? (item.para.level === 0 ? '- ' : '  - ') : '';
        lines.push(`${indent}${bullet}${buildTaggedText(item.para).replace(/<\/?(b|i|u|fld|br)>/g, '')}`);
      }
      const tableShapes = walked.filter((w) => w.shape.table && w.shapeKey.includes('/'));
      const seenCells = new Set<string>();
      for (const item of tableShapes) {
        if (seenCells.has(item.shapeKey)) continue;
        seenCells.add(item.shapeKey);
      }
      if (seenCells.size > 0) {
        lines.push('');
        const tables = new Map<string, { rows: string[][] }>();
        for (const item of tableShapes) {
          const base = item.shapeKey.split('/')[0];
          const cell = item.shapeKey.split('/')[1] ?? '';
          const m = cell.match(/r(\d+)c(\d+)/);
          if (!m) continue;
          const entry = tables.get(base) ?? { rows: [] };
          while (entry.rows.length <= Number(m[1])) entry.rows.push([]);
          entry.rows[Number(m[1])][Number(m[2])] = item.para.text;
          tables.set(base, entry);
        }
        for (const [base, table] of tables) {
          lines.push(md ? `\`${base}\`` : `[${base}]`);
          for (const row of table.rows) {
            lines.push(`| ${row.join(' | ')} |`);
          }
        }
      }
      const notes = walkNotesParagraphs(pkg, slide.partPath);
      if (notes.items.length > 0) {
        lines.push('');
        lines.push(md ? '### Notes' : '[notes]');
        for (const { para } of notes.items) {
          lines.push(`${md ? '> ' : ''}${para.text}`);
        }
      }
      lines.push('');
    }
    const content = lines.join('\n');
    fs.writeFileSync(outputPath, content, 'utf8');
    return {
      output_path: outputPath,
      format,
      bytes: Buffer.byteLength(content, 'utf8'),
      slides: pkg.slides.length,
    };
  },
};

export const readTools: ToolDef[] = [
  readFileTool,
  getOutlineTool,
  grepTool,
  getCommentsTool,
  exportTool,
];
