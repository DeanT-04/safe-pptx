import { z } from 'zod';
import type { ToolDef } from '../tool_catalog.js';
import { resolveAllowedPath } from './path_policy.js';

export const coreTools: ToolDef[] = [
  {
    name: 'get_file_status',
    title: 'Get file status',
    description:
      'Get session metadata for a .pptx: whether a session is open, slide count, dirty parts, ' +
      'audit-log size, and creation time. Sessions auto-create on first use of any tool with a file_path.',
    schema: {
      file_path: z.string().describe('Absolute path to the .pptx file.'),
    },
    handler: async (ctx, args) => {
      const filePath = resolveAllowedPath(String(args.file_path));
      const session = ctx.sessions.get(filePath);
      if (!session) {
        return { open: false, file_path: filePath };
      }
      return {
        open: true,
        file_path: filePath,
        created_at: session.createdAt,
        slide_count: session.pkg.slides.length,
        slides: session.pkg.slides.map((s) => ({ index: s.index, part: s.partPath })),
        dirty_parts: session.pkg.zip.dirtyParts(),
        audit_log_entries: session.audit.length,
      };
    },
  },
  {
    name: 'close_file',
    title: 'Close file',
    description:
      'Close the in-memory session for a .pptx, discarding unsaved edits. ' +
      'Pass clear_all=true + confirm=true to close every open session.',
    schema: {
      file_path: z.string().optional().describe('Absolute path of the session to close.'),
      clear_all: z.boolean().optional().describe('Close all sessions instead of one.'),
      confirm: z
        .boolean()
        .optional()
        .describe('Required true when clear_all is true — an explicit acknowledgment.'),
    },
    handler: async (ctx, args) => {
      if (args.clear_all === true) {
        if (args.confirm !== true) {
          throw new Error('clear_all requires confirm: true');
        }
        const closed = ctx.sessions.closeAll();
        return { closed_all: true, sessions_closed: closed };
      }
      if (typeof args.file_path !== 'string') {
        throw new Error('file_path is required unless clear_all is used');
      }
      const filePath = resolveAllowedPath(args.file_path);
      const closed = ctx.sessions.close(filePath);
      return { closed, file_path: filePath };
    },
  },
];
