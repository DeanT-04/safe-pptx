import type { z } from 'zod';
import type { SessionManager } from './session/manager.js';

export interface ToolContext {
  sessions: SessionManager;
}

export type ToolOutput = Record<string, unknown>;

/**
 * A tool definition in the catalog. Handlers return JSON-serializable
 * objects; the server wraps them as text content. Thrown errors become
 * isError results with the message as text.
 */
export interface ToolDef {
  name: string;
  title: string;
  description: string;
  /** Zod raw shape (record of zod types), passed to registerTool. */
  schema: z.ZodRawShape;
  handler: (ctx: ToolContext, args: Record<string, unknown>) => Promise<ToolOutput> | ToolOutput;
}
