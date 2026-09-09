import type { ToolDef, ToolContext } from './tool_catalog.js';
import { coreTools } from './tools/core_tools.js';

export type { ToolContext };

/**
 * Authoritative tool registry. Phases append their tool arrays here in order.
 * Read tools first, then edit tools, then structure/compare/save.
 */
export const toolCatalog: ToolDef[] = [...coreTools];
