import type { ToolDef, ToolContext } from './tool_catalog.js';
import { coreTools } from './tools/core_tools.js';
import { readTools } from './tools/read_tools.js';
import { editTools } from './tools/edit_tools.js';
import { structureTools } from './tools/structure_tools.js';

export type { ToolContext };

/**
 * Authoritative tool registry. Phases append their tool arrays here in order.
 * Read tools first, then edit tools, then structure/compare/save.
 */
export const toolCatalog: ToolDef[] = [...coreTools, ...readTools, ...editTools, ...structureTools];
