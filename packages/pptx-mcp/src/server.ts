#!/usr/bin/env node
import { pathToFileURL } from 'node:url';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import { sessionManager, type SessionManager } from './session/manager.js';
import { toolCatalog, type ToolContext } from './tool_registry.js';

export function createServer(sessions: SessionManager = sessionManager): McpServer {
  const server = new McpServer(
    { name: 'safe-pptx', version: '0.1.0' },
    {
      instructions:
        'safe-pptx edits .pptx files surgically with format preservation. Read before editing: ' +
        'use read_file to discover slide/shape/paragraph anchors, address edits by those anchors, ' +
        'and save when done. Untouched content is preserved byte-identical.',
    },
  );

  const ctx: ToolContext = { sessions };

  for (const tool of toolCatalog) {
    server.registerTool(
      tool.name,
      { title: tool.title, description: tool.description, inputSchema: tool.schema },
      async (args: Record<string, unknown>) => {
        try {
          const result = await tool.handler(ctx, args ?? {});
          return toResult(result, false);
        } catch (err) {
          const message = err instanceof Error ? err.message : String(err);
          return toResult({ error: message }, true);
        }
      },
    );
  }

  return server;
}

function toResult(payload: unknown, isError: boolean): CallToolResult {
  return {
    isError,
    content: [{ type: 'text', text: JSON.stringify(payload, null, 2) }],
  };
}

export async function main(): Promise<void> {
  const server = createServer();
  const transport = new StdioServerTransport();
  await server.connect(transport);
  // Keep the process alive on stdio; log to stderr only (stdout is the protocol channel).
  process.on('SIGINT', async () => {
    sessionManager.stop();
    await server.close();
    process.exit(0);
  });
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  main().catch((err) => {
    process.stderr.write(`safe-pptx fatal: ${err instanceof Error ? err.stack : String(err)}\n`);
    process.exit(1);
  });
}
