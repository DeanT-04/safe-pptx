#!/usr/bin/env node
/**
 * Stdio end-to-end test: spawns the compiled MCP server and speaks JSON-RPC
 * over stdin/stdout — initialize, tools/list, tools/call (read_file, grep).
 */
import { spawn } from 'node:child_process';
import path from 'node:path';
import fs from 'node:fs';
import { fileURLToPath } from 'node:url';

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const serverJs = path.join(root, 'packages', 'pptx-mcp', 'dist', 'server.js');
const fixturePath = path.join(root, 'fixtures', 'generated', 'sample.pptx');
if (!fs.existsSync(fixturePath)) {
  console.error('fixture missing — run `npm run fixture` first');
  process.exit(1);
}

const child = spawn(process.execPath, [serverJs], { stdio: ['pipe', 'pipe', 'pipe'] });
const checks = [];
function check(name, ok, detail = '') {
  checks.push(ok);
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? ` — ${detail}` : ''}`);
}

let buffer = '';
const pending = new Map();
child.stdout.on('data', (chunk) => {
  buffer += chunk.toString('utf8');
  let idx;
  while ((idx = buffer.indexOf('\n')) !== -1) {
    const line = buffer.slice(0, idx).trim();
    buffer = buffer.slice(idx + 1);
    if (!line) continue;
    try {
      const msg = JSON.parse(line);
      if (msg.id !== undefined && pending.has(msg.id)) {
        pending.get(msg.id)(msg);
        pending.delete(msg.id);
      }
    } catch {
      // non-JSON line — ignore
    }
  }
});
child.stderr.on('data', (chunk) => process.stderr.write(chunk));

function request(id, method, params, timeoutMs = 15000) {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      pending.delete(id);
      reject(new Error(`timeout waiting for response to ${method}`));
    }, timeoutMs);
    pending.set(id, (msg) => {
      clearTimeout(timer);
      resolve(msg);
    });
    child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', id, method, params })}\n`);
  });
}

try {
  const init = await request(1, 'initialize', {
    protocolVersion: '2025-03-26',
    capabilities: {},
    clientInfo: { name: 'safe-pptx-smoke', version: '0.0.0' },
  });
  check('initialize handshake', init.result?.serverInfo?.name === 'safe-pptx', init.result?.serverInfo?.version);
  child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', method: 'notifications/initialized' })}\n`);

  const list = await request(2, 'tools/list', {});
  const names = (list.result?.tools ?? []).map((t) => t.name);
  check('tools/list exposes full catalog', names.length === 23, `${names.length} tools`);
  for (const expected of ['read_file', 'replace_text', 'batch_edit', 'save', 'compare_decks', 'add_slide', 'get_edit_log']) {
    if (!names.includes(expected)) check(`catalog contains ${expected}`, false);
  }
  check('catalog spot-checks present', ['read_file', 'replace_text', 'batch_edit', 'save', 'compare_decks', 'add_slide', 'get_edit_log'].every((t) => names.includes(t)));

  const call = await request(3, 'tools/call', {
    name: 'read_file',
    arguments: { file_path: fixturePath, limit: 4 },
  });
  const text = call.result?.content?.[0]?.text ?? '';
  check('tools/call read_file returns payload', text.includes('"entries"') && text.includes('_bk_'), `${text.length} chars`);

  const grep = await request(4, 'tools/call', {
    name: 'grep',
    arguments: { file_path: fixturePath, pattern: 'Quarterly' },
  });
  check('tools/call grep works', (grep.result?.content?.[0]?.text ?? '').includes('Quarterly'));

  const bad = await request(5, 'tools/call', {
    name: 'replace_text',
    arguments: { file_path: fixturePath, anchor: '_bk_nope', old_string: 'x', new_string: 'y', instruction: 'expect error' },
  });
  check('errors surface as isError', bad.result?.isError === true && (bad.result?.content?.[0]?.text ?? '').includes('not found'));
} catch (err) {
  check('stdio flow completes', false, err instanceof Error ? err.message : String(err));
}

child.kill();
process.exit(checks.every(Boolean) ? 0 : 1);
