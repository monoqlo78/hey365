#!/usr/bin/env node
/**
 * MCP client compatibility gate.
 *
 * Hey365 has to run unchanged under GitHub Copilot, Claude Code, Claude
 * Desktop, Codex, Cursor, Windsurf and Scout. Those hosts differ in which
 * protocol revision they negotiate and how strictly they validate a tool
 * declaration, so this drives the built server over real stdio JSON-RPC and
 * asserts the properties every one of them relies on:
 *
 *   1. both published protocol revisions negotiate,
 *   2. stdout carries JSON-RPC and nothing else,
 *   3. every tool name and input schema is acceptable to a strict host,
 *   4. tools/call actually returns content, not just tools/list.
 */
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const ENTRY = join(root, 'dist', 'index.js');

// Revisions in the wild: Claude Desktop and older Codex builds still open with
// 2024-11-05, current hosts use 2025-06-18.
const PROTOCOL_VERSIONS = ['2024-11-05', '2025-06-18'];

const EXPECTED_TOOLS = [
  'hey365',
  'hey365_triage',
  'hey365_session',
  'hey365_draft',
  'hey365_send',
  'hey365_digest',
  'hey365_snooze',
  'hey365_mutes',
  'hey365_find',
  'hey365_schedule',
  'hey365_health',
  'hey365_setup',
  'hey365_install',
];

const failures = [];
const check = (condition, message) => {
  if (!condition) failures.push(message);
};

/** One request/response session against a freshly spawned server. */
function session(protocolVersion) {
  return new Promise((resolve, reject) => {
    const child = spawn(process.execPath, [ENTRY, 'mcp'], { stdio: ['pipe', 'pipe', 'inherit'] });
    const received = new Map();
    const stdoutViolations = [];
    let buffer = '';

    const finish = (error) => {
      clearTimeout(timer);
      child.kill();
      if (error) reject(error);
      else resolve({ received, stdoutViolations });
    };

    const timer = setTimeout(() => finish(new Error(`${protocolVersion}: timed out`)), 30_000);
    child.on('error', finish);

    const send = (message) => child.stdin.write(`${JSON.stringify(message)}\n`);

    child.stdout.on('data', (chunk) => {
      buffer += chunk.toString('utf8');
      let index;
      while ((index = buffer.indexOf('\n')) >= 0) {
        const line = buffer.slice(0, index).trim();
        buffer = buffer.slice(index + 1);
        if (!line) continue;

        let message;
        try {
          message = JSON.parse(line);
        } catch {
          // A stdio host parses stdout line by line; one stray log line
          // desynchronises the stream and the server looks dead.
          stdoutViolations.push(line.slice(0, 120));
          continue;
        }
        if (message.id !== undefined) received.set(message.id, message);

        if (message.id === 1) {
          send({ jsonrpc: '2.0', method: 'notifications/initialized' });
          send({ jsonrpc: '2.0', id: 2, method: 'tools/list', params: {} });
        }
        if (message.id === 2) {
          // Needs no network and writes nothing: proves the call path itself.
          send({
            jsonrpc: '2.0',
            id: 3,
            method: 'tools/call',
            params: { name: 'hey365_install', arguments: { dryRun: true } },
          });
        }
        if (message.id === 3) finish();
      }
    });

    send({
      jsonrpc: '2.0',
      id: 1,
      method: 'initialize',
      params: {
        protocolVersion,
        capabilities: {},
        clientInfo: { name: 'hey365-compat', version: '0.0.0' },
      },
    });
  });
}

for (const protocolVersion of PROTOCOL_VERSIONS) {
  const { received, stdoutViolations } = await session(protocolVersion);

  check(stdoutViolations.length === 0, `${protocolVersion}: non-JSON stdout: ${stdoutViolations.join(' | ')}`);

  const init = received.get(1)?.result;
  check(Boolean(init), `${protocolVersion}: no initialize result`);
  check(init?.serverInfo?.name === 'hey365', `${protocolVersion}: unexpected serverInfo`);
  // A host that opens with an older revision must not be answered with a newer
  // one; Claude Desktop drops the connection when that happens.
  check(init?.protocolVersion === protocolVersion, `${protocolVersion}: negotiated ${init?.protocolVersion} instead`);

  const tools = received.get(2)?.result?.tools ?? [];
  const names = tools.map((tool) => tool.name);
  for (const expected of EXPECTED_TOOLS) {
    check(names.includes(expected), `${protocolVersion}: missing tool ${expected}`);
  }

  for (const tool of tools) {
    // OpenAI and Anthropic both reject names outside this shape.
    check(/^[a-zA-Z0-9_-]{1,64}$/.test(tool.name), `${protocolVersion}: invalid tool name "${tool.name}"`);
    check(Boolean(tool.description), `${protocolVersion}: ${tool.name} has no description`);
    check(
      tool.inputSchema?.type === 'object' && typeof tool.inputSchema.properties === 'object',
      `${protocolVersion}: ${tool.name} inputSchema is not an object schema`,
    );
    if (tool.outputSchema) {
      check(tool.outputSchema.type === 'object', `${protocolVersion}: ${tool.name} outputSchema is not an object`);
    }
  }

  const call = received.get(3);
  check(!call?.error, `${protocolVersion}: tools/call failed: ${JSON.stringify(call?.error)}`);
  const content = call?.result?.content;
  check(Array.isArray(content) && content.length > 0, `${protocolVersion}: tools/call returned no content`);
  check(content?.[0]?.type === 'text' && typeof content[0].text === 'string', `${protocolVersion}: content is not text`);

  console.log(`${protocolVersion}: ok (${names.length} tools, tools/call returned ${content?.length ?? 0} block(s))`);
}

if (failures.length > 0) {
  console.error('\nMCP compatibility failures:');
  for (const failure of failures) console.error(`  - ${failure}`);
  process.exit(1);
}

console.log('\nMCP compatibility: all checks passed');
