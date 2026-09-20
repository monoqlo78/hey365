#!/usr/bin/env node
// Minimal MCP handshake against the built stdio server: initialize + tools/list.
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const child = spawn(process.execPath, [join(root, 'dist', 'index.js'), 'mcp'], {
  stdio: ['pipe', 'pipe', 'inherit'],
});

const send = (message) => child.stdin.write(`${JSON.stringify(message)}\n`);

let buffer = '';
const timer = setTimeout(() => {
  console.error('timed out waiting for tools/list');
  child.kill();
  process.exit(1);
}, 20_000);

child.stdout.on('data', (chunk) => {
  buffer += chunk.toString('utf8');
  let index;
  while ((index = buffer.indexOf('\n')) >= 0) {
    const line = buffer.slice(0, index).trim();
    buffer = buffer.slice(index + 1);
    if (!line) continue;
    const message = JSON.parse(line);

    if (message.id === 1) {
      console.log(`server: ${message.result.serverInfo.name} ${message.result.serverInfo.version}`);
      send({ jsonrpc: '2.0', method: 'notifications/initialized' });
      send({ jsonrpc: '2.0', id: 2, method: 'tools/list', params: {} });
    }

    if (message.id === 2) {
      const names = message.result.tools.map((tool) => tool.name).sort();
      console.log(`tools (${names.length}): ${names.join(', ')}`);
      clearTimeout(timer);
      child.kill();
      process.exit(names.length >= 8 ? 0 : 1);
    }
  }
});

send({
  jsonrpc: '2.0',
  id: 1,
  method: 'initialize',
  params: {
    protocolVersion: '2024-11-05',
    capabilities: {},
    clientInfo: { name: 'hey365-smoke', version: '0.0.0' },
  },
});
