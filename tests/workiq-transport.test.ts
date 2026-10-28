import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { callWorkIqMcp, closeWorkIqMcp, WORKIQ_MCP_URL } from '../src/services/workiq-mcp.js';

let home: string;
let denied: boolean;
let initializations: number;
let toolCalls: number;
let responseResult: Record<string, unknown>;

describe('real SDK Streamable HTTP transport', () => {
  beforeEach(async () => {
    home = await mkdtemp(join(tmpdir(), 'hey365-transport-test-'));
    vi.stubEnv('HEY365_HOME', home);
    denied = false;
    initializations = 0;
    toolCalls = 0;
    responseResult = { content: [{ type: 'text', text: '{"results":[{"data":{"id":"me"},"statusCode":200}]}' }] };
    vi.stubGlobal('fetch', vi.fn(async (input: string | URL, init?: RequestInit) => {
      expect(String(input)).toBe(WORKIQ_MCP_URL);
      if (init?.method === 'GET') return new Response(null, { status: 405 });
      const request = JSON.parse(String(init?.body));
      if (request.method === 'initialize') {
        initializations++;
        return Response.json({
          jsonrpc: '2.0', id: request.id,
          result: { protocolVersion: request.params.protocolVersion, capabilities: { tools: {} }, serverInfo: { name: 'workiq-test', version: '1' } },
        });
      }
      if (request.method === 'notifications/initialized') return new Response(null, { status: 202 });
      if (request.method === 'tools/call') {
        toolCalls++;
        if (denied) {
          return new Response('Missing scope permissions on the request', {
            status: 403,
            headers: { 'WWW-Authenticate': 'Bearer error="insufficient_scope", scope="Mail.Send"' },
          });
        }
        return Response.json({ jsonrpc: '2.0', id: request.id, result: responseResult });
      }
      throw new Error(`Unexpected RPC method: ${request.method}`);
    }));
  });
  afterEach(async () => {
    await closeWorkIqMcp();
    vi.unstubAllGlobals();
    vi.unstubAllEnvs();
    await rm(home, { recursive: true, force: true });
  });

  it('initializes once and decodes real JSON-RPC tool responses', async () => {
    const result = await callWorkIqMcp('fetch', { entityUrls: ['/me?$select=id'] });
    expect(result).toEqual({ results: [{ data: { id: 'me' }, statusCode: 200 }] });
    responseResult = { content: [], structuredContent: { answer: 'Grounded answer' } };
    expect(await callWorkIqMcp('ask', { question: 'A question' })).toEqual({ answer: 'Grounded answer' });
    expect(initializations).toBe(1);
    expect(toolCalls).toBe(2);
  });

  it('does not upscope or replay a 403 even when the SDK recognises an insufficient_scope challenge', async () => {
    denied = true;
    await expect(callWorkIqMcp('do_action', { actionUrl: '/me/sendMail', jsonBody: {} }))
      .rejects.toMatchObject({ code: 'WORKIQ_PERMISSION_DENIED' });
    expect(initializations).toBe(1);
    expect(toolCalls).toBe(1);
  });

  it('preserves MCP tool failures rather than returning them as successful content', async () => {
    responseResult = { isError: true, content: [{ type: 'text', text: 'HTTP 403 Access denied for path: /me/events' }] };
    await expect(callWorkIqMcp('fetch', { entityUrls: ['/me/events'] }))
      .rejects.toMatchObject({ code: 'WORKIQ_PERMISSION_DENIED' });
    expect(toolCalls).toBe(1);
  });
});
