import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { Hey365Error, classifyFailure } from '../src/utils/errors.js';

vi.mock('../src/services/workiq-mcp.js', async (original) => ({
  ...await original<typeof import('../src/services/workiq-mcp.js')>(),
  callWorkIqMcp: vi.fn(),
}));

const { callWorkIqMcp, workIqBackend } = await import('../src/services/workiq-mcp.js');
const { ask, callFunction, createEntity, doAction, fetchMany, fetchOne, tryFetch, workIqPath } = await import('../src/services/workiq.js');
const { calendarUrl, listCalendar } = await import('../src/services/graph.js');
const call = vi.mocked(callWorkIqMcp);
const envelope = (data: unknown, statusCode = 200) => ({ results: [{ data, statusCode }] });

describe('hosted WorkIQ operations', () => {
  beforeEach(() => {
    vi.stubEnv('HEY365_WORKIQ_BACKEND', 'mcp');
    call.mockReset();
  });
  afterEach(() => vi.unstubAllEnvs());

  it('defaults to the hosted MCP and rejects unknown backends', () => {
    vi.stubEnv('HEY365_WORKIQ_BACKEND', '');
    expect(workIqBackend()).toBe('mcp');
    vi.stubEnv('HEY365_WORKIQ_BACKEND', 'typo');
    expect(() => workIqBackend()).toThrow('must be mcp or cli');
  });

  it('keeps one outcome per URL in a batched MCP response', async () => {
    call.mockResolvedValue({
      results: [
        { data: { id: 'me' }, statusCode: 200 },
        { statusCode: 403, error: { error: { code: 'AccessDenied', message: 'Missing scope permissions on the request' } } },
      ],
    });
    const result = await fetchMany(['/me?$select=id', '/me/events?$top=1&$select=id']);
    expect(call).toHaveBeenCalledWith('fetch', { entityUrls: ['/me?$select=id', '/me/events?$top=1&$select=id'] }, undefined);
    expect(result[0]).toEqual({ ok: true, data: { id: 'me' } });
    expect(result[1]?.error?.code).toBe('WORKIQ_PERMISSION_DENIED');
    expect(call).toHaveBeenCalledTimes(1);
  });

  it('preserves direct error bodies and failures without an error body', async () => {
    call.mockResolvedValue({ results: [
      { statusCode: 403, error: { code: 'Authorization_RequestDenied', message: 'Insufficient privileges' } },
      { statusCode: 401 },
    ] });
    const result = await fetchMany(['/me', '/me/events']);
    expect(result.map((entry) => entry.error?.code)).toEqual(['WORKIQ_PERMISSION_DENIED', 'WORKIQ_NOT_AUTHENTICATED']);
  });

  it('does not turn a missing or null response into success', async () => {
    call.mockResolvedValue({ results: [] });
    await expect(fetchOne('/me')).rejects.toMatchObject({ code: 'INTERNAL_ERROR' });
    call.mockResolvedValue(null);
    await expect(fetchOne('/me')).rejects.toMatchObject({ code: 'INTERNAL_ERROR' });
  });

  it('does not hide permission errors in optional reads', async () => {
    call.mockResolvedValue({ results: [{ statusCode: 403 }] });
    await expect(tryFetch('/me/events')).rejects.toMatchObject({ code: 'WORKIQ_PERMISSION_DENIED' });
  });

  it('uses the official action and create contracts without replaying writes', async () => {
    call.mockResolvedValue({ statusCode: 202 });
    await doAction('/me/messages/exact-id=/reply', { Comment: 'Approved reply' });
    expect(call).toHaveBeenLastCalledWith('do_action', {
      actionUrl: '/me/messages/exact-id=/reply', jsonBody: { Comment: 'Approved reply' },
    }, undefined);
    call.mockResolvedValue({ id: 'draft-id' });
    await createEntity('/me/messages', { subject: 'Draft' });
    expect(call).toHaveBeenLastCalledWith('create_entity', {
      parentUrl: '/me/messages', jsonBody: { subject: 'Draft' },
    }, undefined);
    call.mockRejectedValue(new Hey365Error('WORKIQ_CONNECTION_ERROR', 'timed out'));
    await expect(doAction('/me/sendMail', {})).rejects.toMatchObject({ code: 'WORKIQ_CONNECTION_ERROR' });
    expect(call).toHaveBeenCalledTimes(3);
  });

  it('treats a confirmed 204 response as success even without a payload', async () => {
    call.mockResolvedValue(envelope(null, 204));
    await expect(doAction('/me/messages/id/reply', {})).resolves.toBeUndefined();
  });

  it('routes OData functions and Copilot ask through MCP tool names', async () => {
    call.mockResolvedValue({ value: [] });
    await callFunction('/me/drive/root/search(q=%27Spec%27)');
    expect(call).toHaveBeenLastCalledWith('call_function', { functionUrl: '/me/drive/root/search(q=%27Spec%27)' }, undefined);
    call.mockResolvedValue({ answer: 'A grounded answer', conversationId: 'conversation-id' });
    expect(await ask('Summarize this meeting', { conversationId: 'previous-id' })).toEqual({
      text: 'A grounded answer', conversationId: 'conversation-id',
    });
    expect(call).toHaveBeenLastCalledWith('ask', {
      question: 'Summarize this meeting', conversationId: 'previous-id',
    }, 180_000);
  });

  it('uses the bounded calendarView contract, not semantic ask or a CLI command', async () => {
    call.mockResolvedValue(envelope({ value: [{ id: 'meeting' }] }));
    const start = new Date('2026-10-28T00:00:00Z');
    const end = new Date('2026-10-29T00:00:00Z');
    expect(await listCalendar(start, end, 5)).toEqual([{ id: 'meeting' }]);
    expect(call).toHaveBeenCalledWith('fetch', {
      entityUrls: [workIqPath(calendarUrl(start, end, 5))],
    }, undefined);
    const path = call.mock.calls[0]?.[1].entityUrls as string[];
    expect(path[0]).toContain('/me/calendarView?startDateTime=2026-10-28T00%3A00%3A00Z');
    expect(path[0]).toContain('&$top=5&$select=');
  });

  it('surfaces a denied meeting-list request rather than returning an empty list', async () => {
    call.mockResolvedValue({ results: [{ statusCode: 403 }] });
    await expect(listCalendar(new Date(), new Date())).rejects.toMatchObject({ code: 'WORKIQ_PERMISSION_DENIED' });
    expect(call).toHaveBeenCalledTimes(1);
  });
});

describe('WorkIQ path and error contracts', () => {
  it('encodes values once while preserving property paths and opaque IDs', () => {
    expect(workIqPath('/me/messages/opaque+id=?$filter=from/emailAddress/address eq \'a/b@example.com\''))
      .toBe('/me/messages/opaque+id=?$filter=from/emailAddress/address%20eq%20%27a%2Fb%40example.com%27');
    expect(workIqPath('/me/events?$orderby=start/dateTime&$search=%22Project%20A%22'))
      .toBe('/me/events?$orderby=start/dateTime&$search=%22Project%20A%22');
  });

  it('rejects absolute URLs, version prefixes and malformed query encoding', () => {
    for (const path of ['https://graph.microsoft.com/v1.0/me', '//evil.example/me', '/v1.0/me', '/me?$filter=%ZZ']) {
      expect(() => workIqPath(path)).toThrow();
    }
  });

  it.each([
    'HTTP 403', 'HTTP status: 403', '{"statusCode":403}', 'Access denied for path: /me/events',
    'Authorization_RequestDenied', 'Missing scope permissions on the request', 'Insufficient privileges',
  ])('does not classify %s as a refreshable sign-in failure', (message) => {
    expect(classifyFailure(message)).toBe('WORKIQ_PERMISSION_DENIED');
  });

  it('identifies the documented tenant-provisioning error', () => {
    expect(classifyFailure('AADSTS650052')).toBe('WORKIQ_ADMIN_CONSENT_REQUIRED');
  });

  it('does not let sign-in wording override an explicit permission denial', () => {
    expect(classifyFailure('Failed to acquire token. Please sign in.', 403)).toBe('WORKIQ_PERMISSION_DENIED');
    expect(classifyFailure('HTTP 403: auth login required')).toBe('WORKIQ_PERMISSION_DENIED');
  });
});
