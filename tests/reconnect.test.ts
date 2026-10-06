import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { errorCodeOf, withAuthRecovery, type ToolResponse } from '../src/tools/index.js';
import { Hey365Error, classifyFailure, isAuthFailure } from '../src/utils/errors.js';
import { resetReconnectState } from '../src/services/reconnect.js';

vi.mock('../src/services/reconnect.js', async (importOriginal) => {
  const actual = await importOriginal<typeof import('../src/services/reconnect.js')>();
  return {
    ...actual,
    reconnect: vi.fn(),
    autoReconnectEnabled: vi.fn(() => true),
  };
});

const { autoReconnectEnabled, reconnect } = await import('../src/services/reconnect.js');
const reconnectMock = vi.mocked(reconnect);
const enabledMock = vi.mocked(autoReconnectEnabled);

function authFailure(code: 'WORKIQ_AUTH_EXPIRED' | 'WORKIQ_NOT_AUTHENTICATED' = 'WORKIQ_AUTH_EXPIRED'): ToolResponse {
  const error = new Hey365Error(code);
  return {
    content: [{ type: 'text', text: `⚠ ${error.nextStep('ja')}` }],
    structuredContent: { error: error.toPayload() },
    isError: true,
  };
}

function success(text = 'ok'): ToolResponse {
  return { content: [{ type: 'text', text }] };
}

describe('auth failure classification', () => {
  it('treats only the two sign-in codes as self-healing', () => {
    expect(isAuthFailure('WORKIQ_AUTH_EXPIRED')).toBe(true);
    expect(isAuthFailure('WORKIQ_NOT_AUTHENTICATED')).toBe(true);
    expect(isAuthFailure('WORKIQ_PERMISSION_DENIED')).toBe(false);
    expect(isAuthFailure('WORKIQ_NOT_INSTALLED')).toBe(false);
  });

  it('recognises the Work IQ and MSAL phrasings of an expired session', () => {
    expect(classifyFailure('AADSTS700081: The refresh token has expired')).toBe('WORKIQ_AUTH_EXPIRED');
    expect(classifyFailure('MsalUiRequiredException: interaction_required')).toBe('WORKIQ_AUTH_EXPIRED');
    expect(classifyFailure('InvalidAuthenticationToken')).toBe('WORKIQ_NOT_AUTHENTICATED');
    expect(classifyFailure('No account found in the token cache')).toBe('WORKIQ_NOT_AUTHENTICATED');
    expect(classifyFailure('Failed to acquire token')).toBe('WORKIQ_NOT_AUTHENTICATED');
  });

  it('carries a machine-readable recovery hint on auth errors only', () => {
    expect(new Hey365Error('WORKIQ_AUTH_EXPIRED').toPayload().recovery).toEqual({
      tool: 'hey365_reconnect',
      automatic: true,
      retryOriginalCall: true,
      requiresUserConfirmation: false,
    });
    expect(new Hey365Error('SEND_FAILED').toPayload().recovery).toBeUndefined();
  });

  it('reads the code back out of a tool response', () => {
    expect(errorCodeOf(authFailure())).toBe('WORKIQ_AUTH_EXPIRED');
    expect(errorCodeOf(success())).toBeUndefined();
  });
});

describe('withAuthRecovery', () => {
  beforeEach(() => {
    resetReconnectState();
    reconnectMock.mockReset();
    enabledMock.mockReset();
    enabledMock.mockReturnValue(true);
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  it('passes a successful call straight through without reconnecting', async () => {
    const handler = vi.fn(async () => success('triage'));
    const result = await withAuthRecovery('hey365', handler)({});

    expect(result.isError).toBeUndefined();
    expect(handler).toHaveBeenCalledTimes(1);
    expect(reconnectMock).not.toHaveBeenCalled();
  });

  it('does not reconnect for non-auth failures', async () => {
    const error = new Hey365Error('SEND_FAILED');
    const handler = vi.fn(async (): Promise<ToolResponse> => ({
      content: [{ type: 'text', text: 'nope' }],
      structuredContent: { error: error.toPayload() },
      isError: true,
    }));

    await withAuthRecovery('hey365_send', handler)({});

    expect(handler).toHaveBeenCalledTimes(1);
    expect(reconnectMock).not.toHaveBeenCalled();
  });

  it('reconnects and retries the original call once', async () => {
    const handler = vi
      .fn<[unknown], Promise<ToolResponse>>()
      .mockResolvedValueOnce(authFailure())
      .mockResolvedValueOnce(success('2件'));
    reconnectMock.mockResolvedValue({
      ok: true,
      mode: 'silent',
      account: 'user@example.com',
      detail: '再接続しました',
      detailEn: 'reconnected',
    });

    const result = await withAuthRecovery('hey365', handler)({ hours: 36 });

    expect(handler).toHaveBeenCalledTimes(2);
    expect(handler.mock.calls[1]?.[0]).toEqual({ hours: 36 });
    expect(result.isError).toBeUndefined();
    expect(result.content[0]?.text).toContain('再接続');
    expect(result.content[0]?.text).toContain('2件');
  });

  it('does not retry a second time when the call still fails after reconnecting', async () => {
    const handler = vi.fn(async () => authFailure());
    reconnectMock.mockResolvedValue({
      ok: true,
      mode: 'silent',
      detail: 'ok',
      detailEn: 'ok',
    });

    const result = await withAuthRecovery('hey365', handler)({});

    expect(handler).toHaveBeenCalledTimes(2);
    expect(result.isError).toBe(true);
  });

  it('tells the assistant which tool to call when reconnecting fails', async () => {
    const handler = vi.fn(async () => authFailure('WORKIQ_NOT_AUTHENTICATED'));
    reconnectMock.mockResolvedValue({
      ok: false,
      mode: 'failed',
      detail: 'ブラウザ認証が必要です',
      detailEn: 'browser sign-in required',
    });

    const result = await withAuthRecovery('hey365_digest', handler)({});
    const recovery = (result.structuredContent as { error: { recovery: Record<string, unknown> } }).error.recovery;

    expect(result.isError).toBe(true);
    expect(result.content[0]?.text).toContain('hey365_reconnect');
    expect(result.content[0]?.text).toContain('hey365_digest');
    expect(recovery).toMatchObject({
      tool: 'hey365_reconnect',
      attempted: true,
      retryOriginalCall: true,
      originalTool: 'hey365_digest',
      requiresUserConfirmation: false,
      lastAttempt: 'failed',
    });
  });

  it('still reports the recovery path when auto-reconnect is disabled', async () => {
    enabledMock.mockReturnValue(false);
    const handler = vi.fn(async () => authFailure());

    const result = await withAuthRecovery('hey365', handler)({});
    const recovery = (result.structuredContent as { error: { recovery: Record<string, unknown> } }).error.recovery;

    expect(handler).toHaveBeenCalledTimes(1);
    expect(reconnectMock).not.toHaveBeenCalled();
    expect(recovery).toMatchObject({ tool: 'hey365_reconnect', attempted: false });
  });
});
