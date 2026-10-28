import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { Hey365Error } from '../src/utils/errors.js';

vi.mock('../src/services/workiq.js', () => ({
  describeCommand: () => 'workiq',
  fetchMany: vi.fn(),
  resetCommandCache: vi.fn(),
  resolveWorkIqCommand: vi.fn(),
  runWorkIq: vi.fn(),
}));

vi.mock('../src/services/store.js', () => ({
  rememberAccount: vi.fn(),
}));

const { fetchMany, runWorkIq } = await import('../src/services/workiq.js');
const { rememberAccount } = await import('../src/services/store.js');
const { checkHealth } = await import('../src/services/setup.js');

const fetchManyMock = vi.mocked(fetchMany);
const runWorkIqMock = vi.mocked(runWorkIq);
const rememberAccountMock = vi.mocked(rememberAccount);

const ok = (data: unknown) => ({ ok: true, data });
const failed = () => ({ ok: false, error: new Hey365Error('WORKIQ_CONNECTION_ERROR', 'not ready') });

describe('checkHealth', () => {
  beforeEach(() => {
    vi.stubEnv('HEY365_WORKIQ_BACKEND', 'cli');
    fetchManyMock.mockReset();
    runWorkIqMock.mockReset();
    rememberAccountMock.mockReset();
    runWorkIqMock.mockResolvedValue({
      code: 0,
      stdout: '1.0.1',
      stderr: '',
      timedOut: false,
    });
    afterEach(() => vi.unstubAllEnvs());
  });

  it('checks every M365 surface in one batched Work IQ invocation', async () => {
    fetchManyMock.mockResolvedValue([
      ok({ id: 'me', mail: 'user@example.com' }),
      ok({ id: 'drafts' }),
      ok({ value: [] }),
      ok({ value: [] }),
      ok({ value: [] }),
    ] as never);

    const report = await checkHealth({ deep: true, timeoutMs: 30_000 });

    expect(report.surfaces).toEqual({ outlook: true, teams: true, calendar: true });
    expect(report.writeAccess).toBe(true);
    expect(fetchManyMock).toHaveBeenCalledTimes(1);
    expect(fetchManyMock.mock.calls[0]?.[0]).toHaveLength(5);
    expect(fetchManyMock.mock.calls[0]?.[1]).toEqual({ timeoutMs: 30_000 });
    expect(rememberAccountMock).toHaveBeenCalledWith('user@example.com');
  });

  it('does not call an empty collection unavailable, but does flag a failed surface', async () => {
    fetchManyMock.mockResolvedValue([
      ok({ id: 'me', userPrincipalName: 'user@example.com' }),
      ok({ id: 'drafts' }),
      ok({ value: [] }),
      failed(),
      ok({ value: [] }),
    ] as never);

    const report = await checkHealth({ deep: true });

    expect(report.authenticated).toBe(true);
    expect(report.readAccess).toBe(false);
    expect(report.surfaces).toEqual({ outlook: true, teams: false, calendar: true });
  });

  it('uses hosted reads without launching the CLI or claiming send permission', async () => {
    vi.stubEnv('HEY365_WORKIQ_BACKEND', 'mcp');
    fetchManyMock.mockResolvedValue([
      ok({ id: 'me' }), ok({ value: [] }), ok({ value: [] }), ok({ value: [] }),
    ]);
    const report = await checkHealth({ deep: true });
    expect(report.backend).toBe('mcp');
    expect(report.authenticated).toBe(true);
    expect(report.readAccess).toBe(true);
    expect(report.writeAccess).toBe('unknown');
    expect(fetchManyMock.mock.calls[0]?.[0]).toHaveLength(4);
    expect(runWorkIqMock).not.toHaveBeenCalled();
  });

  it('preserves a 403 as a permission failure, not an expired sign-in', async () => {
    vi.stubEnv('HEY365_WORKIQ_BACKEND', 'mcp');
    fetchManyMock.mockResolvedValue([{ ok: false, error: new Hey365Error('WORKIQ_PERMISSION_DENIED', 'HTTP 403') }]);
    const report = await checkHealth();
    expect(report.workiq).toBe('error');
    expect(report.error?.code).toBe('WORKIQ_PERMISSION_DENIED');
    expect(report.error?.message).toContain('403');
  });
});
