import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { Hey365Error } from '../src/utils/errors.js';

vi.mock('../src/services/workiq.js', () => ({
  describeCommand: () => 'https://workiq.svc.cloud.microsoft/mcp',
  fetchMany: vi.fn(),
  isInteractiveRuntime: vi.fn(() => false),
  resetCommandCache: vi.fn(),
  resolveWorkIqCommand: vi.fn(),
  runWorkIq: vi.fn(),
}));
vi.mock('../src/services/workiq-mcp.js', async (original) => ({
  ...await original<typeof import('../src/services/workiq-mcp.js')>(),
  loginWorkIqMcp: vi.fn(),
}));
vi.mock('../src/services/store.js', () => ({ rememberAccount: vi.fn() }));

const { fetchMany, isInteractiveRuntime, runWorkIq } = await import('../src/services/workiq.js');
const { loginWorkIqMcp } = await import('../src/services/workiq-mcp.js');
const { runSetup } = await import('../src/services/setup.js');
const reads = vi.mocked(fetchMany);
const login = vi.mocked(loginWorkIqMcp);
const healthy = [
  { ok: true, data: { id: 'me' } },
  { ok: true, data: { value: [] } }, { ok: true, data: { value: [] } }, { ok: true, data: { value: [] } },
];

describe('hosted setup', () => {
  beforeEach(() => {
    vi.stubEnv('HEY365_WORKIQ_BACKEND', 'mcp');
    vi.clearAllMocks();
    vi.mocked(isInteractiveRuntime).mockReturnValue(false);
  });
  afterEach(() => vi.unstubAllEnvs());

  it('recognises a connected session without running CLI commands or OAuth', async () => {
    reads.mockResolvedValue(healthy);
    expect((await runSetup()).ok).toBe(true);
    expect(login).not.toHaveBeenCalled();
    expect(runWorkIq).not.toHaveBeenCalled();
  });

  it('does not call partial permission-denied access a successful setup', async () => {
    reads.mockResolvedValue([
      ...healthy.slice(0, 3),
      { ok: false, error: new Hey365Error('WORKIQ_PERMISSION_DENIED', 'HTTP 403 on calendarView') },
    ]);
    expect((await runSetup({ force: true })).ok).toBe(false);
    expect(login).not.toHaveBeenCalled();
    expect(runWorkIq).not.toHaveBeenCalled();
  });

  it('directs a stdio host to terminal setup, without opening a browser', async () => {
    reads.mockResolvedValue([{ ok: false, error: new Hey365Error('WORKIQ_NOT_AUTHENTICATED') }]);
    const result = await runSetup({ interactive: true });
    expect(result.ok).toBe(false);
    expect(result.message).toContain('hey365 setup');
    expect(login).not.toHaveBeenCalled();
  });

  it('signs in from the CLI and verifies all read surfaces before success', async () => {
    vi.mocked(isInteractiveRuntime).mockReturnValue(true);
    reads.mockResolvedValueOnce([{ ok: false, error: new Hey365Error('WORKIQ_NOT_AUTHENTICATED') }])
      .mockResolvedValueOnce(healthy);
    login.mockResolvedValue();
    const result = await runSetup({ interactive: true, loginTimeoutMs: 30_000 });
    expect(result.ok).toBe(true);
    expect(result.health.readAccess).toBe(true);
    expect(login).toHaveBeenCalledWith(30_000);
    expect(runWorkIq).not.toHaveBeenCalled();
  });

  it('does not try to repair a credential-store failure by repeating sign-ins', async () => {
    vi.mocked(isInteractiveRuntime).mockReturnValue(true);
    reads.mockResolvedValue([{ ok: false, error: new Hey365Error('WORKIQ_CREDENTIAL_STORE_ERROR') }]);
    expect((await runSetup()).ok).toBe(false);
    expect(login).not.toHaveBeenCalled();
  });
});
