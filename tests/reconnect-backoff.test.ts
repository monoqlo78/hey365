import { beforeEach, describe, expect, it, vi } from 'vitest';

vi.mock('../src/services/setup.js', () => ({
  checkHealth: vi.fn(),
  runSetup: vi.fn(),
}));

vi.mock('../src/services/workiq.js', () => ({
  runWorkIq: vi.fn(),
  isInteractiveRuntime: () => false,
}));

vi.mock('../src/services/store.js', () => ({
  lastKnownAccount: () => 'user@example.com',
}));

const { checkHealth, runSetup } = await import('../src/services/setup.js');
const { runWorkIq } = await import('../src/services/workiq.js');
const { reconnect, resetReconnectState } = await import('../src/services/reconnect.js');

const healthMock = vi.mocked(checkHealth);
const setupMock = vi.mocked(runSetup);
const workIqMock = vi.mocked(runWorkIq);

const connected = { authenticated: true, account: 'user@example.com' } as never;
const disconnected = { authenticated: false } as never;

function loginResult(code: number) {
  return { code, stdout: '', stderr: '', timedOut: false } as never;
}

describe('reconnect', () => {
  beforeEach(() => {
    resetReconnectState();
    healthMock.mockReset();
    setupMock.mockReset();
    workIqMock.mockReset();
    delete process.env.HEY365_RECONNECT_COOLDOWN_MS;
  });

  it('skips the work entirely when the session is healthy', async () => {
    healthMock.mockResolvedValue(connected);

    const result = await reconnect();

    expect(result.mode).toBe('already-connected');
    expect(workIqMock).not.toHaveBeenCalled();
  });

  it('re-authenticates even when health claims the session is fine, under force', async () => {
    healthMock.mockResolvedValue(connected);
    workIqMock.mockResolvedValue(loginResult(0));

    const result = await reconnect({ force: true });

    expect(result.mode).toBe('silent');
    expect(workIqMock).toHaveBeenCalledWith(['auth', 'login', '--account', 'user@example.com'], expect.anything());
  });

  it('redeems the cached token without touching the browser path', async () => {
    healthMock.mockResolvedValueOnce(disconnected).mockResolvedValueOnce(connected);
    workIqMock.mockResolvedValue(loginResult(0));

    const result = await reconnect({ allowInteractive: true });

    expect(result.mode).toBe('silent');
    expect(setupMock).not.toHaveBeenCalled();
  });

  it('escalates to browser sign-in when the cached token is dead', async () => {
    healthMock.mockResolvedValue(disconnected);
    workIqMock.mockResolvedValue(loginResult(1));
    setupMock.mockResolvedValue({ ok: true, steps: [], health: connected, message: 'ok' });

    const result = await reconnect({ allowInteractive: true });

    expect(result.mode).toBe('interactive');
    expect(setupMock).toHaveBeenCalledWith(
      expect.objectContaining({ interactive: true, force: true, loginTimeoutMs: expect.any(Number) }),
    );
  });

  it('does not open a browser when escalation is not allowed', async () => {
    healthMock.mockResolvedValue(disconnected);
    workIqMock.mockResolvedValue(loginResult(1));

    const result = await reconnect();

    expect(result.ok).toBe(false);
    expect(setupMock).not.toHaveBeenCalled();
  });

  it('backs off after a failure instead of retrying on every call', async () => {
    healthMock.mockResolvedValue(disconnected);
    workIqMock.mockResolvedValue(loginResult(1));

    expect((await reconnect()).mode).toBe('failed');
    expect((await reconnect()).mode).toBe('cooldown');
    expect(workIqMock).toHaveBeenCalledTimes(1);
  });

  it('lets an explicit request retry immediately despite the backoff', async () => {
    healthMock.mockResolvedValue(disconnected);
    workIqMock.mockResolvedValue(loginResult(1));

    await reconnect();
    const result = await reconnect({ ignoreCooldown: true });

    expect(result.mode).toBe('failed');
    expect(workIqMock).toHaveBeenCalledTimes(2);
  });

  it('lengthens the backoff on repeated failures', async () => {
    process.env.HEY365_RECONNECT_COOLDOWN_MS = '1000';
    healthMock.mockResolvedValue(disconnected);
    workIqMock.mockResolvedValue(loginResult(1));

    await reconnect();
    const first = await reconnect();
    await reconnect({ ignoreCooldown: true });
    const second = await reconnect();

    const seconds = (detail: string) => Number(/(\d+) 秒/.exec(detail)?.[1]);
    expect(seconds(second.detail)).toBeGreaterThan(seconds(first.detail));
  });

  it('clears the backoff once a reconnect succeeds', async () => {
    healthMock.mockResolvedValue(disconnected);
    workIqMock.mockResolvedValue(loginResult(1));
    await reconnect();

    healthMock.mockResolvedValueOnce(disconnected).mockResolvedValueOnce(connected);
    workIqMock.mockResolvedValue(loginResult(0));
    expect((await reconnect({ ignoreCooldown: true })).mode).toBe('silent');

    healthMock.mockResolvedValue(disconnected);
    workIqMock.mockResolvedValue(loginResult(1));
    expect((await reconnect()).mode).toBe('failed');
  });

  it('shares one attempt between concurrent callers', async () => {
    healthMock.mockResolvedValue(disconnected);
    let resolveLogin: (value: unknown) => void = () => {};
    workIqMock.mockReturnValue(
      new Promise((resolve) => {
        resolveLogin = resolve;
      }) as never,
    );

    const both = Promise.all([reconnect(), reconnect()]);
    resolveLogin(loginResult(1));
    await both;

    expect(workIqMock).toHaveBeenCalledTimes(1);
  });
});
