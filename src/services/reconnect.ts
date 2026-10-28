import { logger } from '../utils/logger.js';
import { checkHealth, runSetup, type HealthReport } from './setup.js';
import { lastKnownAccount } from './store.js';
import { isInteractiveRuntime, runWorkIq } from './workiq.js';
import { closeWorkIqMcp, workIqBackend } from './workiq-mcp.js';
import { isAuthFailure } from '../utils/errors.js';

/**
 * Automatic Work IQ reconnection.
 *
 * Work IQ tokens expire routinely. Before this existed, every expiry surfaced
 * as a plain error, the assistant relayed "please re-authenticate" to the user,
 * and only a second instruction actually triggered the recovery. Reconnecting
 * is a local, idempotent operation that needs no decision from the user, so
 * Hey365 now performs it itself and retries the original call.
 */

export type ReconnectMode =
  /** Work IQ was already usable; nothing was done. */
  | 'already-connected'
  /** The official Work IQ CLI restored access without the fallback path. */
  | 'silent'
  /** The full setup path (broker off + browser sign-in) repaired it. */
  | 'interactive'
  /** A recent attempt already failed; not retried yet. */
  | 'cooldown'
  /** Disabled through HEY365_AUTO_RECONNECT. */
  | 'disabled'
  | 'failed';

export interface ReconnectResult {
  ok: boolean;
  mode: ReconnectMode;
  account?: string;
  officialCliAttempted?: boolean;
  readinessAttempts?: number;
  detail: string;
  detailEn: string;
  health?: HealthReport;
}

export interface ReconnectOptions {
  /**
   * Allow the full recovery path (disable brokered auth, then browser
   * sign-in) when the cached token can no longer be redeemed.
   */
  allowInteractive?: boolean;
  /**
   * Reconnect even when `checkHealth` reports a working session. Automatic
   * recovery always sets this: the caller already *observed* an auth failure,
   * so a health probe that disagrees is stale, and retrying without a fresh
   * token would simply fail a second time.
   */
  force?: boolean;
  /**
   * Skip the post-failure backoff. Only an explicit user request sets this —
   * the backoff is what stops a dead token cache from opening a browser on
   * every single tool call.
   */
  ignoreCooldown?: boolean;
}

const DEFAULT_SILENT_TIMEOUT_MS = 120_000;
const DEFAULT_INTERACTIVE_TIMEOUT_MS = 180_000;
const DEFAULT_COOLDOWN_MS = 60_000;
const DEFAULT_READINESS_ATTEMPTS = 5;
const DEFAULT_READINESS_INTERVAL_MS = 5_000;
const DEFAULT_READINESS_PROBE_TIMEOUT_MS = 30_000;
const MAX_COOLDOWN_MS = 15 * 60_000;

let inFlight: Promise<ReconnectResult> | undefined;
let lastFailureAt = 0;
let consecutiveFailures = 0;

function silentTimeoutMs(): number {
  const raw = Number(process.env.HEY365_RECONNECT_TIMEOUT_MS);
  return Number.isFinite(raw) && raw > 0 ? raw : DEFAULT_SILENT_TIMEOUT_MS;
}

function interactiveTimeoutMs(): number {
  const raw = Number(process.env.HEY365_RECONNECT_INTERACTIVE_TIMEOUT_MS);
  return Number.isFinite(raw) && raw > 0 ? raw : DEFAULT_INTERACTIVE_TIMEOUT_MS;
}

function baseCooldownMs(): number {
  const raw = Number(process.env.HEY365_RECONNECT_COOLDOWN_MS);
  return Number.isFinite(raw) && raw >= 0 ? raw : DEFAULT_COOLDOWN_MS;
}

function readinessAttempts(): number {
  const raw = Number(process.env.HEY365_READINESS_ATTEMPTS);
  if (!Number.isFinite(raw)) return DEFAULT_READINESS_ATTEMPTS;
  return Math.min(Math.max(Math.trunc(raw), 1), 10);
}

function readinessIntervalMs(): number {
  const raw = Number(process.env.HEY365_READINESS_INTERVAL_MS);
  if (!Number.isFinite(raw) || raw < 0) return DEFAULT_READINESS_INTERVAL_MS;
  return raw;
}

function readinessProbeTimeoutMs(): number {
  const raw = Number(process.env.HEY365_READINESS_PROBE_TIMEOUT_MS);
  if (!Number.isFinite(raw) || raw <= 0) return DEFAULT_READINESS_PROBE_TIMEOUT_MS;
  return raw;
}

/**
 * Exponential backoff. Browser sign-in is now attempted automatically, so a
 * genuinely dead cache must not pop a window open once a minute forever.
 */
function cooldownMs(): number {
  const base = baseCooldownMs();
  if (base === 0 || consecutiveFailures <= 1) return base;
  return Math.min(base * 2 ** (consecutiveFailures - 1), MAX_COOLDOWN_MS);
}

function noteFailure(): void {
  lastFailureAt = Date.now();
  consecutiveFailures += 1;
}

function noteSuccess(): void {
  lastFailureAt = 0;
  consecutiveFailures = 0;
}

/** Automatic reconnection is on unless explicitly disabled. */
export function autoReconnectEnabled(): boolean {
  const raw = process.env.HEY365_AUTO_RECONNECT?.trim().toLowerCase();
  if (!raw) return true;
  return !['0', 'off', 'false', 'no'].includes(raw);
}

/** Test seam: clears the single-flight guard and the failure backoff. */
export function resetReconnectState(): void {
  inFlight = undefined;
  lastFailureAt = 0;
  consecutiveFailures = 0;
}

/**
 * Concurrent tool calls share one attempt; spawning several `auth login`
 * processes against the same token cache corrupts it.
 */
export async function reconnect(options: ReconnectOptions = {}): Promise<ReconnectResult> {
  if (inFlight) return await inFlight;
  inFlight = attempt(options).finally(() => {
    inFlight = undefined;
  });
  return await inFlight;
}

async function attempt(options: ReconnectOptions): Promise<ReconnectResult> {
  if (workIqBackend() === 'mcp') return await reconnectHosted(options);
  // A forced reconnect is only requested after the caller has already
  // observed an auth/read failure. Probing before login adds one full Graph
  // round trip and cannot change the decision, so skip directly to login.
  let health: HealthReport | undefined;
  if (!options.force) {
    health = await checkHealth({ deep: true, timeoutMs: readinessProbeTimeoutMs() });
  }
  if (health?.authenticated && !options.force) {
    if (surfacesReady(health)) {
      noteSuccess();
      return {
        ok: true,
        mode: 'already-connected',
        ...(health.account ? { account: health.account } : {}),
        readinessAttempts: 1,
        detail: `Work IQ は接続済みで、Outlook・Teams・Calendar を読み取れます（${health.account ?? ''}）。`,
        detailEn: `Work IQ is connected and Outlook, Teams and Calendar are readable (${health.account ?? ''}).`,
        health,
      };
    }
    logger.info('Work IQ is authenticated but one or more M365 surfaces are not ready; reconnecting', {
      surfaces: health.surfaces,
    });
  }

  async function reconnectHosted(options: ReconnectOptions): Promise<ReconnectResult> {
    if (!options.ignoreCooldown && lastFailureAt > 0 && Date.now() - lastFailureAt < cooldownMs()) {
      return {
        ok: false, mode: 'cooldown',
        detail: '直前の WorkIQ MCP 再接続が失敗したため待機中です。',
        detailEn: 'Waiting after a failed WorkIQ MCP reconnect attempt.',
      };
    }
    await closeWorkIqMcp();
    const health = await checkHealth({ deep: true, timeoutMs: readinessProbeTimeoutMs() });
    if (health.authenticated && health.readAccess) {
      noteSuccess();
      return {
        ok: true, mode: 'silent', health, account: health.account,
        detail: '公式 WorkIQ MCP に再接続しました。',
        detailEn: 'Reconnected to the official WorkIQ MCP.',
      };
    }
    if (health.error && !isAuthFailure(health.error.code)) {
      noteFailure();
      return { ok: false, mode: 'failed', health, detail: health.error.nextStep, detailEn: 'WorkIQ could not be reached or authorized. Inspect the health error; sign-in was not repeated.' };
    }
    if (options.allowInteractive && isInteractiveRuntime()) {
      const setup = await runSetup({ interactive: true, force: true, loginTimeoutMs: interactiveTimeoutMs() });
      if (setup.ok) noteSuccess(); else noteFailure();
      return {
        ok: setup.ok, mode: setup.ok ? 'interactive' : 'failed', health: setup.health,
        account: setup.health.account, detail: setup.message,
        detailEn: setup.ok ? 'Connected to the official WorkIQ MCP.' : 'WorkIQ MCP setup did not complete.',
      };
    }
    noteFailure();
    return {
      ok: false, mode: 'failed', health,
      detail: '公式 WorkIQ MCP のサインインが必要です。ターミナルで `hey365 setup` を実行してください。旧 CLI への切り替えは行いません。',
      detailEn: 'WorkIQ MCP needs sign-in. Run `hey365 setup` in a terminal. No fallback to the legacy CLI was attempted.',
    };
  }

  const wait = cooldownMs();
  const sinceFailure = Date.now() - lastFailureAt;
  if (!options.ignoreCooldown && lastFailureAt > 0 && sinceFailure < wait) {
    const remaining = Math.ceil((wait - sinceFailure) / 1000);
    return {
      ok: false,
      mode: 'cooldown',
      detail: `直前の再接続が失敗したため ${remaining} 秒待機中です。すぐにやり直すには \`hey365_reconnect\` を force=true で実行してください。`,
      detailEn: `A reconnect attempt failed recently; waiting ${remaining}s before retrying. Call \`hey365_reconnect\` with force=true to retry now.`,
      ...(health ? { health } : {}),
    };
  }

  const officialCliOk = await tryOfficialCliRecovery();
  let officialReadiness: ReadinessResult | undefined;
  if (officialCliOk) {
    officialReadiness = await waitForReadiness();
    if (officialReadiness.ready) {
      noteSuccess();
      return {
        ok: true,
        mode: 'silent',
        ...(officialReadiness.health.account ? { account: officialReadiness.health.account } : {}),
        officialCliAttempted: true,
        readinessAttempts: officialReadiness.attempts,
        detail: `公式 Work IQ CLI で再接続し、${officialReadiness.attempts} 回目の確認で Outlook・Teams・Calendar を読み取れるようになりました（${officialReadiness.health.account ?? ''}）。`,
        detailEn: `Reconnected with the official Work IQ CLI; Outlook, Teams and Calendar became readable on readiness check ${officialReadiness.attempts} (${officialReadiness.health.account ?? ''}).`,
        health: officialReadiness.health,
      };
    }
    logger.warn('official Work IQ CLI returned success but M365 reads are still unavailable; continuing to existing recovery', {
      attempts: officialReadiness.attempts,
      unavailable: unavailableSurfaces(officialReadiness.health),
    });
  }

  if (options.allowInteractive) {
    logger.info('official Work IQ CLI recovery did not restore M365 reads; continuing to existing recovery');
    const setup = await runSetup({ interactive: true, force: true, loginTimeoutMs: interactiveTimeoutMs() });
    if (setup.ok) {
      const readiness = await waitForReadiness();
      if (readiness.ready) {
        noteSuccess();
        return {
          ok: true,
          mode: 'interactive',
          ...(readiness.health.account ? { account: readiness.health.account } : {}),
          officialCliAttempted: true,
          readinessAttempts: readiness.attempts,
          detail: `公式 CLI で復旧できなかったため既存のブラウザ復旧へ進み、${readiness.attempts} 回目の確認で Outlook・Teams・Calendar を読み取れるようになりました（${readiness.health.account ?? ''}）。`,
          detailEn: `The official CLI did not restore access, so the existing browser recovery ran; Outlook, Teams and Calendar became readable on readiness check ${readiness.attempts} (${readiness.health.account ?? ''}).`,
          health: readiness.health,
        };
      }
      return failedReadiness(readiness);
    }
    noteFailure();
    health ??= await checkHealth({ deep: true, timeoutMs: readinessProbeTimeoutMs() });
    return {
      ok: false,
      mode: 'failed',
      officialCliAttempted: true,
      detail: setup.message,
      detailEn: 'Could not reconnect to Work IQ. Run `npx -y @microsoft/workiq auth login` in a terminal and sign in.',
      health: setup.health,
    };
  }

  if (officialReadiness) return failedReadiness(officialReadiness);

  function surfacesReady(health: HealthReport): boolean {
    return (
      health.authenticated &&
      health.surfaces?.outlook === true &&
      health.surfaces.teams === true &&
      health.surfaces.calendar === true
    );
  }

  function unavailableSurfaces(health: HealthReport): string[] {
    if (!health.authenticated) return ['authentication'];
    const unavailable: string[] = [];
    if (health.surfaces?.outlook !== true) unavailable.push('Outlook');
    if (health.surfaces?.teams !== true) unavailable.push('Teams');
    if (health.surfaces?.calendar !== true) unavailable.push('Calendar');
    return unavailable;
  }

  interface ReadinessResult {
    ready: boolean;
    attempts: number;
    health: HealthReport;
  }

  async function waitForReadiness(): Promise<ReadinessResult> {
    const maxAttempts = readinessAttempts();
    let health: HealthReport | undefined;

    for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
      health = await checkHealth({ deep: true, timeoutMs: readinessProbeTimeoutMs() });
      if (surfacesReady(health)) return { ready: true, attempts: attempt, health };

      const unavailable = unavailableSurfaces(health);
      logger.warn('Work IQ reconnect completed but M365 surfaces are not ready yet', {
        attempt,
        maxAttempts,
        unavailable,
      });
      if (attempt < maxAttempts) await sleep(readinessIntervalMs());
    }

    return { ready: false, attempts: maxAttempts, health: health ?? (await checkHealth({ deep: true })) };
  }

  function failedReadiness(readiness: ReadinessResult): ReconnectResult {
    noteFailure();
    const health = readiness.health;
    const unavailable = unavailableSurfaces(health).join('・') || 'Microsoft 365';
    return {
      ok: false,
      mode: 'failed',
      officialCliAttempted: true,
      ...(health.account ? { account: health.account } : {}),
      readinessAttempts: readiness.attempts,
      detail: `再接続後に ${readiness.attempts} 回確認しましたが、${unavailable} をまだ読み取れません。次回の呼び出しでも自動的に再試行します。`,
      detailEn: `Checked readiness ${readiness.attempts} times after reconnecting, but ${unavailable} is still unavailable. The next call will retry automatically.`,
      health,
    };
  }

  async function sleep(ms: number): Promise<void> {
    if (ms <= 0) return;
    await new Promise<void>((resolve) => setTimeout(resolve, ms));
  }

  noteFailure();
  return {
    ok: false,
    mode: 'failed',
    officialCliAttempted: true,
    detail: 'キャッシュされたトークンでは再接続できませんでした。`hey365_reconnect` を実行するとブラウザ認証まで試します。',
    detailEn: 'The cached token could not be redeemed. Call `hey365_reconnect` to also attempt browser sign-in.',
    health,
  };
}

/**
 * First-line recovery through Microsoft's official Work IQ CLI.
 *
 * Work IQ CLI 1.0.1 does not expose a `--force` option. `force=true` is
 * therefore enforced by Hey365: this official command is run unconditionally
 * before the existing broker-disable/browser recovery, even when an earlier
 * health check reported a connected session.
 */
async function tryOfficialCliRecovery(): Promise<boolean> {
  const account = process.env.HEY365_WORKIQ_ACCOUNT?.trim() || lastKnownAccount();
  const args = ['auth', 'login'];
  if (account) args.push('--account', account);

  try {
    logger.info('attempting recovery with the official Work IQ CLI', {
      account: Boolean(account),
      forcedByHey365: true,
    });
    const result = await runWorkIq(args, { timeoutMs: silentTimeoutMs() });
    if (result.timedOut) {
      logger.warn('official Work IQ CLI recovery timed out');
      return false;
    }
    return result.code === 0;
  } catch (error) {
    logger.debug('official Work IQ CLI recovery failed', { error: (error as Error).message });
    return false;
  }
}

/** Human-readable summary used by the `hey365_reconnect` tool and the CLI. */
export function describeReconnect(result: ReconnectResult): string {
  const icon = result.ok ? '✅' : '❌';
  const lines = [`${icon} ${result.detail}`, '', `mode: ${result.mode}`];
  if (result.health) {
    lines.push(
      `authenticated: ${result.health.authenticated}`,
      `readAccess: ${result.health.readAccess}`,
      `writeAccess: ${result.health.writeAccess}`,
    );
  }
  if (!result.ok && !isInteractiveRuntime() && workIqBackend() === 'cli') {
    lines.push('', 'ターミナルで `npx -y @microsoft/workiq auth login` を実行するとブラウザ認証を完了できます。');
  }
  return lines.join('\n');
}
