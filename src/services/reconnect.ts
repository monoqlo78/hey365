import { logger } from '../utils/logger.js';
import { checkHealth, runSetup, type HealthReport } from './setup.js';
import { lastKnownAccount } from './store.js';
import { isInteractiveRuntime, runWorkIq } from './workiq.js';

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
  /** A cached refresh token was redeemed without user interaction. */
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
  const health = await checkHealth();
  if (health.authenticated && !options.force) {
    noteSuccess();
    return {
      ok: true,
      mode: 'already-connected',
      ...(health.account ? { account: health.account } : {}),
      detail: `Work IQ は接続済みです（${health.account ?? ''}）。`,
      detailEn: `Work IQ is already connected (${health.account ?? ''}).`,
      health,
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
      health,
    };
  }

  if (await silentLogin()) {
    const after = await checkHealth();
    if (after.authenticated) {
      noteSuccess();
      return {
        ok: true,
        mode: 'silent',
        ...(after.account ? { account: after.account } : {}),
        detail: `Work IQ に再接続しました（${after.account ?? ''}）。`,
        detailEn: `Reconnected to Work IQ (${after.account ?? ''}).`,
        health: after,
      };
    }
  }

  if (options.allowInteractive) {
    logger.info('cached token could not be redeemed, escalating to browser sign-in');
    const setup = await runSetup({ interactive: true, force: true, loginTimeoutMs: interactiveTimeoutMs() });
    if (setup.ok) {
      noteSuccess();
      return {
        ok: true,
        mode: 'interactive',
        ...(setup.health.account ? { account: setup.health.account } : {}),
        detail: `ブラウザ認証で Work IQ に再接続しました（${setup.health.account ?? ''}）。`,
        detailEn: `Reconnected to Work IQ through browser sign-in (${setup.health.account ?? ''}).`,
        health: setup.health,
      };
    }
    noteFailure();
    return {
      ok: false,
      mode: 'failed',
      detail: setup.message,
      detailEn: 'Could not reconnect to Work IQ. Run `npx -y @microsoft/workiq auth login` in a terminal and sign in.',
      health: setup.health,
    };
  }

  noteFailure();
  return {
    ok: false,
    mode: 'failed',
    detail: 'キャッシュされたトークンでは再接続できませんでした。`hey365_reconnect` を実行するとブラウザ認証まで試します。',
    detailEn: 'The cached token could not be redeemed. Call `hey365_reconnect` to also attempt browser sign-in.',
    health,
  };
}

/**
 * Redeems the cached refresh token. stdin is never inherited, so the CLI can
 * neither block on a console prompt nor read the MCP transport; when a cached
 * token is still valid this returns in a second or two.
 */
async function silentLogin(): Promise<boolean> {
  const account = process.env.HEY365_WORKIQ_ACCOUNT?.trim() || lastKnownAccount();
  const args = ['auth', 'login'];
  if (account) args.push('--account', account);

  try {
    logger.info('attempting silent Work IQ reconnect', { account: Boolean(account) });
    const result = await runWorkIq(args, { timeoutMs: silentTimeoutMs() });
    if (result.timedOut) {
      logger.warn('silent reconnect timed out');
      return false;
    }
    return result.code === 0;
  } catch (error) {
    logger.debug('silent reconnect failed', { error: (error as Error).message });
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
  if (!result.ok && !isInteractiveRuntime()) {
    lines.push('', 'ターミナルで `npx -y @microsoft/workiq auth login` を実行するとブラウザ認証を完了できます。');
  }
  return lines.join('\n');
}
