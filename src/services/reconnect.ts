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
   * sign-in). Automatic retries stay silent-only; the explicit
   * `hey365_reconnect` tool opts into this.
   */
  allowInteractive?: boolean;
  /** Reconnect even when the current session still works. */
  force?: boolean;
}

const DEFAULT_SILENT_TIMEOUT_MS = 120_000;
const DEFAULT_COOLDOWN_MS = 60_000;

let inFlight: Promise<ReconnectResult> | undefined;
let lastFailureAt = 0;

function silentTimeoutMs(): number {
  const raw = Number(process.env.HEY365_RECONNECT_TIMEOUT_MS);
  return Number.isFinite(raw) && raw > 0 ? raw : DEFAULT_SILENT_TIMEOUT_MS;
}

function cooldownMs(): number {
  const raw = Number(process.env.HEY365_RECONNECT_COOLDOWN_MS);
  return Number.isFinite(raw) && raw >= 0 ? raw : DEFAULT_COOLDOWN_MS;
}

/** Automatic reconnection is on unless explicitly disabled. */
export function autoReconnectEnabled(): boolean {
  const raw = process.env.HEY365_AUTO_RECONNECT?.trim().toLowerCase();
  if (!raw) return true;
  return !['0', 'off', 'false', 'no'].includes(raw);
}

/** Test seam: clears the single-flight guard and the failure cooldown. */
export function resetReconnectState(): void {
  inFlight = undefined;
  lastFailureAt = 0;
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
    lastFailureAt = 0;
    return {
      ok: true,
      mode: 'already-connected',
      ...(health.account ? { account: health.account } : {}),
      detail: `Work IQ は接続済みです（${health.account ?? ''}）。`,
      detailEn: `Work IQ is already connected (${health.account ?? ''}).`,
      health,
    };
  }

  const sinceFailure = Date.now() - lastFailureAt;
  if (!options.force && lastFailureAt > 0 && sinceFailure < cooldownMs()) {
    return {
      ok: false,
      mode: 'cooldown',
      detail: `直前の再接続が失敗したため ${Math.ceil((cooldownMs() - sinceFailure) / 1000)} 秒待機中です。ブラウザでのサインインが必要な可能性があります。`,
      detailEn: `A reconnect attempt failed recently; waiting ${Math.ceil((cooldownMs() - sinceFailure) / 1000)}s before retrying. Browser sign-in may be required.`,
      health,
    };
  }

  if (await silentLogin()) {
    const after = await checkHealth();
    if (after.authenticated) {
      lastFailureAt = 0;
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
    const setup = await runSetup({ interactive: true, force: true });
    if (setup.ok) {
      lastFailureAt = 0;
      return {
        ok: true,
        mode: 'interactive',
        ...(setup.health.account ? { account: setup.health.account } : {}),
        detail: `ブラウザ認証で Work IQ に再接続しました（${setup.health.account ?? ''}）。`,
        detailEn: `Reconnected to Work IQ through browser sign-in (${setup.health.account ?? ''}).`,
        health: setup.health,
      };
    }
    lastFailureAt = Date.now();
    return {
      ok: false,
      mode: 'failed',
      detail: setup.message,
      detailEn: 'Could not reconnect to Work IQ. Run `npx -y @microsoft/workiq auth login` in a terminal and sign in.',
      health: setup.health,
    };
  }

  lastFailureAt = Date.now();
  return {
    ok: false,
    mode: 'failed',
    detail: '自動再接続に失敗しました。`hey365_reconnect` を実行するとブラウザ認証まで試します。',
    detailEn: 'Automatic reconnect failed. Call `hey365_reconnect` to also attempt browser sign-in.',
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
