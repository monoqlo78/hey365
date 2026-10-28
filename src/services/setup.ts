import { Hey365Error, asHey365Error, classifyFailure, isAuthFailure, type Hey365ErrorCode } from '../utils/errors.js';
import { logger } from '../utils/logger.js';
import { timezone } from '../utils/time.js';
import { describeCommand, fetchMany, isInteractiveRuntime, resetCommandCache, resolveWorkIqCommand, runWorkIq } from './workiq.js';
import { rememberAccount } from './store.js';
import type { GraphUser } from './graph.js';
import { loginWorkIqMcp, workIqBackend, WORKIQ_CATALOG } from './workiq-mcp.js';

export interface HealthReport {
  hey365: 'ok' | 'degraded';
  workiq: 'connected' | 'unauthenticated' | 'not-installed' | 'error';
  authenticated: boolean;
  readAccess: boolean;
  writeAccess: boolean | 'unknown';
  backend: 'mcp' | 'cli';
  workiqCatalog?: typeof WORKIQ_CATALOG;
  timezone: string;
  account?: string;
  workiqCommand: string;
  workiqVersion?: string;
  surfaces?: { outlook: boolean; teams: boolean; calendar: boolean };
  error?: { code: Hey365ErrorCode; nextStep: string; message?: string };
}

const ME_URL = '/me?$select=id,displayName,mail,userPrincipalName';
const DRAFTS_URL = '/me/mailFolders/Drafts?$select=id';
const SURFACE_URLS = [
  '/me/messages?$top=1&$select=id',
  '/me/chats?$top=1&$select=id',
  '/me/events?$top=1&$select=id',
] as const;

/** Probes Work IQ and reports what Hey365 can currently do (spec §20). */
export async function checkHealth(options: { deep?: boolean; timeoutMs?: number } = {}): Promise<HealthReport> {
  const backend = workIqBackend();
  const report: HealthReport = {
    hey365: 'ok',
    workiq: 'error',
    authenticated: false,
    readAccess: false,
    writeAccess: backend === 'mcp' ? 'unknown' : false,
    backend,
    ...(backend === 'mcp' ? { workiqCatalog: WORKIQ_CATALOG } : {}),
    timezone: timezone(),
    workiqCommand: describeCommand(),
  };

  const baseUrls = backend === 'mcp' ? [ME_URL] : [ME_URL, DRAFTS_URL];
  const now = new Date();
  const calendarProbe = `/me/calendarView?startDateTime=${encodeURIComponent(now.toISOString())}` +
    `&endDateTime=${encodeURIComponent(new Date(now.getTime() + 24 * 3600_000).toISOString())}&$select=id&$top=1`;
  const surfaceUrls = backend === 'mcp' ? [...SURFACE_URLS.slice(0, 2), calendarProbe] : SURFACE_URLS;
  const urls = options.deep ? [...baseUrls, ...surfaceUrls] : baseUrls;
  const versionPromise = backend === 'cli'
    ? runWorkIq(['version'], { timeoutMs: 60_000 }).catch(() => undefined)
    : Promise.resolve(undefined);

  let outcomes;
  try {
    outcomes = await fetchMany(urls, { timeoutMs: options.timeoutMs ?? 60_000 });
  } catch (error) {
    const hey = asHey365Error(error);
    report.workiq = healthState(hey.code);
    report.hey365 = 'degraded';
    report.error = { code: hey.code, nextStep: hey.nextStep('ja'), message: hey.message };
    return report;
  }

  const version = await versionPromise;
  if (version?.code === 0) {
    const match = /\d+\.\d+\.\d+(?:[-+][\w.]+)?/.exec(version.stdout);
    if (match) report.workiqVersion = match[0];
  }

  const meOutcome = outcomes[0];
  const me = meOutcome?.ok ? (meOutcome.data as GraphUser | undefined) : undefined;
  if (!me?.id) {
    const code = meOutcome?.error?.code ?? classifyFailure(meOutcome?.error?.message ?? '');
    report.workiq = healthState(code);
    report.hey365 = 'degraded';
    const error = meOutcome?.error ?? new Hey365Error(code);
    report.error = { code: error.code, nextStep: error.nextStep('ja'), message: error.message };
    return report;
  }

  report.workiq = 'connected';
  report.authenticated = true;
  report.readAccess = true;
  report.account = me.mail ?? me.userPrincipalName ?? '';
  rememberAccount(report.account);

  if (options.deep) {
    const offset = baseUrls.length;
    report.surfaces = {
      outlook: outcomes[offset]?.ok === true,
      teams: outcomes[offset + 1]?.ok === true,
      calendar: outcomes[offset + 2]?.ok === true,
    };
    report.readAccess = Object.values(report.surfaces).every(Boolean);
  }

  // Keep the legacy report, but do not infer hosted send rights from a read.
  if (backend === 'cli') {
    report.writeAccess = outcomes[1]?.ok === true;
    if (!report.writeAccess) {
      const error = new Hey365Error('WORKIQ_WRITE_DISABLED');
      report.error = { code: error.code, nextStep: error.nextStep('ja') };
    }
  }
  const denied = outcomes.find((outcome) => !outcome.ok && (
    outcome.error?.code === 'WORKIQ_PERMISSION_DENIED' || outcome.error?.code === 'WORKIQ_ADMIN_CONSENT_REQUIRED'
  )) ?? outcomes.find((outcome) => !outcome.ok && outcome.error);
  if (denied?.error) {
    report.hey365 = 'degraded';
    report.error = { code: denied.error.code, nextStep: denied.error.nextStep('ja'), message: denied.error.message };
  }

  return report;
}

function healthState(code: string): HealthReport['workiq'] {
  if (code === 'WORKIQ_NOT_INSTALLED') return 'not-installed';
  return code === 'WORKIQ_NOT_AUTHENTICATED' || code === 'WORKIQ_AUTH_EXPIRED' ? 'unauthenticated' : 'error';
}

export interface SetupStep {
  step: string;
  ok: boolean;
  detail?: string;
}

export interface SetupResult {
  ok: boolean;
  steps: SetupStep[];
  health: HealthReport;
  message: string;
}

/**
 * Hosted OAuth is the default. Only explicit CLI mode uses the old
 * install/broker recovery path. Nothing here logs tokens.
 */
export async function runSetup(
  options: { interactive?: boolean; force?: boolean; loginTimeoutMs?: number } = {},
): Promise<SetupResult> {
  const steps: SetupStep[] = [];

  let health = await checkHealth({ deep: workIqBackend() === 'mcp' });
  if (workIqBackend() === 'mcp') {
    if (health.authenticated && health.readAccess && !options.force) {
      return { ok: true, steps: [{ step: 'hosted-mcp-connected', ok: true }], health, message: '公式 WorkIQ MCP に接続済みです。' };
    }
    if (health.error && !isAuthFailure(health.error.code)) {
      // Only a sign-in failure can be repaired by starting OAuth.
      return { ok: false, steps: [], health, message: health.error.nextStep };
    }
    if (options.interactive === false || !isInteractiveRuntime()) {
      return {
        ok: false, steps: [], health,
        message: 'Hey365 専用の WorkIQ サインインが必要です。ターミナルで `hey365 setup` を実行してください。Copilot アプリや旧 CLI のトークンは共有しません。',
      };
    }
    try {
      await loginWorkIqMcp(options.loginTimeoutMs);
      steps.push({ step: 'hosted-mcp-oauth', ok: true });
    } catch (error) {
      const hey = asHey365Error(error);
      return { ok: false, steps: [{ step: 'hosted-mcp-oauth', ok: false, detail: hey.message }], health, message: hey.nextStep('ja') };
    }
    health = await checkHealth({ deep: true });
    return {
      ok: health.authenticated && health.readAccess,
      steps, health,
      message: health.authenticated && health.readAccess
        ? `公式 WorkIQ MCP に接続しました（${health.account ?? ''}）。`
        : health.error?.nextStep ?? 'WorkIQ MCP の読み取りを確認できませんでした。',
    };
  }
  if (health.authenticated && !options.force) {
    return {
      ok: true,
      steps: [{ step: 'already-connected', ok: true, detail: health.account ?? '' }],
      health,
      message: `Work IQ は既に接続済みです（${health.account ?? ''}）。`,
    };
  }

  if (health.workiq === 'not-installed') {
    steps.push(await installWorkIq());
    resetCommandCache();
  }

  steps.push(await acceptEula());
  steps.push(await disableBrokeredAuth());

  if (options.interactive !== false) {
    steps.push(await login(options.loginTimeoutMs));
  } else {
    steps.push({
      step: 'auth-login',
      ok: false,
      detail: 'ブラウザ認証が必要です。ターミナルで `npx -y @microsoft/workiq auth login` を実行してください。',
    });
  }

  health = await checkHealth();
  const ok = health.authenticated;

  return {
    ok,
    steps,
    health,
    message: ok
      ? `Work IQ に接続しました（${health.account ?? ''}）。`
      : 'Work IQ への接続が完了していません。ターミナルで `npx -y @microsoft/workiq auth login` を実行し、ブラウザでサインインしてください。',
  };
}

async function installWorkIq(): Promise<SetupStep> {
  const npm = process.platform === 'win32' ? 'npm.cmd' : 'npm';
  const { spawn } = await import('node:child_process');
  return await new Promise<SetupStep>((resolve) => {
    const child = spawn(npm, ['install', '-g', '@microsoft/workiq'], { stdio: 'ignore', windowsHide: true });
    const timer = setTimeout(() => child.kill(), 300_000);
    child.on('error', (error) => {
      clearTimeout(timer);
      resolve({ step: 'install-workiq', ok: false, detail: error.message });
    });
    child.on('close', (code) => {
      clearTimeout(timer);
      resolve({
        step: 'install-workiq',
        ok: code === 0,
        detail: code === 0 ? '@microsoft/workiq をインストールしました' : `npm install failed (exit ${code})`,
      });
    });
  });
}

async function acceptEula(): Promise<SetupStep> {
  try {
    const result = await runWorkIq(['accept-eula'], { timeoutMs: 60_000 });
    return { step: 'accept-eula', ok: result.code === 0 };
  } catch (error) {
    return { step: 'accept-eula', ok: false, detail: (error as Error).message };
  }
}

/** Step 1 of the recovery flow: turn off Windows WAM / broker auth. */
async function disableBrokeredAuth(): Promise<SetupStep> {
  try {
    const result = await runWorkIq(['config', 'set', 'disableBrokeredAuth=true'], { timeoutMs: 60_000 });
    return {
      step: 'disable-brokered-auth',
      ok: result.code === 0,
      detail: result.code === 0 ? 'WAM/Broker 認証を無効化しました' : 'config set に失敗しました',
    };
  } catch (error) {
    return { step: 'disable-brokered-auth', ok: false, detail: (error as Error).message };
  }
}

/** Step 2: browser-based sign-in. Never logs the resulting tokens. */
async function login(timeoutMs = 300_000): Promise<SetupStep> {
  try {
    logger.info('starting Work IQ browser sign-in');
    const result = await runWorkIq(['auth', 'login'], { timeoutMs, interactive: true });
    return {
      step: 'auth-login',
      ok: result.code === 0,
      detail: result.code === 0 ? 'ブラウザ認証が完了しました' : 'ブラウザ認証が完了しませんでした',
    };
  } catch (error) {
    return { step: 'auth-login', ok: false, detail: asHey365Error(error).nextStep('ja') };
  }
}

export function workIqCommandInfo(): { command: string; via: string } {
  const resolved = resolveWorkIqCommand();
  return { command: describeCommand(), via: resolved.via };
}
