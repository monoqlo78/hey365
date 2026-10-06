import { Hey365Error, asHey365Error, classifyFailure } from '../utils/errors.js';
import { logger } from '../utils/logger.js';
import { timezone } from '../utils/time.js';
import { describeCommand, resetCommandCache, resolveWorkIqCommand, runWorkIq, tryFetch } from './workiq.js';
import { rememberAccount } from './store.js';
import type { GraphUser } from './graph.js';

export interface HealthReport {
  hey365: 'ok' | 'degraded';
  workiq: 'connected' | 'unauthenticated' | 'not-installed' | 'error';
  authenticated: boolean;
  readAccess: boolean;
  writeAccess: boolean;
  timezone: string;
  account?: string;
  workiqCommand: string;
  workiqVersion?: string;
  surfaces?: { outlook: boolean; teams: boolean; calendar: boolean };
  error?: { code: string; nextStep: string };
}

/** Probes Work IQ and reports what Hey365 can currently do (spec §20). */
export async function checkHealth(options: { deep?: boolean } = {}): Promise<HealthReport> {
  const report: HealthReport = {
    hey365: 'ok',
    workiq: 'error',
    authenticated: false,
    readAccess: false,
    writeAccess: false,
    timezone: timezone(),
    workiqCommand: describeCommand(),
  };

  try {
    const version = await runWorkIq(['version'], { timeoutMs: 60_000 });
    if (version.code === 0) {
      const match = /\d+\.\d+\.\d+(?:[-+][\w.]+)?/.exec(version.stdout);
      if (match) report.workiqVersion = match[0];
    }
  } catch (error) {
    const hey = asHey365Error(error);
    if (hey.code === 'WORKIQ_NOT_INSTALLED') {
      report.workiq = 'not-installed';
      report.hey365 = 'degraded';
      report.error = { code: hey.code, nextStep: hey.nextStep('ja') };
      return report;
    }
  }

  const me = await tryFetch<GraphUser>('/me?$select=id,displayName,mail,userPrincipalName');
  if (!me?.id) {
    const probe = await runWorkIq(['fetch', '--urls', '/me'], { timeoutMs: 60_000 }).catch(() => undefined);
    const raw = `${probe?.stderr ?? ''}${probe?.stdout ?? ''}`;
    const code = classifyFailure(raw);
    report.workiq = code === 'WORKIQ_NOT_INSTALLED' ? 'not-installed' : 'unauthenticated';
    report.hey365 = 'degraded';
    const error = new Hey365Error(code === 'INTERNAL_ERROR' ? 'WORKIQ_NOT_AUTHENTICATED' : code);
    report.error = { code: error.code, nextStep: error.nextStep('ja') };
    return report;
  }

  report.workiq = 'connected';
  report.authenticated = true;
  report.readAccess = true;
  report.account = me.mail ?? me.userPrincipalName ?? '';
  rememberAccount(report.account);

  if (options.deep) {
    const [mail, chats, events] = await Promise.all([
      tryFetch('/me/messages?$top=1&$select=id'),
      tryFetch('/me/chats?$top=1'),
      tryFetch('/me/events?$top=1&$select=id'),
    ]);
    report.surfaces = { outlook: Boolean(mail), teams: Boolean(chats), calendar: Boolean(events) };
  }

  // Write capability is inferred from mail folder permissions rather than by
  // sending anything. Hey365 never sends without an explicit user instruction.
  const drafts = await tryFetch('/me/mailFolders/Drafts?$select=id');
  report.writeAccess = Boolean(drafts);
  if (!report.writeAccess) {
    const error = new Hey365Error('WORKIQ_WRITE_DISABLED');
    report.error = { code: error.code, nextStep: error.nextStep('ja') };
  }

  return report;
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
 * Installs / repairs the Work IQ connection (spec §3).
 *
 * On Windows the WAM broker frequently fails inside non-interactive hosts, so
 * the recovery path disables brokered auth before falling back to the browser
 * flow. Nothing here ever logs tokens.
 */
export async function runSetup(options: { interactive?: boolean; force?: boolean } = {}): Promise<SetupResult> {
  const steps: SetupStep[] = [];

  let health = await checkHealth();
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
    steps.push(await login());
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
async function login(): Promise<SetupStep> {
  try {
    logger.info('starting Work IQ browser sign-in');
    const result = await runWorkIq(['auth', 'login'], { timeoutMs: 300_000, interactive: true });
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
