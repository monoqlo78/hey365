import { execFile } from 'node:child_process';
import { homedir, platform } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';
import { fileURLToPath } from 'node:url';

import { Hey365Error } from '../utils/errors.js';
import { logger } from '../utils/logger.js';

const run = promisify(execFile);

/** Comment marker that identifies the cron lines Hey365 owns. */
const CRON_MARKER = '# hey365';
const TASK_PREFIX = 'Hey365';

export type ScheduleJob = 'digest' | 'triage';

export interface ScheduleOptions {
  job?: ScheduleJob;
  /** Local time of day as `HH:MM`. */
  at?: string;
  /** Skip weekends and holidays by having the job exit early on those days. */
  weekdaysOnly?: boolean;
  /** File the run writes its output to. Defaults to ~/.hey365/digest-latest.txt */
  out?: string;
  /** When false the schedule is actually written. Defaults to true. */
  apply?: boolean;
}

export interface SchedulePlan {
  platform: 'windows' | 'unix';
  job: ScheduleJob;
  at: string;
  out: string;
  /** The exact command line the scheduler will run. */
  command: string;
  /** The schtasks argv or the crontab line, for display. */
  entry: string;
  applied: boolean;
}

export interface ScheduleEntry {
  name: string;
  detail: string;
}

function entryPoint(): string {
  // dist/services/schedule.js -> dist/index.js
  return join(fileURLToPath(new URL('..', import.meta.url)), 'index.js');
}

function defaultOut(job: ScheduleJob): string {
  return join(homedir(), '.hey365', `${job}-latest.txt`);
}

function normalizeTime(value: string | undefined): string {
  const at = (value ?? '08:30').trim();
  if (!/^\d{1,2}:\d{2}$/.test(at)) throw new Hey365Error('INVALID_INPUT', undefined, `at=${at}`);
  const [hour, minute] = at.split(':').map(Number) as [number, number];
  if (hour > 23 || minute > 59) throw new Hey365Error('INVALID_INPUT', undefined, `at=${at}`);
  return `${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}`;
}

function isWindows(): boolean {
  return platform() === 'win32';
}

function taskName(job: ScheduleJob): string {
  return `${TASK_PREFIX}${job.charAt(0).toUpperCase()}${job.slice(1)}`;
}

/**
 * Builds the scheduler entry for a daily run. Nothing is written unless
 * `apply` is true, because registering a task changes machine state and the
 * user should see the exact command first.
 */
export async function installSchedule(options: ScheduleOptions = {}): Promise<SchedulePlan> {
  const job = options.job ?? 'digest';
  const at = normalizeTime(options.at);
  const out = options.out ?? defaultOut(job);
  const apply = options.apply === true;

  const args = [entryPoint(), job, '--out', out];
  const command = [process.execPath, ...args].map(quote).join(' ');

  if (isWindows()) {
    // MON-FRI still fires on public holidays; the job itself is cheap and the
    // digest simply reports an empty day, so this stays a scheduler-level hint.
    const schedule = options.weekdaysOnly === false ? ['/SC', 'DAILY'] : ['/SC', 'WEEKLY', '/D', 'MON,TUE,WED,THU,FRI'];
    const argv = ['/Create', '/TN', taskName(job), ...schedule, '/ST', at, '/TR', command, '/F'];
    const entry = `schtasks ${argv.map(quote).join(' ')}`;

    if (apply) {
      await run('schtasks', argv, { windowsHide: true }).catch((error: unknown) => {
        throw new Hey365Error('SCHEDULE_FAILED', (error as Error).message, entry);
      });
      logger.info('schedule installed', { job, at });
    }
    return { platform: 'windows', job, at, out, command, entry, applied: apply };
  }

  const [hour, minute] = at.split(':') as [string, string];
  const days = options.weekdaysOnly === false ? '*' : '1-5';
  const entry = `${Number(minute)} ${Number(hour)} * * ${days} ${command} ${CRON_MARKER}-${job}`;

  if (apply) {
    const existing = await readCrontab();
    const kept = existing.filter((line) => !line.includes(`${CRON_MARKER}-${job}`));
    await writeCrontab([...kept, entry]);
    logger.info('schedule installed', { job, at });
  }
  return { platform: 'unix', job, at, out, command, entry, applied: apply };
}

/** Removes a previously registered schedule. */
export async function removeSchedule(job: ScheduleJob = 'digest', apply = false): Promise<SchedulePlan['entry']> {
  if (isWindows()) {
    const argv = ['/Delete', '/TN', taskName(job), '/F'];
    const entry = `schtasks ${argv.map(quote).join(' ')}`;
    if (apply) {
      await run('schtasks', argv, { windowsHide: true }).catch((error: unknown) => {
        throw new Hey365Error('SCHEDULE_FAILED', (error as Error).message, entry);
      });
    }
    return entry;
  }

  const entry = `crontab: ${CRON_MARKER}-${job} の行を削除`;
  if (apply) {
    const existing = await readCrontab();
    await writeCrontab(existing.filter((line) => !line.includes(`${CRON_MARKER}-${job}`)));
  }
  return entry;
}

/** Lists the schedules Hey365 owns. Returns an empty list when none exist. */
export async function listSchedules(): Promise<ScheduleEntry[]> {
  if (isWindows()) {
    const result = await run('schtasks', ['/Query', '/FO', 'LIST'], { windowsHide: true }).catch(() => undefined);
    if (!result) return [];
    const entries: ScheduleEntry[] = [];
    let current: string | undefined;
    for (const line of result.stdout.split(/\r?\n/)) {
      const name = /^(?:TaskName|タスク名):\s*(.+)$/.exec(line)?.[1]?.trim();
      if (name) {
        current = name.includes(TASK_PREFIX) ? name : undefined;
        continue;
      }
      const next = /^(?:Next Run Time|次回の実行時刻):\s*(.+)$/.exec(line)?.[1]?.trim();
      if (current && next) {
        entries.push({ name: current, detail: `次回 ${next}` });
        current = undefined;
      }
    }
    return entries;
  }

  const lines = await readCrontab();
  return lines
    .filter((line) => line.includes(CRON_MARKER))
    .map((line) => ({ name: line.slice(line.indexOf(CRON_MARKER) + 2), detail: line.split(' ').slice(0, 5).join(' ') }));
}

async function readCrontab(): Promise<string[]> {
  const result = await run('crontab', ['-l']).catch(() => undefined);
  if (!result) return [];
  return result.stdout.split(/\r?\n/).filter((line) => line.trim().length > 0);
}

async function writeCrontab(lines: string[]): Promise<void> {
  await new Promise<void>((resolve, reject) => {
    const child = execFile('crontab', ['-'], (error) => (error ? reject(error) : resolve()));
    child.stdin?.end(`${lines.join('\n')}\n`);
  });
}

function quote(value: string): string {
  if (!/[\s"]/.test(value)) return value;
  // schtasks receives /TR as one argument whose own value is already quoted, so
  // the printed form has to escape those inner quotes to stay copy-pastable.
  return `"${value.replace(/"/g, '\\"')}"`;
}
