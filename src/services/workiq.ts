import { spawn } from 'node:child_process';
import { existsSync, readFileSync, readdirSync } from 'node:fs';
import { homedir } from 'node:os';
import { createRequire } from 'node:module';
import { delimiter, dirname, join } from 'node:path';

import { Hey365Error, classifyFailure } from '../utils/errors.js';
import { logger } from '../utils/logger.js';

export interface RunResult {
  stdout: string;
  stderr: string;
  code: number | null;
  timedOut: boolean;
}

export interface RunOptions {
  timeoutMs?: number;
  /** Set for `auth login`, which must be able to open a browser. */
  interactive?: boolean;
}

/**
 * Under the MCP stdio transport the parent's stdin carries JSON-RPC frames, so
 * a child spawned with `stdio: inherit` would race the transport for those
 * bytes. Interactive spawning is therefore only honoured from the CLI.
 */
type RuntimeMode = 'cli' | 'mcp';
let runtimeMode: RuntimeMode = 'cli';

export function setRuntimeMode(mode: RuntimeMode): void {
  runtimeMode = mode;
}

export function isInteractiveRuntime(): boolean {
  return runtimeMode === 'cli';
}

const DEFAULT_TIMEOUT_MS = Number(process.env.HEY365_WORKIQ_TIMEOUT_MS ?? 120_000);
const WORKIQ_PACKAGE = process.env.HEY365_WORKIQ_PACKAGE ?? '@microsoft/workiq@latest';

interface ResolvedCommand {
  command: string;
  baseArgs: string[];
  via: 'env' | 'script' | 'path' | 'npx';
}

let cachedCommand: ResolvedCommand | undefined;

function binaryNames(): string[] {
  return process.platform === 'win32' ? ['workiq.cmd', 'workiq.exe', 'workiq'] : ['workiq'];
}

function findOnPath(): string | undefined {
  const pathValue = process.env.PATH ?? '';
  for (const dir of pathValue.split(delimiter)) {
    if (!dir) continue;
    for (const name of binaryNames()) {
      const candidate = join(dir, name);
      if (existsSync(candidate)) return candidate;
    }
  }
  return undefined;
}

function scriptFromPackageDir(packageDir: string): string | undefined {
  try {
    const pkg = JSON.parse(readFileSync(join(packageDir, 'package.json'), 'utf8')) as {
      bin?: string | Record<string, string>;
    };
    const rel = typeof pkg.bin === 'string' ? pkg.bin : pkg.bin?.workiq;
    if (!rel) return undefined;
    const script = join(packageDir, rel);
    return existsSync(script) ? script : undefined;
  } catch {
    return undefined;
  }
}

function globalPackageDirs(): string[] {
  const dirs: string[] = [];
  const suffix = join('node_modules', '@microsoft', 'workiq');
  if (process.env.APPDATA) dirs.push(join(process.env.APPDATA, 'npm', suffix));
  dirs.push(join(dirname(process.execPath), suffix));
  dirs.push(join(dirname(process.execPath), '..', 'lib', suffix));
  dirs.push(join('/usr', 'local', 'lib', suffix));
  return dirs;
}

function npxCachePackageDirs(): string[] {
  const cacheRoot = process.env.npm_config_cache
    ? join(process.env.npm_config_cache, '_npx')
    : process.env.LOCALAPPDATA
      ? join(process.env.LOCALAPPDATA, 'npm-cache', '_npx')
      : join(homedir(), '.npm', '_npx');
  if (!existsSync(cacheRoot)) return [];
  try {
    return readdirSync(cacheRoot)
      .map((entry) => join(cacheRoot, entry, 'node_modules', '@microsoft', 'workiq'))
      .filter((dir) => existsSync(dir));
  } catch {
    return [];
  }
}

/**
 * The Work IQ bin is a plain Node script, so spawning it with the current
 * Node binary avoids the `.cmd` shim that Node refuses to spawn on Windows
 * and skips the multi-second npx bootstrap on every call.
 */
function findWorkIqScript(): string | undefined {
  const candidateDirs: string[] = [];
  try {
    const require = createRequire(import.meta.url);
    candidateDirs.push(dirname(require.resolve('@microsoft/workiq/package.json')));
  } catch {
    // Not installed as a dependency of Hey365.
  }
  candidateDirs.push(...globalPackageDirs(), ...npxCachePackageDirs());

  for (const dir of candidateDirs) {
    const script = scriptFromPackageDir(dir);
    if (script) return script;
  }
  return undefined;
}

/**
 * Resolution order: explicit env override → Node script → PATH shim → npx.
 * The result is cached because resolution touches the filesystem.
 */
export function resolveWorkIqCommand(): ResolvedCommand {
  if (cachedCommand) return cachedCommand;

  const override = process.env.HEY365_WORKIQ_COMMAND?.trim();
  if (override) {
    const parts = override.split(/\s+/);
    const [command, ...baseArgs] = parts;
    cachedCommand = { command: command ?? 'workiq', baseArgs, via: 'env' };
    return cachedCommand;
  }

  const script = findWorkIqScript();
  if (script) {
    cachedCommand = { command: process.execPath, baseArgs: [script], via: 'script' };
    return cachedCommand;
  }

  const onPath = findOnPath();
  if (onPath) {
    cachedCommand = { command: onPath, baseArgs: [], via: 'path' };
    return cachedCommand;
  }

  const npx = process.platform === 'win32' ? 'npx.cmd' : 'npx';
  cachedCommand = { command: npx, baseArgs: ['-y', WORKIQ_PACKAGE], via: 'npx' };
  return cachedCommand;
}

export function resetCommandCache(): void {
  cachedCommand = undefined;
}

export function describeCommand(): string {
  const resolved = resolveWorkIqCommand();
  return `${resolved.command} ${resolved.baseArgs.join(' ')}`.trim();
}

interface SpawnTarget {
  command: string;
  args: string[];
  verbatim: boolean;
}

/**
 * Node refuses to spawn `.cmd` / `.bat` shims directly (EINVAL), so they are
 * routed through the command processor with an explicitly quoted command line.
 */
export function toSpawnTarget(command: string, args: string[]): SpawnTarget {
  const isBatch = process.platform === 'win32' && /\.(cmd|bat)$/i.test(command);
  if (!isBatch) return { command, args, verbatim: false };

  const quote = (value: string): string => `"${value.replace(/"/g, '\\"')}"`;
  const commandLine = [command, ...args].map(quote).join(' ');
  return {
    command: process.env.ComSpec ?? 'cmd.exe',
    args: ['/d', '/s', '/c', `"${commandLine}"`],
    verbatim: true,
  };
}

export async function runWorkIq(args: string[], options: RunOptions = {}): Promise<RunResult> {
  const resolved = resolveWorkIqCommand();
  const account = process.env.HEY365_WORKIQ_ACCOUNT?.trim();
  const fullArgs = [...resolved.baseArgs, ...args];
  if (account && !args.includes('--account')) fullArgs.push('--account', account);

  logger.debug('workiq exec', { via: resolved.via, args: args[0] });

  const spawned = toSpawnTarget(resolved.command, fullArgs);
  const inheritStdin = Boolean(options.interactive) && isInteractiveRuntime();

  return await new Promise<RunResult>((resolve, reject) => {
    const child = spawn(spawned.command, spawned.args, {
      stdio: [inheritStdin ? 'inherit' : 'ignore', 'pipe', 'pipe'],
      shell: false,
      windowsHide: true,
      windowsVerbatimArguments: spawned.verbatim,
      env: { ...process.env, NO_COLOR: '1' },
    });

    let stdout = '';
    let stderr = '';
    let timedOut = false;

    const timer = setTimeout(() => {
      timedOut = true;
      child.kill();
    }, options.timeoutMs ?? DEFAULT_TIMEOUT_MS);

    child.stdout?.on('data', (chunk: Buffer) => {
      stdout += chunk.toString('utf8');
    });
    child.stderr?.on('data', (chunk: Buffer) => {
      stderr += chunk.toString('utf8');
    });

    child.on('error', (error: NodeJS.ErrnoException) => {
      clearTimeout(timer);
      if (error.code === 'ENOENT') {
        reject(new Hey365Error('WORKIQ_NOT_INSTALLED', `Cannot launch ${resolved.command}`));
        return;
      }
      reject(new Hey365Error(classifyFailure(error.message), error.message));
    });

    child.on('close', (code) => {
      clearTimeout(timer);
      resolve({ stdout, stderr, code, timedOut });
    });
  });
}

/**
 * The CLI prints a localised confirmation line before the JSON payload for
 * write operations, and emits one JSON document per requested URL for reads.
 * This extracts every top-level JSON value from that mixed output.
 */
export function extractJsonDocuments(raw: string): unknown[] {
  const documents: unknown[] = [];
  let index = 0;

  while (index < raw.length) {
    const objectStart = raw.indexOf('{', index);
    const arrayStart = raw.indexOf('[', index);
    const candidates = [objectStart, arrayStart].filter((position) => position >= 0);
    if (candidates.length === 0) break;
    const start = Math.min(...candidates);
    const open = raw[start] as '{' | '[';
    const close = open === '{' ? '}' : ']';

    let depth = 0;
    let inString = false;
    let escaped = false;
    let end = -1;

    for (let cursor = start; cursor < raw.length; cursor += 1) {
      const char = raw[cursor];
      if (inString) {
        if (escaped) escaped = false;
        else if (char === '\\') escaped = true;
        else if (char === '"') inString = false;
        continue;
      }
      if (char === '"') inString = true;
      else if (char === open) depth += 1;
      else if (char === close) {
        depth -= 1;
        if (depth === 0) {
          end = cursor;
          break;
        }
      }
    }

    if (end < 0) break;
    const slice = raw.slice(start, end + 1);
    try {
      documents.push(JSON.parse(slice));
    } catch {
      logger.debug('skipping unparsable JSON chunk');
    }
    index = end + 1;
  }

  return documents;
}

interface WrappedResult {
  data: unknown;
  statusCode?: number;
  error?: { error?: { code?: string; message?: string } };
}

function isWrapped(value: unknown): value is { results: WrappedResult[] } {
  return typeof value === 'object' && value !== null && Array.isArray((value as { results?: unknown }).results);
}

/** Normalises one CLI JSON document into either a payload or a Hey365Error. */
function unwrap(document: unknown): { ok: true; data: unknown } | { ok: false; error: Hey365Error } {
  if (isWrapped(document)) {
    const first = document.results[0];
    if (!first) return { ok: true, data: undefined };
    if (first.error) {
      const message = first.error.error?.message ?? 'Microsoft Graph request failed';
      const code = classifyFailure(`${first.error.error?.code ?? ''} ${message}`, first.statusCode);
      return { ok: false, error: new Hey365Error(code, message) };
    }
    return { ok: true, data: first.data };
  }
  return { ok: true, data: document };
}

export interface GraphCallOutcome<T> {
  ok: boolean;
  data?: T;
  error?: Hey365Error;
}

function assertCliSuccess(result: RunResult, action: string): void {
  if (result.timedOut) {
    throw new Hey365Error('WORKIQ_CONNECTION_ERROR', `Work IQ timed out while running ${action}`);
  }
  if (result.code === 0) return;
  const combined = `${result.stderr}\n${result.stdout}`;
  // A non-zero exit still carries a JSON error body for Graph failures, which
  // `unwrap` turns into a precise code; only bail out when there is no JSON.
  if (extractJsonDocuments(combined).length > 0) return;
  throw new Hey365Error(classifyFailure(combined), combined.trim().slice(0, 500) || `Work IQ ${action} failed`);
}

/** Fetches one Graph URL and throws on failure. */
export async function fetchOne<T = unknown>(url: string, options?: RunOptions): Promise<T> {
  const [outcome] = await fetchMany<T>([url], options);
  if (!outcome) throw new Hey365Error('INTERNAL_ERROR', 'Work IQ returned no response');
  if (!outcome.ok || outcome.data === undefined) {
    throw outcome.error ?? new Hey365Error('INTERNAL_ERROR', 'Work IQ returned no data');
  }
  return outcome.data;
}

/** Fetches one Graph URL, returning `undefined` instead of throwing. */
export async function tryFetch<T = unknown>(url: string, options?: RunOptions): Promise<T | undefined> {
  try {
    return await fetchOne<T>(url, options);
  } catch (error) {
    logger.debug('fetch failed', { url: url.split('?')[0], error: (error as Error).message });
    return undefined;
  }
}

/**
 * Batches multiple Graph URLs into a single CLI invocation. Process startup
 * dominates latency, so batching is the main throughput lever.
 */
export async function fetchMany<T = unknown>(urls: string[], options?: RunOptions): Promise<Array<GraphCallOutcome<T>>> {
  if (urls.length === 0) return [];
  const result = await runWorkIq(['fetch', '--urls', ...urls], options);
  assertCliSuccess(result, 'fetch');

  const documents = extractJsonDocuments(`${result.stdout}\n${result.stderr}`);
  const outcomes = documents.map((document) => {
    const unwrapped = unwrap(document);
    return unwrapped.ok
      ? ({ ok: true, data: unwrapped.data as T } satisfies GraphCallOutcome<T>)
      : ({ ok: false, error: unwrapped.error } satisfies GraphCallOutcome<T>);
  });

  // Keep positional alignment with the requested URLs even if the CLI dropped
  // a response for one of them.
  while (outcomes.length < urls.length) {
    outcomes.push({ ok: false, error: new Hey365Error('INTERNAL_ERROR', 'Missing response from Work IQ') });
  }
  return outcomes;
}

export async function callFunction<T = unknown>(url: string, options?: RunOptions): Promise<T> {
  const result = await runWorkIq(['call-function', '--url', url], options);
  assertCliSuccess(result, 'call-function');
  const [document] = extractJsonDocuments(`${result.stdout}\n${result.stderr}`);
  const unwrapped = unwrap(document);
  if (!unwrapped.ok) throw unwrapped.error;
  return unwrapped.data as T;
}

export async function doAction<T = unknown>(url: string, body?: unknown, options?: RunOptions): Promise<T> {
  const args = ['do-action', '--url', url];
  if (body !== undefined) args.push('--body', JSON.stringify(body));
  const result = await runWorkIq(args, options);
  assertCliSuccess(result, 'do-action');
  const [document] = extractJsonDocuments(`${result.stdout}\n${result.stderr}`);
  if (document === undefined) return undefined as T;
  const unwrapped = unwrap(document);
  if (!unwrapped.ok) throw unwrapped.error;
  return unwrapped.data as T;
}

export async function createEntity<T = unknown>(url: string, body: unknown, options?: RunOptions): Promise<T> {
  const result = await runWorkIq(['create', '--url', url, '--body', JSON.stringify(body)], options);
  assertCliSuccess(result, 'create');
  const [document] = extractJsonDocuments(`${result.stdout}\n${result.stderr}`);
  if (document === undefined) return undefined as T;
  const unwrapped = unwrap(document);
  if (!unwrapped.ok) throw unwrapped.error;
  return unwrapped.data as T;
}

export interface AskResult {
  text: string;
  conversationId?: string;
}

/**
 * Asks Microsoft 365 Copilot through Work IQ. Used for reply drafting and
 * meeting summarisation so Hey365 needs no extra model credentials.
 */
export async function ask(question: string, options?: RunOptions & { conversationId?: string }): Promise<AskResult> {
  const args = ['ask', '--question', question, '--json'];
  if (options?.conversationId) args.push('--conversation-id', options.conversationId);
  const result = await runWorkIq(args, { timeoutMs: options?.timeoutMs ?? 180_000 });
  assertCliSuccess(result, 'ask');

  const raw = `${result.stdout}\n${result.stderr}`;
  for (const document of extractJsonDocuments(raw)) {
    const text = extractAskText(document);
    if (text) {
      const conversationId =
        typeof document === 'object' && document !== null
          ? ((document as Record<string, unknown>).conversationId as string | undefined)
          : undefined;
      return { text: collapseDoubledAnswer(text), conversationId };
    }
  }

  const fallback = result.stdout.trim();
  if (!fallback) throw new Hey365Error('INTERNAL_ERROR', 'Work IQ ask returned an empty response');
  return { text: collapseDoubledAnswer(fallback) };
}

/**
 * Work IQ concatenates the streamed answer with the final answer, so `ask`
 * responses arrive doubled ("HELLO-PROBEHELLO-PROBE"). Collapse the repetition
 * instead of showing the user the same sentence twice.
 */
export function collapseDoubledAnswer(text: string): string {
  const trimmed = text.trim();
  if (trimmed.length < 8) return trimmed;

  for (const split of [trimmed.length / 2, (trimmed.length - 1) / 2, (trimmed.length + 1) / 2]) {
    if (!Number.isInteger(split)) continue;
    const head = trimmed.slice(0, split).trim();
    const tail = trimmed.slice(split).trim();
    if (head.length >= 4 && head === tail) return head;
  }
  return trimmed;
}

/** The `ask --json` envelope has changed shape across CLI versions. */
function extractAskText(document: unknown): string | undefined {
  if (typeof document === 'string') return document;
  if (typeof document !== 'object' || document === null) return undefined;
  const record = document as Record<string, unknown>;

  for (const key of ['answer', 'text', 'response', 'content', 'message', 'result', 'output']) {
    const value = record[key];
    if (typeof value === 'string' && value.trim()) return value.trim();
    if (value && typeof value === 'object') {
      const nested = extractAskText(value);
      if (nested) return nested;
    }
  }

  for (const key of ['messages', 'parts', 'choices', 'value']) {
    const value = record[key];
    if (Array.isArray(value)) {
      const texts = value
        .map((entry) => extractAskText(entry))
        .filter((entry): entry is string => Boolean(entry && entry.trim()));
      if (texts.length > 0) return texts.join('\n').trim();
    }
  }

  return undefined;
}
