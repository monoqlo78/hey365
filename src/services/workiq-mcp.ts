import { spawn } from 'node:child_process';
import { createCipheriv, createDecipheriv, createHash, randomBytes } from 'node:crypto';
import { mkdir, readFile, rename, writeFile } from 'node:fs/promises';
import { createServer } from 'node:http';
import { homedir } from 'node:os';
import { dirname, join, resolve } from 'node:path';

import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { auth, UnauthorizedError, type OAuthClientProvider } from '@modelcontextprotocol/sdk/client/auth.js';
import { StreamableHTTPClientTransport, StreamableHTTPError } from '@modelcontextprotocol/sdk/client/streamableHttp.js';
import { OAuthTokensSchema, type OAuthTokens } from '@modelcontextprotocol/sdk/shared/auth.js';
import { CallToolResultSchema } from '@modelcontextprotocol/sdk/types.js';

import { Hey365Error, classifyFailure } from '../utils/errors.js';
import { logger, redact } from '../utils/logger.js';

export const WORKIQ_MCP_URL = 'https://workiq.svc.cloud.microsoft/mcp';
export const WORKIQ_CLIENT_ID = 'ba081686-5d24-4bc6-a0d6-d034ecffed87';
const REDIRECT_URL = 'http://localhost:12798/callback';

export function workIqBackend(): 'mcp' | 'cli' {
  const value = process.env.HEY365_WORKIQ_BACKEND?.trim().toLowerCase() || 'mcp';
  if (value !== 'mcp' && value !== 'cli') {
    throw new Hey365Error('INVALID_INPUT', 'HEY365_WORKIQ_BACKEND must be mcp or cli');
  }
  return value;
}

function tokenPath(): string {
  const account = process.env.HEY365_WORKIQ_ACCOUNT?.trim().toLowerCase() || 'default';
  const key = createHash('sha256').update(account).digest('hex').slice(0, 24);
  return resolve(process.env.HEY365_HOME?.trim() || join(homedir(), '.hey365'), `workiq-${key}.enc`);
}

async function credentialKey(path: string, create: boolean): Promise<Buffer | undefined> {
  const { AsyncEntry } = await import('@napi-rs/keyring');
  const entry = new AsyncEntry('Hey365.WorkIQ', createHash('sha256').update(path).digest('hex'));
  let encoded = await entry.getPassword();
  if (!encoded && create) {
    encoded = randomBytes(32).toString('base64');
    await entry.setPassword(encoded);
  }
  return encoded ? Buffer.from(encoded, 'base64') : undefined;
}

/**
 * Tokens can exceed the Windows Credential Manager size limit. Only the
 * encryption key lives there; the token cache is authenticated ciphertext.
 */
export class WorkIqOAuthProvider implements OAuthClientProvider {
  readonly redirectUrl = REDIRECT_URL;
  readonly clientMetadata = {
    client_name: 'Hey365',
    redirect_uris: [REDIRECT_URL],
    grant_types: ['authorization_code', 'refresh_token'],
    response_types: ['code'],
    token_endpoint_auth_method: 'none',
  };
  private readonly cachePath = tokenPath();
  private verifier?: string;
  private readonly nonce = randomBytes(32).toString('hex');
  authorizationUrl?: URL;

  clientInformation() {
    return { client_id: WORKIQ_CLIENT_ID, token_endpoint_auth_method: 'none' };
  }

  state(): string {
    return this.nonce;
  }

  async tokens(): Promise<OAuthTokens | undefined> {
    let encrypted: Buffer;
    try {
      encrypted = await readFile(this.cachePath);
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code === 'ENOENT') return undefined;
      throw new Hey365Error('WORKIQ_CREDENTIAL_STORE_ERROR', 'Cannot read the WorkIQ token cache');
    }
    try {
      const key = await credentialKey(this.cachePath, false);
      if (!key) throw new Error('Encryption key unavailable');
      const decipher = createDecipheriv('aes-256-gcm', key, encrypted.subarray(0, 12));
      decipher.setAuthTag(encrypted.subarray(12, 28));
      const plaintext = Buffer.concat([decipher.update(encrypted.subarray(28)), decipher.final()]);
      return OAuthTokensSchema.parse(JSON.parse(plaintext.toString('utf8')));
    } catch {
      throw new Hey365Error('WORKIQ_CREDENTIAL_STORE_ERROR', 'Cannot unlock the WorkIQ token cache');
    }
  }

  async saveTokens(tokens: OAuthTokens): Promise<void> {
    try {
      const key = await credentialKey(this.cachePath, true);
      if (!key) throw new Error('Encryption key unavailable');
      const iv = randomBytes(12);
      const cipher = createCipheriv('aes-256-gcm', key, iv);
      const ciphertext = Buffer.concat([cipher.update(JSON.stringify(tokens)), cipher.final()]);
      await mkdir(dirname(this.cachePath), { recursive: true, mode: 0o700 });
      const temporaryPath = `${this.cachePath}.${randomBytes(8).toString('hex')}.tmp`;
      await writeFile(temporaryPath, Buffer.concat([iv, cipher.getAuthTag(), ciphertext]), { mode: 0o600 });
      await rename(temporaryPath, this.cachePath);
    } catch {
      throw new Hey365Error('WORKIQ_CREDENTIAL_STORE_ERROR', 'Cannot securely save WorkIQ credentials');
    }
  }

  redirectToAuthorization(url: URL): void {
    const scopes = new Set((url.searchParams.get('scope') || '').split(' ').filter(Boolean));
    scopes.add('offline_access');
    url.searchParams.set('scope', [...scopes].join(' '));
    const account = process.env.HEY365_WORKIQ_ACCOUNT?.trim();
    if (account) url.searchParams.set('login_hint', account);
    this.authorizationUrl = url;
    // Normal tool calls never open a browser. Only explicit CLI setup does.
    throw new Hey365Error('WORKIQ_NOT_AUTHENTICATED');
  }

  saveCodeVerifier(verifier: string): void {
    this.verifier = verifier;
  }

  codeVerifier(): string {
    if (!this.verifier) throw new Hey365Error('WORKIQ_NOT_AUTHENTICATED', 'No active OAuth PKCE flow');
    return this.verifier;
  }

  async invalidateCredentials(scope: 'all' | 'client' | 'tokens' | 'verifier' | 'discovery'): Promise<void> {
    if (scope === 'all' || scope === 'verifier') this.verifier = undefined;
    if (scope === 'all' || scope === 'tokens') {
      // Keep the OS key and overwrite the cache atomically, not with plaintext.
      await this.saveTokens({ access_token: '', token_type: 'Bearer' });
    }
  }
}

function hostedError(error: unknown): Hey365Error {
  if (error instanceof Hey365Error) return error;
  if (error instanceof UnauthorizedError) return new Hey365Error('WORKIQ_NOT_AUTHENTICATED');
  const message = redact(error instanceof Error ? error.message : String(error)).slice(0, 1000);
  const status = error instanceof StreamableHTTPError ? error.code : undefined;
  return new Hey365Error(classifyFailure(message, status), message);
}

let connection: Promise<Client> | undefined;

async function connectedClient(timeoutMs: number): Promise<Client> {
  if (!connection) {
    connection = (async () => {
      const client = new Client({ name: 'hey365', version: '0.2.0' });
      const transport = new StreamableHTTPClientTransport(new URL(WORKIQ_MCP_URL), {
        authProvider: new WorkIqOAuthProvider(),
        fetch: async (input, init) => {
          const response = await fetch(input, init);
          if (response.status === 403) {
            // Do not let SDK upscoping replay an authorization-denied call.
            const detail = redact(await response.text()).slice(0, 1000);
            throw new Hey365Error(classifyFailure(detail, 403), `HTTP 403: ${detail}`);
          }
          return response;
        },
      });
      try {
        await client.connect(transport, { timeout: timeoutMs });
        return client;
      } catch (error) {
        await transport.close();
        throw hostedError(error);
      }
    })();
    connection.catch(() => { connection = undefined; });
  }
  return await connection;
}

export async function closeWorkIqMcp(): Promise<void> {
  const pending = connection;
  connection = undefined;
  if (pending) {
    const client = await pending.catch(() => undefined);
    await client?.close();
  }
}

interface ToolResult {
  isError?: boolean;
  structuredContent?: Record<string, unknown>;
  content: Array<{ type: string; text?: string }>;
}

export function decodeWorkIqResult(result: ToolResult): unknown {
  const text = result.content.filter((part) => part.type === 'text').map((part) => part.text || '').join('\n');
  if (result.isError) {
    const detail = redact(text || JSON.stringify(result.structuredContent) || 'WorkIQ tool failed').slice(0, 1000);
    throw new Hey365Error(classifyFailure(detail), detail);
  }
  if (result.structuredContent) return result.structuredContent;
  if (!text.trim()) throw new Hey365Error('INTERNAL_ERROR', 'WorkIQ MCP returned no result');
  try {
    return JSON.parse(text);
  } catch {
    return text;
  }
}

export async function callWorkIqMcp(name: string, args: Record<string, unknown>, timeoutMs = 120_000): Promise<unknown> {
  try {
    const client = await connectedClient(timeoutMs);
    const result = await client.callTool({ name, arguments: args }, undefined, { timeout: timeoutMs });
    return decodeWorkIqResult(CallToolResultSchema.parse(result));
  } catch (error) {
    // Never replay a tool call here: a timed-out mutation may have succeeded.
    throw hostedError(error);
  }
}

async function openBrowser(url: URL): Promise<void> {
  const command = process.platform === 'win32' ? 'explorer.exe' : process.platform === 'darwin' ? 'open' : 'xdg-open';
  await new Promise<void>((resolve, reject) => {
    const child = spawn(command, [url.href], { stdio: 'ignore', shell: false });
    child.once('error', () => reject(new Hey365Error('WORKIQ_CONNECTION_ERROR', 'Could not open the sign-in browser')));
    child.once('spawn', () => { child.unref(); resolve(); });
  });
}

/** Explicit browser sign-in. No password, auth code or token enters stdio. */
export async function loginWorkIqMcp(timeoutMs = 300_000): Promise<void> {
  await closeWorkIqMcp();
  const provider = new WorkIqOAuthProvider();
  let acceptCode!: (code: string) => void;
  let rejectCode!: (error: Error) => void;
  const codePromise = new Promise<string>((resolve, reject) => { acceptCode = resolve; rejectCode = reject; });
  // The promise can reject while discovery or browser launch is still running.
  void codePromise.catch(() => undefined);
  const server = createServer((request, response) => {
    response.setHeader('Cache-Control', 'no-store');
    let url: URL;
    try {
      url = new URL(request.url || '/', REDIRECT_URL);
    } catch {
      response.writeHead(400).end('Invalid OAuth callback.');
      return;
    }
    if (request.method !== 'GET' || url.pathname !== '/callback' || url.searchParams.get('state') !== provider.state()) {
      response.writeHead(400).end('Invalid OAuth callback.');
      return;
    }
    const error = url.searchParams.get('error');
    const code = url.searchParams.get('code');
    if (error || !code) {
      response.writeHead(400).end('Sign-in did not complete. Return to Hey365.');
      rejectCode(new Hey365Error(classifyFailure(error || ''), 'WorkIQ sign-in was not approved'));
      return;
    }
    response.writeHead(200, { 'Content-Type': 'text/plain; charset=utf-8' }).end('Return to Hey365 to finish sign-in.');
    acceptCode(code);
  });
  const controller = new AbortController();
  const timer = setTimeout(() => {
    controller.abort();
    rejectCode(new Hey365Error('WORKIQ_CONNECTION_ERROR', 'WorkIQ sign-in timed out'));
  }, timeoutMs);
  const authOptions = {
    serverUrl: WORKIQ_MCP_URL,
    fetchFn: (input: string | URL, init?: RequestInit) => fetch(input, { ...init, signal: controller.signal }),
  };
  try {
    await new Promise<void>((resolve, reject) => {
      server.once('error', () => reject(new Hey365Error('WORKIQ_CONNECTION_ERROR', 'Cannot listen on localhost:12798; close other MCP sign-in flows')));
      server.listen(12798, '127.0.0.1', resolve);
    });
    try {
      if (await auth(provider, authOptions) === 'AUTHORIZED') return;
    } catch (error) {
      if (!provider.authorizationUrl) throw error;
    }
    if (!provider.authorizationUrl) throw new Hey365Error('WORKIQ_NOT_AUTHENTICATED');
    logger.info('opening official WorkIQ OAuth sign-in');
    await openBrowser(provider.authorizationUrl);
    const code = await codePromise;
    if (await auth(provider, { ...authOptions, authorizationCode: code }) !== 'AUTHORIZED') {
      throw new Hey365Error('WORKIQ_NOT_AUTHENTICATED');
    }
  } catch (error) {
    if (controller.signal.aborted) throw new Hey365Error('WORKIQ_CONNECTION_ERROR', 'WorkIQ sign-in timed out');
    throw hostedError(error);
  } finally {
    clearTimeout(timer);
    server.closeAllConnections();
    await new Promise<void>((resolve) => server.close(() => resolve()));
  }
}
