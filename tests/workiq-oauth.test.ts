import { mkdtemp, readdir, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { auth } from '@modelcontextprotocol/sdk/client/auth.js';

const store = vi.hoisted(() => ({ passwords: new Map<string, string>(), fail: false }));
vi.mock('@napi-rs/keyring', () => ({
  AsyncEntry: class {
    constructor(private service: string, private account: string) {}
    async getPassword() {
      if (store.fail) throw new Error('Store locked');
      return store.passwords.get(`${this.service}:${this.account}`) ?? null;
    }
    async setPassword(value: string) {
      if (store.fail) throw new Error('Store locked');
      store.passwords.set(`${this.service}:${this.account}`, value);
    }
  },
}));

const { WorkIqOAuthProvider, decodeWorkIqResult, WORKIQ_CLIENT_ID } = await import('../src/services/workiq-mcp.js');
let home: string;

describe('WorkIQ OAuth credentials', () => {
  beforeEach(async () => {
    home = await mkdtemp(join(tmpdir(), 'hey365-workiq-test-'));
    vi.stubEnv('HEY365_HOME', home);
    vi.stubEnv('HEY365_WORKIQ_ACCOUNT', '');
    store.passwords.clear();
    store.fail = false;
  });
  afterEach(async () => {
    await rm(home, { recursive: true, force: true });
    vi.unstubAllEnvs();
  });

  it('uses the official public OAuth client, a loopback redirect and independent state', () => {
    const first = new WorkIqOAuthProvider();
    const second = new WorkIqOAuthProvider();
    expect(first.clientInformation()).toEqual({ client_id: WORKIQ_CLIENT_ID, token_endpoint_auth_method: 'none' });
    expect(first.redirectUrl).toBe('http://localhost:12798/callback');
    expect(first.state()).not.toBe(second.state());
    expect(() => first.codeVerifier()).toThrow('No active OAuth PKCE flow');
    first.saveCodeVerifier('proof-key');
    expect(first.codeVerifier()).toBe('proof-key');
  });

  it('stores tokens as authenticated ciphertext and only the small key in the OS store', async () => {
    const tokens = { access_token: 'access-token-for-test', refresh_token: 'refresh-token-for-test', token_type: 'Bearer' };
    const provider = new WorkIqOAuthProvider();
    expect(await provider.tokens()).toBeUndefined();
    await provider.saveTokens(tokens);
    const [file] = await readdir(home);
    const raw = await readFile(join(home, file!));
    expect(raw.toString()).not.toContain(tokens.access_token);
    expect(raw.toString()).not.toContain(tokens.refresh_token);
    expect([...store.passwords.values()][0]).toHaveLength(44);
    expect(await new WorkIqOAuthProvider().tokens()).toEqual(tokens);
  });

  it('separates account caches instead of borrowing the CLI or Copilot cache', async () => {
    await new WorkIqOAuthProvider().saveTokens({ access_token: 'one', token_type: 'Bearer' });
    vi.stubEnv('HEY365_WORKIQ_ACCOUNT', 'another@example.com');
    expect(await new WorkIqOAuthProvider().tokens()).toBeUndefined();
  });

  it('fails explicitly when the secure store is unavailable', async () => {
    store.fail = true;
    await expect(new WorkIqOAuthProvider().saveTokens({ access_token: 'test', token_type: 'Bearer' }))
      .rejects.toMatchObject({ code: 'WORKIQ_CREDENTIAL_STORE_ERROR' });
    expect(await readdir(home)).toEqual([]);
  });

  it('does not silently ignore a missing encryption key for an existing cache', async () => {
    const provider = new WorkIqOAuthProvider();
    await provider.saveTokens({ access_token: 'test', token_type: 'Bearer' });
    store.passwords.clear();
    await expect(provider.tokens()).rejects.toMatchObject({ code: 'WORKIQ_CREDENTIAL_STORE_ERROR' });
  });

  it('stops normal calls before a browser redirect and requests refresh-token support', () => {
    const provider = new WorkIqOAuthProvider();
    const url = new URL('https://login.microsoftonline.com/organizations/oauth2/v2.0/authorize?scope=WorkIQAgent.Ask');
    expect(() => provider.redirectToAuthorization(url)).toThrow();
    expect(provider.authorizationUrl?.searchParams.get('scope')).toBe('WorkIQAgent.Ask offline_access');
  });

  it('performs SDK discovery, PKCE exchange and refresh using the official client', async () => {
    const provider = new WorkIqOAuthProvider();
    const tokenBodies: URLSearchParams[] = [];
    const fakeFetch: typeof fetch = async (input, init) => {
      const url = String(input);
      if (url.includes('oauth-protected-resource')) {
        return Response.json({
          resource: 'https://workiq.svc.cloud.microsoft/mcp',
          authorization_servers: ['https://login.microsoftonline.com/organizations/v2.0'],
          scopes_supported: ['fdcc1f02-fc51-4226-8753-f668596af7f7/WorkIQAgent.Ask'],
        });
      }
      if (url.includes('.well-known')) {
        return Response.json({
          issuer: 'https://login.microsoftonline.com/organizations/v2.0',
          authorization_endpoint: 'https://login.microsoftonline.com/organizations/oauth2/v2.0/authorize',
          token_endpoint: 'https://login.microsoftonline.com/organizations/oauth2/v2.0/token',
          response_types_supported: ['code'],
          code_challenge_methods_supported: ['S256'],
          token_endpoint_auth_methods_supported: ['none'],
        });
      }
      tokenBodies.push(new URLSearchParams(init?.body as URLSearchParams));
      return Response.json({ access_token: 'test-access', refresh_token: 'test-refresh', token_type: 'Bearer' });
    };
    const options = { serverUrl: 'https://workiq.svc.cloud.microsoft/mcp', fetchFn: fakeFetch };
    await expect(auth(provider, options)).rejects.toMatchObject({ code: 'WORKIQ_NOT_AUTHENTICATED' });
    expect(provider.authorizationUrl?.searchParams.get('client_id')).toBe(WORKIQ_CLIENT_ID);
    expect(provider.authorizationUrl?.searchParams.get('code_challenge_method')).toBe('S256');
    expect(provider.authorizationUrl?.searchParams.get('state')).toBe(provider.state());
    expect(await auth(provider, { ...options, authorizationCode: 'test-code' })).toBe('AUTHORIZED');
    expect(tokenBodies[0]?.get('code_verifier')).toBe(provider.codeVerifier());
    expect(tokenBodies[0]?.get('client_id')).toBe(WORKIQ_CLIENT_ID);
    expect(tokenBodies[0]?.get('client_secret')).toBeNull();
    expect(await auth(new WorkIqOAuthProvider(), options)).toBe('AUTHORIZED');
    expect(tokenBodies[1]?.get('grant_type')).toBe('refresh_token');
    expect(tokenBodies[1]?.get('refresh_token')).toBe('test-refresh');
  });
});

describe('MCP tool result decoding', () => {
  it('prefers structured content, otherwise parses text JSON or returns an answer', () => {
    expect(decodeWorkIqResult({ content: [], structuredContent: { results: [] } })).toEqual({ results: [] });
    expect(decodeWorkIqResult({ content: [{ type: 'text', text: '{"results":[]}' }] })).toEqual({ results: [] });
    expect(decodeWorkIqResult({ content: [{ type: 'text', text: 'Answer' }] })).toBe('Answer');
  });

  it('preserves a tool-level 403 and never treats absent content as success', () => {
    expect(() => decodeWorkIqResult({ isError: true, content: [{ type: 'text', text: 'HTTP 403 Forbidden' }] }))
      .toThrow(expect.objectContaining({ code: 'WORKIQ_PERMISSION_DENIED' }));
    expect(() => decodeWorkIqResult({ content: [] })).toThrow('no result');
  });
});
