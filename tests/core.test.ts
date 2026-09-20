import { describe, expect, it } from 'vitest';

import { collapseDoubledAnswer, extractJsonDocuments } from '../src/services/workiq.js';
import { redact } from '../src/utils/logger.js';
import { detectLanguage, htmlToText, normalizeSubject, similarity, stripQuotedHistory } from '../src/utils/text.js';
import { buildWindow, formatDateTime, formatTime } from '../src/utils/time.js';
import { classifyFailure } from '../src/utils/errors.js';
import { mergeJson, mergeToml, renderConfig } from '../src/services/install.js';
import { dedupeActionItems } from '../src/services/dedupe.js';
import type { ActionItem } from '../src/models/action-item.js';

describe('workiq output parsing', () => {
  it('extracts multiple concatenated JSON documents', () => {
    const raw = '{"a":1}\n\n{"b":{"c":"}"}}';
    expect(extractJsonDocuments(raw)).toEqual([{ a: 1 }, { b: { c: '}' } }]);
  });

  it('ignores the localised prefix that precedes action output', () => {
    const raw = 'アクションが完了しました (HTTP 200):\n{"value":[]}';
    expect(extractJsonDocuments(raw)).toEqual([{ value: [] }]);
  });

  it('survives an unparsable chunk', () => {
    expect(extractJsonDocuments('not json at all')).toEqual([]);
  });

  it('collapses the doubled answer Work IQ returns from ask', () => {
    expect(collapseDoubledAnswer('HELLO-PROBEHELLO-PROBE')).toBe('HELLO-PROBE');
    expect(collapseDoubledAnswer('ご確認ください。 ご確認ください。')).toBe('ご確認ください。');
  });

  it('leaves a normal answer untouched', () => {
    const text = 'お世話になっております。会議リンクを本日中にお送りします。';
    expect(collapseDoubledAnswer(text)).toBe(text);
  });
});

describe('secret redaction', () => {
  it('removes bearer tokens and JWTs', () => {
    const line = 'Authorization: Bearer eyJhbGciOiJIUzI1NiJ9.eyJzdWIiOiIxMjM0NTY3ODkwIn0.dBjftJeZ4CVPmB92K';
    const output = redact(line);
    expect(output).not.toContain('eyJhbGciOiJIUzI1NiJ9');
  });

  it('removes structured secret fields', () => {
    const output = redact('{"access_token":"abc123","refresh_token":"def456","user":"ok"}');
    expect(output).not.toContain('abc123');
    expect(output).not.toContain('def456');
    expect(output).toContain('ok');
  });
});

describe('text helpers', () => {
  it('converts HTML to readable text', () => {
    expect(htmlToText('<p>Hello<br/>World</p><style>x{}</style>')).toBe('Hello\nWorld');
  });

  it('removes quoted history', () => {
    const body = 'ご確認ください。よろしくお願いします。\n\n差出人: Someone\n過去のやり取り';
    expect(stripQuotedHistory(body)).toBe('ご確認ください。よろしくお願いします。');
  });

  it('normalises subjects', () => {
    expect(normalizeSubject('RE: Fwd: [External] Azure SQL')).toBe('Azure SQL');
  });

  it('detects the dominant language', () => {
    expect(detectLanguage('この構成で問題ありませんか')).toBe('ja');
    expect(detectLanguage('Could you please confirm the configuration?')).toBe('en');
    expect(detectLanguage('Please confirm the plan. Thanks, 曽我部')).toBe('en');
  });

  it('scores similarity for Japanese without whitespace', () => {
    expect(similarity('Fabric PoC の開始日について', 'Fabric PoC の開始日について確認')).toBeGreaterThan(0.6);
    expect(similarity('請求書の件', 'Azure の障害報告')).toBeLessThan(0.3);
  });
});

describe('time window', () => {
  it('uses now minus N hours, not a calendar day', () => {
    const now = new Date('2026-09-20T10:00:00Z');
    const window = buildWindow(36, now);
    expect(window.startIso).toBe('2026-09-18T22:00:00.000Z');
    expect(window.endIso).toBe('2026-09-20T10:00:00.000Z');
  });

  it('falls back to 36 hours for invalid input', () => {
    expect(buildWindow(Number.NaN).hours).toBe(36);
  });

  it('formats in the requested timezone', () => {
    expect(formatTime('2026-09-20T01:32:00Z', 'Asia/Tokyo')).toBe('10:32');
    expect(formatDateTime('2026-09-20T05:00:00Z', 'Asia/Tokyo')).toBe('2026-09-20 14:00');
  });
});

describe('error classification', () => {
  it('maps common failures', () => {
    expect(classifyFailure('spawn workiq ENOENT')).toBe('WORKIQ_NOT_INSTALLED');
    expect(classifyFailure('AADSTS65001: consent_required')).toBe('WORKIQ_ADMIN_CONSENT_REQUIRED');
    expect(classifyFailure('no cached account found')).toBe('WORKIQ_NOT_AUTHENTICATED');
    expect(classifyFailure('AADSTS700082 token expired')).toBe('WORKIQ_AUTH_EXPIRED');
    expect(classifyFailure('Forbidden', 403)).toBe('WORKIQ_PERMISSION_DENIED');
    expect(classifyFailure('getaddrinfo ENOTFOUND graph.microsoft.com')).toBe('WORKIQ_CONNECTION_ERROR');
  });
});

describe('client config generation', () => {
  const spec = { command: 'node', args: ['/opt/hey365/dist/index.js', 'mcp'] };

  it('renders the VS Code envelope', () => {
    const config = JSON.parse(renderConfig('vscode-servers', spec));
    expect(config.servers.hey365.type).toBe('stdio');
    expect(config.servers.hey365.args).toContain('mcp');
  });

  it('renders TOML for Codex', () => {
    const toml = renderConfig('toml', spec);
    expect(toml).toContain('[mcp_servers.hey365]');
    expect(toml).toContain('args = ["/opt/hey365/dist/index.js", "mcp"]');
  });

  it('preserves other servers when merging JSON', () => {
    const existing = JSON.stringify({ mcpServers: { other: { command: 'x', args: [] } } });
    const merged = JSON.parse(mergeJson(existing, 'mcpServers', 'hey365', spec));
    expect(merged.mcpServers.other).toBeDefined();
    expect(merged.mcpServers.hey365.command).toBe('node');
  });

  it('replaces an existing TOML table instead of duplicating it', () => {
    const existing = '[mcp_servers.hey365]\ncommand = "old"\nargs = []\n\n[mcp_servers.other]\ncommand = "keep"\n';
    const merged = mergeToml(existing, 'hey365', spec);
    expect(merged.match(/\[mcp_servers\.hey365\]/g)).toHaveLength(1);
    expect(merged).toContain('[mcp_servers.other]');
    expect(merged).not.toContain('"old"');
  });
});

describe('de-duplication', () => {
  const base: ActionItem = {
    index: 1,
    id: 'a',
    source: 'outlook',
    conversationId: 'c1',
    subject: 'Fabric PoC の開始日について',
    sender: { name: 'Yamada', address: 'yamada@contoso.com' },
    participants: [{ name: 'Yamada', address: 'yamada@contoso.com' }],
    lastMessageTime: '2026-09-20T01:00:00Z',
    needsReply: true,
    alreadyReplied: false,
    importance: 'medium',
    score: 30,
    reason: 'r',
    reasonEn: 'r',
    signals: [],
    summary: 'Fabric PoC を10月1日に開始してよいか確認したい',
    excerpt: 'Fabric PoC を10月1日に開始してよいか確認したい',
    language: 'ja',
    routing: { kind: 'outlook', messageId: 'm1', conversationId: 'c1', subject: '' },
  };

  it('merges the same matter arriving over two channels', () => {
    const teamsCopy: ActionItem = {
      ...base,
      index: 2,
      id: 'b',
      source: 'teams-chat',
      conversationId: '19:abc@thread.v2',
      subject: 'Re: Fabric PoC の開始日について',
      lastMessageTime: '2026-09-20T02:00:00Z',
      score: 40,
      routing: { kind: 'teams-chat', chatId: '19:abc@thread.v2' },
    };
    const merged = dedupeActionItems([base, teamsCopy]);
    expect(merged).toHaveLength(1);
    expect(merged[0]!.mergedFrom).toHaveLength(1);
    expect(merged[0]!.score).toBe(40);
  });

  it('keeps unrelated items separate', () => {
    const other: ActionItem = {
      ...base,
      index: 2,
      id: 'b',
      conversationId: 'c2',
      subject: '請求書の送付について',
      excerpt: '9月分の請求書を送付します',
      sender: { name: 'Tanaka', address: 'tanaka@contoso.com' },
    };
    expect(dedupeActionItems([base, other])).toHaveLength(2);
  });
});
