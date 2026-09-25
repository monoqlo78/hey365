import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

import { findMute, listMutes, muteConversation, resetStateCache, unmuteConversation } from '../src/services/store.js';
import { collapseDoubledAnswer, extractJsonDocuments } from '../src/services/workiq.js';
import { redact } from '../src/utils/logger.js';
import { detectLanguage, htmlToText, normalizeSubject, similarity, stripQuotedHistory } from '../src/utils/text.js';
import {
  addBusinessDays,
  buildBusinessDayWindow,
  buildWindow,
  businessDaysBetween,
  formatDateTime,
  formatTime,
  japaneseHolidays,
  localDate,
  parseAsOf,
  resolveWindow,
} from '../src/utils/time.js';
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
  it('counts working time only, stepping over the weekend', () => {
    // Sunday 19:00 JST — Sat/Sun are skipped, so 36h reaches back into Thursday.
    const now = new Date('2026-09-20T10:00:00Z');
    const window = buildWindow(36, now, 'Asia/Tokyo');
    expect(window.startIso).toBe('2026-09-17T03:00:00.000Z');
    expect(window.endIso).toBe('2026-09-20T10:00:00.000Z');
    expect(window.skippedDays).toEqual(['2026-09-20', '2026-09-19']);
  });

  it('steps over Japanese public holidays', () => {
    // Wed 2026-09-23 is 秋分の日, Mon 09-21 敬老の日, Tue 09-22 国民の休日.
    const holidays = japaneseHolidays(2026);
    expect(holidays.has('2026-09-21')).toBe(true);
    expect(holidays.has('2026-09-22')).toBe(true);
    expect(holidays.has('2026-09-23')).toBe(true);
    expect(holidays.has('2026-03-20')).toBe(true);

    const now = new Date('2026-09-21T01:00:00Z'); // Mon 10:00 JST, a holiday
    const window = buildWindow(36, now, 'Asia/Tokyo');
    // Mon holiday + Sun + Sat skipped -> 24h of Friday + 12h of Thursday.
    expect(window.startIso).toBe('2026-09-17T03:00:00.000Z');
    expect(window.skippedDays).toEqual(['2026-09-21', '2026-09-20', '2026-09-19']);
  });

  it('falls back to a plain range when business days are disabled', () => {
    process.env.HEY365_BUSINESS_DAYS = 'off';
    try {
      const window = buildWindow(36, new Date('2026-09-20T10:00:00Z'), 'Asia/Tokyo');
      expect(window.startIso).toBe('2026-09-18T22:00:00.000Z');
      expect(window.skippedDays).toEqual([]);
    } finally {
      delete process.env.HEY365_BUSINESS_DAYS;
    }
  });

  it('falls back to 36 hours for invalid input', () => {
    expect(buildWindow(Number.NaN).hours).toBe(36);
  });

  it('formats in the requested timezone', () => {
    expect(formatTime('2026-09-20T01:32:00Z', 'Asia/Tokyo')).toBe('10:32');
    expect(formatDateTime('2026-09-20T05:00:00Z', 'Asia/Tokyo')).toBe('2026-09-20 14:00');
  });
});

describe('business-day window', () => {
  it('counts the reference day as the first business day', () => {
    // Wed 2026-10-07 -> Wed, Tue, Mon; starts at Monday 00:00 JST.
    const window = buildBusinessDayWindow(3, new Date('2026-10-07T05:00:00Z'), 'Asia/Tokyo');
    expect(window.startIso).toBe('2026-10-04T15:00:00.000Z');
    expect(window.businessDays).toBe(3);
    expect(window.skippedDays).toEqual([]);
  });

  it('reaches into the previous week from a Monday', () => {
    // Mon 2026-10-05 -> Mon, (Sun/Sat skipped), Fri, Thu.
    const window = buildBusinessDayWindow(3, new Date('2026-10-05T05:00:00Z'), 'Asia/Tokyo');
    expect(window.startIso).toBe('2026-09-30T15:00:00.000Z');
    expect(window.skippedDays).toEqual(['2026-10-04', '2026-10-03']);
  });

  it('does not let a holiday reference day consume one of the days', () => {
    // 2026-09-23 is 秋分の日, and 09-21/09-22 are holidays too.
    const window = buildBusinessDayWindow(2, new Date('2026-09-23T05:00:00Z'), 'Asia/Tokyo');
    expect(window.startIso).toBe('2026-09-16T15:00:00.000Z'); // Thu 09-17 00:00 JST
    expect(window.skippedDays).toContain('2026-09-23');
  });

  it('reads a bare date as the whole local day', () => {
    expect(parseAsOf('2026-10-07', 'Asia/Tokyo')?.toISOString()).toBe('2026-10-07T14:59:59.999Z');
    expect(parseAsOf('2026-10-07T09:00:00Z', 'Asia/Tokyo')?.toISOString()).toBe('2026-10-07T09:00:00.000Z');
    expect(parseAsOf('not a date', 'Asia/Tokyo')).toBeUndefined();
  });

  it('ends a dated business-day window at the end of that day', () => {
    const window = resolveWindow({ businessDays: 3, asOf: '2026-10-07' }, 'Asia/Tokyo');
    expect(window.startIso).toBe('2026-10-04T15:00:00.000Z');
    expect(window.endIso).toBe('2026-10-07T14:59:59.999Z');
    expect(window.asOf).toBe('2026-10-07T14:59:59.999Z');
  });

  it('still honours hours when no business-day count is given', () => {
    const window = resolveWindow({ hours: 36, now: new Date('2026-09-20T10:00:00Z') }, 'Asia/Tokyo');
    expect(window.startIso).toBe('2026-09-17T03:00:00.000Z');
    expect(window.businessDays).toBeUndefined();
    expect(window.asOf).toBeUndefined();
  });

  it('rejects an unparseable reference date', () => {
    expect(() => resolveWindow({ businessDays: 3, asOf: 'yesterday' }, 'Asia/Tokyo')).toThrow();
  });
});

describe('business-day arithmetic', () => {
  it('counts calendar days elapsed, not hours', () => {
    // Friday 18:00 JST -> Monday 09:00 JST is one business day, not zero.
    const from = new Date('2026-10-02T09:00:00Z');
    const to = new Date('2026-10-05T00:00:00Z');
    expect(businessDaysBetween(from, to, 'Asia/Tokyo')).toBe(1);
  });

  it('returns zero within the same day', () => {
    const from = new Date('2026-10-07T00:00:00Z');
    const to = new Date('2026-10-07T09:00:00Z');
    expect(businessDaysBetween(from, to, 'Asia/Tokyo')).toBe(0);
  });

  it('skips weekends when adding days', () => {
    // Friday + 1 business day = Monday.
    const monday = addBusinessDays(new Date('2026-10-02T05:00:00Z'), 1, 'Asia/Tokyo');
    expect(localDate(monday, 'Asia/Tokyo')).toBe('2026-10-05');
  });

  it('skips a public holiday when adding days', () => {
    // 2026-09-21 is 敬老の日; from Friday 09-18 the next working day is 09-24.
    const next = addBusinessDays(new Date('2026-09-18T05:00:00Z'), 1, 'Asia/Tokyo');
    expect(localDate(next, 'Asia/Tokyo')).toBe('2026-09-24');
  });

  it('reports the local calendar date across the UTC boundary', () => {
    expect(localDate(new Date('2026-10-06T16:00:00Z'), 'Asia/Tokyo')).toBe('2026-10-07');
    expect(localDate(new Date('2026-10-06T14:00:00Z'), 'Asia/Tokyo')).toBe('2026-10-06');
  });
});

describe('snooze and done', () => {
  const stateFile = join(tmpdir(), `hey365-test-${process.pid}-${Math.random().toString(36).slice(2)}.json`);

  beforeEach(() => {
    process.env.HEY365_STATE_FILE = stateFile;
    resetStateCache();
    rmSync(stateFile, { force: true });
  });

  afterEach(() => {
    rmSync(stateFile, { force: true });
    delete process.env.HEY365_STATE_FILE;
    resetStateCache();
  });

  const base = {
    conversationId: 'AAQk',
    subject: '見積の確認',
    lastMessageTime: '2026-10-05T01:00:00.000Z',
  };

  it('hides a conversation the user marked done', () => {
    muteConversation({ ...base, kind: 'done' });
    expect(findMute(base.conversationId, base.lastMessageTime)).toBeDefined();
  });

  it('shows it again once a newer message arrives', () => {
    muteConversation({ ...base, kind: 'done' });
    expect(findMute(base.conversationId, '2026-10-06T02:00:00.000Z')).toBeUndefined();
  });

  it('keeps hiding it while the thread is unchanged', () => {
    muteConversation({ ...base, kind: 'snoozed', until: '2026-10-06T00:00:00.000Z' });
    const before = new Date('2026-10-05T12:00:00.000Z');
    expect(findMute(base.conversationId, base.lastMessageTime, before)).toBeDefined();
  });

  it('expires a snooze once the deadline passes', () => {
    muteConversation({ ...base, kind: 'snoozed', until: '2026-10-06T00:00:00.000Z' });
    const after = new Date('2026-10-06T01:00:00.000Z');
    expect(findMute(base.conversationId, base.lastMessageTime, after)).toBeUndefined();
    expect(listMutes(after)).toHaveLength(0);
  });

  it('restores an item on request', () => {
    muteConversation({ ...base, kind: 'done' });
    expect(unmuteConversation(base.conversationId)).toBe(true);
    expect(unmuteConversation(base.conversationId)).toBe(false);
    expect(findMute(base.conversationId, base.lastMessageTime)).toBeUndefined();
  });

  it('replaces an existing mute instead of stacking duplicates', () => {
    muteConversation({ ...base, kind: 'snoozed', until: '2026-10-06T00:00:00.000Z' });
    muteConversation({ ...base, kind: 'done' });
    const mutes = listMutes(new Date('2026-10-05T12:00:00.000Z'));
    expect(mutes).toHaveLength(1);
    expect(mutes[0]?.kind).toBe('done');
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
