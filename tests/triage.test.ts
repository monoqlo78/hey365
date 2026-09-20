import { describe, expect, it } from 'vitest';

import { buildThreads, hasAnsweredLatestInbound, unansweredInbound } from '../src/services/conversation.js';
import { triageThread } from '../src/services/triage.js';
import type { NormalizedMessage } from '../src/models/conversation.js';
import type { Identity } from '../src/services/graph.js';

const ME: Identity = { id: 'me-id', name: 'Masaaki Sogabe', addresses: ['me@contoso.com'] };

function message(overrides: Partial<NormalizedMessage> & { id: string; createdDateTime: string }): NormalizedMessage {
  return {
    conversationId: 'conv-1',
    source: 'outlook',
    from: { name: 'Yamada', address: 'yamada@contoso.com' },
    to: [{ name: 'Masaaki Sogabe', address: 'me@contoso.com' }],
    cc: [],
    subject: 'Azure SQL の構成について',
    body: '',
    isFromMe: false,
    isCcOnlyForMe: false,
    mentionsMe: false,
    routing: { kind: 'outlook', messageId: overrides.id, conversationId: 'conv-1', subject: '' },
    ...overrides,
  };
}

function fromMe(overrides: Partial<NormalizedMessage> & { id: string; createdDateTime: string }): NormalizedMessage {
  return message({
    ...overrides,
    from: { name: 'Masaaki Sogabe', address: 'me@contoso.com' },
    to: [{ name: 'Yamada', address: 'yamada@contoso.com' }],
    isFromMe: true,
  });
}

describe('thread state', () => {
  it('treats a thread as answered when my reply is the newest message', () => {
    const [thread] = buildThreads([
      message({ id: 'm1', createdDateTime: '2026-09-19T01:00:00Z', body: '確認お願いします。' }),
      fromMe({ id: 'm2', createdDateTime: '2026-09-19T02:00:00Z', body: '確認しました。問題ありません。' }),
    ]);
    expect(hasAnsweredLatestInbound(thread!)).toBe(true);
    expect(triageThread(thread!, ME).needsReply).toBe(false);
  });

  it('requires a reply when they followed up after my last message', () => {
    const [thread] = buildThreads([
      fromMe({ id: 'm1', createdDateTime: '2026-09-19T01:00:00Z', body: '確認します。' }),
      message({ id: 'm2', createdDateTime: '2026-09-19T05:00:00Z', body: 'その後どうでしょうか？' }),
    ]);
    expect(hasAnsweredLatestInbound(thread!)).toBe(false);
    const decision = triageThread(thread!, ME);
    expect(decision.needsReply).toBe(true);
    expect(decision.signals.some((signal) => signal.code === 'follow_up')).toBe(true);
  });

  it('returns every inbound message that arrived after my last reply', () => {
    const [thread] = buildThreads([
      message({ id: 'm1', createdDateTime: '2026-09-19T01:00:00Z', body: 'ping 1' }),
      fromMe({ id: 'm2', createdDateTime: '2026-09-19T02:00:00Z', body: 'ok' }),
      message({ id: 'm3', createdDateTime: '2026-09-19T03:00:00Z', body: 'ping 2' }),
      message({ id: 'm4', createdDateTime: '2026-09-19T04:00:00Z', body: 'ping 3' }),
    ]);
    expect(unansweredInbound(thread!).map((entry) => entry.id)).toEqual(['m3', 'm4']);
  });
});

describe('triage signals', () => {
  it('flags an approval request as high importance', () => {
    const [thread] = buildThreads([
      message({
        id: 'm1',
        createdDateTime: '2026-09-19T06:00:00Z',
        body: 'この構成で進めてよいでしょうか？本日中にご承認をお願いします。',
      }),
    ]);
    const decision = triageThread(thread!, ME);
    expect(decision.needsReply).toBe(true);
    expect(decision.importance).toBe('high');
    expect(decision.deadlineText).toBeDefined();
  });

  it('detects an English direct question', () => {
    const [thread] = buildThreads([
      message({
        id: 'm1',
        createdDateTime: '2026-09-19T06:00:00Z',
        subject: 'Workspace creation',
        body: 'Could you confirm whether we can start the PoC on Oct 1?',
      }),
    ]);
    const decision = triageThread(thread!, ME);
    expect(decision.needsReply).toBe(true);
  });

  it('excludes automated notifications', () => {
    const [thread] = buildThreads([
      message({
        id: 'm1',
        createdDateTime: '2026-09-19T06:00:00Z',
        from: { name: 'Azure DevOps', address: 'noreply@azure.com' },
        body: 'Build failed. Please check the pipeline. ご確認ください。',
      }),
    ]);
    const decision = triageThread(thread!, ME);
    expect(decision.needsReply).toBe(false);
    expect(decision.excludeReason).toBe('automated');
  });

  it('excludes a pure FYI broadcast where I am only on CC', () => {
    const [thread] = buildThreads([
      message({
        id: 'm1',
        createdDateTime: '2026-09-19T06:00:00Z',
        to: [{ name: 'Someone Else', address: 'other@contoso.com' }],
        cc: [{ name: 'Masaaki Sogabe', address: 'me@contoso.com' }],
        isCcOnlyForMe: true,
        body: 'FYI',
      }),
    ]);
    expect(triageThread(thread!, ME).needsReply).toBe(false);
  });

  it('excludes a closed conversation', () => {
    const [thread] = buildThreads([
      fromMe({ id: 'm1', createdDateTime: '2026-09-19T01:00:00Z', body: '対応しました。' }),
      message({ id: 'm2', createdDateTime: '2026-09-19T02:00:00Z', body: 'ありがとうございます。助かりました。' }),
    ]);
    const decision = triageThread(thread!, ME);
    expect(decision.needsReply).toBe(false);
    expect(decision.excludeReason).toBe('completed');
  });

  it('flags a request even when the message opens with an acknowledgement', () => {
    const [thread] = buildThreads([
      fromMe({ id: 'm1', createdDateTime: '2026-09-19T01:00:00Z', body: '三社会議の候補日を調整します。' }),
      message({
        id: 'm2',
        createdDateTime: '2026-09-19T02:00:00Z',
        body: '承知しました。調整頂けましたら会議リンクの送付をお願いいたします！',
      }),
    ]);
    const decision = triageThread(thread!, ME);
    expect(decision.needsReply).toBe(true);
    expect(decision.excludeReason).toBeUndefined();
  });

  it('does not treat a closing greeting as a request', () => {
    const [thread] = buildThreads([
      fromMe({ id: 'm1', createdDateTime: '2026-09-19T01:00:00Z', body: '環境は削除済みです。' }),
      message({
        id: 'm2',
        createdDateTime: '2026-09-19T02:00:00Z',
        body: 'ご確認頂きありがとうございます。引き続きよろしくお願いいたします。',
      }),
    ]);
    expect(triageThread(thread!, ME).needsReply).toBe(false);
  });

  it('skips a bot message that mentions somebody else', () => {
    const [thread] = buildThreads([
      message({
        id: 'm1',
        createdDateTime: '2026-09-19T06:00:00Z',
        source: 'teams-channel',
        from: { name: 'azureavabot' },
        body: 'Can you please provide the outcome of the conversation?',
        mentionsOthers: true,
        routing: { kind: 'teams-channel', teamId: 't1', channelId: 'c1', rootMessageId: 'm1' },
      }),
    ]);
    expect(triageThread(thread!, ME).needsReply).toBe(false);
  });

  it('bases the stated reason on the newest message', () => {
    const [thread] = buildThreads([
      message({ id: 'm1', createdDateTime: '2026-09-19T01:00:00Z', body: '以下、ご都合いかがでしょうか。' }),
      message({ id: 'm2', createdDateTime: '2026-09-19T02:00:00Z', body: '見積書のご確認をお願いします。' }),
    ]);
    const decision = triageThread(thread!, ME);
    expect(decision.needsReply).toBe(true);
    expect(decision.reason).toContain('見積書');
  });

  it('ignores a reaction-only reply', () => {    const [thread] = buildThreads([
      fromMe({ id: 'm1', createdDateTime: '2026-09-19T01:00:00Z', body: '資料を送りました。' }),
      message({ id: 'm2', createdDateTime: '2026-09-19T02:00:00Z', body: '👍' }),
    ]);
    expect(triageThread(thread!, ME).needsReply).toBe(false);
  });

  it('boosts VIP senders', () => {
    const build = () =>
      buildThreads([
        message({
          id: 'm1',
          createdDateTime: '2026-09-19T06:00:00Z',
          from: { name: 'Boss', address: 'boss@contoso.com' },
          body: '来週の件、ご意見いただけますか。',
        }),
      ])[0]!;
    const plain = triageThread(build(), ME);
    const vip = triageThread(build(), ME, { vipAddresses: ['boss@contoso.com'] });
    expect(vip.score).toBeGreaterThan(plain.score);
  });
});
