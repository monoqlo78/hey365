import { z } from 'zod';

import type { ActionItem } from '../models/action-item.js';
import { runTriage } from '../services/collector.js';
import { formatDigest, formatFind, formatSessionSummary, formatTriage } from '../services/format.js';
import { buildDigest } from '../services/digest.js';
import { findConversations } from '../services/find.js';
import { installSchedule, listSchedules, removeSchedule } from '../services/schedule.js';
import { generateDraft, refineDraft } from '../services/reply-generator.js';
import { buildThreads } from '../services/conversation.js';
import { buildIdentity, getMe, normalizeChatMessage, normalizeMailMessage } from '../services/graph.js';
import { sendItem, type SendOutcome } from '../services/sender.js';
import { summarizeSession } from '../services/session-summary.js';
import { checkHealth, runSetup } from '../services/setup.js';
import { getTriage, listMutes, muteConversation, resolveItems, storeSession, storeTriage, unmuteConversation, updateItemDraft } from '../services/store.js';
import { Hey365Error, asHey365Error } from '../utils/errors.js';
import { logger } from '../utils/logger.js';
import { timezone, parseAsOf, addBusinessDays, formatDateTime } from '../utils/time.js';
import { CLIENTS, buildServerSpec, findClient, installForClient } from '../services/install.js';

export interface ToolResponse {
  // The MCP SDK allows arbitrary extra fields on a tool result.
  [key: string]: unknown;
  content: Array<{ type: 'text'; text: string }>;
  structuredContent?: Record<string, unknown>;
  isError?: boolean;
}

function ok(text: string, structured?: Record<string, unknown>): ToolResponse {
  return {
    content: [{ type: 'text', text }],
    ...(structured ? { structuredContent: structured } : {}),
  };
}

function fail(error: unknown): ToolResponse {
  const hey = asHey365Error(error);
  logger.warn('tool failed', { code: hey.code });
  const text = [`⚠ ${hey.nextStep('ja')}`, '', `code: ${hey.code}`].join('\n');
  return { content: [{ type: 'text', text }], structuredContent: { error: hey.toPayload() }, isError: true };
}

/* ------------------------------------------------------------------ *
 * hey365 / hey365_triage                                              *
 * ------------------------------------------------------------------ */

export const triageInputSchema = {
  hours: z.number().int().min(1).max(720).optional().describe('何時間前までを対象にするか（既定 36）。businessDays 指定時は無視'),
  businessDays: z
    .number()
    .int()
    .min(1)
    .max(60)
    .optional()
    .describe('何営業日さかのぼるか。基準日自身を1日目として数える（例 3 = 基準日＋前の2営業日）'),
  asOf: z
    .string()
    .optional()
    .describe('基準日。`YYYY-MM-DD` ならその日の終わりまで、ISO 日時ならその時刻まで。未指定なら現在'),
  includeDrafts: z.boolean().default(true).describe('返信案も生成するか'),
  limit: z.number().int().min(1).max(50).default(10).describe('最大件数'),
  sources: z
    .array(z.enum(['outlook', 'teams', 'meetings']))
    .optional()
    .describe('対象ソース。未指定なら全て'),
};

export async function triageTool(input: {
  hours?: number;
  businessDays?: number;
  asOf?: string;
  includeDrafts?: boolean;
  limit?: number;
  sources?: Array<'outlook' | 'teams' | 'meetings'>;
}): Promise<ToolResponse> {
  try {
    if (input.asOf !== undefined && !parseAsOf(input.asOf)) {
      throw new Hey365Error('INVALID_INPUT', undefined, `asOf=${input.asOf}`);
    }

    const sources = input.sources;
    const result = await runTriage({
      ...(input.hours !== undefined ? { hours: input.hours } : {}),
      ...(input.businessDays !== undefined ? { businessDays: input.businessDays } : {}),
      ...(input.asOf !== undefined ? { asOf: input.asOf } : {}),
      limit: input.limit ?? 10,
      includeOutlook: !sources || sources.includes('outlook'),
      includeTeams: !sources || sources.includes('teams'),
      includeMeetings: !sources || sources.includes('meetings'),
    });

    if (input.includeDrafts !== false && result.items.length > 0) {
      await attachDrafts(result.items, result.me.name);
    }

    storeTriage(result);
    return ok(formatTriage(result), { triage: result });
  } catch (error) {
    return fail(error);
  }
}

/**
 * Regenerates the conversation for each item and drafts a reply. Drafts are
 * produced sequentially because `workiq ask` is a metered LLM call.
 */
async function attachDrafts(items: ActionItem[], myName: string): Promise<void> {
  const me = await getMe();
  const identity = buildIdentity(me);

  for (const item of items) {
    try {
      const thread = await rebuildThread(item, identity);
      if (!thread) continue;
      item.draft = await generateDraft({ thread, myName });
    } catch (error) {
      logger.debug('draft generation failed', { index: item.index, error: (error as Error).message });
    }
  }
}

async function rebuildThread(item: ActionItem, identity: ReturnType<typeof buildIdentity>) {
  const { tryFetch } = await import('../services/workiq.js');
  if (item.routing.kind === 'outlook') {
    const collection = await tryFetch<{ value?: Parameters<typeof normalizeMailMessage>[0][] }>(
      `/me/messages?$filter=conversationId eq '${item.conversationId}'&$top=20` +
        '&$select=id,conversationId,subject,body,bodyPreview,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,webLink',
    );
    const messages = (collection?.value ?? [])
      .map((message) => normalizeMailMessage(message, identity))
      .filter((message): message is NonNullable<typeof message> => Boolean(message));
    return buildThreads(messages)[0];
  }

  if (item.routing.kind === 'teams-chat') {
    const { getChatMessages } = await import('../services/graph.js');
    const messages = (await getChatMessages(item.routing.chatId, 15))
      .map((message) => normalizeChatMessage(message, identity, { chatId: (item.routing as { chatId: string }).chatId }))
      .filter((message): message is NonNullable<typeof message> => Boolean(message));
    return buildThreads(messages)[0];
  }

  const { getChannelThread } = await import('../services/graph.js');
  const { teamId, channelId, rootMessageId } = item.routing;
  const thread = await getChannelThread(teamId, channelId, rootMessageId, 20);
  const messages = [thread.root, ...thread.replies]
    .filter((message): message is NonNullable<typeof message> => Boolean(message))
    .map((message) => normalizeChatMessage(message, identity))
    .filter((message): message is NonNullable<typeof message> => Boolean(message));
  return buildThreads(messages)[0];
}

/* ------------------------------------------------------------------ *
 * hey365_session                                                      *
 * ------------------------------------------------------------------ */

export const sessionInputSchema = {
  sessionId: z.string().min(1).describe('会議 ID / Teams thread ID / Outlook conversation ID / 会議名'),
  includeDraft: z.boolean().default(true).describe('返信が必要な場合に返信案も生成するか'),
};

export async function sessionTool(input: { sessionId: string; includeDraft?: boolean }): Promise<ToolResponse> {
  try {
    const summary = await summarizeSession(input.sessionId, { withDraft: input.includeDraft !== false });
    if (!summary.candidates?.length) storeSession(summary);
    return ok(formatSessionSummary(summary, timezone()), { session: summary });
  } catch (error) {
    return fail(error);
  }
}

/* ------------------------------------------------------------------ *
 * hey365_draft                                                        *
 * ------------------------------------------------------------------ */

export const draftInputSchema = {
  itemId: z.string().min(1).describe('返信案の番号（例 "1"）または item id'),
  instruction: z.string().min(1).describe('修正指示（例「もう少し柔らかく」「英語にして」「短く」）'),
};

export async function draftTool(input: { itemId: string; instruction: string }): Promise<ToolResponse> {
  try {
    const [item] = resolveItems([input.itemId]);
    if (!item) throw new Hey365Error('DRAFT_NOT_FOUND');

    const me = await getMe();
    const identity = buildIdentity(me);
    const thread = await rebuildThread(item, identity).catch(() => undefined);

    const current =
      item.draft ??
      (thread
        ? await generateDraft({ thread, myName: identity.name })
        : (() => {
            throw new Hey365Error('DRAFT_NOT_FOUND');
          })());

    const refined = await refineDraft(current, input.instruction, {
      ...(thread ? { thread } : {}),
      myName: identity.name,
    });
    updateItemDraft(item.index, refined);

    const text = [
      `${item.index}. ${item.sender.name} / ${item.subject}`,
      '',
      '修正後の返信案:',
      refined.text,
      '',
      `送信する場合は「${item.index}を送って」と指示してください。`,
    ].join('\n');
    return ok(text, { item: { ...item, draft: refined } });
  } catch (error) {
    return fail(error);
  }
}

/* ------------------------------------------------------------------ *
 * hey365_send                                                         *
 * ------------------------------------------------------------------ */

export const sendInputSchema = {
  itemIds: z
    .array(z.string().min(1))
    .min(1)
    .describe('送信する番号の配列。例 ["1","3"]。全件送信は ["all"]'),
  confirm: z
    .boolean()
    .default(true)
    .describe('ユーザーが明示的に送信を指示した場合のみ true にしてください'),
};

export async function sendTool(input: { itemIds: string[]; confirm?: boolean }): Promise<ToolResponse> {
  try {
    if (input.confirm === false) {
      return ok('送信は行いませんでした。送信する場合は「1を送って」のように明示的に指示してください。');
    }

    const triage = getTriage();
    if (!triage) throw new Hey365Error('DRAFT_NOT_FOUND');

    const items = resolveItems(input.itemIds);
    if (items.length === 0) throw new Hey365Error('DRAFT_NOT_FOUND');

    const outcomes: SendOutcome[] = [];
    for (const item of items) {
      outcomes.push(await sendItem(item, item.draft?.hash));
    }

    const lines: string[] = [];
    const sent = outcomes.filter((outcome) => outcome.ok);
    const failed = outcomes.filter((outcome) => !outcome.ok);

    if (sent.length > 0) {
      lines.push(`${sent.length}件を送信しました。`, '');
      for (const outcome of sent) lines.push(`✅ ${outcome.index}. ${outcome.target}`);
    }
    if (failed.length > 0) {
      if (lines.length > 0) lines.push('');
      lines.push(`${failed.length}件は送信できませんでした。`, '');
      for (const outcome of failed) {
        lines.push(`❌ ${outcome.index}. ${outcome.target}`);
        if (outcome.error) lines.push(`   ${outcome.error.nextStep}`);
      }
    }

    return ok(lines.join('\n'), { outcomes });
  } catch (error) {
    return fail(error);
  }
}

/* ------------------------------------------------------------------ *
 * hey365_snooze / hey365_done                                         *
 * ------------------------------------------------------------------ */

export const snoozeInputSchema = {
  itemIds: z
    .array(z.string().min(1))
    .min(1)
    .describe('対象の番号の配列。例 ["1","3"]。全件は ["all"]'),
  action: z
    .enum(['snooze', 'done', 'unmute'])
    .default('snooze')
    .describe('snooze=一定期間隠す / done=対応済みとして隠す / unmute=再表示する'),
  businessDays: z
    .number()
    .int()
    .min(1)
    .max(60)
    .optional()
    .describe('snooze の場合に何営業日後まで隠すか（既定 1）'),
  note: z.string().optional().describe('メモ（後で理由が分かるように）'),
};

export async function snoozeTool(input: {
  itemIds: string[];
  action?: 'snooze' | 'done' | 'unmute';
  businessDays?: number;
  note?: string;
}): Promise<ToolResponse> {
  try {
    const items = resolveItems(input.itemIds, { requireDraft: false });
    if (items.length === 0) throw new Hey365Error('DRAFT_NOT_FOUND');

    const action = input.action ?? 'snooze';
    if (action === 'unmute') {
      const restored = items.filter((item) => unmuteConversation(item.conversationId));
      const text =
        restored.length > 0
          ? [`${restored.length}件を再表示します。`, '', ...restored.map((item) => `↩ ${item.sender.name} / ${item.subject}`)].join('\n')
          : '再表示できる項目はありませんでした（もともと非表示になっていません）。';
      return ok(text, { restored: restored.map((item) => item.id) });
    }

    // Snoozing hides the item until the start of the Nth next business day, so
    // "また明日" really means the next morning you are at work.
    const until =
      action === 'snooze'
        ? addBusinessDays(new Date(), input.businessDays ?? 1).toISOString()
        : undefined;

    const entries = items.map((item) =>
      muteConversation({
        conversationId: item.conversationId,
        subject: item.subject,
        kind: action === 'snooze' ? 'snoozed' : 'done',
        lastMessageTime: item.lastMessageTime,
        ...(until ? { until } : {}),
        ...(input.note ? { note: input.note } : {}),
      }),
    );

    const lines =
      action === 'done'
        ? [`${entries.length}件を対応済みにしました。次回以降は表示しません。`, '']
        : [`${entries.length}件を ${formatDateTime(until!, timezone())} まで隠します。`, ''];
    for (const item of items) lines.push(`✅ ${item.sender.name} / ${item.subject}`);
    lines.push('', '新しいメッセージが届いた場合は、隠していても再度表示されます。');

    return ok(lines.join('\n'), { muted: entries });
  } catch (error) {
    return fail(error);
  }
}

export const mutesInputSchema = {};

export async function mutesTool(): Promise<ToolResponse> {
  try {
    const mutes = listMutes();
    if (mutes.length === 0) return ok('現在、非表示にしている項目はありません。');

    const tz = timezone();
    const lines = [`非表示中の項目が ${mutes.length}件あります。`, ''];
    for (const mute of mutes) {
      const state = mute.kind === 'done' ? '対応済み' : `${formatDateTime(mute.until ?? '', tz)} まで`;
      lines.push(`- ${mute.subject}（${state}）${mute.note ? ` — ${mute.note}` : ''}`);
    }
    return ok(lines.join('\n'), { mutes });
  } catch (error) {
    return fail(error);
  }
}

/* ------------------------------------------------------------------ *
 * hey365_digest                                                       *
 * ------------------------------------------------------------------ */

export const digestInputSchema = {
  asOf: z.string().optional().describe('基準日。`YYYY-MM-DD` またはISO日時。未指定なら今日'),
  businessDays: z.number().int().min(1).max(60).optional().describe('メール/Teams を何営業日さかのぼるか（既定 3）'),
  limit: z.number().int().min(1).max(50).default(15).describe('未返信の最大件数'),
};

export async function digestTool(input: { asOf?: string; businessDays?: number; limit?: number }): Promise<ToolResponse> {
  try {
    if (input.asOf !== undefined && !parseAsOf(input.asOf)) {
      throw new Hey365Error('INVALID_INPUT', undefined, `asOf=${input.asOf}`);
    }

    const digest = await buildDigest({
      ...(input.asOf !== undefined ? { asOf: input.asOf } : {}),
      ...(input.businessDays !== undefined ? { businessDays: input.businessDays } : {}),
      limit: input.limit ?? 15,
    });

    // The digest numbers items, so keep the same store the reply tools read.
    storeTriage(digest.triage);
    return ok(formatDigest(digest), { digest });
  } catch (error) {
    return fail(error);
  }
}

/* ------------------------------------------------------------------ *
 * hey365_find                                                         *
 * ------------------------------------------------------------------ */

export const findInputSchema = {
  keyword: z.string().min(1).describe('人名・案件名・製品名など。Outlook と Teams を横断して検索する'),
  limit: z.number().int().min(1).max(50).default(10).describe('最大件数'),
  days: z.number().int().min(1).max(3650).optional().describe('Teams を何日さかのぼるか（既定 180）'),
  sources: z.array(z.enum(['outlook', 'teams'])).optional().describe('対象ソース。未指定なら両方'),
};

export async function findTool(input: {
  keyword: string;
  limit?: number;
  days?: number;
  sources?: Array<'outlook' | 'teams'>;
}): Promise<ToolResponse> {
  try {
    const sources = input.sources;
    const result = await findConversations({
      keyword: input.keyword,
      limit: input.limit ?? 10,
      ...(input.days !== undefined ? { days: input.days } : {}),
      includeOutlook: !sources || sources.includes('outlook'),
      includeTeams: !sources || sources.includes('teams'),
    });
    return ok(formatFind(result), { find: result });
  } catch (error) {
    return fail(error);
  }
}

/* ------------------------------------------------------------------ *
 * hey365_schedule                                                     *
 * ------------------------------------------------------------------ */

export const scheduleInputSchema = {
  action: z.enum(['list', 'add', 'remove']).default('list').describe('list=一覧 / add=登録 / remove=解除'),
  job: z.enum(['digest', 'triage']).default('digest').describe('定期実行する内容'),
  at: z.string().optional().describe('実行時刻 `HH:MM`（既定 08:30）'),
  weekdaysOnly: z.boolean().default(true).describe('平日だけ実行するか'),
  out: z.string().optional().describe('結果の書き出し先ファイル'),
  apply: z
    .boolean()
    .default(false)
    .describe('true で実際に OS のスケジューラに登録/解除する。既定は内容の表示のみ'),
};

export async function scheduleTool(input: {
  action?: 'list' | 'add' | 'remove';
  job?: 'digest' | 'triage';
  at?: string;
  weekdaysOnly?: boolean;
  out?: string;
  apply?: boolean;
}): Promise<ToolResponse> {
  try {
    const action = input.action ?? 'list';
    const job = input.job ?? 'digest';

    if (action === 'list') {
      const entries = await listSchedules();
      if (entries.length === 0) {
        return ok('Hey365 の定期実行は登録されていません。\n\n登録するには action="add" を指定してください。', { schedules: [] });
      }
      const lines = [`定期実行が ${entries.length}件あります。`, ''];
      for (const entry of entries) lines.push(`- ${entry.name}（${entry.detail}）`);
      return ok(lines.join('\n'), { schedules: entries });
    }

    if (action === 'remove') {
      const entry = await removeSchedule(job, input.apply === true);
      const text =
        input.apply === true
          ? `定期実行（${job}）を解除しました。`
          : ['以下を実行すると解除されます。', '', entry, '', '実行してよければ apply=true で呼び直してください。'].join('\n');
      return ok(text, { entry, applied: input.apply === true });
    }

    const plan = await installSchedule({
      job,
      ...(input.at !== undefined ? { at: input.at } : {}),
      ...(input.weekdaysOnly !== undefined ? { weekdaysOnly: input.weekdaysOnly } : {}),
      ...(input.out !== undefined ? { out: input.out } : {}),
      apply: input.apply === true,
    });

    const header = plan.applied
      ? `定期実行を登録しました。毎日 ${plan.at}${input.weekdaysOnly === false ? '' : '（平日のみ）'} に実行されます。`
      : `以下の内容で登録します（まだ登録していません）。`;

    const lines = [
      header,
      '',
      `内容: ${plan.job}`,
      `時刻: ${plan.at}${input.weekdaysOnly === false ? '' : '（平日のみ）'}`,
      `出力: ${plan.out}`,
      '',
      plan.entry,
    ];
    if (!plan.applied) lines.push('', 'これでよければ apply=true で呼び直してください。');

    return ok(lines.join('\n'), { plan });
  } catch (error) {
    return fail(error);
  }
}

/* ------------------------------------------------------------------ *
 * hey365_health / hey365_setup / hey365_install                       *
 * ------------------------------------------------------------------ */

export const healthInputSchema = {
  deep: z.boolean().default(false).describe('Outlook / Teams / Calendar の到達性も確認する'),
};

export async function healthTool(input: { deep?: boolean }): Promise<ToolResponse> {
  try {
    const report = await checkHealth({ deep: input.deep ?? false });
    return ok(JSON.stringify(report, null, 2), { health: report as unknown as Record<string, unknown> });
  } catch (error) {
    return fail(error);
  }
}

export const setupInputSchema = {
  interactive: z
    .boolean()
    .default(true)
    .describe('true の場合、必要に応じてブラウザ認証を開始します'),
  force: z.boolean().default(false).describe('接続済みでも再セットアップする'),
};

export async function setupTool(input: { interactive?: boolean; force?: boolean }): Promise<ToolResponse> {
  try {
    const result = await runSetup({
      interactive: input.interactive !== false,
      force: input.force ?? false,
    });
    const lines = [result.message, ''];
    for (const step of result.steps) {
      lines.push(`${step.ok ? '✅' : '❌'} ${step.step}${step.detail ? ` — ${step.detail}` : ''}`);
    }
    lines.push('', `authenticated: ${result.health.authenticated}`, `writeAccess: ${result.health.writeAccess}`);
    return ok(lines.join('\n'), { setup: result as unknown as Record<string, unknown> });
  } catch (error) {
    return fail(error);
  }
}

export const installInputSchema = {
  clients: z
    .array(z.enum(['copilot-cli', 'vscode', 'claude-code', 'claude-desktop', 'codex', 'cursor', 'windsurf', 'scout']))
    .optional()
    .describe('設定を書き込む MCP クライアント。未指定なら設定内容を表示のみ'),
  dryRun: z.boolean().default(false).describe('ファイルに書かず内容だけ表示する'),
};

export async function installTool(input: { clients?: string[]; dryRun?: boolean }): Promise<ToolResponse> {
  try {
    const spec = buildServerSpec();
    const targets = (input.clients ?? []).map(findClient).filter((client): client is NonNullable<typeof client> => Boolean(client));

    if (targets.length === 0) {
      const lines = ['Hey365 を各 MCP クライアントに登録する設定です。', ''];
      for (const client of CLIENTS) {
        lines.push(`## ${client.label} (${client.id})`, `path: ${client.path(process.cwd())}`, '', '```', '```');
      }
      return ok(
        [
          'クライアントを指定してください（clients パラメータ）。利用可能:',
          ...CLIENTS.map((client) => `- ${client.id}: ${client.label}`),
        ].join('\n'),
        { clients: CLIENTS.map((client) => ({ id: client.id, label: client.label, path: client.path(process.cwd()) })) },
      );
    }

    const outcomes = targets.map((client) => installForClient(client, spec, { dryRun: input.dryRun ?? false }));
    const lines = outcomes.map((outcome) =>
      outcome.error
        ? `❌ ${outcome.label}: ${outcome.error}`
        : `${outcome.written ? '✅' : 'ℹ️'} ${outcome.label} → ${outcome.path}`,
    );
    return ok(lines.join('\n'), { outcomes });
  } catch (error) {
    return fail(error);
  }
}
