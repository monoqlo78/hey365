import { z } from 'zod';

import type { ActionItem } from '../models/action-item.js';
import { runTriage } from '../services/collector.js';
import { formatSessionSummary, formatTriage } from '../services/format.js';
import { generateDraft, refineDraft } from '../services/reply-generator.js';
import { buildThreads } from '../services/conversation.js';
import { buildIdentity, getMe, normalizeChatMessage, normalizeMailMessage } from '../services/graph.js';
import { sendItem, type SendOutcome } from '../services/sender.js';
import { summarizeSession } from '../services/session-summary.js';
import { checkHealth, runSetup } from '../services/setup.js';
import { getTriage, resolveItems, storeSession, storeTriage, updateItemDraft } from '../services/store.js';
import { Hey365Error, asHey365Error } from '../utils/errors.js';
import { logger } from '../utils/logger.js';
import { timezone } from '../utils/time.js';
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
  hours: z.number().int().min(1).max(720).default(36).describe('何時間前までを対象にするか（既定 36）'),
  includeDrafts: z.boolean().default(true).describe('返信案も生成するか'),
  limit: z.number().int().min(1).max(50).default(10).describe('最大件数'),
  sources: z
    .array(z.enum(['outlook', 'teams', 'meetings']))
    .optional()
    .describe('対象ソース。未指定なら全て'),
};

export async function triageTool(input: {
  hours?: number;
  includeDrafts?: boolean;
  limit?: number;
  sources?: Array<'outlook' | 'teams' | 'meetings'>;
}): Promise<ToolResponse> {
  try {
    const hours = input.hours ?? 36;
    const sources = input.sources;
    const result = await runTriage({
      hours,
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
