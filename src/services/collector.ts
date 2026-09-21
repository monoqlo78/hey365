import type { ActionItem, TriageResult } from '../models/action-item.js';
import type { ConversationThread, NormalizedMessage, Participant } from '../models/conversation.js';
import { buildThreads } from './conversation.js';
import { dedupeActionItems } from './dedupe.js';
import {
  buildIdentity,
  getChannelThread,
  getChat,
  getChatMessages,
  getMe,
  listCalendar,
  listRecentMail,
  normalizeChatMessage,
  normalizeMailMessage,
  searchTeamsMessages,
  type GraphChatMessage,
  type GraphEvent,
  type Identity,
} from './graph.js';
import { toActionItem, triageThread, type TriageDecision } from './triage.js';
import { logger } from '../utils/logger.js';
import { Hey365Error, asHey365Error } from '../utils/errors.js';
import { buildWindow, isWithin, timezone, type TimeWindow } from '../utils/time.js';

export interface CollectOptions {
  hours: number;
  /** Cap on Teams conversations expanded with a full thread read. */
  maxTeamsThreads?: number;
  includeTeams?: boolean;
  includeOutlook?: boolean;
  includeMeetings?: boolean;
  now?: Date;
}

export interface CollectedContext {
  identity: Identity;
  window: TimeWindow;
  threads: ConversationThread[];
  events: GraphEvent[];
  warnings: string[];
  counts: { outlookMessages: number; teamsMessages: number; meetings: number };
}

function vipAddresses(): string[] {
  return (process.env.HEY365_VIP ?? '')
    .split(',')
    .map((entry) => entry.trim())
    .filter(Boolean);
}

/**
 * Gathers the raw Microsoft 365 surface for the window: Outlook threads, Teams
 * chat/channel threads and calendar events.
 */
export async function collect(options: CollectOptions): Promise<CollectedContext> {
  const window = buildWindow(options.hours, options.now);
  const warnings: string[] = [];

  const me = await getMe();
  const identity = buildIdentity(me);

  const includeOutlook = options.includeOutlook !== false;
  const includeTeams = options.includeTeams !== false;
  const includeMeetings = options.includeMeetings !== false;

  const [mailMessages, teamsHits, events] = await Promise.all([
    includeOutlook
      ? listRecentMail(window.start).catch((error) => {
          warnings.push(describeWarning('Outlook', error));
          return [];
        })
      : Promise.resolve([]),
    includeTeams
      ? searchTeamsMessages(window.start).catch((error) => {
          warnings.push(describeWarning('Teams', error));
          return [] as GraphChatMessage[];
        })
      : Promise.resolve([] as GraphChatMessage[]),
    includeMeetings
      ? listCalendar(window.start, new Date(window.end.getTime() + 7 * 24 * 3600_000)).catch((error) => {
          warnings.push(describeWarning('Calendar', error));
          return [] as GraphEvent[];
        })
      : Promise.resolve([] as GraphEvent[]),
  ]);

  const mailNormalized = mailMessages
    .map((message) => normalizeMailMessage(message, identity))
    .filter((message): message is NormalizedMessage => Boolean(message));

  const teamsNormalized = await expandTeamsThreads(teamsHits, identity, window, options.maxTeamsThreads ?? 14, warnings);

  const threads = buildThreads([...mailNormalized, ...teamsNormalized]).filter((thread) =>
    // Keep threads that saw activity inside the window; older context is only
    // used to decide whether the user already replied.
    thread.messages.some((message) => isWithin(message.createdDateTime, window)),
  );

  return {
    identity,
    window,
    threads,
    events: events.filter((event) => !event.isCancelled),
    warnings,
    counts: {
      outlookMessages: mailNormalized.length,
      teamsMessages: teamsNormalized.length,
      meetings: events.length,
    },
  };
}

/**
 * The search index only returns matching messages, not whole conversations.
 * For the most promising conversations we read the actual thread so that the
 * "did I already answer?" check has real data.
 */
async function expandTeamsThreads(
  hits: GraphChatMessage[],
  identity: Identity,
  window: TimeWindow,
  maxThreads: number,
  warnings: string[],
): Promise<NormalizedMessage[]> {
  const inWindow = hits.filter((hit) => isWithin(hit.createdDateTime ?? hit.lastModifiedDateTime, window));

  const chatIds = new Set<string>();
  const channelKeys = new Map<string, { teamId: string; channelId: string; messageId: string }>();

  for (const hit of inWindow) {
    const channel = hit.channelIdentity;
    if (channel?.teamId && channel.channelId) {
      const key = `${channel.teamId}|${channel.channelId}|${hit.replyToId || hit.id}`;
      if (!channelKeys.has(key)) {
        channelKeys.set(key, {
          teamId: channel.teamId,
          channelId: channel.channelId,
          messageId: hit.replyToId || hit.id,
        });
      }
    } else if (hit.chatId) {
      chatIds.add(hit.chatId);
    }
  }

  const messages: NormalizedMessage[] = [];
  const chatBudget = Math.max(1, Math.ceil(maxThreads * 0.7));
  const channelBudget = Math.max(1, maxThreads - chatBudget);

  for (const chatId of [...chatIds].slice(0, chatBudget)) {
    try {
      const [chat, chatMessages] = await Promise.all([getChat(chatId), getChatMessages(chatId)]);
      const participants: Participant[] = (chat?.members ?? []).map((member) => ({
        name: member.displayName ?? member.email ?? 'unknown',
        ...(member.email ? { address: member.email } : {}),
        ...(member.userId ? { id: member.userId } : {}),
      }));
      for (const message of chatMessages) {
        const normalized = normalizeChatMessage(message, identity, {
          chatId,
          participants,
          ...(chat?.topic ? { topic: chat.topic } : {}),
        });
        if (normalized) messages.push(normalized);
      }
    } catch (error) {
      logger.debug('chat expand failed', { error: (error as Error).message });
    }
  }

  for (const target of [...channelKeys.values()].slice(0, channelBudget)) {
    try {
      const thread = await getChannelThread(target.teamId, target.channelId, target.messageId);
      for (const message of [thread.root, ...thread.replies]) {
        if (!message) continue;
        const normalized = normalizeChatMessage(message, identity);
        if (normalized) messages.push(normalized);
      }
    } catch (error) {
      logger.debug('channel expand failed', { error: (error as Error).message });
    }
  }

  // Search hits that we could not expand still carry enough context to triage.
  const expandedIds = new Set(messages.map((message) => message.id));
  for (const hit of inWindow) {
    if (expandedIds.has(hit.id)) continue;
    const normalized = normalizeChatMessage(hit, identity);
    if (normalized) messages.push(normalized);
  }

  if (inWindow.length > 0 && messages.length === 0) {
    warnings.push('Teams のメッセージ本文を取得できませんでした（権限またはポリシーの可能性があります）。');
  }

  return messages;
}

export interface TriageRunOptions extends CollectOptions {
  /** Include low-confidence items that fell below the needsReply threshold. */
  includeLowConfidence?: boolean;
  limit?: number;
}

export async function runTriage(options: TriageRunOptions): Promise<TriageResult> {
  const context = await collect(options);
  const vip = vipAddresses();

  const skipped = { alreadyReplied: 0, automated: 0, ccOnly: 0, fyi: 0, noSignal: 0 };
  const candidates: Array<{ thread: ConversationThread; decision: TriageDecision }> = [];

  for (const thread of context.threads) {
    const decision = triageThread(thread, context.identity, { vipAddresses: vip });
    if (decision.needsReply || (options.includeLowConfidence && decision.score > 0)) {
      candidates.push({ thread, decision });
      continue;
    }
    switch (decision.excludeReason) {
      case 'already_replied':
        skipped.alreadyReplied += 1;
        break;
      case 'automated':
      case 'reaction':
        skipped.automated += 1;
        break;
      case 'cc_only':
      case 'broadcast':
      case 'delegated':
        skipped.ccOnly += 1;
        break;
      case 'fyi':
      case 'completed':
        skipped.fyi += 1;
        break;
      default:
        skipped.noSignal += 1;
    }
  }

  candidates.sort((a, b) => {
    if (b.decision.score !== a.decision.score) return b.decision.score - a.decision.score;
    return new Date(b.thread.lastMessage.createdDateTime).getTime() - new Date(a.thread.lastMessage.createdDateTime).getTime();
  });

  const limited = candidates.slice(0, options.limit ?? 25);
  const items: ActionItem[] = dedupeActionItems(
    limited.map((candidate, position) => toActionItem(candidate.thread, candidate.decision, position + 1)),
  ).map((item, position) => ({ ...item, index: position + 1 }));

  return {
    generatedAt: new Date().toISOString(),
    windowHours: context.window.hours,
    windowStart: context.window.startIso,
    windowEnd: context.window.endIso,
    skippedDays: context.window.skippedDays,
    timezone: timezone(),
    me: { name: context.identity.name, address: context.identity.addresses[0] ?? '', id: context.identity.id },
    items,
    scanned: {
      outlookMessages: context.counts.outlookMessages,
      teamsMessages: context.counts.teamsMessages,
      meetings: context.counts.meetings,
      conversations: context.threads.length,
    },
    skipped,
    warnings: context.warnings,
  };
}

function describeWarning(surface: string, error: unknown): string {
  const hey = error instanceof Hey365Error ? error : asHey365Error(error);
  return `${surface}: ${hey.nextStep('ja')}`;
}
