import type { ConversationThread, NormalizedMessage } from '../models/conversation.js';
import { buildThreads } from './conversation.js';
import {
  buildIdentity,
  getMe,
  normalizeChatMessage,
  normalizeMailMessage,
  searchMail,
  searchTeamsMessages,
  type GraphChatMessage,
} from './graph.js';
import { logger } from '../utils/logger.js';
import { businessDaysBetween, timezone } from '../utils/time.js';

export interface FindMatch {
  conversationId: string;
  source: ConversationThread['source'];
  subject: string;
  participants: string[];
  lastMessageTime: string;
  lastMessageFrom: string;
  excerpt: string;
  /** True when the newest message in the thread is mine. */
  iRepliedLast: boolean;
  /** Business days since the newest message. */
  ageBusinessDays: number;
  webLink?: string;
}

export interface FindResult {
  keyword: string;
  timezone: string;
  matches: FindMatch[];
  warnings: string[];
  scanned: { mail: number; teams: number };
}

export interface FindOptions {
  /** Free text: a person's name, a project, a product — anything searchable. */
  keyword: string;
  limit?: number;
  /** How far back Teams search reaches, in days. Defaults to 180. */
  days?: number;
  includeOutlook?: boolean;
  includeTeams?: boolean;
}

/**
 * Answers "what is the state of X?" across Outlook and Teams at once. Unlike
 * triage this ignores the time window and the needs-reply scoring: it simply
 * shows every thread that mentions the keyword and who spoke last, which is
 * what you need before a meeting or when picking a dropped thread back up.
 */
export async function findConversations(options: FindOptions): Promise<FindResult> {
  const tz = timezone();
  const keyword = options.keyword.trim();
  const warnings: string[] = [];
  const limit = options.limit ?? 10;
  const since = new Date(Date.now() - (options.days ?? 180) * 24 * 3600_000);

  const me = await getMe();
  const identity = buildIdentity(me);

  const [mail, teams] = await Promise.all([
    options.includeOutlook === false
      ? Promise.resolve([])
      : searchMail(keyword, 40).catch((error: unknown) => {
          warnings.push(`Outlook の検索に失敗しました: ${(error as Error).message}`);
          return [];
        }),
    options.includeTeams === false
      ? Promise.resolve([] as GraphChatMessage[])
      : searchTeamsMessages(since, 100, keyword).catch((error: unknown) => {
          warnings.push(`Teams の検索に失敗しました: ${(error as Error).message}`);
          return [] as GraphChatMessage[];
        }),
  ]);

  const normalized: NormalizedMessage[] = [];
  for (const message of mail) {
    const entry = normalizeMailMessage(message, identity);
    if (entry) normalized.push(entry);
  }
  for (const message of teams) {
    try {
      const entry = normalizeChatMessage(message, identity);
      if (entry) normalized.push(entry);
    } catch (error) {
      logger.debug('teams normalize failed', { error: (error as Error).message });
    }
  }

  const now = new Date();
  const matches = buildThreads(normalized)
    .map((thread) => toMatch(thread, now, tz))
    .sort((a, b) => b.lastMessageTime.localeCompare(a.lastMessageTime))
    .slice(0, limit);

  return { keyword, timezone: tz, matches, warnings, scanned: { mail: mail.length, teams: teams.length } };
}

function toMatch(thread: ConversationThread, now: Date, tz: string): FindMatch {
  const last = thread.messages[thread.messages.length - 1]!;
  const lastTime = last.createdDateTime;
  return {
    conversationId: thread.conversationId,
    source: thread.source,
    subject: thread.subject || '(件名なし)',
    participants: thread.participants.map((participant) => participant.name).filter(Boolean),
    lastMessageTime: lastTime,
    lastMessageFrom: last.from?.name ?? '(不明)',
    excerpt: (last.body ?? '').replace(/\s+/g, ' ').slice(0, 160),
    iRepliedLast: Boolean(last.isFromMe),
    ageBusinessDays: businessDaysBetween(new Date(lastTime), now, tz),
    ...(last.webLink ? { webLink: last.webLink } : {}),
  };
}
