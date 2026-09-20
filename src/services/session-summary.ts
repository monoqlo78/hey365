import type { SessionCandidate, SessionSummary } from '../models/session.js';
import type { ConversationThread, NormalizedMessage, Participant } from '../models/conversation.js';
import { Hey365Error } from '../utils/errors.js';
import { logger } from '../utils/logger.js';
import { detectLanguage, htmlToText, truncate, type Language } from '../utils/text.js';
import { buildThreads, renderTranscript } from './conversation.js';
import {
  buildIdentity,
  getChannelThread,
  getChat,
  getChatMessages,
  getMe,
  normalizeChatMessage,
  normalizeMailMessage,
  type GraphEvent,
  type GraphMessage,
  type Identity,
} from './graph.js';
import { generateDraft } from './reply-generator.js';
import { ask, tryFetch } from './workiq.js';

const TEAMS_THREAD = /^19:[^@]+@thread\.(v2|tacv2|skype)$/i;
const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const OUTLOOK_ID = /^(AAMk|AAQk|AQMk)/;

interface ResolvedSession {
  kind: 'meeting' | 'teams-chat' | 'teams-channel' | 'outlook';
  title: string;
  event?: GraphEvent;
  thread?: ConversationThread;
  webLink?: string;
  participants: Participant[];
}

/** Resolves a user-supplied session id without guessing its type (spec 13). */
export async function resolveSession(
  sessionId: string,
  identity: Identity,
): Promise<{ resolved?: ResolvedSession; candidates: SessionCandidate[] }> {
  const trimmed = sessionId.trim();
  if (!trimmed) return { candidates: [] };

  if (TEAMS_THREAD.test(trimmed)) {
    const resolved = await resolveTeamsThread(trimmed, identity);
    return resolved ? { resolved, candidates: [] } : { candidates: [] };
  }

  if (OUTLOOK_ID.test(trimmed)) {
    const resolved = (await resolveOutlookConversation(trimmed, identity)) ?? (await resolveEventById(trimmed, identity));
    return resolved ? { resolved, candidates: [] } : { candidates: [] };
  }

  if (GUID.test(trimmed)) {
    const event = await findEventByICalUid(trimmed);
    if (event) return { resolved: await buildMeetingSession(event, identity), candidates: [] };
  }

  // Fall back to a keyword search across meetings and Teams threads.
  return await searchCandidates(trimmed, identity);
}

async function resolveTeamsThread(threadId: string, identity: Identity): Promise<ResolvedSession | undefined> {
  const chat = await getChat(threadId);
  const messages = await getChatMessages(threadId, 40);
  if (!chat && messages.length === 0) return undefined;

  const participants: Participant[] = (chat?.members ?? []).map((member) => ({
    name: member.displayName ?? member.email ?? 'unknown',
    ...(member.email ? { address: member.email } : {}),
    ...(member.userId ? { id: member.userId } : {}),
  }));

  const normalized = messages
    .map((message) =>
      normalizeChatMessage(message, identity, {
        chatId: threadId,
        participants,
        ...(chat?.topic ? { topic: chat.topic } : {}),
      }),
    )
    .filter((message): message is NormalizedMessage => Boolean(message));

  const [thread] = buildThreads(normalized);
  if (!thread) return undefined;

  const session: ResolvedSession = {
    kind: 'teams-chat',
    title: chat?.topic || thread.subject || 'Teams conversation',
    thread,
    participants: participants.length > 0 ? participants : thread.participants,
  };
  if (chat?.webUrl) session.webLink = chat.webUrl;

  // A meeting chat is better presented with its calendar event attached.
  if (chat?.chatType === 'meeting') {
    const event = await findEventByChatTopic(chat.topic ?? '');
    if (event) {
      session.kind = 'meeting';
      session.event = event;
      session.title = event.subject ?? session.title;
    }
  }
  return session;
}

async function resolveOutlookConversation(id: string, identity: Identity): Promise<ResolvedSession | undefined> {
  const single = await tryFetch<GraphMessage>(
    `/me/messages/${id}?$select=id,conversationId,subject,body,from,toRecipients,ccRecipients,receivedDateTime,webLink`,
  );
  const conversationId = single?.conversationId ?? id;
  const collection = await tryFetch<{ value?: GraphMessage[] }>(
    `/me/messages?$filter=conversationId eq '${conversationId}'&$top=30&$select=id,conversationId,subject,body,bodyPreview,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,webLink`,
  );
  const messages = (collection?.value ?? (single ? [single] : []))
    .map((message) => normalizeMailMessage(message, identity))
    .filter((message): message is NormalizedMessage => Boolean(message));
  if (messages.length === 0) return undefined;

  const [thread] = buildThreads(messages);
  if (!thread) return undefined;
  const session: ResolvedSession = {
    kind: 'outlook',
    title: thread.subject,
    thread,
    participants: thread.participants,
  };
  if (thread.webLink) session.webLink = thread.webLink;
  return session;
}

async function resolveEventById(id: string, identity: Identity): Promise<ResolvedSession | undefined> {
  const event = await tryFetch<GraphEvent>(`/me/events/${id}`);
  if (!event?.id) return undefined;
  return await buildMeetingSession(event, identity);
}

async function findEventByICalUid(uid: string): Promise<GraphEvent | undefined> {
  const result = await tryFetch<{ value?: GraphEvent[] }>(`/me/events?$filter=iCalUId eq '${uid}'&$top=1`);
  return result?.value?.[0];
}

async function findEventByChatTopic(topic: string): Promise<GraphEvent | undefined> {
  if (!topic) return undefined;
  const escaped = topic.replace(/'/g, "''").slice(0, 80);
  const result = await tryFetch<{ value?: GraphEvent[] }>(
    `/me/events?$filter=subject eq '${escaped}'&$top=1&$orderby=start/dateTime desc`,
  );
  return result?.value?.[0];
}

async function buildMeetingSession(event: GraphEvent, identity: Identity): Promise<ResolvedSession> {
  const participants: Participant[] = (event.attendees ?? []).map((attendee) => ({
    name: attendee.emailAddress?.name ?? attendee.emailAddress?.address ?? 'unknown',
    ...(attendee.emailAddress?.address ? { address: attendee.emailAddress.address } : {}),
  }));

  const session: ResolvedSession = {
    kind: 'meeting',
    title: event.subject ?? 'Meeting',
    event,
    participants,
  };
  if (event.webLink) session.webLink = event.webLink;

  // Attach the meeting chat when the join URL exposes its thread id.
  const threadId = extractThreadId(event.onlineMeeting?.joinUrl);
  if (threadId) {
    const chatSession = await resolveTeamsThread(threadId, identity);
    if (chatSession?.thread) session.thread = chatSession.thread;
  }
  return session;
}

export function extractThreadId(joinUrl: string | undefined): string | undefined {
  if (!joinUrl) return undefined;
  const match = /19%3ameeting_[^%]+%40thread\.v2/i.exec(joinUrl);
  if (!match) return undefined;
  return decodeURIComponent(match[0]);
}

async function searchCandidates(
  query: string,
  identity: Identity,
): Promise<{ resolved?: ResolvedSession; candidates: SessionCandidate[] }> {
  const escaped = query.replace(/'/g, "''");
  const events = await tryFetch<{ value?: GraphEvent[] }>(
    `/me/events?$filter=contains(subject,'${escaped}')&$top=5&$orderby=start/dateTime desc`,
  );
  const eventCandidates: SessionCandidate[] = (events?.value ?? []).map((event) => ({
    id: event.id,
    kind: 'meeting',
    title: event.subject ?? 'Meeting',
    ...(event.start?.dateTime ? { when: event.start.dateTime } : {}),
    hint: 'calendar event',
  }));

  if (eventCandidates.length === 1 && eventCandidates[0]) {
    const event = events?.value?.[0];
    if (event) return { resolved: await buildMeetingSession(event, identity), candidates: [] };
  }

  return { candidates: eventCandidates };
}

/* ------------------------------------------------------------------ *
 * Summarisation                                                       *
 * ------------------------------------------------------------------ */

interface RawSummary {
  summary?: string;
  decisions?: string[];
  keyPoints?: string[];
  actionItems?: Array<{ owner?: string; text?: string; due?: string }>;
  openQuestions?: string[];
  replyNeeded?: boolean;
  replyReason?: string;
}

export async function summarizeSession(sessionId: string, options: { withDraft?: boolean } = {}): Promise<SessionSummary> {
  const me = await getMe();
  const identity = buildIdentity(me);
  const { resolved, candidates } = await resolveSession(sessionId, identity);

  if (!resolved) {
    if (candidates.length > 0) {
      return {
        sessionId,
        source: 'meeting',
        title: '',
        participants: [],
        summary: '',
        decisions: [],
        keyPoints: [],
        actionItems: [],
        openQuestions: [],
        replyNeeded: false,
        language: 'ja',
        candidates,
      };
    }
    throw new Hey365Error('SESSION_NOT_FOUND', `No meeting or conversation matched "${sessionId}"`);
  }

  const context = buildSessionContext(resolved);
  const language = detectLanguage(context);
  const raw = await askForSummary(context, identity.name, language);

  const summary: SessionSummary = {
    sessionId,
    source: resolved.kind === 'meeting' ? 'meeting' : resolved.kind,
    title: resolved.title,
    participants: resolved.participants,
    summary: raw.summary ?? '',
    decisions: raw.decisions ?? [],
    keyPoints: raw.keyPoints ?? [],
    actionItems: (raw.actionItems ?? []).map((entry) => ({
      owner: entry.owner ?? 'unassigned',
      isMine: isMine(entry.owner ?? '', identity),
      text: entry.text ?? '',
      ...(entry.due ? { due: entry.due } : {}),
    })),
    openQuestions: raw.openQuestions ?? [],
    replyNeeded: raw.replyNeeded ?? false,
    language,
  };
  if (raw.replyReason) summary.replyReason = raw.replyReason;
  if (resolved.event?.start?.dateTime) summary.startTime = resolved.event.start.dateTime;
  if (resolved.event?.end?.dateTime) summary.endTime = resolved.event.end.dateTime;
  if (resolved.webLink) summary.webLink = resolved.webLink;

  if (summary.replyNeeded && resolved.thread && options.withDraft !== false) {
    try {
      summary.draft = await generateDraft({ thread: resolved.thread, myName: identity.name, language });
    } catch (error) {
      logger.debug('session draft failed', { error: (error as Error).message });
    }
  }

  return summary;
}

function buildSessionContext(resolved: ResolvedSession): string {
  const parts: string[] = [];
  if (resolved.event) {
    parts.push(`MEETING: ${resolved.event.subject ?? ''}`);
    parts.push(`WHEN: ${resolved.event.start?.dateTime ?? ''} - ${resolved.event.end?.dateTime ?? ''}`);
    parts.push(`ORGANIZER: ${resolved.event.organizer?.emailAddress?.name ?? ''}`);
    parts.push(`ATTENDEES: ${resolved.participants.map((p) => p.name).join(', ')}`);
    const body = htmlToText(resolved.event.body?.content) || resolved.event.bodyPreview || '';
    if (body) parts.push(`INVITE BODY:\n${truncate(body, 1500)}`);
  }
  if (resolved.thread) {
    parts.push(`CONVERSATION (oldest first, "ME" is the user):\n${renderTranscript(resolved.thread, 30, 700)}`);
  }
  return parts.join('\n\n');
}

async function askForSummary(context: string, myName: string, language: Language): Promise<RawSummary> {
  const prompt = [
    `Summarise the following Microsoft 365 session for ${myName}.`,
    'Respond with ONLY a JSON object, no markdown fences, using exactly these keys:',
    '{"summary": string, "decisions": string[], "keyPoints": string[], "actionItems": [{"owner": string, "text": string, "due": string}], "openQuestions": string[], "replyNeeded": boolean, "replyReason": string}',
    `Write all free text in ${language === 'ja' ? 'Japanese' : 'English'}.`,
    `"owner" must be a participant name, or "${myName}" when the action belongs to the user.`,
    '"replyNeeded" is true only when somebody is waiting for a response from the user.',
    'Never invent facts that are not present in the material below.',
    '',
    context,
  ].join('\n');

  try {
    const response = await ask(prompt);
    const parsed = parseJsonObject(response.text);
    if (parsed) return parsed;
    return { summary: response.text.trim() };
  } catch (error) {
    logger.warn('session summarisation failed', { error: (error as Error).message });
    return { summary: '' };
  }
}

export function parseJsonObject(raw: string): RawSummary | undefined {
  const text = raw.replace(/^```[a-z]*\s*/i, '').replace(/```\s*$/i, '');
  const start = text.indexOf('{');
  const end = text.lastIndexOf('}');
  if (start < 0 || end <= start) return undefined;
  try {
    return JSON.parse(text.slice(start, end + 1)) as RawSummary;
  } catch {
    return undefined;
  }
}

function isMine(owner: string, identity: Identity): boolean {
  const normalized = owner.trim().toLowerCase();
  if (!normalized) return false;
  if (['me', '私', '自分', 'myself'].includes(normalized)) return true;
  if (identity.addresses.includes(normalized)) return true;
  return identity.name.toLowerCase().includes(normalized) || normalized.includes(identity.name.toLowerCase());
}

export { getChannelThread };
