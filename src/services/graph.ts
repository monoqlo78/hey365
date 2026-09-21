import type { NormalizedMessage, Participant, ReplyRouting, SourceKind } from '../models/conversation.js';
import { detectLanguage, htmlToText, stripQuotedHistory, truncate } from '../utils/text.js';
import { toGraphTimestamp, toKqlDate } from '../utils/time.js';
import { doAction, fetchMany, fetchOne, tryFetch } from './workiq.js';

/* ------------------------------------------------------------------ *
 * Raw Microsoft Graph shapes (only the fields Hey365 actually reads)  *
 * ------------------------------------------------------------------ */

export interface GraphUser {
  id: string;
  displayName?: string;
  mail?: string;
  userPrincipalName?: string;
  givenName?: string;
  surname?: string;
  preferredLanguage?: string;
  mailboxSettings?: { timeZone?: string };
}

interface GraphEmailAddress {
  name?: string;
  address?: string;
}

interface GraphRecipient {
  emailAddress?: GraphEmailAddress;
}

export interface GraphMessage {
  id: string;
  conversationId?: string;
  subject?: string;
  bodyPreview?: string;
  body?: { contentType?: string; content?: string };
  from?: GraphRecipient;
  sender?: GraphRecipient;
  toRecipients?: GraphRecipient[];
  ccRecipients?: GraphRecipient[];
  receivedDateTime?: string;
  sentDateTime?: string;
  isRead?: boolean;
  importance?: string;
  webLink?: string;
  isDraft?: boolean;
  inferenceClassification?: string;
  /** Set locally when the message was read out of the SentItems folder. */
  fromSentItems?: boolean;
  '@odata.type'?: string;
}

export interface GraphChatMessage {
  id: string;
  chatId?: string;
  replyToId?: string | null;
  messageType?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  deletedDateTime?: string | null;
  subject?: string | null;
  summary?: string | null;
  importance?: string;
  webUrl?: string;
  webLink?: string;
  body?: { contentType?: string; content?: string };
  from?: { user?: { id?: string; displayName?: string }; application?: unknown; emailAddress?: GraphEmailAddress };
  mentions?: Array<{ mentioned?: { user?: { id?: string; displayName?: string } } }>;
  channelIdentity?: { teamId?: string; channelId?: string };
  eventDetail?: unknown;
}

export interface GraphChat {
  id: string;
  topic?: string | null;
  chatType?: string;
  createdDateTime?: string;
  lastUpdatedDateTime?: string;
  webUrl?: string;
  members?: Array<{ displayName?: string; userId?: string; email?: string }>;
  lastMessagePreview?: { id?: string; createdDateTime?: string; from?: GraphChatMessage['from'] };
}

export interface GraphEvent {
  id: string;
  iCalUId?: string;
  subject?: string;
  bodyPreview?: string;
  body?: { content?: string; contentType?: string };
  start?: { dateTime?: string; timeZone?: string };
  end?: { dateTime?: string; timeZone?: string };
  organizer?: GraphRecipient;
  attendees?: Array<{ emailAddress?: GraphEmailAddress; status?: { response?: string } }>;
  onlineMeeting?: { joinUrl?: string };
  webLink?: string;
  isCancelled?: boolean;
  responseStatus?: { response?: string };
}

interface GraphCollection<T> {
  value?: T[];
  '@odata.nextLink'?: string;
}

/* ------------------------------------------------------------------ *
 * URL builders                                                        *
 * ------------------------------------------------------------------ */

const MAIL_SELECT =
  'id,conversationId,subject,bodyPreview,from,sender,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,importance,webLink,isDraft,inferenceClassification';

export function mailUrl(since: Date, top: number): string {
  return (
    `/me/messages?$filter=receivedDateTime ge ${toGraphTimestamp(since)}` +
    `&$orderby=receivedDateTime desc&$top=${top}&$select=${MAIL_SELECT}`
  );
}

export function sentMailUrl(since: Date, top: number): string {
  return (
    `/me/mailFolders/SentItems/messages?$filter=sentDateTime ge ${toGraphTimestamp(since)}` +
    `&$orderby=sentDateTime desc&$top=${top}&$select=${MAIL_SELECT}`
  );
}

export function calendarUrl(start: Date, end: Date, top: number): string {
  return (
    `/me/calendarView?startDateTime=${toGraphTimestamp(start)}&endDateTime=${toGraphTimestamp(end)}` +
    `&$orderby=start/dateTime&$top=${top}` +
    '&$select=id,iCalUId,subject,bodyPreview,start,end,organizer,attendees,onlineMeeting,webLink,isCancelled,responseStatus'
  );
}

export function chatMessagesUrl(chatId: string, top: number): string {
  return `/me/chats/${encodeURIComponent(chatId)}/messages?$top=${top}`;
}

export function channelRepliesUrl(teamId: string, channelId: string, messageId: string, top: number): string {
  return `/teams/${teamId}/channels/${encodeURIComponent(channelId)}/messages/${messageId}/replies?$top=${top}`;
}

export function channelMessageUrl(teamId: string, channelId: string, messageId: string): string {
  return `/teams/${teamId}/channels/${encodeURIComponent(channelId)}/messages/${messageId}`;
}

/** Recent channel posts with their replies, so a thread resolves in one call. */
export function channelMessagesUrl(teamId: string, channelId: string, top: number): string {
  return `/teams/${teamId}/channels/${encodeURIComponent(channelId)}/messages?$top=${top}&$expand=replies`;
}

/* ------------------------------------------------------------------ *
 * Reads                                                               *
 * ------------------------------------------------------------------ */

export async function getMe(): Promise<GraphUser> {
  return await fetchOne<GraphUser>('/me?$select=id,displayName,mail,userPrincipalName,givenName,surname,preferredLanguage');
}

export async function getMailboxTimeZone(): Promise<string | undefined> {
  const settings = await tryFetch<{ timeZone?: string }>('/me/mailboxSettings?$select=timeZone');
  return settings?.timeZone;
}

export async function listRecentMail(since: Date, top = 80): Promise<GraphMessage[]> {
  const [inbox, sent] = await fetchMany<GraphCollection<GraphMessage>>([
    mailUrl(since, top),
    // Sent items reach further back so that "I already answered" is detectable
    // for threads that started before the window.
    sentMailUrl(new Date(since.getTime() - 7 * 24 * 3600_000), top),
  ]);
  // Anything pulled from SentItems is mine even if `from` comes back empty.
  const sentItems = (sent?.data?.value ?? []).map((message) => ({ ...message, fromSentItems: true }));
  return [...(inbox?.data?.value ?? []), ...sentItems];
}

export async function listSentMail(since: Date, top = 80): Promise<GraphMessage[]> {
  const result = await tryFetch<GraphCollection<GraphMessage>>(sentMailUrl(since, top));
  return result?.value ?? [];
}

export async function listCalendar(start: Date, end: Date, top = 50): Promise<GraphEvent[]> {
  const result = await tryFetch<GraphCollection<GraphEvent>>(calendarUrl(start, end, top));
  return result?.value ?? [];
}

interface SearchHit {
  hitId?: string;
  rank?: number;
  summary?: string;
  resource?: GraphChatMessage & { '@odata.type'?: string };
}

interface SearchResponse {
  value?: Array<{
    hitsContainers?: Array<{ hits?: SearchHit[]; total?: number; moreResultsAvailable?: boolean }>;
  }>;
}

/**
 * `/me/chats/getAllMessages` is not available in delegated context, so Teams
 * coverage (1:1 chats, group chats and channel posts) goes through the
 * Microsoft Search API instead.
 */
export async function searchTeamsMessages(since: Date, size = 300): Promise<GraphChatMessage[]> {
  // The Search API caps a single request at 100 hits and silently drops the
  // rest, so a wider window needs paging or the oldest messages disappear.
  const pageSize = 100;
  const messages: GraphChatMessage[] = [];
  const seen = new Set<string>();
  for (let from = 0; from < size; from += pageSize) {
    const body = {
      requests: [
        {
          entityTypes: ['chatMessage'],
          query: { queryString: `sent>=${toKqlDate(since)}` },
          from,
          size: Math.min(pageSize, size - from),
        },
      ],
    };
    const response = await doAction<SearchResponse>('/search/query', body);
    const containers = response?.value?.flatMap((entry) => entry.hitsContainers ?? []) ?? [];
    const hits = containers.flatMap((container) => container.hits ?? []);
    for (const hit of hits) {
      const resource = hit.resource;
      if (!resource?.id || seen.has(resource.id)) continue;
      seen.add(resource.id);
      // The search index returns a trimmed body; keep the summary as fallback.
      messages.push({ ...resource, summary: resource.summary ?? hit.summary ?? null });
    }
    if (hits.length === 0) break;
    if (!containers.some((container) => container.moreResultsAvailable === true)) break;
  }
  return messages;
}

export async function getChat(chatId: string): Promise<GraphChat | undefined> {
  return await tryFetch<GraphChat>(`/me/chats/${encodeURIComponent(chatId)}?$expand=members`);
}

export async function getChatMessages(chatId: string, top = 12): Promise<GraphChatMessage[]> {
  const result = await tryFetch<GraphCollection<GraphChatMessage>>(chatMessagesUrl(chatId, top));
  return result?.value ?? [];
}

export async function getChannelThread(
  teamId: string,
  channelId: string,
  messageId: string,
  top = 20,
): Promise<{ root?: GraphChatMessage; replies: GraphChatMessage[] }> {
  // A search hit may be either a root post or a reply, and Graph 404s when a
  // reply id is addressed as a root. Expanding replies over the recent posts
  // resolves both cases in one call.
  const recent = await tryFetch<GraphCollection<GraphChatMessage & { replies?: GraphChatMessage[] }>>(
    channelMessagesUrl(teamId, channelId, Math.max(10, Math.ceil(top / 2))),
  );

  const threads = recent?.value ?? [];
  const match =
    threads.find((thread) => thread.id === messageId) ??
    threads.find((thread) => (thread.replies ?? []).some((reply) => reply.id === messageId));

  if (match) {
    const replies = [...(match.replies ?? [])]
      .filter((reply) => reply.id)
      .sort((a, b) => (a.createdDateTime ?? '').localeCompare(b.createdDateTime ?? ''));
    return { root: match, replies: replies.slice(-top) };
  }

  // Older thread: fall back to addressing it directly.
  const message = await tryFetch<GraphChatMessage>(channelMessageUrl(teamId, channelId, messageId));
  const rootId = message?.replyToId || messageId;
  const [root, replies] = await fetchMany<unknown>([
    channelMessageUrl(teamId, channelId, rootId),
    channelRepliesUrl(teamId, channelId, rootId, top),
  ]);
  return {
    root: (root?.data as GraphChatMessage | undefined) ?? message,
    replies: ((replies?.data as GraphCollection<GraphChatMessage> | undefined)?.value ?? []).filter((r) => r.id),
  };
}

/* ------------------------------------------------------------------ *
 * Normalisation                                                       *
 * ------------------------------------------------------------------ */

export interface Identity {
  id: string;
  name: string;
  addresses: string[];
  /** Extra spellings of my name, e.g. 曽我部 — from HEY365_MY_NAMES. */
  aliases?: string[];
}

export function buildIdentity(user: GraphUser): Identity {
  const addresses = [user.mail, user.userPrincipalName]
    .filter((value): value is string => Boolean(value))
    .map((value) => value.toLowerCase());
  const aliases = [user.givenName, user.surname, ...(process.env.HEY365_MY_NAMES ?? '').split(',')]
    .map((value) => value?.trim().toLowerCase() ?? '')
    .filter((value) => value.length >= 2);
  return {
    id: user.id,
    name: user.displayName ?? user.mail ?? 'me',
    addresses,
    aliases: [...new Set(aliases)],
  };
}

function toParticipant(recipient: GraphRecipient | undefined): Participant {
  const email = recipient?.emailAddress;
  return {
    name: email?.name ?? email?.address ?? 'unknown',
    ...(email?.address ? { address: email.address } : {}),
  };
}

function isMeAddress(identity: Identity, address?: string): boolean {
  if (!address) return false;
  return identity.addresses.includes(address.toLowerCase());
}

export function normalizeMailMessage(message: GraphMessage, identity: Identity): NormalizedMessage | undefined {
  if (!message.id) return undefined;
  const from = toParticipant(message.from ?? message.sender);
  const to = (message.toRecipients ?? []).map(toParticipant);
  const cc = (message.ccRecipients ?? []).map(toParticipant);
  const isFromMe = message.fromSentItems === true || isMeAddress(identity, from.address) || from.name === identity.name;
  const inTo = to.some((recipient) => isMeAddress(identity, recipient.address));
  const inCc = cc.some((recipient) => isMeAddress(identity, recipient.address));
  const bodyRaw = message.body?.contentType?.toLowerCase() === 'html'
    ? htmlToText(message.body.content)
    : (message.body?.content ?? message.bodyPreview ?? '');
  const body = stripQuotedHistory(bodyRaw || message.bodyPreview || '');
  const timestamp = message.receivedDateTime ?? message.sentDateTime ?? new Date().toISOString();
  const conversationId = message.conversationId ?? `mail:${message.id}`;

  return {
    id: message.id,
    conversationId,
    source: 'outlook',
    from,
    to,
    cc,
    subject: message.subject ?? '(no subject)',
    body,
    createdDateTime: timestamp,
    isFromMe,
    isCcOnlyForMe: !isFromMe && !inTo && inCc,
    // Delivered through a distribution list: my address is on neither line.
    ...(!isFromMe && !inTo && !inCc ? { isBroadcast: true } : {}),
    ...(message.inferenceClassification?.toLowerCase() === 'other' ? { isLowPriorityInbox: true } : {}),
    ...(message['@odata.type']?.toLowerCase().includes('eventmessage') ? { isEventMessage: true } : {}),
    mentionsMe: mentionsIdentity(body, identity),
    ...(message.isRead !== undefined ? { isRead: message.isRead } : {}),
    ...(message.importance ? { importance: message.importance } : {}),
    ...(message.webLink ? { webLink: message.webLink } : {}),
    routing: { kind: 'outlook', messageId: message.id, conversationId, subject: message.subject ?? '' },
  };
}

export function normalizeChatMessage(
  message: GraphChatMessage,
  identity: Identity,
  context?: { chatId?: string; participants?: Participant[]; topic?: string },
): NormalizedMessage | undefined {
  if (!message.id || message.deletedDateTime) return undefined;
  if (message.messageType && message.messageType !== 'message') return undefined;

  const fromUser = message.from?.user;
  const fromEmail = message.from?.emailAddress;
  const from: Participant = {
    name: fromUser?.displayName ?? fromEmail?.name ?? fromEmail?.address ?? 'unknown',
    ...(fromEmail?.address ? { address: fromEmail.address } : {}),
    ...(fromUser?.id ? { id: fromUser.id } : {}),
  };
  const isFromMe = from.id === identity.id || isMeAddress(identity, from.address) || from.name === identity.name;

  const bodyRaw = message.body?.contentType?.toLowerCase() === 'html'
    ? htmlToText(message.body.content)
    : (message.body?.content ?? '');
  const body = bodyRaw || htmlToText(message.summary ?? '') || '';

  const channel = message.channelIdentity;
  const source: SourceKind = channel?.teamId ? 'teams-channel' : 'teams-chat';
  const chatId = context?.chatId ?? message.chatId ?? channel?.channelId ?? '';
  const conversationId = channel?.teamId ? `${channel.teamId}|${channel.channelId}` : chatId;

  const explicitMentions = message.mentions ?? [];
  const mentionedMeExplicitly = explicitMentions.some((mention) => mention.mentioned?.user?.id === identity.id);
  const mentionedMe = mentionedMeExplicitly || mentionsIdentity(body, identity);
  const mentionsOthers = explicitMentions.length > 0 && !mentionedMeExplicitly;

  const routing: ReplyRouting = channel?.teamId
    ? {
        kind: 'teams-channel',
        teamId: channel.teamId,
        channelId: channel.channelId ?? '',
        rootMessageId: message.replyToId || message.id,
      }
    : { kind: 'teams-chat', chatId };

  return {
    id: message.id,
    conversationId: conversationId || `teams:${message.id}`,
    source,
    from,
    to: context?.participants ?? [],
    cc: [],
    subject: message.subject || context?.topic || '',
    body,
    createdDateTime: message.createdDateTime ?? message.lastModifiedDateTime ?? new Date().toISOString(),
    isFromMe,
    isCcOnlyForMe: false,
    mentionsMe: mentionedMe,
    ...(mentionsOthers ? { mentionsOthers: true } : {}),
    ...(message.importance ? { importance: message.importance } : {}),
    ...(message.webUrl || message.webLink ? { webLink: message.webUrl ?? message.webLink } : {}),
    routing,
  };
}

/** Detects "Masaaki さん" / "@Masaaki" / "曽我部様" style addressing in free text. */
export function mentionsIdentity(text: string, identity: Identity): boolean {
  if (!text) return false;
  const haystack = text.toLowerCase();
  if (identity.addresses.some((address) => haystack.includes(address))) return true;
  const nameParts = identity.name.split(/\s+/).filter((part) => part.length >= 2);
  const candidates = [...nameParts, ...(identity.aliases ?? [])];
  return candidates.some((part) => haystack.includes(part.toLowerCase()));
}

export function summariseEvent(event: GraphEvent): string {
  const preview = event.bodyPreview ?? htmlToText(event.body?.content);
  return truncate(preview, 400);
}

export function eventLanguage(event: GraphEvent): 'ja' | 'en' {
  return detectLanguage(`${event.subject ?? ''} ${event.bodyPreview ?? ''}`);
}
