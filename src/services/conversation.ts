import type { ConversationThread, NormalizedMessage, Participant } from '../models/conversation.js';
import { detectLanguage } from '../utils/text.js';

/**
 * Groups normalized messages into threads and computes the state that triage
 * depends on: who spoke last, whether the signed-in user already answered the
 * latest inbound message, and who is involved (spec section 7).
 */
export function buildThreads(messages: NormalizedMessage[]): ConversationThread[] {
  const byConversation = new Map<string, NormalizedMessage[]>();
  for (const message of messages) {
    const key = `${message.source === 'outlook' ? 'mail' : 'teams'}:${message.conversationId}`;
    const bucket = byConversation.get(key);
    if (bucket) bucket.push(message);
    else byConversation.set(key, [message]);
  }

  const threads: ConversationThread[] = [];
  for (const bucket of byConversation.values()) {
    const thread = toThread(bucket);
    if (thread) threads.push(thread);
  }
  return threads.sort(
    (a, b) => new Date(b.lastMessage.createdDateTime).getTime() - new Date(a.lastMessage.createdDateTime).getTime(),
  );
}

function toThread(messages: NormalizedMessage[]): ConversationThread | undefined {
  const deduped = dedupeById(messages).sort(
    (a, b) => new Date(a.createdDateTime).getTime() - new Date(b.createdDateTime).getTime(),
  );
  const lastMessage = deduped[deduped.length - 1];
  if (!lastMessage) return undefined;

  const lastInboundMessage = [...deduped].reverse().find((message) => !message.isFromMe);
  const lastOutboundMessage = [...deduped].reverse().find((message) => message.isFromMe);

  const subject =
    deduped.find((message) => message.subject && message.subject.trim())?.subject ?? lastMessage.subject ?? '';

  const languageSample = deduped
    .slice(-4)
    .map((message) => `${message.subject} ${message.body}`)
    .join(' ');

  const thread: ConversationThread = {
    conversationId: lastMessage.conversationId,
    source: lastMessage.source,
    subject,
    participants: collectParticipants(deduped),
    messages: deduped,
    lastMessage,
    language: detectLanguage(languageSample),
  };
  if (lastInboundMessage) thread.lastInboundMessage = lastInboundMessage;
  if (lastOutboundMessage) thread.lastOutboundMessage = lastOutboundMessage;
  const webLink = deduped.find((message) => message.webLink)?.webLink;
  if (webLink) thread.webLink = webLink;
  return thread;
}

function dedupeById(messages: NormalizedMessage[]): NormalizedMessage[] {
  const seen = new Map<string, NormalizedMessage>();
  for (const message of messages) {
    if (!seen.has(message.id)) seen.set(message.id, message);
  }
  return [...seen.values()];
}

function collectParticipants(messages: NormalizedMessage[]): Participant[] {
  const seen = new Map<string, Participant>();
  for (const message of messages) {
    for (const participant of [message.from, ...message.to, ...message.cc]) {
      const key = (participant.address ?? participant.id ?? participant.name).toLowerCase();
      if (!seen.has(key)) seen.set(key, participant);
    }
  }
  return [...seen.values()];
}

/**
 * True when the user has already responded to the newest inbound message.
 * This is the difference between "山田さん→私→私が回答済み" (no action) and
 * "私→山田さん→山田さんが再質問" (action required).
 */
export function hasAnsweredLatestInbound(thread: ConversationThread): boolean {
  const { lastInboundMessage, lastOutboundMessage } = thread;
  if (!lastInboundMessage) return true;
  if (!lastOutboundMessage) return false;
  return new Date(lastOutboundMessage.createdDateTime).getTime() > new Date(lastInboundMessage.createdDateTime).getTime();
}

/** Messages the user has not answered yet, oldest first. */
export function unansweredInbound(thread: ConversationThread): NormalizedMessage[] {
  const lastOutboundTime = thread.lastOutboundMessage
    ? new Date(thread.lastOutboundMessage.createdDateTime).getTime()
    : 0;
  return thread.messages.filter(
    (message) => !message.isFromMe && new Date(message.createdDateTime).getTime() > lastOutboundTime,
  );
}

/** Renders the tail of a thread as plain text for prompts and summaries. */
export function renderTranscript(thread: ConversationThread, maxMessages = 8, maxCharsPerMessage = 900): string {
  return thread.messages
    .slice(-maxMessages)
    .map((message) => {
      const who = message.isFromMe ? 'ME' : message.from.name;
      const body = message.body.length > maxCharsPerMessage
        ? `${message.body.slice(0, maxCharsPerMessage)}…`
        : message.body;
      return `[${message.createdDateTime}] ${who}: ${body}`;
    })
    .join('\n\n');
}
