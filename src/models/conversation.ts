import type { Language } from '../utils/text.js';

export type SourceKind = 'outlook' | 'teams-chat' | 'teams-channel' | 'meeting';

export interface Participant {
  name: string;
  address?: string;
  id?: string;
}

export interface NormalizedMessage {
  id: string;
  /** Stable id of the thread this message belongs to. */
  conversationId: string;
  source: SourceKind;
  from: Participant;
  to: Participant[];
  cc: Participant[];
  subject: string;
  body: string;
  createdDateTime: string;
  isFromMe: boolean;
  /** True when the signed-in user only appears on CC. */
  isCcOnlyForMe: boolean;
  mentionsMe: boolean;
  /** The message @-mentions somebody, and that somebody is not me. */
  mentionsOthers?: boolean;
  /** Mail that reached me through a distribution list, not as a named recipient. */
  isBroadcast?: boolean;
  /** Outlook filed this under "Other", i.e. newsletters and bulk mail. */
  isLowPriorityInbox?: boolean;
  /** A calendar invitation delivered as mail (`#microsoft.graph.eventMessage`). */
  isEventMessage?: boolean;
  isRead?: boolean;
  importance?: string;
  webLink?: string;
  /** Extra routing info needed to reply into the original thread. */
  routing: ReplyRouting;
}

export type ReplyRouting =
  | { kind: 'outlook'; messageId: string; conversationId: string; subject: string }
  | { kind: 'teams-chat'; chatId: string }
  | { kind: 'teams-channel'; teamId: string; channelId: string; rootMessageId: string };

export interface ConversationThread {
  conversationId: string;
  source: SourceKind;
  subject: string;
  participants: Participant[];
  messages: NormalizedMessage[];
  /** Newest message in the thread regardless of author. */
  lastMessage: NormalizedMessage;
  /** Newest message written by somebody other than the signed-in user. */
  lastInboundMessage?: NormalizedMessage;
  /** Newest message written by the signed-in user. */
  lastOutboundMessage?: NormalizedMessage;
  webLink?: string;
  language: Language;
}
