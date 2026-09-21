import type { Language } from '../utils/text.js';
import type { Participant, ReplyRouting, SourceKind } from './conversation.js';

export type Importance = 'high' | 'medium' | 'low';

export type TriageReasonCode =
  | 'direct_question'
  | 'explicit_request'
  | 'approval_request'
  | 'decision_required'
  | 'scheduling'
  | 'deadline'
  | 'follow_up'
  | 'mentioned'
  | 'assigned_action'
  | 'unanswered_thread';

export interface TriageSignal {
  code: TriageReasonCode;
  weight: number;
  evidence: string;
  /** Human-readable reason for the matched pattern, per locale. */
  label?: { ja: string; en: string };
}

export interface ActionItem {
  /** 1-based display number shown to the user and used by `hey365_send`. */
  index: number;
  id: string;
  source: SourceKind;
  conversationId: string;
  subject: string;
  sender: Participant;
  participants: Participant[];
  lastMessageTime: string;
  needsReply: boolean;
  alreadyReplied: boolean;
  importance: Importance;
  score: number;
  /** Human readable explanation of why this needs the user's response. */
  reason: string;
  reasonEn: string;
  signals: TriageSignal[];
  deadline?: string;
  deadlineText?: string;
  summary: string;
  excerpt: string;
  language: Language;
  webLink?: string;
  routing: ReplyRouting;
  /** Other conversations merged into this item by the de-duplication pass. */
  mergedFrom?: Array<{ source: SourceKind; conversationId: string; subject: string }>;
  draft?: ReplyDraft;
}

export interface ReplyDraft {
  text: string;
  language: Language;
  /** sha256 prefix of `text`, verified again right before sending. */
  hash: string;
  generator: 'workiq-ask' | 'template' | 'user-edited';
  /** Set when Hey365 cannot answer without a human decision (spec 15). */
  needsUserDecision?: string;
  updatedAt: string;
}

export interface TriageResult {
  generatedAt: string;
  windowHours: number;
  windowStart: string;
  windowEnd: string;
  /** Weekends/holidays the window stepped over, as `YYYY-MM-DD`. */
  skippedDays?: string[];
  timezone: string;
  me: Participant;
  items: ActionItem[];
  scanned: {
    outlookMessages: number;
    teamsMessages: number;
    meetings: number;
    conversations: number;
  };
  skipped: {
    alreadyReplied: number;
    automated: number;
    ccOnly: number;
    fyi: number;
    noSignal: number;
  };
  warnings: string[];
}
