import type { Participant, SourceKind } from './conversation.js';
import type { ReplyDraft } from './action-item.js';

export interface SessionActionItem {
  owner: string;
  isMine: boolean;
  text: string;
  due?: string;
}

export interface SessionSummary {
  sessionId: string;
  source: SourceKind;
  title: string;
  startTime?: string;
  endTime?: string;
  participants: Participant[];
  summary: string;
  decisions: string[];
  keyPoints: string[];
  actionItems: SessionActionItem[];
  openQuestions: string[];
  replyNeeded: boolean;
  replyReason?: string;
  draft?: ReplyDraft;
  webLink?: string;
  language: 'ja' | 'en';
  /** Populated when the session id matched more than one entity (spec 13). */
  candidates?: SessionCandidate[];
}

export interface SessionCandidate {
  id: string;
  kind: SourceKind;
  title: string;
  when?: string;
  hint?: string;
}
