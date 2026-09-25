import type { ActionItem, Importance, TriageSignal } from '../models/action-item.js';
import type { ConversationThread, NormalizedMessage } from '../models/conversation.js';
import type { Identity } from './graph.js';
import { mentionsIdentity } from './graph.js';
import { hasAnsweredLatestInbound, unansweredInbound } from './conversation.js';
import { shortHash, stripEmphasis, truncate } from '../utils/text.js';
import { businessDaysBetween } from '../utils/time.js';

/* ------------------------------------------------------------------ *
 * Signal dictionaries                                                 *
 * ------------------------------------------------------------------ */

interface Pattern {
  code: TriageSignal['code'];
  weight: number;
  regex: RegExp;
  label: { ja: string; en: string };
}

const REQUEST_PATTERNS: Pattern[] = [
  {
    code: 'explicit_request',
    weight: 32,
    regex: /(ご回答|回答(を)?(ください|お願い)|ご返信|返信(を)?(ください|お願い)|お返事|please reply|please respond|awaiting your (reply|response)|let me know)/i,
    label: { ja: 'あなたへの返信依頼', en: 'A reply is explicitly requested' },
  },
  {
    code: 'explicit_request',
    weight: 28,
    regex: /(ご確認(ください|のほど|お願い)|確認(を)?お願い|チェックお願い|please (check|review|confirm)|could you (check|review|confirm)|review (this|the))/i,
    label: { ja: '確認を依頼されています', en: 'You are asked to check or review something' },
  },
  {
    code: 'approval_request',
    weight: 34,
    regex: /(ご承認|承認(を)?(ください|お願い)|approval|approve|sign[- ]?off|決裁|許可をいただ)/i,
    label: { ja: '承認待ちです', en: 'An approval is pending on you' },
  },
  {
    code: 'decision_required',
    weight: 30,
    regex: /(どう(しますか|されますか|しましょう)|いかがでしょうか|ご判断|判断(を)?(ください|お願い)|どちらが|進めてよい|進めても(よろしい|良い)|what do you (think|want)|your (call|decision|thoughts)|which (one|option))/i,
    label: { ja: 'あなたの判断が求められています', en: 'Your decision is required' },
  },
  {
    code: 'direct_question',
    weight: 26,
    regex: /(可能でしょうか|できますでしょうか|いただけますか|いただけますでしょうか|教えてください|ご教示|ご意見|are you able to|can you|could you|would you|do you know|any thoughts)/i,
    label: { ja: 'あなたに直接質問しています', en: 'You are asked a direct question' },
  },
  {
    code: 'scheduling',
    weight: 24,
    regex: /(日程|ご都合|空いて(いますか|ますか)|候補日|スケジュール調整|打ち合わせ(の)?(日|設定)|availability|reschedul|propose a time|does .{0,20}work for you|book a (slot|meeting)|時間を(いただけ|取れ))/i,
    label: { ja: '日程調整の回答待ちです', en: 'A scheduling answer is pending' },
  },
  {
    code: 'follow_up',
    weight: 22,
    regex: /(その後いかが|その後どう|進捗(は)?いかが|リマインド|再送|再度のご連絡|follow(ing)?[- ]?up|gentle reminder|any update|status update|bump)/i,
    label: { ja: 'フォローアップの催促です', en: 'This is a follow-up nudge' },
  },
  {
    code: 'assigned_action',
    weight: 20,
    regex: /(お願いできますか|ご対応(ください|お願い)|対応(を)?お願い|やっていただけ|assign(ed)? to you|action for you|over to you|your action)/i,
    label: { ja: 'あなたへの作業依頼です', en: 'A task is assigned to you' },
  },
  {
    code: 'assigned_action',
    weight: 26,
    // "〜をお願いいたします" is a request, but the closing "よろしくお願いいたします" is not.
    regex: /(?<!よろしく)(?<!宜しく)(?<!何卒)(?<!引き続き)(?<!今後とも)お願い(いたし|致し|し|申し上げ)ます/,
    label: { ja: '依頼を受けています', en: 'You are being asked to do something' },
  },
  {
    code: 'explicit_request',
    weight: 24,
    regex: /(いただけますと(幸い|助かり)|ご連絡(を)?(ください|お待ち)|お手数ですが|ご手配(を)?(ください|お願い)|送付(を)?(ください|お願い)|共有(を)?(ください|お願い))/,
    label: { ja: 'あなたからの連絡待ちです', en: 'They are waiting to hear from you' },
  },
];

const DEADLINE_PATTERNS: Array<{ regex: RegExp; label: string }> = [
  { regex: /本日中|今日中|by (end of day|eod|today)|today/i, label: '今日中' },
  { regex: /明日(まで|中)|by tomorrow/i, label: '明日まで' },
  { regex: /今週(中|末)まで|by (end of week|eow|friday)/i, label: '今週中' },
  { regex: /(\d{1,2})[\/月](\d{1,2})日?(まで|迄)/, label: '期限指定' },
  { regex: /by \d{1,2}[:：]\d{2}/i, label: '時刻指定' },
  { regex: /(至急|大至急|urgent|asap|as soon as possible)/i, label: '至急' },
  { regex: /締切|期限|deadline|due date/i, label: '締切あり' },
];

const AUTOMATED_SENDER = /(no[-\s]?reply|do[-\s]?not[-\s]?reply|donotreply|notifications?@|noreply@|mailer-daemon|postmaster|automated|alerts?@|newsletter|bounce)/i;

/**
 * Display names that only a system uses: "Contoso Alerting Engine",
 * "Build Notification Service". Applied to the name alone, never to a body,
 * so a human writing about an alerting engine is not filtered out.
 */
const AUTOMATED_SENDER_NAME =
  /(alert(ing)?\s+engine|notification\s+(service|centre|center|system)|automated?\s+(alerts?|notifications?|reports?)|\bdaemon\b|\bwebhook\b)/i;

/** Bulk/marketing senders: `team@marcom.example.com`, `info@`, `campaign@`… */
const MARKETING_SENDER = /(@|\.)(marcom|marketing|mailing|campaign|mktg|email|em|news|info)\.|^(marketing|campaigns?|webinars?|events?|sales|info|hello|contact)@/i;

const AUTOMATED_BODY =
  /(unsubscribe|配信停止|このメールは自動|自動送信|automatically generated|do not reply to this|本メールは送信専用|system notification|ticket (has been )?(created|updated)|build (succeeded|failed)|pipeline (succeeded|failed))/i;

/**
 * Outlook/Teams status notices such as "フォローしています: <会議>" or
 * "Following: <meeting>" — generated by the calendar, never by a human.
 */
const NOTIFICATION_SUBJECT =
  /^(フォローしています|辞退|仮承諾|承諾|転送されたメッセージ|配信不能|following|declined|accepted|tentative|canceled|cancelled|undeliverable|automatic reply|自動応答)\s*[:：]/i;

const NOTIFICATION_BODY =
  /(この会議をフォローしています|は、この会議を|さんがこの会議を|is following this meeting|has (accepted|declined|tentatively accepted) this meeting)/i;

const FYI_ONLY =
  /^(\s*)(fyi|ご参考|参考まで|共有です|情報共有|for your information|sharing|heads[- ]?up|周知|ご連絡まで|お知らせ)(です|のみ)?[\s。.!:：]*$/im;

const COMPLETION_MARKERS =
  /(ありがとうございま|助かりました|承知(いた)?しました|了解(です|しました)|対応(完了|済み)|クローズ|thanks[,.! ]|thank you|got it|sounds good|perfect|no further action|closing this)/i;

const REACTION_ONLY = /^(\s*)(\p{Emoji_Presentation}|\p{Extended_Pictographic}|👍|ok|ok!|了解|承知|thanks|thx|\+1|同意)[\s!！。.]*$/iu;

/** Teams/Outlook meeting-invite boilerplate — an invitation, not a request. */
const MEETING_INVITE_BODY =
  /(Microsoft Teams 会議\s*参加する|会議 ID\s*[:：]|Join the meeting now|Meeting ID\s*[:：]|_{20,})/;

function stripUrls(text: string): string {
  return text.replace(/https?:\/\/\S+/gi, ' ');
}

/** Opening salutation: "Hi Jaganathan,", "Dear Sato", "山田さん、", "平野 様". */
const SALUTATION_EN = /^\s*(?:hi|hello|hey|dear)\s+([\p{L}][\p{L}'’.-]*(?:\s+[\p{L}][\p{L}'’.-]*)?)\s*[,，、:：!！\n]/iu;
const SALUTATION_JA = /^\s*([\p{L}\p{N}]{1,12}(?:\s*[\p{L}\p{N}]{1,12})?)\s*(?:さん|さま|様|先生)\s*[,，、:：\n]/u;

/**
 * A message that opens by greeting somebody else ("Hi Jaganathan,") is that
 * person's to-do, even when it lands in a channel I follow.
 */
export function addressedToSomeoneElse(body: string, identity: Identity): boolean {
  const opening = body.trimStart().split(/\n/).slice(0, 2).join('\n');
  const match = SALUTATION_EN.exec(opening) ?? SALUTATION_JA.exec(opening);
  const named = match?.[1]?.trim();
  if (!named) return false;
  if (/^(all|team|everyone|folks|there|皆|みな|みんな|みなさん|皆さん|皆様|各位|関係者|全員)$/i.test(named)) return false;
  return !mentionsIdentity(named, identity);
}

/** "本日担当の〇〇さんに折り返し対応を依頼しました" — somebody else owns it. */
const DELEGATED_TO_OTHERS =
  /(担当の\s*\S{2,20}?\s*さん(に|へ).{0,16}(依頼|お願い)(し|いたし)ました|\S{2,20}さん(に|へ)(対応|折り返し|確認)を(依頼|お願い)(し|いたし)ました|assigned (this|it|the case) to \w+|handed (this|it) (over )?to \w+)/;

/* ------------------------------------------------------------------ *
 * Triage                                                              *
 * ------------------------------------------------------------------ */

export interface TriageDecision {
  needsReply: boolean;
  alreadyReplied: boolean;
  importance: Importance;
  score: number;
  reason: string;
  reasonEn: string;
  signals: TriageSignal[];
  deadlineText?: string;
  excludeReason?:
    | 'already_replied'
    | 'automated'
    | 'cc_only'
    | 'broadcast'
    | 'delegated'
    | 'fyi'
    | 'no_signal'
    | 'completed'
    | 'reaction';
}

export interface TriageOptions {
  /** VIP addresses get an importance boost (comma separated env var). */
  vipAddresses?: string[];
  minScore?: number;
}

export function triageThread(thread: ConversationThread, identity: Identity, options: TriageOptions = {}): TriageDecision {
  const alreadyReplied = hasAnsweredLatestInbound(thread);
  const pending = unansweredInbound(thread);
  const latest = pending[pending.length - 1] ?? thread.lastInboundMessage;

  if (!latest) {
    return decision(false, alreadyReplied, 'no_signal', [], 0, thread);
  }

  if (alreadyReplied) {
    return decision(false, true, 'already_replied', [], 0, thread);
  }

  if (isAutomated(latest)) {
    return decision(false, false, 'automated', [], 0, thread);
  }

  // A meeting invitation arriving as mail is handled by the calendar view.
  if (latest.source === 'outlook' && (latest.isEventMessage || MEETING_INVITE_BODY.test(latest.body))) {
    return decision(false, false, 'automated', [], 0, thread);
  }

  // A bot @-mentioning someone else is that person's task, not mine.
  if (latest.mentionsOthers && !latest.mentionsMe) {
    return decision(false, false, 'cc_only', [], 0, thread);
  }

  // "Hi Jaganathan, just following up…" — greeted by name, and it is not me.
  if (!latest.mentionsMe && addressedToSomeoneElse(latest.body, identity)) {
    return decision(false, false, 'cc_only', [], 0, thread);
  }

  // Reached me through a distribution list, so nobody addressed me by name.
  if (latest.isBroadcast && !latest.mentionsMe) {
    return decision(false, false, 'broadcast', [], 0, thread);
  }

  // A channel post is a broadcast by default. It becomes my task only when I am
  // @-mentioned or have already taken part in the thread; otherwise I am a
  // bystander watching two other people arrange something between themselves.
  if (thread.source === 'teams-channel' && !latest.mentionsMe && !thread.lastOutboundMessage) {
    return decision(false, false, 'broadcast', [], 0, thread);
  }

  const combinedText = pending.map((message) => `${message.subject}\n${message.body}`).join('\n');
  const latestText = `${latest.subject}\n${latest.body}`;

  if (REACTION_ONLY.test(latest.body.trim())) {
    return decision(false, false, 'reaction', [], 0, thread);
  }

  // The newest message decides first, so the reason we show always matches the
  // excerpt we show. Older unanswered messages are only a fallback.
  const latestSignals = detectSignals(latestText, latest, identity);

  // A pure FYI with no request signal is excluded even if unanswered.
  if (latestSignals.length === 0 && FYI_ONLY.test(latest.body.trim())) {
    return decision(false, false, 'fyi', [], 0, thread);
  }

  // Their last word closes the loop ("ありがとうございました") and asks nothing.
  if (latestSignals.length === 0 && COMPLETION_MARKERS.test(latest.body) && latest.body.length < 400) {
    return decision(false, false, 'completed', [], 0, thread);
  }

  // The work was handed to a named colleague, and that colleague is not me.
  if (DELEGATED_TO_OTHERS.test(latest.body) && !latest.mentionsMe) {
    return decision(false, false, 'delegated', [], 0, thread);
  }

  const usingFallback = latestSignals.length === 0;
  const signals = usingFallback ? detectSignals(combinedText, latest, identity) : latestSignals;

  // Being on CC only is somebody else's task unless I am named in the body.
  if (latest.isCcOnlyForMe && !latest.mentionsMe) {
    return decision(false, false, 'cc_only', [], 0, thread);
  }

  let score = signals.reduce((total, signal) => total + signal.weight, 0);
  // Evidence from an older message in the thread is weaker than a fresh ask.
  if (usingFallback && signals.length > 0) score -= 8;

  if (latest.mentionsMe) score += 12;
  if (latest.isCcOnlyForMe) score -= 6;
  const isDirectRecipient = latest.to.some((recipient) =>
    identity.addresses.some((address) => recipient.address?.toLowerCase() === address.toLowerCase()),
  );
  if (thread.source === 'outlook' && latest.to.length === 1 && isDirectRecipient) score += 8;
  if (thread.source === 'teams-chat' && thread.participants.length <= 3) score += 10;
  if (latest.importance === 'high') score += 8;
  if (options.vipAddresses?.some((vip) => latest.from.address?.toLowerCase() === vip.toLowerCase())) score += 14;
  if (isExternal(latest, identity)) score += 10;
  // Somebody restarted a thread that I previously participated in.
  if (thread.lastOutboundMessage && pending.length > 0) score += 8;

  const deadline = detectDeadline(usingFallback ? combinedText : latestText);
  if (deadline) score += deadline.urgent ? 20 : 10;

  const minScore = options.minScore ?? 22;
  // Boosts alone never make an action item: something must actually be asked.
  const needsReply = signals.length > 0 && score >= minScore;

  const result = decision(needsReply, false, needsReply ? undefined : 'no_signal', signals, score, thread);
  if (deadline) result.deadlineText = deadline.label;
  return result;
}

function decision(
  needsReply: boolean,
  alreadyReplied: boolean,
  excludeReason: TriageDecision['excludeReason'] | undefined,
  signals: TriageSignal[],
  score: number,
  thread: ConversationThread,
): TriageDecision {
  const importance = toImportance(score, signals);
  const reasons = buildReason(signals, thread);
  const result: TriageDecision = {
    needsReply,
    alreadyReplied,
    importance,
    score,
    reason: reasons.ja,
    reasonEn: reasons.en,
    signals,
  };
  if (excludeReason) result.excludeReason = excludeReason;
  return result;
}

function toImportance(score: number, signals: TriageSignal[]): Importance {
  const hasUrgentSignal = signals.some(
    (signal) => signal.code === 'approval_request' || signal.code === 'decision_required' || signal.code === 'deadline',
  );
  if (score >= 55 || (hasUrgentSignal && score >= 45)) return 'high';
  if (score >= 30) return 'medium';
  return 'low';
}

function buildReason(signals: TriageSignal[], thread: ConversationThread): { ja: string; en: string } {
  if (signals.length === 0) {
    return {
      ja: `${thread.lastInboundMessage?.from.name ?? '相手'}からの連絡に、まだあなたが返信していません。`,
      en: `You have not replied to ${thread.lastInboundMessage?.from.name ?? 'the sender'} yet.`,
    };
  }
  const top = [...signals].sort((a, b) => b.weight - a.weight).slice(0, 2);
  const label = (signal: TriageSignal, locale: 'ja' | 'en') =>
    signal.label?.[locale] ??
    REQUEST_PATTERNS.find((pattern) => pattern.code === signal.code)?.label[locale] ??
    (locale === 'ja' ? 'あなたの対応が必要です' : 'Your response is required');
  const evidence = top[0]?.evidence ? `「${truncate(top[0].evidence, 60)}」` : '';
  return {
    ja: `${evidence}${evidence ? 'のように、' : ''}${label(top[0]!, 'ja')}。`,
    en: `${label(top[0]!, 'en')}${top[0]?.evidence ? `: "${truncate(top[0].evidence, 60)}"` : ''}.`,
  };
}

export function detectSignals(text: string, latest: NormalizedMessage, identity: Identity): TriageSignal[] {
  const signals: TriageSignal[] = [];
  const seen = new Set<string>();
  // URLs carry query strings ("?p=..."), which are not questions to me.
  const scanned = stripUrls(text);

  for (const pattern of REQUEST_PATTERNS) {
    const match = pattern.regex.exec(scanned);
    if (!match) continue;
    const key = `${pattern.code}:${pattern.weight}`;
    if (seen.has(key)) continue;
    seen.add(key);
    signals.push({
      code: pattern.code,
      weight: pattern.weight,
      evidence: extractSentence(scanned, match.index),
      label: pattern.label,
    });
  }

  // A question mark aimed at me, when I am a direct recipient.
  if (/[?？]/.test(scanned) && !latest.isCcOnlyForMe && !signals.some((s) => s.code === 'direct_question')) {
    const index = scanned.search(/[?？]/);
    signals.push({ code: 'direct_question', weight: 18, evidence: extractSentence(scanned, index) });
  }

  // In email every greeting contains my name ("Hi Masaaki,"), so a name match
  // is only an action signal in Teams, where mentions are explicit.
  if (latest.mentionsMe && latest.source !== 'outlook' && !signals.some((s) => s.code === 'mentioned')) {
    signals.push({
      code: 'mentioned',
      weight: 14,
      evidence: `${identity.name} が名指しされています`,
    });
  }

  const deadline = detectDeadline(text);
  if (deadline) {
    signals.push({ code: 'deadline', weight: deadline.urgent ? 22 : 12, evidence: deadline.label });
  }

  return signals;
}

export function detectDeadline(text: string): { label: string; urgent: boolean } | undefined {
  for (const pattern of DEADLINE_PATTERNS) {
    const match = pattern.regex.exec(text);
    if (match) {
      const urgent = /至急|大至急|urgent|asap|本日中|今日中|eod|today/i.test(match[0]);
      return { label: match[0].trim() || pattern.label, urgent };
    }
  }
  return undefined;
}

function isAutomated(message: NormalizedMessage): boolean {
  if (AUTOMATED_SENDER.test(message.from.address ?? '')) return true;
  if (AUTOMATED_SENDER.test(message.from.name)) return true;
  if (AUTOMATED_SENDER_NAME.test(message.from.name)) return true;
  if (MARKETING_SENDER.test(message.from.address ?? '')) return true;
  if (isBotSender(message.from.name)) return true;
  if (AUTOMATED_BODY.test(message.body)) return true;
  if (NOTIFICATION_SUBJECT.test(message.subject.trim())) return true;
  if (NOTIFICATION_BODY.test(message.body)) return true;
  return false;
}

/**
 * Teams bots post as a single-token handle such as `azureavabot`, or as a
 * display name ending in "Bot". Human names keep their given/family parts.
 */
export function isBotSender(name: string): boolean {
  const trimmed = name.trim();
  if (!trimmed) return false;
  if (/\b(bot|assistant)\b/i.test(trimmed)) return true;
  return !/\s/.test(trimmed) && /bot$/i.test(trimmed);
}

function isExternal(message: NormalizedMessage, identity: Identity): boolean {
  const senderDomain = message.from.address?.split('@')[1]?.toLowerCase();
  if (!senderDomain) return false;
  const myDomains = identity.addresses.map((address) => address.split('@')[1]?.toLowerCase()).filter(Boolean);
  return !myDomains.includes(senderDomain);
}

function extractSentence(text: string, index: number): string {
  const start = Math.max(0, text.lastIndexOf('\n', index) + 1, text.lastIndexOf('。', index) + 1);
  const endCandidates = [text.indexOf('\n', index), text.indexOf('。', index), text.indexOf('. ', index)]
    .filter((position) => position > index)
    .sort((a, b) => a - b);
  const end = endCandidates[0] ?? Math.min(text.length, index + 120);
  return truncate(stripEmphasis(text.slice(start, end + 1)), 140);
}

/**
 * Business days a thread may sit unanswered before it is called out. Replying
 * within the same or next working day is normal; past this it is a backlog.
 */
export function staleAfterDays(): number {
  const raw = Number(process.env.HEY365_STALE_AFTER_DAYS);
  return Number.isFinite(raw) && raw > 0 ? Math.floor(raw) : 3;
}

/** Builds the user-facing action item from a thread and its triage decision. */
export function toActionItem(
  thread: ConversationThread,
  decisionResult: TriageDecision,
  index: number,
  now: Date = new Date(),
): ActionItem {
  const latest = thread.lastInboundMessage ?? thread.lastMessage;
  const arrived = new Date(latest.createdDateTime);
  const age = Number.isNaN(arrived.getTime()) ? 0 : businessDaysBetween(arrived, now);
  const stale = age >= staleAfterDays();

  const item: ActionItem = {
    index,
    id: shortHash(`${thread.source}:${thread.conversationId}:${latest.id}`),
    source: thread.source,
    conversationId: thread.conversationId,
    subject: thread.subject || truncate(latest.body, 60) || '(no subject)',
    sender: latest.from,
    participants: thread.participants,
    lastMessageTime: latest.createdDateTime,
    needsReply: decisionResult.needsReply,
    alreadyReplied: decisionResult.alreadyReplied,
    // Something left hanging for days outranks whatever the text scored.
    importance: stale ? 'high' : decisionResult.importance,
    score: decisionResult.score,
    reason: decisionResult.reason,
    reasonEn: decisionResult.reasonEn,
    signals: decisionResult.signals,
    summary: truncate(stripEmphasis(latest.body), 220),
    excerpt: truncate(stripEmphasis(latest.body), 600),
    language: thread.language,
    routing: latest.routing,
    ageBusinessDays: age,
    ...(stale ? { stale: true } : {}),
  };
  if (decisionResult.deadlineText) item.deadlineText = decisionResult.deadlineText;
  if (thread.webLink) item.webLink = thread.webLink;
  return item;
}
