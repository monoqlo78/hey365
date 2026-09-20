import type { ReplyDraft } from '../models/action-item.js';
import type { ConversationThread } from '../models/conversation.js';
import { renderTranscript, unansweredInbound } from './conversation.js';
import { ask } from './workiq.js';
import { logger } from '../utils/logger.js';
import { detectLanguage, shortHash, truncate, type Language } from '../utils/text.js';

export interface DraftContext {
  thread: ConversationThread;
  myName: string;
  /** Extra instruction such as "もう少し柔らかく" or "make it English". */
  instruction?: string;
  /** Force a language instead of mirroring the conversation. */
  language?: Language;
}

const AI_TELLS = [
  /^(certainly|sure(,| thing)|of course|i hope this (email|message) finds you well)[!,.]?\s*/i,
  /^(承知いたしました[。、]\s*)(?=.)/,
  /^(お世話になっております。\s*){2,}/,
  /\bas an ai\b.*$/i,
  /^(はい、?承知しました。?\s*)(?=.{20,})/,
];

const UNKNOWN_MARKERS =
  /(判断が必要|判断できません|情報が不足|わかりません|不明です|cannot determine|a decision is required|not enough information|i (don'?t|do not) (know|have))/i;

/**
 * Generates a reply draft. Microsoft 365 Copilot (through `workiq ask`) does
 * the writing so Hey365 needs no extra model credentials; if that is
 * unavailable we fall back to a conservative template that never invents facts.
 */
export async function generateDraft(context: DraftContext): Promise<ReplyDraft> {
  const language = context.language ?? context.thread.language;
  const prompt = buildPrompt(context, language);

  try {
    const response = await ask(prompt);
    const cleaned = cleanDraft(response.text, language);
    if (cleaned && cleaned.length >= 8) {
      const needsDecision = UNKNOWN_MARKERS.test(cleaned) ? missingDecisionNote(context.thread, language) : undefined;
      return draft(cleaned, language, 'workiq-ask', needsDecision);
    }
    logger.debug('ask returned an unusable draft, falling back to template');
  } catch (error) {
    logger.debug('ask failed, falling back to template', { error: (error as Error).message });
  }

  return templateDraft(context, language);
}

/** Applies a natural-language edit ("softer", "shorter", "in English"). */
export async function refineDraft(
  current: ReplyDraft,
  instruction: string,
  context: { thread?: ConversationThread; myName: string },
): Promise<ReplyDraft> {
  const language = inferInstructionLanguage(instruction, current.language);
  const prompt = [
    'You are editing a draft reply that the user will send from their own account.',
    'Rewrite the draft according to the instruction. Output ONLY the rewritten reply body.',
    'Do not add commentary, headings, quotation marks, or a signature block.',
    `Write the reply in ${language === 'ja' ? 'Japanese' : 'English'}.`,
    '',
    `INSTRUCTION: ${instruction}`,
    '',
    'CURRENT DRAFT:',
    current.text,
    context.thread ? `\nCONVERSATION CONTEXT:\n${renderTranscript(context.thread, 4, 500)}` : '',
  ].join('\n');

  try {
    const response = await ask(prompt);
    const cleaned = cleanDraft(response.text, language);
    if (cleaned && cleaned.length >= 5) return draft(cleaned, language, 'workiq-ask');
  } catch (error) {
    logger.debug('refine failed', { error: (error as Error).message });
  }

  // Deterministic fallbacks so the tool still does something useful offline.
  const local = applyLocalRefinement(current.text, instruction);
  return draft(local, language, 'user-edited');
}

export function draft(
  text: string,
  language: Language,
  generator: ReplyDraft['generator'],
  needsUserDecision?: string,
): ReplyDraft {
  const trimmed = text.trim();
  return {
    text: trimmed,
    language,
    hash: shortHash(trimmed),
    generator,
    ...(needsUserDecision ? { needsUserDecision } : {}),
    updatedAt: new Date().toISOString(),
  };
}

function buildPrompt(context: DraftContext, language: Language): string {
  const { thread, myName } = context;
  const pending = unansweredInbound(thread);
  const channel = thread.source === 'outlook' ? 'Outlook email' : 'Microsoft Teams message';
  const styleRules =
    thread.source === 'outlook'
      ? language === 'ja'
        ? '- メールなので簡潔な挨拶（例: お世話になっております）を1行だけ入れてよい\n- 署名は書かない'
        : '- A one-line greeting is fine for email\n- Do not add a signature block'
      : language === 'ja'
        ? '- Teams なのでメールより短く、挨拶は不要\n- 2〜4文程度'
        : '- Teams chat: shorter than email, no greeting, 1-3 sentences';

  return [
    `You are drafting a reply that ${myName} will send in a ${channel}.`,
    'Output ONLY the reply body text. No subject line, no headings, no explanation, no quotes.',
    `Write in ${language === 'ja' ? 'Japanese' : 'English'} (match the conversation language).`,
    'Rules:',
    '- Answer the sender\'s actual question directly, in the first sentence if possible.',
    '- Be concise and professional; do not sound like generated boilerplate.',
    '- Never invent facts, dates, numbers, commitments or approvals that are not in the conversation.',
    `- If the conversation does not contain enough information to answer, reply with exactly: ${
      language === 'ja' ? '「返信するには判断が必要です」' : '"A decision is required before replying"'
    } followed by what is missing.`,
    styleRules,
    context.instruction ? `- Additional instruction: ${context.instruction}` : '',
    '',
    `SUBJECT: ${thread.subject}`,
    `PARTICIPANTS: ${thread.participants.map((p) => p.name).join(', ')}`,
    `MESSAGES AWAITING MY RESPONSE: ${pending.length}`,
    '',
    'CONVERSATION (oldest first, "ME" is the user writing the reply):',
    renderTranscript(thread, 8, 800),
  ]
    .filter(Boolean)
    .join('\n');
}

function templateDraft(context: DraftContext, language: Language): ReplyDraft {
  const { thread } = context;
  const pending = unansweredInbound(thread);
  const latest = pending[pending.length - 1] ?? thread.lastInboundMessage ?? thread.lastMessage;
  const topic = truncate(latest.body, 80);
  const note = missingDecisionNote(thread, language);

  const text =
    language === 'ja'
      ? [
          `${latest.from.name} さん`,
          '',
          `ご連絡ありがとうございます。「${topic}」の件、確認のうえ改めてご連絡します。`,
        ].join('\n')
      : [
          `Hi ${latest.from.name.split(' ')[0] ?? latest.from.name},`,
          '',
          `Thanks for the note about "${topic}". I will check and get back to you shortly.`,
        ].join('\n');

  return draft(text, language, 'template', note);
}

function missingDecisionNote(thread: ConversationThread, language: Language): string {
  const subject = thread.subject || '本件';
  return language === 'ja'
    ? `返信するには「${truncate(subject, 40)}」についてあなたの判断が必要です。`
    : `A decision from you is required about "${truncate(subject, 40)}" before this can be answered.`;
}

/** Removes model preamble, code fences and AI tells from a generated reply. */
export function cleanDraft(raw: string, language: Language): string {
  let text = raw.trim();

  text = text.replace(/^```[a-z]*\s*/i, '').replace(/```\s*$/i, '');
  text = text.replace(/^(?:draft|reply|返信案|以下が返信案です)[:：]\s*/i, '');
  text = text.replace(/^(?:here(?:'s| is) (?:a|the) (?:draft|reply)[^\n]*\n)/i, '');
  text = text.replace(/^(?:以下のように返信[^\n]*\n)/, '');
  text = text.replace(/^["'「](.*)["'」]$/s, '$1');
  // The model sometimes quotes the required phrase verbatim from the prompt.
  text = text.replace(/^「(返信するには判断が必要です)」/, '$1');
  text = text.replace(/^"(A decision is required before replying)"/i, '$1');

  for (const tell of AI_TELLS) text = text.replace(tell, '');

  // Strip trailing meta-commentary the model sometimes appends.
  text = text.replace(/\n+(?:この返信案は|Let me know if you|ご希望に応じて|必要に応じて調整)[^\n]*$/i, '');

  if (language === 'ja') text = text.replace(/\n{3,}/g, '\n\n');
  return text.trim();
}

function inferInstructionLanguage(instruction: string, fallback: Language): Language {
  if (/英語|english|in english/i.test(instruction)) return 'en';
  if (/日本語|japanese|in japanese/i.test(instruction)) return 'ja';
  return fallback ?? detectLanguage(instruction);
}

/** Offline approximations of the most common refinement requests. */
function applyLocalRefinement(text: string, instruction: string): string {
  if (/短く|簡潔|shorter|brief/i.test(instruction)) {
    const sentences = text.split(/(?<=[。.!?！？])\s*/).filter(Boolean);
    return sentences.slice(0, Math.max(1, Math.ceil(sentences.length / 2))).join('');
  }
  if (/丁寧|polite|formal/i.test(instruction)) {
    return text.replace(/です。/g, 'でございます。').replace(/ます。/g, 'ます。');
  }
  return text;
}
