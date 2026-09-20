import type { ActionItem, TriageResult } from '../models/action-item.js';
import type { SessionSummary } from '../models/session.js';
import { formatDateTime, formatTime, relativeLabel } from '../utils/time.js';

const MARKERS: Record<ActionItem['importance'], string> = {
  high: '🔴',
  medium: '🟠',
  low: '🟡',
};

const SOURCE_LABEL: Record<ActionItem['source'], string> = {
  outlook: 'Outlook',
  'teams-chat': 'Teams',
  'teams-channel': 'Teams Channel',
  meeting: 'Meeting',
};

const DIVIDER = '------------------------------------------------';

/** Renders the triage result exactly in the shape described by spec §8. */
export function formatTriage(result: TriageResult): string {
  const lines: string[] = ['Hey365 チェック完了', ''];

  if (result.items.length === 0) {
    lines.push(
      `過去${result.windowHours}時間で、あなたの対応が必要そうなものは見つかりませんでした。`,
      '',
      `スキャン: メール ${result.scanned.outlookMessages}件 / Teams ${result.scanned.teamsMessages}件 / 会議 ${result.scanned.meetings}件`,
      `除外: 返信済み ${result.skipped.alreadyReplied} / 自動通知 ${result.skipped.automated} / CCのみ ${result.skipped.ccOnly} / FYI ${result.skipped.fyi}`,
    );
    appendWarnings(lines, result.warnings);
    return lines.join('\n');
  }

  lines.push(`過去${result.windowHours}時間で、あなたの対応が必要と思われるものが ${result.items.length}件あります。`, '');

  result.items.forEach((item, position) => {
    lines.push(...formatItem(item, result.timezone));
    if (position < result.items.length - 1) lines.push('', DIVIDER, '');
  });

  lines.push(
    '',
    DIVIDER,
    '',
    '送信したいものがあれば、',
    '',
    '「1を送って」',
    '「1と3を送って」',
    '「全部送って」',
    '',
    'のように指示してください。',
    '',
    '文章を修正する場合は、',
    '「1をもう少し柔らかく」',
    'のように指示してください。',
  );

  appendWarnings(lines, result.warnings);
  return lines.join('\n');
}

function formatItem(item: ActionItem, tz: string): string[] {
  const lines: string[] = [];
  const time = formatTime(item.lastMessageTime, tz);
  const relative = relativeLabel(item.lastMessageTime);
  lines.push(`${MARKERS[item.importance]} ${item.index}. ${item.sender.name} / ${SOURCE_LABEL[item.source]}`);
  lines.push(`${time}（${relative}）`);
  lines.push('');
  lines.push('内容:');
  const subject = item.subject?.trim() ?? '';
  const summary = item.summary?.trim() ?? '';
  if (subject) lines.push(subject);
  // Teams items often carry the first body line as their subject; printing both
  // would show the same sentence twice.
  if (summary && !isSameText(summary, subject)) lines.push(summary);
  lines.push('');
  lines.push('返信が必要な理由:');
  lines.push(item.reason);

  if (item.deadlineText) {
    lines.push('');
    lines.push(`期限: ${item.deadlineText}`);
  }

  if (item.mergedFrom?.length) {
    lines.push('');
    lines.push(`関連: ${item.mergedFrom.map((entry) => SOURCE_LABEL[entry.source]).join(' / ')} でも同じ案件が届いています。`);
  }

  if (item.draft) {
    lines.push('');
    if (item.draft.needsUserDecision) {
      lines.push('返信案:');
      lines.push(`⚠ ${item.draft.needsUserDecision}`);
      lines.push('');
      lines.push(item.draft.text);
    } else {
      lines.push('返信案:');
      lines.push(item.draft.text);
    }
  }

  return lines.filter((line, index, all) => !(line === '' && all[index - 1] === ''));
}

/** True when one string is the other, or a truncated prefix of it. */
function isSameText(a: string, b: string): boolean {
  if (!a || !b) return false;
  const normalize = (value: string) => value.replace(/[\s…]+/g, '').replace(/\.{3}$/, '');
  const left = normalize(a);
  const right = normalize(b);
  if (!left || !right) return false;
  return left === right || left.startsWith(right) || right.startsWith(left);
}

function appendWarnings(lines: string[], warnings: string[]): void {  if (warnings.length === 0) return;
  lines.push('', DIVIDER, '', '注意:');
  for (const warning of warnings) lines.push(`- ${warning}`);
}

/** Renders a session summary in the shape described by spec §12. */
export function formatSessionSummary(summary: SessionSummary, tz: string): string {
  if (summary.candidates?.length) {
    const lines = ['Session の候補が複数あります。どれを要約しますか？', ''];
    summary.candidates.forEach((candidate, index) => {
      lines.push(`${index + 1}. ${candidate.title}`);
      if (candidate.when) lines.push(`   ${formatDateTime(candidate.when, tz)}`);
      lines.push(`   id: ${candidate.id}`);
    });
    return lines.join('\n');
  }

  const lines: string[] = ['Session Summary', ''];
  lines.push('会議:', summary.title, '');

  if (summary.startTime) {
    lines.push('日時:', formatDateTime(summary.startTime, tz), '');
  }

  if (summary.participants.length > 0) {
    lines.push('参加者:');
    for (const participant of summary.participants.slice(0, 20)) lines.push(participant.name);
    lines.push('');
  }

  lines.push('■ 要約', '', summary.summary || '(要約を生成できませんでした)', '');

  if (summary.decisions.length > 0) {
    lines.push('■ 決定事項', '');
    summary.decisions.forEach((decision, index) => lines.push(`${index + 1}. ${decision}`));
    lines.push('');
  }

  if (summary.keyPoints.length > 0) {
    lines.push('■ 重要ポイント', '');
    for (const point of summary.keyPoints) lines.push(`- ${point}`);
    lines.push('');
  }

  if (summary.actionItems.length > 0) {
    lines.push('■ Action Items', '');
    const mine = summary.actionItems.filter((item) => item.isMine);
    const others = summary.actionItems.filter((item) => !item.isMine);
    const byOwner = new Map<string, string[]>();
    for (const item of others) {
      const bucket = byOwner.get(item.owner) ?? [];
      bucket.push(item.due ? `${item.text}（${item.due}）` : item.text);
      byOwner.set(item.owner, bucket);
    }
    for (const [owner, texts] of byOwner) {
      lines.push(owner);
      for (const text of texts) lines.push(`- ${text}`);
      lines.push('');
    }
    if (mine.length > 0) {
      lines.push('私');
      for (const item of mine) lines.push(`- ${item.due ? `${item.text}（${item.due}）` : item.text}`);
      lines.push('');
    }
  }

  if (summary.openQuestions.length > 0) {
    lines.push('■ 未解決', '');
    for (const question of summary.openQuestions) lines.push(question);
    lines.push('');
  }

  lines.push('■ 私からの返信', '', summary.replyNeeded ? '必要' : '不要', '');
  if (summary.replyReason) lines.push('理由:', summary.replyReason, '');
  if (summary.draft) {
    lines.push('返信案:', '');
    if (summary.draft.needsUserDecision) lines.push(`⚠ ${summary.draft.needsUserDecision}`, '');
    lines.push(summary.draft.text);
  }

  return lines.join('\n').replace(/\n{3,}/g, '\n\n').trim();
}
