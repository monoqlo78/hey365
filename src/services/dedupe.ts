import type { ActionItem } from '../models/action-item.js';
import { normalizeSubject, similarity } from '../utils/text.js';

const TIME_PROXIMITY_MS = 12 * 3600_000;

/**
 * Collapses the same real-world request when it arrives through several
 * channels (mail + Teams + meeting chat), per spec section 16.
 */
export function dedupeActionItems(items: ActionItem[]): ActionItem[] {
  const kept: ActionItem[] = [];

  for (const item of items) {
    const duplicateOf = kept.find((candidate) => isSameMatter(candidate, item));
    if (!duplicateOf) {
      kept.push(item);
      continue;
    }

    const merged = duplicateOf.mergedFrom ?? [];
    merged.push({ source: item.source, conversationId: item.conversationId, subject: item.subject });
    duplicateOf.mergedFrom = merged;

    // Keep the strongest signal set and the freshest timestamp.
    if (item.score > duplicateOf.score) {
      duplicateOf.score = item.score;
      duplicateOf.importance = item.importance;
      duplicateOf.reason = item.reason;
      duplicateOf.reasonEn = item.reasonEn;
      duplicateOf.signals = item.signals;
    }
    if (new Date(item.lastMessageTime).getTime() > new Date(duplicateOf.lastMessageTime).getTime()) {
      duplicateOf.lastMessageTime = item.lastMessageTime;
      duplicateOf.summary = item.summary;
      duplicateOf.excerpt = item.excerpt;
    }
    if (item.deadlineText && !duplicateOf.deadlineText) duplicateOf.deadlineText = item.deadlineText;
  }

  return kept;
}

export function isSameMatter(a: ActionItem, b: ActionItem): boolean {
  if (a.conversationId === b.conversationId && a.source === b.source) return true;

  const timeDelta = Math.abs(new Date(a.lastMessageTime).getTime() - new Date(b.lastMessageTime).getTime());
  if (timeDelta > TIME_PROXIMITY_MS) return false;

  const subjectA = normalizeSubject(a.subject);
  const subjectB = normalizeSubject(b.subject);
  const subjectScore = subjectA && subjectB ? similarity(subjectA, subjectB) : 0;

  const sameSender =
    Boolean(a.sender.address && b.sender.address && a.sender.address.toLowerCase() === b.sender.address.toLowerCase()) ||
    a.sender.name === b.sender.name;

  if (subjectScore >= 0.8 && sameSender) return true;
  if (subjectScore >= 0.9) return true;

  // Different channels, same people, near-identical content.
  if (sameSender && participantOverlap(a, b) >= 0.6) {
    const bodyScore = similarity(a.excerpt, b.excerpt);
    if (bodyScore >= 0.55) return true;
  }

  return false;
}

function participantOverlap(a: ActionItem, b: ActionItem): number {
  const keyOf = (participant: { address?: string; name: string }) =>
    (participant.address ?? participant.name).toLowerCase();
  const setA = new Set(a.participants.map(keyOf));
  const setB = new Set(b.participants.map(keyOf));
  if (setA.size === 0 || setB.size === 0) return 0;
  let intersection = 0;
  for (const key of setA) if (setB.has(key)) intersection += 1;
  return intersection / Math.min(setA.size, setB.size);
}
