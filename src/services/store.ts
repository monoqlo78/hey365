import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { homedir } from 'node:os';
import { dirname, join } from 'node:path';

import type { ActionItem, TriageResult } from '../models/action-item.js';
import type { SessionSummary } from '../models/session.js';
import { logger } from '../utils/logger.js';

/**
 * Draft store.
 *
 * `hey365` only ever produces drafts; `hey365_send` later resolves the numbers
 * the user typed ("1と3を送って") back to concrete conversations. The state is
 * persisted so a client restart between triage and send does not lose context.
 */
export interface StoredState {
  version: 1;
  updatedAt: string;
  triage?: TriageResult;
  sessions: Record<string, SessionSummary>;
  sent: Array<{ id: string; index: number; sentAt: string; target: string }>;
  /** Conversations the user chose to hide, keyed by conversation id. */
  muted: MuteEntry[];
  /** Last Work IQ account that authenticated, used for silent reconnects. */
  lastAccount?: string;
}

export interface MuteEntry {
  conversationId: string;
  subject: string;
  kind: 'snoozed' | 'done';
  /** Hidden only while a newer message has not arrived than this one. */
  lastMessageTime: string;
  /** For `snoozed`, the instant the item comes back by itself. */
  until?: string;
  note?: string;
  createdAt: string;
}

const EMPTY: StoredState = {
  version: 1,
  updatedAt: new Date(0).toISOString(),
  sessions: {},
  sent: [],
  muted: [],
};

function statePath(): string {
  const override = process.env.HEY365_STATE_FILE?.trim();
  if (override) return override;
  return join(process.env.HEY365_HOME?.trim() || join(homedir(), '.hey365'), 'state.json');
}

let cache: StoredState | undefined;

export function loadState(): StoredState {
  if (cache) return cache;
  try {
    const raw = readFileSync(statePath(), 'utf8');
    const parsed = JSON.parse(raw) as StoredState;
    cache = {
      ...EMPTY,
      ...parsed,
      sessions: parsed.sessions ?? {},
      sent: parsed.sent ?? [],
      muted: parsed.muted ?? [],
    };
  } catch {
    cache = { ...EMPTY, sessions: {}, sent: [], muted: [] };
  }
  return cache;
}

export function saveState(state: StoredState): void {
  cache = { ...state, updatedAt: new Date().toISOString() };
  const path = statePath();
  try {
    mkdirSync(dirname(path), { recursive: true });
    writeFileSync(path, JSON.stringify(cache, null, 2), { encoding: 'utf8', mode: 0o600 });
  } catch (error) {
    logger.warn('could not persist Hey365 state', { error: (error as Error).message });
  }
}

export function storeTriage(result: TriageResult): void {
  const state = loadState();
  state.triage = result;
  saveState(state);
}

export function getTriage(): TriageResult | undefined {
  return loadState().triage;
}

/**
 * Remembers which account last authenticated so a silent reconnect can pass
 * `--account` and pick the right cached token on multi-account machines.
 */
export function rememberAccount(account: string | undefined): void {
  const trimmed = account?.trim();
  if (!trimmed) return;
  const state = loadState();
  if (state.lastAccount === trimmed) return;
  state.lastAccount = trimmed;
  saveState(state);
}

export function lastKnownAccount(): string | undefined {
  return loadState().lastAccount;
}

export function storeSession(summary: SessionSummary): void {
  const state = loadState();
  state.sessions[summary.sessionId] = summary;
  saveState(state);
}

export function getSession(sessionId: string): SessionSummary | undefined {
  return loadState().sessions[sessionId];
}

/** Resolves user input such as `["1","3"]`, `["all"]` or `["全部"]`. */
export function resolveItems(selectors: string[], options: { requireDraft?: boolean } = {}): ActionItem[] {
  const triage = getTriage();
  if (!triage) return [];
  const requireDraft = options.requireDraft !== false;
  const wantsAll = selectors.some((selector) => /^(all|全部|すべて|全て)$/i.test(selector.trim()));
  if (wantsAll) return requireDraft ? triage.items.filter((item) => item.draft) : triage.items;

  const wanted = new Set(selectors.map((selector) => selector.trim()));
  return triage.items.filter((item) => wanted.has(String(item.index)) || wanted.has(item.id));
}

export function updateItemDraft(index: number, draft: ActionItem['draft']): ActionItem | undefined {
  const state = loadState();
  const item = state.triage?.items.find((entry) => entry.index === index);
  if (!item) return undefined;
  item.draft = draft;
  saveState(state);
  return item;
}

export function recordSend(item: ActionItem, target: string): void {
  const state = loadState();
  state.sent.unshift({ id: item.id, index: item.index, sentAt: new Date().toISOString(), target });
  state.sent = state.sent.slice(0, 100);
  saveState(state);
}

export function resetStateCache(): void {
  cache = undefined;
}

/* ------------------------------------------------------------------ *
 * Snooze / done                                                       *
 * ------------------------------------------------------------------ */

/**
 * Hides a conversation from future triage runs.
 *
 * The mute is pinned to the message that was on screen when the user dismissed
 * it. Anything newer re-opens the item, because "I have dealt with this" was
 * only ever a statement about what the user had already read.
 */
export function muteConversation(entry: Omit<MuteEntry, 'createdAt'>): MuteEntry {
  const state = loadState();
  const stored: MuteEntry = { ...entry, createdAt: new Date().toISOString() };
  state.muted = [stored, ...state.muted.filter((mute) => mute.conversationId !== entry.conversationId)].slice(0, 500);
  saveState(state);
  return stored;
}

export function unmuteConversation(conversationId: string): boolean {
  const state = loadState();
  const before = state.muted.length;
  state.muted = state.muted.filter((mute) => mute.conversationId !== conversationId);
  if (state.muted.length === before) return false;
  saveState(state);
  return true;
}

export function listMutes(now: Date = new Date()): MuteEntry[] {
  return loadState().muted.filter((mute) => !isExpired(mute, now));
}

function isExpired(mute: MuteEntry, now: Date): boolean {
  if (!mute.until) return false;
  const until = new Date(mute.until);
  return !Number.isNaN(until.getTime()) && until.getTime() <= now.getTime();
}

/** The active mute for a conversation, or undefined when it should be shown. */
export function findMute(
  conversationId: string,
  lastMessageTime: string,
  now: Date = new Date(),
): MuteEntry | undefined {
  const mute = loadState().muted.find((entry) => entry.conversationId === conversationId);
  if (!mute || isExpired(mute, now)) return undefined;

  const seen = new Date(mute.lastMessageTime).getTime();
  const latest = new Date(lastMessageTime).getTime();
  // A reply arrived after the user dismissed it, so this is a new matter.
  if (Number.isFinite(seen) && Number.isFinite(latest) && latest > seen) return undefined;
  return mute;
}

/** Drops mutes that have expired or been overtaken by a newer message. */
export function pruneMutes(now: Date = new Date()): number {
  const state = loadState();
  const before = state.muted.length;
  state.muted = state.muted.filter((mute) => !isExpired(mute, now));
  if (state.muted.length === before) return 0;
  saveState(state);
  return before - state.muted.length;
}
