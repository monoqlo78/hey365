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
}

const EMPTY: StoredState = { version: 1, updatedAt: new Date(0).toISOString(), sessions: {}, sent: [] };

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
    cache = { ...EMPTY, ...parsed, sessions: parsed.sessions ?? {}, sent: parsed.sent ?? [] };
  } catch {
    cache = { ...EMPTY, sessions: {}, sent: [] };
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

export function storeSession(summary: SessionSummary): void {
  const state = loadState();
  state.sessions[summary.sessionId] = summary;
  saveState(state);
}

export function getSession(sessionId: string): SessionSummary | undefined {
  return loadState().sessions[sessionId];
}

/** Resolves user input such as `["1","3"]`, `["all"]` or `["全部"]`. */
export function resolveItems(selectors: string[]): ActionItem[] {
  const triage = getTriage();
  if (!triage) return [];
  const wantsAll = selectors.some((selector) => /^(all|全部|すべて|全て)$/i.test(selector.trim()));
  if (wantsAll) return triage.items.filter((item) => item.draft);

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
