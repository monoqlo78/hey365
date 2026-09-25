import type { ActionItem, TriageResult } from '../models/action-item.js';
import { collect, triageCollected, type TriageRunOptions } from './collector.js';
import type { GraphEvent } from './graph.js';
import { listCalendar } from './graph.js';
import { localDate, localDayBounds, parseAsOf, timezone } from '../utils/time.js';

export interface DigestMeeting {
  subject: string;
  start: string;
  end?: string;
  organizer?: string;
  joinUrl?: string;
  /** True when the user has not accepted or declined yet. */
  needsResponse: boolean;
}

export interface DigestResult {
  /** Local date the digest describes, as `YYYY-MM-DD`. */
  date: string;
  timezone: string;
  triage: TriageResult;
  meetings: DigestMeeting[];
  /** Items whose detected deadline falls on or before the digest date. */
  dueToday: ActionItem[];
  /** Items that have gone unanswered past the stale threshold. */
  stale: ActionItem[];
}

export interface DigestOptions extends TriageRunOptions {
  /** Business days of mail/Teams history to include. Defaults to 3. */
  businessDays?: number;
}

/**
 * One screen for the start of the day: today's meetings, anything whose
 * deadline has arrived, and the replies still owed. Everything comes from a
 * single collect() pass so the digest costs the same as one triage run.
 */
export async function buildDigest(options: DigestOptions = {}): Promise<DigestResult> {
  const tz = timezone();
  const reference = (options.asOf ? parseAsOf(options.asOf, tz) : undefined) ?? options.now ?? new Date();
  const date = localDate(reference, tz);

  const context = await collect({
    ...options,
    businessDays: options.businessDays ?? 3,
    // The triage window's calendar read is capped and ordered from its start,
    // so a multi-day window pushes today's meetings past the limit. Read the
    // digest day directly instead.
    includeMeetings: false,
  });
  const triage = triageCollected(context, options);

  const bounds = localDayBounds(date, tz);
  const events = await listCalendar(bounds.start, bounds.end, 50).catch((error: unknown) => {
    triage.warnings.push(`カレンダーを取得できませんでした: ${(error as Error).message}`);
    return [] as GraphEvent[];
  });

  const meetings = events
    .filter((event) => !event.isCancelled && isOnDay(event, date, tz))
    .map(toDigestMeeting)
    .sort((a, b) => a.start.localeCompare(b.start));

  const dueToday = triage.items.filter((item) => isDueBy(item, date, tz));
  const stale = triage.items.filter((item) => item.stale);

  return { date, timezone: tz, triage, meetings, dueToday, stale };
}

function eventStart(event: GraphEvent): string | undefined {
  const raw = event.start?.dateTime;
  if (!raw) return undefined;
  // Graph returns calendar times without a zone suffix; they are UTC unless the
  // request asked otherwise, which Hey365 never does.
  return /[Zz]|[+-]\d{2}:\d{2}$/.test(raw) ? raw : `${raw}Z`;
}

function eventEnd(event: GraphEvent): string | undefined {
  const raw = event.end?.dateTime;
  if (!raw) return undefined;
  return /[Zz]|[+-]\d{2}:\d{2}$/.test(raw) ? raw : `${raw}Z`;
}

function isOnDay(event: GraphEvent, date: string, tz: string): boolean {
  const start = eventStart(event);
  if (!start) return false;
  const parsed = new Date(start);
  if (Number.isNaN(parsed.getTime())) return false;
  return localDate(parsed, tz) === date;
}

function toDigestMeeting(event: GraphEvent): DigestMeeting {
  const response = event.responseStatus?.response ?? 'none';
  return {
    subject: event.subject?.trim() || '(件名なし)',
    start: eventStart(event) ?? '',
    ...(eventEnd(event) ? { end: eventEnd(event) as string } : {}),
    ...(event.organizer?.emailAddress?.name ? { organizer: event.organizer.emailAddress.name } : {}),
    ...(event.onlineMeeting?.joinUrl ? { joinUrl: event.onlineMeeting.joinUrl } : {}),
    needsResponse: response === 'none' || response === 'notResponded',
  };
}

function isDueBy(item: ActionItem, date: string, tz: string): boolean {
  if (!item.deadline) return false;
  const parsed = new Date(item.deadline);
  if (Number.isNaN(parsed.getTime())) return false;
  return localDate(parsed, tz) <= date;
}
