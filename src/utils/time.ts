/**
 * Time helpers. Internally Hey365 always works in ISO 8601 UTC (spec section
 * 17); only presentation is converted to the user's local timezone.
 */

export const DEFAULT_TIMEZONE = 'Asia/Tokyo';

export function timezone(): string {
  return process.env.HEY365_TIMEZONE?.trim() || DEFAULT_TIMEZONE;
}

/** Graph `$filter` comparisons require a `Z`-suffixed timestamp without millis. */
export function toGraphTimestamp(date: Date): string {
  return `${date.toISOString().slice(0, 19)}Z`;
}

/** KQL `sent>=` comparisons in /search/query use plain dates. */
export function toKqlDate(date: Date): string {
  return date.toISOString().slice(0, 10);
}

export interface TimeWindow {
  hours: number;
  start: Date;
  end: Date;
  startIso: string;
  endIso: string;
}

/** "Past N hours" means `now - N` .. `now`, never a calendar day boundary. */
export function buildWindow(hours: number, now: Date = new Date()): TimeWindow {
  const safeHours = Number.isFinite(hours) && hours > 0 ? Math.min(hours, 24 * 30) : 36;
  const start = new Date(now.getTime() - safeHours * 3600_000);
  return {
    hours: safeHours,
    start,
    end: now,
    startIso: start.toISOString(),
    endIso: now.toISOString(),
  };
}

export function parseDate(value: string | null | undefined): Date | undefined {
  if (!value) return undefined;
  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? undefined : parsed;
}

/** Formats an instant as `HH:mm` in the display timezone. */
export function formatTime(iso: string, tz: string = timezone()): string {
  const date = parseDate(iso);
  if (!date) return '';
  return new Intl.DateTimeFormat('en-GB', {
    timeZone: tz,
    hour: '2-digit',
    minute: '2-digit',
    hour12: false,
  }).format(date);
}

/** Formats an instant as `YYYY-MM-DD HH:mm` in the display timezone. */
export function formatDateTime(iso: string, tz: string = timezone()): string {
  const date = parseDate(iso);
  if (!date) return '';
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: tz,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    hour12: false,
  }).formatToParts(date);
  const get = (type: string) => parts.find((p) => p.type === type)?.value ?? '';
  return `${get('year')}-${get('month')}-${get('day')} ${get('hour')}:${get('minute')}`;
}

/** Human-friendly "3時間前" / "3h ago" style label. */
export function relativeLabel(iso: string, locale: 'ja' | 'en' = 'ja', now: Date = new Date()): string {
  const date = parseDate(iso);
  if (!date) return '';
  const minutes = Math.max(0, Math.round((now.getTime() - date.getTime()) / 60_000));
  if (locale === 'ja') {
    if (minutes < 60) return `${minutes}分前`;
    const hours = Math.round(minutes / 60);
    if (hours < 24) return `${hours}時間前`;
    return `${Math.round(hours / 24)}日前`;
  }
  if (minutes < 60) return `${minutes}m ago`;
  const hours = Math.round(minutes / 60);
  if (hours < 24) return `${hours}h ago`;
  return `${Math.round(hours / 24)}d ago`;
}

export function isWithin(iso: string | undefined, window: TimeWindow): boolean {
  const date = parseDate(iso);
  if (!date) return false;
  return date.getTime() >= window.start.getTime() && date.getTime() <= window.end.getTime() + 60_000;
}
