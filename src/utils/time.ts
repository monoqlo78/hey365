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
  /** Non-business days that the window skipped over, as `YYYY-MM-DD`. */
  skippedDays: string[];
}

/* ------------------------------------------------------------------ *
 * Business calendar                                                   *
 * ------------------------------------------------------------------ */

function tzOffsetMs(instant: number, tz: string): number {
  const parts = new Intl.DateTimeFormat('en-US', {
    timeZone: tz,
    hour12: false,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
  }).formatToParts(new Date(instant));
  const get = (type: string) => Number(parts.find((p) => p.type === type)?.value ?? '0');
  // Some ICU builds render midnight as hour 24 of the previous date; Date.UTC
  // rolls that over correctly, so the value must not be wrapped with `% 24`.
  const wall = Date.UTC(get('year'), get('month') - 1, get('day'), get('hour'), get('minute'), get('second'));
  // The formatter drops milliseconds, so round to the minute — every real UTC
  // offset is a whole number of minutes.
  return Math.round((wall - instant) / 60_000) * 60_000;
}

interface CalendarDay {
  year: number;
  month: number;
  day: number;
  weekday: number;
  /** Instant of local midnight that starts this day. */
  startsAt: number;
}

function calendarDay(instant: number, tz: string): CalendarDay {
  const offset = tzOffsetMs(instant, tz);
  const local = new Date(instant + offset);
  const year = local.getUTCFullYear();
  const month = local.getUTCMonth() + 1;
  const day = local.getUTCDate();
  const midnightLocal = Date.UTC(year, month - 1, day);
  const startsAt = midnightLocal - tzOffsetMs(midnightLocal - offset, tz);
  return { year, month, day, weekday: local.getUTCDay(), startsAt };
}

function isoDate({ year, month, day }: CalendarDay): string {
  return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
}

function nthWeekday(year: number, month: number, weekday: number, nth: number): number {
  const first = new Date(Date.UTC(year, month - 1, 1)).getUTCDay();
  return 1 + ((weekday - first + 7) % 7) + (nth - 1) * 7;
}

/** Valid for 1980–2099 (Japanese National Astronomical Observatory approximation). */
function equinoxDay(year: number, base: number): number {
  return Math.floor(base + 0.242194 * (year - 1980) - Math.floor((year - 1980) / 4));
}

/** Japanese public holidays, including 振替休日 and 国民の休日. */
export function japaneseHolidays(year: number): Set<string> {
  const fixed: Array<[number, number]> = [
    [1, 1],
    [2, 11],
    [2, 23],
    [4, 29],
    [5, 3],
    [5, 4],
    [5, 5],
    [8, 11],
    [11, 3],
    [11, 23],
  ];
  const dates = new Set<string>();
  const add = (month: number, day: number) =>
    dates.add(`${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`);

  for (const [month, day] of fixed) add(month, day);
  add(1, nthWeekday(year, 1, 1, 2)); // 成人の日
  add(7, nthWeekday(year, 7, 1, 3)); // 海の日
  add(9, nthWeekday(year, 9, 1, 3)); // 敬老の日
  add(10, nthWeekday(year, 10, 1, 2)); // スポーツの日
  add(3, equinoxDay(year, 20.8431)); // 春分の日
  add(9, equinoxDay(year, 23.2488)); // 秋分の日

  // 振替休日: a holiday landing on Sunday moves to the next free weekday.
  for (const iso of [...dates]) {
    const [y, m, d] = iso.split('-').map(Number);
    const date = new Date(Date.UTC(y!, m! - 1, d!));
    if (date.getUTCDay() !== 0) continue;
    do {
      date.setUTCDate(date.getUTCDate() + 1);
    } while (dates.has(date.toISOString().slice(0, 10)));
    dates.add(date.toISOString().slice(0, 10));
  }

  // 国民の休日: a weekday sandwiched between two holidays (e.g. Silver Week).
  for (const iso of [...dates]) {
    const [y, m, d] = iso.split('-').map(Number);
    const gap = new Date(Date.UTC(y!, m! - 1, d! + 1));
    const after = new Date(Date.UTC(y!, m! - 1, d! + 2));
    if (gap.getUTCDay() === 0 || gap.getUTCDay() === 6) continue;
    if (dates.has(gap.toISOString().slice(0, 10))) continue;
    if (dates.has(after.toISOString().slice(0, 10))) dates.add(gap.toISOString().slice(0, 10));
  }

  return dates;
}

const holidayCache = new Map<number, Set<string>>();

function holidaysFor(year: number): Set<string> {
  let cached = holidayCache.get(year);
  if (!cached) {
    cached = japaneseHolidays(year);
    holidayCache.set(year, cached);
  }
  return cached;
}

function extraHolidays(): Set<string> {
  return new Set(
    (process.env.HEY365_HOLIDAYS ?? '')
      .split(',')
      .map((value) => value.trim())
      .filter((value) => /^\d{4}-\d{2}-\d{2}$/.test(value)),
  );
}

/** Whether the business calendar is applied at all (`HEY365_BUSINESS_DAYS=off` disables it). */
export function businessDaysEnabled(): boolean {
  const raw = (process.env.HEY365_BUSINESS_DAYS ?? '').trim().toLowerCase();
  return !['off', 'false', '0', 'no'].includes(raw);
}

function useJapaneseCalendar(tz: string): boolean {
  const override = (process.env.HEY365_HOLIDAY_CALENDAR ?? '').trim().toLowerCase();
  if (override === 'jp' || override === 'ja') return true;
  if (override === 'none' || override === 'off') return false;
  return tz === 'Asia/Tokyo';
}

/** Weekend or public holiday in the display timezone. */
export function isBusinessDay(day: CalendarDay, tz: string): boolean {
  if (day.weekday === 0 || day.weekday === 6) return false;
  const iso = isoDate(day);
  if (extraHolidays().has(iso)) return false;
  if (useJapaneseCalendar(tz) && holidaysFor(day.year).has(iso)) return false;
  return true;
}

/**
 * "Past N hours" counts working time only: weekends and public holidays are
 * stepped over rather than consumed, so a Monday morning check still reaches
 * back into the previous working week. The resulting range stays contiguous,
 * so messages that did arrive over the weekend are still picked up.
 */
export function buildWindow(hours: number, now: Date = new Date(), tz: string = timezone()): TimeWindow {
  const safeHours = Number.isFinite(hours) && hours > 0 ? Math.min(hours, 24 * 30) : 36;
  const end = now;
  const skippedDays: string[] = [];

  if (!businessDaysEnabled()) {
    const start = new Date(end.getTime() - safeHours * 3600_000);
    return { hours: safeHours, start, end, startIso: start.toISOString(), endIso: end.toISOString(), skippedDays };
  }

  let remaining = safeHours * 3600_000;
  let cursor = end.getTime();
  let day = calendarDay(cursor, tz);
  let startMs: number | undefined;

  for (let guard = 0; guard < 90; guard += 1) {
    if (isBusinessDay(day, tz)) {
      const available = cursor - day.startsAt;
      if (remaining <= available) {
        startMs = cursor - remaining;
        break;
      }
      remaining -= available;
    } else if (cursor > day.startsAt) {
      skippedDays.push(isoDate(day));
    }
    cursor = day.startsAt;
    day = calendarDay(cursor - 1, tz);
  }

  const start = new Date(startMs ?? end.getTime() - safeHours * 3600_000);
  return {
    hours: safeHours,
    start,
    end,
    startIso: start.toISOString(),
    endIso: end.toISOString(),
    skippedDays,
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
