// Turns the accessibility labels Outlook on the web puts on calendar event cards into
// events. Outlook renders every event's details into its aria-label, e.g.
//   "Weekly sync, 10:00 AM to 10:30 AM, Monday, October 13, 2026, Teams Meeting, Jane Doe, Busy"
// which is more stable than its generated CSS class names. English labels only.

export interface OutlookEvent {
  id: string;
  title: string;
  /** Start timestamp (ms), interpreted in this machine's time zone */
  start: number;
  /** Duration (ms) */
  duration: number;
  isAllDay: boolean;
  /** Remaining label text: location, organizer, status */
  details?: string;
}

const MONTHS = [
  'jan',
  'feb',
  'mar',
  'apr',
  'may',
  'jun',
  'jul',
  'aug',
  'sep',
  'oct',
  'nov',
  'dec',
];
const DAY_MS = 24 * 60 * 60 * 1000;

const TIME = String.raw`(\d{1,2}(?::\d{2})?\s?(?:[ap]\.?m\.?)?)`;
// "Monday, October 13, 2026", "October 13, 2026", "Mon, Oct 13, 2026", "13 October 2026"
const DATE = String.raw`((?:[a-z]+,?\s+)?(?:[a-z]+\.?\s+\d{1,2}|\d{1,2}\s+[a-z]+\.?),?\s+\d{4})`;
const RANGE = String.raw`\s*(?:to|-|–|—)\s*`;
const REST = String.raw`(.*)$`;

// Ordered from most to least specific: the lazy title must not swallow a pattern
// that a more specific one would match.
const MULTI_DAY_TIMED = new RegExp(
  String.raw`^(.+?),\s*${TIME},?\s+${DATE}${RANGE}${TIME},?\s+${DATE}${REST}`,
  'i',
);
const SAME_DAY_TIMED = new RegExp(
  String.raw`^(.+?),\s*${TIME}${RANGE}${TIME},\s*${DATE}${REST}`,
  'i',
);
const DATE_FIRST_TIMED = new RegExp(
  String.raw`^(.+?),\s*${DATE},?\s+${TIME}${RANGE}${TIME}${REST}`,
  'i',
);
const ALL_DAY = new RegExp(
  String.raw`^(.+?),\s*all[- ]day(?:\s+event)?,\s*${DATE}(?:${RANGE}${DATE})?${REST}`,
  'i',
);

const CANCELED = /^(?:canceled|cancelled)\s*:/i;

/** Midnight (local time) of a date like "Monday, October 13, 2026", or null. */
export const parseDate = (str: string): Date | null => {
  const s = str.trim().replace(/^(?:mon|tue|wed|thu|fri|sat|sun)[a-z]*\.?,?\s+/i, '');
  const mdy = /^([a-z]+)\.?\s+(\d{1,2}),?\s+(\d{4})$/i.exec(s);
  const dmy = /^(\d{1,2})\s+([a-z]+)\.?,?\s+(\d{4})$/i.exec(s);
  const [monthName, day, year] = mdy
    ? [mdy[1], mdy[2], mdy[3]]
    : dmy
      ? [dmy[2], dmy[1], dmy[3]]
      : [];
  if (!monthName) return null;
  const month = MONTHS.indexOf(monthName.slice(0, 3).toLowerCase());
  if (month < 0) return null;
  const date = new Date(Number(year), month, Number(day));
  return date.getDate() === Number(day) ? date : null;
};

/** Minutes after midnight for "9:00 AM", "9 AM", "21:30", or null. */
export const parseTime = (str: string): number | null => {
  const m = /^(\d{1,2})(?::(\d{2}))?\s?([ap])?\.?m?\.?$/i.exec(str.trim());
  if (!m) return null;
  let hours = Number(m[1]);
  const minutes = Number(m[2] ?? 0);
  const meridiem = m[3]?.toLowerCase();
  // A bare hour ("9") is not a time; it is how a slot index or count looks.
  if (!meridiem && m[2] === undefined) return null;
  if (meridiem) {
    if (hours < 1 || hours > 12) return null;
    hours = (hours % 12) + (meridiem === 'p' ? 12 : 0);
  }
  if (hours > 23 || minutes > 59) return null;
  return hours * 60 + minutes;
};

const at = (day: Date, minutes: number): number =>
  new Date(day.getFullYear(), day.getMonth(), day.getDate(), 0, minutes).getTime();

// djb2: short, stable ids without a crypto dependency.
const hash = (str: string): string => {
  let h = 5381;
  for (let i = 0; i < str.length; i++) h = ((h << 5) + h + str.charCodeAt(i)) | 0;
  return (h >>> 0).toString(36);
};

const cleanDetails = (rest: string): string | undefined => {
  const details = rest.replace(/^[\s,]+/, '').trim();
  return details || undefined;
};

const build = (
  title: string,
  start: number,
  end: number,
  isAllDay: boolean,
  rest: string,
): OutlookEvent | null => {
  const cleanTitle = title.trim();
  if (!cleanTitle || CANCELED.test(cleanTitle) || !(end > start)) return null;
  return {
    id: `owa-${start.toString(36)}-${hash(cleanTitle)}`,
    title: cleanTitle,
    start,
    duration: end - start,
    isAllDay,
    details: cleanDetails(rest),
  };
};

/** Parses one aria-label; returns null for labels that are not an event card. */
export const parseLabel = (rawLabel: string): OutlookEvent | null => {
  const label = rawLabel.replace(/\s+/g, ' ').trim();

  let m = MULTI_DAY_TIMED.exec(label);
  if (m) {
    const [, title, startTime, startDate, endTime, endDate, rest] = m;
    const sDay = parseDate(startDate);
    const eDay = parseDate(endDate);
    const sMin = parseTime(startTime);
    const eMin = parseTime(endTime);
    if (sDay && eDay && sMin !== null && eMin !== null) {
      return build(title, at(sDay, sMin), at(eDay, eMin), false, rest);
    }
  }

  const sameDay = SAME_DAY_TIMED.exec(label);
  const dateFirst = sameDay ? null : DATE_FIRST_TIMED.exec(label);
  const timed = sameDay
    ? { title: sameDay[1], startTime: sameDay[2], endTime: sameDay[3], dateStr: sameDay[4], rest: sameDay[5] }
    : dateFirst
      ? { title: dateFirst[1], dateStr: dateFirst[2], startTime: dateFirst[3], endTime: dateFirst[4], rest: dateFirst[5] }
      : null;
  if (timed) {
    const { title, startTime, endTime, dateStr, rest } = timed;
    const day = parseDate(dateStr);
    const sMin = parseTime(startTime);
    const eMin = parseTime(endTime);
    if (day && sMin !== null && eMin !== null) {
      // "11:00 PM to 12:30 AM" ends on the next day.
      const endMin = eMin <= sMin ? eMin + 24 * 60 : eMin;
      return build(title, at(day, sMin), at(day, endMin), false, rest);
    }
  }

  m = ALL_DAY.exec(label);
  if (m) {
    const [, title, startDate, endDate, rest] = m;
    const sDay = parseDate(startDate);
    const eDay = endDate ? parseDate(endDate) : sDay;
    if (sDay && eDay) {
      // Outlook names the last day of a multi-day all-day event, so the end is inclusive.
      return build(title, sDay.getTime(), eDay.getTime() + DAY_MS, true, rest);
    }
  }

  return null;
};

/** Parses all labels, dropping non-events and duplicates (Outlook can render a card twice). */
export const parseLabels = (labels: unknown[]): OutlookEvent[] => {
  const byId = new Map<string, OutlookEvent>();
  for (const label of labels) {
    if (typeof label !== 'string') continue;
    const event = parseLabel(label);
    if (event && !byId.has(event.id)) byId.set(event.id, event);
  }
  return [...byId.values()].sort((a, b) => a.start - b.start);
};
