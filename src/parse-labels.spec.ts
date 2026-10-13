import { describe, expect, it } from 'vitest';
import { parseDate, parseLabel, parseLabels, parseTime } from './parse-labels';

const local = (y: number, mo: number, d: number, h = 0, mi = 0): number =>
  new Date(y, mo - 1, d, h, mi).getTime();
const MIN = 60 * 1000;

describe('parseTime', () => {
  it('parses 12h and 24h times', () => {
    expect(parseTime('9:00 AM')).toBe(9 * 60);
    expect(parseTime('12:00 AM')).toBe(0);
    expect(parseTime('12:30 PM')).toBe(12 * 60 + 30);
    expect(parseTime('9 PM')).toBe(21 * 60);
    expect(parseTime('9:15pm')).toBe(21 * 60 + 15);
    expect(parseTime('21:30')).toBe(21 * 60 + 30);
  });

  it('rejects bare numbers and impossible times', () => {
    expect(parseTime('9')).toBeNull();
    expect(parseTime('13:00 PM')).toBeNull();
    expect(parseTime('24:00')).toBeNull();
  });
});

describe('parseDate', () => {
  it('parses month-first and day-first dates with or without a weekday', () => {
    expect(parseDate('Monday, October 13, 2026')?.getTime()).toBe(local(2026, 10, 13));
    expect(parseDate('Oct 13, 2026')?.getTime()).toBe(local(2026, 10, 13));
    expect(parseDate('Monday 13 October 2026')?.getTime()).toBe(local(2026, 10, 13));
  });

  it('rejects unknown months and impossible days', () => {
    expect(parseDate('Smarch 13, 2026')).toBeNull();
    expect(parseDate('February 30, 2026')).toBeNull();
  });
});

describe('parseLabel', () => {
  it('parses a same-day meeting and keeps the rest as details', () => {
    const event = parseLabel(
      'Weekly sync, 10:00 AM to 10:30 AM, Monday, October 13, 2026, Microsoft Teams Meeting, Jane Doe, Busy, Recurring event',
    );
    expect(event).toMatchObject({
      title: 'Weekly sync',
      start: local(2026, 10, 13, 10),
      duration: 30 * MIN,
      isAllDay: false,
      details: 'Microsoft Teams Meeting, Jane Doe, Busy, Recurring event',
    });
  });

  it('keeps commas inside titles', () => {
    expect(
      parseLabel('Plan Q4: budget, hiring, 2:00 PM to 3:00 PM, Tuesday, October 14, 2026')
        ?.title,
    ).toBe('Plan Q4: budget, hiring');
  });

  it('parses 24h labels', () => {
    expect(
      parseLabel('Standup, 09:00 to 09:15, Wednesday, October 15, 2026'),
    ).toMatchObject({ start: local(2026, 10, 15, 9), duration: 15 * MIN });
  });

  it('parses date-first labels', () => {
    expect(
      parseLabel('1:1 with Sam, Thursday, October 16, 2026 3:00 PM to 3:30 PM, Room 4'),
    ).toMatchObject({
      title: '1:1 with Sam',
      start: local(2026, 10, 16, 15),
      duration: 30 * MIN,
      details: 'Room 4',
    });
  });

  it('rolls a meeting that ends after midnight into the next day', () => {
    expect(
      parseLabel('Release, 11:00 PM to 12:30 AM, Friday, October 17, 2026'),
    ).toMatchObject({ start: local(2026, 10, 17, 23), duration: 90 * MIN });
  });

  it('parses timed events spanning several days', () => {
    expect(
      parseLabel(
        'Offsite, 9:00 AM Monday, October 13, 2026 to 5:00 PM Tuesday, October 14, 2026, Berlin',
      ),
    ).toMatchObject({
      title: 'Offsite',
      start: local(2026, 10, 13, 9),
      duration: local(2026, 10, 14, 17) - local(2026, 10, 13, 9),
      isAllDay: false,
      details: 'Berlin',
    });
  });

  it('parses single and multi-day all-day events with an inclusive end day', () => {
    expect(
      parseLabel('Company holiday, all day event, Monday, October 13, 2026, Free'),
    ).toMatchObject({
      start: local(2026, 10, 13),
      duration: local(2026, 10, 14) - local(2026, 10, 13),
      isAllDay: true,
      details: 'Free',
    });
    expect(
      parseLabel(
        'Conference, all day, Monday, October 13, 2026 to Wednesday, October 15, 2026',
      ),
    ).toMatchObject({
      start: local(2026, 10, 13),
      duration: local(2026, 10, 16) - local(2026, 10, 13),
    });
  });

  it('ignores labels that are not event cards', () => {
    expect(parseLabel('Monday, October 13, 2026')).toBeNull();
    expect(parseLabel('9:00 AM, Monday, October 13, 2026')).toBeNull();
    expect(parseLabel('New event')).toBeNull();
  });

  it('skips canceled meetings', () => {
    expect(
      parseLabel('Canceled: Weekly sync, 10:00 AM to 10:30 AM, Monday, October 13, 2026'),
    ).toBeNull();
  });

  it('gives the same event the same id across syncs', () => {
    const label = 'Weekly sync, 10:00 AM to 10:30 AM, Monday, October 13, 2026';
    expect(parseLabel(label)?.id).toBe(parseLabel(`${label}, Busy`)?.id);
  });
});

describe('label format captured from outlook.live.com (2026-10, titles replaced)', () => {
  it('parses event cards and skips the untitled slot label', () => {
    const events = parseLabels([
      'Design review, 10:30 AM to 1:00 PM, Tuesday, October 13, 2026, Busy',
      'Sprint planning, 9:00 AM to 12:00 PM, Wednesday, October 14, 2026, Busy',
      '5:00 PM to 5:30 PM, Tuesday, October 13, 2026',
    ]);
    expect(events).toEqual([
      expect.objectContaining({
        title: 'Design review',
        start: local(2026, 10, 13, 10, 30),
        duration: 150 * MIN,
        details: 'Busy',
      }),
      expect.objectContaining({
        title: 'Sprint planning',
        start: local(2026, 10, 14, 9),
        duration: 180 * MIN,
      }),
    ]);
  });
});

describe('parseLabels', () => {
  it('drops duplicates and non-strings, and sorts by start', () => {
    const late = 'B, 2:00 PM to 3:00 PM, Monday, October 13, 2026';
    const early = 'A, 9:00 AM to 10:00 AM, Monday, October 13, 2026';
    const events = parseLabels([late, early, late, 42, 'Monday, October 13, 2026']);
    expect(events.map((e) => e.title)).toEqual(['A', 'B']);
  });
});
