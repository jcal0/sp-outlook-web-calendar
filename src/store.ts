import type { OutlookEvent } from './parse-labels';

// Meetings are work data: keep them in this device's localStorage (plugin.js runs in
// the host renderer) instead of persistDataSynced, so they never leave the device.
const KEY = 'sp-plugin.outlook-web-calendar.events.v1';

export interface StoredWeek {
  syncedAt: number;
  events: OutlookEvent[];
}

export const loadWeek = (): StoredWeek | null => {
  try {
    const raw = localStorage.getItem(KEY);
    const parsed = raw ? (JSON.parse(raw) as StoredWeek) : null;
    return parsed && Array.isArray(parsed.events) ? parsed : null;
  } catch {
    return null;
  }
};

/** Each sync replaces the previous one: the captured week is the source of truth. */
export const saveWeek = (events: OutlookEvent[]): void => {
  try {
    localStorage.setItem(KEY, JSON.stringify({ syncedAt: Date.now(), events }));
  } catch {
    // Storage full or blocked: the next sync tries again.
  }
};
