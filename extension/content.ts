import { collectLabels } from '../src/collect-labels';

declare const chrome: {
  runtime: { sendMessage(message: unknown): Promise<unknown> };
};

// Outlook renders the week incrementally: send once the labels have held still for
// a few polls. Only watch for a minute after each load or in-app navigation (e.g.
// switching weeks), so an Outlook tab left open does not keep scanning the page.
const POLL_MS = 1500;
const WATCH_MS = 60000;
const MIN_ON_CALENDAR_MS = 6000;

let watchUntil = Date.now() + WATCH_MS;
let lastHref = location.href;
let calendarSince = 0;
let lastCount = -1;
let stablePolls = 0;
let lastSent = '';

const tick = (): void => {
  if (location.href !== lastHref) {
    lastHref = location.href;
    watchUntil = Date.now() + WATCH_MS;
    calendarSince = 0;
  }
  // The extension runs on all of Outlook (Mail too), so check the cheap URL first.
  if (Date.now() > watchUntil || !/\/calendar/.test(location.pathname)) return;

  const state = collectLabels();
  if (state.isLogin || !state.isCalendar) {
    calendarSince = 0;
    return;
  }
  if (!calendarSince) calendarSince = Date.now();
  stablePolls = state.labels.length === lastCount ? stablePolls + 1 : 0;
  lastCount = state.labels.length;

  const key = state.labels.join('\n');
  if (stablePolls >= 2 && Date.now() - calendarSince >= MIN_ON_CALENDAR_MS && key !== lastSent) {
    lastSent = key;
    chrome.runtime
      .sendMessage({ type: 'outlook-week', url: state.url, labels: state.labels })
      .catch(() => undefined);
  }
};

setInterval(tick, POLL_MS);
