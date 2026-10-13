import type { HostPluginAPI, PluginIssue, PluginSearchResult } from './host-api';
import { collectLabels } from './collect-labels';
import { LISTEN_SCRIPT, ListenResult, SYNC_PORT } from './listen-script';
import { OutlookEvent, parseLabels } from './parse-labels';
import { loadWeek, saveWeek } from './store';

declare const PluginAPI: HostPluginAPI;

const DEFAULT_CALENDAR_URL = 'https://outlook.office.com/calendar/view/week';
// The executor caps scripts at 5 minutes; leave room to start and report back.
const WAIT_MS = 4 * 60 * 1000;
const EXEC_TIMEOUT_MS = WAIT_MS + 30 * 1000;
const IMPORT_TEXTAREA_ID = 'outlook-web-calendar-import';

const getCalendarUrl = async (): Promise<string> => {
  const cfg = await PluginAPI.getConfig<{ calendarUrl?: string }>().catch(() => null);
  const url = cfg?.calendarUrl?.trim();
  return url && /^https:\/\//.test(url) ? url : DEFAULT_CALENDAR_URL;
};

const toSearchResult = (e: OutlookEvent): PluginSearchResult => ({
  id: e.id,
  title: e.title,
  start: e.start,
  duration: e.duration,
  isAllDay: e.isAllDay,
  // Timed meetings become tasks scheduled at their start; all-day ones get a due day.
  dueWithTime: e.isAllDay ? undefined : e.start,
  description: e.details,
});

const formatWhen = (e: OutlookEvent): string =>
  e.isAllDay
    ? PluginAPI.formatDate(e.start, 'medium')
    : `${PluginAPI.formatDate(e.start, 'datetime')} – ${PluginAPI.formatDate(e.start + e.duration, 'time')}`;

const storeEvents = (labels: unknown[], source: string): number => {
  const events = parseLabels(labels);
  saveWeek(events);
  // Only counts: meeting titles are work content and the log is exportable.
  PluginAPI.log.info(`outlook-web-calendar: ${source} stored ${events.length} events`);
  return events.length;
};

const syncedMsg = (count: number): string =>
  count
    ? `Synced ${count} Outlook meeting${count === 1 ? '' : 's'}. Not in the schedule after 30 seconds? Connect "Outlook Web Calendar" in the issue provider panel (P).`
    : 'Outlook sync finished, but no meetings were found on the page.';

// --- Sync via the Chrome extension ---

const SYNC_ERRORS: Record<string, string> = {
  PORT_IN_USE: `Another sync is running (or port ${SYNC_PORT} is in use).`,
  TIMEOUT:
    'No calendar received from Chrome. Is the "Super Productivity Outlook Sync" extension installed and the week view open?',
  SERVER_ERROR: 'Could not start the local sync listener.',
};

let isSyncing = false;

const syncWeek = async (): Promise<void> => {
  if (isSyncing) return;
  if (!PluginAPI.executeNodeScript) {
    PluginAPI.showSnack({
      msg: 'Outlook sync needs the desktop app. Use "Import Outlook meetings" instead.',
      type: 'WARNING',
    });
    return;
  }
  isSyncing = true;
  try {
    PluginAPI.showSnack({ msg: 'Opening Outlook in Chrome…', type: 'INFO', ico: 'sync' });
    const res = await PluginAPI.executeNodeScript({
      script: LISTEN_SCRIPT,
      args: [{ calendarUrl: await getCalendarUrl(), port: SYNC_PORT, waitMs: WAIT_MS }],
      timeout: EXEC_TIMEOUT_MS,
    });
    const result = res.result as ListenResult | undefined;
    if (!res.success || !result) {
      PluginAPI.showSnack({
        msg: `Outlook sync failed: ${res.error ?? 'no result'}`,
        type: 'ERROR',
      });
      return;
    }
    if (!result.ok) {
      PluginAPI.showSnack({ msg: SYNC_ERRORS[result.error] ?? result.error, type: 'ERROR' });
      return;
    }
    PluginAPI.showSnack({ msg: syncedMsg(storeEvents(result.labels, 'sync')), type: 'SUCCESS' });
  } finally {
    isSyncing = false;
  }
};

// --- Fallback: bookmarklet + paste ---

const escapeHtml = (s: string): string =>
  s.replace(/[&<>"']/g, (c) => `&#${c.charCodeAt(0)};`);

const buildBookmarklet = (): string => {
  const body = `(() => {
  const d = (${collectLabels.toString()})();
  const text = JSON.stringify({ type: 'outlook-web-calendar', url: d.url, labels: d.labels });
  navigator.clipboard.writeText(text).then(
    () => alert('Copied ' + d.labels.length + ' calendar entries. Paste them in Super Productivity: Import Outlook meetings.'),
    () => prompt('Copy this and paste it in Super Productivity:', text),
  );
})()`;
  return `javascript:${encodeURIComponent(body)}`;
};

/** Accepts the bookmarklet's JSON, a JSON array of labels, or one label per line. */
const labelsFromPaste = (text: string): string[] => {
  try {
    const data = JSON.parse(text) as unknown;
    if (Array.isArray(data)) return data.filter((l): l is string => typeof l === 'string');
    const labels = (data as { labels?: unknown })?.labels;
    if (Array.isArray(labels)) return labels.filter((l): l is string => typeof l === 'string');
  } catch {
    // not JSON: fall through to lines
  }
  return text.split(/\r?\n/);
};

const openImportDialog = async (): Promise<void> => {
  await PluginAPI.openDialog({
    title: 'Import Outlook meetings',
    htmlContent: `
      <p>On your Outlook calendar week, click the bookmarklet, then paste here.</p>
      <textarea id="${IMPORT_TEXTAREA_ID}" rows="8" style="width: 100%"></textarea>`,
    buttons: [
      { label: 'Cancel' },
      {
        label: 'Import',
        color: 'primary',
        raised: true,
        onClick: () => {
          // plugin.js runs in the host page, so it can read the dialog's textarea.
          const field = document.getElementById(IMPORT_TEXTAREA_ID) as HTMLTextAreaElement | null;
          const count = storeEvents(labelsFromPaste(field?.value ?? ''), 'import');
          PluginAPI.showSnack({ msg: syncedMsg(count), type: count ? 'SUCCESS' : 'WARNING' });
        },
      },
    ],
  });
};

const openBookmarkletDialog = async (): Promise<void> => {
  await PluginAPI.openDialog({
    title: 'Outlook bookmarklet',
    htmlContent: `
      <p>For when the Chrome extension cannot be installed. Create a bookmark and paste
      this as its URL. Click it on your Outlook calendar week, then use
      <strong>Import Outlook meetings</strong>.</p>
      <textarea readonly rows="6" style="width: 100%">${escapeHtml(buildBookmarklet())}</textarea>`,
    buttons: [{ label: 'Close' }],
  });
};

// --- Registration ---

PluginAPI.registerHeaderButton({
  label: 'Sync Outlook week',
  icon: 'event_repeat',
  onClick: () => void syncWeek(),
});
PluginAPI.registerMenuEntry({
  label: 'Import Outlook meetings',
  icon: 'content_paste',
  onClick: () => void openImportDialog(),
});
PluginAPI.registerMenuEntry({
  label: 'Outlook bookmarklet',
  icon: 'bookmark',
  onClick: () => void openBookmarkletDialog(),
});
PluginAPI.registerShortcut({
  id: 'sync-outlook-week',
  label: 'Sync Outlook week',
  onExec: () => void syncWeek(),
});

const storedEvents = (): OutlookEvent[] => loadWeek()?.events ?? [];

PluginAPI.registerIssueProvider({
  configFields: [],
  getHeaders: () => ({}),

  async searchIssues(searchTerm: string): Promise<PluginSearchResult[]> {
    const term = searchTerm.trim().toLowerCase();
    return storedEvents()
      .filter((e) => !term || e.title.toLowerCase().includes(term))
      .map(toSearchResult);
  },

  async getById(issueId: string): Promise<PluginIssue> {
    const event = storedEvents().find((e) => e.id === issueId);
    if (!event) {
      // Outside the last synced week (or moved): keep the task as it is.
      return { id: issueId, title: '', state: 'confirmed' };
    }
    return {
      id: event.id,
      title: event.title,
      body: event.details,
      state: 'confirmed',
      when: formatWhen(event),
      details: event.details,
    };
  },

  getIssueLink: () => DEFAULT_CALENDAR_URL,

  async getNewIssuesForBacklog(): Promise<PluginSearchResult[]> {
    return storedEvents().map(toSearchResult);
  },

  issueDisplay: [
    { field: 'title', label: 'Title', type: 'text' },
    { field: 'when', label: 'When', type: 'text' },
    { field: 'details', label: 'Details', type: 'text', hideEmpty: true },
  ],
});
