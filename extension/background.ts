import { SYNC_HEADER, SYNC_PATH, SYNC_PORT } from '../src/sync-protocol';

declare const chrome: {
  runtime: {
    onMessage: {
      addListener(fn: (message: { type?: string; url?: string; labels?: string[] }) => void): void;
    };
  };
};

// The content script cannot reach 127.0.0.1 itself (Outlook's CSP and CORS); the
// service worker can, through the extension's host permission. Super Productivity
// only listens while a sync runs, so a failed request just means nobody asked.
chrome.runtime.onMessage.addListener((message) => {
  if (message?.type !== 'outlook-week') return;
  fetch(`http://127.0.0.1:${SYNC_PORT}${SYNC_PATH}`, {
    method: 'POST',
    headers: { 'content-type': 'application/json', [SYNC_HEADER]: '1' },
    body: JSON.stringify({ url: message.url, labels: message.labels }),
  }).catch(() => undefined);
});
