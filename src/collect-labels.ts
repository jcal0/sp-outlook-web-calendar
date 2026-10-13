export interface CollectedLabels {
  url: string;
  isLogin: boolean;
  isCalendar: boolean;
  labels: string[];
}

/**
 * Runs inside the Outlook on the web page (Chrome extension content script, and the
 * bookmarklet via `toString()`). Collects the raw aria-labels that carry a date and a
 * time; parsing happens in the plugin, where it can be tested.
 *
 * Must stay self-contained (no outer references): the bookmarklet serializes it.
 */
export function collectLabels(): CollectedLabels {
  const year = /\b(19|20)\d{2}\b/;
  const timeOrAllDay = /\d{1,2}:\d{2}|\d\s?[ap]\.?m\b|all[- ]day/i;
  const labels: string[] = [];
  const seen: Record<string, true> = {};
  const nodes = document.querySelectorAll('[aria-label]');
  for (let i = 0; i < nodes.length; i++) {
    const label = (nodes[i].getAttribute('aria-label') || '').replace(/\s+/g, ' ').trim();
    if (
      label.length > 15 &&
      label.length < 1500 &&
      !seen[label] &&
      year.test(label) &&
      timeOrAllDay.test(label)
    ) {
      seen[label] = true;
      labels.push(label);
    }
  }
  return {
    url: location.origin + location.pathname,
    isLogin:
      /(^|\.)login\.|microsoftonline/.test(location.hostname) ||
      !!document.querySelector('input[type=password], input[name=loginfmt]'),
    isCalendar: /\/calendar/.test(location.pathname),
    labels,
  };
}
