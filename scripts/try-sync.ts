// Dev-only: runs the listener script the way the plugin's executeNodeScript call
// does, then prints the parsed events. Checks the extension against a real
// Outlook page without installing the plugin.
//   npm run try-sync -- [calendarUrl] [--raw out.json]
import { writeFileSync } from 'fs';
import { LISTEN_SCRIPT, ListenResult, SYNC_PORT } from '../src/listen-script';
import { parseLabels } from '../src/parse-labels';

const argv = process.argv.slice(2);
const rawIdx = argv.indexOf('--raw');
const rawOut = rawIdx >= 0 ? argv.splice(rawIdx, 2)[1] : undefined;
const calendarUrl = argv[0] || 'https://outlook.office.com/calendar/view/week';

const AsyncFunction = Object.getPrototypeOf(async () => {}).constructor;
const run = new AsyncFunction('args', 'require', LISTEN_SCRIPT);

(async () => {
  console.log(`Listening on 127.0.0.1:${SYNC_PORT}, opening ${calendarUrl}...`);
  const res: ListenResult = await run(
    [{ calendarUrl, port: SYNC_PORT, waitMs: 120000 }],
    require,
  );
  if (!res.ok) {
    console.error('Sync failed:', res.error);
    process.exit(1);
  }
  if (rawOut) writeFileSync(rawOut, JSON.stringify(res.labels, null, 2));
  const events = parseLabels(res.labels);
  console.log(`From ${res.url}: ${res.labels.length} candidate labels, ${events.length} events:`);
  for (const e of events) {
    const when = new Date(e.start).toLocaleString();
    console.log(`  ${when}  ${Math.round(e.duration / 60000)}m${e.isAllDay ? ' (all day)' : ''}  ${e.title}`);
  }
})();
