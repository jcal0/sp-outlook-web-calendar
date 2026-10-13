# Outlook Web Calendar for Super Productivity

Shows your Outlook on the web meetings for the week in the
[Super Productivity](https://super-productivity.com) schedule and planner, as calendar
events that block time (not tasks).

It's for accounts where you can't export the calendar: no "Publish calendar" ICS link
and no API access. Instead it reads the week from the Outlook page in your own,
already signed-in Chrome.

> If your Outlook lets you **publish your calendar** (Settings → Calendar → Shared
> calendars → Publish a calendar), you don't need this: paste the ICS link into
> Super Productivity's built-in "Outlook 365" calendar instead.

## How it works

1. You click **Sync Outlook week** in Super Productivity.
2. The plugin starts a short-lived listener on `127.0.0.1:39217` and opens your Outlook
   calendar week in Chrome.
3. The companion Chrome extension waits for the week to render, reads each meeting from
   its accessibility label (title, time, date, location) and sends the list to that listener.
4. The plugin stores the meetings **on this device only** (never synced or exported),
   and Super Productivity shows them in the schedule.

The listener only accepts requests carrying a custom header. Web pages can't send that
header without permission (a CORS preflight), and the listener never grants it, so only
the extension can post to it. The listener stops after one sync or about four minutes.

## Requirements

- Super Productivity **desktop app** (macOS, Windows or Linux), v19 or newer. Syncing
  needs the desktop app because the plugin runs a small Node.js script.
- Google Chrome, with permission to load an unpacked extension (Developer mode).
- Outlook on the web in **English**.

## Install

Download both zips from the [latest release](../../releases/latest).

**Plugin**

1. In Super Productivity, open **Settings → Plugins → Choose Plugin File** and pick
   `outlook-web-calendar-plugin-vX.Y.Z.zip`.
2. In the plugin's settings, set **Calendar week URL**:
   - work or school accounts: `https://outlook.office.com/calendar/view/week`
     (or `/view/workweek`)
   - personal accounts: `https://outlook.live.com/calendar/0/view/week`
3. Open the **issue provider panel** (press `P`, or click the webhook icon in the header).
   Under **Connect Calendar**, choose **Outlook Web Calendar** (not "Outlook 365") and
   save. Leave auto-import as tasks off if you want meetings to stay calendar events.

**Chrome extension**

1. Unzip `outlook-web-calendar-chrome-extension-vX.Y.Z.zip` to a folder you'll keep;
   Chrome loads the extension from that folder.
2. Open `chrome://extensions`, turn on **Developer mode** and click **Load unpacked**.
   Select the unzipped folder.

## Use

Click **Sync Outlook week** in the header. You can also bind a key to it in Settings →
Keyboard, under Plugin Shortcuts. The first sync asks you to allow Node.js execution
for the plugin.

Each sync replaces the stored meetings with the week shown in Chrome. Click a meeting in
the schedule to turn it into a task if you want to track time on it.

### If Chrome extensions are blocked (e.g. a managed work laptop)

Use the bookmarklet instead:

1. In the plugin menu, open **Outlook bookmarklet** and copy the code into a new bookmark's URL.
2. On your Outlook calendar week, click the bookmark. It copies the week to the clipboard.
3. In Super Productivity, open **Import Outlook meetings** from the plugin menu and paste.

## Limitations

- Only reads the week currently shown in Outlook.
- English Outlook labels only. Times are read in your computer's time zone, so keep it
  the same as Outlook's.
- It depends on how Outlook labels event cards. If a Microsoft update changes that,
  sync may find no meetings until the parser is updated. Please open an issue with an
  example label (with the meeting title replaced).

## Development

```bash
npm install
npm test            # parser tests
npm run typecheck
npm run build       # dist/plugin, dist/extension and both release zips
npm run try-sync -- https://outlook.live.com/calendar/0/view/week
                    # runs the listener without Super Productivity and prints what the
                    # extension sends (load dist/extension in Chrome first)
```

To release, bump `version` in `package.json`, commit, then tag and push:
`git tag v0.2.0 && git push --tags`. The Release workflow builds and attaches both zips.

## License

MIT
