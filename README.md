# Outlook Extension PoC

A proof of concept for an Outlook add-in (Office.js, React + Vite + TypeScript) that works while the user composes a message:

- **Ribbon button and side panel** in the compose window ("Outlook Extension PoC", under Apps in Outlook on the web).
- **Items** from a test endpoint: checked items are inserted as one block at the top of the message body. Changing the selection replaces the block instead of adding another one.
- **List of emails** from a test endpoint: clicking an address adds it to **To** (one way only; removing it from the panel does not remove it from the message).
- **Send**: the message is posted to a test endpoint as JSON, and " Outlook extension PoC" is appended to the subject:
  ```json
  { "itemIds": ["itm-002", "itm-005"], "email": "<html…>", "recipients": ["to…", "cc…", "bcc…"] }
  ```
  - Outlook's own **Send** button is intercepted with the `OnMessageSend` event (Smart Alerts, `SendMode="SoftBlock"`). The handler never blocks: if anything fails, or after 20 s, the message is sent anyway.
  - The panel's **Send** button posts the JSON, updates the subject and sends the message with `item.sendAsync()`. The send handler then skips posting a second time.
  - Both only act on messages where the panel was opened. Other mail is sent untouched.

## Layout

| Path | What |
|---|---|
| `src/App.tsx` | The side panel UI. |
| `src/office.ts` | All Office.js calls (body, recipients, subject, session data), wrapped in promises. |
| `manifest.template.xml` | Add-in only (XML) manifest template (compose button, task pane, `OnMessageSend` launch event); `${...}` values come from `.env`. |
| `public/launchevent.js` | The `OnMessageSend` handler. Plain JS without imports, because classic Outlook on Windows runs it in a JavaScript-only runtime. |
| `public/commands.html` | Loads Office.js and the handler for Outlook on the web and new Outlook. |
| `server/` | The test endpoint: a dependency-free Node service (`GET items`, `GET emails`, `POST submit`, and the `drafts` endpoints used by the prototype), with its systemd unit and nginx snippet. See `server/README.md`. |
| `prototype/` | A clickable prototype: a web page that looks like Outlook on the web and runs the real panel inside it, without Outlook. See below. |

State between the panel and the send handler is kept in `item.sessionData` and, so that a reopened draft keeps it, in the item's custom properties.

## Clickable prototype

`prototype/` is a standalone page that imitates Outlook on the web (folder pane, message list, reading pane, compose, the add-in pane) and is filled with fictitious bank/trading sample data (lorem ipsum, `example.com` addresses). It exists to click through the idea without installing the add-in.

- The real panel (`src/App.tsx`, `src/office.ts`) runs unchanged. `prototype/fakeOffice.ts` stands in for Office.js and reads and writes the message held in `prototype/mailStore.ts` (in memory, reset on reload).
- **New mail** creates a draft on the test endpoint (`POST /drafts`) and stores its id in the message as the `x-idea-id` header. While the draft is edited, the selected items, recipients and subject are mirrored to the endpoint (`PUT /drafts/:id`).
- **Send** (Outlook's button or the panel's) does what `launchevent.js` does in Outlook: posts the JSON with the `ideaId`, appends the subject suffix, and the endpoint marks the draft as sent. The message moves to Sent Items and a copy is delivered to the Inbox so it can be opened as a recipient.
- Opening a received message that carries `x-idea-id` and pressing **Outlook Extension PoC** shows the read-mode pane (`prototype/ReadPanel.tsx`), which resolves the items from `GET /drafts/:id`. "View message headers" shows the header.

Run it locally:

```bash
cp .env.example .env         # ADDIN_BASE_URL=http://localhost:5173/outlook-addin, VITE_API_BASE=http://localhost:8799/api/outlook-poc
npm run prototype            # starts the endpoint on 8799 and Vite on 5173, opens the prototype in the browser
```

Or by hand: `ALLOWED_ORIGINS=http://localhost:5173 PORT=8799 node server/server.js`, then `npm run dev` and open http://localhost:5173/outlook-addin/prototype/index.html.

`npm run build` also emits the prototype at `dist/prototype/index.html`. `server/package.json` marks the server folder as CommonJS so `node server/server.js` runs inside this ESM package.

## Build and host

```bash
npm install
npm run build        # -> dist/
npm run validate     # Microsoft's manifest validator on dist/manifest.xml (needs Node 20+)
```

Host-specific values live in `.env` (git-ignored): copy `.env.example` and fill in the HTTPS folder `dist/` is served from, the test endpoint URL, your own add-in GUID and provider name. The build writes `dist/manifest.xml` from `manifest.template.xml` and puts the endpoint URL into the task pane and `dist/launchevent.js`. `dist/` must be served over HTTPS.

## Install for a test mailbox

Outlook on the web: **Apps → Get add-ins → My add-ins → Custom Addins → Add a custom add-in → Add from File…** and pick `dist/manifest.xml`. ("Add from URL" is disabled for personal Outlook.com accounts.) The add-in then shows up in every Outlook client signed in to that mailbox.

For an organization, an admin uploads the manifest in the Microsoft 365 admin center (**Settings → Integrated apps → Upload custom apps**).

## Tested

Outlook on the web with a personal Outlook.com mailbox:

- insert, replace and remove the items block;
- add addresses to To;
- post the JSON and send, both with Outlook's Send button and with the panel's Send button;
- the subject suffix.

Not yet tested: classic Outlook on Windows and new Outlook.
