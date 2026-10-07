# Outlook add-in: clickable prototype

A web page that looks like Outlook on the web and lets you click through an add-in idea without Outlook, an installed add-in or a server. Everything runs in the browser: the task pane runs on a stand-in for Office.js, and the test endpoint is mocked in memory. All data is fictitious (lorem ipsum text, `example.com` addresses).

```bash
npm install
npm run dev        # opens the prototype in the browser
npm run build      # static site in dist/, to be served under https://<host>/static/outlook/
```

## The idea being prototyped

An add-in with a side panel in the compose window:

- **Items** from an endpoint: checked items are inserted as one block at the top of the message body. Changing the selection replaces the block.
- **List of emails** from the endpoint: clicking an address adds it to **To**.
- **Draft linked to a backend record.** New mail creates a draft on the endpoint (`POST /drafts`) and stores its id in the message as the `x-idea-id` header. While the draft is edited, the selected items, recipients and subject are mirrored to it (`PUT /drafts/:id`).
- **Send** (Outlook's button or the panel's) posts the message as JSON with the `ideaId`, appends " Outlook extension PoC" to the subject, and the endpoint marks the draft as sent. The message moves to Sent Items and a copy is delivered to the Inbox so it can be opened as a recipient.
- **Read mode.** Opening a received message that carries `x-idea-id` and pressing **Outlook Extension PoC** shows the panel in read mode, which resolves the items from `GET /drafts/:id`. "View message headers" shows the header.
- The **cog** at the bottom of the left rail opens the mock backend console: drafts, submissions, request log, Reset.

## Layout

| Path | What |
|---|---|
| `src/pane/App.tsx` | The side panel UI, written against Office.js as the real add-in would be. |
| `src/pane/office.ts` | Office.js calls (body, recipients, subject, session data) wrapped in promises, and the endpoint calls. |
| `src/shell/Shell.tsx`, `shell.css` | The Outlook look: folder pane, message list, reading pane, compose, the add-in pane host, the send flow. |
| `src/shell/mailStore.ts` | The in-memory mailbox (folders, messages, headers, custom properties), seeded from `samples.json`. |
| `src/shell/fakeOffice.ts` | Office.js stand-in: `item.body`, `subject`, `to/cc/bcc`, `sessionData`, custom properties, `sendAsync`. |
| `src/shell/ReadPanel.tsx` | The panel in read mode. |
| `src/shell/BackendInspector.tsx` | The mock backend console. |
| `src/be-mock/index.js` | The endpoint mock, plain JS: entities (`Draft`, `Submission`), the in-memory `db`, one function per operation, a `route()` table, and `installMockBackend()` which answers `fetch` calls to the API base. |
| `src/be-mock/samples.json` | All sample data, edit by hand: `me`, `contacts` (the panel's list and the inbox senders), `items`, `inbox` (initial messages; `daysAgo` + `time` keep them recent, `body` is a list of HTML paragraphs). |

## Endpoint the mock implements

| Route | |
|---|---|
| `GET items` | `{ id, title, html }[]` |
| `GET emails` | `{ name, email }[]` |
| `POST drafts` | 201 `{ ok, draft }`, a new `{ id, status: "draft", itemIds, recipients, subject, items, createdAt, updatedAt }` |
| `GET drafts` | list without recipient addresses |
| `GET drafts/:id` | one draft with its `items` resolved, 404 if unknown |
| `PUT drafts/:id` | merges `itemIds`, `recipients`, `subject`, `status: "sent"` |
| `POST submit` | `{ itemIds, email, recipients, ideaId? }`, 200 `{ ok, id, receivedAt }`; a known `ideaId` marks the draft sent |

The API base is `VITE_API_BASE` if set, else `/api/outlook-poc`. To point the prototype at a real service with these routes, drop the `installMockBackend` call in `src/main.tsx`.

## Hosting

The build assumes the path `/static/outlook/` (`base` in `vite.config.ts`): copy `dist/` to the folder served at `https://<host>/static/outlook/`. The dev server uses the same path, http://localhost:5173/static/outlook/.

## Stack

Vite 8, React 18, TypeScript. `@types/office-js` gives the pane the Office.js types; no Office.js script is loaded.
