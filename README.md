# Outlook add-in: clickable prototype

A web page that looks like Outlook on the web and lets you click through an add-in idea without Outlook, an installed add-in or a server. Everything runs in the browser: the task pane runs on a stand-in for Office.js, and the test endpoint is mocked in memory. All data is fictitious (lorem ipsum text, `example.com` addresses).

```bash
npm install
npm run dev        # opens the prototype in the browser
npm run build      # static site in dist/, to be served under https://<host>/static/outlook/
```

## The idea being prototyped

An add-in with a side panel in the compose window, shown as a 2-step wizard (Key elements, Select recipients). Only step 1 is built so far.

- **Key Elements** (step 1). A card for one idea: status (New / Pipeline / RevEnq), headline, trading area, and instrument legs added from a list. Each leg has an EMEA checkbox; when checked, a "factual market comment" checkbox appears; a leg that is EMEA and not factual is an investment recommendation and gets Buy / Pay / Receive / Sell, Price, Underlying(s) (called Instrument ID for Cash Equities and Cash Bonds) and Time Horizon. The card is titled IDEA, or INVESTMENT RECOMMENDATION when any leg is one. **Save** switches to a read-only view card (`NEW INVESTMENT RECOMMENDATION`, trading area, legs with side chips); **Edit** goes back. No validation. The mail body is not touched.
- **Draft linked to a backend record.** New mail creates a draft on the endpoint (`POST /drafts`) and stores its id in the message as the `x-idea-id` header. The key elements, recipients and subject are mirrored to it while editing (`PUT /drafts/:id`).
- **Send** (Outlook's button) posts the key elements, the body and the recipients as JSON with the `ideaId`, appends " Outlook extension PoC" to the subject, and the endpoint marks the draft as sent. The message moves to Sent Items and a copy is delivered to the Inbox so it can be opened as a recipient.
- **Read mode.** Opening a received message that carries `x-idea-id` and pressing **Outlook Extension PoC** shows the same view card, resolved from `GET /drafts/:id`. "View message headers" shows the header.
- The **cog** at the bottom of the left rail opens the mock backend console: drafts, submissions, request log, Reset.

## Layout

| Path | What |
|---|---|
| `src/pane/App.tsx` | The side panel: stepper, Key Elements edit/view, Save/Edit. Written against Office.js as the real add-in would be; state is saved with the draft in custom properties. |
| `src/pane/KeyElements.tsx`, `keyElementsModel.ts` | The Key Elements form and view card, and the record type, defaults and rules (recommendation, card title, underlying label). |
| `src/pane/office.ts` | Office.js calls (recipients, subject, session data, custom properties) wrapped in promises, and the endpoint calls. |
| `src/shell/Shell.tsx`, `shell.css` | The Outlook look: folder pane, message list, reading pane, compose, the add-in pane host, the send flow. |
| `src/shell/mailStore.ts` | The in-memory mailbox (folders, messages, headers, custom properties), seeded from `samples.json`. |
| `src/shell/fakeOffice.ts` | Office.js stand-in: `item.body`, `subject`, `to/cc/bcc`, `sessionData`, custom properties, `sendAsync`. |
| `src/shell/ReadPanel.tsx` | The panel in read mode. |
| `src/shell/BackendInspector.tsx` | The mock backend console. |
| `src/be-mock/index.js` | The endpoint mock, plain JS: entities (`Draft`, `Submission`), the in-memory `db`, one function per operation, a `route()` table, and `installMockBackend()` which answers `fetch` calls to the API base. |
| `src/be-mock/samples.json` | All sample data, edit by hand: `me`, `contacts` (inbox senders), `tradingAreas`, `instruments` (name + which label the underlying field gets), `inbox` (initial messages; `daysAgo` + `time` keep them recent, `body` is a list of HTML paragraphs). |

## Endpoint the mock implements

| Route | |
|---|---|
| `GET emails` | `{ name, email }[]` |
| `POST drafts` | 201 `{ ok, draft }`, a new `{ id, status: "draft", keyElements: null, recipients, subject, createdAt, updatedAt }` |
| `GET drafts` | list without recipient addresses |
| `GET drafts/:id` | one draft, 404 if unknown |
| `PUT drafts/:id` | merges `keyElements`, `recipients`, `subject`, `status: "sent"` |
| `POST submit` | `{ keyElements, email, recipients, ideaId? }`, 200 `{ ok, id, receivedAt }`; a known `ideaId` marks the draft sent |

The API base is `VITE_API_BASE` if set, else `/api/outlook-poc`. To point the prototype at a real service with these routes, drop the `installMockBackend` call in `src/main.tsx`.

## Hosting

The build assumes the path `/static/outlook/` (`base` in `vite.config.ts`): copy `dist/` to the folder served at `https://<host>/static/outlook/`. The dev server uses the same path, http://localhost:5173/static/outlook/.

## Stack

Vite 8, React 18, TypeScript. `@types/office-js` gives the pane the Office.js types; no Office.js script is loaded.
