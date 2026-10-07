# outlook-poc server

Tiny dependency-free Node HTTP service backing an Outlook add-in PoC. Serves
sample "items" and "email addresses", and accepts a submit POST that logs to
a local JSONL file (capped at the last 200 entries). No endpoint reads
submissions back publicly, since they contain email content.

## Endpoints (public prefix `/api/outlook-poc/`)

- `GET /api/outlook-poc/health` — `{ ok: true }`
- `GET /api/outlook-poc/items` — 6 sample items `{ id, title, html }`
- `GET /api/outlook-poc/emails` — 6 sample addresses `{ name, email }` (all `@example.com`)
- `POST /api/outlook-poc/submit` — body `{ itemIds: string[], email: string, recipients: string[], ideaId?: string }`
  - 400 if shapes don't match, 413 if body > 1 MB
  - 200 `{ ok: true, id, receivedAt }` on success; with a known `ideaId` the draft is marked `sent`

Drafts (in memory, at most 500, oldest dropped; used by the clickable prototype to link a mail to a backend record through its `x-idea-id` header):

- `POST /api/outlook-poc/drafts` — 201 `{ ok: true, draft }`, a new draft `{ id, status: "draft", itemIds: [], recipients: [], subject: "", items: [], createdAt, updatedAt }`
- `GET /api/outlook-poc/drafts` — list `{ id, status, itemIds, recipientCount, subject, createdAt, updatedAt }`
- `GET /api/outlook-poc/drafts/:id` — one draft with its `items` resolved; 404 if unknown
- `PUT /api/outlook-poc/drafts/:id` — body with any of `itemIds: string[]`, `recipients: string[]`, `subject: string`, `status: "sent"`; returns the draft

The draft endpoints return recipients and subjects, so they are meant for the prototype and test data only.

CORS allows the origins in `ALLOWED_ORIGINS` (comma-separated env variable: the add-in's own host),
`https://outlook.live.com`, `https://outlook.office.com`,
`https://outlook.office365.com`, and anything ending in
`.officeapps.live.com` / `.outlook.com`, plus `Origin: null`.

## Local run

```
PORT=8799 DATA_FILE=./tmp/submissions.jsonl node server.js
```

## Deploy (manual, on the server)

1. Copy `server.js` to `/opt/outlook-poc/server.js`.
2. Create the system user/group and state dir:
   ```
   sudo useradd --system --no-create-home --shell /usr/sbin/nologin outlook-poc
   sudo mkdir -p /opt/outlook-poc
   sudo chown outlook-poc:outlook-poc /opt/outlook-poc
   ```
3. Put host settings in `/etc/outlook-poc/env` (format: `server/.env.example`), e.g.
   `ALLOWED_ORIGINS=https://addin.example.com`.
4. Install the unit: copy `outlook-poc.service` to `/etc/systemd/system/outlook-poc.service`, then:
   ```
   sudo systemctl daemon-reload
   sudo systemctl enable --now outlook-poc
   ```
5. Add the contents of `nginx-location.conf` inside the site's existing
   `server { ... listen 443 ... }` block, then:
   ```
   sudo nginx -t
   sudo systemctl reload nginx
   ```

## Reading submissions

```
sudo tail -n 5 /var/lib/outlook-poc/submissions.jsonl | jq
```
