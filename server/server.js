// Tiny test HTTP service for the Outlook add-in PoC.
// Zero npm dependencies: only node:http, node:fs, node:path, node:crypto.
// Public paths are served behind nginx at /api/outlook-poc/ (prefix kept).

const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');

const PORT = Number(process.env.PORT) || 8790;
const DATA_FILE = path.resolve(process.env.DATA_FILE || './submissions.jsonl');
const MAX_BODY_BYTES = 1 * 1024 * 1024; // 1 MB
const MAX_SUBMISSIONS = 200;

// The add-in's own host(s) come from the environment (comma-separated), e.g.
// ALLOWED_ORIGINS=https://addin.example.com. The Outlook origins below are always allowed.
const ALLOWED_ORIGINS = new Set([
  ...(process.env.ALLOWED_ORIGINS || '').split(',').map((s) => s.trim()).filter(Boolean),
  'https://outlook.live.com',
  'https://outlook.office.com',
  'https://outlook.office365.com',
]);
const ALLOWED_ORIGIN_SUFFIXES = ['.officeapps.live.com', '.outlook.com'];

// Sample data. Everything here is fictitious: invented desks, invented people, lorem ipsum text,
// example.com addresses. The add-in inserts "items" into a mail and offers "emails" for the To line.
const SAMPLE_ITEMS = [
  {
    id: 'itm-001',
    title: 'FX price indication EUR/USD (fictitious)',
    html: '<p>Lorem ipsum indicative spot <strong>1.0842 / 1.0845</strong>, valid for lorem minutes. Dolor sit amet, consectetur adipiscing elit.</p>',
  },
  {
    id: 'itm-002',
    title: 'Term sheet summary: 3Y EUR note (fictitious)',
    html: '<p>Notional lorem 10m, coupon <strong>ipsum 3.10%</strong> p.a., maturity dolor 2029. Sed do eiusmod tempor incididunt. Full term sheet: <a href="https://example.com/term-sheet">example.com/term-sheet</a>.</p>',
  },
  {
    id: 'itm-003',
    title: 'Research snippet: rates outlook (fictitious)',
    html: '<p>Lorem ipsum: we expect the curve to <strong>consectetur</strong> in Q3, driven by adipiscing elit. Ut enim ad minim veniam, quis nostrud exercitation.</p>',
  },
  {
    id: 'itm-004',
    title: 'Trade confirmation summary #LRM-4821 (fictitious)',
    html: '<p>Buy lorem 5,000 ipsum shares at dolor 42.10, settlement T+2. <strong>Status: ipsum confirmed</strong>. Duis aute irure dolor in reprehenderit.</p>',
  },
  {
    id: 'itm-005',
    title: 'Settlement reminder (fictitious)',
    html: '<p>Lorem ipsum settlement for trade #LRM-4790 is due in <strong>2 business days</strong>. Please dolor sit amet the instructions at <a href="https://example.com/settlement">example.com/settlement</a>.</p>',
  },
  {
    id: 'itm-006',
    title: 'KYC refresh notice (fictitious)',
    html: '<p>Lorem ipsum KYC documents for Lorem Holdings expire on dolor 30. <strong>Action: sit amet</strong> before the review date. Excepteur sint occaecat cupidatat non proident.</p>',
  },
];

const SAMPLE_EMAILS = [
  { name: 'Lorem Ipsum (FX desk, fictitious)', email: 'fx.desk@example.com' },
  { name: 'Dolor Sit (Settlements, fictitious)', email: 'settlements@example.com' },
  { name: 'Amet Consectetur (Compliance, fictitious)', email: 'compliance@example.com' },
  { name: 'Adipiscing Elit (Client Onboarding, fictitious)', email: 'onboarding@example.com' },
  { name: 'Tempor Incididunt (Research, fictitious)', email: 'research@example.com' },
  { name: 'Magna Aliqua (Lorem Holdings, fictitious)', email: 'treasury@lorem-holdings.example.com' },
];

// Drafts live in memory only: a draft is created when the add-in opens in a new mail, updated as items
// and recipients change, and marked sent when the mail is submitted. The id travels in the mail as the
// x-idea-id header, so a recipient's panel can look the items up. Restarting the service forgets them.
const MAX_DRAFTS = 500;
const drafts = new Map();

function createDraft() {
  const now = new Date().toISOString();
  const draft = { id: crypto.randomUUID(), status: 'draft', itemIds: [], recipients: [], subject: '', createdAt: now, updatedAt: now };
  drafts.set(draft.id, draft);
  if (drafts.size > MAX_DRAFTS) drafts.delete(drafts.keys().next().value);
  return draft;
}

function draftView(draft) {
  return { ...draft, items: SAMPLE_ITEMS.filter((i) => draft.itemIds.includes(i.id)) };
}

function readJsonBody(req, res, onBody) {
  let totalBytes = 0;
  const chunks = [];
  let rejected = false;
  req.on('data', (chunk) => {
    if (rejected) return;
    totalBytes += chunk.length;
    if (totalBytes > MAX_BODY_BYTES) {
      rejected = true;
      sendJson(res, 413, { ok: false, error: 'payload too large' });
      req.destroy();
      return;
    }
    chunks.push(chunk);
  });
  req.on('end', () => {
    if (rejected) return;
    let body;
    try {
      body = JSON.parse(Buffer.concat(chunks).toString('utf8') || '{}');
    } catch {
      sendJson(res, 400, { ok: false, error: 'invalid JSON body' });
      return;
    }
    onBody(body);
  });
  req.on('error', () => {});
}

function handleDraftUpdate(req, res, draft) {
  readJsonBody(req, res, (body) => {
    if (body.itemIds !== undefined && !isStringArray(body.itemIds)) return sendJson(res, 400, { ok: false, error: 'itemIds must be string[]' });
    if (body.recipients !== undefined && !isStringArray(body.recipients)) return sendJson(res, 400, { ok: false, error: 'recipients must be string[]' });
    if (body.subject !== undefined && typeof body.subject !== 'string') return sendJson(res, 400, { ok: false, error: 'subject must be a string' });
    if (body.status !== undefined && body.status !== 'sent') return sendJson(res, 400, { ok: false, error: 'status can only be set to "sent"' });
    if (body.itemIds !== undefined) draft.itemIds = body.itemIds;
    if (body.recipients !== undefined) draft.recipients = body.recipients;
    if (body.subject !== undefined) draft.subject = body.subject.slice(0, 500);
    if (body.status === 'sent' && draft.status !== 'sent') { draft.status = 'sent'; draft.sentAt = new Date().toISOString(); }
    draft.updatedAt = new Date().toISOString();
    sendJson(res, 200, { ok: true, draft: draftView(draft) });
  });
}

function isAllowedOrigin(origin) {
  if (!origin) return false;
  if (ALLOWED_ORIGINS.has(origin)) return true;
  return ALLOWED_ORIGIN_SUFFIXES.some((suffix) => origin.endsWith(suffix));
}

function applyCors(req, res) {
  const origin = req.headers.origin;
  res.setHeader('Vary', 'Origin');
  if (origin === 'null') {
    res.setHeader('Access-Control-Allow-Origin', 'null');
  } else if (isAllowedOrigin(origin)) {
    res.setHeader('Access-Control-Allow-Origin', origin);
  }
  // else: no ACAO header at all.
}

function sendJson(res, status, obj) {
  res.writeHead(status, {
    'Content-Type': 'application/json; charset=utf-8',
    'Cache-Control': 'no-store',
  });
  res.end(JSON.stringify(obj));
}

function readSubmissions() {
  try {
    const text = fs.readFileSync(DATA_FILE, 'utf8');
    return text.split('\n').filter((line) => line.trim().length > 0);
  } catch {
    return [];
  }
}

function appendSubmission(entry) {
  fs.mkdirSync(path.dirname(DATA_FILE), { recursive: true });
  const lines = readSubmissions();
  lines.push(JSON.stringify(entry));
  const trimmed = lines.slice(-MAX_SUBMISSIONS);
  fs.writeFileSync(DATA_FILE, trimmed.join('\n') + '\n', 'utf8');
}

function isStringArray(value) {
  return Array.isArray(value) && value.every((v) => typeof v === 'string');
}

function handleSubmit(req, res) {
  let totalBytes = 0;
  const chunks = [];
  let rejected = false;

  req.on('data', (chunk) => {
    if (rejected) return;
    totalBytes += chunk.length;
    if (totalBytes > MAX_BODY_BYTES) {
      rejected = true;
      sendJson(res, 413, { ok: false, error: 'payload too large' });
      req.destroy();
      return;
    }
    chunks.push(chunk);
  });

  req.on('end', () => {
    if (rejected) return;

    let body;
    try {
      body = JSON.parse(Buffer.concat(chunks).toString('utf8'));
    } catch {
      sendJson(res, 400, { ok: false, error: 'invalid JSON body' });
      return;
    }

    if (
      !body ||
      typeof body !== 'object' ||
      !isStringArray(body.itemIds) ||
      typeof body.email !== 'string' ||
      !isStringArray(body.recipients) ||
      (body.ideaId !== undefined && typeof body.ideaId !== 'string')
    ) {
      sendJson(res, 400, {
        ok: false,
        error: 'expected { itemIds: string[], email: string, recipients: string[], ideaId?: string }',
      });
      return;
    }

    const id = crypto.randomUUID();
    const receivedAt = new Date().toISOString();
    const draft = body.ideaId ? drafts.get(body.ideaId) : undefined;
    if (draft) {
      Object.assign(draft, { status: 'sent', itemIds: body.itemIds, recipients: body.recipients, sentAt: receivedAt, submissionId: id, updatedAt: receivedAt });
    }
    appendSubmission({
      id,
      receivedAt,
      ideaId: body.ideaId,
      itemIds: body.itemIds,
      recipients: body.recipients,
      emailLength: body.email.length,
      email: body.email,
    });

    sendJson(res, 200, { ok: true, id, receivedAt });
  });

  req.on('error', () => {
    // client aborted mid-upload; nothing to respond to.
  });
}

function logRequest(method, pathname, status) {
  // Never log email body content.
  console.log(`${method} ${pathname} ${status}`);
}

const server = http.createServer((req, res) => {
  const { pathname } = new URL(req.url, 'http://localhost');
  applyCors(req, res);

  if (req.method === 'OPTIONS') {
    res.writeHead(204, {
      'Access-Control-Allow-Methods': 'GET, POST, PUT, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type',
      'Access-Control-Max-Age': '600',
    });
    res.end();
    logRequest(req.method, pathname, 204);
    return;
  }

  if (req.method === 'GET' && pathname === '/api/outlook-poc/health') {
    sendJson(res, 200, { ok: true });
    logRequest(req.method, pathname, 200);
    return;
  }

  if (req.method === 'GET' && pathname === '/api/outlook-poc/items') {
    sendJson(res, 200, SAMPLE_ITEMS);
    logRequest(req.method, pathname, 200);
    return;
  }

  if (req.method === 'GET' && pathname === '/api/outlook-poc/emails') {
    sendJson(res, 200, SAMPLE_EMAILS);
    logRequest(req.method, pathname, 200);
    return;
  }

  if (pathname === '/api/outlook-poc/drafts') {
    if (req.method === 'POST') {
      const draft = createDraft();
      sendJson(res, 201, { ok: true, draft: draftView(draft) });
      logRequest(req.method, pathname, 201);
      return;
    }
    if (req.method === 'GET') {
      const list = [...drafts.values()].map(({ id, status, itemIds, recipients, subject, createdAt, updatedAt, sentAt }) =>
        ({ id, status, itemIds, recipientCount: recipients.length, subject, createdAt, updatedAt, sentAt }));
      sendJson(res, 200, list);
      logRequest(req.method, pathname, 200);
      return;
    }
  }

  const draftMatch = /^\/api\/outlook-poc\/drafts\/([0-9a-f-]{36})$/.exec(pathname);
  if (draftMatch) {
    const draft = drafts.get(draftMatch[1]);
    if (!draft) {
      sendJson(res, 404, { ok: false, error: 'draft not found' });
      logRequest(req.method, pathname, 404);
      return;
    }
    if (req.method === 'GET') {
      sendJson(res, 200, { ok: true, draft: draftView(draft) });
      logRequest(req.method, pathname, 200);
      return;
    }
    if (req.method === 'PUT') {
      const origEnd = res.end.bind(res);
      res.end = (...args) => {
        logRequest(req.method, pathname, res.statusCode);
        return origEnd(...args);
      };
      handleDraftUpdate(req, res, draft);
      return;
    }
  }

  if (req.method === 'POST' && pathname === '/api/outlook-poc/submit') {
    // Status is logged inside handleSubmit's response paths.
    const origEnd = res.end.bind(res);
    res.end = (...args) => {
      logRequest(req.method, pathname, res.statusCode);
      return origEnd(...args);
    };
    handleSubmit(req, res);
    return;
  }

  sendJson(res, 404, { ok: false, error: 'not found' });
  logRequest(req.method, pathname, 404);
});

server.listen(PORT, '127.0.0.1', () => {
  console.log(`outlook-poc server listening on 127.0.0.1:${PORT}`);
});
