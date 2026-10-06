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

const SAMPLE_ITEMS = [
  {
    id: 'itm-001',
    title: 'Quarterly report summary',
    html: '<p>Revenue grew <strong>8%</strong> quarter over quarter, with the biggest gains in the EU region. Full breakdown attached.</p>',
  },
  {
    id: 'itm-002',
    title: 'Meeting agenda',
    html: '<p>Agenda for Thursday: roadmap review, budget check-in, and open Q&amp;A. See <a href="https://example.com/agenda">agenda doc</a>.</p>',
  },
  {
    id: 'itm-003',
    title: 'Product update',
    html: '<p>Version 2.3 ships next week with <strong>faster search</strong> and a redesigned settings panel.</p>',
  },
  {
    id: 'itm-004',
    title: 'Support ticket #4821',
    html: '<p>Customer reports intermittent sync failures on mobile. Logs attached, <strong>priority: high</strong>.</p>',
  },
  {
    id: 'itm-005',
    title: 'Invoice reminder',
    html: '<p>Invoice #10432 is due in 5 days. Pay online at <a href="https://example.com/billing">example.com/billing</a>.</p>',
  },
  {
    id: 'itm-006',
    title: 'Release notes 2.3',
    html: '<p>Highlights: bug fixes, <strong>performance improvements</strong>, and a new dark mode theme.</p>',
  },
];

const SAMPLE_EMAILS = [
  { name: 'Alice Martin', email: 'alice.martin@example.com' },
  { name: 'Bob Chen', email: 'bob.chen@example.com' },
  { name: 'Carla Diaz', email: 'carla.diaz@example.com' },
  { name: 'David Kim', email: 'david.kim@example.com' },
  { name: 'Elena Rossi', email: 'elena.rossi@example.com' },
  { name: 'Farid Hussain', email: 'farid.hussain@example.com' },
];

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
      !isStringArray(body.recipients)
    ) {
      sendJson(res, 400, {
        ok: false,
        error: 'expected { itemIds: string[], email: string, recipients: string[] }',
      });
      return;
    }

    const id = crypto.randomUUID();
    const receivedAt = new Date().toISOString();
    appendSubmission({
      id,
      receivedAt,
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
      'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
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
