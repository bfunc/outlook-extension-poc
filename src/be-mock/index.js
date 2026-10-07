// Mock of the add-in test endpoint, running in the browser.
//
// Same routes and JSON shapes as the real service would have, state kept in memory for the life of the page.
// `installMockBackend(apiBase)` wraps `fetch`, so pane/office.ts and the shell call it exactly as
// they would call the real service. Sample data lives in samples.json.

import samples from "./samples.json";

// ---------------------------------------------------------------------------------------------
// Entities
// ---------------------------------------------------------------------------------------------

/**
 * @typedef {Object} KeyElements  What the pane edits (see pane/keyElementsModel.ts): one idea.
 * @property {"New" | "Pipeline" | "RevEnq"} status
 * @property {string} headline
 * @property {string} tradingArea            one of samples.json: tradingAreas
 * @property {Leg[]} legs
 *
 * @typedef {Object} Leg
 * @property {string} id
 * @property {string} instrument             one of samples.json: instruments
 * @property {boolean} emea                  mainly traded in EMEA
 * @property {boolean} factual               factual market comment, not a recommendation
 * @property {"Buy" | "Pay" | "Receive" | "Sell"} side
 * @property {string} price
 * @property {string} underlying
 * @property {string} timeHorizon
 *
 * @typedef {Object} Contact      An address the pane offers for To (samples.json: contacts).
 * @property {string} name
 * @property {string} email
 *
 * @typedef {Object} Draft        The backend record linked to a mail through its x-idea-id header.
 * @property {string} id                       uuid, the value of the x-idea-id header
 * @property {"draft" | "sent"} status
 * @property {KeyElements | null} keyElements  what the pane saved with the mail
 * @property {string[]} recipients             e-mail addresses (To, Cc, Bcc)
 * @property {string} subject
 * @property {string} createdAt                ISO date
 * @property {string} updatedAt                ISO date
 * @property {string} [sentAt]                 ISO date, once sent
 * @property {string} [submissionId]           the Submission created when the mail was sent
 *
 * @typedef {Object} Submission   What the add-in posts when a mail is sent (POST submit).
 * @property {string} id
 * @property {string} receivedAt               ISO date
 * @property {KeyElements | null} keyElements
 * @property {string} email                    the mail body as HTML
 * @property {string[]} recipients
 * @property {string} [ideaId]                 x-idea-id of the mail, when known
 */

/** @type {Contact[]} */
export const CONTACTS = samples.contacts;

export const MAX_DRAFTS = 500;

// ---------------------------------------------------------------------------------------------
// State
// ---------------------------------------------------------------------------------------------

export const db = {
  /** @type {Map<string, Draft>} */
  drafts: new Map(),
  /** @type {Submission[]} */
  submissions: [],
  /** @type {{ at: string, line: string }[]} */
  log: [],
};

const listeners = new Set();
let version = 0;

/** Lets a UI (the prototype's backend inspector) re-render when the state changes. */
export function subscribe(listener) {
  listeners.add(listener);
  return () => listeners.delete(listener);
}

export function getVersion() {
  return version;
}

function changed(line) {
  version += 1;
  db.log.push({ at: now(), line });
  if (db.log.length > 200) db.log.shift();
  listeners.forEach((l) => l());
}

function now() {
  return new Date().toISOString();
}

export function resetMockBackend() {
  db.drafts.clear();
  db.submissions.length = 0;
  db.log.length = 0;
  seedFromSamples();
  changed("reset");
}

/** Inbox mails in samples.json that carry ideaId + keyElements get a matching sent draft. */
function seedFromSamples() {
  for (const mail of samples.inbox) {
    if (!mail.ideaId || !mail.keyElements) continue;
    const [hours, minutes] = String(mail.time || "09:00").split(":").map(Number);
    const date = new Date();
    date.setDate(date.getDate() - mail.daysAgo);
    date.setHours(hours, minutes, 0, 0);
    const at = date.toISOString();
    db.drafts.set(mail.ideaId, {
      id: mail.ideaId,
      status: "sent",
      keyElements: withLegIds(mail.keyElements),
      recipients: [samples.me.email],
      subject: mail.subject,
      createdAt: at,
      updatedAt: at,
      sentAt: at,
    });
  }
}

function withLegIds(keyElements) {
  return { ...keyElements, legs: keyElements.legs.map((leg, i) => ({ id: leg.id || `leg-${i + 1}`, ...leg })) };
}

seedFromSamples();

// ---------------------------------------------------------------------------------------------
// Operations (one per route)
// ---------------------------------------------------------------------------------------------

/** @returns {Draft} */
export function createDraft() {
  const draft = {
    id: crypto.randomUUID(),
    status: "draft",
    keyElements: null,
    recipients: [],
    subject: "",
    createdAt: now(),
    updatedAt: now(),
  };
  db.drafts.set(draft.id, draft);
  while (db.drafts.size > MAX_DRAFTS) db.drafts.delete(db.drafts.keys().next().value);
  changed(`POST drafts -> ${short(draft.id)}`);
  return draft;
}

/** @returns {Draft | undefined} */
export function getDraft(id) {
  return db.drafts.get(id);
}

/** A draft as the API returns it. */
export function draftView(draft) {
  return { ...draft };
}

/** The list view, without recipient addresses. */
export function listDrafts() {
  return [...db.drafts.values()].map(({ recipients, ...rest }) => ({ ...rest, recipientCount: recipients.length }));
}

/**
 * Merges `patch` into a draft. Allowed keys: keyElements, recipients, subject, status ("sent" only).
 * @param {Draft} draft
 * @param {Partial<Draft>} patch
 * @returns {string | undefined} a validation error, or nothing on success
 */
export function updateDraft(draft, patch) {
  if (patch.keyElements !== undefined && !isKeyElements(patch.keyElements)) return "keyElements must be an object with legs[] or null";
  if (patch.recipients !== undefined && !isStringArray(patch.recipients)) return "recipients must be string[]";
  if (patch.subject !== undefined && typeof patch.subject !== "string") return "subject must be a string";
  if (patch.status !== undefined && patch.status !== "sent") return 'status can only be set to "sent"';

  if (patch.keyElements !== undefined) draft.keyElements = patch.keyElements;
  if (patch.recipients !== undefined) draft.recipients = patch.recipients;
  if (patch.subject !== undefined) draft.subject = patch.subject.slice(0, 500);
  if (patch.status === "sent") markSent(draft, now());
  draft.updatedAt = now();
  changed(`PUT drafts/${short(draft.id)} ${Object.keys(patch).join(",")}`);
  return undefined;
}

/**
 * Records what the add-in posts on send. With a known ideaId the linked draft becomes "sent".
 * @returns {{ error: string } | { submission: Submission }}
 */
export function submit(body) {
  const valid =
    body &&
    isKeyElements(body.keyElements) &&
    typeof body.email === "string" &&
    isStringArray(body.recipients) &&
    (body.ideaId === undefined || typeof body.ideaId === "string");
  if (!valid) return { error: "expected { keyElements: object | null, email: string, recipients: string[], ideaId?: string }" };

  const submission = {
    id: crypto.randomUUID(),
    receivedAt: now(),
    keyElements: body.keyElements,
    email: body.email,
    recipients: body.recipients,
    ideaId: body.ideaId,
  };
  db.submissions.push(submission);

  const draft = body.ideaId ? db.drafts.get(body.ideaId) : undefined;
  if (draft) {
    markSent(draft, submission.receivedAt);
    draft.submissionId = submission.id;
    draft.updatedAt = submission.receivedAt;
  }
  changed(`POST submit -> ${short(submission.id)}${draft ? ` (draft ${short(draft.id)} sent)` : ""}`);
  return { submission };
}

function markSent(draft, at) {
  if (draft.status === "sent") return;
  draft.status = "sent";
  draft.sentAt = at;
}

function isKeyElements(value) {
  return value === null || (typeof value === "object" && Array.isArray(value.legs));
}

function isStringArray(value) {
  return Array.isArray(value) && value.every((x) => typeof x === "string");
}

function short(id) {
  return id.slice(0, 8);
}

// ---------------------------------------------------------------------------------------------
// Routes: method + path (relative to the API base) -> { status, body }
// ---------------------------------------------------------------------------------------------

export function route(method, path, body) {
  if (method === "GET" && path === "health") return ok(200, { ok: true });
  if (method === "GET" && path === "emails") return ok(200, CONTACTS);

  if (method === "POST" && path === "drafts") return ok(201, { ok: true, draft: draftView(createDraft()) });
  if (method === "GET" && path === "drafts") return ok(200, listDrafts());

  const one = /^drafts\/([^/]+)$/.exec(path);
  if (one) {
    const draft = getDraft(one[1]);
    if (!draft) return fail(404, "draft not found");
    if (method === "GET") return ok(200, { ok: true, draft: draftView(draft) });
    if (method === "PUT") {
      const error = updateDraft(draft, body || {});
      return error ? fail(400, error) : ok(200, { ok: true, draft: draftView(draft) });
    }
  }

  if (method === "POST" && path === "submit") {
    const result = submit(body);
    if ("error" in result) return fail(400, result.error);
    return ok(200, { ok: true, id: result.submission.id, receivedAt: result.submission.receivedAt });
  }

  return fail(404, "not found");
}

function ok(status, body) {
  return { status, body };
}

function fail(status, error) {
  return { status, body: { ok: false, error } };
}

// ---------------------------------------------------------------------------------------------
// fetch interceptor
// ---------------------------------------------------------------------------------------------

export const MOCK_LATENCY_MS = 60;

/** Routes every fetch to `${apiBase}/…` into the mock; everything else goes to the network. */
export function installMockBackend(apiBase) {
  const base = apiBase.replace(/\/$/, "") + "/";
  const realFetch = globalThis.fetch.bind(globalThis);

  globalThis.fetch = async (input, init) => {
    const url = typeof input === "string" ? input : input instanceof URL ? input.href : input.url;
    if (!url.startsWith(base)) return realFetch(input, init);

    const method = (init?.method || (input instanceof Request ? input.method : "GET")).toUpperCase();
    const path = url.slice(base.length).replace(/\?.*$/, "");
    let body;
    if (init?.body) {
      try {
        body = JSON.parse(String(init.body));
      } catch {
        return jsonResponse(400, { ok: false, error: "invalid JSON" });
      }
    }
    await new Promise((resolve) => setTimeout(resolve, MOCK_LATENCY_MS));
    const result = route(method, path, body);
    return jsonResponse(result.status, result.body);
  };
}

function jsonResponse(status, body) {
  return new Response(JSON.stringify(body), { status, headers: { "Content-Type": "application/json" } });
}
