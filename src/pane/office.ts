// Thin promise layer over Office.js for the compose item. The task pane only talks to Outlook through
// these functions; outside Outlook (no mailbox) `inOutlook()` is false and the UI runs read-only.

export const API: string = import.meta.env.VITE_API_BASE || "/api/outlook-poc";
export const SUBJECT_SUFFIX = " Outlook extension PoC";

export type Item = { id: string; title: string; html: string };
export type Contact = { name: string; email: string };
export type Payload = { itemIds: string[]; email: string; recipients: string[] };

const BLOCK_START = "poc-items-block";
const BLOCK_END = "poc-items-end";

function item(): Office.MessageCompose {
  return Office.context.mailbox.item as unknown as Office.MessageCompose;
}

function call<T>(fn: (cb: (r: Office.AsyncResult<T>) => void) => void): Promise<T> {
  return new Promise((resolve, reject) =>
    fn((r) => (r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(r.error))),
  );
}

export function inOutlook(): boolean {
  return typeof Office !== "undefined" && !!Office.context?.mailbox?.item;
}

export async function fetchJson<T>(path: string): Promise<T> {
  const res = await fetch(`${API}/${path}`, { cache: "no-store" });
  if (!res.ok) throw new Error(`${path}: HTTP ${res.status}`);
  return res.json();
}

// State shared with the add-in's OnMessageSend handler. sessionData lives for this compose
// session; custom properties are saved with the draft, so a reopened draft keeps its selection too.
export async function setSession(key: string, value: string): Promise<void> {
  await call<void>((cb) => item().sessionData.setAsync(key, value, cb)).catch(() => {});
  const props = await call<Office.CustomProperties>((cb) => item().loadCustomPropertiesAsync(cb));
  props.set(key, value);
  await call<void>((cb) => props.saveAsync(cb));
}

export async function getSaved(key: string): Promise<string> {
  const props = await call<Office.CustomProperties>((cb) => item().loadCustomPropertiesAsync(cb));
  return (props.get(key) as string) || "";
}

function escapeHtml(s: string): string {
  return s.replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[c]!);
}

function blockHtml(items: Item[]): string {
  if (items.length === 0) return "";
  const parts = items.map((i) => `<p><strong>${escapeHtml(i.title)}</strong></p>${i.html}`).join("");
  return (
    `<div id="${BLOCK_START}" style="border-left:3px solid #0f6cbd;padding:2px 10px;margin:0 0 12px">` +
    parts +
    `<p id="${BLOCK_END}" style="margin:0;font-size:1px;line-height:1px">&nbsp;</p></div>`
  );
}

// Outlook may rename ids (e.g. "x_poc-items-block"), so match on the id substring. The block runs from
// its opening div to the end marker paragraph and the div that closes it.
const BLOCK_RE = new RegExp(
  `<div[^>]*id="[^"]*${BLOCK_START}[^"]*"[^>]*>[\\s\\S]*?<p[^>]*id="[^"]*${BLOCK_END}[^"]*"[^>]*>[\\s\\S]*?</p>\\s*</div>`,
  "i",
);

/** Puts the selected items at the beginning of the body, replacing the block inserted earlier. */
export async function applyItems(items: Item[]): Promise<void> {
  const body = await call<string>((cb) => item().body.getAsync(Office.CoercionType.Html, cb));
  const block = blockHtml(items);
  if (BLOCK_RE.test(body)) {
    const html = body.replace(BLOCK_RE, block);
    await call<void>((cb) => item().body.setAsync(html, { coercionType: Office.CoercionType.Html }, cb));
  } else if (block) {
    await call<void>((cb) => item().body.prependAsync(block, { coercionType: Office.CoercionType.Html }, cb));
  }
  await setSession("pocItemIds", JSON.stringify(items.map((i) => i.id)));
}

export function addTo(c: Contact): Promise<void> {
  return call<void>((cb) => item().to.addAsync([{ displayName: c.name, emailAddress: c.email }], cb));
}

const emails = (list: Office.EmailAddressDetails[]) => list.map((r) => r.emailAddress).filter(Boolean);

export async function buildPayload(itemIds: string[]): Promise<Payload> {
  const it = item();
  const [email, to, cc, bcc] = await Promise.all([
    call<string>((cb) => it.body.getAsync(Office.CoercionType.Html, cb)),
    call<Office.EmailAddressDetails[]>((cb) => it.to.getAsync(cb)),
    call<Office.EmailAddressDetails[]>((cb) => it.cc.getAsync(cb)),
    call<Office.EmailAddressDetails[]>((cb) => it.bcc.getAsync(cb)),
  ]);
  return { itemIds, email, recipients: [...emails(to), ...emails(cc), ...emails(bcc)] };
}

export async function appendSubject(): Promise<void> {
  const subject = (await call<string>((cb) => item().subject.getAsync(cb))) || "";
  if (subject.endsWith(SUBJECT_SUFFIX)) return;
  await call<void>((cb) => item().subject.setAsync(subject + SUBJECT_SUFFIX, cb));
}

export async function submit(payload: Payload): Promise<{ id: string }> {
  const res = await fetch(`${API}/submit`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  const data = await res.json();
  if (!res.ok || !data.ok) throw new Error(data.error || `HTTP ${res.status}`);
  return data;
}

/** item.sendAsync exists only in newer requirement sets (Mailbox 1.15); returns false when missing. */
export async function trySend(): Promise<boolean> {
  const it = item() as unknown as { sendAsync?: (cb: (r: Office.AsyncResult<void>) => void) => void };
  if (typeof it.sendAsync !== "function") return false;
  await call<void>((cb) => it.sendAsync!(cb));
  return true;
}
