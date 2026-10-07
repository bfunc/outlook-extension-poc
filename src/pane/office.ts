// Thin promise layer over Office.js for the compose item. The task pane only talks to Outlook through
// these functions; outside Outlook (no mailbox) `inOutlook()` is false and the UI runs read-only.

import { parseKeyElements, type KeyElements } from "./keyElementsModel";

export const API: string = import.meta.env.VITE_API_BASE || "/api/outlook-poc";
export const SUBJECT_SUFFIX = " Outlook extension PoC";

export type Contact = { name: string; email: string };
export type Payload = { keyElements: KeyElements | null; email: string; recipients: string[] };

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
// session; custom properties are saved with the draft, so a reopened draft keeps its state too.
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

const emails = (list: Office.EmailAddressDetails[]) => list.map((r) => r.emailAddress).filter(Boolean);

export function addTo(c: Contact): Promise<void> {
  return call<void>((cb) => item().to.addAsync([{ displayName: c.name, emailAddress: c.email }], cb));
}

export async function currentTo(): Promise<string[]> {
  return emails(await call<Office.EmailAddressDetails[]>((cb) => item().to.getAsync(cb)));
}

/** What is posted on send: the key elements saved with the draft, the body and the recipients. */
export async function buildPayload(): Promise<Payload> {
  const it = item();
  const [saved, email, to, cc, bcc] = await Promise.all([
    getSaved("pocKeyElements"),
    call<string>((cb) => it.body.getAsync(Office.CoercionType.Html, cb)),
    call<Office.EmailAddressDetails[]>((cb) => it.to.getAsync(cb)),
    call<Office.EmailAddressDetails[]>((cb) => it.cc.getAsync(cb)),
    call<Office.EmailAddressDetails[]>((cb) => it.bcc.getAsync(cb)),
  ]);
  return { keyElements: parseKeyElements(saved), email, recipients: [...emails(to), ...emails(cc), ...emails(bcc)] };
}

export async function appendSubject(): Promise<void> {
  const subject = (await call<string>((cb) => item().subject.getAsync(cb))) || "";
  if (subject.endsWith(SUBJECT_SUFFIX)) return;
  await call<void>((cb) => item().subject.setAsync(subject + SUBJECT_SUFFIX, cb));
}

export async function submit(payload: Payload & { ideaId?: string }): Promise<{ id: string }> {
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
