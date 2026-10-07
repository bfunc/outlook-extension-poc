// A stand-in for Office.js, covering only what src/office.ts calls, so the real task pane runs
// inside the prototype unchanged. Every call reads and writes the draft that is open in the
// prototype's compose window.

import { type Address, mailStore } from "./mailStore";

type Callback<T> = (result: { status: string; value: T; error?: { message: string } }) => void;

const SUCCEEDED = "succeeded";
const FAILED = "failed";

let composingId: string | undefined;

/** Tells the fake Office which draft the pane is working on. */
export function setComposingMessage(id: string | undefined): void {
  composingId = id;
}

export function composingMessageId(): string | undefined {
  return composingId;
}

function ok<T>(cb: Callback<T>, value: T): void {
  setTimeout(() => cb({ status: SUCCEEDED, value }), 0);
}

function fail<T>(cb: Callback<T>, message: string): void {
  setTimeout(() => cb({ status: FAILED, value: undefined as unknown as T, error: { message } }), 0);
}

function draft() {
  const m = composingId ? mailStore.get(composingId) : undefined;
  if (!m) throw new Error("No message is being composed");
  return m;
}

function recipients(field: "to" | "cc" | "bcc") {
  return {
    getAsync(cb: Callback<{ displayName: string; emailAddress: string }[]>) {
      try {
        ok(cb, draft()[field].map((a) => ({ displayName: a.name, emailAddress: a.email })));
      } catch (e) {
        fail(cb, (e as Error).message);
      }
    },
    addAsync(list: { displayName: string; emailAddress: string }[], cb: Callback<void>) {
      try {
        const id = draft().id;
        const added: Address[] = list.map((r) => ({ name: r.displayName, email: r.emailAddress }));
        mailStore.update(id, (m) => ({ [field]: [...m[field], ...added.filter((a) => !m[field].some((x) => x.email === a.email))] }));
        ok(cb, undefined);
      } catch (e) {
        fail(cb, (e as Error).message);
      }
    },
  };
}

function withDraft<T>(cb: Callback<T>, fn: (m: ReturnType<typeof draft>) => T): void {
  try {
    ok(cb, fn(draft()));
  } catch (e) {
    fail(cb, (e as Error).message);
  }
}

const item = {
  body: {
    getAsync(_coercion: string, cb: Callback<string>) {
      withDraft(cb, (m) => m.bodyHtml);
    },
    setAsync(html: string, _options: unknown, cb: Callback<void>) {
      withDraft(cb, (m) => {
        mailStore.update(m.id, { bodyHtml: html });
      });
    },
    prependAsync(html: string, _options: unknown, cb: Callback<void>) {
      withDraft(cb, (m) => {
        mailStore.update(m.id, { bodyHtml: html + m.bodyHtml });
      });
    },
  },
  subject: {
    getAsync(cb: Callback<string>) {
      withDraft(cb, (m) => m.subject);
    },
    setAsync(subject: string, cb: Callback<void>) {
      withDraft(cb, (m) => {
        mailStore.update(m.id, { subject });
      });
    },
  },
  to: recipients("to"),
  cc: recipients("cc"),
  bcc: recipients("bcc"),
  sessionData: {
    getAsync(key: string, cb: Callback<string>) {
      withDraft(cb, (m) => m.customProperties[`session:${key}`] || "");
    },
    setAsync(key: string, value: string, cb: Callback<void>) {
      withDraft(cb, (m) => {
        mailStore.update(m.id, (cur) => ({ customProperties: { ...cur.customProperties, [`session:${key}`]: value } }));
      });
    },
  },
  loadCustomPropertiesAsync(cb: Callback<unknown>) {
    withDraft(cb, (m) => {
      const pending: Record<string, string> = {};
      return {
        get: (key: string) => pending[key] ?? m.customProperties[key],
        set: (key: string, value: string) => {
          pending[key] = value;
        },
        saveAsync: (saveCb: Callback<void>) => {
          mailStore.update(m.id, (cur) => ({ customProperties: { ...cur.customProperties, ...pending } }));
          ok(saveCb, undefined);
        },
      };
    });
  },
  sendAsync(cb: Callback<void>) {
    withDraft(cb, (m) => {
      sendListeners.forEach((l) => l(m.id));
    });
  },
};

const sendListeners = new Set<(id: string) => void>();

/** The shell registers here to run its own send flow when the pane calls item.sendAsync(). */
export function onPaneSend(listener: (id: string) => void): () => void {
  sendListeners.add(listener);
  return () => sendListeners.delete(listener);
}

export function installFakeOffice(): void {
  const fake = {
    context: { mailbox: { item } },
    AsyncResultStatus: { Succeeded: SUCCEEDED, Failed: FAILED },
    CoercionType: { Html: "html", Text: "text" },
    onReady: (cb?: () => void) => {
      cb?.();
      return Promise.resolve({ host: null, platform: null });
    },
    actions: { associate: () => {} },
  };
  (globalThis as unknown as { Office: unknown }).Office = fake;
}
