// In-memory mailbox for the prototype. The initial inbox, the contacts and the user come from
// be-mock/samples.json, so everything is fictitious and editable by hand.

import samples from "../be-mock/samples.json";

export type Address = { name: string; email: string };

export type Folder = "inbox" | "drafts" | "sent";

export type Message = {
  id: string;
  folder: Folder;
  from: Address;
  to: Address[];
  cc: Address[];
  bcc: Address[];
  subject: string;
  bodyHtml: string;
  date: Date;
  unread: boolean;
  headers: Record<string, string>;
  customProperties: Record<string, string>;
};

export const ME: Address = samples.me;

let seq = 0;
export function nextId(prefix = "msg"): string {
  seq += 1;
  return `${prefix}-${seq.toString().padStart(4, "0")}`;
}

function at(daysAgo: number, time: string): Date {
  const [hours, minutes] = time.split(":").map(Number);
  const d = new Date();
  d.setDate(d.getDate() - daysAgo);
  d.setHours(hours, minutes, 0, 0);
  return d;
}

type InboxSample = {
  from: string;
  subject: string;
  daysAgo: number;
  time: string;
  unread: boolean;
  body: string[];
  ideaId?: string;
};

export const INITIAL_MESSAGES: Message[] = (samples.inbox as InboxSample[]).map((m) => ({
  id: nextId(),
  folder: "inbox",
  from: samples.contacts.find((c) => c.email === m.from) ?? { name: m.from, email: m.from },
  to: [ME],
  cc: [],
  bcc: [],
  subject: m.subject,
  bodyHtml: m.body.map((p) => `<p>${p}</p>`).join(""),
  date: at(m.daysAgo, m.time),
  unread: m.unread,
  headers: (m.ideaId ? { "x-idea-id": m.ideaId } : {}) as Record<string, string>,
  customProperties: {},
}));

type Listener = () => void;

export class MailStore {
  messages: Message[] = INITIAL_MESSAGES.map((m) => ({ ...m }));
  private listeners = new Set<Listener>();

  subscribe = (listener: Listener): (() => void) => {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  };

  private notify() {
    this.messages = [...this.messages];
    this.listeners.forEach((l) => l());
  }

  getSnapshot = () => this.messages;

  byFolder(folder: Folder): Message[] {
    return this.messages.filter((m) => m.folder === folder).sort((a, b) => b.date.getTime() - a.date.getTime());
  }

  get(id: string): Message | undefined {
    return this.messages.find((m) => m.id === id);
  }

  update(id: string, patch: Partial<Message> | ((m: Message) => Partial<Message>)): void {
    const m = this.get(id);
    if (!m) return;
    const next = { ...m, ...(typeof patch === "function" ? patch(m) : patch) };
    this.messages = this.messages.map((x) => (x.id === id ? next : x));
    this.notify();
  }

  createDraft(): Message {
    const draft: Message = {
      id: nextId("draft"),
      folder: "drafts",
      from: ME,
      to: [],
      cc: [],
      bcc: [],
      subject: "",
      bodyHtml: "<p></p>",
      date: new Date(),
      unread: false,
      headers: {},
      customProperties: {},
    };
    this.messages.push(draft);
    this.notify();
    return draft;
  }

  discard(id: string): void {
    this.messages = this.messages.filter((m) => m.id !== id);
    this.notify();
  }

  /** Moves the draft to Sent and delivers a copy to the Inbox, as if a recipient were reading it. */
  send(id: string): void {
    const draft = this.get(id);
    if (!draft) return;
    const now = new Date();
    const sent: Message = { ...draft, folder: "sent", date: now };
    this.messages = this.messages.map((x) => (x.id === id ? sent : x));
    const delivered: Message = {
      ...sent,
      id: nextId(),
      folder: "inbox",
      date: now,
      unread: true,
      headers: { ...draft.headers },
      customProperties: {},
    };
    this.messages.push(delivered);
    this.notify();
  }
}

export const mailStore = new MailStore();
