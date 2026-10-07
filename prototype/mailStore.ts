// In-memory mailbox for the prototype. Everything here is fictitious: invented banks and desks,
// lorem ipsum text, example.com addresses.

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

export const ME: Address = { name: "Lorem Ipsum (you, fictitious)", email: "lorem.ipsum@example.com" };

const LOREM =
  "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.";
const LOREM2 =
  "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.";

function paragraphs(...lines: string[]): string {
  return lines.map((l) => `<p>${l}</p>`).join("");
}

function daysAgo(days: number, hour: number, minute = 0): Date {
  const d = new Date();
  d.setDate(d.getDate() - days);
  d.setHours(hour, minute, 0, 0);
  return d;
}

let seq = 0;
export function nextId(prefix = "msg"): string {
  seq += 1;
  return `${prefix}-${seq.toString().padStart(4, "0")}`;
}

function mail(
  from: Address,
  subject: string,
  body: string,
  date: Date,
  unread = false,
  extra: Partial<Message> = {},
): Message {
  return {
    id: nextId(),
    folder: "inbox",
    from,
    to: [ME],
    cc: [],
    bcc: [],
    subject,
    bodyHtml: body,
    date,
    unread,
    headers: {},
    customProperties: {},
    ...extra,
  };
}

const FX: Address = { name: "Dolor Sit (FX desk, fictitious)", email: "fx.desk@example.com" };
const SETTLEMENTS: Address = { name: "Amet Consectetur (Settlements, fictitious)", email: "settlements@example.com" };
const COMPLIANCE: Address = { name: "Adipiscing Elit (Compliance, fictitious)", email: "compliance@example.com" };
const RESEARCH: Address = { name: "Tempor Incididunt (Research, fictitious)", email: "research@example.com" };
const CLIENT: Address = { name: "Magna Aliqua (Lorem Holdings, fictitious)", email: "treasury@lorem-holdings.example.com" };
const RISK: Address = { name: "Veniam Quis (Market Risk, fictitious)", email: "market.risk@example.com" };
const ONBOARDING: Address = { name: "Nostrud Exercitation (Client Onboarding, fictitious)", email: "onboarding@example.com" };

export const INITIAL_MESSAGES: Message[] = [
  mail(
    CLIENT,
    "RE: FX hedge for Q3 receivables (fictitious)",
    paragraphs(
      `Dear Lorem,`,
      `Thank you for the indication. ${LOREM} Could you send the <strong>EUR/USD forward</strong> levels for ipsum 3 and 6 months?`,
      LOREM2,
      `Kind regards,<br/>Magna Aliqua<br/>Treasury, Lorem Holdings (fictitious)`,
    ),
    daysAgo(0, 9, 12),
    true,
  ),
  mail(
    SETTLEMENTS,
    "Settlement query: trade #LRM-4790 (fictitious)",
    paragraphs(
      `Hi,`,
      `Lorem ipsum settlement instructions for trade <strong>#LRM-4790</strong> are missing the dolor sit amet. ${LOREM2}`,
      `Please confirm by ipsum 16:00.`,
      `Settlements (fictitious)`,
    ),
    daysAgo(0, 8, 41),
    true,
  ),
  mail(
    RISK,
    "Margin call notice: Lorem Holdings (fictitious)",
    paragraphs(
      `${LOREM} The margin requirement for <strong>Lorem Holdings</strong> increased by ipsum 1.2m overnight.`,
      `Action required by dolor 12:00. ${LOREM2}`,
      `Market Risk (fictitious)`,
    ),
    daysAgo(0, 7, 55),
    true,
  ),
  mail(
    RESEARCH,
    "Morning note: rates and FX outlook (fictitious)",
    paragraphs(
      `<strong>Rates:</strong> ${LOREM}`,
      `<strong>FX:</strong> ${LOREM2}`,
      `<strong>Equities:</strong> Lorem ipsum dolor sit amet, consectetur adipiscing elit.`,
      `Research (fictitious). Not investment advice; this is sample text.`,
    ),
    daysAgo(0, 6, 30),
  ),
  mail(
    COMPLIANCE,
    "KYC refresh due: Lorem Holdings (fictitious)",
    paragraphs(
      `${LOREM} The KYC file for <strong>Lorem Holdings</strong> expires on dolor 30.`,
      `Please collect the ipsum documents listed at <a href="https://example.com/kyc">example.com/kyc</a>. ${LOREM2}`,
      `Compliance (fictitious)`,
    ),
    daysAgo(1, 16, 5),
  ),
  mail(
    FX,
    "Price indication EUR/USD 3M forward (fictitious)",
    paragraphs(
      `Indicative: <strong>1.0901 / 1.0905</strong>, lorem ipsum valid 15 minutes. ${LOREM2}`,
      `FX desk (fictitious)`,
    ),
    daysAgo(1, 14, 20),
  ),
  mail(
    ONBOARDING,
    "New client onboarding: Ipsum Capital (fictitious)",
    paragraphs(
      `${LOREM} Onboarding of <strong>Ipsum Capital</strong> is at step 3 of 5.`,
      `Outstanding: dolor sit amet form, consectetur signature. ${LOREM2}`,
      `Client Onboarding (fictitious)`,
    ),
    daysAgo(1, 11, 48),
  ),
  mail(
    SETTLEMENTS,
    "Trade confirmation #LRM-4821 (fictitious)",
    paragraphs(
      `Confirmed: buy lorem 5,000 ipsum shares at dolor 42.10, settlement T+2. ${LOREM}`,
      `Settlements (fictitious)`,
    ),
    daysAgo(2, 17, 2),
  ),
  mail(
    CLIENT,
    "Meeting request: portfolio review (fictitious)",
    paragraphs(
      `Dear Lorem,`,
      `${LOREM} Would ipsum Thursday 10:00 suit for the quarterly review? ${LOREM2}`,
      `Magna Aliqua (fictitious)`,
    ),
    daysAgo(2, 10, 15),
  ),
  mail(
    RISK,
    "Limit utilisation report, weekly (fictitious)",
    paragraphs(
      `${LOREM} Utilisation is at <strong>ipsum 68%</strong> of the dolor limit. ${LOREM2}`,
      `Market Risk (fictitious)`,
    ),
    daysAgo(3, 9, 0),
  ),
  mail(
    RESEARCH,
    "Sector update: lorem ipsum banks (fictitious)",
    paragraphs(`${LOREM}`, `${LOREM2}`, `Research (fictitious). Sample text only.`),
    daysAgo(4, 8, 10),
  ),
  mail(
    COMPLIANCE,
    "Reminder: annual attestation (fictitious)",
    paragraphs(`${LOREM} Please complete the ipsum attestation by dolor 15.`, `Compliance (fictitious)`),
    daysAgo(5, 15, 40),
  ),
];

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
