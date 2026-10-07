import { useCallback, useEffect, useRef, useState, useSyncExternalStore } from "react";
import { App } from "../src/App";
import { API, appendSubject, buildPayload, submit } from "../src/office";
import { composingMessageId, onPaneSend, setComposingMessage } from "./fakeOffice";
import { ME, mailStore, type Address, type Folder, type Message } from "./mailStore";
import { ReadPanel } from "./ReadPanel";
import { BackendInspector } from "./BackendInspector";

const IDEA_HEADER = "x-idea-id";

type View = { kind: "read"; id: string } | { kind: "compose"; id: string } | { kind: "none" };

const FOLDERS: { id: Folder; label: string }[] = [
  { id: "inbox", label: "Inbox" },
  { id: "drafts", label: "Drafts" },
  { id: "sent", label: "Sent Items" },
];

function formatDate(d: Date): string {
  const today = new Date();
  const sameDay = d.toDateString() === today.toDateString();
  return sameDay
    ? d.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })
    : d.toLocaleDateString([], { weekday: "short", day: "numeric", month: "short" });
}

function initials(a: Address): string {
  return a.name
    .replace(/\(.*\)/, "")
    .trim()
    .split(/\s+/)
    .slice(0, 2)
    .map((w) => w[0]?.toUpperCase() ?? "")
    .join("");
}

function textOf(html: string): string {
  const div = document.createElement("div");
  div.innerHTML = html;
  return (div.textContent || "").replace(/\s+/g, " ").trim();
}

function parseAddresses(text: string): Address[] {
  return text
    .split(/[;,]/)
    .map((s) => s.trim())
    .filter(Boolean)
    .map((s) => {
      const m = /^(.*?)\s*<([^>]+)>$/.exec(s);
      return m ? { name: m[1] || m[2], email: m[2] } : { name: s, email: s };
    });
}

function joinAddresses(list: Address[]): string {
  return list.map((a) => (a.name && a.name !== a.email ? `${a.name} <${a.email}>` : a.email)).join("; ");
}

async function createBackendDraft(): Promise<string> {
  const res = await fetch(`${API}/drafts`, { method: "POST" });
  const data = await res.json();
  if (!res.ok || !data.ok) throw new Error(data.error || `HTTP ${res.status}`);
  return data.draft.id as string;
}

async function updateBackendDraft(id: string, patch: { itemIds?: string[]; recipients?: string[]; subject?: string }) {
  await fetch(`${API}/drafts/${id}`, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(patch),
  }).catch(() => {});
}

export function Shell() {
  const messages = useSyncExternalStore(mailStore.subscribe, mailStore.getSnapshot);
  const [folder, setFolder] = useState<Folder>("inbox");
  const [view, setView] = useState<View>({ kind: "none" });
  const [paneOpen, setPaneOpen] = useState(false);
  const [toast, setToast] = useState<string | null>(null);
  const [inspector, setInspector] = useState(false);

  const list = mailStore.byFolder(folder);
  const current = view.kind === "none" ? undefined : mailStore.get(view.id);
  const composing = view.kind === "compose" ? current : undefined;

  useEffect(() => {
    setComposingMessage(composing?.id);
  }, [composing?.id]);

  const say = useCallback((text: string) => {
    setToast(text);
    window.setTimeout(() => setToast(null), 4000);
  }, []);

  function openMessage(m: Message) {
    if (m.folder === "drafts") {
      setView({ kind: "compose", id: m.id });
      return;
    }
    if (m.unread) mailStore.update(m.id, { unread: false });
    setView({ kind: "read", id: m.id });
  }

  function newMail() {
    const draft = mailStore.createDraft();
    setView({ kind: "compose", id: draft.id });
    setPaneOpen(true);
    createBackendDraft()
      .then((ideaId) => mailStore.update(draft.id, (m) => ({ headers: { ...m.headers, [IDEA_HEADER]: ideaId } })))
      .catch((e) => say(`Could not create the draft on the server: ${(e as Error).message}`));
  }

  // The draft on the server follows the message: item selection (written by the pane into the
  // custom properties), recipients and subject.
  const synced = useRef<Record<string, string>>({});
  useEffect(() => {
    if (!composing) return;
    const ideaId = composing.headers[IDEA_HEADER];
    if (!ideaId) return;
    const patch = {
      itemIds: JSON.parse(composing.customProperties.pocItemIds || "[]") as string[],
      recipients: [...composing.to, ...composing.cc, ...composing.bcc].map((a) => a.email),
      subject: composing.subject,
    };
    const key = JSON.stringify(patch);
    if (synced.current[ideaId] === key) return;
    const timer = window.setTimeout(() => {
      synced.current[ideaId] = key;
      void updateBackendDraft(ideaId, patch);
    }, 400);
    return () => window.clearTimeout(timer);
  }, [composing]);

  // Outlook's Send: what launchevent.js does on OnMessageSend, then the message goes out.
  const sendMessage = useCallback(
    async (id: string) => {
      const m = mailStore.get(id);
      if (!m) return;
      if (m.to.length + m.cc.length + m.bcc.length === 0) {
        say("Add at least one recipient.");
        return;
      }
      const props = m.customProperties;
      if (props["session:pocUsed"] === "1" || props.pocUsed === "1") {
        try {
          await appendSubject();
          const ideaId = m.headers[IDEA_HEADER];
          if (props["session:pocSubmitted"] !== "1" && props.pocSubmitted !== "1") {
            const payload = await buildPayload(JSON.parse(props.pocItemIds || "[]"));
            await submit({ ...payload, ...(ideaId ? { ideaId } : {}) });
          } else if (ideaId) {
            await fetch(`${API}/drafts/${ideaId}`, {
              method: "PUT",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify({ status: "sent" }),
            });
          }
        } catch (e) {
          say(`Posting to the test endpoint failed, the message is sent anyway: ${(e as Error).message}`);
        }
      }
      mailStore.send(id);
      setView({ kind: "none" });
      setPaneOpen(false);
      setFolder("inbox");
      say("Message sent. A copy was delivered to your Inbox so you can open it as a recipient.");
    },
    [say],
  );

  useEffect(() => onPaneSend((id) => void sendMessage(id)), [sendMessage]);

  const unread = mailStore.byFolder("inbox").filter((m) => m.unread).length;

  return (
    <div className="ol">
      <div className="ol-banner">Prototype. Everything on this page is fictitious sample data.</div>
      <header className="ol-top">
        <div className="ol-waffle" aria-hidden="true">⋮⋮⋮</div>
        <div className="ol-brand">Outlook</div>
        <input className="ol-search" placeholder="Search" readOnly />
        <button className={`ol-top-btn${inspector ? " active" : ""}`} onClick={() => setInspector((v) => !v)}>Backend (mock)</button>
        <div className="ol-avatar" title={ME.email}>{initials(ME)}</div>
      </header>
      <div className="ol-body">
        <nav className="ol-rail" aria-label="Apps">
          <button className="active" title="Mail">✉</button>
          <button title="Calendar">▦</button>
          <button title="People">☺</button>
          <button title="Apps">⊞</button>
        </nav>
        <aside className="ol-folders">
          <button className="ol-new" onClick={newMail}>+ New mail</button>
          <div className="ol-account">{ME.email}</div>
          <ul>
            {FOLDERS.map((f) => (
              <li key={f.id}>
                <button className={folder === f.id ? "active" : ""} onClick={() => setFolder(f.id)}>
                  <span>{f.label}</span>
                  {f.id === "inbox" && unread > 0 && <span className="ol-badge">{unread}</span>}
                </button>
              </li>
            ))}
          </ul>
        </aside>
        <section className="ol-list" aria-label="Message list">
          <div className="ol-list-head">{FOLDERS.find((f) => f.id === folder)?.label}</div>
          {list.length === 0 && <p className="ol-empty">Nothing here.</p>}
          {list.map((m) => {
            const who = folder === "inbox" ? m.from : m.to[0] ?? { name: "(no recipient)", email: "" };
            return (
              <button
                key={m.id}
                className={`ol-row${m.unread ? " unread" : ""}${current?.id === m.id ? " selected" : ""}`}
                onClick={() => openMessage(m)}
              >
                <span className="ol-row-avatar">{initials(who)}</span>
                <span className="ol-row-text">
                  <span className="ol-row-from">{folder === "inbox" ? who.name : `To: ${who.name}`}</span>
                  <span className="ol-row-subject">{m.subject || "(no subject)"}</span>
                  <span className="ol-row-preview">{textOf(m.bodyHtml).slice(0, 90)}</span>
                </span>
                <span className="ol-row-meta">
                  <span>{formatDate(m.date)}</span>
                  {m.headers[IDEA_HEADER] && <span className="ol-chip" title={m.headers[IDEA_HEADER]}>idea</span>}
                </span>
              </button>
            );
          })}
        </section>
        <section className="ol-pane">
          {view.kind === "none" && <div className="ol-placeholder">Select an item to read, or press New mail.</div>}
          {view.kind === "read" && current && (
            <ReadView message={current} paneOpen={paneOpen} togglePane={() => setPaneOpen((v) => !v)} />
          )}
          {view.kind === "compose" && current && (
            <ComposeView
              message={current}
              paneOpen={paneOpen}
              togglePane={() => setPaneOpen((v) => !v)}
              onSend={() => void sendMessage(current.id)}
              onDiscard={() => {
                mailStore.discard(current.id);
                setView({ kind: "none" });
                setPaneOpen(false);
              }}
            />
          )}
        </section>
        {paneOpen && view.kind !== "none" && current && (
          <aside className="ol-addin" aria-label="Add-in pane">
            <div className="ol-addin-head">
              <span>Outlook Extension PoC</span>
              <button onClick={() => setPaneOpen(false)} aria-label="Close pane">×</button>
            </div>
            <div className="ol-addin-body">
              {view.kind === "compose" ? (
                <App key={current.id} />
              ) : (
                <ReadPanel key={current.id} ideaId={current.headers[IDEA_HEADER]} />
              )}
            </div>
          </aside>
        )}
      </div>
      {inspector && <BackendInspector onClose={() => setInspector(false)} />}
      {toast && <div className="ol-toast" role="status">{toast}</div>}
    </div>
  );
}

function AddinButton({ open, onClick }: { open: boolean; onClick: () => void }) {
  return (
    <button className={`ol-cmd ol-cmd-addin${open ? " active" : ""}`} onClick={onClick}>
      <span className="ol-cmd-icon">⚙</span> Outlook Extension PoC
    </button>
  );
}

function HeaderLine({ label, list }: { label: string; list: Address[] }) {
  if (list.length === 0) return null;
  return (
    <div className="ol-read-line">
      <span className="ol-read-label">{label}</span>
      <span>{list.map((a) => a.name).join("; ")}</span>
    </div>
  );
}

function ReadView({ message, paneOpen, togglePane }: { message: Message; paneOpen: boolean; togglePane: () => void }) {
  const [showHeaders, setShowHeaders] = useState(false);
  const ideaId = message.headers[IDEA_HEADER];
  return (
    <div className="ol-read">
      <div className="ol-cmdbar">
        <button className="ol-cmd">↩ Reply</button>
        <button className="ol-cmd">↪ Forward</button>
        <button className="ol-cmd">🗑 Delete</button>
        <span className="ol-cmd-spacer" />
        <AddinButton open={paneOpen} onClick={togglePane} />
      </div>
      <h1 className="ol-read-subject">{message.subject || "(no subject)"}</h1>
      <div className="ol-read-head">
        <span className="ol-row-avatar big">{initials(message.from)}</span>
        <div>
          <div className="ol-read-from">{message.from.name}</div>
          <HeaderLine label="To:" list={message.to} />
          <HeaderLine label="Cc:" list={message.cc} />
        </div>
        <div className="ol-read-date">{message.date.toLocaleString()}</div>
      </div>
      {ideaId && (
        <div className="ol-read-idea">
          Linked to backend draft <code>{ideaId}</code>
          <button className="ol-link" onClick={() => setShowHeaders((v) => !v)}>
            {showHeaders ? "Hide message headers" : "View message headers"}
          </button>
        </div>
      )}
      {showHeaders && (
        <pre className="ol-headers">
          {`From: ${joinAddresses([message.from])}\nTo: ${joinAddresses(message.to)}\nSubject: ${message.subject}\nDate: ${message.date.toUTCString()}\n${IDEA_HEADER}: ${ideaId}`}
        </pre>
      )}
      <div className="ol-read-body" dangerouslySetInnerHTML={{ __html: message.bodyHtml }} />
    </div>
  );
}

function ComposeView({
  message,
  paneOpen,
  togglePane,
  onSend,
  onDiscard,
}: {
  message: Message;
  paneOpen: boolean;
  togglePane: () => void;
  onSend: () => void;
  onDiscard: () => void;
}) {
  const bodyRef = useRef<HTMLDivElement>(null);
  const [toText, setToText] = useState(joinAddresses(message.to));
  const [ccText, setCcText] = useState(joinAddresses(message.cc));

  // The pane changes the message too (items block, To from the list), so the editor follows the
  // store, and the store follows the editor only when the user types.
  const storeTo = joinAddresses(message.to);
  const storeCc = joinAddresses(message.cc);
  const lastTo = useRef(storeTo);
  const lastCc = useRef(storeCc);
  useEffect(() => {
    if (storeTo !== lastTo.current) {
      lastTo.current = storeTo;
      setToText(storeTo);
    }
    if (storeCc !== lastCc.current) {
      lastCc.current = storeCc;
      setCcText(storeCc);
    }
  }, [storeTo, storeCc]);
  useEffect(() => {
    const el = bodyRef.current;
    if (el && el.innerHTML !== message.bodyHtml && document.activeElement !== el) el.innerHTML = message.bodyHtml;
  }, [message.bodyHtml]);

  function commitTo() {
    const to = parseAddresses(toText);
    lastTo.current = joinAddresses(to);
    mailStore.update(message.id, { to });
  }
  function commitCc() {
    const cc = parseAddresses(ccText);
    lastCc.current = joinAddresses(cc);
    mailStore.update(message.id, { cc });
  }

  return (
    <div className="ol-compose">
      <div className="ol-cmdbar">
        <button className="ol-cmd primary" onClick={onSend}>➤ Send</button>
        <button className="ol-cmd" onClick={onDiscard}>🗑 Discard</button>
        <button className="ol-cmd">📎 Attach</button>
        <span className="ol-cmd-spacer" />
        <AddinButton open={paneOpen} onClick={togglePane} />
      </div>
      <div className="ol-compose-fields">
        <label>
          <span>To</span>
          <input value={toText} onChange={(e) => setToText(e.target.value)} onBlur={commitTo} placeholder="Name <address@example.com>; …" />
        </label>
        <label>
          <span>Cc</span>
          <input value={ccText} onChange={(e) => setCcText(e.target.value)} onBlur={commitCc} />
        </label>
        <label>
          <span>Subject</span>
          <input value={message.subject} onChange={(e) => mailStore.update(message.id, { subject: e.target.value })} placeholder="Add a subject" />
        </label>
      </div>
      <div
        ref={bodyRef}
        className="ol-compose-body"
        contentEditable
        suppressContentEditableWarning
        onInput={(e) => mailStore.update(message.id, { bodyHtml: (e.target as HTMLDivElement).innerHTML })}
      />
      <div className="ol-compose-foot">
        Draft {message.id}
        {message.headers[IDEA_HEADER] ? (
          <> · linked to backend draft <code>{message.headers[IDEA_HEADER]}</code> (sent as header <code>{IDEA_HEADER}</code>)</>
        ) : (
          <> · creating the backend draft…</>
        )}
      </div>
    </div>
  );
}
