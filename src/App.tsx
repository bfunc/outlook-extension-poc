import { useEffect, useState } from "react";
import {
  addTo, appendSubject, applyItems, buildPayload, fetchJson, getSaved, inOutlook, setSession, submit, trySend,
  type Contact, type Item,
} from "./office";

type Status = { kind: "info" | "ok" | "error"; text: string } | null;

export function App() {
  const outlook = inOutlook();
  const [items, setItems] = useState<Item[]>([]);
  const [contacts, setContacts] = useState<Contact[]>([]);
  const [selected, setSelected] = useState<string[]>([]);
  const [added, setAdded] = useState<string[]>([]);
  const [busy, setBusy] = useState(false);
  const [status, setStatus] = useState<Status>(null);

  useEffect(() => {
    Promise.all([fetchJson<Item[]>("items"), fetchJson<Contact[]>("emails")])
      .then(([i, c]) => { setItems(i); setContacts(c); })
      .catch((e) => setStatus({ kind: "error", text: `Could not load the lists: ${e.message}` }));
    if (!outlook) return;
    // Marks this message for the OnMessageSend handler, which acts only on messages where the pane was used.
    setSession("pocUsed", "1").catch(() => {});
    // A reopened draft: restore the items selected last time.
    getSaved("pocItemIds")
      .then((v) => { const ids = JSON.parse(v || "[]"); if (Array.isArray(ids)) setSelected(ids); })
      .catch(() => {});
  }, [outlook]);

  async function run(label: string, fn: () => Promise<string | void>) {
    setBusy(true);
    setStatus({ kind: "info", text: label });
    try {
      const done = await fn();
      setStatus(done ? { kind: "ok", text: done } : null);
    } catch (e) {
      setStatus({ kind: "error", text: (e as Error).message || String(e) });
    } finally {
      setBusy(false);
    }
  }

  function toggleItem(id: string) {
    const next = selected.includes(id) ? selected.filter((s) => s !== id) : [...selected, id];
    setSelected(next);
    if (!outlook) return;
    const chosen = items.filter((i) => next.includes(i.id));
    run("Updating the message…", () => applyItems(chosen));
  }

  function addContact(c: Contact) {
    if (!outlook || added.includes(c.email)) return;
    run(`Adding ${c.email}…`, async () => {
      await addTo(c);
      setAdded((a) => [...a, c.email]);
      return `${c.email} added to To`;
    });
  }

  function send() {
    run("Sending to the test endpoint…", async () => {
      const payload = await buildPayload(selected);
      if (payload.recipients.length === 0) throw new Error("Add at least one recipient first.");
      const { id } = await submit(payload);
      await setSession("pocSubmitted", "1");
      await appendSubject();
      const sent = await trySend();
      return sent
        ? `Posted (${id.slice(0, 8)}) and sent.`
        : `Posted (${id.slice(0, 8)}), subject updated. Now press Send in Outlook.`;
    });
  }

  return (
    <main>
      <header>
        <h1>Outlook Extension PoC</h1>
        {!outlook && <p className="note">Open this panel from a new message in Outlook. Lists are read-only here.</p>}
      </header>

      <section>
        <h2>Items <span className="count">{selected.length}/{items.length}</span></h2>
        <p className="hint">Selected items go to the top of the message.</p>
        <ul className="list">
          {items.map((i) => (
            <li key={i.id}>
              <label>
                <input type="checkbox" checked={selected.includes(i.id)} disabled={busy} onChange={() => toggleItem(i.id)} />
                <span>{i.title}</span>
              </label>
            </li>
          ))}
        </ul>
      </section>

      <section>
        <h2>List of emails</h2>
        <p className="hint">Click to add to To.</p>
        <ul className="list">
          {contacts.map((c) => {
            const isAdded = added.includes(c.email);
            return (
              <li key={c.email}>
                <button className="contact" disabled={busy || isAdded || !outlook} onClick={() => addContact(c)}>
                  <span className="name">{c.name}</span>
                  <span className="email">{c.email}</span>
                  <span className="tag">{isAdded ? "added" : "+ To"}</span>
                </button>
              </li>
            );
          })}
        </ul>
      </section>

      <footer>
        <button className="primary" disabled={busy || !outlook} onClick={send}>Send</button>
        {status && <p className={`status ${status.kind}`} role="status">{status.text}</p>}
      </footer>
    </main>
  );
}
