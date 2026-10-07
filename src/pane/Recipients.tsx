import { useEffect, useState } from "react";
import { addTo, currentTo, fetchJson, inOutlook, type Contact } from "./office";

/** Step 2: the list of emails from the endpoint; a click adds the address to To. */
export function Recipients() {
  const outlook = inOutlook();
  const [contacts, setContacts] = useState<Contact[]>([]);
  const [added, setAdded] = useState<string[]>([]);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    fetchJson<Contact[]>("emails")
      .then(setContacts)
      .catch((e) => setError(`Could not load the list: ${(e as Error).message}`));
    if (outlook) currentTo().then(setAdded).catch(() => {});
  }, [outlook]);

  function add(c: Contact) {
    if (!outlook || added.includes(c.email)) return;
    setBusy(true);
    addTo(c)
      .then(() => setAdded((a) => [...a, c.email]))
      .catch((e) => setError((e as Error).message))
      .finally(() => setBusy(false));
  }

  return (
    <>
      <h1>Select recipients</h1>
      <p className="note">Click an address to add it to To.</p>
      <ul className="contacts">
        {contacts.map((c) => {
          const isAdded = added.includes(c.email);
          return (
            <li key={c.email}>
              <button type="button" className="contact" disabled={busy || isAdded || !outlook} onClick={() => add(c)}>
                <span className="name">{c.name}</span>
                <span className="email">{c.email}</span>
                <span className="tag">{isAdded ? "added" : "+ To"}</span>
              </button>
            </li>
          );
        })}
      </ul>
      {error && <p className="status error" role="status">{error}</p>}
    </>
  );
}
