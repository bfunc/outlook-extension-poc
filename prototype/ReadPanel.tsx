import { useEffect, useState } from "react";
import { API, type Item } from "../src/office";

type Draft = { id: string; status: string; items: Item[]; sentAt?: string };

/** The pane as a recipient sees it: the items that were sent with this mail, looked up by x-idea-id. */
export function ReadPanel({ ideaId }: { ideaId: string | undefined }) {
  const [state, setState] = useState<{ draft?: Draft; error?: string; loading: boolean }>({ loading: !!ideaId });

  useEffect(() => {
    if (!ideaId) return;
    let cancelled = false;
    setState({ loading: true });
    fetch(`${API}/drafts/${ideaId}`, { cache: "no-store" })
      .then(async (res) => {
        const data = await res.json();
        if (!res.ok || !data.ok) throw new Error(data.error || `HTTP ${res.status}`);
        if (!cancelled) setState({ loading: false, draft: data.draft });
      })
      .catch((e) => {
        if (!cancelled) setState({ loading: false, error: (e as Error).message });
      });
    return () => {
      cancelled = true;
    };
  }, [ideaId]);

  return (
    <main>
      <header>
        <h1>Outlook Extension PoC</h1>
        <p className="note">Read mode: items linked to this message through its <code>x-idea-id</code> header.</p>
      </header>
      <section>
        <h2>Items {state.draft && <span className="count">{state.draft.items.length}</span>}</h2>
        {!ideaId && <p className="hint">This message carries no x-idea-id header, so there is nothing to show.</p>}
        {state.loading && <p className="hint">Loading…</p>}
        {state.error && <p className="status error">{state.error}</p>}
        {state.draft && state.draft.items.length === 0 && <p className="hint">The sender selected no items.</p>}
        <ul className="list">
          {state.draft?.items.map((i) => (
            <li key={i.id} className="read-item">
              <strong>{i.title}</strong>
              <div dangerouslySetInnerHTML={{ __html: i.html }} />
            </li>
          ))}
        </ul>
      </section>
      {ideaId && (
        <footer>
          <p className="hint">
            x-idea-id <code>{ideaId}</code>
            {state.draft?.sentAt && <> · sent {new Date(state.draft.sentAt).toLocaleString()}</>}
          </p>
        </footer>
      )}
    </main>
  );
}
