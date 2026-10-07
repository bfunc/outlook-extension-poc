import { useEffect, useState } from "react";
import { KeyElementsView } from "../pane/KeyElements";
import type { KeyElements } from "../pane/keyElementsModel";
import { API } from "../pane/office";

type Draft = { id: string; status: string; keyElements: KeyElements | null; sentAt?: string };

/** The pane as a recipient sees it: the key elements sent with this mail, looked up by x-idea-id. */
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
      <h1>Key Elements</h1>
      <p className="note">Read mode: what was sent with this message, looked up by its <code>x-idea-id</code> header.</p>
      {!ideaId && <p className="note">This message carries no x-idea-id header, so there is nothing to show.</p>}
      {state.loading && <p className="note">Loading…</p>}
      {state.error && <p className="status error">{state.error}</p>}
      {state.draft && (state.draft.keyElements ? <KeyElementsView value={state.draft.keyElements} /> : <p className="note">The sender saved no key elements.</p>)}
      {ideaId && (
        <footer>
          <p className="note">
            x-idea-id <code>{ideaId}</code>
            {state.draft?.sentAt && <> · sent {new Date(state.draft.sentAt).toLocaleString()}</>}
          </p>
        </footer>
      )}
    </main>
  );
}
