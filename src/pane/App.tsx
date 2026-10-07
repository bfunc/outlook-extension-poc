import { useEffect, useRef, useState } from "react";
import { KeyElementsForm, KeyElementsView } from "./KeyElements";
import { emptyKeyElements, parseKeyElements, type KeyElements } from "./keyElementsModel";
import { getSaved, inOutlook, setSession } from "./office";
import { Recipients } from "./Recipients";

type Mode = "edit" | "view";
type Tab = 0 | 1;

const TABS = ["Key elements", "Select recipients"];

export function App() {
  const outlook = inOutlook();
  const [ke, setKe] = useState<KeyElements>(emptyKeyElements);
  const [mode, setMode] = useState<Mode>("edit");
  const [tab, setTab] = useState<Tab>(0);
  const [ready, setReady] = useState(!outlook);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (!outlook) return;
    // Marks this message for the OnMessageSend handler, which acts only on messages where the pane was used.
    setSession("pocUsed", "1").catch(() => {});
    // A reopened draft: restore what was entered last time.
    Promise.all([getSaved("pocKeyElements"), getSaved("pocMode")])
      .then(([saved, savedMode]) => {
        const parsed = parseKeyElements(saved);
        if (parsed) setKe(parsed);
        if (savedMode === "view" && parsed) setMode("view");
      })
      .catch(() => {})
      .finally(() => setReady(true));
  }, [outlook]);

  // Every change is saved with the draft, after a short pause, so the send handler and the backend
  // see the latest state.
  const first = useRef(true);
  useEffect(() => {
    if (!outlook || !ready) return;
    if (first.current) {
      first.current = false;
      return;
    }
    const t = window.setTimeout(() => {
      Promise.all([setSession("pocKeyElements", JSON.stringify(ke)), setSession("pocMode", mode)]).catch((e) =>
        setError(`Could not save to the draft: ${(e as Error).message}`),
      );
    }, 300);
    return () => window.clearTimeout(t);
  }, [ke, mode, outlook, ready]);

  return (
    <main>
      <div className="steps" role="tablist" aria-label="Steps">
        {TABS.map((s, i) => (
          <button key={s} type="button" role="tab" aria-selected={tab === i} className={tab === i ? "active" : ""} onClick={() => setTab(i as Tab)}>
            <span className="num">{tab === i ? "●" : i + 1}</span>
            <span>{s}</span>
          </button>
        ))}
      </div>
      {!outlook && <p className="note">Open this panel from a new message in Outlook.</p>}
      {tab === 0 ? (
        <>
          <h1>Key Elements</h1>
          {mode === "edit" ? <KeyElementsForm value={ke} onChange={setKe} /> : <KeyElementsView value={ke} />}
          <footer>
            {mode === "edit" ? (
              <button className="primary" disabled={!outlook} onClick={() => setMode("view")}>Save</button>
            ) : (
              <button className="primary" disabled={!outlook} onClick={() => setMode("edit")}>Edit</button>
            )}
            {error && <p className="status error" role="status">{error}</p>}
          </footer>
        </>
      ) : (
        <Recipients />
      )}
    </main>
  );
}
