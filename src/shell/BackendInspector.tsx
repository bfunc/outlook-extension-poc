import { useSyncExternalStore } from "react";
import { db, getVersion, resetMockBackend, subscribe } from "../be-mock";

/** Shows the state of the in-browser mock backend: drafts, submissions, request log. */
export function BackendInspector({ onClose }: { onClose: () => void }) {
  useSyncExternalStore(subscribe, getVersion);
  const drafts = [...db.drafts.values()].reverse();
  const submissions = [...db.submissions].reverse();
  const log = [...db.log].reverse().slice(0, 30);

  return (
    <section className="ol-inspector" aria-label="Mock backend">
      <div className="ol-inspector-head">
        <span>Backend (mock, in this browser tab) · src/be-mock</span>
        <span className="ol-cmd-spacer" />
        <button className="ol-link" onClick={resetMockBackend}>Reset</button>
        <button onClick={onClose} aria-label="Close inspector">×</button>
      </div>
      <div className="ol-inspector-cols">
        <div>
          <h3>Drafts ({drafts.length})</h3>
          {drafts.length === 0 && <p className="ol-empty">None yet. Press New mail.</p>}
          <table>
            <thead>
              <tr><th>x-idea-id</th><th>status</th><th>items</th><th>recipients</th><th>subject</th></tr>
            </thead>
            <tbody>
              {drafts.map((d) => (
                <tr key={d.id}>
                  <td><code>{d.id.slice(0, 8)}</code></td>
                  <td>{d.status}</td>
                  <td>{d.itemIds.join(", ") || "–"}</td>
                  <td>{d.recipients.join(", ") || "–"}</td>
                  <td>{d.subject || "–"}</td>
                </tr>
              ))}
            </tbody>
          </table>
          <h3>Submissions ({submissions.length})</h3>
          {submissions.length === 0 && <p className="ol-empty">None yet. Send a mail.</p>}
          <table>
            <thead>
              <tr><th>id</th><th>received</th><th>ideaId</th><th>items</th><th>recipients</th><th>email</th></tr>
            </thead>
            <tbody>
              {submissions.map((s) => (
                <tr key={s.id}>
                  <td><code>{s.id.slice(0, 8)}</code></td>
                  <td>{new Date(s.receivedAt).toLocaleTimeString()}</td>
                  <td><code>{s.ideaId ? s.ideaId.slice(0, 8) : "–"}</code></td>
                  <td>{s.itemIds.join(", ") || "–"}</td>
                  <td>{s.recipients.join(", ")}</td>
                  <td>{s.email.length} chars of HTML</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div>
          <h3>Requests</h3>
          <pre>{log.map((l) => `${l.at.slice(11, 19)}  ${l.line}`).join("\n") || "none yet"}</pre>
        </div>
      </div>
    </section>
  );
}
