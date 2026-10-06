import { createRoot } from "react-dom/client";
import { App } from "./App";
import "./app.css";

const root = createRoot(document.getElementById("root")!);

// Office.onReady resolves outside Outlook too (host null), so the pane also opens in a plain browser.
if (typeof Office !== "undefined") Office.onReady(() => root.render(<App />));
else root.render(<App />);
