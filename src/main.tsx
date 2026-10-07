import { createRoot } from "react-dom/client";
import { API } from "./pane/office";
import { installMockBackend } from "./be-mock";
import { installFakeOffice } from "./shell/fakeOffice";
import { Shell } from "./shell/Shell";
import "./pane/app.css";
import "./shell/shell.css";

// Both installed before the task pane code runs: pane/office.ts then sees a mailbox item, and its
// fetch calls to the test endpoint are answered by the in-browser mock (src/be-mock).
installMockBackend(API);
installFakeOffice();

createRoot(document.getElementById("root")!).render(<Shell />);
