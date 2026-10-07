import { createRoot } from "react-dom/client";
import { API } from "../src/office";
import { installMockBackend } from "../src/be-mock";
import { installFakeOffice } from "./fakeOffice";
import { Shell } from "./Shell";
import "../src/app.css";
import "./shell.css";

// Both installed before the task pane code runs: src/office.ts then sees a mailbox item, and its
// fetch calls to the test endpoint are answered by the in-browser mock (src/be-mock).
installMockBackend(API);
installFakeOffice();

createRoot(document.getElementById("root")!).render(<Shell />);
