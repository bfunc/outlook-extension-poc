import { createRoot } from "react-dom/client";
import { installFakeOffice } from "./fakeOffice";
import { Shell } from "./Shell";
import "../src/app.css";
import "./shell.css";

// Installed before the task pane code runs, so src/office.ts sees a mailbox item.
installFakeOffice();

createRoot(document.getElementById("root")!).render(<Shell />);
