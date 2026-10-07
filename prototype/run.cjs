// Starts the test endpoint and the Vite dev server together for the clickable prototype.
// Usage: npm run prototype, then open http://localhost:5173/outlook-addin/prototype/index.html
const { spawn } = require("node:child_process");
const { mkdirSync } = require("node:fs");
const { resolve } = require("node:path");

const root = resolve(__dirname, "..");
mkdirSync(resolve(root, "server", "tmp"), { recursive: true });

const server = spawn(process.execPath, ["server/server.js"], {
  cwd: root,
  stdio: "inherit",
  env: {
    ...process.env,
    PORT: process.env.PORT || "8799",
    ALLOWED_ORIGINS: process.env.ALLOWED_ORIGINS || "http://localhost:5173",
    DATA_FILE: process.env.DATA_FILE || "server/tmp/submissions.jsonl",
  },
});

const vite = spawn(process.execPath, ["node_modules/vite/bin/vite.js", "--port", "5173", "--strictPort", "--open", "/outlook-addin/prototype/index.html"], {
  cwd: root,
  stdio: "inherit",
});

function stop() {
  server.kill();
  vite.kill();
  process.exit(0);
}
process.on("SIGINT", stop);
process.on("SIGTERM", stop);
server.on("exit", stop);
vite.on("exit", stop);
