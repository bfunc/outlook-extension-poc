import { readFileSync, writeFileSync } from "node:fs";
import { resolve } from "node:path";
import { defineConfig, loadEnv, type Plugin } from "vite";
import react from "@vitejs/plugin-react";

// Host-specific values (URLs, add-in id, provider) come from .env, see .env.example.
const REQUIRED = ["ADDIN_BASE_URL", "ADDIN_ID", "ADDIN_PROVIDER", "ADDIN_SUPPORT_URL", "VITE_API_BASE"];

// Writes dist/manifest.xml from manifest.template.xml and puts the API URL into dist/launchevent.js
// (a plain public file that Vite copies untouched).
function addinFiles(env: Record<string, string>): Plugin {
  const values: Record<string, string> = { ...env, ADDIN_ORIGIN: new URL(env.ADDIN_BASE_URL).origin };
  return {
    name: "addin-files",
    apply: "build",
    closeBundle() {
      const manifest = readFileSync("manifest.template.xml", "utf8").replace(/\$\{(\w+)\}/g, (m, key) => {
        if (!(key in values)) throw new Error(`manifest.template.xml: no value for ${key}`);
        return values[key];
      });
      writeFileSync(resolve("dist", "manifest.xml"), manifest);
      const handler = resolve("dist", "launchevent.js");
      writeFileSync(handler, readFileSync(handler, "utf8").replace("__POC_API__", env.VITE_API_BASE));
    },
  };
}

export default defineConfig(({ mode }) => {
  const env = loadEnv(mode, process.cwd(), "");
  const missing = REQUIRED.filter((k) => !env[k]);
  if (missing.length) throw new Error(`Missing in .env: ${missing.join(", ")} (copy .env.example)`);
  return {
    base: new URL(env.ADDIN_BASE_URL).pathname.replace(/\/?$/, "/"),
    plugins: [react(), addinFiles(env)],
    build: { target: "es2020" },
  };
});
