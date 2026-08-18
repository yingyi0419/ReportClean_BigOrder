import fs from "node:fs/promises";
import { build } from "vite";

await fs.rm("dist", { recursive: true, force: true });
await build();
await fs.mkdir("dist/server", { recursive: true });
await fs.mkdir("dist/.openai", { recursive: true });
await fs.copyFile(".openai/hosting.json", "dist/.openai/hosting.json");
await fs.writeFile(
  "dist/server/index.js",
  `export default {
  async fetch(request, env) {
    if (env?.ASSETS?.fetch) return env.ASSETS.fetch(request);
    return new Response("Site assets unavailable", { status: 503 });
  }
};
`,
);
