// Playwright config. Serves the built add-in (dist/) via a tiny static server.
// Run with: npm run build && npm run test:e2e
import { defineConfig } from "@playwright/test";

const PORT = 8214;

export default defineConfig({
  testDir: "./e2e",
  fullyParallel: false,
  forbidOnly: !!process.env.CI,
  retries: process.env.CI ? 1 : 0,
  workers: 1,
  reporter: process.env.CI ? [["github"], ["list"]] : "list",
  use: {
    baseURL: `http://localhost:${PORT}`,
    headless: true,
  },
  webServer: {
    command: "node scripts/serve-dist.mjs",
    port: PORT,
    reuseExistingServer: !process.env.CI,
    timeout: 30_000,
  },
});
