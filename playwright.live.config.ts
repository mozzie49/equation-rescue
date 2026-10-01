import { defineConfig, devices } from "@playwright/test";

if (!process.env.LIVE_URL) throw new Error("LIVE_URL is required for deployed-site checks");

export default defineConfig({
  testDir: "./tests",
  testMatch: "live.spec.ts",
  forbidOnly: true,
  retries: 0,
  workers: 1,
  reporter: "list",
  use: {
    baseURL: process.env.LIVE_URL,
    locale: "en-US",
    trace: "retain-on-failure",
    screenshot: "only-on-failure",
  },
  projects: [{ name: "deployed-chromium", use: { ...devices["Desktop Chrome"] } }],
});
