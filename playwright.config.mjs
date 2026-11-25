import { defineConfig } from "@playwright/test";

export default defineConfig({
  testDir: "./tests/browser",
  timeout: 240_000,
  expect: { timeout: 15_000 },
  fullyParallel: false,
  reporter: "list",
  use: {
    baseURL: "http://127.0.0.1:4173",
    browserName: "chromium",
    headless: true,
    trace: "retain-on-failure",
  },
  webServer: {
    command: "python3 -m http.server 4173 --bind 127.0.0.1 --directory site",
    url: "http://127.0.0.1:4173/",
    reuseExistingServer: !process.env.CI,
    timeout: 10_000,
  },
});