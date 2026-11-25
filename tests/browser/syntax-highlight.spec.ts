import { expect, test } from "@playwright/test";
import { resolve } from "node:path";

declare global {
  interface Window {
    __highlightPerformanceProbe: {
      active: boolean;
      previousFrame: number;
      maxFrameGapMs: number;
      maxLongTaskMs: number;
      longTaskCount: number;
    };
  }
}

test("large VBA highlighting avoids multi-second main-thread stalls", async ({ page }) => {
  test.setTimeout(240_000);

  await page.addInitScript(() => {
    const probe = {
      active: false,
      previousFrame: 0,
      maxFrameGapMs: 0,
      maxLongTaskMs: 0,
      longTaskCount: 0,
    };
    window.__highlightPerformanceProbe = probe;

    if ("PerformanceObserver" in window) {
      try {
        new PerformanceObserver((list) => {
          if (!probe.active) return;
          for (const entry of list.getEntries()) {
            probe.longTaskCount++;
            probe.maxLongTaskMs = Math.max(probe.maxLongTaskMs, entry.duration);
          }
        }).observe({ type: "longtask", buffered: true });
      } catch {}
    }

    const sampleFrame = (timestamp: number) => {
      if (probe.active) {
        if (probe.previousFrame) {
          probe.maxFrameGapMs = Math.max(probe.maxFrameGapMs, timestamp - probe.previousFrame);
        }
        probe.previousFrame = timestamp;
      }
      requestAnimationFrame(sampleFrame);
    };
    requestAnimationFrame(sampleFrame);
  });

  await page.goto("/");
  await page.locator("#file-input").setInputFiles(
    resolve("samples/13 Packs Solitaire Excel Edition.xlsm"),
  );
  await expect(page.locator("#open-vba-source")).toBeVisible({ timeout: 180_000 });
  await page.locator("#open-vba-source").click();
  await expect.poll(
    () => page.locator("#vba-source-output").evaluate((element) => element.textContent?.length ?? 0),
    { timeout: 30_000 },
  ).toBeGreaterThan(50_000);

  await page.evaluate(() => {
    const probe = window.__highlightPerformanceProbe;
    probe.active = true;
    probe.previousFrame = 0;
    probe.maxFrameGapMs = 0;
    probe.maxLongTaskMs = 0;
    probe.longTaskCount = 0;
  });
  await page.locator("#vba-source-highlight").click();
  const status = page.locator("#vba-source-highlight-status");
  await expect(status).toContainText("max frame gap", { timeout: 60_000 });

  const measurements = await page.evaluate(() => {
    window.__highlightPerformanceProbe.active = false;
    return {
      summary: document.querySelector("#vba-source-highlight-status")?.textContent,
      ...window.__highlightPerformanceProbe,
    };
  });
  console.log("Large VBA syntax highlighting performance:", measurements);

  expect(measurements.summary).toMatch(/\d[\d,]* chars/);
  expect(measurements.maxFrameGapMs).toBeLessThan(2_000);
  expect(measurements.maxLongTaskMs).toBeLessThan(2_000);
});