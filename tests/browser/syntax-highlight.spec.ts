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

test("large VBA highlighting renders only the viewport and stays responsive", async ({ page }) => {
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
  const sourceChars = await expect.poll(
    () => page.locator("#vba-source-output").evaluate((element) => element.textContent?.length ?? 0),
    { timeout: 30_000 },
  ).toBeGreaterThan(50_000).then(() =>
    page.locator("#vba-source-output").evaluate((element) => element.textContent?.length ?? 0),
  );

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
  await expect(status).toContainText("Highlighting lines", { timeout: 60_000 });

  const output = page.locator("#vba-source-output");
  const initialViewport = await output.evaluate((element) => ({
    sourceChars: Number(element.getAttribute("data-source-chars")),
    totalLines: Number(element.getAttribute("data-source-lines")),
    renderedLines: Number(element.getAttribute("data-rendered-lines")),
    firstLine: Number(element.getAttribute("data-visible-line-start")),
    highlightedSpans: element.querySelectorAll(".syntax-token").length,
    domTextChars: element.textContent?.length ?? 0,
  }));
  expect(initialViewport.sourceChars).toBe(sourceChars);
  expect(initialViewport.renderedLines).toBeLessThan(initialViewport.totalLines);
  expect(initialViewport.highlightedSpans).toBeGreaterThan(0);
  expect(initialViewport.domTextChars).toBeLessThan(sourceChars / 2);

  await page.locator("#vba-source-output").evaluate((element) => {
    const scrollContainer = element.closest("pre");
    if (scrollContainer) scrollContainer.scrollTop = scrollContainer.scrollHeight;
  });
  await expect.poll(
    () => output.evaluate((element) => Number(element.getAttribute("data-visible-line-start"))),
  ).toBeGreaterThan(initialViewport.firstLine);
  const bottomViewport = await output.evaluate((element) => ({
    firstLine: Number(element.getAttribute("data-visible-line-start")),
    highlightedSpans: element.querySelectorAll(".syntax-token").length,
  }));
  expect(bottomViewport.highlightedSpans).toBeGreaterThan(0);

  const measurements = await page.evaluate(() => {
    window.__highlightPerformanceProbe.active = false;
    return {
      summary: document.querySelector("#vba-source-highlight-status")?.textContent,
      ...window.__highlightPerformanceProbe,
    };
  });
  console.log("Large VBA syntax highlighting performance:", measurements);

  expect(measurements.summary).toContain("Highlighting lines");
  expect(measurements.maxFrameGapMs).toBeLessThan(2_000);
  expect(measurements.maxLongTaskMs).toBeLessThan(2_000);
});