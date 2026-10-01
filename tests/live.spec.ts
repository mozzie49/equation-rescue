import { test, expect } from "@playwright/test";
import { readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { SafeZip } from "../src/core/zip";
import { SaxesParser } from "saxes";

const samplePath = fileURLToPath(new URL("../public/examples/teaching-handout.docx", import.meta.url));

test("deployed sample, selective downloads, upload and unchanged export", async ({ page }, testInfo) => {
  await page.goto("./");
  await expect(page.getByRole("heading", { level: 1 })).toContainText("Rescue the math");
  await page.getByRole("button", { name: "Try the teaching sample", exact: true }).click();
  await expect(page.locator("#filename")).toHaveText("teaching-handout.docx");
  await expect(page.locator(".equation-card")).toHaveCount(8);
  await expect(page.getByTestId("ready-count")).toHaveText("7");
  await expect(page.getByTestId("unsupported-count")).toHaveText("1");
  await page.getByRole("checkbox", { name: "Convert equation 1", exact: true }).uncheck();
  await expect(page.getByTestId("ready-count")).toHaveText("6");
  await expect(page.getByTestId("skipped-count")).toHaveText("1");
  await page.screenshot({ path: testInfo.outputPath("live-review.png"), fullPage: true });

  const docxPromise = page.waitForEvent("download");
  await page.getByRole("button", { name: "Download repaired DOCX", exact: true }).click();
  const docx = await docxPromise;
  expect(docx.suggestedFilename()).toBe("teaching-handout-repaired.docx");
  const docxPath = testInfo.outputPath("live-selective-repaired.docx");
  await docx.saveAs(docxPath);
  const output = new SafeZip(new Uint8Array(await readFile(docxPath)));
  const original = new SafeZip(new Uint8Array(await readFile(samplePath)));
  const xml = new TextDecoder().decode(output.read("word/document.xml"));
  let mathCount = 0;
  let documentText = "";
  const parser = new SaxesParser({ xmlns: true });
  parser.on("opentag", (tag) => {
    if (tag.local === "oMath" && tag.uri === "http://schemas.openxmlformats.org/officeDocument/2006/math") mathCount++;
  });
  parser.on("text", (text) => { documentText += text; });
  parser.write(xml).close();
  expect(mathCount).toBe(6);
  expect(documentText).toContain(String.raw`\(x^2 + 2x + 1\)`);
  expect(xml).toContain(String.raw`\begin{matrix}`);
  for (const entry of original.entries.filter((e) => e.name !== "word/document.xml")) {
    expect(output.entries.find((e) => e.name === entry.name)?.raw).toEqual(entry.raw);
  }

  const reportPromise = page.waitForEvent("download");
  await page.getByRole("button", { name: "Download JSON report", exact: true }).click();
  const reportDownload = await reportPromise;
  const reportPath = testInfo.outputPath("live-selective-report.json");
  await reportDownload.saveAs(reportPath);
  const report = JSON.parse(await readFile(reportPath, "utf8"));
  expect(report).toMatchObject({ converted: 6, skipped: 1, unsupported: 1, filename: "teaching-handout.docx" });

  await page.getByRole("button", { name: "中文", exact: true }).click();
  await expect(page.getByRole("heading", { name: "逐项检查公式", exact: true })).toBeVisible();
  await expect(page.getByTestId("ready-count")).toHaveText("6");
  await page.getByRole("button", { name: "重置", exact: true }).click();
  await page.getByRole("button", { name: "EN", exact: true }).click();
  await page.locator("#file-input").setInputFiles(samplePath);
  await expect(page.locator(".equation-card")).toHaveCount(8);
  await page.getByRole("button", { name: "Skip all", exact: true }).click();
  const unchangedPromise = page.waitForEvent("download");
  await page.getByRole("button", { name: "Download unchanged DOCX", exact: true }).click();
  const unchanged = await unchangedPromise;
  expect(await readFile((await unchanged.path())!)).toEqual(await readFile(samplePath));
});
