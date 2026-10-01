import { test, expect, type Page } from "@playwright/test";
import { readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { zipSync, strToU8 } from "fflate";
import { SafeZip } from "../src/core/zip";

const samplePath = fileURLToPath(
  new URL("../public/examples/teaching-handout.docx", import.meta.url),
);
const sampleName = "teaching-handout.docx";
const docxType =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
function fixture(text: string, extraParagraphXml = "") {
  const escaped = text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
  return Buffer.from(
    zipSync({
      "[Content_Types].xml": strToU8(
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
      ),
      "_rels/.rels": strToU8(
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
      ),
      "word/document.xml": strToU8(
        `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>${escaped}</w:t></w:r>${extraParagraphXml}</w:p><w:sectPr/></w:body></w:document>`,
      ),
    }),
  );
}
async function uploadSample(page: Page) {
  await page.locator("#file-input").setInputFiles(samplePath);
  await expect(page.locator("#filename")).toHaveText(sampleName);
  await expect(page.locator(".equation-card")).toHaveCount(8);
}
async function assertNoHorizontalOverflow(page: Page) {
  expect(
    await page.evaluate(
      () =>
        document.documentElement.scrollWidth <=
        document.documentElement.clientWidth,
    ),
  ).toBe(true);
}

test.beforeEach(async ({ page }) => {
  await page.goto("/");
});

test("upload, review, individual selection, select all and skip all", async ({
  page,
}) => {
  await expect(page.getByRole("heading", { level: 1 })).toContainText(
    "Rescue the math",
  );
  await expect(page.getByText("v0 · Experimental")).toBeVisible();
  await uploadSample(page);
  await expect(page.getByTestId("ready-count")).toHaveText("7");
  await expect(page.getByTestId("unsupported-count")).toHaveText("1");
  const supported = page.locator("input[data-equation]:enabled");
  await expect(supported).toHaveCount(7);
  await expect(page.locator("input[data-equation]:disabled")).toHaveCount(1);
  await expect(page.locator(".unsupported-reason")).not.toBeEmpty();
  await supported.first().uncheck();
  await expect(page.getByTestId("ready-count")).toHaveText("6");
  await expect(page.getByTestId("skipped-count")).toHaveText("1");
  await page.getByRole("button", { name: "Skip all", exact: true }).click();
  await expect(page.getByTestId("ready-count")).toHaveText("0");
  await expect(page.getByTestId("skipped-count")).toHaveText("7");
  await expect(
    page.getByRole("button", { name: "Download unchanged DOCX" }),
  ).toBeEnabled();
  await page
    .getByRole("button", { name: "Select supported", exact: true })
    .click();
  await expect(page.getByTestId("ready-count")).toHaveText("7");
  await expect(page.locator("input[data-equation]:checked")).toHaveCount(7);
});

test("export repaired DOCX and a separate report with source text", async ({
  page,
}) => {
  await uploadSample(page);
  await page.locator("input[data-equation]:enabled").first().uncheck();
  const docxDownload = page.waitForEvent("download");
  await page.getByRole("button", { name: "Download repaired DOCX" }).click();
  const docx = await docxDownload;
  expect(docx.suggestedFilename()).toBe("teaching-handout-repaired.docx");
  const output = new SafeZip(
    new Uint8Array(await readFile((await docx.path())!)),
  );
  const xml = new TextDecoder().decode(output.read("word/document.xml"));
  expect(xml).toContain("<m:oMath");
  const original = new SafeZip(new Uint8Array(await readFile(samplePath)));
  for (const entry of original.entries.filter(
    (e) => e.name !== "word/document.xml",
  )) {
    expect(output.entries.find((e) => e.name === entry.name)?.raw).toEqual(
      entry.raw,
    );
  }
  const reportDownload = page.waitForEvent("download");
  await page.getByRole("button", { name: "Download JSON report" }).click();
  const reportFile = await reportDownload;
  expect(reportFile.suggestedFilename()).toBe(
    "teaching-handout-equation-report.json",
  );
  const report = JSON.parse(await readFile((await reportFile.path())!, "utf8"));
  expect(report).toMatchObject({
    converted: 6,
    skipped: 1,
    unsupported: 1,
    filename: sampleName,
  });
  expect(
    report.equations.every(
      (equation: { source: string }) =>
        typeof equation.source === "string" && equation.source.length > 0,
    ),
  ).toBe(true);
  await expect(
    page.getByText(
      "The report includes the filename and original equation source text. Review it before sharing.",
    ),
  ).toBeVisible();
});

test("skip all exports a byte-identical separate copy", async ({ page }) => {
  await uploadSample(page);
  await page.getByRole("button", { name: "Skip all", exact: true }).click();
  const downloadPromise = page.waitForEvent("download");
  await page.getByRole("button", { name: "Download unchanged DOCX" }).click();
  const download = await downloadPromise;
  expect(download.suggestedFilename()).toBe("teaching-handout-unchanged.docx");
  expect(await readFile((await download.path())!)).toEqual(
    await readFile(samplePath),
  );
});

test("sample, language switch, reset, and repeated file selection stay coherent", async ({
  page,
}) => {
  await page.getByRole("button", { name: "Try the teaching sample" }).click();
  await expect(page.locator("#filename")).toHaveText(sampleName);
  await page.locator("input[data-equation]:enabled").first().uncheck();
  await page.getByRole("button", { name: "中文", exact: true }).click();
  await expect(page.locator("html")).toHaveAttribute("lang", "zh-CN");
  await expect(
    page.getByRole("heading", { name: "逐项检查公式", exact: true }),
  ).toBeVisible();
  await expect(page.getByTestId("ready-count")).toHaveText("6");
  await expect(
    page.getByRole("button", { name: "下载修复后的 DOCX" }),
  ).toBeVisible();
  await page.getByRole("button", { name: "重置", exact: true }).click();
  await expect(
    page.getByRole("heading", { name: "从你的文档开始" }),
  ).toBeVisible();
  await expect(page.locator(".equation-card")).toHaveCount(0);
  await page.getByRole("button", { name: "EN", exact: true }).click();
  await uploadSample(page);
  await page.getByRole("button", { name: "Reset", exact: true }).click();
  await uploadSample(page);
  await expect(page.getByTestId("ready-count")).toHaveText("7");
});

test("wrong extension, malformed DOCX and oversize files are rejected explicitly", async ({
  page,
}) => {
  await page
    .locator("#file-input")
    .setInputFiles({
      name: "notes.txt",
      mimeType: "text/plain",
      buffer: Buffer.from("hello"),
    });
  await expect(page.getByRole("alert")).toContainText(
    "Please choose a .docx file",
  );
  await page
    .locator("#file-input")
    .setInputFiles({
      name: "broken.docx",
      mimeType: docxType,
      buffer: Buffer.from("not a zip"),
    });
  await expect(page.getByRole("alert")).toContainText(
    "This document could not be inspected",
  );
  await page
    .locator("#file-input")
    .setInputFiles({
      name: "oversize.docx",
      mimeType: docxType,
      buffer: Buffer.alloc(20 * 1024 * 1024 + 1),
    });
  await expect(page.getByRole("alert")).toContainText("larger than 20 MB");
  await expect(page.locator(".equation-card")).toHaveCount(0);
  await uploadSample(page);
  await expect(page.getByRole("alert")).toHaveCount(0);
});

test("no-candidate document has a clear empty state", async ({ page }) => {
  await page
    .locator("#file-input")
    .setInputFiles({
      name: "ordinary.docx",
      mimeType: docxType,
      buffer: fixture("Just ordinary text with $single dollars$."),
    });
  await expect(
    page.getByRole("heading", { name: "No complete equations found" }),
  ).toBeVisible();
  await expect(page.getByTestId("ready-count")).toHaveText("0");
  await expect(
    page.getByRole("button", { name: "Download unchanged DOCX" }),
  ).toBeEnabled();
});

test("untrusted source is literal text and complex paragraphs stay unchanged", async ({
  page,
}) => {
  const unsafe = String.raw`\(\unknown{<img src=x onerror=alert(1)>}\)`;
  await page
    .locator("#file-input")
    .setInputFiles({
      name: "<img src=x>.docx",
      mimeType: docxType,
      buffer: fixture(unsafe),
    });
  await expect(page.locator(".equation-source pre")).toHaveText(unsafe);
  await expect(page.locator("#filename")).toHaveText("<img src=x>.docx");
  await expect(page.locator(".equation-card img, #filename img")).toHaveCount(
    0,
  );
  await expect(page.locator("input[data-equation]")).toBeDisabled();
  await expect(page.locator(".unsupported-reason")).toContainText(
    "Unsupported command",
  );
  await page
    .locator("#file-input")
    .setInputFiles({
      name: "bookmark.docx",
      mimeType: docxType,
      buffer: fixture(
        String.raw`\(x^2\)`,
        '<w:bookmarkStart w:id="0" w:name="example"/><w:bookmarkEnd w:id="0"/>',
      ),
    });
  await expect(page.locator("input[data-equation]")).toBeDisabled();
  await expect(page.locator(".unsupported-reason")).toContainText("bookmarks");
  await page.getByText("What v0 can safely handle", { exact: true }).click();
  await expect(
    page.getByText(
      "Paragraphs containing existing equations, fields, links, bookmarks, objects, or other complex structures are left unchanged.",
      { exact: false },
    ),
  ).toBeVisible();
});

test("canceling a pending sample cannot replace a newer uploaded file", async ({
  page,
}) => {
  let release!: () => void;
  const gate = new Promise<void>((resolve) => {
    release = resolve;
  });
  let intercepted!: () => void;
  const interception = new Promise<void>((resolve) => {
    intercepted = resolve;
  });
  await page.route("**/examples/teaching-handout.docx", async (route) => {
    intercepted();
    await gate;
    await route
      .fulfill({ body: await readFile(samplePath), contentType: docxType })
      .catch(() => {});
  });
  await page.getByRole("button", { name: "Try the teaching sample" }).click();
  await interception;
  await expect(
    page.getByRole("heading", { name: "Inspecting your document…" }),
  ).toBeVisible();
  await page.getByRole("button", { name: "Cancel", exact: true }).click();
  await page
    .locator("#file-input")
    .setInputFiles({
      name: "newer.docx",
      mimeType: docxType,
      buffer: fixture(String.raw`\(x^2\)`),
    });
  release();
  await expect(page.locator("#filename")).toHaveText("newer.docx");
  await expect(page.locator(".equation-card")).toHaveCount(1);
  await expect(page.getByTestId("ready-count")).toHaveText("1");
});

test("narrow phone layout and local-only requests", async ({ page }) => {
  const requests: string[] = [];
  page.on("request", (request) => requests.push(request.url()));
  await page.setViewportSize({ width: 375, height: 812 });
  await page.reload();
  await assertNoHorizontalOverflow(page);
  await page.getByRole("button", { name: "Try the teaching sample" }).click();
  await expect(page.locator(".equation-card")).toHaveCount(8);
  await assertNoHorizontalOverflow(page);
  await page.getByRole("button", { name: "中文", exact: true }).click();
  await assertNoHorizontalOverflow(page);
  expect(
    requests.filter(
      (url) =>
        !url.startsWith("http://127.0.0.1:4173/") &&
        !url.startsWith("blob:") &&
        !url.startsWith("data:"),
    ),
  ).toEqual([]);
  await expect(
    page.getByRole("heading", { name: "最后，请在 Word 中检查" }),
  ).toBeVisible();
});
