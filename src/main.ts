/// <reference types="vite/client" />
import katex from "katex";
import "katex/dist/katex.min.css";
import "./style.css";
import {
  inspectDocx,
  convertDocx,
  type Inspection,
  type Candidate,
} from "./core/docx";

type Language = "en" | "zh";
type Notice = {
  kind: "error" | "success";
  key:
    | "invalidType"
    | "tooLarge"
    | "readFailed"
    | "sampleFailed"
    | "exportFailed"
    | "docxSaved"
    | "reportSaved";
  detail?: string;
};
const MAX_FILE_BYTES = 20 * 1024 * 1024;
const DOCX_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const words = {
  en: {
    skip: "Skip to document workspace",
    tagline: "A careful repair for your documents",
    local: "Runs in your browser",
    experimental: "v0 · Experimental",
    language: "Language",
    eyebrow: "FROM LATEX TO WORD EQUATIONS",
    title: "Rescue the math.\nKeep the document.",
    intro:
      "Turn supported LaTeX into native Word equations inside your existing .docx. Review every change, then download a separate copy.",
    before: "BEFORE · SOURCE TEXT",
    after: "AFTER · APPROXIMATE PREVIEW",
    sampleCaption:
      "Your document. The same surrounding text. A real equation in place of raw LaTeX.",
    workflow: ["Open a DOCX", "Review the equations", "Download a new copy"],
    workspace: "Document workspace",
    openTitle: "Start with your document",
    openDescription: "Drop a .docx here, or choose a file to inspect.",
    choose: "Choose a DOCX",
    trySample: "Try the teaching sample",
    fileHint: "DOCX only · Up to 20 MB · Original stays unchanged",
    privacyTitle: "Your document stays on this device",
    privacy:
      "Files are processed in this browser. No account, analytics, uploads, or backend. The sample is the only document fetched from this site.",
    reading: "Inspecting your document…",
    readingHint:
      "Looking for complete LaTeX delimiters and checking what can be safely converted.",
    cancel: "Cancel",
    replace: "Change file",
    reset: "Reset",
    resetDone: "Document cleared. Choose another DOCX to begin.",
    review: "Review the equations",
    reviewDescription:
      "Supported equations are selected. Uncheck anything you want to keep as text.",
    found: (n: number, p: number) =>
      `${n} equation${n === 1 ? "" : "s"} found across ${p} body paragraph${p === 1 ? "" : "s"}`,
    ready: "Ready to convert",
    skipped: "Skipped by you",
    unsupported: "Unsupported",
    selectAll: "Select supported",
    skipAll: "Skip all",
    equation: (n: number) => `Equation ${n}`,
    selectEquation: (n: number) => `Convert equation ${n}`,
    paragraph: (n: number) => `Paragraph ${n}`,
    table: "Table cell",
    inline: "Inline",
    display: "Display",
    source: "ORIGINAL LATEX",
    preview: "APPROXIMATE PREVIEW",
    keepText: "Kept as original text",
    selected: "Will convert",
    notSelected: "Will keep as text",
    reason: "Why it stays as text",
    context: "Show paragraph context",
    previewUnavailable:
      "Browser preview unavailable. This does not change the converter’s support decision.",
    emptyTitle: "No complete equations found",
    emptyText:
      "This version looks for explicit LaTeX delimiters in ordinary body paragraphs and table cells. Plain math, single dollar signs, images, and equations in other document areas are left untouched.",
    zeroSelected:
      "Nothing is selected. The downloaded DOCX will be an unchanged copy.",
    exportTitle: "Make a separate copy",
    exportDescription:
      "Only selected equations are replaced. Untouched package parts are preserved, and your original file is never overwritten.",
    export: "Download repaired DOCX",
    exportUnchanged: "Download unchanged DOCX",
    exportBusy: "Preparing download…",
    report: "Download JSON report",
    reportPrivacy:
      "The report includes the filename and original equation source text. Review it before sharing.",
    verifyTitle: "One final check in Word",
    verify:
      "Browser previews are approximate, not Word rendering. Microsoft Word rendering and editability have not been verified. Open the downloaded file in Microsoft Word, check its layout, and click each repaired equation to confirm it is editable before relying on it.",
    limits: "What v0 can safely handle",
    limitedTitle: "A small, deliberate first version",
    limitedText:
      "Fractions, square roots, scripts, Greek letters, selected symbols, and simple sums/products. Unsupported syntax is kept as text with a reason; it is never silently guessed.",
    delimiters: "Explicit delimiters",
    delimiterHint:
      "Inline: \\(…\\) · Display: \\[…\\] or $$…$$. Display equations must occupy their own paragraph. Single $…$ is not scanned.",
    scope: "Document boundaries",
    scopeText:
      "Ordinary body paragraphs and table cells only. Headers, footers, notes, text boxes, comments, and existing equations are left untouched. Paragraphs containing existing equations, fields, links, bookmarks, objects, or other complex structures are left unchanged.",
    restrictions: "Not supported in v0",
    restrictionsText:
      "DOC, DOCM, PDFs, encrypted or digitally signed files, documents with tracked changes, and hidden text. Matrices, environments, custom macros, and unrecognized commands are not converted.",
    technical: "Technical detail",
    validation: "Package validation",
    validationText:
      "Export checks XML well-formedness and that raw ZIP records outside word/document.xml are unchanged. These checks do not substitute for opening the result in Word.",
    footer: "EquationRescue",
    footerDetail: "Repair in place. Download a new copy.",
    warnings: "Document notes",
    invalidType:
      "Please choose a .docx file. Other file types cannot be inspected.",
    tooLarge: "This file is larger than 20 MB. Choose a smaller DOCX.",
    readFailed:
      "This document could not be inspected. Choose a supported, unencrypted DOCX.",
    sampleFailed:
      "The sample could not be loaded. You can still choose your own DOCX.",
    exportFailed: "The copy could not be prepared. No output was downloaded.",
    docxSaved:
      "Your DOCX download is ready. Open the copy in Word and check the equations.",
    reportSaved:
      "Your JSON report download is ready. It includes original equation text.",
    multipleFiles: "Choose one DOCX at a time.",
    scanNote:
      "Only ordinary body paragraphs and table cells are scanned. Headers, footers, notes, text boxes and comments are left untouched.",
    previewNote:
      "Browser preview is approximate. Open the output in Word and check layout and equation editability before relying on it.",
    embeddedNote:
      "This document contains embedded files. They are preserved, not inspected or opened.",
    delimiterNote:
      "Incomplete delimiters or equations split across paragraphs are left untouched.",
    preservedNote:
      "Existing document links and embedded content are preserved. This tool is not a malware scanner; use trusted documents.",
    layoutNote:
      "Equations may change line heights and pagination. Surrounding formatting is preserved, but identical layout is not guaranteed.",
  },
  zh: {
    skip: "跳转到文档工作区",
    tagline: "谨慎修复，保留文档",
    local: "在浏览器中本地处理",
    experimental: "v0 · 实验版本",
    language: "语言",
    eyebrow: "从 LATEX 到 WORD 公式",
    title: "修复公式，\n保留原有文档。",
    intro:
      "将现有 .docx 中支持的 LaTeX 转换为原生 Word 公式。逐项确认后，下载一个新副本。",
    before: "转换前 · 源文本",
    after: "转换后 · 近似预览",
    sampleCaption: "保留原有文档与周围文字，将 LaTeX 源文本替换为公式。",
    workflow: ["打开 DOCX", "逐项检查公式", "下载新副本"],
    workspace: "文档工作区",
    openTitle: "从你的文档开始",
    openDescription: "将 .docx 拖到这里，或选择文件进行检查。",
    choose: "选择 DOCX",
    trySample: "试用教学示例",
    fileHint: "仅限 DOCX · 最大 20 MB · 原文件保持不变",
    privacyTitle: "文档始终留在此设备上",
    privacy:
      "文件仅在此浏览器中处理，无需账号，无分析追踪、上传或后端。只有示例文档会从本站获取。",
    reading: "正在检查文档…",
    readingHint: "查找完整的 LaTeX 定界符，并检查哪些公式可以安全转换。",
    cancel: "取消",
    replace: "更换文件",
    reset: "重置",
    resetDone: "文档已清除。请选择另一个 DOCX。",
    review: "逐项检查公式",
    reviewDescription: "已默认选中支持的公式。取消勾选即可保留其原始文本。",
    found: (n: number, p: number) => `在 ${p} 个正文段落中找到 ${n} 个公式`,
    ready: "待转换",
    skipped: "已跳过",
    unsupported: "暂不支持",
    selectAll: "选中支持的公式",
    skipAll: "全部跳过",
    equation: (n: number) => `公式 ${n}`,
    selectEquation: (n: number) => `转换公式 ${n}`,
    paragraph: (n: number) => `第 ${n} 段`,
    table: "表格单元格",
    inline: "行内公式",
    display: "独立公式",
    source: "原始 LATEX",
    preview: "近似预览",
    keepText: "保留原始文本",
    selected: "将转换",
    notSelected: "将保留文本",
    reason: "保留文本的原因",
    context: "查看段落上下文",
    previewUnavailable:
      "浏览器预览不可用。这不影响转换器对公式是否支持的判断。",
    emptyTitle: "未找到完整的公式",
    emptyText:
      "此版本仅检查普通正文段落和表格单元格中带有明确 LaTeX 定界符的内容。普通数学文本、单美元符号、图片及其他文档区域中的公式均保留原样。",
    zeroSelected: "当前未选中公式。下载的 DOCX 将是未经修改的副本。",
    exportTitle: "生成一个新副本",
    exportDescription:
      "仅替换选中的公式。其余文档包组件保持不变，原始文件不会被覆盖。",
    export: "下载修复后的 DOCX",
    exportUnchanged: "下载未修改的 DOCX",
    exportBusy: "正在准备下载…",
    report: "下载 JSON 报告",
    reportPrivacy: "报告包含文件名和公式源文本。分享前请先检查其中的内容。",
    verifyTitle: "最后，请在 Word 中检查",
    verify:
      "浏览器预览仅为近似效果，不是 Word 的实际渲染。尚未验证 Microsoft Word 中的实际显示及可编辑性。请用 Microsoft Word 打开下载的文件，检查排版，并逐个点击已修复的公式以确认可编辑，再正式使用。",
    limits: "v0 的安全处理范围",
    limitedTitle: "刻意保持精简的首个版本",
    limitedText:
      "支持分数、平方根、上下标、希腊字母、部分符号及简单的求和与连乘。不支持的语法将保留为文本并说明原因，不会擅自猜测。",
    delimiters: "明确的定界符",
    delimiterHint:
      "行内：\\(…\\)；独立：\\[…\\] 或 $$…$$。独立公式必须单独占据一个段落。不扫描单个 $…$。",
    scope: "文档处理范围",
    scopeText:
      "仅处理普通正文段落与表格单元格。页眉、页脚、注释、文本框、批注和现有公式保持原样。包含现有公式、域、链接、书签、对象或其他复杂结构的段落保持原样。",
    restrictions: "v0 暂不支持",
    restrictionsText:
      "DOC、DOCM、PDF、加密或数字签名文件，以及带修订或隐藏文本的文档。矩阵、环境、自定义宏及无法识别的命令不会被转换。",
    technical: "技术详情",
    validation: "文档包验证",
    validationText:
      "导出时检查 XML 格式，并确认 word/document.xml 之外的原始 ZIP 记录保持不变。这些检查不能代替在 Word 中打开文件进行验证。",
    footer: "EquationRescue",
    footerDetail: "在原文档中修复，下载独立副本。",
    warnings: "文档说明",
    invalidType: "请选择 .docx 文件。无法检查其他文件类型。",
    tooLarge: "此文件超过 20 MB，请选择较小的 DOCX。",
    readFailed: "无法检查此文档。请选择受支持且未加密的 DOCX。",
    sampleFailed: "无法加载示例。你仍然可以选择自己的 DOCX。",
    exportFailed: "无法生成副本，未下载任何输出文件。",
    docxSaved: "DOCX 已准备下载。请在 Word 中打开副本并检查公式。",
    reportSaved: "JSON 报告已准备下载，其中包含原始公式文本。",
    multipleFiles: "请一次选择一个 DOCX。",
    scanNote:
      "仅扫描普通正文段落和表格单元格。页眉、页脚、注释、文本框和批注保持不变。",
    previewNote:
      "浏览器预览仅为近似效果。请在 Word 中打开输出文件，检查排版及公式可编辑性后再使用。",
    embeddedNote: "此文档包含嵌入文件。这些文件将保留，但不会被检查或打开。",
    delimiterNote: "未闭合或跨段落的公式将保持不变。",
    preservedNote:
      "现有链接和嵌入内容将保留。本工具不进行恶意软件扫描，请仅使用可信文档。",
    layoutNote:
      "公式可能改变行高和分页。周围的格式将保留，但不能保证排版完全相同。",
  },
};
let language: Language = navigator.language.toLowerCase().startsWith("zh")
  ? "zh"
  : "en";
let inspection: Inspection | null = null;
let selected = new Set<string>();
let reading = false;
let exporting = false;
let operation = 0;
let sampleRequest: AbortController | null = null;
let notice: Notice | null = null;
let liveMessage = "";
const app = document.querySelector<HTMLDivElement>("#app")!;
const icon = (name: "file" | "arrow" | "shield" | "upload" | "check") => {
  const paths = {
    file: '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6M8 13h8M8 17h5"/>',
    arrow: '<path d="M4 12h16M14 6l6 6-6 6"/>',
    shield:
      '<path d="m12 3 8 3v6c0 5-8 9-8 9s-8-4-8-9V6z"/><path d="m8 12 3 3 5-6"/>',
    upload:
      '<path d="M12 16V3m-5 5 5-5 5 5M4 15v5a1 1 0 0 0 1 1h14a1 1 0 0 0 1-1v-5"/>',
    check: '<path d="m5 12 4 4L19 6"/>',
  };
  return `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${paths[name]}</svg>`;
};
function el<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  className?: string,
  text?: string,
): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text !== undefined) node.textContent = text;
  return node;
}
function counts() {
  const candidates = inspection?.candidates ?? [];
  const ready = candidates.filter(
    (c) => c.supported && selected.has(c.id),
  ).length;
  const unsupported = candidates.filter((c) => !c.supported).length;
  return {
    ready,
    unsupported,
    skipped: candidates.length - ready - unsupported,
  };
}
function renderMath(target: HTMLElement, latex: string, display = true) {
  try {
    katex.render(latex, target, {
      displayMode: display,
      throwOnError: true,
      strict: "error",
      trust: false,
      output: "htmlAndMathml",
      maxExpand: 100,
      maxSize: 10,
    });
  } catch {
    target.replaceChildren(
      el("p", "preview-fallback", words[language].previewUnavailable),
    );
  }
}
function render() {
  const t = words[language];
  const c = counts();
  document.documentElement.lang = language === "zh" ? "zh-CN" : "en";
  document.title =
    language === "zh"
      ? "EquationRescue · 修复 Word 公式"
      : "EquationRescue · Repair Word equations";
  app.innerHTML = `
    <a class="skip-link" href="#workspace">${t.skip}</a>
    <header class="site-header"><a class="brand" href="#" aria-label="EquationRescue"><span class="brand-mark" aria-hidden="true">E<span>r</span></span><span>Equation<span class="brand-light">Rescue</span></span></a><div class="header-tools"><span class="local-status"><span class="status-dot"></span>${t.local}</span><fieldset class="language-switch"><legend class="visually-hidden">${t.language}</legend><button type="button" data-lang="en" aria-pressed="${language === "en"}">EN</button><button type="button" data-lang="zh" aria-pressed="${language === "zh"}">中文</button></fieldset></div></header>
    <main>
      <section class="hero ${inspection || reading ? "hero-compact" : ""}" aria-labelledby="hero-title"><div class="hero-copy"><div class="eyebrow"><span>${t.eyebrow}</span><span class="version">${t.experimental}</span></div><h1 id="hero-title">${t.title.replace("\n", inspection || reading ? " " : "<br>")}</h1><p class="hero-intro">${t.intro}</p><ol class="workflow">${t.workflow.map((step, i) => `<li><span>${String(i + 1).padStart(2, "0")}</span>${step}</li>`).join("")}</ol></div><div class="transformation"><div class="before-example"><span class="small-label">${t.before}</span><pre id="example-source"></pre></div><div class="transformation-arrow">${icon("arrow")}</div><div class="after-example"><span class="small-label">${t.after}</span><div id="example-preview"></div></div><p class="example-caption">${t.sampleCaption}</p></div></section>
      <div class="workspace-shell"><section id="workspace" class="workspace" aria-label="${t.workspace}" tabindex="-1"><div id="notice-slot"></div><input id="file-input" type="file" accept=".docx,${DOCX_TYPE}" class="visually-hidden" aria-label="${t.choose}" tabindex="-1"><div id="workspace-content"></div></section><aside class="assurance"><div class="privacy-icon">${icon("shield")}</div><div><h2>${t.privacyTitle}</h2><p>${t.privacy}</p></div></aside></div>
      <section class="limits-section"><details id="limits"><summary>${t.limits}<span aria-hidden="true">+</span></summary><div class="limits-grid"><div><h3>${t.limitedTitle}</h3><p>${t.limitedText}</p></div><div><h3>${t.delimiters}</h3><p>${t.delimiterHint}</p></div><div><h3>${t.scope}</h3><p>${t.scopeText}</p></div><div><h3>${t.restrictions}</h3><p>${t.restrictionsText}</p></div><div><h3>${t.validation}</h3><p>${t.validationText}</p></div></div></details></section>
      <p class="experimental-note">${t.verify}</p>
    </main><footer><span>${t.footer} <span class="footer-version">/ v0.1</span></span><span>${t.footerDetail}</span></footer><div id="announcer" class="visually-hidden" aria-live="polite" aria-atomic="true"></div>`;
  const example = String.raw`x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}`;
  app.querySelector("#example-source")!.textContent =
    String.raw`\(${example}\)`;
  renderMath(app.querySelector("#example-preview")!, example);
  app.querySelectorAll<HTMLButtonElement>("[data-lang]").forEach((button) =>
    button.addEventListener("click", () => {
      language = button.dataset.lang as Language;
      render();
      app
        .querySelector<HTMLButtonElement>(`[data-lang="${language}"]`)
        ?.focus();
    }),
  );
  app.querySelector(".brand")!.addEventListener("click", (e) => {
    e.preventDefault();
    window.scrollTo({ top: 0, behavior: "smooth" });
  });
  app
    .querySelector<HTMLInputElement>("#file-input")!
    .addEventListener("change", (event) => {
      const input = event.target as HTMLInputElement;
      const file = input.files?.[0];
      input.value = "";
      if (file) void loadFile(file);
    });
  const content = app.querySelector<HTMLDivElement>("#workspace-content")!;
  if (reading) {
    content.innerHTML = `<div class="loading-state"><span class="spinner" aria-hidden="true"></span><h2>${t.reading}</h2><p>${t.readingHint}</p><button class="button button-secondary" id="cancel">${t.cancel}</button></div>`;
    content.querySelector("#cancel")!.addEventListener("click", reset);
  } else if (!inspection) {
    content.innerHTML = `<div id="dropzone" class="dropzone"><div class="upload-symbol">${icon("upload")}</div><h2>${t.openTitle}</h2><p>${t.openDescription}</p><div class="upload-actions"><button id="choose" class="button button-primary">${t.choose}${icon("arrow")}</button><button id="sample" class="button button-secondary">${t.trySample}</button></div><span class="file-hint">${t.fileHint}</span></div>`;
    content.querySelector("#choose")!.addEventListener("click", chooseFile);
    content
      .querySelector("#sample")!
      .addEventListener("click", () => void loadSample());
    const dropzone = content.querySelector<HTMLElement>("#dropzone")!;
    dropzone.addEventListener("dragover", (event) => {
      event.preventDefault();
      dropzone.classList.add("drag-over");
    });
    dropzone.addEventListener("dragleave", (event) => {
      if (!dropzone.contains(event.relatedTarget as Node | null))
        dropzone.classList.remove("drag-over");
    });
    dropzone.addEventListener("drop", (event) => {
      event.preventDefault();
      dropzone.classList.remove("drag-over");
      const files = event.dataTransfer?.files;
      if (files?.length === 1) void loadFile(files[0]);
      else if (files?.length) {
        notice = { kind: "error", key: "readFailed", detail: t.multipleFiles };
        renderNotice();
      }
    });
  } else {
    content.innerHTML = `<div class="document-bar"><div class="document-identity">${icon("file")}<div><p id="filename" class="filename"></p><p class="document-meta">${t.found(inspection.candidates.length, inspection.paragraphCount)}</p></div></div><div class="document-actions"><button id="replace" class="text-button">${t.replace}</button><span aria-hidden="true">/</span><button id="reset" class="text-button">${t.reset}</button></div></div><div class="review-header"><div><p class="section-number">02 / ${t.workspace}</p><h2>${t.review}</h2><p>${t.reviewDescription}</p></div></div><div class="counts"><div class="count-ready"><strong data-testid="ready-count">${c.ready}</strong><span>${t.ready}</span></div><div><strong data-testid="skipped-count">${c.skipped}</strong><span>${t.skipped}</span></div><div><strong data-testid="unsupported-count">${c.unsupported}</strong><span>${t.unsupported}</span></div></div><div id="review-list"></div><div id="document-notes"></div><div class="export-panel"><div class="export-copy"><span class="section-number">03 / DOCX + JSON</span><h2>${t.exportTitle}</h2><p>${t.exportDescription}</p><p id="zero-selection" class="zero-selection" ${c.ready ? "hidden" : ""}>${t.zeroSelected}</p></div><div class="export-actions"><button id="export-docx" class="button button-primary" ${exporting ? "disabled" : ""}>${exporting ? t.exportBusy : c.ready ? t.export : t.exportUnchanged}${icon("arrow")}</button><button id="export-report" class="button button-secondary" ${exporting ? "disabled" : ""}>${t.report}</button><p>${t.reportPrivacy}</p></div></div><aside class="verify-notice"><span class="verify-symbol" aria-hidden="true">!</span><div><h3>${t.verifyTitle}</h3><p>${t.verify}</p></div></aside>`;
    content.querySelector("#filename")!.textContent = inspection.filename;
    content.querySelector("#replace")!.addEventListener("click", chooseFile);
    content.querySelector("#reset")!.addEventListener("click", reset);
    content
      .querySelector("#export-docx")!
      .addEventListener("click", () => void downloadOutput("docx"));
    content
      .querySelector("#export-report")!
      .addEventListener("click", () => void downloadOutput("report"));
    const review = content.querySelector("#review-list")!;
    if (!inspection.candidates.length) {
      const empty = el("div", "no-candidates");
      empty.append(el("h3", "", t.emptyTitle), el("p", "", t.emptyText));
      review.append(empty);
    } else {
      const toolbar = el("div", "review-toolbar");
      const all = el("button", "text-button", t.selectAll);
      all.type = "button";
      all.id = "select-all";
      all.disabled =
        exporting || !inspection.candidates.some((x) => x.supported);
      const none = el("button", "text-button", t.skipAll);
      none.type = "button";
      none.id = "skip-all";
      none.disabled =
        exporting || !inspection.candidates.some((x) => x.supported);
      all.addEventListener("click", () => {
        selected = new Set(
          inspection!.candidates.filter((x) => x.supported).map((x) => x.id),
        );
        selectionChanged();
      });
      none.addEventListener("click", () => {
        selected.clear();
        selectionChanged();
      });
      toolbar.append(all, el("span", "toolbar-divider", "/"), none);
      const list = el("ol", "equation-list");
      inspection.candidates.forEach((candidate, index) =>
        list.append(candidateCard(candidate, index)),
      );
      review.append(toolbar, list);
    }
    renderDocumentNotes(content.querySelector("#document-notes")!);
  }
  renderNotice();
  app.querySelector("#announcer")!.textContent = reading
    ? t.reading
    : liveMessage;
}
function candidateCard(candidate: Candidate, index: number) {
  const t = words[language];
  const card = el(
    "li",
    `equation-card ${candidate.supported ? "supported" : "unsupported"} ${selected.has(candidate.id) ? "is-selected" : ""}`,
  );
  card.dataset.candidate = candidate.id;
  const top = el("div", "equation-top");
  const label = el("label", "equation-selection");
  const checkbox = el("input");
  checkbox.type = "checkbox";
  checkbox.checked = selected.has(candidate.id);
  checkbox.disabled = !candidate.supported || exporting;
  checkbox.setAttribute("aria-label", t.selectEquation(index + 1));
  checkbox.dataset.equation = candidate.id;
  checkbox.addEventListener("change", () => {
    if (checkbox.checked) selected.add(candidate.id);
    else selected.delete(candidate.id);
    selectionChanged();
  });
  label.append(checkbox, el("span", "equation-title", t.equation(index + 1)));
  const location = el("div", "equation-location");
  location.append(
    el("span", "", t.paragraph(candidate.paragraph)),
    el("span", "location-dot", "·"),
    el("span", "", candidate.display ? t.display : t.inline),
  );
  if (candidate.inTable) location.append(el("span", "table-tag", t.table));
  const badge = el(
    "span",
    `equation-status ${candidate.supported ? "status-supported" : "status-unsupported"}`,
    !candidate.supported
      ? t.unsupported
      : selected.has(candidate.id)
        ? t.selected
        : t.notSelected,
  );
  top.append(label, location, badge);
  const comparison = el("div", "equation-comparison");
  const source = el("div", "equation-source");
  source.append(
    el("span", "small-label", t.source),
    el("pre", "", candidate.source),
  );
  const preview = el("div", "equation-preview");
  preview.append(
    el("span", "small-label", candidate.supported ? t.preview : t.keepText),
  );
  if (candidate.supported) {
    const math = el("div", "math-preview");
    renderMath(math, candidate.latex, candidate.display);
    preview.append(math);
  } else {
    const reason = el(
      "p",
      "unsupported-reason",
      candidate.reason ?? t.unsupported,
    );
    reason.id = `reason-${candidate.id}`;
    checkbox.setAttribute("aria-describedby", reason.id);
    preview.append(el("span", "reason-label", t.reason), reason);
  }
  comparison.append(source, preview);
  const details = el("details", "context-details");
  details.append(
    el("summary", "", t.context),
    el("p", "context-text", candidate.context),
  );
  card.append(top, comparison, details);
  return card;
}
function renderDocumentNotes(target: Element) {
  if (!inspection?.warnings.length) return;
  const t = words[language];
  const details = el("details", "document-notes");
  details.append(el("summary", "", t.warnings));
  const list = el("ul");
  for (const warning of inspection.warnings) {
    let message = warning;
    if (warning.startsWith("Only ordinary body paragraphs"))
      message = t.scanNote;
    if (warning.startsWith("Browser preview is approximate"))
      message = t.previewNote;
    if (warning.startsWith("This document contains embedded files"))
      message = t.embeddedNote;
    if (
      warning.startsWith("No complete supported delimiter pairs") ||
      warning.startsWith("Some delimiters are incomplete")
    )
      message = t.delimiterNote;
    if (warning.startsWith("Existing document links"))
      message = t.preservedNote;
    if (warning.startsWith("Equations may change line heights"))
      message = t.layoutNote;
    list.append(el("li", "", message));
  }
  details.append(list);
  target.append(details);
}
function renderNotice() {
  const target = app.querySelector("#notice-slot")!;
  target.replaceChildren();
  if (!notice) return;
  const box = el("div", `notice notice-${notice.kind}`);
  box.setAttribute("role", notice.kind === "error" ? "alert" : "status");
  box.append(el("p", "", words[language][notice.key]));
  if (notice.detail) {
    const details = el("details");
    details.append(
      el("summary", "", words[language].technical),
      el("p", "technical-detail", notice.detail),
    );
    box.append(details);
  }
  target.append(box);
}
function selectionChanged() {
  operation++;
  notice = null;
  renderNotice();
  const t = words[language],
    c = counts();
  app.querySelector('[data-testid="ready-count"]')!.textContent = String(
    c.ready,
  );
  app.querySelector('[data-testid="skipped-count"]')!.textContent = String(
    c.skipped,
  );
  app
    .querySelectorAll<HTMLInputElement>("input[data-equation]")
    .forEach((input) => {
      const candidate = inspection!.candidates.find(
        (x) => x.id === input.dataset.equation,
      )!;
      input.checked = selected.has(candidate.id);
      const card = input.closest(".equation-card")!;
      card.classList.toggle("is-selected", input.checked);
      if (candidate.supported)
        card.querySelector(".equation-status")!.textContent = input.checked
          ? t.selected
          : t.notSelected;
    });
  app.querySelector<HTMLParagraphElement>("#zero-selection")!.hidden =
    c.ready !== 0;
  app.querySelector<HTMLButtonElement>("#export-docx")!.innerHTML =
    `${c.ready ? t.export : t.exportUnchanged}${icon("arrow")}`;
  liveMessage = `${t.ready}: ${c.ready}. ${t.skipped}: ${c.skipped}. ${t.unsupported}: ${c.unsupported}.`;
  app.querySelector("#announcer")!.textContent = liveMessage;
}
function chooseFile() {
  app.querySelector<HTMLInputElement>("#file-input")!.click();
}
function beginRead() {
  operation++;
  sampleRequest?.abort();
  sampleRequest = null;
  inspection = null;
  selected.clear();
  notice = null;
  liveMessage = "";
  reading = true;
  exporting = false;
  render();
  return operation;
}
function finishRead(bytes: Uint8Array, filename: string, current: number) {
  if (current !== operation) return;
  const result = inspectDocx(bytes, filename);
  if (current !== operation) return;
  inspection = result;
  selected = new Set(
    result.candidates.filter((c) => c.supported).map((c) => c.id),
  );
  reading = false;
  liveMessage = words[language].found(
    result.candidates.length,
    result.paragraphCount,
  );
  render();
  app.querySelector<HTMLElement>("#workspace")!.focus({ preventScroll: true });
}
async function loadFile(file: File) {
  const current = beginRead();
  if (!/\.docx$/i.test(file.name)) {
    reading = false;
    notice = { kind: "error", key: "invalidType" };
    render();
    return;
  }
  if (file.size > MAX_FILE_BYTES) {
    reading = false;
    notice = { kind: "error", key: "tooLarge" };
    render();
    return;
  }
  try {
    const bytes = new Uint8Array(await file.arrayBuffer());
    finishRead(bytes, file.name, current);
  } catch (error) {
    if (current !== operation) return;
    reading = false;
    notice = {
      kind: "error",
      key: "readFailed",
      detail: error instanceof Error ? error.message : undefined,
    };
    render();
  }
}
async function loadSample() {
  const current = beginRead();
  const controller = new AbortController();
  sampleRequest = controller;
  try {
    const response = await fetch(
      `${import.meta.env.BASE_URL}examples/teaching-handout.docx`,
      { signal: controller.signal, credentials: "same-origin" },
    );
    if (!response.ok) throw new Error(`Sample response: ${response.status}`);
    const bytes = new Uint8Array(await response.arrayBuffer());
    if (bytes.length > MAX_FILE_BYTES) throw new Error("Sample exceeds 20 MB");
    finishRead(bytes, "teaching-handout.docx", current);
  } catch (error) {
    if (current !== operation) return;
    reading = false;
    notice = {
      kind: "error",
      key: "sampleFailed",
      detail: error instanceof Error ? error.message : undefined,
    };
    render();
  } finally {
    if (current === operation) sampleRequest = null;
  }
}
function reset() {
  operation++;
  sampleRequest?.abort();
  sampleRequest = null;
  inspection = null;
  selected.clear();
  reading = false;
  exporting = false;
  notice = null;
  liveMessage = words[language].resetDone;
  render();
  app.querySelector<HTMLButtonElement>("#choose")?.focus();
}
function safeBaseName(filename: string) {
  return (
    filename
      .normalize("NFKC")
      .replace(/\.docx$/i, "")
      .replace(
        /[<>:"/\\|?*\u0000-\u001f\u007f\u202a-\u202e\u2066-\u2069]/g,
        "_",
      )
      .replace(/^\.+/, "")
      .trim()
      .slice(0, 100) || "document"
  );
}
function saveDownload(bytes: Uint8Array, type: string, filename: string) {
  const buffer = new ArrayBuffer(bytes.byteLength);
  new Uint8Array(buffer).set(bytes);
  const url = URL.createObjectURL(new Blob([buffer], { type }));
  const anchor = el("a");
  anchor.href = url;
  anchor.download = filename;
  anchor.hidden = true;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 60_000);
}
async function downloadOutput(kind: "docx" | "report") {
  if (!inspection || exporting) return;
  const currentInspection = inspection,
    currentSelection = [...selected],
    current = ++operation;
  exporting = true;
  notice = null;
  render();
  // Yield once so the busy state paints and a reset/new file can invalidate this export.
  await new Promise<void>((resolve) => window.setTimeout(resolve, 0));
  if (current !== operation) return;
  try {
    const result = convertDocx(currentInspection, currentSelection);
    if (current !== operation) return;
    const base = safeBaseName(currentInspection.filename);
    if (kind === "docx")
      saveDownload(
        result.bytes,
        DOCX_TYPE,
        `${base}-${result.report.converted ? "repaired" : "unchanged"}.docx`,
      );
    else
      saveDownload(
        new TextEncoder().encode(JSON.stringify(result.report, null, 2)),
        "application/json;charset=utf-8",
        `${base}-equation-report.json`,
      );
    notice = {
      kind: "success",
      key: kind === "docx" ? "docxSaved" : "reportSaved",
    };
  } catch (error) {
    notice = {
      kind: "error",
      key: "exportFailed",
      detail: error instanceof Error ? error.message : undefined,
    };
  } finally {
    if (current === operation) {
      exporting = false;
      render();
      app
        .querySelector<HTMLButtonElement>(
          kind === "docx" ? "#export-docx" : "#export-report",
        )
        ?.focus({ preventScroll: true });
    }
  }
}
// Avoid navigating away if a file is dropped just outside the target.
window.addEventListener("dragover", (event) => {
  if (event.dataTransfer?.types.includes("Files")) event.preventDefault();
});
window.addEventListener("drop", (event) => {
  if (event.dataTransfer?.types.includes("Files")) event.preventDefault();
});
render();
