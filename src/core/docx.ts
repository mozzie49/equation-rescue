import { SafeZip, LIMITS } from "./zip";
import { escapeXml, latexToOmml } from "./math";
import { parseXml, walk, isW, W_NS, type XmlNode } from "./xml";
export type Candidate = {
  id: string;
  source: string;
  latex: string;
  display: boolean;
  paragraph: number;
  inTable: boolean;
  context: string;
  start: number;
  end: number;
  supported: boolean;
  reason?: string;
  omml?: string;
};
type Run = { node: XmlNode; text: string; start: number; end: number };
type Paragraph = {
  node: XmlNode;
  runs: Run[];
  text: string;
  safe: boolean;
  reason: string;
  index: number;
};
export type Inspection = {
  filename: string;
  candidates: Candidate[];
  warnings: string[];
  paragraphCount: number;
  zip: SafeZip;
  xml: string;
  paragraphs: Paragraph[];
};
export type ConversionReport = {
  version: string;
  filename: string;
  converted: number;
  skipped: number;
  unsupported: number;
  changedParts: string[];
  warnings: string[];
  validation: string;
  equations: {
    id: string;
    source: string;
    paragraph: number;
    status: string;
    reason?: string;
  }[];
};
const encoder = new TextEncoder(),
  decoder = new TextDecoder("utf-8", { fatal: true });
function readXml(zip: SafeZip, name: string) {
  const e = zip.entries.find((e) => e.name === name);
  if (!e) throw new Error(`Missing DOCX part: ${name}`);
  if (e.size > LIMITS.xml) throw new Error("XML part exceeds the 8 MB limit");
  const text = decoder.decode(zip.read(name));
  if (/encoding\s*=\s*["'](?!utf-8["'])/i.test(text.slice(0, 150)))
    throw new Error("Only UTF-8 XML is supported");
  return text;
}
function matches(
  text: string,
): {
  source: string;
  latex: string;
  display: boolean;
  start: number;
  end: number;
}[] {
  const out = [];
  const re = /\\\(([\s\S]*?)\\\)|\\\[([\s\S]*?)\\\]|\$\$([\s\S]*?)\$\$/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text))) {
    const previous = text.slice(0, m.index).match(/\\+$/)?.[0].length ?? 0;
    if (previous % 2) continue;
    out.push({
      source: m[0],
      latex: m[1] ?? m[2] ?? m[3],
      display: m[1] === undefined,
      start: m.index,
      end: m.index + m[0].length,
    });
  }
  return out;
}
export function inspectDocx(
  bytes: Uint8Array,
  filename = "document.docx",
): Inspection {
  if (!/\.docx$/i.test(filename))
    throw new Error(
      "Choose a .docx file. DOC, DOCM, PDF and encrypted documents are unsupported",
    );
  const zip = new SafeZip(bytes);
  if (zip.entries.some((e) => /vbaProject|_xmlsignatures\//i.test(e.name)))
    throw new Error(
      "Macro-enabled and digitally signed documents are unsupported",
    );
  const types = parseXml(readXml(zip, "[Content_Types].xml"));
  if (
    types.local !== "Types" ||
    types.uri !== "http://schemas.openxmlformats.org/package/2006/content-types"
  )
    throw new Error("Invalid DOCX content types");
  const main = walk(types).find(
    (n) =>
      n.local === "Override" && n.attrs["{}PartName"] === "/word/document.xml",
  );
  if (
    main?.attrs["{}ContentType"] !==
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
  )
    throw new Error("Not a supported Word DOCX document");
  const rootRels = parseXml(readXml(zip, "_rels/.rels"));
  if (
    rootRels.local !== "Relationships" ||
    rootRels.uri !==
      "http://schemas.openxmlformats.org/package/2006/relationships"
  )
    throw new Error("Invalid DOCX root relationships");
  if (
    !rootRels.children.some(
      (n) =>
        n.attrs["{}Type"] ===
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" &&
        ["word/document.xml", "/word/document.xml"].includes(
          n.attrs["{}Target"],
        ) &&
        !n.attrs["{}TargetMode"],
    )
  )
    throw new Error("Unsupported main document relationship");
  if (zip.entries.some((e) => e.name === "word/settings.xml")) {
    const settings = walk(parseXml(readXml(zip, "word/settings.xml")));
    if (
      settings.some(
        (n) =>
          isW(n, "trackRevisions") &&
          !["0", "false", "off"].includes(n.attrs[`{${W_NS}}val`] ?? ""),
      )
    )
      throw new Error(
        "Tracking is enabled in this document. Turn Track Changes off in a separate copy first",
      );
    if (
      settings.some(
        (n) =>
          isW(n, "documentProtection") &&
          ["1", "true", "on"].includes(n.attrs[`{${W_NS}}enforcement`] ?? ""),
      )
    )
      throw new Error("Protected documents are unsupported");
  }
  const xml = readXml(zip, "word/document.xml"),
    root = parseXml(xml),
    all = walk(root);
  if (!isW(root, "document"))
    throw new Error("Only transitional WordprocessingML DOCX is supported");
  const hidden = (nodes: XmlNode[]) =>
    nodes.some(
      (n) =>
        n.uri === W_NS &&
        ["vanish", "webHidden", "specVanish"].includes(n.local) &&
        !["0", "false", "off"].includes(n.attrs[`{${W_NS}}val`] ?? ""),
    );
  if (
    hidden(all) ||
    (zip.entries.some((e) => e.name === "word/styles.xml") &&
      hidden(walk(parseXml(readXml(zip, "word/styles.xml")))))
  )
    throw new Error(
      "Documents containing or defining hidden text are unsupported",
    );
  if (
    all.some(
      (n) =>
        n.uri === W_NS &&
        /^(ins|del|moveFrom|moveTo|.*PrChange|moveFromRangeStart|moveToRangeStart)$/.test(
          n.local,
        ),
    )
  )
    throw new Error(
      "Documents with tracked changes are unsupported. Accept or reject changes in a separate copy first",
    );
  const paragraphs: Paragraph[] = [],
    candidates: Candidate[] = [];
  for (const node of all.filter((n) => isW(n, "p"))) {
    let ancestor = node.parent,
      safe = true;
    while (ancestor && !isW(ancestor, "document")) {
      if (
        ancestor.uri !== W_NS ||
        !["body", "tbl", "tr", "tc"].includes(ancestor.local)
      )
        safe = false;
      ancestor = ancestor.parent;
    }
    if (
      node.special ||
      node.children.some((n) => !isW(n, "r") && !isW(n, "pPr"))
    )
      safe = false;
    const runs: Run[] = [];
    let offset = 0;
    for (const r of node.children.filter((n) => isW(n, "r"))) {
      const texts = r.children.filter((n) => isW(n, "t"));
      const text = texts.map((t) => t.text).join("");
      if (
        r.special ||
        r.children.some((n) => !isW(n, "rPr") && !isW(n, "t")) ||
        texts.some((t) => t.children.length)
      )
        safe = false;
      runs.push({ node: r, text, start: offset, end: offset + text.length });
      offset += text.length;
    }
    // Read text across all descendant runs for reporting unsafe paragraphs too.
    const text = safe
      ? runs.map((r) => r.text).join("")
      : walk(node)
          .filter((n) => isW(n, "t"))
          .map((n) => n.text)
          .join("");
    const index = paragraphs.length + 1,
      p: Paragraph = {
        node,
        runs,
        text,
        safe,
        reason:
          "Paragraph contains fields, links, bookmarks, objects or other unsupported structure",
        index,
      };
    paragraphs.push(p);
    for (const found of matches(text)) {
      if (candidates.length >= 500)
        throw new Error(
          "More than 500 equations. Split this document into smaller files",
        );
      const candidate: Candidate = {
        ...found,
        id: `eq-${candidates.length + 1}`,
        paragraph: index,
        inTable: false,
        context: text.slice(
          Math.max(0, found.start - 60),
          Math.min(text.length, found.end + 60),
        ),
        supported: false,
      };
      let a = node.parent;
      while (a) {
        if (isW(a, "tc")) candidate.inTable = true;
        a = a.parent;
      }
      try {
        if (!safe) throw new Error(p.reason);
        if (found.display && text.replace(found.source, "").trim())
          throw new Error("Display equations must occupy their own paragraph");
        candidate.omml = latexToOmml(found.latex, found.display);
        candidate.supported = true;
      } catch (error) {
        candidate.reason =
          error instanceof Error ? error.message : "Unsupported equation";
      }
      candidates.push(candidate);
    }
  }
  const warnings = [
    "Existing document links and embedded content are preserved. This tool is not a malware scanner; use trusted documents.",
    "Equations may change line heights and pagination. Surrounding formatting is preserved, but identical layout is not guaranteed.",
    "Only ordinary body paragraphs and table cells are scanned. Headers, footers, notes, text boxes and comments are left untouched.",
    "Browser preview is approximate. Open the output in Word and check layout and equation editability before relying on it.",
  ];
  if (zip.entries.some((e) => e.name.startsWith("word/embeddings/")))
    warnings.push(
      "This document contains embedded files. They are preserved, not inspected or opened.",
    );
  if (
    paragraphs.some((p) =>
      /\\[([]|\$\$/.test(
        matches(p.text).reduce((text, m) => text.replace(m.source, ""), p.text),
      ),
    )
  )
    warnings.push(
      "Some delimiters are incomplete or split across paragraphs. Those source spans are left untouched.",
    );
  return {
    filename,
    candidates,
    warnings,
    paragraphCount: paragraphs.length,
    zip,
    xml,
    paragraphs,
  };
}
function sliceRun(run: Run, from: number, to: number, xml: string): string {
  if (from === 0 && to === run.text.length)
    return xml.slice(run.node.start, run.node.end);
  if (from === to) return "";
  const n = run.node,
    prefix = n.name.includes(":") ? n.name.split(":")[0] + ":" : "",
    properties = n.children
      .filter((c) => isW(c, "rPr"))
      .map((c) => xml.slice(c.start, c.end))
      .join("");
  return `${xml.slice(n.start, n.openEnd)}${properties}<${prefix}t xml:space="preserve">${escapeXml(run.text.slice(from, to))}</${prefix}t></${n.name}>`;
}
function patchParagraph(
  p: Paragraph,
  selected: Candidate[],
  xml: string,
): string {
  // Patch only runs touched by a selected equation; unrelated source spans stay exact.
  const groups: { first: number; last: number; equations: Candidate[] }[] = [];
  for (const c of selected) {
    const first = p.runs.findIndex(
        (r) => c.start >= r.start && c.start < r.end,
      ),
      last = p.runs.findIndex((r) => c.end > r.start && c.end <= r.end);
    if (first < 0 || last < first)
      throw new Error("Equation source mapping failed");
    const previous = groups.at(-1);
    if (previous && first <= previous.last) {
      previous.last = Math.max(last, previous.last);
      previous.equations.push(c);
    } else groups.push({ first, last, equations: [c] });
  }
  let patched = xml.slice(p.node.start, p.node.end);
  for (const group of groups.reverse()) {
    const runs = p.runs.slice(group.first, group.last + 1),
      begin = runs[0].start,
      end = runs.at(-1)!.end;
    const content = (start: number, stop: number) =>
      runs
        .map((r) => {
          const a = Math.max(start, r.start),
            b = Math.min(stop, r.end);
          return a < b ? sliceRun(r, a - r.start, b - r.start, xml) : "";
        })
        .join("");
    let result = "",
      cursor = begin;
    for (const c of group.equations) {
      result += content(cursor, c.start) + c.omml;
      cursor = c.end;
    }
    result += content(cursor, end);
    const sourceStart = runs[0].node.start - p.node.start,
      sourceEnd = runs.at(-1)!.node.end - p.node.start;
    patched = patched.slice(0, sourceStart) + result + patched.slice(sourceEnd);
  }
  return patched;
}
export function convertDocx(
  inspection: Inspection,
  selectedIds: Iterable<string>,
): { bytes: Uint8Array; report: ConversionReport } {
  const ids = new Set(selectedIds),
    chosen = inspection.candidates.filter((c) => c.supported && ids.has(c.id));
  let xml = inspection.xml;
  for (const p of [...inspection.paragraphs].reverse()) {
    const cs = chosen.filter((c) => c.paragraph === p.index);
    if (cs.length)
      xml =
        xml.slice(0, p.node.start) +
        patchParagraph(p, cs, inspection.xml) +
        xml.slice(p.node.end);
  }
  // Reparse generated XML before an output can be offered.
  parseXml(xml);
  const bytes = chosen.length
    ? inspection.zip.replace("word/document.xml", encoder.encode(xml))
    : inspection.zip.bytes.slice();
  const verification = new SafeZip(bytes);
  if (chosen.length)
    parseXml(decoder.decode(verification.read("word/document.xml")));
  for (const entry of inspection.zip.entries) {
    if (entry.name === "word/document.xml") continue;
    const actual = verification.entries.find((e) => e.name === entry.name);
    if (
      !actual ||
      actual.raw.length !== entry.raw.length ||
      actual.raw.some((b, i) => b !== entry.raw[i])
    )
      throw new Error("Preservation check failed; output was not created");
  }
  return {
    bytes,
    report: {
      version: "0.1.0",
      filename: inspection.filename,
      converted: chosen.length,
      skipped: inspection.candidates.filter(
        (c) => c.supported && !ids.has(c.id),
      ).length,
      unsupported: inspection.candidates.filter((c) => !c.supported).length,
      changedParts: chosen.length ? ["word/document.xml"] : [],
      warnings: inspection.warnings,
      validation:
        "Well-formed XML, supported native OMML structures, and unchanged raw ZIP records outside word/document.xml. Microsoft Word application rendering and editability are not automatically verified.",
      equations: inspection.candidates.map((c) => ({
        id: c.id,
        source: c.source,
        paragraph: c.paragraph,
        status: !c.supported
          ? "unsupported"
          : ids.has(c.id)
            ? "converted"
            : "skipped",
        ...(c.reason ? { reason: c.reason } : {}),
      })),
    },
  };
}
