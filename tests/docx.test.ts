import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { strToU8 } from "fflate";
import { inspectDocx, convertDocx } from "../src/core/docx";
import { SafeZip } from "../src/core/zip";
import { parseXml, walk } from "../src/core/xml";
import { fixture, paragraph, W } from "./helpers";
const output = (bytes: Uint8Array, ids?: string[]) => {
  const i = inspectDocx(bytes);
  return new TextDecoder().decode(
    new SafeZip(
      convertDocx(i, ids ?? i.candidates.map((c) => c.id)).bytes,
    ).read("word/document.xml"),
  );
};
describe("DOCX repair", () => {
  it("supports all three delimiters; ignores single dollar currency", () => {
    const i = inspectDocx(
      fixture(
        paragraph("Cost $5. \\(x^2\\)") +
          paragraph("\\[x_1\\]") +
          paragraph("$$a+b$$"),
      ),
    );
    expect(i.candidates).toHaveLength(3);
    expect(i.candidates.every((c) => c.supported)).toBe(true);
  });
  it("uses namespace URI rather than a hardcoded prefix", () => {
    const bytes = fixture(
      "<a:p><a:r><a:t>\\(x^2\\)</a:t></a:r></a:p>",
      {},
      "a",
    );
    expect(output(bytes)).toContain("<m:sSup>");
  });
  it("preserves unsupported and deselected source verbatim", () => {
    const bytes = fixture(
      paragraph("\\(x^2\\) then \\(\\unknown{x}\\) then \\(y_1\\)"),
    );
    const i = inspectDocx(bytes);
    expect(i.candidates.map((c) => c.supported)).toEqual([true, false, true]);
    const result = convertDocx(i, ["eq-1"]);
    expect(result.report).toMatchObject({
      converted: 1,
      skipped: 1,
      unsupported: 1,
    });
    const xml = new TextDecoder().decode(
      new SafeZip(result.bytes).read("word/document.xml"),
    );
    expect(xml).toContain("\\(\\unknown{x}\\)");
    expect(xml).toContain("\\(y_1\\)");
  });
  it("preserves unrelated run attributes, properties and exact paragraph source", () => {
    const before =
      '<w:p w:rsidR="1234"><w:r><w:t>Leave me alone</w:t></w:r></w:p>';
    const bytes = fixture(
      before +
        '<w:p><w:r w:rsidR="5678"><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">A \\(x^2\\) B \\(y_1\\) C</w:t></w:r></w:p>',
    );
    const xml = output(bytes);
    expect(xml).toContain(before);
    expect(xml.match(/w:rsidR="5678"/g)).toHaveLength(3);
    expect(xml).toContain("A ");
    expect(xml).toContain(" B ");
    expect(xml).toContain(" C");
    parseXml(xml);
  });
  it("skips paragraphs with bookmarks, hyperlinks, fields, objects and comments", () => {
    for (const extra of [
      '<w:bookmarkStart w:id="0" w:name="x"/>',
      "<w:hyperlink/>",
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>',
      "<w:r><w:drawing/></w:r>",
      "<!-- keep -->",
    ]) {
      const bytes = fixture(
        `<w:p>${extra}<w:r><w:t>\\(x^2\\)</w:t></w:r></w:p>`,
      );
      expect(inspectDocx(bytes).candidates[0].supported).toBe(false);
      expect(convertDocx(inspectDocx(bytes), ["eq-1"]).bytes).toEqual(bytes);
    }
  });
  it("rejects tracked changes and macros", () => {
    expect(() =>
      inspectDocx(fixture("<w:p><w:ins><w:r><w:t>x</w:t></w:r></w:ins></w:p>")),
    ).toThrow(/tracked/);
    expect(() =>
      inspectDocx(
        fixture(paragraph("x"), { "word/vbaProject.bin": new Uint8Array([1]) }),
      ),
    ).toThrow(/Macro/);
  });
  it("leaves mixed display equations and unmatched delimiters unchanged", () => {
    const i = inspectDocx(
      fixture(paragraph("Before $$x^2$$ after") + paragraph("Unclosed \\(x")),
    );
    expect(i.candidates).toHaveLength(1);
    expect(i.candidates[0].supported).toBe(false);
    expect(i.candidates[0].reason).toContain("own paragraph");
  });
  it("rejects XML entities and false namespaces", () => {
    const bytes = fixture(paragraph("x"), {
      "word/document.xml": strToU8(
        '<!DOCTYPE root [<!ENTITY x "boom">]><root>&x;</root>',
      ),
    });
    expect(() => inspectDocx(bytes)).toThrow(/DTD/);
    expect(() =>
      inspectDocx(
        fixture(paragraph("x"), {
          "word/document.xml": strToU8('<w:document xmlns:w="urn:fake"/>'),
        }),
      ),
    ).toThrow(/WordprocessingML/);
  });
  it("bundled teaching handout converts exactly seven and preserves its matrix", () => {
    const i = inspectDocx(
      new Uint8Array(readFileSync("public/examples/teaching-handout.docx")),
    );
    expect(i.candidates).toHaveLength(8);
    expect(i.candidates.filter((c) => c.supported)).toHaveLength(7);
    const result = convertDocx(
      i,
      i.candidates.map((c) => c.id),
    );
    const zip = new SafeZip(result.bytes),
      root = parseXml(new TextDecoder().decode(zip.read("word/document.xml")));
    expect(walk(root).filter((n) => n.local === "oMath")).toHaveLength(7);
    expect(result.report.unsupported).toBe(1);
    for (const e of i.zip.entries.filter((e) => e.name !== "word/document.xml"))
      expect(zip.entries.find((x) => x.name === e.name)?.raw).toEqual(e.raw);
  });
  it("refuses unsupported filename extensions", () =>
    expect(() => inspectDocx(fixture(paragraph("x")), "x.docm")).toThrow(
      /\.docx/,
    ));
});

describe("offset and ambiguity regressions", () => {
  it("handles non-BMP characters and XML entities before formula offsets", () => {
    const xml = output(fixture(paragraph("🧮 & 📐 \\(x^2\\) after 🚀")));
    expect(xml).toContain("🧮 &amp; 📐 ");
    expect(xml).toContain(" after 🚀");
    expect(xml).toContain("<m:sSup>");
    expect(() => parseXml(xml)).not.toThrow();
  });
  it("retains namespace declarations on split runs", () => {
    const bytes = fixture(
      `<w:p><a:r xmlns:a="${W}"><a:rPr><a:b/></a:rPr><a:t>A \\(x</a:t></a:r><b:r xmlns:b="${W}"><b:rPr><b:i/></b:rPr><b:t>^2\\) B</b:t></b:r></w:p>`,
    );
    const xml = output(bytes);
    const root = parseXml(xml);
    expect(
      walk(root)
        .filter((n) => n.local === "t" && n.uri === W)
        .map((n) => n.text)
        .join(""),
    ).toBe("A  B");
    expect(xml).toContain("<a:b/>");
    expect(xml).toContain("<b:i/>");
  });
  it("does not interpret escaped open delimiters", () => {
    const i = inspectDocx(fixture(paragraph(String.raw`\\(x^2\) and \(y_1\)`)));
    expect(i.candidates).toHaveLength(1);
    expect(i.candidates[0].latex).toBe("y_1");
  });
  it("warns about an unmatched delimiter even when another formula converts", () => {
    const i = inspectDocx(
      fixture(paragraph(String.raw`\(x^2\) and \(unclosed`)),
    );
    expect(i.candidates).toHaveLength(1);
    expect(i.warnings.some((w) => w.includes("incomplete"))).toBe(true);
    expect(output(i.zip.bytes)).toContain("\\(unclosed");
  });
  it("rejects track-enabled and protected settings", () => {
    for (const setting of [
      "<w:trackRevisions/>",
      '<w:documentProtection w:enforcement="1"/>',
    ])
      expect(() =>
        inspectDocx(
          fixture(paragraph("\\(x\\)"), {
            "word/settings.xml": strToU8(
              `<w:settings xmlns:w="${W}">${setting}</w:settings>`,
            ),
          }),
        ),
      ).toThrow(/Tracking|Protected/);
  });
});

describe("hidden content safety", () => {
  it("refuses hidden source rather than making it visible as an equation", () => {
    expect(() =>
      inspectDocx(
        fixture(
          "<w:p><w:r><w:rPr><w:vanish/></w:rPr><w:t>\\(x^2\\)</w:t></w:r></w:p>",
        ),
      ),
    ).toThrow(/hidden text/);
  });
  it("refuses styles that define hidden text, including inherited possibilities", () => {
    expect(() =>
      inspectDocx(
        fixture(paragraph("\\(x^2\\)"), {
          "word/styles.xml": strToU8(
            `<w:styles xmlns:w="${W}"><w:style w:styleId="secret"><w:rPr><w:vanish/></w:rPr></w:style></w:styles>`,
          ),
        }),
      ),
    ).toThrow(/hidden text/);
  });
});
