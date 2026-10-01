import { describe, it, expect } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { inspectDocx, convertDocx } from "../src/core/docx";
import { SafeZip } from "../src/core/zip";
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
function fixture() {
  return zipSync({
    "[Content_Types].xml": strToU8(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
    ),
    "_rels/.rels": strToU8(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    ),
    "word/document.xml": strToU8(
      `<w:document xmlns:w="${W}"><w:body><w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Keep bold \\(x</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>^2\\) and italic</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>\\[\\frac{a}{b}\\]</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
    ),
    "word/styles.xml": strToU8("<styles>untouched</styles>"),
    "word/media/image.png": new Uint8Array([1, 2, 3, 4]),
  });
}
describe("feasibility gate", () => {
  it("patches mixed runs and table cells while preserving untouched package entries", () => {
    const input = fixture(),
      inspection = inspectDocx(input);
    expect(inspection.candidates).toHaveLength(2);
    expect(inspection.candidates.every((c) => c.supported)).toBe(true);
    const result = convertDocx(
      inspection,
      inspection.candidates.map((c) => c.id),
    );
    expect(result.report.converted).toBe(2);
    const zip = new SafeZip(result.bytes),
      xml = new TextDecoder().decode(zip.read("word/document.xml"));
    expect(xml).toContain("<m:sSup>");
    expect(xml).toContain("<m:f>");
    expect(xml).toContain("<w:b/>");
    expect(xml).toContain("<w:i/>");
    expect(xml).toContain("Keep bold ");
    expect(xml).toContain(" and italic");
    for (const e of inspection.zip.entries.filter(
      (e) => e.name !== "word/document.xml",
    ))
      expect(zip.entries.find((x) => x.name === e.name)?.raw).toEqual(e.raw);
  });
  it("returns a byte identical copy when nothing is selected", () => {
    const bytes = fixture();
    expect(convertDocx(inspectDocx(bytes), []).bytes).toEqual(bytes);
  });
});
