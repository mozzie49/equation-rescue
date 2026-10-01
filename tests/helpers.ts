import { zipSync, strToU8 } from "fflate";
export const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
export function fixture(
  body: string,
  extra: Record<string, Uint8Array> = {},
  prefix = "w",
) {
  return zipSync({
    "[Content_Types].xml": strToU8(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
    ),
    "_rels/.rels": strToU8(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    ),
    "word/document.xml": strToU8(
      `<${prefix}:document xmlns:${prefix}="${W}"><${prefix}:body>${body}<${prefix}:sectPr/></${prefix}:body></${prefix}:document>`,
    ),
    "word/styles.xml": strToU8("<styles>untouched</styles>"),
    ...extra,
  });
}
export const paragraph = (s: string) =>
  `<w:p><w:r><w:t>${s.replace(/&/g, "&amp;").replace(/</g, "&lt;")}</w:t></w:r></w:p>`;
