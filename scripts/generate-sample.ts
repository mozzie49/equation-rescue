// Generate the after example from the checked-in original. The Python script creates the before fixture.
import { readFileSync, writeFileSync } from "node:fs";
import { inspectDocx, convertDocx } from "../src/core/docx";
const input = readFileSync("public/examples/teaching-handout.docx");
const inspection = inspectDocx(new Uint8Array(input), "teaching-handout.docx");
const output = convertDocx(
  inspection,
  inspection.candidates.filter((c) => c.supported).map((c) => c.id),
);
writeFileSync("public/examples/teaching-handout.repaired.docx", output.bytes);
writeFileSync(
  "public/examples/teaching-handout.report.json",
  JSON.stringify(output.report, null, 2),
);
console.log(
  `${output.report.converted} converted; ${output.report.unsupported} unsupported`,
);
