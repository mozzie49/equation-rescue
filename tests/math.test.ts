import { describe, it, expect } from "vitest";
import { latexToOmml, MATH_NS } from "../src/core/math";
import { parseXml, walk } from "../src/core/xml";
describe("strict LaTeX subset", () => {
  it.each([
    ["x^2", "sSup"],
    ["x_i", "sSub"],
    ["x_i^2", "sSubSup"],
    ["\\frac{a+b}{c}", "f"],
    ["\\sqrt{x^2+1}", "rad"],
    ["\\alpha+\\beta", "r"],
    ["\\sum_{i=1}^{n}i^2", "nary"],
    ["\\prod_{i=1}^{n}i", "nary"],
  ])("renders %s as native %s", (latex, tag) => {
    const root = parseXml(latexToOmml(latex));
    expect(root.uri).toBe(MATH_NS);
    expect(walk(root).some((n) => n.local === tag && n.uri === MATH_NS)).toBe(
      true,
    );
  });
  it.each([
    "\\begin{matrix}1&2\\end{matrix}",
    "\\text{hello}",
    "\\left(x\\right)",
    "\\sqrt[3]{x}",
    "\\unknown{x}",
    "x^",
    "x_",
    "\\frac{1}",
    "{x",
    "x}",
    "x^2^3",
    "\\sum",
    "x^\\sum",
    "",
    "% comment",
    "x&y",
    "\\href{https://example.com}{x}",
    "\\newcommand{\\x}{y}",
    "x^{}",
  ])("refuses unsupported/ambiguous %s", (latex) =>
    expect(() => latexToOmml(latex)).toThrow(),
  );
  it("escapes XML math operators", () => {
    expect(latexToOmml("a<b")).toContain("&lt;");
  });
  it("caps nesting and length", () => {
    expect(() => latexToOmml("{".repeat(34) + "x" + "}".repeat(34))).toThrow();
    expect(() => latexToOmml("x".repeat(2001))).toThrow();
  });
});
