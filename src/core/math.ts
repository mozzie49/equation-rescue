/** Deliberately small LaTeX parser. Never guesses what an unknown command means. */
export const MATH_NS =
  "http://schemas.openxmlformats.org/officeDocument/2006/math";
export const escapeXml = (s: string) =>
  s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
const symbols: Record<string, string> = {
  alpha: "α",
  beta: "β",
  gamma: "γ",
  delta: "δ",
  epsilon: "ϵ",
  varepsilon: "ε",
  zeta: "ζ",
  eta: "η",
  theta: "θ",
  vartheta: "ϑ",
  iota: "ι",
  kappa: "κ",
  lambda: "λ",
  mu: "μ",
  nu: "ν",
  xi: "ξ",
  pi: "π",
  rho: "ρ",
  sigma: "σ",
  tau: "τ",
  upsilon: "υ",
  phi: "ϕ",
  varphi: "φ",
  chi: "χ",
  psi: "ψ",
  omega: "ω",
  Gamma: "Γ",
  Delta: "Δ",
  Theta: "Θ",
  Lambda: "Λ",
  Xi: "Ξ",
  Pi: "Π",
  Sigma: "Σ",
  Upsilon: "Υ",
  Phi: "Φ",
  Psi: "Ψ",
  Omega: "Ω",
  times: "×",
  cdot: "⋅",
  pm: "±",
  mp: "∓",
  le: "≤",
  leq: "≤",
  ge: "≥",
  geq: "≥",
  ne: "≠",
  neq: "≠",
  approx: "≈",
  infty: "∞",
  in: "∈",
  notin: "∉",
  to: "→",
  rightarrow: "→",
  partial: "∂",
  nabla: "∇",
  ldots: "…",
};
const run = (s: string) =>
  `<m:r><m:t xml:space="preserve">${escapeXml(s)}</m:t></m:r>`;
type Atom = { xml: string; nary?: string };
export function latexToOmml(source: string, display = false): string {
  if (!source.trim()) throw new Error("Empty equation");
  if (source.length > 2000)
    throw new Error("Equation exceeds 2,000 characters");
  if (/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(source))
    throw new Error("Control characters are unsupported");
  let pos = 0,
    depth = 0,
    nodes = 0;
  const fail = (s: string): never => {
    throw new Error(s);
  };
  function ws() {
    while (/\s/.test(source[pos] ?? "") && pos < source.length) pos++;
  }
  function group(): string {
    ws();
    if (source[pos] !== "{")
      fail("Use braces around fraction and root arguments");
    pos++;
    const x = seq("}");
    if (!x) fail("Empty group");
    return x;
  }
  function script(): string {
    ws();
    if (source[pos] === "{") return group();
    const value = atom();
    if (value.nary) fail("Sum/product is unsupported as an unbraced script");
    if (!value.xml) fail("Empty script");
    return value.xml;
  }
  function atom(): Atom {
    ws();
    if (++nodes > 500) fail("Equation is too complex");
    const c = source[pos++];
    if (!c) fail("Missing argument");
    if (c === "{") {
      const xml = seq("}");
      if (!xml) fail("Empty group");
      return { xml };
    }
    if (c === "\\") {
      const rest = source.slice(pos),
        match = /^[A-Za-z]+/.exec(rest);
      if (!match) fail("Escaped symbols and spacing commands are unsupported");
      const name = match![0];
      pos += name.length;
      if (symbols[name]) return { xml: run(symbols[name]) };
      if (name === "frac")
        return {
          xml: `<m:f><m:num>${group()}</m:num><m:den>${group()}</m:den></m:f>`,
        };
      if (name === "sqrt") {
        ws();
        if (source[pos] === "[") fail("Indexed roots are not supported yet");
        return {
          xml: `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>${group()}</m:e></m:rad>`,
        };
      }
      if (name === "sum" || name === "prod")
        return { xml: "", nary: name === "sum" ? "∑" : "∏" };
      fail(`Unsupported command: \\${name}`);
    }
    if (/[}\^_$%&#~]/.test(c)) fail(`Unexpected character: ${c}`);
    // ASCII math and a conservative Unicode symbol range. No TeX control semantics.
    if (
      !/[A-Za-z0-9+\-=(),.\[\]|:;!?<>/]/.test(c) &&
      !/[\u0370-\u03ff\u2200-\u22ff]/.test(c)
    )
      fail(`Unsupported character: ${c}`);
    return { xml: run(c) };
  }
  function scripted(): string {
    const base = atom();
    let sub = "",
      sup = "",
      hasSub = false,
      hasSup = false;
    ws();
    while (source[pos] === "_" || source[pos] === "^") {
      const mode = source[pos++];
      if (mode === "_") {
        if (hasSub) fail("Duplicate subscript");
        hasSub = true;
        sub = script();
      } else {
        if (hasSup) fail("Duplicate superscript");
        hasSup = true;
        sup = script();
      }
      ws();
    }
    if (base.nary) {
      if (pos >= source.length || source[pos] === "}")
        fail("Sum/product requires a following term");
      const term = scripted();
      return `<m:nary><m:naryPr><m:chr m:val="${base.nary}"/><m:limLoc m:val="${display ? "undOvr" : "subSup"}"/><m:subHide m:val="${hasSub ? "0" : "1"}"/><m:supHide m:val="${hasSup ? "0" : "1"}"/></m:naryPr><m:sub>${sub}</m:sub><m:sup>${sup}</m:sup><m:e>${term}</m:e></m:nary>`;
    }
    if (hasSub && hasSup)
      return `<m:sSubSup><m:e>${base.xml}</m:e><m:sub>${sub}</m:sub><m:sup>${sup}</m:sup></m:sSubSup>`;
    if (hasSub)
      return `<m:sSub><m:e>${base.xml}</m:e><m:sub>${sub}</m:sub></m:sSub>`;
    if (hasSup)
      return `<m:sSup><m:e>${base.xml}</m:e><m:sup>${sup}</m:sup></m:sSup>`;
    return base.xml;
  }
  function seq(end?: string): string {
    if (++depth > 32) fail("Equation nesting exceeds 32 levels");
    let xml = "";
    for (;;) {
      ws();
      if (pos >= source.length) {
        if (end) fail("Unclosed group");
        break;
      }
      if (source[pos] === end) {
        pos++;
        break;
      }
      xml += scripted();
    }
    depth--;
    return xml;
  }
  const inner = seq();
  if (!inner) fail("Empty equation");
  const math = `<m:oMath xmlns:m="${MATH_NS}">${inner}</m:oMath>`;
  return display
    ? `<m:oMathPara xmlns:m="${MATH_NS}">${math}</m:oMathPara>`
    : math;
}
