# Supported syntax in v0.1

Only source between explicit `\(...\)`, `\[...\]`, or `$$...$$` delimiters is considered. Single dollar signs are ignored. The latter two forms require a paragraph containing only that equation and whitespace.

## Supported

- ASCII letters, digits and basic punctuation/operators: `+ - = ( ) , . [ ] | : ; ! ? < > /`
- Groups in braces, at most 32 levels
- Subscript and superscript, including both: `x_i`, `x^2`, `x_i^{n+1}`
- Fractions with braced arguments: `\frac{a+b}{c}`
- Square roots with a braced argument: `\sqrt{x^2+1}`
- Sums and products with optional limits and a following term: `\sum_{i=1}^{n} i^2`, `\prod_{i=1}^{n} i`
- Greek commands listed in `src/core/math.ts`, including `\alpha`, `\beta`, `\pi`, `\Sigma` and variant forms
- `\times`, `\cdot`, `\pm`, `\mp`, `\le`, `\leq`, `\ge`, `\geq`, `\ne`, `\neq`, `\approx`, `\infty`, `\in`, `\notin`, `\to`, `\rightarrow`, `\partial`, `\nabla`, `\ldots`

Sum/product conversion places the immediately following atom (or braced group) inside the native n-ary operand and retains following terms in sequence. It does not infer mathematical semantics or simplify expressions.

Whitespace inside mathematical source follows TeX-style spacing rather than literal text spacing. Source run bold/italic settings are retained for surrounding text; converted math uses native math typography.

## Intentionally refused

Matrices/environments, `\left` and `\right`, indexed roots, `\text` and font commands, accents, integrals, named functions, escaped special symbols, spacing commands, alignment, macros, labels, references, packages, and arbitrary HTML or links. Unknown characters/commands are not dropped. Malformed, empty, excessively long or deeply nested source stays unchanged.

Support in the KaTeX preview does not imply export support. The independent export parser is deliberately stricter, and unsupported candidates remain unselected and unchanged even if a renderer could display them.

## Document eligibility

Only transitional WordprocessingML with the main document at `word/document.xml` is accepted. Ordinary body/table paragraphs may contain paragraph properties and simple runs containing text/run properties. Any other direct child or non-text run content makes the entire paragraph ineligible. Existing OMML, bookmarks, fields, hyperlink wrappers and drawings therefore cause a conservative skip.

Multi-paragraph formulas are not joined. Headers, footers, footnotes, endnotes, comments and text boxes are not repaired. No OCR or image conversion is performed.

Documents containing or defining hidden text are refused, so converting a hidden source run cannot unexpectedly expose it as visible math. This deliberately also refuses unused hidden-text style definitions.
