# Validation record

Checked on 2026-10-01. This record describes the local pre-publication build; it is not a certification of arbitrary DOCX compatibility.

## Actually exercised

- 62 unit/regression tests pass
- TypeScript typecheck and Vite production build pass
- `npm audit` reports zero known vulnerabilities in the installed dependency tree
- Playwright lists nine end-to-end tests; discovery only, not execution

- Original one-page teaching handout: 8 candidates, 7 conversions, 1 deliberately unsupported matrix kept verbatim
- Mixed bold/italic runs, ordinary paragraphs and table cells
- Native OMML fractions, roots, sub/superscripts, Greek symbols and sums
- Exact untouched ZIP local-record comparison; no-selection byte-identical output
- Unit regressions for namespace aliases, unsupported structures, malformed math/XML, ZIP traversal/duplicates/lying expanded sizes and limits
- Source and repaired DOCX both rendered through the installed LibreOffice renderer; every resulting page visually inspected without clipping, overlap or missing equation glyphs
- Before/after pictures in docs/images are those actual document renders

The repaired handout remains one page, but equations change line heights. The matrix is still raw source. Surrounding bold/italic text, headers, footers and table structure remain.

## Not established

- Microsoft Word desktop, Word Online or mobile rendering/editability
- Complete OOXML schema conformance or cross-application fidelity
- Browser end-to-end test success in this local environment (local browser/socket execution is restricted)
- A GitHub Actions run, live deployment or public release
- A full security audit or compatibility across arbitrary real-world documents

The CI workflow is prepared to run a production build and Playwright Chromium checks in GitHub Actions after publication is authorized. It must actually run and pass before reporting browser tests as verified.

## Human release gate

1. Open the repaired handout in supported Word versions
2. Click each equation and edit an exponent, numerator and summation limit
3. Confirm source/answer fidelity and surrounding text, layout and unsupported-source preservation
4. Run the Chromium CI suite and inspect any failures/screenshots
5. Exercise two or three original, non-sensitive documents from the target workflow before broad claims
