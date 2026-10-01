# Validation record

Checked on 2026-10-01. This is an experimental alpha, not a certification of arbitrary DOCX compatibility.

[Live demo](https://mozzie49.github.io/equation-rescue/) · [All workflow runs](https://github.com/mozzie49/equation-rescue/actions/workflows/ci.yml)

## Actually exercised

- 62 unit/regression tests pass locally and in GitHub Actions
- TypeScript typecheck and Vite production build pass
- Local full dependency audit reported zero known vulnerabilities; CI production dependency audit passes
- All nine Chromium end-to-end tests pass: upload/review/selection, repaired DOCX/report download, byte-identical no-selection export, language/reset/repeated selection, invalid/oversized inputs, empty state, untrusted literal text/complex paragraphs, cancellation race, and 375px layout/same-host requests
- [First passing nine-test CI run](https://github.com/mozzie49/equation-rescue/actions/runs/36842032628) on `76ca45884b06d7da12567c8da5f7ed93a01bf5ae`
- [Passing checks, deployment and live-URL smoke test](https://github.com/mozzie49/equation-rescue/actions/runs/36843425689) on `3767c7805ee534ec57aec48c18a3cb61eaafa204`

The live smoke check opens the real GitHub Pages URL, loads the bundled handout, deselects one supported equation, and downloads both output files. It verifies six native OMML equations, the skipped equation's source text, preserved matrix source, exact untouched ZIP records, and report counts of 6 converted / 1 skipped / 1 unsupported. It also checks Chinese language switching, reset, a new local-file selection, and a byte-identical skip-all download. The workflow saves the actual screenshot, repaired DOCX and report as a short-lived `deployed-browser-evidence` artifact.

A separate cloud-browser inspection checked the live page, sample, selective counts, unsupported reason and Chinese/reset flow. Its download-event tool timed out, so that observation alone was not counted as download verification. The successful real-URL Chromium smoke test above captured and checked the downloaded files.

The initial Chromium run found one exact-text selector that accidentally included a decorative plus sign. The test now targets the details summary, asserts its accessible name and open state, and retains the original content/security checks. No test was skipped or assertion removed to make CI pass.

## Document fixture checks

- Original one-page teaching handout: 8 candidates, 7 conversions, 1 deliberately unsupported matrix kept verbatim
- Mixed bold/italic runs, ordinary paragraphs and table cells
- Native OMML fractions, roots, sub/superscripts, Greek symbols and sums
- Exact untouched ZIP local-record comparison; no-selection byte-identical output
- Unit regressions for namespace aliases, unsupported structures, malformed math/XML, ZIP traversal/duplicates/lying expanded sizes and limits
- Source and repaired DOCX both rendered through LibreOffice; every resulting page visually inspected without clipping, overlap or missing equation glyphs
- Before/after pictures in docs/images are those actual document renders

The repaired handout remains one page, but equations change line heights. The matrix is still raw source. Surrounding bold/italic text, headers, footers and table structure remain.

## Not established

- Microsoft Word desktop, Word Online or mobile rendering/editability
- Complete OOXML schema conformance or cross-application fidelity
- A full independent security audit or compatibility across arbitrary real-world documents
- Safari, Firefox or mobile-native browser compatibility

## Remaining human release gate

1. Open the repaired handout in supported Word versions
2. Click each equation and edit an exponent, numerator and summation limit
3. Confirm source/answer fidelity and surrounding text, layout and unsupported-source preservation
4. Exercise two or three original, non-sensitive documents from the target workflow before broad compatibility claims

For changes, rerun the complete Chromium CI suite and inspect failures and screenshots. Deployment is gated on unit/build/browser checks; the real-site smoke check runs after deployment and must also pass before calling that deployment verified.
