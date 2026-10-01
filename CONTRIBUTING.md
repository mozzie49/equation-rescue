# Contributing

Install Node 22.12+ or 24, run `npm ci`, then `npm run check`.

Please keep changes narrow. New notation needs:

1. A refusal or failure fixture showing the existing behavior
2. Explicit OMML structure assertions and source-preservation tests
3. A rendered output check in a document application
4. A documented compatibility boundary, including what remains unsupported

Do not broaden the parser by silently ignoring commands. Do not regenerate a whole input document through a document library. Preserve all unrelated package records and source spans.

Run `npm run test:e2e` with Playwright Chromium installed when changing the interface. Tests must cover interrupted/repeated uploads, selections, export, invalid files, language switching and third-party request absence. Keep malicious-content tests as small synthetic fixtures.

For sample changes, recreate the original with `python scripts/generate-sample.py` (python-docx), run `npm run samples`, render both DOCX files and inspect all pages. Update the validation record honestly. A KaTeX preview is not Word output, and a structural test is not an application rendering test.

Never commit personal documents, credentials, node_modules, build output or test traces containing private data. The checked-in teaching handout is original project content under the MIT license.
