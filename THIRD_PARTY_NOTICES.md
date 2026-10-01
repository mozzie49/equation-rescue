# Third-party notices

The application bundles these runtime dependencies and preserves their license notices in `public/licenses/` (copied into the static build):

| Package  | Version | License                         | Purpose                                            |
| -------- | ------- | ------------------------------- | -------------------------------------------------- |
| fflate   | 0.8.3   | MIT                             | Deflate compression/decompression primitives       |
| saxes    | 6.0.0   | ISC, including upstream notices | Namespace-aware XML parser                         |
| xmlchars | 2.2.0   | MIT                             | XML character classification used by saxes         |
| KaTeX    | 0.16.22 | MIT                             | Approximate browser math preview and bundled fonts |

Dependency versions are locked in package-lock.json. Development tools retain their own licenses in their installed packages. The complete project is not relicensed under a dependency's license.

The saxes npm package does not include its standalone LICENSE file; its notice was retrieved from the official v6.0.0 repository: https://raw.githubusercontent.com/lddubeau/saxes/v6.0.0/LICENSE

Prior-art links in the README are acknowledgments, not dependencies. No conversion code from AfterMath, Converter or seewo-doc/docx-math-converter is bundled. We do not claim their algorithms or workflows as original inventions.
