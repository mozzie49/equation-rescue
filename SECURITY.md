# Security and privacy

EquationRescue processes selected DOCX bytes in browser memory. It has no upload endpoint, account, telemetry, analytics or external conversion service. Bundled libraries and fonts load from the app's own static host. The example button fetches only the bundled example.

## Boundaries

- Archive/file/member/entry-count/compression-ratio limits are checked before inflation
- Deflate input is fed in 1 KiB chunks and actual expanded output is checked before copying into a fixed-size target; forged small headers cannot allocate the entire claimed expansion
- ZIP local and central headers must agree; unsafe, absolute, traversal and ambiguous duplicate paths are refused; no archive is extracted onto the filesystem
- DTDs and entities in inspected XML are rejected; parsing uses namespace URIs, complexity bounds and a strict parser
- Documents containing or defining hidden text are refused to avoid exposing hidden source as a visible equation
- Formula source is bounded; unknown commands remain verbatim
- The output is a new copy. There is no overwrite API
- Untouched ZIP records are compared against the original before an output is returned

The tool is **not a malware scanner or document sanitizer**. It preserves embedded files and external relationships without opening/following them. Uninspected package members retain their original content. Use trusted inputs, keep backups and inspect the result in Word. Reflow, unsupported layout structures and application compatibility remain real risks.

A downloaded report includes the original filename and formula source. Do not publish it when the document is private.

## Reporting

Do not attach a private real document to a public issue. Reproduce the issue with a minimal synthetic DOCX and explain the expected behavior. For a security-sensitive issue, use the repository's private security-reporting channel if its owner has enabled one; otherwise contact the maintainer through a verified private channel. No unverified security contact address is supplied here.

No independent penetration test or comprehensive Office security audit has been completed. Tests are regression checks, not a proof of safety.
