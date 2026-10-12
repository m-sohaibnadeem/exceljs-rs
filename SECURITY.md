# Security

## Reporting a vulnerability

Use [GitHub private vulnerability reporting](https://github.com/m-sohaibnadeem/excel-rs/security/advisories/new).
The repository owner must enable private reporting before the public launch.
If that form is unavailable, open an issue requesting a private contact without
including an exploit, credentials, or confidential workbook data.

Include the package version, Node.js version, operating system, a minimal
reproduction, and expected impact. Remove personal information and secrets
from sample files. Security fixes target the latest release.

## Processing untrusted files

- Run parsing in a separate process with enforced memory, CPU, and wall-time
  limits. A worker thread alone does not isolate process memory exhaustion.
- Limit upload size and concurrency. Compressed size does not bound decompressed
  size. Streaming is not protection against ZIP bombs.
- Large ranges, shared strings, styles, images, and deeply nested input can
  exhaust resources. The parser does not provide a complete resource budget.
- Authorize file access. Resolve user-controlled paths within your application's
  storage directory before passing them to file APIs.
- CSV quoting does not prevent spreadsheet formula injection. Sanitize untrusted
  values beginning with formula control characters before exporting data for
  spreadsheet applications.
- Treat hyperlinks, formulas, and embedded content as untrusted. excel-rs stores
  formulas and cached values; it does not calculate formulas.
- Worksheet protection restricts editing. It is not encryption or access control.

The Rust engine forbids handwritten unsafe code. Native dependencies and
generated bindings remain part of the trusted code base. Malformed-file tests
are regression checks, not proof that arbitrary input cannot crash.
