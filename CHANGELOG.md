# Changelog

## 0.2.1 - 2026-10-15

### Enhancements

- Add optional XLSX read limits, cancellation, and progress callbacks to document
  reads, WorkbookReader, and batch import. Limits cover input and actual inflated
  bytes, ZIP entries, XML tokens, cells, shared strings, and validation and merge
  expansion. Cancellation cooperatively stops native parsing.
- Add `WorkbookWriter.writeRows(worksheet, rows, options?)` for synchronous and
  asynchronous row sources. Compression checkpoints and destination backpressure
  let callers await bulk exports while preserving synchronous row commits.
- Add `stream.xlsx.readBatches(input, options?)` for native value batches in
  workbook order without constructing Workbook, Row, or Cell objects. Values
  remain sparse and one-based; early loop exits close the native reader.
- Add TypeScript declarations, API examples, and documentation for all new
  options and helpers. Existing ExcelJS 4.4.0 API defaults remain unchanged.

### Fixes

- Retain the original destination error across workbook commit races, handle
  early stream errors and premature closes, and close owned resources on failure.
- Service timer cancellation during bulk exports with fast output streams and
  synchronous sources. Progress-triggered cancellation stops before another row
  is pulled or a batch is delivered.
- Fail compatibility runs when a spec process crashes, including after writing
  its test report.
- Make native reader shutdown tests reliable when metadata has already completed
  on fast platforms, while still requiring worker shutdown and validation abort.

### Validation and limits

- 113 Node tests and 44 Rust tests pass. CI validates Node 22 and 24 on Linux,
  macOS, and Windows; the selected native ExcelJS suite passes 269 tests with one
  upstream pending test. Unmodified ExcelJS 4.4.0 fails its issue-1328 streaming
  shared-string/date test in the reference run; expectations are unchanged.
- Compressed read input is still buffered. Controlled reads add a ZIP preflight
  pass. Batch import omits styles, editing metadata, and merged slave values.
  Bulk exports require sequential worksheets; shared strings and individual
  large rows can still grow memory. See README for API and ownership details.

## 0.2.0 - not published

The release tag was retained after optimized macOS arm64 validation exposed a
timing-dependent test assertion. No npm packages were published for this version;
0.2.1 includes the enhancements and the corrected regression.

## 0.1.0

Initial release:

- XLSX and CSV reading and writing targeting ExcelJS 4.4.0 API compatibility.
- Rust parsing, rendering, compression, and CSV tokenization.
- Workbook, worksheet, row, cell, style, table, and streaming APIs.
- TypeScript declarations and platform-specific native addons.
- Fixture regression tests and selected ExcelJS public API compatibility specs.

See [COMPATIBILITY.md](COMPATIBILITY.md) for known differences and
[SECURITY.md](SECURITY.md) for input handling limits.
