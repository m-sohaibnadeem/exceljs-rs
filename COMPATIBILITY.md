# Compatibility with ExcelJS

exceljs-rs targets ExcelJS 4.4.0 public API compatibility. The compatibility
suite runs selected upstream public API specs, with exclusions recorded in
`compat/suite.js`. Passing those specs does not guarantee that every workbook
or API usage behaves identically.

This page records known differences. Report unexpected behavior with a minimal,
non-confidential reproduction through [GitHub issues](https://github.com/m-sohaibnadeem/exceljs-rs/issues).

## Bugs fixed from ExcelJS

These ExcelJS 4.4.0 behaviours are bugs, and exceljs-rs fixes them instead of
copying them. Each one is shown by a test file (named in parentheses) and
tested in `crates/exceljs-rs-node/test/fixed-bugs.test.js`. The selected upstream specs run in CI.

- **B1. `_xHHHH_` escapes round-trip.** OOXML writes characters as `_xHHHH_`
  in text. ExcelJS decodes them only in shared strings, only with upper-case
  hex, and never escapes literal text of that shape when writing, so
  `"_x0041_ literal"` is written as is and read back as `"A literal"`
  (`strings.xlsx` A10). exceljs-rs decodes upper- and lower-case hex in shared
  strings, inline strings and `t="str"` values (both readers), and writes the
  `_` of escape-shaped literal text as `_x005F_` (shared strings, rich text,
  inline and `t="str"` cells, formula string results). Control characters
  are still dropped on write, as ExcelJS's issue-234 spec requires.
  (`edge-cells.xlsx` B1, C4; `strings-inline.xlsx` A10.)
- **B2. Text is XML-decoded once.** ExcelJS runs `utils.xmlDecode` on
  `t="str"` text and on sheet names after the XML parser has already decoded
  them, so the literal text `&amp; &lt;` became `& <` (both readers).
  exceljs-rs decodes once (`strings-inline.xlsx` A9, `edge-cells.xlsx` D4,
  `edge-sheet-name-entity.xlsx` sheet `a &lt; b &amp; c`).
- **B3. Empty strings survive and `useSharedStrings: false` works.** A string
  cell with an empty `<v></v>` (what the streaming writer writes for `''` with
  `useSharedStrings: false`) or `<t/>` reads as `''`; ExcelJS dropped the cell
  (`strings-inline.xlsx` A2). `workbook.xlsx.write*({useSharedStrings:
  false})` writes strings in the cells (`t="str"`) and no shared-strings part;
  ExcelJS ignored the option.
- **B4. Unknown elements inside `<c>` are skipped.** ExcelJS ended the cell
  at the close tag of any element it does not know, such as a phonetic
  `<rPh>` run, which lost the rest of the row and every later row
  (`edge-inline-phonetic.xlsx`: ExcelJS reads B1 as `"basePHONETIC"` with no
  type, and loses C1 and row 2). exceljs-rs skips such elements whole. The
  streaming reader also joins the text of every run of a rich inline string
  (ExcelJS kept the last `<t>` only).
- **B5. `<si/>` is the empty string.** ExcelJS read it as the object `{}`,
  typed String, and wrote it back as `[object Object]`
  (`edge-sst-empty-si.xlsx`).
- **B6. A `t="s"` cell in a package without `xl/sharedStrings.xml` is
  empty** (type Null, style kept). ExcelJS made it a String cell whose value
  is the index number (`edge-no-shared-strings.xlsx`, ExcelJS's
  `missing-bits.xlsx`, which its spec only requires to load).
- **B7. A formula result of `0`, `false` or `''` is kept.**
  `FormulaValue._copyModel` copied only truthy fields, so `cell.value` lost
  such results (`formulas.xlsx` C3, whose XML has `<v>0</v>`).
- **B8. CDATA is text.** ExcelJS registers no CDATA handler, so
  `<t><![CDATA[cdata text]]></t>` read as `""` (`edge-cells.xlsx` E1).
- **B9. A cell's address matches its row.** `<row r="8"><c r="A99">` puts
  the cell in row 8; ExcelJS kept the address `"A99"`, so
  `ws.getCell('A8')` and the cell disagreed. exceljs-rs addresses it `A8`
  (`edge-cells.xlsx`).
- **B10. `t="b"` booleans.** ExcelJS's `parseInt(text) !== 0` made any
  non-numeric text true, `false` included. exceljs-rs reads a non-zero integer
  or the word `true` (any case) as true and anything else as false (both
  readers; `edge-cells.xlsx` C2, E2, F2).
- **B11. The streaming writer's shared strings tell rich texts apart.**
  `utils/shared-strings.js` used the value as an object key, so every rich
  text was `"[object Object]"` and, with `useSharedStrings: true`, every
  rich-text cell showed the first rich text written. exceljs-rs keys rich text
  by its XML, as the xlsx writer does.
- **B12. The streaming reader's shared strings** (`_parseSharedStrings`):
  `<color rgb>` now gives `argb` (ExcelJS read the absent `argb`
  attribute), `<rFont val>` gives `name` (ExcelJS: `undefined`), `theme` is a
  number, only `<t>` text counts (ExcelJS added the whitespace between
  elements and let a phonetic `<rPh><t>` replace the text), escapes are
  decoded (B1), and `<si/>` is `''` (`edge-sst-rich-fonts.xlsx`,
  `edge-cells.xlsx`).
- **B13. The streaming reader with `styles: 'ignore'` (the default) reads a
  row with `s="0"`.** ExcelJS's style manager threw for style 0
  (`edge-row-style-zero.xlsx`).
- **B14. The streaming reader reads a `<c>` without `r`** as the next
  column, like the xlsx reader. ExcelJS threw (`edge-cells.xlsx` row 5).

## ExcelJS limitations exceljs-rs keeps

These aren't bugs, but they limit which files load. exceljs-rs behaves the same
way.

- Part names are fixed (`xl/workbook.xml`, `xl/_rels/workbook.xml.rels`,
  `xl/sharedStrings.xml`) and not read from `_rels/.rels`. A package without
  `xl/workbook.xml` loads with no worksheets, and writing it back fails
  (`edge-no-workbook.xlsx`).
- A worksheet part not named `xl/worksheets/sheetN.xml` is skipped
  (`edge-sheet-path-not-sheetN.xlsx`), as are chartsheets.
- Element names are matched as written: prefixed SpreadsheetML (`<x:c>`) is
  not recognised.
- Rejected files: lower-case cell references (`edge-lowercase-ref.xlsx`), a
  first `<c>` without `r` (`edge-first-cell-no-r.xlsx`), a column past `XFD`
  (`edge-column-out-of-range.xlsx`), a shared-string index that is not a valid
  index (`edge-sst-index-*.xlsx`), and anything that is not a ZIP
  (`edge-not-a-zip.xlsx`).
- Numbers use JavaScript `parseFloat` (`"12abc"` is 12, `"abc"` is `NaN`);
  shared-string indices and booleans use `parseInt`.
- A data validation on a whole-sheet range (`A1:XFD1048576`) is expanded to
  every address (ExcelJS's `DataValidations`), which takes minutes and
  gigabytes in both (ExcelJS's `issue-1842` spec data; its spec reads the file
  with `ignoreNodes: ['dataValidations']`).
- No complete resource budget: compressed size checks and streaming do not
  prevent memory exhaustion. Isolate untrusted input with process memory and
  time limits. See [SECURITY.md](SECURITY.md).

## Other differences from ExcelJS

- **D1. Lone UTF-16 surrogates** (e.g. a lone `_xD83D_` escape, or a lone
  surrogate in a cell value being written) become U+FFFD, because the engine
  works in UTF-8. JavaScript keeps the lone surrogate.
- **D2. `WorkbookReader` holds the compressed file.**
  `stream.xlsx.WorkbookReader` reads its input stream to the end, then
  parses: each worksheet is inflated as it is parsed and emitted in batches,
  so memory does not grow with the sheet (500,000 rows: 134 MB peak; ExcelJS
  144 MB), but the compressed file is in memory. ExcelJS parses as the data
  arrives. (`workbook.csv.read` parses each chunk as it arrives, as
  fast-csv does.)
- **D3. `workbook.xlsx.write(stream)` writes the file in one chunk**, once
  it is complete (the archive is built in memory, compressed: about a
  tenth of the sheet XML). `workbook.csv.write` passes its output on in
  pieces of about 1 MB, and `stream.xlsx.WorkbookWriter` as rows are
  committed.
- **D4. CSV parser option types.** fast-csv compares option values with
  tokens with `===`; exceljs-rs passes `delimiter` as a string
  (`String(delimiter)`) and any non-string `quote`, `escape`, `comment` as
  "never matches". A one-character non-string delimiter behaves differently.
- **D5. ZIP output.** Entries are deflated by zlib-rs rather than JSZip or
  archiver, so the compressed bytes differ (the uncompressed parts are
  identical). The streaming writer's archive uses data
  descriptors.
- **D6. A load that fails at a cell leaves what it loaded.** The rows are
  read while the workbook's objects are built, so an error found in a cell
  (e.g. a shared-string index that does not exist) rejects
  `workbook.xlsx.load`/`readFile` after the workbook's model was replaced:
  the workbook has the new worksheets, partly loaded. ExcelJS finds such
  errors before touching the workbook (its own model setter errors, like a
  `<row>` without `r`, leave it partly replaced too). Shown by
  `test/fixed-bugs.test.js` ("D6").

## Additional API corrections

- Opt-in XLSX `limits`, `signal`, and `onProgress` controls share native
  preflight budgets across document reads and WorkbookReader. Defaults and
  synchronous model/row commit behavior are preserved. Limits include actual
  compressed/inflated bytes, ZIP entries, raw cells/shared strings, raw XML
  tokens, and validation/merge expansion counts. See README for units and
  ownership semantics. Controlled reads still buffer compressed input and
  retain document metadata; they add a preflight inflation pass.
- `stream.xlsx.readBatches(input, options?)` returns native reconciled value
  batches in workbook order without constructing workbook/cell objects.
  Empty worksheets yield nothing; values are sparse and one-based. Styles,
  merge slave values, and other editing metadata are omitted. It uses the
  document reader's existing value behavior and does not evaluate formulas.
- `WorkbookWriter.writeRows(worksheet, rows, options?)` is an additive async
  helper that awaits compression and destination drain before pulling more
  rows. It requires sequential worksheet export; existing `row.commit(): void`
  remains synchronous. Shared strings/styles and single very large rows can
  still grow memory. Failed/aborted helpers leave the writer unusable and retain
  the original failure for `commit()`.

- CSV input and output stream errors reject their promises. CSV literals such
  as `__proto__`, `constructor`, and `toString` remain strings instead of
  being looked up through an object prototype.
- TypeScript uses Node.js Buffer types and exposes the accepted XLSX input
  types and read limits. `WorkbookReader` is an EventEmitter, not an editable
  Workbook. `Anchor` is a type, not a top-level constructor. The unsupported
  `config.setValue` declaration has been removed.

- XLSX stream writes reject destination errors and wait for destination
  finalization before resolving.

- Native Node.js ES modules can use named imports such as
  `import {Workbook} from "exceljs-rs"`. CommonJS imports are also supported.
