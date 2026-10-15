# exceljs-rs

Read, write, and stream XLSX and CSV files in Node.js with a Rust engine and
an ExcelJS-compatible JavaScript API.

exceljs-rs targets the public API of ExcelJS 4.4.0. It includes TypeScript
declarations, workbook editing, cell styles, formulas, tables, and streaming.
Review the [known differences](COMPATIBILITY.md) and test your workbooks when
migrating. This project is not affiliated with ExcelJS.

[Documentation](https://m-sohaibnadeem.github.io/exceljs-rs/) ?
[API guide](https://m-sohaibnadeem.github.io/exceljs-rs/guide.html) ?
[Contributing](CONTRIBUTING.md)

## Install

```sh
npm install exceljs-rs
```

Use Node.js 22 or later. Release targets are Linux x64 and arm64 (glibc),
macOS x64 and arm64, and Windows x64. The package loads a native addon and
does not run in browsers. Platform binaries are optional dependencies;
do not install with `--omit=optional`.

For an unpublished checkout, follow the [source build instructions](CONTRIBUTING.md#build).

## Get started

Save this as `example.cjs` and run `node example.cjs`:

```js
const {Workbook} = require('exceljs-rs');

async function main() {
  const workbook = new Workbook();
  const sheet = workbook.addWorksheet('Sales');
  sheet.columns = [
    {header: 'Item', key: 'item', width: 24},
    {header: 'Amount', key: 'amount', width: 14},
  ];
  sheet.addRow({item: 'Notebook', amount: 12.5});
  sheet.getRow(1).font = {bold: true};
  sheet.getColumn('amount').numFmt = '#,##0.00';
  await workbook.xlsx.writeFile('sales.xlsx');

  const saved = new Workbook();
  await saved.xlsx.readFile('sales.xlsx');
  console.log(saved.getWorksheet('Sales').getCell('B2').value);
}

main().catch(error => {
  console.error(error);
  process.exitCode = 1;
});
```

ES modules and TypeScript:

```ts
import {Workbook} from 'exceljs-rs';
const workbook = new Workbook();
workbook.addWorksheet('Data').addRow(['Hello', 42]);
```

Install `@types/node` in TypeScript projects. The declarations use Node.js
buffer and stream types.

## ExcelJS compatibility

To try an existing application with exceljs-rs, change the import:

```diff
- const ExcelJS = require('exceljs');
+ const ExcelJS = require('exceljs-rs');
```

The compatibility suite exercises selected ExcelJS 4.4.0 public API specs.
It does not establish compatibility with every feature or workbook.
[COMPATIBILITY.md](COMPATIBILITY.md) describes known differences and limitations.
Formulas store expressions and cached results; exceljs-rs does not calculate them.

## Large workbooks and CSV

Inside an async function:

```js
const Excel = require('exceljs-rs');
const writer = new Excel.stream.xlsx.WorkbookWriter({filename: 'data.xlsx'});
const sheet = writer.addWorksheet('Data');
for (let i = 0; i < 100_000; i++) {
  sheet.addRow([i, 'row ' + i]).commit();
}
await writer.commit();

const workbook = new Excel.Workbook();
await workbook.csv.readFile('input.csv');
await workbook.csv.writeFile('output.csv');
```

Commit streaming rows when complete. A committed row cannot be edited.
Keep each worksheet within Excel's limit of 1,048,576 rows. The streaming
reader buffers compressed input before parsing; memory also depends on
shared strings, styles, and workbook structure.

The following opt-in extensions preserve the existing ExcelJS APIs:

```js
const controller = new AbortController();
await workbook.xlsx.readFile('input.xlsx', {
  limits: {
    maxInputBytes: 50 * 1024 * 1024,
    maxInflatedBytes: 250 * 1024 * 1024,
    maxEntryBytes: 100 * 1024 * 1024,
    maxZipEntries: 1000,
    maxCells: 1_000_000,
    maxSharedStrings: 200_000,
    maxTokenBytes: 1024 * 1024,
    maxValidationCells: 1_000_000,
    maxMergeCells: 1_000_000,
  },
  signal: controller.signal,
  onProgress: progress => console.log(progress.phase, progress.rows),
});
```

`load`, `read`, `readFile`, `WorkbookReader`, and `readBatches` share these
controls. Every limit is inclusive, accepts a nonnegative safe integer, and
treats zero as a real bound. Byte limits use bytes; the other limits count
entries, raw worksheet cells, shared-string items, or cells covered by validation
and merge declarations. Expansion counts include the master and repeated or
overlapping declarations. `maxTokenBytes` bounds each raw XML tag, comment,
CDATA section, declaration, or contiguous text run, including malformed and
ignored XML parts. It does not bound the combined text of multiple runs.

Controlled reads make an extra preflight pass: every ZIP entry is inflated
with bounded reads and counted once, including directories, duplicate entries,
media, and unknown or ignored parts. Actual inflated bytes are measured rather
than trusting ZIP size headers. Subsequent parsing passes are uncharged.
The compressed input is still buffered; metadata, shared strings, and document
objects consume additional memory. Progress counters are monotonic totals;
`rows` counts parsed rows and may lead the rows delivered to the caller.
Phases are `input`, `archive`, `rows`, `merges`, and `done`.

```js
async function* records() {
  for (let i = 0; i < 100_000; i++) yield [i, `row ${i}`];
}
const bulkWriter = new Excel.stream.xlsx.WorkbookWriter({filename: 'bulk.xlsx'});
const bulkSheet = bulkWriter.addWorksheet('Data');
const written = await bulkWriter.writeRows(bulkSheet, records(), {
  batchSize: 128, // rows between compression and destination drain checkpoints
  signal: controller.signal,
  onProgress: progress => console.log(progress.rows),
});
bulkSheet.commit();
await bulkWriter.commit();
console.log(written);

for await (const {worksheet, rows} of Excel.stream.xlsx.readBatches('bulk.xlsx', {
  batchSize: 1000,
  limits: {maxInputBytes: 50 * 1024 * 1024, maxCells: 1_000_000},
  signal: controller.signal,
})) {
  console.log(worksheet.name, rows[0].number, rows[0].values[1]);
}
```

`writeRows` accepts synchronous or asynchronous iterables of arrays or keyed
objects accepted by `addRow`. It commits each row and returns the invocation's
row count; worksheet and workbook finalization remain the caller's responsibility.
It waits for destination drain before requesting more source rows. Export
worksheets in creation order and commit each before writing the next. Foreign or
committed worksheets, concurrent bulk operations, and concurrent workbook commits
are rejected. A single very large row and accumulated shared strings/styles can
still consume substantial memory. In-memory output intentionally retains the
whole archive. Abort or source/progress/destination failure leaves the writer
unusable; a later `commit()` rejects with the original reason.

`readBatches` uses the native document reader without constructing ExcelJS
Workbook, Row, or Cell objects. It accepts a filename, readable, Buffer,
ArrayBuffer, typed array, or DataView. It yields `{worksheet: {id, name}, rows}`
in workbook sheet order, with at most `batchSize` rows from a single worksheet.
Empty worksheets yield no batches. Each row is `{number, values}` with sparse,
one-based values (index zero unused). Dates, cached formula objects, rich text,
errors, and hyperlinks follow the document reader; formulas are not evaluated.
Styles and other row/cell metadata are omitted. Merge slaves are omitted and
only a stored master value is returned. `maxRows`, `maxCols`, and `ignoreNodes`
are supported; base64 strings are not supported by this filename-oriented API.

Read failures and progress callback exceptions reject and close native cursors;
early batch-loop exits also close the cursor. Cancellation preserves
`signal.reason` and stops native parsing cooperatively. Controlled document reads
can leave a partially loaded workbook, as ordinary reads do. Caller-owned input
and output streams remain open on abort; the library removes only its own input
listeners. Owned filename streams are closed. Consume controlled WorkbookReader
worksheets inside the outer iteration. Export errors invoke the source iterator's
`return()`; an arbitrary source that ignores cancellation or blocks in `next()`
may also block `return()`, so cleanup is requested without delaying rejection.

## Performance

The [benchmark page](https://m-sohaibnadeem.github.io/exceljs-rs/benchmarks.html)
records measurements and methodology. Results depend on workload, hardware,
and build. Reproduce them with `node bench/run.js` after a release build
and `npm ci` in `bench/`.

## Security

Read [SECURITY.md](SECURITY.md) before processing uploads. Use a separate
process with memory and time limits for untrusted files. Worksheet
protection does not encrypt a file.

## Development and releases

See [CONTRIBUTING.md](CONTRIBUTING.md) for builds and tests and
[RELEASING.md](RELEASING.md) for local npm publication, GitHub Actions releases,
and GitHub Pages setup.

## License

[MIT](LICENSE). Third-party components retain their own licenses;
retain [LICENSE-EXCELJS](LICENSE-EXCELJS) and
[LICENSE-THIRD-PARTY](LICENSE-THIRD-PARTY), and
[LICENSE-RUST-DEPENDENCIES](LICENSE-RUST-DEPENDENCIES) when redistributing it.
