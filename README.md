# excel-rs

Read, write, and stream XLSX and CSV files in Node.js with a Rust engine and
an ExcelJS-compatible JavaScript API.

excel-rs targets the public API of ExcelJS 4.4.0. It includes TypeScript
declarations, workbook editing, cell styles, formulas, tables, and streaming.
Review the [known differences](COMPATIBILITY.md) and test your workbooks when
migrating. This project is not affiliated with ExcelJS.

[Documentation](https://m-sohaibnadeem.github.io/excel-rs/) ?
[API guide](https://m-sohaibnadeem.github.io/excel-rs/guide.html) ?
[Contributing](CONTRIBUTING.md)

## Install

```sh
npm install excel-rs
```

Use Node.js 22 or later. Release targets are Linux x64 and arm64 (glibc),
macOS x64 and arm64, and Windows x64. The package loads a native addon and
does not run in browsers. Platform binaries are optional dependencies;
do not install with `--omit=optional`.

For an unpublished checkout, follow the [source build instructions](CONTRIBUTING.md#build).

## Get started

Save this as `example.cjs` and run `node example.cjs`:

```js
const {Workbook} = require('excel-rs');

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
import {Workbook} from 'excel-rs';
const workbook = new Workbook();
workbook.addWorksheet('Data').addRow(['Hello', 42]);
```

Install `@types/node` in TypeScript projects. The declarations use Node.js
buffer and stream types.

## ExcelJS compatibility

To try an existing application with excel-rs, change the import:

```diff
- const ExcelJS = require('exceljs');
+ const ExcelJS = require('excel-rs');
```

The compatibility suite exercises selected ExcelJS 4.4.0 public API specs.
It does not establish compatibility with every feature or workbook.
[COMPATIBILITY.md](COMPATIBILITY.md) describes known differences and limitations.
Formulas store expressions and cached results; excel-rs does not calculate them.

## Large workbooks and CSV

Inside an async function:

```js
const Excel = require('excel-rs');
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

## Performance

The [benchmark page](https://m-sohaibnadeem.github.io/excel-rs/benchmarks.html)
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
