import {Workbook, stream, type Cell, type Anchor} from '..';
import {Buffer} from 'node:buffer';

async function example() {
  const workbook = new Workbook();
  const sheet = workbook.addWorksheet('Sales');
  sheet.columns = [{header: 'Amount', key: 'amount'}];
  sheet.addRow({amount: 12.5});
  const cell: Cell = sheet.getCell('A2');
  cell.numFmt = '#,##0.00';
  const bytes: Buffer = await workbook.xlsx.writeBuffer();
  await workbook.xlsx.load(new Uint8Array(bytes), {maxRows: 100, maxCols: 10});
  await workbook.xlsx.load(bytes.buffer as ArrayBuffer);
  const controller = new AbortController();
  await workbook.xlsx.load(bytes, {
    limits: {maxCells: 100, maxInputBytes: bytes.length, maxValidationCells: 0},
    signal: controller.signal,
    onProgress(progress) { console.log(progress.phase, progress.inflatedBytes); },
  });
  const bulk = new stream.xlsx.WorkbookWriter({filename: 'bulk.xlsx'});
  const bulkSheet = bulk.addWorksheet('Data');
  const count: number = await bulk.writeRows(bulkSheet, [[0, false, '']], {
    batchSize: 100, signal: controller.signal, onProgress(progress) { console.log(progress.rows); },
  });
  async function* source() { yield {amount: count}; }
  await bulk.writeRows(bulkSheet, source());
  for await (const batch of stream.xlsx.readBatches(new Uint8Array(bytes), {batchSize: 2, signal: controller.signal})) {
    const name: string = batch.worksheet.name;
    console.log(name, batch.rows[0].values[1]);
  }
  // @ts-expect-error Unknown read budget names are rejected.
  await workbook.xlsx.load(bytes, {limits: {maxBytes: 1}});
  // @ts-expect-error A batch size must be numeric.
  stream.xlsx.readBatches(bytes, {batchSize: '2'});
  // @ts-expect-error Bulk export requires an iterable of rows.
  bulk.writeRows(bulkSheet, 123);
  const reader = new stream.xlsx.WorkbookReader('sales.xlsx');
  reader.on('error', (error: Error) => console.error(error));
  for await (const worksheet of reader) {
    for await (const row of worksheet) console.log(row.getCell(1).value);
  }
  const anchor = {} as Anchor;
  console.log(anchor.col);
}
void example;
