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
  const reader = new stream.xlsx.WorkbookReader('sales.xlsx');
  reader.on('error', (error: Error) => console.error(error));
  for await (const worksheet of reader) {
    for await (const row of worksheet) console.log(row.getCell(1).value);
  }
  const anchor = {} as Anchor;
  console.log(anchor.col);
}
void example;
