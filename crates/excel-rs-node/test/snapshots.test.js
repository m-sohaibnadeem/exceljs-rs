'use strict';

// Snapshot test of the Node API: what it shows for every fixtures/*.xlsx
// (sheets, dimensions, counts, merges, and every cell's address, type, value
// and style) must equal fixtures/<name>.json. The snapshots began as what the
// real ExcelJS 4.4.0 shows; intended changes (COMPATIBILITY.md) are written
// back with `UPDATE_SNAPSHOTS=1 npm test`. Format: see `dump` below.

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');

const FIXTURES = path.join(__dirname, '..', '..', '..', 'fixtures');
const ExcelJS = require('..');
const UPDATE = Boolean(process.env.UPDATE_SNAPSHOTS);

const TYPE_NAMES = Object.fromEntries(Object.entries(ExcelJS.ValueType).map(([name, num]) => [num, name]));

// Keys sorted, undefined dropped: the snapshot's form.
function sortKeys(value) {
  if (Array.isArray(value)) return value.map(sortKeys);
  if (value && typeof value === 'object' && !(value instanceof Date)) {
    const out = {};
    for (const key of Object.keys(value).sort()) {
      if (value[key] !== undefined) out[key] = sortKeys(value[key]);
    }
    return out;
  }
  return value;
}

function encodeScalar(v) {
  if (typeof v === 'number' && !Number.isFinite(v)) return {nonFinite: String(v)};
  if (v instanceof Date) {
    return {date: Number.isNaN(v.getTime()) ? 'Invalid Date' : v.toISOString()};
  }
  if (v && typeof v === 'object' && Array.isArray(v.richText)) {
    return {richText: v.richText.map(r => ({text: r.text, font: r.font}))};
  }
  return v;
}

// Files that load but cannot be written back (COMPATIBILITY.md,
// limitations), with the error.
const ROUND_TRIP_FAILS = {
  // No xl/workbook.xml: loads with no worksheets, as in ExcelJS.
  'edge-no-workbook.xlsx': /reading 'date1904'/,
};

function encodeValue(type, cell) {
  const v = cell.value;
  switch (type) {
    case 'Merge':
      return {master: cell.master.address};
    case 'Formula':
      return {
        formula: v.formula,
        sharedFormula: v.sharedFormula,
        shareType: v.shareType,
        ref: v.ref,
        result: v.result === undefined ? undefined : encodeScalar(v.result),
      };
    case 'Hyperlink':
      return {text: encodeScalar(v.text), hyperlink: v.hyperlink, tooltip: v.tooltip};
    default:
      return encodeScalar(v);
  }
}

// JSON with sorted keys and without undefined (the snapshot's form).
function plain(v) {
  return v === undefined ? undefined : JSON.parse(JSON.stringify(sortKeys(v)));
}

// The snapshot of `file`:
// {
//   "file": "strings.xlsx", "date1904": false,
//   "worksheets": [{
//     "id", "name", "state", "orderNo", "dimensions": "A1:A15" | null,
//     "rowCount", "columnCount", "actualRowCount", "actualColumnCount",
//     "merges": ["A1:C1"],
//     "cells": [{"address", "row", "col", "type": "String", "value", "style"?}]
//   }]
// }
// or {"file", "error": message} if the file is rejected. Cells are those
// row.eachCell() visits. Values: numbers, strings and booleans as JSON;
// non-finite numbers {"nonFinite": "NaN"}; dates {"date": ISO string};
// rich text {"richText": [{text, font?}]}; formulas {formula?,
// sharedFormula?, shareType?, ref?, result?}; errors {"error": "#N/A"};
// hyperlinks {text, hyperlink, tooltip?}; merged cells {"master": "A1"}.
async function dump(file) {
  const wb = new ExcelJS.Workbook();
  try {
    await wb.xlsx.readFile(path.join(FIXTURES, file));
  } catch (err) {
    return plain({file, error: String(err.message)});
  }
  return plain({
    file,
    date1904: Boolean(wb.properties && wb.properties.date1904),
    worksheets: wb.worksheets.map(dumpWorksheet),
  });
}

function dumpWorksheet(ws) {
  const cells = [];
  ws.eachRow(row => {
    row.eachCell(cell => {
      const type = TYPE_NAMES[cell.type];
      const entry = {address: cell.address, row: cell.row, col: cell.col, type, value: encodeValue(type, cell)};
      if (cell.style && Object.keys(cell.style).length > 0) entry.style = cell.style;
      cells.push(entry);
    });
  });
  const dims = ws.dimensions;
  return {
    id: ws.id,
    name: ws.name,
    state: ws.state,
    orderNo: ws.orderNo,
    dimensions: dims && dims.model && dims.model.top ? dims.range : null,
    rowCount: ws.rowCount,
    columnCount: ws.columnCount,
    actualRowCount: ws.actualRowCount,
    actualColumnCount: ws.actualColumnCount,
    merges: Object.keys(ws._merges || {})
      .map(k => ws._merges[k].range)
      .sort(),
    cells,
  };
}

const xlsxFiles = fs
  .readdirSync(FIXTURES)
  .filter(f => f.endsWith('.xlsx'))
  .sort();

test('fixtures exist', () => {
  assert.ok(xlsxFiles.length >= 20);
});

for (const file of xlsxFiles) {
  test(`snapshot: ${file}`, async () => {
    const snapshotFile = path.join(FIXTURES, file.replace(/\.xlsx$/, '.json'));
    const actual = await dump(file);
    const expected = fs.existsSync(snapshotFile) ? JSON.parse(fs.readFileSync(snapshotFile, 'utf8')) : undefined;
    if (UPDATE && JSON.stringify(actual) !== JSON.stringify(expected)) {
      fs.writeFileSync(snapshotFile, `${JSON.stringify(actual, null, 2)}\n`);
      return;
    }
    assert.deepEqual(actual, expected, `${file}: differs from its snapshot (UPDATE_SNAPSHOTS=1 writes it if intended)`);
  });
}

// Writing a fixture back and reading the result must show the same cells
// (the writer's regression test; ExcelJS's spec suite in compat/ checks
// the written XML).
for (const file of xlsxFiles) {
  test(`write round trip: ${file}`, async () => {
    const original = await dump(file);
    if (original.error) return;
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(path.join(FIXTURES, file));
    let buffer;
    try {
      buffer = await wb.xlsx.writeBuffer();
    } catch (err) {
      assert.ok(ROUND_TRIP_FAILS[file], `${file}: write failed: ${err.message}`);
      assert.match(err.message, ROUND_TRIP_FAILS[file]);
      return;
    }
    const back = new ExcelJS.Workbook();
    await back.xlsx.load(buffer);
    const again = plain({file, date1904: Boolean(back.properties && back.properties.date1904), worksheets: back.worksheets.map(dumpWorksheet)});
    assert.deepEqual(again, original);
  });
}

test('readFile on a missing file rejects like ExcelJS', async () => {
  const wb = new ExcelJS.Workbook();
  await assert.rejects(wb.xlsx.readFile('/no/such/file.xlsx'), /^Error: File not found: /);
});

test('load accepts Buffer, ArrayBuffer and Uint8Array', async () => {
  const buf = fs.readFileSync(path.join(FIXTURES, 'strings.xlsx'));
  for (const data of [buf, new Uint8Array(buf), buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.length)]) {
    const wb = await new ExcelJS.Workbook().xlsx.load(data);
    assert.equal(wb.getWorksheet(1).getCell('A1').value, 'hello');
  }
});
