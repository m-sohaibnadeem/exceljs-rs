'use strict';
const test = require('node:test');
const {execFileSync} = require('node:child_process');
const {pathToFileURL} = require('node:url');

test('native ESM supports the documented named and namespace imports', () => {
  const url = pathToFileURL(require.resolve('..')).href;
  execFileSync(process.execPath, ['--input-type=module', '-e', `
    import {Workbook, stream, ValueType} from ${JSON.stringify(url)};
    import * as Excel from ${JSON.stringify(url)};
    import assert from 'node:assert/strict';
    assert.equal(Excel.Workbook, Workbook);
    assert.equal(typeof stream.xlsx.WorkbookReader, 'function');
    const book = new Workbook();
    const cell = book.addWorksheet('Data').getCell('A1');
    cell.value = 42;
    assert.equal(cell.type, ValueType.Number);
  `], {stdio: 'pipe'});
});
