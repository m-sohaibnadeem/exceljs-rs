'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const {Readable, Writable} = require('node:stream');
const {Workbook} = require('..');

test('CSV input errors reject without removing caller listeners', async () => {
  const input = new Readable({read() { this.destroy(new Error('input failed')); }});
  const listener = () => {};
  input.on('data', listener);
  await assert.rejects(new Workbook().csv.read(input), /input failed/);
  assert.ok(input.listeners('data').includes(listener));
});

test('CSV output errors reject', async () => {
  const workbook = new Workbook();
  workbook.addWorksheet('Data').addRow(['hello']);
  const output = new Writable({write(chunk, encoding, callback) { callback(new Error('output failed')); }});
  await assert.rejects(workbook.csv.write(output), /output failed/);
});

test('CSV prototype property names remain literal strings', async () => {
  const sheet = await new Workbook().csv.read(Readable.from(['__proto__,constructor,toString\n']));
  assert.deepEqual(sheet.getRow(1).values.slice(1), ['__proto__', 'constructor', 'toString']);
});

test('XLSX output errors reject', async () => {
  const workbook = new Workbook();
  workbook.addWorksheet('Data').addRow(['hello']);
  const output = new Writable({write(chunk, encoding, callback) { callback(new Error('xlsx output failed')); }});
  await assert.rejects(workbook.xlsx.write(output), /xlsx output failed/);
});

test('XLSX write waits for destination finalization', async () => {
  const workbook = new Workbook();
  workbook.addWorksheet('Data').addRow(['hello']);
  let finalized = false;
  const output = new Writable({
    write(chunk, encoding, callback) { callback(); },
    final(callback) { setImmediate(() => { finalized = true; callback(); }); },
  });
  await workbook.xlsx.write(output);
  assert.equal(finalized, true);
});
