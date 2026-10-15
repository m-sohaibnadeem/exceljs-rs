'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const {Readable, Writable} = require('node:stream');
const {setTimeout: delay} = require('node:timers/promises');
const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');
const zlib = require('node:zlib');
const {spawnSync} = require('node:child_process');
const ExcelJS = require('..');
let Oracle;
try { Oracle = require('../../../compat/node_modules/exceljs'); } catch (error) {
  if (error.code !== 'MODULE_NOT_FOUND') throw error;
}
const native = require('../lib/native');
const FIXTURES = path.join(__dirname, '../../../fixtures');

// Small ZIP author for deliberately malformed/unknown parts, with a
// deliberately inaccurate size header supported by the ZIP decoder.
function crc32(bytes) {
  let crc = 0xffffffff;
  for (const b of bytes) {
    crc ^= b;
    for (let i = 0; i < 8; i++) crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0);
  }
  return (crc ^ 0xffffffff) >>> 0;
}
function zip(parts) {
  const records = [];
  const directory = [];
  let offset = 0;
  for (const [name, content, reportedSize] of parts) {
    const data = Buffer.from(content);
    const compressed = zlib.deflateRawSync(data);
    const filename = Buffer.from(name);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(8, 8);
    local.writeUInt32LE(crc32(data), 14);
    local.writeUInt32LE(compressed.length, 18);
    local.writeUInt32LE(reportedSize === undefined ? data.length : reportedSize, 22);
    local.writeUInt16LE(filename.length, 26);
    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(8, 10);
    local.copy(central, 16, 14, 26);
    central.writeUInt16LE(filename.length, 28);
    central.writeUInt32LE(offset, 42);
    records.push(local, filename, compressed);
    directory.push(central, filename);
    offset += local.length + filename.length + compressed.length;
  }
  const dir = Buffer.concat(directory);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50);
  end.writeUInt16LE(parts.length, 8);
  end.writeUInt16LE(parts.length, 10);
  end.writeUInt32LE(dir.length, 12);
  end.writeUInt32LE(offset, 16);
  return Buffer.concat([...records, dir, end]);
}
async function bytes(rows = [[0, false, '', 'hello']]) {
  const workbook = new ExcelJS.Workbook();
  workbook.addWorksheet('Data').addRows(rows);
  return workbook.xlsx.writeBuffer();
}
async function collect(input, options) {
  const batches = [];
  for await (const batch of ExcelJS.stream.xlsx.readBatches(input, options)) batches.push(batch);
  return batches;
}
async function consumeReader(input, options) {
  for await (const sheet of new ExcelJS.stream.xlsx.WorkbookReader(input, options)) {
    for await (const row of sheet) void row;
  }
}

test('all read APIs share inclusive byte/count budgets; zero and unknown limits are validated', async () => {
  const data = await bytes();
  await new ExcelJS.Workbook().xlsx.load(data, {limits: {maxInputBytes: data.length, maxCells: 4}});
  for (const name of ['maxInputBytes', 'maxZipEntries', 'maxEntryBytes', 'maxInflatedBytes', 'maxCells', 'maxSharedStrings', 'maxTokenBytes']) {
    await assert.rejects(new ExcelJS.Workbook().xlsx.load(data, {limits: {[name]: 0}}), new RegExp(name));
  }
  await assert.rejects(new ExcelJS.Workbook().xlsx.load(data, {limits: {maxCells: 3}}), /maxCells/);
  for (const limits of [{maxCells: -1}, {maxCells: 1.5}, {maxCells: Infinity}, {maxCells: Number.MAX_SAFE_INTEGER + 1}, {typo: 1}, null]) {
    await assert.rejects(new ExcelJS.Workbook().xlsx.load(data, {limits}));
  }
  await assert.rejects(collect(data, {limits: {maxCells: 3}}), /maxCells/);
  await assert.rejects(consumeReader(Readable.from([data]), {limits: {maxCells: 3}}), /maxCells/);
  await consumeReader(Readable.from([data]), {limits: {maxCells: 4}});
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(data, null);
  assert.equal(workbook.worksheets[0].getCell('A1').value, 0);
});

test('archive preflight counts actual bytes of unknown/media/ignored entries and bounds raw tokens', async () => {
  const unknown = zip([['unknown.bin', 'x'.repeat(10000)]]);
  for (const limit of ['maxEntryBytes', 'maxInflatedBytes']) {
    await assert.rejects(new ExcelJS.Workbook().xlsx.load(unknown, {limits: {[limit]: 9999}}), new RegExp(limit));
    await assert.rejects(consumeReader(Readable.from([unknown]), {limits: {[limit]: 9999}}), new RegExp(limit));
  }
  await assert.rejects(collect(zip([['unknown.bin', 'x'.repeat(10000), 1]]), {limits: {maxEntryBytes: 10000}}), /declared uncompressed size/);
  const entries = zip([['unknown.bin', '123'], ['xl/media/image1.png', '4567']]);
  await new ExcelJS.Workbook().xlsx.load(entries, {limits: {maxEntryBytes: 4, maxInflatedBytes: 7, maxZipEntries: 2}});
  await assert.rejects(collect(entries, {limits: {maxInflatedBytes: 6}}), /maxInflatedBytes/);
  for (const token of ['<a attribute="' + 'x'.repeat(9000) + '"/>', '<!--' + 'x'.repeat(9000) + '-->', '<![CDATA[' + 'x'.repeat(9000) + ']]>', '<a>' + 'x'.repeat(9000)]) {
    await assert.rejects(collect(zip([['ignored.xml', token]]), {limits: {maxTokenBytes: 8192}}), /maxTokenBytes/);
  }
  const cells = zip([['xl/worksheets/sheet1.xml', '<worksheet><sheetData><row r="1"><c r="A1"/><c r="B1"/></row></sheetData></worksheet>']]);
  await assert.rejects(new ExcelJS.Workbook().xlsx.load(cells, {ignoreNodes: ['sheetData'], limits: {maxCells: 1}}), /maxCells/);
  await new ExcelJS.Workbook().xlsx.load(zip([['ignored.xml', '<a/>']]), {limits: {maxTokenBytes: 4}});
  await assert.rejects(collect(zip([['ignored.xml', '<a/>']]), {limits: {maxTokenBytes: 3}}), /maxTokenBytes/);
});

test('validation budgets reject expansion before allocation, including JS whitespace', async () => {
  const content = '<worksheet><dataValidations><dataValidation sqref="A1:B2&#xFEFF;D1"/></dataValidations></worksheet>';
  const data = zip([['xl/worksheets/sheet1.xml', content]]);
  await new ExcelJS.Workbook().xlsx.load(data, {limits: {maxValidationCells: 5}});
  await assert.rejects(collect(data, {limits: {maxValidationCells: 4}}), /maxValidationCells/);
  await assert.rejects(consumeReader(Readable.from([data]), {limits: {maxValidationCells: 0}}), /maxValidationCells/);
  const bomb = zip([['xl/worksheets/sheet1.xml', '<worksheet><dataValidations><dataValidation sqref="A1:XFD1048576"/></dataValidations></worksheet>']]);
  await assert.rejects(collect(bomb, {limits: {maxValidationCells: 100}}), /maxValidationCells/);
});

test('merge expansion is budgeted and controlled document loading can abort inside a merge', async () => {
  const merged = range => zip([
    ['xl/worksheets/sheet1.xml', `<worksheet><sheetData><row r="1"><c r="A1" t="str"><v>master</v></c></row></sheetData><mergeCells><mergeCell ref="${range}"/></mergeCells></worksheet>`],
    ['xl/workbook.xml', '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Merged" sheetId="1" r:id="rId1"/></sheets></workbook>'],
    ['xl/_rels/workbook.xml.rels', '<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>'],
  ]);
  const data = merged('A1:B2');
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(data, {limits: {maxMergeCells: 4}});
  assert.equal(workbook.worksheets[0].getCell('B2').value, 'master');
  for (const maxMergeCells of [0, 3]) {
    await assert.rejects(workbook.xlsx.load(data, {limits: {maxMergeCells}}), /maxMergeCells/);
    await assert.rejects(collect(data, {limits: {maxMergeCells}}), /maxMergeCells/);
    await assert.rejects(consumeReader(Readable.from([data]), {limits: {maxMergeCells}}), /maxMergeCells/);
  }
  const controller = new AbortController();
  const reason = new Error('abort merge');
  await assert.rejects(workbook.xlsx.load(merged('A1:XFD1048576'), {
    signal: controller.signal,
    onProgress(progress) { if (progress.phase === 'merges') controller.abort(reason); },
  }), error => error === reason);
  assert.ok(workbook.worksheets[0].getRow(1).cellCount <= 1024);
});

test('input cancellation preserves reasons, removes only owned listeners and leaves caller stream alive', async () => {
  const controller = new AbortController();
  const source = new Readable({read() {}});
  const callerError = () => {};
  source.on('error', callerError);
  const before = Object.fromEntries(['readable', 'end', 'close', 'error'].map(event => [event, source.listenerCount(event)]));
  const reason = {cancel: 'input'};
  const pending = new ExcelJS.Workbook().xlsx.read(source, {signal: controller.signal});
  await delay(10);
  controller.abort(reason);
  await assert.rejects(pending, error => error === reason);
  assert.equal(source.destroyed, false);
  for (const [event, count] of Object.entries(before)) assert.equal(source.listenerCount(event), count, event);
  source.destroy();
  const pre = new AbortController();
  pre.abort(null);
  await assert.rejects(new ExcelJS.Workbook().xlsx.readFile('does-not-exist.xlsx', {signal: pre.signal}), error => error === null);
});

test('native cancellation interrupts preflight and validation loops rather than leaving a worker running', async () => {
  const data = zip([['unknown.bin', 'x'.repeat(2_000_000)]]);
  const reader = new native.XlsxReader(data, JSON.stringify({_controlled: true}));
  const pending = reader.model();
  reader.close();
  await assert.rejects(pending);
  for (let i = 0; i < 1000 && reader.isReading(); i++) await delay(5);
  assert.equal(reader.isReading(), false);
  const first = reader.progress();
  await delay(30);
  assert.equal(reader.progress(), first);
  const bomb = zip([['xl/worksheets/sheet1.xml', '<worksheet><dataValidations><dataValidation sqref="A1:XFD1048576"/></dataValidations></worksheet>']]);
  const controller = new AbortController();
  const reason = new Error('stop validation');
  const reading = new ExcelJS.Workbook().xlsx.load(bomb, {signal: controller.signal});
  await delay(30);
  controller.abort(reason);
  await assert.rejects(reading, error => error === reason);
});

test('aborting an owned input during file open does not emit an unhandled late error', () => {
  const script = `const Excel = require(${JSON.stringify(path.join(__dirname, '..'))});
    const a = new AbortController();
    const reader = Excel.stream.xlsx.readBatches('Z:/not-here-98329.xlsx', {signal: a.signal});
    const pending = reader.next(); a.abort();
    pending.catch(error => { if (error.name !== 'AbortError') process.exitCode = 2; });`;
  const child = spawnSync(process.execPath, ['-e', script], {encoding: 'utf8', timeout: 5000});
  assert.equal(child.status, 0, child.stderr);
});

test('progress counters are monotonic and callback exceptions reject; controlled D6 and sharing survive', async () => {
  const data = await bytes([[1], [2], [3]]);
  const progress = [];
  await new ExcelJS.Workbook().xlsx.read(Readable.from([data.subarray(0, 20), data.subarray(20)]), {onProgress: p => progress.push(p)});
  assert.equal(progress.at(-1).phase, 'done');
  assert.equal(progress.at(-1).rows, 3);
  assert.equal(progress.at(-1).cells, 3);
  for (let i = 1; i < progress.length; i++) {
    for (const key of Object.keys(progress[i]).filter(key => key !== 'phase')) assert.ok(progress[i][key] >= progress[i - 1][key], key);
  }
  const streamProgress = [];
  await consumeReader(Readable.from([data]), {onProgress: p => streamProgress.push(p)});
  assert.equal(streamProgress.at(-1).rows, 3);
  const reason = new Error('progress failed');
  await assert.rejects(new ExcelJS.Workbook().xlsx.load(data, {onProgress() { throw reason; }}), error => error === reason);
  await assert.rejects(collect(data, {onProgress(p) { if (p.phase === 'rows') throw reason; }}), error => error === reason);
  const workbook = new ExcelJS.Workbook();
  workbook.addWorksheet('Mine');
  await assert.rejects(workbook.xlsx.readFile(path.join(FIXTURES, 'edge-sst-index-out-of-range.xlsx'), {limits: {maxCells: 100}}), /richText/);
  assert.deepEqual(workbook.worksheets.map(sheet => sheet.name), ['Edge']);
  const original = new ExcelJS.Workbook();
  const sheet = original.addWorksheet('Styles');
  const font = {bold: true};
  sheet.addRow([1, 2]).eachCell(cell => { cell.font = font; });
  await workbook.xlsx.load(await original.xlsx.writeBuffer(), {limits: {maxCells: 2}});
  assert.equal(workbook.worksheets[0].getCell('A1').font, workbook.worksheets[0].getCell('B1').font);
});

test('readBatches splits worksheets, preserves sparse/falsy/rich/formula/date values and omits merge slaves', async () => {
  const workbook = new ExcelJS.Workbook();
  workbook.properties.date1904 = true;
  const a = workbook.addWorksheet('First');
  a.getCell('B3').value = 0;
  a.getCell('D3').value = false;
  a.getCell('E3').value = '';
  a.getCell('F3').value = {formula: '1-1', result: 0};
  a.getCell('A4').value = new Date('2024-01-02T00:00:00Z');
  a.getCell('B4').value = {richText: [{text: 'hey', font: {bold: true}}]};
  a.getCell('C4').value = {error: '#DIV/0!'};
  a.getCell('D4').value = {text: 'link', hyperlink: 'https://example.com', tooltip: 'tip'};
  a.mergeCells('A5:B5');
  a.getCell('A5').value = 'master';
  workbook.addWorksheet('Empty');
  workbook.addWorksheet('Last').addRows([[1], [2], [3]]);
  const data = await workbook.xlsx.writeBuffer();
  const result = await collect(new Uint8Array(data), {batchSize: 2});
  assert.deepEqual(result.map(batch => [batch.worksheet.name, batch.rows.length]), [['First', 2], ['First', 1], ['Last', 2], ['Last', 1]]);
  assert.equal(result[0].worksheet.id, a.id);
  assert.equal(result[0].rows[0].number, 3);
  const row = result[0].rows[0].values;
  assert.equal(0 in row, false);
  assert.equal(1 in row, false);
  assert.deepEqual([row[2], row[4], row[5], row[6]], [0, false, '', {formula: '1-1', result: 0}]);
  const next = result[0].rows[1].values;
  assert.deepEqual(next[1], a.getCell('A4').value);
  assert.deepEqual(next[2], a.getCell('B4').value);
  assert.deepEqual(next[3], a.getCell('C4').value);
  const loaded = new ExcelJS.Workbook();
  await loaded.xlsx.load(data);
  assert.deepEqual(next[4], loaded.worksheets[0].getCell('D4').value);
  assert.equal(result[1].rows[0].values[1], 'master');
  assert.equal(2 in result[1].rows[0].values, false);
  if (Oracle) {
    const oracle = new Oracle.Workbook();
    await oracle.xlsx.load(data);
    assert.deepEqual(next[1], oracle.worksheets[0].getCell('A4').value);
  }
  for (const batchSize of [0, -1, 0.5, NaN]) await assert.rejects(collect(data, {batchSize}), /batchSize/);
  await assert.rejects(collect({}), /input/);
  await assert.rejects(collect(Buffer.from('bad zip')));
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'exceljs-rs-batches-'));
  try {
    const filename = path.join(dir, 'input.xlsx');
    fs.writeFileSync(filename, data);
    assert.deepEqual(await collect(filename, {batchSize: 2}), result);
  } finally { fs.rmSync(dir, {recursive: true}); }
});

test('batch iteration has no public model construction and closes native reader on early return', async () => {
  const data = await bytes(Array.from({length: 5000}, (_, i) => [i]));
  const Row = require('../lib/doc/row');
  const Cell = require('../lib/doc/cell');
  const previousRow = Object.getOwnPropertyDescriptor(Row.prototype, 'model');
  const previousCell = Object.getOwnPropertyDescriptor(Cell.prototype, 'value');
  Object.defineProperty(Row.prototype, 'model', {configurable: true, get() { throw new Error('Row construction path'); }, set() { throw new Error('Row model'); }});
  Object.defineProperty(Cell.prototype, 'value', {configurable: true, get() { throw new Error('Cell construction path'); }, set() { throw new Error('Cell value'); }});
  const originalClose = native.XlsxReader.prototype.close;
  let closed = 0;
  native.XlsxReader.prototype.close = function() { closed++; return originalClose.call(this); };
  try {
    for await (const batch of ExcelJS.stream.xlsx.readBatches(data, {batchSize: 7})) { assert.equal(batch.rows.length, 7); break; }
    assert.equal(closed, 1);
  } finally {
    Object.defineProperty(Row.prototype, 'model', previousRow);
    Object.defineProperty(Cell.prototype, 'value', previousCell);
    native.XlsxReader.prototype.close = originalClose;
  }
});

test('batch worksheet order follows workbook order with reordered ZIP entries and sparse IDs', async () => {
  const data = zip([
    ['xl/worksheets/sheet2.xml', '<worksheet><sheetData><row r="1"><c r="A1" t="str"><v>second</v></c></row></sheetData></worksheet>'],
    ['xl/worksheets/sheet1.xml', '<worksheet><sheetData><row r="1"><c r="A1" t="str"><v>first</v></c><c r="B1" t="str"><v>discarded slave</v></c></row></sheetData><mergeCells><mergeCell ref="A1:B1"/></mergeCells></worksheet>'],
    ['xl/workbook.xml', '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="First" sheetId="9" r:id="rId1"/><sheet name="Second" sheetId="3" r:id="rId2"/></sheets></workbook>'],
    ['xl/_rels/workbook.xml.rels', '<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Target="worksheets/sheet2.xml"/></Relationships>'],
  ]);
  const batches = await collect(data, {batchSize: 1});
  assert.deepEqual(batches.map(batch => [batch.worksheet.id, batch.worksheet.name, batch.rows[0].values[1]]), [[9, 'First', 'first'], [3, 'Second', 'second']]);
  assert.equal(2 in batches[0].rows[0].values, false);
});

test('source and progress failures close owned filename output', async () => {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'exceljs-rs-owned-'));
  try {
    const filename = path.join(directory, 'failed.xlsx');
    const writer = new ExcelJS.stream.xlsx.WorkbookWriter({filename});
    const reason = new Error('source failed');
    const source = { *[Symbol.iterator]() { yield [1]; throw reason; } };
    await assert.rejects(writer.writeRows(writer.addWorksheet('Data'), source), error => error === reason);
    for (let i = 0; i < 100 && !writer.stream.closed; i++) await delay(2);
    assert.equal(writer.stream.closed, true);
    assert.equal(writer.stream.listenerCount('error'), 0);
    fs.unlinkSync(filename);
  } finally { fs.rmSync(directory, {recursive: true}); }
});

test('bulk export pauses source at drain checkpoints and round-trips with ExcelJS', async () => {
  let release;
  let block = false;
  const chunks = [];
  const stream = new Writable({highWaterMark: 1, write(chunk, encoding, callback) {
    chunks.push(Buffer.from(chunk));
    if (block) release = callback;
    else callback();
  }});
  const writer = new ExcelJS.stream.xlsx.WorkbookWriter({stream});
  const sheet = writer.addWorksheet('Data');
  let pulled = 0;
  const rows = {async *[Symbol.asyncIterator]() {
    for (let i = 0; i < 40; i++) { pulled++; yield [i, 'text-' + i]; }
  }};
  // Drain constructor output first, then block the first row checkpoint.
  await writer.promise;
  writer._engine.checkpoint();
  writer._flush();
  await delay(5);
  block = true;
  const pending = writer.writeRows(sheet, rows, {batchSize: 4});
  for (let i = 0; i < 100 && !release; i++) await delay(5);
  assert.ok(release, 'a compressed checkpoint reached the destination');
  const atBlock = pulled;
  await delay(20);
  assert.equal(pulled, atBlock);
  assert.ok(pulled <= 4);
  block = false;
  release();
  assert.equal(await pending, 40);
  assert.equal(sheet.committed, false);
  await writer.commit();
  const data = Buffer.concat(chunks);
  for (const Library of Oracle ? [ExcelJS, Oracle] : [ExcelJS]) {
    const workbook = new Library.Workbook();
    await workbook.xlsx.load(data);
    assert.equal(workbook.worksheets[0].rowCount, 40);
    assert.equal(workbook.worksheets[0].getCell('A40').value, 39);
  }
});

test('bulk export rejects foreign/committed/concurrent/non-head sheets without pulling input', async () => {
  const writer = new ExcelJS.stream.xlsx.WorkbookWriter({});
  const a = writer.addWorksheet('A');
  const b = writer.addWorksheet('B');
  const foreign = new ExcelJS.stream.xlsx.WorkbookWriter({});
  await assert.rejects(writer.writeRows(foreign.addWorksheet('Foreign'), []), /belong/);
  await foreign.commit();
  await assert.rejects(writer.writeRows(b, [[1]]), /earlier/);
  let release;
  const source = { [Symbol.asyncIterator]() { return {next: () => new Promise(resolve => { release = resolve; }), return: async () => ({done: true})}; } };
  const pending = writer.writeRows(a, source);
  for (let i = 0; i < 100 && !release; i++) await delay(2);
  await assert.rejects(writer.writeRows(a, []), /progress/);
  await assert.rejects(writer.commit(), /progress/);
  release({done: true});
  assert.equal(await pending, 0);
  a.commit();
  await assert.rejects(writer.writeRows(a, []), /committed/);
  assert.equal(await writer.writeRows(b, [[2], [3]]), 2);
  await writer.commit();
});

test('bulk abort during pending next returns promptly, calls return and retains falsy reason for commit', async () => {
  const writer = new ExcelJS.stream.xlsx.WorkbookWriter({});
  const sheet = writer.addWorksheet('Abort');
  let returned = false;
  let requested = false;
  const source = {[Symbol.asyncIterator]() { return {
    next() { requested = true; return new Promise(() => {}); },
    return() { returned = true; return new Promise(() => {}); },
  }; }};
  const controller = new AbortController();
  const pending = writer.writeRows(sheet, source, {signal: controller.signal});
  for (let i = 0; i < 100 && !requested; i++) await delay(2);
  controller.abort(null);
  await assert.rejects(pending, error => error === null);
  assert.equal(returned, true);
  await assert.rejects(writer.commit(), error => error === null);
});

test('destination lifecycle keeps early errors, finalization errors and premature closes; caller listeners survive', async () => {
  const reason = new Error('sink failed');
  const stream = new Writable({write(chunk, encoding, callback) { callback(reason); }});
  const caller = () => {};
  stream.on('error', caller);
  const writer = new ExcelJS.stream.xlsx.WorkbookWriter({stream});
  writer._engine.checkpoint();
  writer._flush();
  await delay(10);
  await assert.rejects(writer.commit(), error => error === reason);
  assert.equal(stream.listeners('error').includes(caller), true);
  for (const final of ['error', 'close']) {
    const output = new Writable({write(chunk, encoding, callback) { callback(); }, final(callback) {
      if (final === 'error') callback(reason);
      else this.destroy();
    }});
    const exporting = new ExcelJS.stream.xlsx.WorkbookWriter({stream: output});
    await exporting.writeRows(exporting.addWorksheet('S'), [[1]]);
    await assert.rejects(exporting.commit(), final === 'error' ? error => error === reason : /closed/);
  }
  const throwing = new Writable({write(chunk, encoding, callback) { callback(); }});
  throwing.write = () => { throw reason; };
  const bad = new ExcelJS.stream.xlsx.WorkbookWriter({stream: throwing});
  bad._engine.checkpoint();
  bad._flush();
  await assert.rejects(bad.commit(), error => error === reason);
});


test('batch import skips worksheet ZIP parts absent from workbook metadata', async () => {
  const orphan = ['xl/worksheets/sheet1.xml', '<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>'];
  assert.deepEqual(await collect(zip([orphan])), []);
  assert.deepEqual(await collect(zip([orphan, ['xl/workbook.xml', '<workbook><sheets/></workbook>']])), []);
});

test('progress-triggered bulk abort never pulls another source row or leaves a rejected promise', () => {
  const child = spawnSync(process.execPath, ['--unhandled-rejections=strict', '-e', `
    const assert = require('node:assert/strict');
    const Excel = require(${JSON.stringify(path.resolve(__dirname, '..'))});
    (async () => {
      const writer = new Excel.stream.xlsx.WorkbookWriter();
      const sheet = writer.addWorksheet('Data');
      const controller = new AbortController();
      const reason = new Error('stop');
      let pulls = 0, returned = 0;
      const source = {[Symbol.asyncIterator]() {return this;},
        next() {return Promise.resolve({value: [++pulls], done: false});},
        return() {returned++; return Promise.resolve({done: true});}};
      await assert.rejects(writer.writeRows(sheet, source, {
        batchSize: 1, signal: controller.signal,
        onProgress(p) {if (p.phase === 'rows') controller.abort(reason);},
      }), error => error === reason);
      await new Promise(resolve => setImmediate(resolve));
      assert.equal(pulls, 1);
      assert.equal(returned, 1);
      await assert.rejects(writer.commit(), error => error === reason);
    })().catch(error => {console.error(error); process.exitCode = 1;});
  `], {encoding: 'utf8', timeout: 10000});
  assert.equal(child.status, 0, child.stderr);
});

test('batch progress abort rejects before delivering a batch', async () => {
  const book = new ExcelJS.Workbook();
  book.addWorksheet('Data').addRow([1]);
  const controller = new AbortController();
  const reason = new Error('stop before yield');
  const reader = ExcelJS.stream.xlsx.readBatches(await book.xlsx.writeBuffer(), {
    signal: controller.signal,
    onProgress(p) {if (p.phase === 'rows') controller.abort(reason);},
  });
  await assert.rejects(reader.next(), error => error === reason);
});

test('bulk checkpoints service timer cancellation with synchronous input and a fast destination', async () => {
  const writer = new ExcelJS.stream.xlsx.WorkbookWriter();
  const sheet = writer.addWorksheet('Data');
  const controller = new AbortController();
  const reason = new Error('timer abort');
  let pulls = 0;
  let returned = false;
  const source = { *[Symbol.iterator]() {
    try { for (let i = 0; i < 2000; i++) {pulls++; yield [i];} }
    finally {returned = true;}
  }};
  await assert.rejects(writer.writeRows(sheet, source, {
    batchSize: 64, signal: controller.signal,
    onProgress(progress) {if (progress.rows === 64) setImmediate(() => controller.abort(reason));},
  }), error => error === reason);
  assert.ok(pulls >= 64 && pulls < 2000);
  assert.equal(returned, true);
  await assert.rejects(writer.commit(), error => error === reason);
});


test('batch early return without controls cancels its actual native worker', async () => {
  const data = zip([
    ['xl/worksheets/sheet1.xml', '<worksheet><sheetData><row r="1"><c r="A1" t="str"><v>' + 'a'.repeat(1_100_000) + '</v></c></row><row r="2"><c r="A2" t="str"><v>' + 'b'.repeat(20_000_000) + '</v></c></row></sheetData></worksheet>'],
    ['xl/workbook.xml', '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>'],
    ['xl/_rels/workbook.xml.rels', '<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>'],
  ]);
  const originalClose = native.XlsxReader.prototype.close;
  let captured;
  native.XlsxReader.prototype.close = function() {captured = this; return originalClose.call(this);};
  try {
    for await (const batch of ExcelJS.stream.xlsx.readBatches(data, {batchSize: 1})) {
      assert.equal(batch.rows[0].number, 1);
      break;
    }
    assert.ok(captured);
    // The owner and parser must share counters/control even without opt-in
    // preflight. Merely dropping the channel leaves these counters at zero.
    assert.ok(JSON.parse(captured.progress()).rows >= 1);
    for (let i = 0; i < 200 && captured.isReading(); i++) await delay(5);
    assert.equal(captured.isReading(), false);
  } finally {native.XlsxReader.prototype.close = originalClose;}
});
