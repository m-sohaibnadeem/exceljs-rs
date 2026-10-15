'use strict';

const native = require('../../native');
const codec = require('../../xlsx/codec');
const {AsyncRowBatches, readStrings} = require('../../xlsx/rows');
const {ReadControls, readInput} = require('../../xlsx/read-controls');
const colCache = require('../../utils/col-cache');
const {ValueType} = require('../../doc/enums');

function values(row, merges) {
  const result = [];
  let previous = 0;
  for (const cell of row.cells || []) {
    if (cell.type === ValueType.Merge) continue;
    const col = cell.address ? colCache.decodeAddress(cell.address).col : previous + 1;
    previous = col;
    if (cell.type === ValueType.Null) continue;
    if (merges.some(range => row.number >= range.top && row.number <= range.bottom &&
      col >= range.left && col <= range.right && (row.number !== range.top || col !== range.left))) continue;
    let value;
    switch (cell.type) {
      case ValueType.Hyperlink:
        value = {text: cell.text, hyperlink: cell.hyperlink};
        if (cell.tooltip) value.tooltip = cell.tooltip;
        break;
      case ValueType.Formula:
        value = {};
        for (const key of ['formula', 'sharedFormula', 'result', 'ref', 'shareType']) {
          if (cell[key] !== undefined && cell[key] !== null) value[key] = cell[key];
        }
        break;
      default:
        value = cell.value;
    }
    result[col] = value;
  }
  return result;
}

// Document-native reconciliation, without creating Workbook, Row or Cell.
// Merge slaves are omitted; only the stored master value is returned.
async function* readBatches(input, options = {}) {
  const batchSize = options.batchSize === undefined ? 1000 : options.batchSize;
  if (options.base64 !== undefined) throw new TypeError('readBatches string input is a filename; base64 is not supported');
  if (!Number.isSafeInteger(batchSize) || batchSize <= 0) throw new RangeError('batchSize must be a positive safe integer');
  const controls = new ReadControls(options);
  try {
    const buffer = await readInput(input, controls);
    const reader = new native.XlsxReader(buffer, codec.encode({...controls.nativeOptions(), _batchOrder: true, _cancellable: true}));
    controls.attach(reader);
    const head = await controls.wait(reader.model());
    const table = [];
    codec.addRefs(head.refs, table);
    const model = codec.parseWithRefs(head.model, table);
    const batches = new AsyncRowBatches(reader, table, readStrings(head.strings), controls);
    let current = -1;
    let worksheet;
    let merges = [];
    let rows = [];
    for (;;) {
      const batch = await batches.nextAsync();
      if (!batch) break;
      if (current !== batch.sheet) {
        if (rows.length) { controls.notify('rows'); yield {worksheet, rows}; rows = []; }
        current = batch.sheet;
        const sheet = model.worksheets[current];
        worksheet = {id: sheet.id, name: sheet.name};
        merges = (sheet.mergeCells || []).map(range => colCache.decodeEx(range));
      }
      for (let i = 0; i < batch.count; i++) {
        controls.check();
        const row = batch.rows.row();
        rows.push({number: row.number, values: values(row, merges)});
        if (rows.length === batchSize) {
          controls.notify('rows');
          yield {worksheet, rows};
          rows = [];
        }
      }
      if (batch.last && rows.length) {
        controls.notify('rows');
        yield {worksheet, rows};
        rows = [];
      }
    }
    if (rows.length) { controls.notify('rows'); yield {worksheet, rows}; }
    controls.finish();
  } finally { controls.close(); }
}

module.exports = readBatches;
