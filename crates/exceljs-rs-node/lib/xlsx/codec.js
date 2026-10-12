'use strict';

// Moves ExcelJS workbook models across the napi boundary as JSON.
// JSON has no Date or binary type, so they are tagged:
//   Date            <-> {"$date": <ms since epoch> | null (invalid date)}
//   Buffer/bytes    <-> {"$buffer": "<base64>"}
//   NaN / ±Infinity <-> {"$num": "NaN" | "Infinity" | "-Infinity"}
//   a key whose value is undefined  <-  {"$undef": true} (engine -> JS only)
// The Rust engine uses the same tags (crates/exceljs-rs/src/jsv.rs, model_json.rs).

function encodeReplacer(key, value) {
  // `value` has already been through toJSON(); look at the raw value.
  const raw = this[key];
  if (typeof raw === 'number' && !Number.isFinite(raw)) {
    return {$num: String(raw)};
  }
  if (raw instanceof Date) {
    const ms = raw.getTime();
    return {$date: Number.isNaN(ms) ? null : ms};
  }
  if (Buffer.isBuffer(raw) || raw instanceof Uint8Array) {
    return {$buffer: Buffer.from(raw.buffer, raw.byteOffset, raw.byteLength).toString('base64')};
  }
  if (raw instanceof ArrayBuffer) {
    return {$buffer: Buffer.from(raw).toString('base64')};
  }
  return value;
}

function decodeReviver(key, value) {
  if (value !== null && typeof value === 'object' && !Array.isArray(value)) {
    if ('$date' in value) return new Date(value.$date === null ? NaN : value.$date);
    if ('$buffer' in value) return Buffer.from(value.$buffer, 'base64');
    if ('$num' in value) return Number(value.$num);
  }
  return value;
}

// Shared objects: the reader returns a table of objects (each may refer to
// earlier entries) and a model that refers to them as {"$ref": n}. Resolving
// them restores ExcelJS's object sharing (e.g. one style object per style id).
function isRef(v) {
  return v !== null && typeof v === 'object' && !Array.isArray(v) && typeof v.$ref === 'number' && Object.keys(v).length === 1;
}

function resolveIn(value, table) {
  if (isRef(value)) return table[value.$ref];
  if (value !== null && typeof value === 'object' && !Buffer.isBuffer(value) && !(value instanceof Date)) {
    for (const k of Object.keys(value)) {
      value[k] = resolveIn(value[k], table);
    }
  }
  return value;
}

// Parses tagged JSON. A reviver returning undefined deletes the key, so keys
// that must exist with the value undefined are set after parsing.
function parse(json, table) {
  const undefs = [];
  const result = JSON.parse(json, function reviver(key, value) {
    if (value !== null && typeof value === 'object' && !Array.isArray(value)) {
      if (table && isRef(value)) return table[value.$ref];
      if (value.$undef === true) {
        undefs.push(this, key);
        return null;
      }
    }
    return decodeReviver(key, value);
  });
  for (let i = 0; i < undefs.length; i += 2) {
    undefs[i][undefs[i + 1]] = undefined;
  }
  return result;
}

// The result of native.readXlsx: the shared-object table, the model and,
// when given, each worksheet's rows (binary, lib/xlsx/rows.js). With `lazy`,
// a worksheet's `rows` only supports forEach (what the model setter uses).
function decodeRead({refs, model, rows}, {lazy = false} = {}) {
  const table = [];
  for (const entry of parse(refs)) {
    table.push(resolveIn(entry, table));
  }
  const result = parse(model, table);
  if (rows) {
    // eslint-disable-next-line global-require
    const reader = new (require('./rows').RowsReader)(rows, table);
    (result.worksheets || []).forEach((ws, i) => {
      if (ws && i < reader.sheets.length) ws.rows = lazy ? reader.lazyRows(i) : reader.rows(i);
    });
  }
  return result;
}

// Appends the shared objects of JSON array `json` to `table` (each may
// refer to earlier entries).
function addRefs(json, table) {
  for (const entry of parse(json)) {
    table.push(resolveIn(entry, table));
  }
}

// The model of a native XlsxReader (`await reader.model()`): worksheet rows
// come from the reader in batches as the model setter loads them.
function decodeStreamed({refs, model, strings}, reader) {
  const table = [];
  addRefs(refs, table);
  const result = parse(model, table);
  // eslint-disable-next-line global-require
  const {RowBatches, readStrings} = require('./rows');
  const batches = new RowBatches(reader, table, readStrings(strings));
  (result.worksheets || []).forEach((ws, i) => {
    if (ws) ws.rows = batches.source(i);
  });
  return result;
}

module.exports = {
  encode: model => JSON.stringify(model, encodeReplacer),
  decode: json => parse(json),
  addRefs,
  decodeRead,
  decodeStreamed,
  parseWithRefs: parse,
};
