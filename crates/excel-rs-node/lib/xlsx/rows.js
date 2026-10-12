'use strict';

// Encodes worksheet row models for the engine's writers
// (crates/excel-rs/src/rows.rs). Rows are the bulk of a workbook, so they
// skip JSON: each row and cell model is written into one Buffer with only
// the keys RowXform/CellXform read, and each `style` object becomes an index
// into a table of distinct styles (sent as JSON). The binary layout is
// described in rows.rs.

const codec = require('./codec');
const Range = require('../doc/range');
const colCache = require('../utils/col-cache');
const Row = require('../doc/row');
const Cell = require('../doc/cell');

const CELL_KEYS = [
  'address',
  'type',
  'value',
  'comment',
  'text',
  'hyperlink',
  'tooltip',
  'master',
  'shareType',
  'formula',
  'sharedFormula',
  'result',
  'ref',
  'si',
  'range',
  'date1904',
  'ssId',
  'styleId',
];
const ROW_KEYS = ['number', 'min', 'max', 'height', 'hidden', 'outlineLevel', 'collapsed', 'styleId'];

const END = 0xff;
const CELLS = 31;
const STYLE = 30;
const F64 = 1;
const STR = 2;
const TRUE = 3;
const FALSE = 4;
const NULL = 5;
const DATE = 6;
const JSON_VALUE = 7;
const INDEX = 8;
const REF = 9;
const SST = 10;
const ADDR = 11;
const STYLE_VALUE = 29;
const BEST_FIT = 8;

// `x` as V8 keeps a parsed integer: a small integer, not a boxed double
// (which takes 16 more bytes per value held). -0 stays -0.
function small(x) {
  const i = x | 0;
  return i === x && (i !== 0 || 1 / x > 0) ? i : x;
}

class Writer {
  constructor(size = 1 << 16) {
    this.buf = Buffer.allocUnsafe(size);
    this.pos = 0;
  }

  ensure(n) {
    if (this.pos + n > this.buf.length) {
      const next = Buffer.allocUnsafe(Math.max(this.buf.length * 2, this.pos + n));
      this.buf.copy(next, 0, 0, this.pos);
      this.buf = next;
    }
  }

  u8(v) {
    this.ensure(1);
    this.buf[this.pos++] = v;
  }

  u32(v) {
    this.ensure(4);
    this.buf.writeUInt32LE(v, this.pos);
    this.pos += 4;
  }

  f64(v) {
    this.ensure(8);
    this.buf.writeDoubleLE(v, this.pos);
    this.pos += 8;
  }

  str(s) {
    this.ensure(s.length * 3 + 4);
    const n = this.buf.write(s, this.pos + 4, 'utf8');
    this.buf.writeUInt32LE(n, this.pos);
    this.pos += 4 + n;
  }

  // A field `key` with value `v`; nothing when JSON.stringify would drop it.
  field(key, v) {
    switch (typeof v) {
      case 'number':
        this.u8(key);
        this.u8(F64);
        this.f64(v);
        return;
      case 'string':
        this.u8(key);
        this.u8(STR);
        this.str(v);
        return;
      case 'boolean':
        this.u8(key);
        this.u8(v ? TRUE : FALSE);
        return;
      case 'undefined':
        return;
      case 'object':
        if (v === null) {
          this.u8(key);
          this.u8(NULL);
          return;
        }
        if (v instanceof Date) {
          this.u8(key);
          this.u8(DATE);
          this.f64(v.getTime());
          return;
        }
        break;
      default:
        break;
    }
    const json = codec.encode(v);
    if (json !== undefined) {
      this.u8(key);
      this.u8(JSON_VALUE);
      this.str(json);
    }
  }

  result() {
    return this.buf.subarray(0, this.pos);
  }
}

// Distinct style objects. A style is keyed by its own keys and values, with
// object values by identity: cells get fresh style objects whose parts
// (font, fill, ...) are shared, so this finds the repeats without
// serialising every style. Equal styles with different part objects get
// separate entries; the engine maps both to the same style id.
class StyleTable {
  constructor() {
    this.partIds = new WeakMap();
    this.nextPartId = 0;
    this.byKey = new Map();
    this.list = [];
    this.sent = 0;
  }

  partId(o) {
    let id = this.partIds.get(o);
    if (id === undefined) {
      id = this.nextPartId++;
      this.partIds.set(o, id);
    }
    return id;
  }

  key(style) {
    if (style === null || typeof style !== 'object') return `${typeof style}:${String(style)}`;
    let k = '';
    for (const name of Object.keys(style)) {
      const v = style[name];
      if (v !== null && typeof v === 'object') {
        k += `${JSON.stringify(name)}#${this.partId(v)};`;
      } else if (typeof v === 'string') {
        k += `${JSON.stringify(name)}=${JSON.stringify(v)};`;
      } else {
        k += `${JSON.stringify(name)}:${typeof v}:${String(v)};`;
      }
    }
    return k;
  }

  index(style) {
    const key = this.key(style);
    let i = this.byKey.get(key);
    if (i === undefined) {
      i = this.list.length;
      const json = codec.encode(style);
      this.list.push(json === undefined ? 'null' : json);
      this.byKey.set(key, i);
    }
    return i;
  }

  // JSON array of the styles added since the last call.
  takeNew() {
    const json = `[${this.list.slice(this.sent).join(',')}]`;
    this.sent = this.list.length;
    return json;
  }
}

function keyIds(keys) {
  const ids = Object.create(null);
  keys.forEach((k, i) => {
    ids[k] = i;
  });
  return ids;
}
const CELL_IDS = keyIds(CELL_KEYS);
const ROW_IDS = keyIds(ROW_KEYS);

// The model's own enumerable keys, as JSON.stringify would see them.
function writeFields(w, model, ids, styles) {
  for (const key in model) {
    if (!Object.prototype.hasOwnProperty.call(model, key)) continue;
    const v = model[key];
    if (v === undefined) continue;
    const id = ids[key];
    if (id !== undefined) {
      w.field(id, v);
    } else if (key === 'style') {
      w.u8(STYLE);
      w.u8(INDEX);
      w.u32(styles.index(v));
    }
  }
}

function writeRow(w, row, styles) {
  writeFields(w, row, ROW_IDS, styles);
  const cells = row.cells || [];
  w.u8(CELLS);
  w.u32(cells.length);
  for (let i = 0; i < cells.length; i++) {
    writeFields(w, cells[i], CELL_IDS, styles);
    w.u8(END);
  }
  w.u8(END);
}

function writeSheet(w, rows, styles) {
  const list = rows || [];
  w.u32(list.length);
  for (let i = 0; i < list.length; i++) {
    writeRow(w, list[i], styles);
  }
}

// Rows of a streaming worksheet, encoded as they are committed and taken
// in batches for native StreamWorkbook.writeRows.
class RowBatch {
  constructor(styles) {
    this.styles = styles;
    this.reset();
  }

  // One buffer is reused: take()'s result is valid until the next add().
  reset() {
    if (!this.w || this.w.buf.length > 1 << 22) this.w = new Writer();
    this.w.pos = 0;
    this.w.u32(1);
    this.w.u32(0); // row count, set by take()
    this.count = 0;
    this.taken = false;
  }

  add(row) {
    if (this.taken) this.reset();
    writeRow(this.w, row, this.styles);
    this.count++;
  }

  take() {
    if (this.taken) this.reset();
    this.w.buf.writeUInt32LE(this.count, 4);
    const out = this.w.result();
    this.count = 0;
    this.taken = true;
    return out;
  }
}

// workbook.model -> the arguments of native.writeXlsx: the model without
// rows (JSON), every worksheet's rows (Buffer) and the styles (JSON).
function encodeWorkbook(model) {
  const styles = new StyleTable();
  const w = new Writer(1 << 20);
  const sheets = model.worksheets || [];
  w.u32(sheets.length);
  for (const ws of sheets) {
    if (ws && ws._worksheet) {
      // a worksheet's _modelForWrite: its rows, one at a time
      const at = w.pos;
      w.u32(0);
      let count = 0;
      ws.dimensions = new Range();
      ws._worksheet.eachRowModel(row => {
        writeRow(w, row, styles);
        count++;
      }, ws.dimensions);
      w.buf.writeUInt32LE(count, at);
    } else {
      writeSheet(w, ws && ws.rows, styles);
      if (ws) ws.rows = undefined;
    }
  }
  return {model: codec.encode(model), rows: w.result(), styles: styles.takeNew()};
}

// ---------------------------------------------------------------------------
// Reading: the engine's typed rows -> row and cell models

const ROW_NAMES = [...ROW_KEYS];
ROW_NAMES[BEST_FIT] = 'bestFit';
ROW_NAMES[STYLE_VALUE] = 'style';
const CELL_NAMES = [...CELL_KEYS];
CELL_NAMES[STYLE_VALUE] = 'style';

// Reads native readXlsx's rows: the plain shared strings, then each
// worksheet's rows as an independent byte range, decoded on demand. Given
// `strings`, `buf` is one batch of XlsxReader rows instead (no header).
class RowsReader {
  constructor(buf, table, strings) {
    this.buf = buf;
    this.table = table;
    this.pos = 0;
    this.rowNumber = '';
    if (strings) {
      this.strings = strings;
      this.sheets = [];
      return;
    }
    this.strings = readStrings(buf, this);
    this.sheets = new Array(this.u32());
    for (let s = 0; s < this.sheets.length; s++) {
      const len = this.u32();
      this.sheets[s] = this.pos;
      this.pos += len;
    }
  }

  u32() {
    const v = this.buf.readUInt32LE(this.pos);
    this.pos += 4;
    return v;
  }

  f64() {
    const v = this.buf.readDoubleLE(this.pos);
    this.pos += 8;
    return v;
  }

  str() {
    const n = this.u32();
    const s = this.buf.toString('utf8', this.pos, this.pos + n);
    this.pos += n;
    return s;
  }

  value() {
    const tag = this.buf[this.pos++];
    switch (tag) {
      case F64:
        return small(this.f64());
      case STR:
        return this.str();
      case TRUE:
        return true;
      case FALSE:
        return false;
      case NULL:
        return null;
      case DATE:
        return new Date(this.f64());
      case JSON_VALUE: {
        const json = this.str();
        // every row's default style
        return json === '{}' ? {} : codec.parseWithRefs(json, this.table);
      }
      case REF:
        return this.table[this.u32()];
      case SST:
        return this.strings[this.u32()];
      case ADDR:
        return colCache.n2l(this.u32()) + this.rowNumber;
      default:
        throw new Error(`excel-rs: bad row value tag ${tag}`);
    }
  }

  row() {
    const {buf} = this;
    const row = {};
    this.rowNumber = '';
    for (;;) {
      const key = buf[this.pos++];
      if (key === END) return row;
      if (key === CELLS) {
        const cells = new Array(this.u32());
        for (let c = 0; c < cells.length; c++) {
          const cell = {};
          for (;;) {
            const k = buf[this.pos++];
            if (k === END) break;
            cell[CELL_NAMES[k]] = this.value();
          }
          cells[c] = cell;
        }
        row.cells = cells;
      } else {
        const v = this.value();
        row[ROW_NAMES[key]] = v;
        if (key === 0) this.rowNumber = String(v);
      }
    }
  }

  // Calls fn(rowModel, index) for each row of sheet `s`.
  eachRow(s, fn) {
    this.pos = this.sheets[s];
    const n = this.u32();
    for (let r = 0; r < n; r++) {
      fn(this.row(), r);
    }
  }

  // The rows of sheet `s` as an array.
  rows(s) {
    const out = [];
    this.eachRow(s, row => out.push(row));
    return out;
  }

  // The rows of sheet `s` as a source for Worksheet._parseRows: `loadInto`
  // builds the Row and Cell objects straight from the buffer, `forEach`
  // gives the row models.
  lazyRows(s) {
    return {
      forEach: fn => this.eachRow(s, fn),
      loadInto: worksheet => this.loadRows(s, worksheet),
    };
  }

  // What `worksheet._parseRows({rows})` does (a Row per row model, then
  // `row.model = rowModel`), without the row models in between.
  loadRows(s, worksheet) {
    worksheet._rows = [];
    this.pos = this.sheets[s];
    const n = this.u32();
    for (let r = 0; r < n; r++) {
      this.loadRow(worksheet);
    }
  }

  loadRow(worksheet) {
    const {buf} = this;
    let number;
    let height;
    let hidden;
    let outlineLevel;
    let style;
    this.rowNumber = '';
    for (;;) {
      const key = buf[this.pos++];
      if (key === CELLS) break;
      if (key === END) throw new Error('excel-rs: row without cells');
      const v = this.value();
      switch (key) {
        case 0:
          number = v;
          this.rowNumber = String(v);
          break;
        case 3:
          height = v;
          break;
        case 4:
          hidden = v;
          break;
        case 5:
          outlineLevel = v;
          break;
        case STYLE_VALUE:
          style = v;
          break;
        default:
          break;
      }
    }
    const row = new Row(worksheet, number);
    worksheet._rows[row.number - 1] = row;
    // Row.model setter (NaN for a <row> without r)
    if (number !== row._number) {
      throw new Error('Invalid row number in model');
    }
    const count = this.u32();
    let previous;
    for (let c = 0; c < count; c++) {
      const model = {};
      let col = 0;
      for (;;) {
        const k = buf[this.pos++];
        if (k === END) break;
        if (k === 0 && buf[this.pos] === ADDR) {
          this.pos++;
          col = this.u32();
          model.address = colCache.n2l(col) + this.rowNumber;
        } else {
          model[CELL_NAMES[k]] = this.value();
        }
      }
      if (model.type === Cell.Types.Merge) continue;
      let address;
      if (col) {
        address = {row: number, col, address: model.address};
      } else if (model.address) {
        address = colCache.decodeAddress(model.address);
        if (address.row !== number) {
          // B9: the cell is in this row
          address = {row: number, col: address.col, address: colCache.encodeAddress(number, address.col)};
        }
      } else if (previous) {
        address = {row: previous.row, col: previous.col + 1, address: colCache.encodeAddress(previous.row, previous.col + 1)};
      }
      previous = address;
      if (row._cells[address.col - 1]) {
        row.getCellEx(address).model = model;
      } else {
        row._cells[address.col - 1] = new Cell(row, worksheet.getColumn(address.col), address.address, model);
      }
    }
    if (buf[this.pos++] !== END) throw new Error('excel-rs: bad row end');
    if (height) {
      row.height = height;
    } else {
      delete row.height;
    }
    row.hidden = hidden;
    row.outlineLevel = outlineLevel || 0;
    row._setLoadedStyle(style);
  }
}

// The plain shared strings at `at.pos` of `buf` (u32 count, then each
// string); `at.pos` moves past them.
function readStrings(buf, at = {pos: 0}) {
  const n = buf.readUInt32LE(at.pos);
  at.pos += 4;
  const strings = new Array(n);
  for (let i = 0; i < n; i++) {
    const len = buf.readUInt32LE(at.pos);
    strings[i] = buf.toString('utf8', at.pos + 4, at.pos + 4 + len);
    at.pos += 4 + len;
  }
  return strings;
}

// The rows of native XlsxReader, pulled in batches as the worksheets are
// loaded: each worksheet's `rows` is `source(i)`, whose `loadInto` builds
// its Row and Cell objects while the engine reads the next rows.
class RowBatches {
  constructor(reader, table, strings) {
    this.reader = reader;
    this.table = table;
    this.strings = strings;
    this.pending = null;
  }

  // The next batch: {sheet, last, count, rows (a RowsReader at its rows)};
  // the shared objects it adds go into the table.
  next() {
    if (this.pending) {
      const p = this.pending;
      this.pending = null;
      return p;
    }
    const buf = this.reader.nextRows();
    if (!buf) return null;
    const rows = new RowsReader(buf, this.table, this.strings);
    const sheet = rows.u32();
    const last = buf[rows.pos++] === 1;
    codec.addRefs(rows.str(), this.table);
    const count = rows.u32();
    return {sheet, last, count, rows};
  }

  // Calls fn(batch) for each batch of worksheet `i`.
  each(i, fn) {
    for (;;) {
      const b = this.next();
      if (!b) return;
      // batches of a worksheet the model setter left out
      if (b.sheet < i) continue;
      if (b.sheet > i) {
        this.pending = b;
        return;
      }
      fn(b);
      if (b.last) return;
    }
  }

  source(i) {
    return {
      loadInto: worksheet => {
        worksheet._rows = [];
        this.each(i, b => {
          for (let k = 0; k < b.count; k++) b.rows.loadRow(worksheet);
        });
      },
      forEach: fn => {
        let index = 0;
        this.each(i, b => {
          for (let k = 0; k < b.count; k++) fn(b.rows.row(), index++);
        });
      },
    };
  }
}

// Decodes every worksheet's rows into arrays.
function decodeSheets(buf, table) {
  const reader = new RowsReader(buf, table);
  return reader.sheets.map((_, s) => reader.rows(s));
}

module.exports = {decodeSheets, encodeWorkbook, readStrings, RowBatch, RowBatches, RowsReader, StyleTable};
