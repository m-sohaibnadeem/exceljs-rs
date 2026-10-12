// From ExcelJS 4.4.0 lib/stream/xlsx/worksheet-reader.js (MIT, see LICENSE-EXCELJS).
// The SAX state machine runs in the Rust engine on a worker thread and
// sends batches of cols/row/hyperlink events; the cell value rules below are
// ExcelJS's.
'use strict';

const {EventEmitter} = require('events');
const _ = require('../../utils/under-dash');
const utils = require('../../utils/utils');
const colCache = require('../../utils/col-cache');
const Dimensions = require('../../doc/range');
const Row = require('../../doc/row');
const Column = require('../../doc/column');
const codec = require('../../xlsx/codec');

// Event kinds of the engine's batches (excel_rs::EVENT_ROW).
const EVENT_ROW = 1;
const EVENT_NULL_ROW = 2;
const EVENT_HYPERLINK = 3;
const EVENT_COLS = 4;

// Rows per batch from the engine.
const ROW_BATCH = 500;

class WorksheetReader extends EventEmitter {
  constructor({workbook, id, iterator, options}) {
    super();

    this.workbook = workbook;
    this.id = id;
    this.iterator = iterator;
    this.options = options || {};

    // and a name
    this.name = `Sheet${this.id}`;

    // column definitions
    this._columns = null;
    this._keys = {};

    // keep a record of dimensions
    this._dimensions = new Dimensions();
  }

  // destroy - not a valid operation for a streaming writer
  // even though some streamers might be able to, it's a bad idea.
  destroy() {
    throw new Error('Invalid Operation: destroy');
  }

  // return the current dimensions of the writer
  get dimensions() {
    return this._dimensions;
  }

  // =========================================================================
  // Columns

  // get the current columns array.
  get columns() {
    return this._columns;
  }

  // get a single column by col number. If it doesn't exist, it and any gaps before it
  // are created.
  getColumn(c) {
    if (typeof c === 'string') {
      // if it matches a key'd column, return that
      const col = this._keys[c];
      if (col) {
        return col;
      }

      // otherise, assume letter
      c = colCache.l2n(c);
    }
    if (!this._columns) {
      this._columns = [];
    }
    if (c > this._columns.length) {
      let n = this._columns.length + 1;
      while (n <= c) {
        this._columns.push(new Column(this, n++));
      }
    }
    return this._columns[c - 1];
  }

  getColumnKey(key) {
    return this._keys[key];
  }

  setColumnKey(key, value) {
    this._keys[key] = value;
  }

  deleteColumnKey(key) {
    delete this._keys[key];
  }

  eachColumnKey(f) {
    _.each(this._keys, f);
  }

  async read() {
    try {
      for await (const events of this.parse()) {
        for (const {eventType, value} of events) {
          this.emit(eventType, value);
        }
      }
      this.emit('finished');
    } catch (error) {
      this.emit('error', error);
    }
  }

  async *[Symbol.asyncIterator]() {
    for await (const events of this.parse()) {
      for (const {eventType, value} of events) {
        if (eventType === 'row') {
          yield value;
        }
      }
    }
  }

  async *parse() {
    const {iterator, options} = this;
    let emitSheet = false;
    let emitHyperlinks = false;
    let hyperlinks = null;
    switch (options.worksheets) {
      case 'emit':
        emitSheet = true;
        break;
      case 'prep':
        break;
      default:
        break;
    }
    switch (options.hyperlinks) {
      case 'emit':
        emitHyperlinks = true;
        break;
      case 'cache':
        this.hyperlinks = hyperlinks = {};
        break;
      default:
        break;
    }
    if (!emitSheet && !emitHyperlinks && !hyperlinks) {
      return;
    }

    // references
    const {sharedStrings, styles, properties} = this.workbook;

    // eslint-disable-next-line no-underscore-dangle
    const cursor = this.workbook._archive.openWorksheet(
      iterator.index,
      iterator.path,
      emitSheet,
      !!(emitHyperlinks || hyperlinks),
      ROW_BATCH
    );
    for (;;) {
      const batch = await cursor.next();
      if (batch === null) {
        break;
      }
      const worksheetEvents = [];
      // The batch layout is documented at excel_rs::EVENT_ROW.
      let pos = 0;
      const f64 = () => {
        const v = batch.readDoubleLE(pos);
        pos += 8;
        return v;
      };
      const str = () => {
        const n = batch.readUInt32LE(pos);
        const v = batch.toString('utf8', pos + 4, pos + 4 + n);
        pos += 4 + n;
        return v;
      };
      const opt = () => (batch[pos++] === 0 ? undefined : str());
      while (pos < batch.length) {
        const eventType = batch[pos++];
        switch (eventType) {
          case EVENT_COLS:
            this._columns = Column.fromModel(codec.decode(str()));
            break;
          case EVENT_HYPERLINK: {
            const ref = opt();
            const rId = opt();
            const value = {ref, rId};
            if (emitHyperlinks) {
              worksheetEvents.push({eventType: 'hyperlink', value});
            } else {
              hyperlinks[value.ref] = value;
            }
            break;
          }
          case EVENT_NULL_ROW:
          case EVENT_ROW: {
            let row = null;
            if (eventType === EVENT_ROW) {
              row = new Row(this, f64());
              if (batch[pos++]) {
                row.height = f64();
              }
              if (batch[pos++]) {
                const style = styles.getStyleModel(f64());
                if (style) {
                  row.style = style;
                }
              }
              const n = batch.readUInt32LE(pos);
              pos += 4;
              let col = 0;
              for (let i = 0; i < n; i++) {
                const c = {};
                const ref = opt();
                if (ref !== undefined) c.ref = ref;
                c.s = f64();
                const t = opt();
                if (t !== undefined) c.t = t;
                const f = opt();
                if (f !== undefined) c.f = f;
                const v = opt();
                if (v !== undefined) c.v = v;
                col = this._readCell(row, c, col, sharedStrings, styles, properties, hyperlinks);
              }
            }
            this._dimensions.expandRow(row);
            worksheetEvents.push({eventType: 'row', value: row});
            break;
          }
          default:
            throw new Error(`excel-rs: bad worksheet event ${eventType}`);
        }
      }
      if (worksheetEvents.length > 0) {
        yield worksheetEvents;
      }
    }
  }

  // The closing </c> of ExcelJS's parse loop; returns the cell's column.
  // A <c> without r is the next column, as in the xlsx reader. (ExcelJS
  // throws: B14 in COMPATIBILITY.md.)
  _readCell(row, c, previousCol, sharedStrings, styles, properties, hyperlinks) {
    const col = c.ref === undefined ? previousCol + 1 : colCache.decodeAddress(c.ref).col;
    const cell = row.getCell(col);
    if (c.s) {
      const style = styles.getStyleModel(c.s);
      if (style) {
        cell.style = style;
      }
    }
    if (c.f !== undefined) {
      const cellValue = {
        formula: c.f,
      };
      if (c.v !== undefined) {
        if (c.t === 'str') {
          // The XML parser has decoded the text; ExcelJS decodes it again
          // (B2 in COMPATIBILITY.md).
          cellValue.result = c.v;
        } else {
          cellValue.result = parseFloat(c.v);
        }
      }
      cell.value = cellValue;
    } else if (c.v !== undefined) {
      switch (c.t) {
        case 's': {
          const index = parseInt(c.v, 10);
          if (sharedStrings) {
            cell.value = sharedStrings[index];
          } else {
            cell.value = {
              sharedString: index,
            };
          }
          break;
        }
        case 'inlineStr':
        case 'str':
          cell.value = c.v;
          break;
        case 'e':
          cell.value = {error: c.v};
          break;
        case 'b':
          cell.value = utils.parseBooleanText(c.v);
          break;
        default:
          if (utils.isDateFmt(cell.numFmt)) {
            cell.value = utils.excelToDate(parseFloat(c.v), properties.model && properties.model.date1904);
          } else {
            cell.value = parseFloat(c.v);
          }
          break;
      }
    }
    if (hyperlinks) {
      const hyperlink = hyperlinks[c.ref === undefined ? cell.address : c.ref];
      if (hyperlink) {
        cell.text = cell.value;
        cell.value = undefined;
        cell.hyperlink = hyperlink;
      }
    }
    return col;
  }
}

module.exports = WorksheetReader;
