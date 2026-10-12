// From ExcelJS 4.4.0 lib/doc/row.js (MIT, see LICENSE-EXCELJS).
// Object model; file I/O goes through the Rust engine.
'use strict';

const _ = require('../utils/under-dash');

const Enums = require('./enums');
const colCache = require('../utils/col-cache');
const Cell = require('./cell');

function isEmptyObject(o) {
  if (typeof o !== 'object' || Array.isArray(o)) return false;
  for (const k in o) {
    if (Object.prototype.hasOwnProperty.call(o, k)) return false;
  }
  return true;
}

class Row {
  constructor(worksheet, number) {
    this._worksheet = worksheet;
    this._number = number;
    this._cells = [];
    this.style = {};
    this.outlineLevel = 0;
  }

  // return the row number
  get number() {
    return this._number;
  }

  get worksheet() {
    return this._worksheet;
  }

  // Inform Streaming Writer that this row (and all rows before it) are complete
  // and ready to write. Has no effect on Worksheet document
  commit() {
    this._worksheet._commitRow(this); // eslint-disable-line no-underscore-dangle
  }

  // helps GC by breaking cyclic references
  destroy() {
    delete this._worksheet;
    delete this._cells;
    delete this.style;
  }

  findCell(colNumber) {
    return this._cells[colNumber - 1];
  }

  // given {address, row, col}, find or create new cell
  getCellEx(address) {
    let cell = this._cells[address.col - 1];
    if (!cell) {
      const column = this._worksheet.getColumn(address.col);
      cell = new Cell(this, column, address.address);
      this._cells[address.col - 1] = cell;
    }
    return cell;
  }

  // get cell by key, letter or column number
  getCell(col) {
    if (typeof col === 'string') {
      // is it a key?
      const column = this._worksheet.getColumnKey(col);
      if (column) {
        col = column.number;
      } else {
        col = colCache.l2n(col);
      }
    }
    return (
      this._cells[col - 1] ||
      this.getCellEx({
        address: colCache.encodeAddress(this._number, col),
        row: this._number,
        col,
      })
    );
  }

  // remove cell(s) and shift all higher cells down by count
  splice(start, count, ...inserts) {
    const nKeep = start + count;
    const nExpand = inserts.length - count;
    const nEnd = this._cells.length;
    let i;
    let cSrc;
    let cDst;

    if (nExpand < 0) {
      // remove cells
      for (i = start + inserts.length; i <= nEnd; i++) {
        cDst = this._cells[i - 1];
        cSrc = this._cells[i - nExpand - 1];
        if (cSrc) {
          cDst = this.getCell(i);
          cDst.value = cSrc.value;
          cDst.style = cSrc.style;
          // eslint-disable-next-line no-underscore-dangle
          cDst._comment = cSrc._comment;
        } else if (cDst) {
          cDst.value = null;
          cDst.style = {};
          // eslint-disable-next-line no-underscore-dangle
          cDst._comment = undefined;
        }
      }
    } else if (nExpand > 0) {
      // insert new cells
      for (i = nEnd; i >= nKeep; i--) {
        cSrc = this._cells[i - 1];
        if (cSrc) {
          cDst = this.getCell(i + nExpand);
          cDst.value = cSrc.value;
          cDst.style = cSrc.style;
          // eslint-disable-next-line no-underscore-dangle
          cDst._comment = cSrc._comment;
        } else {
          this._cells[i + nExpand - 1] = undefined;
        }
      }
    }

    // now add the new values
    for (i = 0; i < inserts.length; i++) {
      cDst = this.getCell(start + i);
      cDst.value = inserts[i];
      cDst.style = {};
      // eslint-disable-next-line no-underscore-dangle
      cDst._comment = undefined;
    }
  }

  // Iterate over all non-null cells in this row
  eachCell(options, iteratee) {
    if (!iteratee) {
      iteratee = options;
      options = null;
    }
    if (options && options.includeEmpty) {
      const n = this._cells.length;
      for (let i = 1; i <= n; i++) {
        iteratee(this.getCell(i), i);
      }
    } else {
      this._cells.forEach((cell, index) => {
        if (cell && cell.type !== Enums.ValueType.Null) {
          iteratee(cell, index + 1);
        }
      });
    }
  }

  // ===========================================================================
  // Page Breaks
  addPageBreak(lft, rght) {
    const ws = this._worksheet;
    const left = Math.max(0, lft - 1) || 0;
    const right = Math.max(0, rght - 1) || 16838;
    const pb = {
      id: this._number,
      max: right,
      man: 1,
    };
    if (left) pb.min = left;

    ws.rowBreaks.push(pb);
  }

  // return a sparse array of cell values
  get values() {
    const values = [];
    this._cells.forEach(cell => {
      if (cell && cell.type !== Enums.ValueType.Null) {
        values[cell.col] = cell.value;
      }
    });
    return values;
  }

  // set the values by contiguous or sparse array, or by key'd object literal
  set values(value) {
    // this operation is not additive - any prior cells are removed
    this._cells = [];
    if (!value) {
      // empty row
    } else if (value instanceof Array) {
      let offset = 0;
      if (value.hasOwnProperty('0')) {
        // contiguous array - start at column 1
        offset = 1;
      }
      value.forEach((item, index) => {
        if (item !== undefined) {
          this._setNewValue(index + offset, item);
        }
      });
    } else {
      // assume object with column keys
      this._worksheet.eachColumnKey((column, key) => {
        if (value[key] !== undefined) {
          this._setNewValue(column.number, value[key]);
        }
      });
    }
  }

  // (excel-rs) `getCellEx({address, row, col}).value = value` for the values
  // setter, which has just emptied the row
  _setNewValue(col, value) {
    if (this._cells[col - 1]) {
      this._cells[col - 1].value = value;
      return;
    }
    const column = this._worksheet.getColumn(col);
    this._cells[col - 1] = Cell.withValue(this, column, colCache.encodeAddress(this._number, col), value);
  }

  // returns true if the row includes at least one cell with a value
  get hasValues() {
    return _.some(this._cells, cell => cell && cell.type !== Enums.ValueType.Null);
  }

  get cellCount() {
    return this._cells.length;
  }

  get actualCellCount() {
    let count = 0;
    this.eachCell(() => {
      count++;
    });
    return count;
  }

  // get the min and max column number for the non-null cells in this row or null
  get dimensions() {
    let min = 0;
    let max = 0;
    this._cells.forEach(cell => {
      if (cell && cell.type !== Enums.ValueType.Null) {
        if (!min || min > cell.col) {
          min = cell.col;
        }
        if (max < cell.col) {
          max = cell.col;
        }
      }
    });
    return min > 0
      ? {
          min,
          max,
        }
      : null;
  }

  // =========================================================================
  // styles
  _applyStyle(name, value) {
    this.style[name] = value;
    this._cells.forEach(cell => {
      if (cell) {
        cell[name] = value;
      }
    });
    return value;
  }

  get numFmt() {
    return this.style.numFmt;
  }

  set numFmt(value) {
    this._applyStyle('numFmt', value);
  }

  get font() {
    return this.style.font;
  }

  set font(value) {
    this._applyStyle('font', value);
  }

  get alignment() {
    return this.style.alignment;
  }

  set alignment(value) {
    this._applyStyle('alignment', value);
  }

  get protection() {
    return this.style.protection;
  }

  set protection(value) {
    this._applyStyle('protection', value);
  }

  get border() {
    return this.style.border;
  }

  set border(value) {
    this._applyStyle('border', value);
  }

  get fill() {
    return this.style.fill;
  }

  set fill(value) {
    this._applyStyle('fill', value);
  }

  get hidden() {
    return !!this._hidden;
  }

  set hidden(value) {
    this._hidden = value;
  }

  get outlineLevel() {
    return this._outlineLevel || 0;
  }

  set outlineLevel(value) {
    this._outlineLevel = value;
  }

  get collapsed() {
    return !!(
      this._outlineLevel && this._outlineLevel >= this._worksheet.properties.outlineLevelRow
    );
  }

  // =========================================================================
  get model() {
    const cells = [];
    let min = 0;
    let max = 0;
    this._cells.forEach(cell => {
      if (cell) {
        const cellModel = cell.model;
        if (cellModel) {
          if (!min || min > cell.col) {
            min = cell.col;
          }
          if (max < cell.col) {
            max = cell.col;
          }
          cells.push(cellModel);
        }
      }
    });

    return this.height || cells.length
      ? {
          cells,
          number: this.number,
          min,
          max,
          height: this.height,
          style: this.style,
          hidden: this.hidden,
          outlineLevel: this.outlineLevel,
          collapsed: this.collapsed,
        }
      : null;
  }

  set model(value) {
    if (value.number !== this._number) {
      throw new Error('Invalid row number in model');
    }
    this._cells = [];
    let previousAddress;
    value.cells.forEach(cellModel => {
      switch (cellModel.type) {
        case Cell.Types.Merge:
          // special case - don't add this types
          break;
        default: {
          let address;
          if (cellModel.address) {
            address = colCache.decodeAddress(cellModel.address);
            if (address.row !== this._number) {
              // <row r="8"><c r="A99">: the cell is in row 8, so its address
              // is A8. (ExcelJS keeps "A99": B9 in COMPATIBILITY.md.)
              const {col} = address;
              address = {
                row: this._number,
                col,
                address: colCache.encodeAddress(this._number, col),
                $col$row: `$${colCache.n2l(col)}$${this._number}`,
              };
            }
          } else if (previousAddress) {
            // This is a <c> element without an r attribute
            // Assume that it's the cell for the next column
            const {row} = previousAddress;
            const col = previousAddress.col + 1;
            address = {
              row,
              col,
              address: colCache.encodeAddress(row, col),
              $col$row: `$${colCache.n2l(col)}$${row}`,
            };
          }
          previousAddress = address;
          if (this._cells[address.col - 1]) {
            this.getCellEx(address).model = cellModel;
          } else {
            const column = this._worksheet.getColumn(address.col);
            this._cells[address.col - 1] = new Cell(this, column, address.address, cellModel);
          }
          break;
        }
      }
    });

    if (value.height) {
      this.height = value.height;
    } else {
      delete this.height;
    }

    this.hidden = value.hidden;
    this.outlineLevel = value.outlineLevel || 0;

    this._setLoadedStyle(value.style);
  }

  // `this.style` of the model setter: a copy (excel-rs: an empty style needs no
  // JSON round trip).
  _setLoadedStyle(style) {
    this.style = (style && (isEmptyObject(style) ? {} : JSON.parse(JSON.stringify(style)))) || {};
  }
}

module.exports = Row;
