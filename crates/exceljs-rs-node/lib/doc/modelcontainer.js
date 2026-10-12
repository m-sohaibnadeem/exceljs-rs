// From ExcelJS 4.4.0 lib/doc/modelcontainer.js (MIT, see LICENSE-EXCELJS).
// Object model; file I/O goes through the Rust engine.
'use strict';

const XLSX = require('../xlsx/xlsx');

class ModelContainer {
  constructor(model) {
    this.model = model;
  }

  get xlsx() {
    if (!this._xlsx) {
      this._xlsx = new XLSX(this);
    }
    return this._xlsx;
  }
}

module.exports = ModelContainer;
