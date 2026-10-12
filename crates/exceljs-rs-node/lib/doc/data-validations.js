// From ExcelJS 4.4.0 lib/doc/data-validations.js (MIT, see LICENSE-EXCELJS).
// Object model; file I/O goes through the Rust engine.
class DataValidations {
  constructor(model) {
    this.model = model || {};
  }

  add(address, validation) {
    return (this.model[address] = validation);
  }

  find(address) {
    return this.model[address];
  }

  remove(address) {
    this.model[address] = undefined;
  }
}

module.exports = DataValidations;
