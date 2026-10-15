'use strict';

// Public entry point, same shape as ExcelJS lib/exceljs.nodejs.js.

const ExcelJS = {
  Workbook: require('./doc/workbook'),
  ModelContainer: require('./doc/modelcontainer'),
  stream: {
    xlsx: {
      WorkbookWriter: require('./stream/xlsx/workbook-writer'),
      WorkbookReader: require('./stream/xlsx/workbook-reader'),
      readBatches: require('./stream/xlsx/read-batches'),
    },
  },
};

Object.assign(ExcelJS, require('./doc/enums'));

module.exports = ExcelJS;
