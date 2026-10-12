// Replaces ExcelJS's spec/utils/verquire.js (installed by fetch-upstream.sh).
//
// verquire('exceljs') is the library under test:
//   EXCELRS_TARGET=exceljs-rs (default) -> crates/exceljs-rs-node
//   EXCELRS_TARGET=exceljs        -> the real ExcelJS 4.4.0 (baseline)
// Any other path is one of ExcelJS's internal modules. Specs in the suite
// only use those as test helpers, so they always come from real ExcelJS.

/* eslint-disable import/no-dynamic-require, global-require */
const path = require('path');

const target = process.env.EXCELRS_TARGET || 'exceljs-rs';
const exceljsLib = path.dirname(require.resolve('exceljs/lib/exceljs.nodejs.js'));
const libs = {
  exceljs:
    target === 'exceljs'
      ? require('exceljs/lib/exceljs.nodejs.js')
      : require(path.join(__dirname, '..', '..', '..', '..', 'crates', 'exceljs-rs-node')),
};

module.exports = function verquire(p) {
  if (!libs[p]) libs[p] = require(path.join(exceljsLib, p));
  return libs[p];
};
