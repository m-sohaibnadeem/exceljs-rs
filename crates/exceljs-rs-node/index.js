'use strict';

const api = require('./lib/exceljs');

// Explicit assignments let Node.js detect named exports for native ESM imports.
exports.Workbook = api.Workbook;
exports.ModelContainer = api.ModelContainer;
exports.stream = api.stream;
exports.ValueType = api.ValueType;
exports.FormulaType = api.FormulaType;
exports.RelationshipType = api.RelationshipType;
exports.DocumentType = api.DocumentType;
exports.ReadingOrder = api.ReadingOrder;
exports.ErrorValue = api.ErrorValue;
