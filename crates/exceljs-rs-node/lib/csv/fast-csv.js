// The parts of @fast-csv/parse and @fast-csv/format 4.3.6 (MIT) that
// ExcelJS's CSV code uses. Tokenizing runs in the Rust engine
// (native.parseCsv, a port of Parser/RowParser/ColumnParser/Scanner); the
// row handling (headers, skipLines, skipRows, maxRows) and the formatter
// are ported here, since they call user functions.
'use strict';

const {StringDecoder} = require('string_decoder');
const native = require('../native');

// ---------------------------------------------------------------------------
// parse

class ParserOptions {
  constructor(opts) {
    this.objectMode = true;
    this.delimiter = ',';
    this.ignoreEmpty = false;
    this.quote = '"';
    this.escape = null;
    this.escapeChar = this.quote;
    this.comment = null;
    this.supportsComments = false;
    this.ltrim = false;
    this.rtrim = false;
    this.trim = false;
    this.headers = null;
    this.renameHeaders = false;
    this.strictColumnHandling = false;
    this.discardUnmappedColumns = false;
    this.carriageReturn = '\r';
    this.encoding = 'utf8';
    this.limitRows = false;
    this.maxRows = 0;
    this.skipLines = 0;
    this.skipRows = 0;
    Object.assign(this, opts || {});
    if (this.delimiter.length > 1) {
      throw new Error('delimiter option must be one character long');
    }
    this.escapeChar = this.escape !== null && this.escape !== undefined ? this.escape : this.quote;
    this.supportsComments = this.comment !== null && this.comment !== undefined;
    if (this.maxRows > 0) {
      this.limitRows = true;
    }
  }

  // What the engine's tokenizer reads. Option values compare with tokens by
  // ===, so anything but a string never matches one.
  toEngine() {
    const str = v => (typeof v === 'string' ? v : null);
    return JSON.stringify({
      delimiter: typeof this.delimiter === 'string' ? this.delimiter : String(this.delimiter),
      quote: str(this.quote),
      escapeChar: str(this.escapeChar),
      comment: this.supportsComments ? str(this.comment) : null,
      ltrim: !!this.ltrim,
      rtrim: !!this.rtrim,
      trim: !!this.trim,
      ignoreEmpty: !!this.ignoreEmpty,
      carriageReturn: str(this.carriageReturn),
    });
  }
}

const isFunction = v => typeof v === 'function';

class HeaderTransformer {
  constructor(parserOptions) {
    this.headers = null;
    this.receivedHeaders = false;
    this.shouldUseFirstRow = false;
    this.processedFirstRow = false;
    this.headersLength = 0;
    this.parserOptions = parserOptions;
    if (parserOptions.headers === true) {
      this.shouldUseFirstRow = true;
    } else if (Array.isArray(parserOptions.headers)) {
      this.setHeaders(parserOptions.headers);
    } else if (isFunction(parserOptions.headers)) {
      this.headersTransform = parserOptions.headers;
    }
  }

  transform(row) {
    if (!this.shouldMapRow(row)) {
      return {row: null, isValid: true};
    }
    return this.processRow(row);
  }

  shouldMapRow(row) {
    const {parserOptions} = this;
    if (!this.headersTransform && parserOptions.renameHeaders && !this.processedFirstRow) {
      if (!this.receivedHeaders) {
        throw new Error('Error renaming headers: new headers must be provided in an array');
      }
      this.processedFirstRow = true;
      return false;
    }
    if (!this.receivedHeaders && Array.isArray(row)) {
      if (this.headersTransform) {
        this.setHeaders(this.headersTransform(row));
      } else if (this.shouldUseFirstRow) {
        this.setHeaders(row);
      } else {
        return true;
      }
      return false;
    }
    return true;
  }

  processRow(row) {
    if (!this.headers) {
      return {row, isValid: true};
    }
    const {parserOptions} = this;
    if (!parserOptions.discardUnmappedColumns && row.length > this.headersLength) {
      if (!parserOptions.strictColumnHandling) {
        throw new Error(
          `Unexpected Error: column header mismatch expected: ${this.headersLength} columns got: ${row.length}`
        );
      }
      return {
        row,
        isValid: false,
        reason: `Column header mismatch expected: ${this.headersLength} columns got: ${row.length}`,
      };
    }
    if (parserOptions.strictColumnHandling && row.length < this.headersLength) {
      return {
        row,
        isValid: false,
        reason: `Column header mismatch expected: ${this.headersLength} columns got: ${row.length}`,
      };
    }
    return {row: this.mapHeaders(row), isValid: true};
  }

  mapHeaders(row) {
    const rowMap = {};
    const {headers, headersLength} = this;
    for (let i = 0; i < headersLength; i += 1) {
      const header = headers[i];
      if (header !== undefined) {
        const val = row[i];
        rowMap[header] = val === undefined ? '' : val;
      }
    }
    return rowMap;
  }

  setHeaders(headers) {
    const filteredHeaders = headers.filter(h => !!h);
    if (new Set(filteredHeaders).size !== filteredHeaders.length) {
      const counts = {};
      filteredHeaders.forEach(h => {
        counts[h] = (counts[h] || 0) + 1;
      });
      const duplicates = Object.keys(counts).filter(dup => counts[dup] > 1);
      throw new Error(`Duplicate headers found ${JSON.stringify(duplicates)}`);
    }
    this.headers = headers;
    this.receivedHeaders = true;
    this.headersLength = (this.headers && this.headers.length) || 0;
  }
}

// CsvParserStream over a whole input: `onData(row)` gets each row;
// `onHeaders(headers)` the headers. Returns the row count.
// fast-csv's parser stream: write(chunk) is _transform, end() is _flush
// (and returns the row count). Each chunk is tokenized as it comes, with
// an unfinished line carried to the next.
function createParser(opts, onData) {
  const parserOptions = new ParserOptions(opts);
  const engineOptions = parserOptions.toEngine();
  const headerTransformer = new HeaderTransformer(parserOptions);
  let lines = '';
  let rowCount = 0;
  let parsedRowCount = 0;
  let parsedLineCount = 0;
  const hasHitRowLimit = () => parserOptions.limitRows && rowCount >= parserOptions.maxRows;

  const parseChunk = (data, hasMoreData) => {
    if (!data) {
      return [];
    }
    const {rows, rest} = native.parseCsv(data, engineOptions, hasMoreData);
    lines = rest;
    return JSON.parse(rows);
  };

  const processRows = rows => {
    for (let i = 0; i < rows.length && !hasHitRowLimit(); i++) {
      parsedLineCount += 1;
      if (parsedLineCount <= parserOptions.skipLines) {
        continue;
      }
      rowCount += 1;
      parsedRowCount += 1;
      const withHeaders = headerTransformer.transform(rows[i]);
      if (!withHeaders.isValid) {
        if (parsedRowCount <= parserOptions.skipRows) {
          rowCount -= 1;
        }
        // invalid rows go to 'data-invalid', which ExcelJS does not listen to
        continue;
      }
      if (withHeaders.row) {
        if (parsedRowCount > parserOptions.skipRows) {
          onData(withHeaders.row);
        } else {
          rowCount -= 1;
        }
        continue;
      }
      // a header row
      rowCount -= 1;
      parsedRowCount -= 1;
    }
  };

  const decoder = new StringDecoder(parserOptions.encoding);
  return {
    write(chunk) {
      const text = typeof chunk === 'string' ? chunk : decoder.write(chunk);
      if (!hasHitRowLimit()) processRows(parseChunk(lines + text, true));
    },
    end() {
      if (!hasHitRowLimit()) processRows(parseChunk(lines + decoder.end(), false));
      return rowCount;
    },
  };
}

// All of `input` at once: _transform, then _flush.
function parse(input, opts, onData) {
  const parser = createParser(opts, onData);
  parser.write(input);
  return parser.end();
}

// ---------------------------------------------------------------------------
// format

class FormatterOptions {
  constructor(opts = {}) {
    this.objectMode = true;
    this.delimiter = ',';
    this.rowDelimiter = '\n';
    this.quote = '"';
    this.escape = this.quote;
    this.quoteColumns = false;
    this.quoteHeaders = this.quoteColumns;
    this.headers = null;
    this.includeEndRowDelimiter = false;
    this.writeBOM = false;
    this.BOM = '﻿';
    this.alwaysWriteHeaders = false;
    Object.assign(this, opts || {});
    if (typeof (opts && opts.quoteHeaders) === 'undefined') {
      this.quoteHeaders = this.quoteColumns;
    }
    if (opts && opts.quote === true) {
      this.quote = '"';
    } else if (opts && opts.quote === false) {
      this.quote = '';
    }
    if (typeof (opts && opts.escape) !== 'string') {
      this.escape = this.quote;
    }
    this.shouldWriteHeaders = !!this.headers && (opts.writeHeaders !== null && opts.writeHeaders !== undefined ? opts.writeHeaders : true);
    this.headers = Array.isArray(this.headers) ? this.headers : null;
    this.escapedQuote = `${this.escape}${this.quote}`;
  }
}

const escapeRegExp = s => String(s).replace(/[\\^$.*+?()[\]{}|]/g, '\\$&');

class FieldFormatter {
  constructor(formatterOptions) {
    this._headers = null;
    this.formatterOptions = formatterOptions;
    if (formatterOptions.headers !== null) {
      this.headers = formatterOptions.headers;
    }
    this.REPLACE_REGEXP = new RegExp(formatterOptions.quote, 'g');
    const escapePattern = `[${formatterOptions.delimiter}${escapeRegExp(formatterOptions.rowDelimiter)}|\r|\n]`;
    this.ESCAPE_REGEXP = new RegExp(escapePattern);
  }

  set headers(headers) {
    this._headers = headers;
  }

  shouldQuote(fieldIndex, isHeader) {
    const quoteConfig = isHeader ? this.formatterOptions.quoteHeaders : this.formatterOptions.quoteColumns;
    if (typeof quoteConfig === 'boolean') {
      return quoteConfig;
    }
    if (Array.isArray(quoteConfig)) {
      return quoteConfig[fieldIndex];
    }
    if (this._headers !== null) {
      return quoteConfig[this._headers[fieldIndex]];
    }
    return false;
  }

  format(field, fieldIndex, isHeader) {
    const preparedField = `${field === null || field === undefined ? '' : field}`.replace(/\0/g, '');
    const {formatterOptions} = this;
    if (formatterOptions.quote !== '') {
      const shouldEscape = preparedField.indexOf(formatterOptions.quote) !== -1;
      if (shouldEscape) {
        return this.quoteField(preparedField.replace(this.REPLACE_REGEXP, formatterOptions.escapedQuote));
      }
    }
    const hasEscapeCharacters = preparedField.search(this.ESCAPE_REGEXP) !== -1;
    if (hasEscapeCharacters || this.shouldQuote(fieldIndex, isHeader)) {
      return this.quoteField(preparedField);
    }
    return preparedField;
  }

  quoteField(field) {
    const {quote} = this.formatterOptions;
    return `${quote}${field}${quote}`;
  }
}

// lodash.isequal for the values a CSV row holds.
function isEqual(a, b) {
  if (a === b || (a !== a && b !== b)) return true; // eslint-disable-line no-self-compare
  if (typeof a !== 'object' || typeof b !== 'object' || a === null || b === null) return false;
  if (Object.getPrototypeOf(a) !== Object.getPrototypeOf(b)) return false;
  if (a instanceof Date) return a.getTime() === b.getTime();
  const ka = Object.keys(a);
  const kb = Object.keys(b);
  return ka.length === kb.length && ka.every(k => Object.prototype.hasOwnProperty.call(b, k) && isEqual(a[k], b[k]));
}

class RowFormatter {
  constructor(formatterOptions) {
    this.rowCount = 0;
    this.formatterOptions = formatterOptions;
    this.fieldFormatter = new FieldFormatter(formatterOptions);
    this.headers = formatterOptions.headers;
    this.shouldWriteHeaders = formatterOptions.shouldWriteHeaders;
    this.hasWrittenHeaders = false;
    if (this.headers !== null) {
      this.fieldFormatter.headers = this.headers;
    }
    if (formatterOptions.transform) {
      if (!isFunction(formatterOptions.transform)) {
        throw new TypeError('The transform should be a function');
      }
      this.rowTransform = formatterOptions.transform;
    }
  }

  static isRowHashArray(row) {
    if (Array.isArray(row)) {
      return Array.isArray(row[0]) && row[0].length === 2;
    }
    return false;
  }

  static isRowArray(row) {
    return Array.isArray(row) && !this.isRowHashArray(row);
  }

  static gatherHeaders(row) {
    if (RowFormatter.isRowHashArray(row)) {
      return row.map(it => it[0]);
    }
    if (Array.isArray(row)) {
      return row;
    }
    return Object.keys(row);
  }

  // Synchronous transforms only (`transform.length === 1`), as ExcelJS
  // passes none of its own.
  format(row) {
    const transformedRow = this.rowTransform ? this.rowTransform(row) : row;
    if (!row) {
      return [];
    }
    const rows = [];
    if (transformedRow) {
      const {shouldFormatColumns, headers} = this.checkHeaders(transformedRow);
      if (this.shouldWriteHeaders && headers && !this.hasWrittenHeaders) {
        rows.push(this.formatColumns(headers, true));
        this.hasWrittenHeaders = true;
      }
      if (shouldFormatColumns) {
        const columns = this.gatherColumns(transformedRow);
        rows.push(this.formatColumns(columns, false));
      }
    }
    return rows;
  }

  finish() {
    const rows = [];
    if (this.formatterOptions.alwaysWriteHeaders && this.rowCount === 0) {
      if (!this.headers) {
        throw new Error('`alwaysWriteHeaders` option is set to true but `headers` option not provided.');
      }
      rows.push(this.formatColumns(this.headers, true));
    }
    if (this.formatterOptions.includeEndRowDelimiter) {
      rows.push(this.formatterOptions.rowDelimiter);
    }
    return rows;
  }

  checkHeaders(row) {
    if (this.headers) {
      return {shouldFormatColumns: true, headers: this.headers};
    }
    const headers = RowFormatter.gatherHeaders(row);
    this.headers = headers;
    this.fieldFormatter.headers = headers;
    if (!this.shouldWriteHeaders) {
      return {shouldFormatColumns: true, headers: null};
    }
    return {shouldFormatColumns: !isEqual(headers, row), headers};
  }

  gatherColumns(row) {
    if (this.headers === null) {
      throw new Error('Headers is currently null');
    }
    if (!Array.isArray(row)) {
      return this.headers.map(header => row[header]);
    }
    if (RowFormatter.isRowHashArray(row)) {
      return this.headers.map((header, i) => {
        const col = row[i];
        if (col) {
          return col[1];
        }
        return '';
      });
    }
    if (RowFormatter.isRowArray(row) && !this.shouldWriteHeaders) {
      return row;
    }
    return this.headers.map((header, i) => row[i]);
  }

  formatColumns(columns, isHeadersRow) {
    const formattedCols = columns
      .map((field, i) => this.fieldFormatter.format(field, i, isHeadersRow))
      .join(this.formatterOptions.delimiter);
    const {rowCount} = this;
    this.rowCount += 1;
    if (rowCount) {
      return [this.formatterOptions.rowDelimiter, formattedCols].join('');
    }
    return formattedCols;
  }
}

// CsvFormatterStream as a batch: collects the output of `write(row)` calls.
class Formatter {
  constructor(opts) {
    this.formatterOptions = new FormatterOptions(opts);
    this.rowFormatter = new RowFormatter(this.formatterOptions);
    this.chunks = [];
    this.size = 0;
    this.hasWrittenBOM = !this.formatterOptions.writeBOM;
  }

  write(row) {
    if (!this.hasWrittenBOM) {
      this.chunks.push(this.formatterOptions.BOM);
      this.hasWrittenBOM = true;
    }
    // writableObjectMode: a null row ends the stream in Node; ExcelJS never
    // writes one.
    for (const r of this.rowFormatter.format(row)) {
      this.chunks.push(r);
      this.size += r.length;
    }
  }

  // (exceljs-rs) The output so far, which is then dropped: the writer passes it on
  // in pieces instead of holding the whole file.
  take() {
    const out = Buffer.from(this.chunks.join(''), 'utf8');
    this.chunks = [];
    this.size = 0;
    return out;
  }

  end() {
    for (const r of this.rowFormatter.finish()) {
      this.chunks.push(r);
    }
    return this.take();
  }
}

module.exports = {createParser, parse, Formatter};
