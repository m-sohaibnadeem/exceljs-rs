// From ExcelJS 4.4.0 lib/csv/csv.js (MIT, see LICENSE-EXCELJS).
// fast-csv and dayjs are replaced by ports (./fast-csv.js, ../utils/dayjs-lite.js);
// CSV tokenizing runs in the Rust engine, chunk by chunk as the input
// arrives, as fast-csv does.
'use strict';

const fs = require('fs');
const fastCsv = require('./fast-csv');
const dayjs = require('../utils/dayjs-lite');
const StreamBuf = require('../utils/stream-buf');

const {
  fs: {exists},
} = require('../utils/utils');

/* eslint-disable quote-props */
const SpecialValues = {
  true: true,
  false: false,
  '#N/A': {error: '#N/A'},
  '#REF!': {error: '#REF!'},
  '#NAME?': {error: '#NAME?'},
  '#DIV/0!': {error: '#DIV/0!'},
  '#NULL!': {error: '#NULL!'},
  '#VALUE!': {error: '#VALUE!'},
  '#NUM!': {error: '#NUM!'},
};
/* eslint-ensable quote-props */

class CSV {
  constructor(workbook) {
    this.workbook = workbook;
    this.worksheet = null;
  }

  async readFile(filename, options) {
    options = options || {};
    if (!(await exists(filename))) {
      throw new Error(`File not found: ${filename}`);
    }
    const stream = fs.createReadStream(filename);
    try {
      return await this.read(stream, options);
    } finally {
      stream.destroy();
    }
  }

  read(stream, options) {
    options = options || {};

    return new Promise((resolve, reject) => {
      const worksheet = this.workbook.addWorksheet(options.sheetName);

      const dateFormats = options.dateFormats || [
        'YYYY-MM-DD[T]HH:mm:ssZ',
        'YYYY-MM-DD[T]HH:mm:ss',
        'MM-DD-YYYY',
        'YYYY-MM-DD',
      ];
      const map =
        options.map ||
        function(datum) {
          if (datum === '') {
            return null;
          }
          const datumNumber = Number(datum);
          if (!Number.isNaN(datumNumber) && datumNumber !== Infinity) {
            return datumNumber;
          }
          const dt = dateFormats.reduce((matchingDate, currentDateFormat) => {
            if (matchingDate) {
              return matchingDate;
            }
            const date = dayjs.parseStrict(datum, currentDateFormat);
            if (dayjs.isValid(date)) {
              return date;
            }
            return null;
          }, null);
          if (dt) {
            return new Date(dt.valueOf());
          }
          const special = Object.hasOwn(SpecialValues, datum) ? SpecialValues[datum] : undefined;
          if (special !== undefined) {
            return special;
          }
          return datum;
        };

      // Parsed as it arrives, chunk by chunk, as fast-csv does.
      let parser;
      let failed = false;
      const cleanup = () => {
        stream.removeListener('data', onData);
        stream.removeListener('end', onEnd);
        stream.removeListener('error', fail);
      };
      const fail = error => {
        failed = true;
        cleanup();
        reject(error);
      };
      try {
        parser = fastCsv.createParser(options.parserOptions, data => {
          worksheet.addRow(data.map(map));
        });
      } catch (error) {
        reject(error);
        return;
      }
      const onData = chunk => {
        if (failed) return;
        try {
          parser.write(chunk);
        } catch (error) {
          fail(error);
        }
      };
      const onEnd = () => {
        if (failed) return;
        cleanup();
        try {
          parser.end();
          resolve(worksheet);
        } catch (error) {
          reject(error);
        }
      };
      stream.on('error', fail);
      stream.on('data', onData);
      stream.on('end', onEnd);
    });
  }

  /**
   * @deprecated since version 4.0. You should use `CSV#read` instead. Please follow upgrade instruction: https://github.com/exceljs/exceljs/blob/master/UPGRADE-4.0.md
   */
  createInputStream() {
    throw new Error(
      '`CSV#createInputStream` is deprecated. You should use `CSV#read` instead. This method will be removed in version 5.0. Please follow upgrade instruction: https://github.com/exceljs/exceljs/blob/master/UPGRADE-4.0.md'
    );
  }

  write(stream, options) {
    return new Promise((resolve, reject) => {
      options = options || {};
      // const encoding = options.encoding || 'utf8';
      // const separator = options.separator || ',';
      // const quoteChar = options.quoteChar || '\'';

      const worksheet = this.workbook.getWorksheet(options.sheetName || options.sheetId);

      let csvStream;
      try {
        csvStream = new fastCsv.Formatter(options.formatterOptions);
      } catch (error) {
        reject(error);
        return;
      }
      const onError = error => {
        stream.removeListener('finish', onFinish);
        reject(error);
      };
      const onFinish = () => {
        stream.removeListener('error', onError);
        resolve();
      };
      stream.once('error', onError);
      stream.once('finish', onFinish);

      const {dateFormat, dateUTC} = options;
      const map =
        options.map ||
        (value => {
          if (value) {
            if (value.text || value.hyperlink) {
              return value.hyperlink || value.text || '';
            }
            if (value.formula || value.result) {
              return value.result || '';
            }
            if (value instanceof Date) {
              if (dateFormat) {
                return dayjs.format(value, dateFormat, !!dateUTC);
              }
              return dayjs.format(value, undefined, !!dateUTC);
            }
            if (value.error) {
              return value.error;
            }
            if (typeof value === 'object') {
              return JSON.stringify(value);
            }
          }
          return value;
        });

      const includeEmptyRows = options.includeEmptyRows === undefined || options.includeEmptyRows;
      let lastRow = 1;
      let output;
      try {
        if (worksheet) {
          worksheet.eachRow((row, rowNumber) => {
            if (includeEmptyRows) {
              while (lastRow++ < rowNumber - 1) {
                csvStream.write([]);
              }
            }
            const {values} = row;
            values.shift();
            csvStream.write(values.map(map));
            lastRow = rowNumber;
            // (excel-rs) passed on in pieces, not held whole
            if (csvStream.size >= 1 << 20) {
              stream.write(csvStream.take());
            }
          });
        }
        output = csvStream.end();
      } catch (error) {
        stream.removeListener('error', onError);
        stream.removeListener('finish', onFinish);
        reject(error);
        return;
      }
      if (output.length) {
        stream.write(output);
      }
      stream.end();
    });
  }

  writeFile(filename, options) {
    options = options || {};

    const streamOptions = {
      encoding: options.encoding || 'utf8',
    };
    const stream = fs.createWriteStream(filename, streamOptions);

    return this.write(stream, options);
  }

  async writeBuffer(options) {
    const stream = new StreamBuf();
    await this.write(stream, options);
    return stream.read();
  }
}

module.exports = CSV;
