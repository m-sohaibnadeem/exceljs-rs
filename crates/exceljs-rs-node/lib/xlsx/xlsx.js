'use strict';

// workbook.xlsx: the same methods as ExcelJS lib/xlsx/xlsx.js, with all
// .xlsx parsing and serialisation done by the Rust engine. The boundary is
// ExcelJS's workbook model: the reader returns the model ExcelJS's reader
// gives to `workbook.model = ...`; the writer takes `workbook.model`.

const fs = require('fs');
const native = require('../native');
const codec = require('./codec');
const {encodeWorkbook} = require('./rows');

// The input types JSZip.loadAsync accepts, with JSZip's error otherwise.
function toBuffer(data) {
  if (Buffer.isBuffer(data)) return data;
  if (data instanceof ArrayBuffer) return Buffer.from(data);
  if (ArrayBuffer.isView(data)) return Buffer.from(data.buffer, data.byteOffset, data.byteLength);
  if (typeof data === 'string') return Buffer.from(data, 'binary');
  if (Array.isArray(data)) return Buffer.from(data);
  throw new Error(
    "Can't read the data of 'the loaded zip file'. Is it in a supported JavaScript type (String, Blob, ArrayBuffer, etc) ?"
  );
}

class XLSX {
  constructor(workbook) {
    this.workbook = workbook;
  }

  async readFile(filename, options) {
    if (!(await fs.promises.access(filename).then(() => true, () => false))) {
      throw new Error(`File not found: ${filename}`);
    }
    return this.load(await fs.promises.readFile(filename), options);
  }

  async read(stream, options) {
    const chunks = [];
    for await (const chunk of stream) {
      chunks.push(typeof chunk === 'string' ? Buffer.from(chunk) : chunk);
    }
    return this.load(Buffer.concat(chunks), options);
  }

  async load(data, options) {
    let buffer;
    if (options && options.base64) {
      buffer = Buffer.from(data.toString(), 'base64');
    } else {
      buffer = toBuffer(await data);
    }
    // The engine reads on a thread of its own and sends the rows in
    // batches, which become Rows and Cells while it reads the next ones.
    const reader = new native.XlsxReader(buffer, codec.encode(options || {}));
    try {
      const head = await reader.model();
      this.workbook.model = codec.decodeStreamed(head, reader);
    } finally {
      reader.close();
    }
    return this.workbook;
  }

  async writeBuffer(options) {
    const {model, rows, styles} = encodeWorkbook(this.workbook._modelForWrite());
    return native.writeXlsx(model, rows, styles, codec.encode(options || {}));
  }

  async write(stream, options) {
    const buffer = await this.writeBuffer(options);
    await new Promise((resolve, reject) => {
      // Writable errors can be emitted as well as passed to a callback.
      stream.once('error', reject);
      stream.write(buffer, error => {
        if (error) {
          reject(error);
        } else if (typeof stream.end === 'function') {
          stream.end(error => {
            if (error) reject(error);
            else {
              stream.removeListener('error', reject);
              resolve();
            }
          });
        } else {
          stream.removeListener('error', reject);
          resolve();
        }
      });
    });
    return this;
  }

  writeFile(filename, options) {
    return this.writeBuffer(options).then(buffer => fs.promises.writeFile(filename, buffer));
  }
}

module.exports = XLSX;
