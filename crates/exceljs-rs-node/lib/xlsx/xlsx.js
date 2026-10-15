'use strict';

// workbook.xlsx: the same methods as ExcelJS lib/xlsx/xlsx.js, with all
// .xlsx parsing and serialisation done by the Rust engine. The boundary is
// ExcelJS's workbook model: the reader returns the model ExcelJS's reader
// gives to `workbook.model = ...`; the writer takes `workbook.model`.

const fs = require('fs');
const native = require('../native');
const codec = require('./codec');
const {encodeWorkbook} = require('./rows');
const {ReadControls, readInput} = require('./read-controls');

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
    const controls = new ReadControls(options);
    try {
      if (!(await controls.wait(fs.promises.access(filename).then(() => true, () => false)))) {
        throw new Error(`File not found: ${filename}`);
      }
      const buffer = controls.enabled ? await readInput(filename, controls) : await fs.promises.readFile(filename);
      return await this._loadBuffer(buffer, controls);
    } finally { controls.close(); }
  }

  async read(stream, options) {
    const controls = new ReadControls(options);
    try {
      let buffer;
      if (controls.enabled) buffer = await readInput(stream, controls);
      else {
        const chunks = [];
        for await (const chunk of stream) chunks.push(typeof chunk === 'string' ? Buffer.from(chunk) : chunk);
        buffer = Buffer.concat(chunks);
      }
      return await this._loadBuffer(buffer, controls);
    } finally { controls.close(); }
  }

  async load(data, options) {
    const controls = new ReadControls(options);
    try {
      const resolved = await controls.wait(Promise.resolve(data));
      const buffer = options && options.base64 ? Buffer.from(resolved.toString(), 'base64') : toBuffer(resolved);
      if (controls.enabled) controls.input(buffer.length);
      return await this._loadBuffer(buffer, controls);
    } finally { controls.close(); }
  }

  async _loadBuffer(buffer, controls) {
    // The engine reads on a thread of its own and sends the rows in
    // batches, which become Rows and Cells while it reads the next ones.
    const reader = new native.XlsxReader(buffer, codec.encode(controls.nativeOptions()));
    controls.attach(reader);
    try {
      const head = await controls.wait(reader.model());
      if (controls.enabled) {
        await this.workbook._loadModelAsync(codec.decodeAsync(head, reader, controls), controls);
        controls.finish();
      } else this.workbook.model = codec.decodeStreamed(head, reader);
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
