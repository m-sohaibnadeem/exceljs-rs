'use strict';

const fs = require('fs');

const LIMITS = new Set(['maxInputBytes', 'maxZipEntries', 'maxEntryBytes', 'maxInflatedBytes',
  'maxCells', 'maxSharedStrings', 'maxTokenBytes', 'maxValidationCells', 'maxMergeCells']);

function abortReason(signal) {
  return signal.reason === undefined ? new DOMException('The operation was aborted', 'AbortError') : signal.reason;
}

class ReadControls {
  constructor(options = {}) {
    options = options || {};
    this.options = options;
    this.enabled = options.limits !== undefined || options.signal !== undefined || options.onProgress !== undefined;
    if (options.limits !== undefined) {
      if (!options.limits || typeof options.limits !== 'object' || Array.isArray(options.limits)) {
        throw new TypeError('limits must be an object');
      }
      for (const [name, value] of Object.entries(options.limits)) {
        if (!LIMITS.has(name)) throw new TypeError(`Unknown XLSX limit: ${name}`);
        if (!Number.isSafeInteger(value) || value < 0) throw new RangeError(`${name} must be a nonnegative safe integer`);
      }
    }
    const {signal, onProgress} = options;
    if (signal !== undefined && (!signal || typeof signal.addEventListener !== 'function' || typeof signal.removeEventListener !== 'function' || typeof signal.aborted !== 'boolean')) {
      throw new TypeError('signal must be an AbortSignal');
    }
    if (onProgress !== undefined && typeof onProgress !== 'function') throw new TypeError('onProgress must be a function');
    if (signal && signal.aborted) throw abortReason(signal);
    this.failed = false;
    this.error = undefined;
    this.native = null;
    this.counters = {inputBytes: 0, inflatedBytes: 0, entries: 0, cells: 0, sharedStrings: 0, validationCells: 0, mergeCells: 0, rows: 0};
    this.failure = new Promise((resolve, reject) => { this.reject = reject; });
    this.failure.catch(() => {});
    this.onAbort = () => this.fail(abortReason(signal));
    if (signal) signal.addEventListener('abort', this.onAbort, {once: true});
  }

  fail(error) {
    if (this.failed) return;
    this.failed = true;
    this.error = error;
    if (this.native) this.native.close();
    this.reject(error);
  }

  check() {
    if (this.failed) throw this.error;
    if (this.options.signal && this.options.signal.aborted) {
      this.fail(abortReason(this.options.signal));
      throw this.error;
    }
  }

  async wait(promise) {
    // The caller may already have started work. Observe its rejection even
    // when cancellation wins before this wait is entered.
    const pending = Promise.race([promise, this.failure]);
    pending.catch(() => {});
    this.check();
    const result = await pending;
    this.check();
    return result;
  }

  input(n) {
    this.check();
    this.counters.inputBytes += n;
    const max = this.options.limits && this.options.limits.maxInputBytes;
    if (max !== undefined && this.counters.inputBytes > max) throw new Error('XLSX limit exceeded: maxInputBytes');
    this.notify('input');
  }

  attach(reader) {
    this.check();
    this.native = reader;
    if (this.options.onProgress) {
      this.timer = setInterval(() => {
        try { this.notify(this.phase || 'archive'); } catch (error) { this.fail(error); }
      }, 16);
    }
  }

  notify(phase) {
    this.phase = phase;
    this.check();
    if (this.native && this.native.progress) {
      const counters = JSON.parse(this.native.progress());
      for (const [key, value] of Object.entries(counters)) this.counters[key] = Math.max(this.counters[key], value);
    }
    if (this.options.onProgress) this.options.onProgress({phase, ...this.counters});
    this.check();
  }

  finish() { this.notify('done'); }

  close() {
    clearInterval(this.timer);
    if (this.options.signal) this.options.signal.removeEventListener('abort', this.onAbort);
    if (this.native) { this.native.close(); this.native = null; }
  }

  nativeOptions() {
    const {signal, onProgress, batchSize, ...options} = this.options;
    return {...options, _controlled: this.enabled};
  }
}

function toBuffer(data) {
  if (Buffer.isBuffer(data)) return data;
  if (data instanceof ArrayBuffer) return Buffer.from(data);
  if (ArrayBuffer.isView(data)) return Buffer.from(data.buffer, data.byteOffset, data.byteLength);
  if (typeof data === 'string') return Buffer.from(data, 'binary');
  if (Array.isArray(data)) return Buffer.from(data);
  throw new Error("Can't read the data of 'the loaded zip file'. Is it in a supported JavaScript type (String, Blob, ArrayBuffer, etc) ?");
}

async function readInput(input, controls) {
  controls.check();
  if (Buffer.isBuffer(input) || input instanceof ArrayBuffer || ArrayBuffer.isView(input)) {
    const buffer = toBuffer(input);
    controls.input(buffer.length);
    return buffer;
  }
  const owned = typeof input === 'string';
  const stream = owned ? fs.createReadStream(input) : input;
  if (!stream || typeof stream.read !== 'function' || typeof stream.on !== 'function') throw new TypeError('input must be a filename, readable, Buffer, ArrayBuffer or typed array');
  // Explicit pull listeners can be removed even when the source never
  // supplies a chunk. Node's async iterator queues return() behind next().
  const chunks = [];
  let closing = false;
  const onError = error => { if (!closing) controls.fail(error); };
  stream.on('error', onError);
  try {
    for (;;) {
      controls.check();
      const value = stream.read();
      if (value === null) {
        if (stream.readableEnded) break;
        if (stream.destroyed) throw new Error('XLSX input closed before end');
        let onReadable;
        let onEnd;
        let onClose;
        const ready = new Promise((resolve, reject) => {
          onReadable = resolve;
          onEnd = resolve;
          onClose = () => stream.readableEnded ? resolve() : reject(new Error('XLSX input closed before end'));
          stream.once('readable', onReadable);
          stream.once('end', onEnd);
          stream.once('close', onClose);
        });
        try { await controls.wait(ready); }
        finally {
          stream.removeListener('readable', onReadable);
          stream.removeListener('end', onEnd);
          stream.removeListener('close', onClose);
        }
        continue;
      }
      const buffer = typeof value === 'string' ? Buffer.from(value) : toBuffer(value);
      controls.input(buffer.length);
      chunks.push(buffer);
    }
    return Buffer.concat(chunks);
  } finally {
    if (owned) {
      closing = true;
      // An asynchronous open error may arrive after abort/destroy. Keep
      // only our error absorber until the file stream's terminal close.
      if (stream.closed) stream.removeListener('error', onError);
      else stream.once('close', () => stream.removeListener('error', onError));
      stream.destroy();
    } else stream.removeListener('error', onError);
  }
}

module.exports = {ReadControls, readInput, toBuffer, abortReason};
