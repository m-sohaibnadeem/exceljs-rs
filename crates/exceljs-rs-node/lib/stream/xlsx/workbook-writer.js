// From ExcelJS 4.4.0 lib/stream/xlsx/workbook-writer.js (MIT, see LICENSE-EXCELJS).
// The archive, styles, shared strings and XML rendering live in the Rust
// engine (native StreamWorkbook); this keeps ExcelJS's object model and
// commit protocol.
'use strict';

const fs = require('fs');
const EventEmitter = require('events');

const native = require('../../native');
const codec = require('../../xlsx/codec');
const {StyleTable} = require('../../xlsx/rows');
const StreamBuf = require('../../utils/stream-buf');
const DefinedNames = require('../../doc/defined-names');

const WorksheetWriter = require('./worksheet-writer');
const {ReadControls} = require('../../xlsx/read-controls');

const RelType = {
  OfficeDocument: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
  Worksheet: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
  Styles: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
  SharedStrings: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
  Theme: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
  CoreProperties: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
  ExtenderProperties: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
};

const theme1Xml = fs.readFileSync(require.resolve('../../xlsx/theme1.xml'), 'utf8');

// One zip entry, written through the engine (what `_openStream` returns).
class EntryStream extends EventEmitter {
  constructor(workbook, name) {
    super();
    this._workbook = workbook;
    this.name = name;
    workbook._engine.open(name);
  }

  pause() {}

  write(data) {
    const engine = this._workbook._engine;
    if (Buffer.isBuffer(data)) {
      engine.writeBytes(this.name, data);
    } else if (data && typeof data.toBuffer === 'function') {
      engine.writeBytes(this.name, data.toBuffer());
    } else {
      engine.write(this.name, String(data));
    }
    this._workbook._flush();
    return true;
  }

  end() {
    this._workbook._engine.end(this.name);
    this._workbook._flush();
    process.nextTick(() => {
      this.emit('finish');
      this.emit('zipped');
    });
  }
}

class WorkbookWriter {
  constructor(options) {
    options = options || {};

    this.created = options.created || new Date();
    this.modified = options.modified || this.created;
    this.creator = options.creator || 'ExcelJS';
    this.lastModifiedBy = options.lastModifiedBy || 'ExcelJS';
    this.lastPrinted = options.lastPrinted;

    // using shared strings creates a smaller xlsx file but may use more memory
    this.useSharedStrings = options.useSharedStrings || false;
    const sharedStrings = {
      get count() {
        return engine.sharedStringsCount();
      },
    };
    this.sharedStrings = sharedStrings;

    // style manager
    const zipOptions = options.zip;
    const level = zipOptions && zipOptions.zlib && zipOptions.zlib.level;
    const engine = new native.StreamWorkbook(!!options.useStyles, typeof level === 'number' ? level : null);
    this._engine = engine;
    this._styleTable = new StyleTable();
    this.styles = {
      get xml() {
        return engine.render('styles');
      },
    };

    // defined names
    this._definedNames = new DefinedNames();

    this._worksheets = [];
    this.views = [];

    this.zipOptions = options.zip;

    this.media = [];
    this.commentRefs = [];

    if (options.stream) {
      this.stream = options.stream;
    } else if (options.filename) {
      this.stream = fs.createWriteStream(options.filename);
      this._ownsFile = true;
    } else {
      this.stream = new StreamBuf();
    }

    this._needsDrain = false;
    this._destinationFailure = new Promise((resolve, reject) => { this._rejectDestination = reject; });
    this._destinationFailure.catch(() => {});
    this._onDestinationError = error => this._failDestination(error);
    this._onDestinationFinish = () => {
      this._destinationFinished = true;
      if (this.stream instanceof StreamBuf || (this.stream._writableState && !this.stream._writableState.autoDestroy)) {
        this._releaseDestinationListeners();
      }
    };
    this._onDestinationClose = () => {
      if (!this._destinationFinished) this._failDestination(new Error('XLSX destination closed before finish'));
      this._releaseDestinationListeners();
    };
    this.stream.on('error', this._onDestinationError);
    this.stream.on('finish', this._onDestinationFinish);
    this.stream.on('close', this._onDestinationClose);
    this._onDestinationDrain = () => { this._needsDrain = false; };
    this.stream.on('drain', this._onDestinationDrain);

    // these bits can be added right now
    this.promise = Promise.all([this.addThemes(), this.addOfficeRels()]);
    this.promise.catch(() => {});
  }

  // Moves the archive bytes produced so far to the output stream.
  _flush() {
    if (this._destinationFailed) return;
    const out = this._engine.takeOutput();
    if (out.length) {
      try {
        if (this.stream.write(out, error => { if (error) this._failDestination(error); }) === false) this._needsDrain = true;
      } catch (error) { this._failDestination(error); }
    }
  }

  _failDestination(error) {
    if (this._destinationFailed) return;
    this._destinationFailed = true;
    this._destinationError = error;
    this._rejectDestination(error);
    try { this._engine.abort(); } catch (_) { /* Keep the original destination/source error. */ }
    if (this._ownsFile && !this.stream.destroyed) this.stream.destroy();
  }

  _checkDestination() {
    if (this._destinationFailed) throw this._destinationError;
    if (this._destinationFinished || this.stream.destroyed || this.stream.writableEnded) {
      const error = new Error('XLSX destination is already closed or ended');
      this._failDestination(error);
      throw error;
    }
  }

  async _waitDrain(controls) {
    this._checkDestination();
    if (!this._needsDrain && !this.stream.writableNeedDrain) return;
    let onDrain;
    const drained = new Promise(resolve => { onDrain = resolve; this.stream.once('drain', onDrain); });
    try {
      await controls.wait(Promise.race([drained, this._destinationFailure]));
      this._checkDestination();
    } finally { this.stream.removeListener('drain', onDrain); }
  }

  _releaseDestinationListeners() {
    this.stream.removeListener('error', this._onDestinationError);
    this.stream.removeListener('finish', this._onDestinationFinish);
    this.stream.removeListener('close', this._onDestinationClose);
    this.stream.removeListener('drain', this._onDestinationDrain);
  }

  async _bulkCheckpoint(worksheet, controls) {
    this._checkDestination();
    controls.check();
    worksheet._flushRows();
    this._engine.checkpoint();
    this._flush();
    await this._waitDrain(controls);
    // Fast destinations and synchronous sources otherwise form only a
    // microtask chain, preventing timers and I/O from delivering aborts.
    await controls.wait(Promise.race([new Promise(resolve => setImmediate(resolve)), this._destinationFailure]));
    this._checkDestination();
  }

  async writeRows(worksheet, rows, options = {}) {
    if (!worksheet || worksheet._workbook !== this || this._worksheets[worksheet.id] !== worksheet) {
      throw new TypeError('worksheet must belong to this WorkbookWriter');
    }
    if (worksheet.committed) throw new Error('worksheet is already committed');
    if (this._bulkActive || this._committing) throw new Error('A bulk export or workbook commit is already in progress');
    const batchSize = options.batchSize === undefined ? 128 : options.batchSize;
    if (!Number.isSafeInteger(batchSize) || batchSize <= 0) throw new RangeError('batchSize must be a positive safe integer');
    if (options.onProgress !== undefined && typeof options.onProgress !== 'function') throw new TypeError('onProgress must be a function');
    const factory = rows && (rows[Symbol.asyncIterator] || rows[Symbol.iterator]);
    if (typeof factory !== 'function') throw new TypeError('rows must be a synchronous or asynchronous iterable');
    this._checkDestination();
    if (!this._engine.canWriteRows(worksheet.stream.name)) {
      throw new Error('Commit earlier worksheets before bulk exporting this worksheet');
    }
    const controls = new ReadControls({signal: options.signal});
    this._bulkActive = true;
    let iterator;
    let count = 0;
    try {
      iterator = factory.call(rows);
      await controls.wait(Promise.race([this.promise, this._destinationFailure]));
      await this._bulkCheckpoint(worksheet, controls);
      for (;;) {
        await this._waitDrain(controls);
        if (worksheet.committed) throw new Error('worksheet was committed during bulk export');
        controls.check();
        const item = await controls.wait(Promise.race([Promise.resolve().then(() => iterator.next()), this._destinationFailure]));
        if (item.done) break;
        worksheet.addRow(item.value).commit();
        count++;
        if (count % batchSize === 0) {
          await this._bulkCheckpoint(worksheet, controls);
          if (options.onProgress) options.onProgress({phase: 'rows', rows: count});
          controls.check();
        }
      }
      await this._bulkCheckpoint(worksheet, controls);
      if (options.onProgress) options.onProgress({phase: 'done', rows: count});
      controls.check();
      return count;
    } catch (error) {
      this._failDestination(error);
      if (iterator && iterator.return) {
        // A source with an unresolved next() may also block return(). Invoke
        // cleanup without making rejection wait for an uncooperative source.
        try { Promise.resolve(iterator.return()).catch(() => {}); } catch (_) {}
      }
      throw error;
    } finally {
      controls.close();
      this._bulkActive = false;
    }
  }

  _render(kind, model) {
    return this._engine.render(kind, model === undefined ? undefined : codec.encode(model));
  }

  _append(name, data) {
    if (Buffer.isBuffer(data)) {
      this._engine.appendBytes(name, data);
    } else {
      this._engine.append(name, data);
    }
    this._flush();
  }

  get definedNames() {
    return this._definedNames;
  }

  _openStream(path) {
    return new EntryStream(this, path);
  }

  _commitWorksheets() {
    const commitWorksheet = function(worksheet) {
      if (!worksheet.committed) {
        return new Promise(resolve => {
          worksheet.stream.on('zipped', () => {
            resolve();
          });
          worksheet.commit();
        });
      }
      return Promise.resolve();
    };
    // if there are any uncommitted worksheets, commit them now and wait
    const promises = this._worksheets.map(commitWorksheet);
    if (promises.length) {
      return Promise.all(promises);
    }
    return Promise.resolve();
  }

  async commit() {
    if (this._bulkActive) throw new Error('A bulk export is still in progress');
    if (this._committing) return this._committing;
    this._committing = this._commit();
    return this._committing;
  }

  async _commit() {
    this._checkDestination();
    // commit all worksheets, then add suplimentary files
    await this.promise;
    this._checkDestination();
    await this.addMedia();
    await this._commitWorksheets();
    await Promise.all([
      this.addContentTypes(),
      this.addApp(),
      this.addCore(),
      this.addSharedStrings(),
      this.addStyles(),
      this.addWorkbookRels(),
    ]);
    await this.addWorkbook();
    return this._finalize();
  }

  get nextId() {
    // find the next unique spot to add worksheet
    let i;
    for (i = 1; i < this._worksheets.length; i++) {
      if (!this._worksheets[i]) {
        return i;
      }
    }
    return this._worksheets.length || 1;
  }

  addImage(image) {
    const id = this.media.length;
    const medium = Object.assign({}, image, {type: 'image', name: `image${id}.${image.extension}`});
    this.media.push(medium);
    return id;
  }

  getImage(id) {
    return this.media[id];
  }

  addWorksheet(name, options) {
    // it's possible to add a worksheet with different than default
    // shared string handling
    // in fact, it's even possible to switch it mid-sheet
    options = options || {};
    const useSharedStrings =
      options.useSharedStrings !== undefined ? options.useSharedStrings : this.useSharedStrings;

    if (options.tabColor) {
      // eslint-disable-next-line no-console
      console.trace('tabColor option has moved to { properties: tabColor: {...} }');
      options.properties = Object.assign(
        {
          tabColor: options.tabColor,
        },
        options.properties
      );
    }

    const id = this.nextId;
    name = name || `sheet${id}`;

    const worksheet = new WorksheetWriter({
      id,
      name,
      workbook: this,
      useSharedStrings,
      properties: options.properties,
      state: options.state,
      pageSetup: options.pageSetup,
      views: options.views,
      autoFilter: options.autoFilter,
      headerFooter: options.headerFooter,
    });

    this._worksheets[id] = worksheet;
    return worksheet;
  }

  getWorksheet(id) {
    if (id === undefined) {
      return this._worksheets.find(() => true);
    }
    if (typeof id === 'number') {
      return this._worksheets[id];
    }
    if (typeof id === 'string') {
      return this._worksheets.find(worksheet => worksheet && worksheet.name === id);
    }
    return undefined;
  }

  addStyles() {
    return new Promise(resolve => {
      this._append('xl/styles.xml', this.styles.xml);
      resolve();
    });
  }

  addThemes() {
    return new Promise(resolve => {
      this._append('xl/theme/theme1.xml', theme1Xml);
      resolve();
    });
  }

  addOfficeRels() {
    return new Promise(resolve => {
      const xml = this._render('relationships', [
        {Id: 'rId1', Type: RelType.OfficeDocument, Target: 'xl/workbook.xml'},
        {Id: 'rId2', Type: RelType.CoreProperties, Target: 'docProps/core.xml'},
        {Id: 'rId3', Type: RelType.ExtenderProperties, Target: 'docProps/app.xml'},
      ]);
      this._append('/_rels/.rels', xml);
      resolve();
    });
  }

  _sheetModels() {
    return this._worksheets
      .filter(Boolean)
      .map(ws => ({id: ws.id, name: ws.name, state: ws.state, rId: ws.rId, pageSetup: ws.pageSetup}));
  }

  addContentTypes() {
    return new Promise(resolve => {
      const model = {
        worksheets: this._sheetModels(),
        commentRefs: this.commentRefs,
        media: this.media.map(m => ({type: m.type, name: m.name, extension: m.extension})),
      };
      this._append('[Content_Types].xml', this._render('contentTypes', model));
      resolve();
    });
  }

  addMedia() {
    return Promise.all(
      this.media.map(async medium => {
        if (medium.type === 'image') {
          const filename = `xl/media/${medium.name}`;
          if (medium.filename) {
            this._append(filename, await fs.promises.readFile(medium.filename));
            return;
          }
          if (medium.buffer) {
            this._append(filename, Buffer.from(medium.buffer));
            return;
          }
          if (medium.base64) {
            const dataimg64 = medium.base64;
            const content = dataimg64.substring(dataimg64.indexOf(',') + 1);
            this._append(filename, Buffer.from(content, 'base64'));
            return;
          }
        }
        throw new Error('Unsupported media');
      })
    );
  }

  addApp() {
    return new Promise(resolve => {
      this._append('docProps/app.xml', this._render('app', {worksheets: this._sheetModels()}));
      resolve();
    });
  }

  addCore() {
    return new Promise(resolve => {
      const model = {
        creator: this.creator,
        lastModifiedBy: this.lastModifiedBy,
        lastPrinted: this.lastPrinted,
        created: this.created,
        modified: this.modified,
        title: this.title,
        subject: this.subject,
        keywords: this.keywords,
        category: this.category,
        description: this.description,
        language: this.language,
        revision: this.revision,
        contentStatus: this.contentStatus,
        identifier: this.identifier,
      };
      this._append('docProps/core.xml', this._render('core', model));
      resolve();
    });
  }

  addSharedStrings() {
    if (this.sharedStrings.count) {
      return new Promise(resolve => {
        this._append('/xl/sharedStrings.xml', this._render('sharedStrings'));
        resolve();
      });
    }
    return Promise.resolve();
  }

  addWorkbookRels() {
    let count = 1;
    const relationships = [
      {Id: `rId${count++}`, Type: RelType.Styles, Target: 'styles.xml'},
      {Id: `rId${count++}`, Type: RelType.Theme, Target: 'theme/theme1.xml'},
    ];
    if (this.sharedStrings.count) {
      relationships.push({
        Id: `rId${count++}`,
        Type: RelType.SharedStrings,
        Target: 'sharedStrings.xml',
      });
    }
    this._worksheets.forEach(worksheet => {
      if (worksheet) {
        worksheet.rId = `rId${count++}`;
        relationships.push({
          Id: worksheet.rId,
          Type: RelType.Worksheet,
          Target: `worksheets/sheet${worksheet.id}.xml`,
        });
      }
    });
    return new Promise(resolve => {
      this._append('/xl/_rels/workbook.xml.rels', this._render('relationships', relationships));
      resolve();
    });
  }

  addWorkbook() {
    const model = {
      worksheets: this._sheetModels(),
      definedNames: this._definedNames.model,
      views: this.views,
      properties: {},
      calcProperties: {},
    };

    return new Promise(resolve => {
      this._append('/xl/workbook.xml', this._render('workbook', model));
      resolve();
    });
  }

  _finalize() {
    return new Promise((resolve, reject) => {
      const cleanup = () => {
        this.stream.removeListener('error', onError);
        this.stream.removeListener('finish', onFinish);
        this.stream.removeListener('close', onClose);
      };
      const onError = error => { cleanup(); reject(error); };
      const onFinish = () => { cleanup(); resolve(this); };
      const onClose = () => { if (!this._destinationFinished) onError(this._destinationError || new Error('XLSX destination closed before finish')); };
      this.stream.once('error', onError);
      this.stream.once('finish', onFinish);
      this.stream.once('close', onClose);
      try {
        this._checkDestination();
        this._engine.finalize();
        this._flush();
        if (this._destinationFailed) throw this._destinationError;
        this.stream.end();
      } catch (error) {
        this._failDestination(error);
        onError(error);
        return;
      }
    });
  }
}

module.exports = WorkbookWriter;
