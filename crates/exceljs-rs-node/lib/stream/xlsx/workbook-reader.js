// From ExcelJS 4.4.0 lib/stream/xlsx/workbook-reader.js (MIT, see LICENSE-EXCELJS).
// The archive and the SAX parsing are in the Rust engine (native
// StreamReader); entry order, deferred worksheets and events follow ExcelJS.
// Difference: the input stream is read to the end before parsing starts
// (COMPATIBILITY.md, D2).
'use strict';

const fs = require('fs');
const {EventEmitter} = require('events');
const nodeStream = require('stream');

const native = require('../../native');
const codec = require('../../xlsx/codec');
const {ReadControls, readInput} = require('../../xlsx/read-controls');

const WorksheetReader = require('./worksheet-reader');
const HyperlinkReader = require('./hyperlink-reader');

// styles: 'ignore': no style is known. (ExcelJS's StylesXform after init()
// throws for id 0, so a <row s="0"> fails: B13 in COMPATIBILITY.md.)
const initStyleManager = {
  getStyleModel() {
    return null;
  },
};

// StylesXform after parsing styles.xml (styles: 'cache').
function cachedStyleManager(json) {
  const [refs, model] = JSON.parse(json);
  const models = codec.decodeRead({refs, model});
  return {
    getStyleModel(id) {
      const model = models[id];
      return model === undefined ? null : model;
    },
  };
}

async function readAll(stream) {
  const chunks = [];
  for await (const chunk of stream) {
    chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
  }
  return Buffer.concat(chunks);
}

class WorkbookReader extends EventEmitter {
  constructor(input, options = {}) {
    super();

    this.input = input;

    this.options = {
      worksheets: 'emit',
      sharedStrings: 'cache',
      hyperlinks: 'ignore',
      styles: 'ignore',
      entries: 'ignore',
      ...options,
    };

    this.styles = initStyleManager;
  }

  _getStream(input) {
    if (input instanceof nodeStream.Readable || (input && input._readableState && typeof input.pipe === 'function')) {
      return input;
    }
    if (typeof input === 'string') {
      return fs.createReadStream(input);
    }
    throw new Error(`Could not recognise input: ${input}`);
  }

  async read(input, options) {
    try {
      for await (const {eventType, value} of this.parse(input, options)) {
        switch (eventType) {
          case 'shared-strings':
            this.emit(eventType, value);
            break;
          case 'worksheet':
            this.emit(eventType, value);
            await value.read();
            break;
          case 'hyperlinks':
            this.emit(eventType, value);
            break;
        }
      }
      this.emit('end');
      this.emit('finished');
    } catch (error) {
      this.emit('error', error);
    }
  }

  async *[Symbol.asyncIterator]() {
    for await (const {eventType, value} of this.parse()) {
      if (eventType === 'worksheet') {
        yield value;
      }
    }
  }

  async *parse(input, options) {
    if (options) this.options = options;
    const controls = new ReadControls(this.options);
    this._controls = controls;
    const source = input || this.input;
    let stream;
    try {
      const buffer = controls.enabled ? await readInput(source, controls) : await readAll(stream = (this.stream = this._getStream(source)));
      const archive = new native.StreamReader(buffer, codec.encode(controls.nativeOptions()));
      controls.attach(archive);
      if (controls.enabled) await controls.wait(archive.validate());
      this._archive = archive;

      // worksheets, deferred for parsing after shared strings reading
      const waitingWorkSheets = [];

      const names = archive.entryNames();
      for (let index = 0; index < names.length; index++) {
        const path = names[index];
        const entry = {index, path};
        let match;
        let sheetNo;
        switch (path) {
          case '_rels/.rels':
            break;
          case 'xl/_rels/workbook.xml.rels':
            await this._parseRels(entry);
            break;
          case 'xl/workbook.xml':
            await this._parseWorkbook(entry);
            break;
          case 'xl/sharedStrings.xml':
            yield* this._parseSharedStrings(entry);
            break;
          case 'xl/styles.xml':
            await this._parseStyles(entry);
            break;
          default:
            if (path.match(/xl\/worksheets\/sheet\d+[.]xml/)) {
              match = path.match(/xl\/worksheets\/sheet(\d+)[.]xml/);
              sheetNo = match[1];
              if (this.sharedStrings && this.workbookRels) {
                yield* this._parseWorksheet(entry, sheetNo);
              } else {
                waitingWorkSheets.push({sheetNo, entry});
              }
            } else if (path.match(/xl\/worksheets\/_rels\/sheet\d+[.]xml.rels/)) {
              match = path.match(/xl\/worksheets\/_rels\/sheet(\d+)[.]xml.rels/);
              sheetNo = match[1];
              yield* this._parseHyperlinks(entry, sheetNo);
            }
            break;
        }
      }

      for (const {sheetNo, entry} of waitingWorkSheets) {
        yield* this._parseWorksheet(entry, sheetNo);
      }
      if (controls.enabled) controls.finish();
    } finally {
      // Legacy readers keep their archive for worksheet objects consumed
      // after outer iteration. Controlled readers scope it to this parse.
      if (!controls.enabled) controls.native = null;
      controls.close();
      if (controls.enabled) this._archive = null;
      if (stream && typeof source === 'string') stream.destroy();
    }
  }

  _readEntry(entry, kind) {
    return this._controls.wait(this._archive.readEntry(entry.index, entry.path, kind));
  }

  _emitEntry(payload) {
    if (this.options.entries === 'emit') {
      this.emit('entry', payload);
    }
  }

  async _parseRels(entry) {
    this.workbookRels = codec.decode(await this._readEntry(entry, 'rels'));
  }

  async _parseWorkbook(entry) {
    this._emitEntry({type: 'workbook'});

    const {model, properties} = codec.decode(await this._readEntry(entry, 'workbook'));
    this.properties = {model: properties};
    this.model = model;
  }

  async *_parseSharedStrings(entry) {
    this._emitEntry({type: 'shared-strings'});
    switch (this.options.sharedStrings) {
      case 'cache':
      case 'emit':
        break;
      default:
        return;
    }
    const strings = codec.decode(await this._readEntry(entry, 'sharedStrings'));
    if (this.options.sharedStrings === 'cache') {
      this.sharedStrings = strings;
    } else {
      for (let index = 0; index < strings.length; index++) {
        yield {eventType: 'shared-strings', value: {index, text: strings[index]}};
      }
    }
  }

  async _parseStyles(entry) {
    this._emitEntry({type: 'styles'});
    if (this.options.styles === 'cache') {
      this.styles = cachedStyleManager(await this._readEntry(entry, 'styles'));
    }
  }

  *_parseWorksheet(entry, sheetNo) {
    this._emitEntry({type: 'worksheet', id: sheetNo});
    const worksheetReader = new WorksheetReader({
      workbook: this,
      id: sheetNo,
      iterator: entry,
      options: this.options,
    });

    const matchingRel = (this.workbookRels || []).find(rel => rel.Target === `worksheets/sheet${sheetNo}.xml`);
    const matchingSheet = matchingRel && (this.model.sheets || []).find(sheet => sheet.rId === matchingRel.Id);
    if (matchingSheet) {
      worksheetReader.id = matchingSheet.id;
      worksheetReader.name = matchingSheet.name;
      worksheetReader.state = matchingSheet.state;
    }
    if (this.options.worksheets === 'emit') {
      yield {eventType: 'worksheet', value: worksheetReader};
    }
  }

  *_parseHyperlinks(entry, sheetNo) {
    this._emitEntry({type: 'hyperlinks', id: sheetNo});
    const hyperlinksReader = new HyperlinkReader({
      workbook: this,
      id: sheetNo,
      iterator: entry,
      options: this.options,
    });
    if (this.options.hyperlinks === 'emit') {
      yield {eventType: 'hyperlinks', value: hyperlinksReader};
    }
  }
}

// for reference - these are the valid values for options
WorkbookReader.Options = {
  worksheets: ['emit', 'ignore'],
  sharedStrings: ['cache', 'emit', 'ignore'],
  hyperlinks: ['cache', 'emit', 'ignore'],
  styles: ['cache', 'ignore'],
  entries: ['emit', 'ignore'],
};

module.exports = WorkbookReader;
