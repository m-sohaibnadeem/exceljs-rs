// From ExcelJS 4.4.0 lib/stream/xlsx/hyperlink-reader.js (MIT, see LICENSE-EXCELJS).
// The relationships are parsed by the Rust engine.
'use strict';

const {EventEmitter} = require('events');

const codec = require('../../xlsx/codec');

class HyperlinkReader extends EventEmitter {
  constructor({workbook, id, iterator, options}) {
    super();

    this.workbook = workbook;
    this.id = id;
    this.iterator = iterator;
    this.options = options;
  }

  get count() {
    return (this.hyperlinks && this.hyperlinks.length) || 0;
  }

  each(fn) {
    return this.hyperlinks.forEach(fn);
  }

  async read() {
    const {iterator, options} = this;
    let emitHyperlinks = false;
    let hyperlinks = null;
    switch (options.hyperlinks) {
      case 'emit':
        emitHyperlinks = true;
        break;
      case 'cache':
        this.hyperlinks = hyperlinks = {};
        break;
      default:
        break;
    }

    if (!emitHyperlinks && !hyperlinks) {
      this.emit('finished');
      return;
    }

    try {
      // eslint-disable-next-line no-underscore-dangle
      const archive = this.workbook._archive;
      const relationships = codec.decode(await archive.readEntry(iterator.index, iterator.path, 'hyperlinks'));
      for (const relationship of relationships) {
        if (emitHyperlinks) {
          this.emit('hyperlink', relationship);
        } else {
          hyperlinks[relationship.rId] = relationship;
        }
      }
      this.emit('finished');
    } catch (error) {
      this.emit('error', error);
    }
  }
}

module.exports = HyperlinkReader;
