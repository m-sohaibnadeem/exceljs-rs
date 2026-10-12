//! # exceljs-rs
//!
//! The engine of exceljs-rs, a spreadsheet library for Node.js that is
//! API-compatible with [ExcelJS](https://github.com/exceljs/exceljs). It is
//! published to Node.js through the `exceljs-rs-node` bindings; this crate
//! holds all parsing and logic.
//!
//! The Node package keeps the object model in JS and calls this crate
//! for file I/O, exchanging ExcelJS's workbook model as JSON:
//! [`read_xlsx_model`] / [`write_xlsx_model`] for `workbook.xlsx`,
//! [`StreamWorkbook`] / [`StreamArchive`] for the streaming writer and
//! reader, and [`parse_csv`] for `workbook.csv`. It keeps ExcelJS 4.4.0's
//! behaviour; COMPATIBILITY.md lists every difference.

#![forbid(unsafe_code)]

mod col_cache;
mod csv;
mod error;
mod js;
mod jsv;
mod model_json;
mod reader;
mod rows;
mod sheet_data;
mod stream_reader;
mod stream_writer;
mod writer;
mod xform;
mod xml;
mod xml_stream;
mod zip_thread;

pub use csv::{parse_csv, CsvParseOptions};
pub use error::{Error, Result};
pub use reader::{
    read as read_xlsx_model, read_events as read_xlsx_events, ReadEvent, ReadOptions, ReadOutput,
};
pub use rows::SheetRows;
pub use stream_reader::{
    read_hyperlinks, read_relationships, read_shared_strings, read_style_models, read_workbook,
    read_worksheet_events, SheetReadOptions, StreamArchive, EVENT_COLS, EVENT_HYPERLINK,
    EVENT_NULL_ROW, EVENT_ROW,
};
pub use stream_writer::{StreamOptions, StreamWorkbook};
pub use writer::{write as write_xlsx_model, write_xlsx_parts, WriteOptions};
