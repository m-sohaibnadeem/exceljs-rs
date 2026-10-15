//! The engine behind ExcelJS's streaming `WorkbookWriter`
//! (`lib/stream/xlsx/workbook-writer.js`, `worksheet-writer.js`).
//!
//! The JS side keeps ExcelJS's object model (rows are buffered until
//! committed) and hands committed rows and parts here. This side owns the
//! style manager, the shared strings, the per-sheet shared-formula state and
//! the zip archive. Like ExcelJS's archiver, entries are written in the order
//! they were opened: the oldest open entry streams straight into the
//! archive and later ones are buffered until it ends. The archive's bytes
//! are collected for the JS side to take and write to its stream.

use std::collections::{HashMap, VecDeque};
use std::io::Write;
use std::sync::{Arc, Mutex};

use serde_json::{json, Value};
use zip::write::SimpleFileOptions;
use zip::ZipWriter;

use crate::error::{Error, Result};
use crate::jsv::get;
use crate::rows::{
    finish_master_range, prepare_cell, prepare_row, render_row, CellContext, CellPrep, MasterRef,
    Masters, Merges, RowModel, StyleTable,
};
use crate::writer::{
    prepare_cfs, prepare_cols, render_data_validations, SharedStrings, StyleManager,
};
use crate::xform::base::ListXform;
use crate::xform::book::{prepare as prepare_workbook, WorkbookXform};
use crate::xform::cf::ConditionalFormattingsXform;
use crate::xform::comment::render_comment_pair;
use crate::xform::core::{render_app, render_content_types, render_core, render_relationships};
use crate::xform::sheet::{
    auto_filter, col, hyperlink, page_margins, page_setup, picture, sheet_format_properties,
    sheet_protection, HeaderFooterXform, RowBreaksXform, SheetPropertiesXform, SheetViewXform,
};
use crate::xform::strings::{render_shared_strings, RichTextXform};
use crate::xform::to_xml;
use crate::xml_stream::XmlStream;
use crate::zip_thread::{ZipStreamOut, ZipThread};

/// The archive's output, shared between the zip writer and the taker.
#[derive(Debug, Clone, Default)]
struct Sink(Arc<Mutex<Vec<u8>>>);

impl Sink {
    fn take(&self) -> Vec<u8> {
        self.0
            .lock()
            .map(|mut v| std::mem::take(&mut *v))
            .unwrap_or_default()
    }

    /// `f` on the bytes produced so far, which are then dropped (the buffer
    /// is kept for the next ones).
    fn take_with<T>(&self, f: impl FnOnce(&[u8]) -> T) -> T {
        match self.0.lock() {
            Ok(mut v) => {
                let out = f(&v);
                v.clear();
                out
            }
            Err(_) => f(&[]),
        }
    }
}

impl Write for Sink {
    fn write(&mut self, buf: &[u8]) -> std::io::Result<usize> {
        let mut v = self
            .0
            .lock()
            .map_err(|_| std::io::Error::other("poisoned"))?;
        v.extend_from_slice(buf);
        Ok(buf.len())
    }
    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }
}

#[derive(Debug)]
struct Entry {
    name: String,
    pending: Vec<u8>,
    ended: bool,
}

#[derive(Debug, Default)]
struct SheetState {
    /// Formula cells by address (`this._formulae`), as (row, cell) positions
    /// in the order rows were written; a later clone updates its master.
    formulae: HashMap<String, MasterRef>,
    masters: Masters,
    /// Rows written so far.
    rows: usize,
}

/// Options of `new WorkbookWriter(options)` the engine needs.
#[derive(Debug, Clone, Default)]
pub struct StreamOptions {
    /// `useStyles` (StylesXform or its Mock).
    pub use_styles: bool,
    /// Deflate level from `zip` options, when given.
    pub compression_level: Option<i64>,
}

/// A streaming `.xlsx` archive plus the workbook-wide writer state.
pub struct StreamWorkbook {
    /// Compressed on another thread while JS builds the next rows.
    zip: Option<ZipThread<ZipStreamOut<Sink>>>,
    sink: Sink,
    entries: VecDeque<Entry>,
    head_started: bool,
    styles: StyleManager,
    table: StyleTable,
    shared_strings: SharedStrings,
    sheets: HashMap<String, SheetState>,
}

impl std::fmt::Debug for StreamWorkbook {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("StreamWorkbook")
            .field("entries", &self.entries.len())
            .finish_non_exhaustive()
    }
}

fn invalid(m: String) -> Error {
    Error::invalid("stream", m)
}

impl StreamWorkbook {
    /// `new WorkbookWriter(options)` (engine part).
    pub fn new(options: &StreamOptions) -> Self {
        let sink = Sink::default();
        StreamWorkbook {
            zip: Some(ZipThread::new(ZipStreamOut {
                zip: ZipWriter::new_stream(sink.clone()),
                options: SimpleFileOptions::default()
                    .compression_method(zip::CompressionMethod::Deflated)
                    .compression_level(options.compression_level),
            })),
            sink,
            entries: VecDeque::new(),
            head_started: false,
            styles: StyleManager::new(options.use_styles),
            table: StyleTable::default(),
            shared_strings: SharedStrings::default(),
            sheets: HashMap::new(),
        }
    }

    /// Opens an entry (`_openStream(path)`); a leading `/` is dropped as
    /// archiver does.
    pub fn open(&mut self, name: &str) {
        let name = name.trim_start_matches('/').to_owned();
        self.entries.push_back(Entry {
            name,
            pending: Vec::new(),
            ended: false,
        });
    }

    fn entry(&mut self, name: &str) -> Result<&mut Entry> {
        let name = name.trim_start_matches('/');
        self.entries
            .iter_mut()
            .find(|e| e.name == name && !e.ended)
            .ok_or_else(|| invalid(format!("write after end: {name}")))
    }

    /// Appends data to an open entry.
    pub fn write(&mut self, name: &str, data: &[u8]) -> Result<()> {
        self.entry(name)?.pending.extend_from_slice(data);
        self.pump()
    }

    /// Ends an entry.
    pub fn end(&mut self, name: &str) -> Result<()> {
        self.entry(name)?.ended = true;
        self.pump()
    }

    /// Adds a complete entry (`zip.append(data, {name})`).
    pub fn append(&mut self, name: &str, data: &[u8]) -> Result<()> {
        self.open(name);
        self.write(name, data)?;
        self.end(name)
    }

    fn pump(&mut self) -> Result<()> {
        let zip = self
            .zip
            .as_mut()
            .ok_or_else(|| invalid("archive already finalized".into()))?;
        while let Some(head) = self.entries.front_mut() {
            if !self.head_started {
                zip.start(head.name.clone())?;
                self.head_started = true;
            }
            if !head.pending.is_empty() {
                zip.data(std::mem::take(&mut head.pending))?;
            }
            if !head.ended {
                break;
            }
            self.entries.pop_front();
            self.head_started = false;
        }
        Ok(())
    }

    /// `zip.finalize()`: every entry must have ended.
    pub fn finalize(&mut self) -> Result<()> {
        if let Some(e) = self.entries.iter().find(|e| !e.ended) {
            return Err(invalid(format!("entry not ended: {}", e.name)));
        }
        self.pump()?;
        if let Some(zip) = self.zip.take() {
            zip.finish()?;
        }
        Ok(())
    }

    /// Takes the archive bytes produced so far.
    pub fn take_output(&mut self) -> Vec<u8> {
        self.sink.take()
    }

    /// The additive bulk helper may only write the active archive entry.
    pub fn can_write_rows(&self, name: &str) -> bool {
        self.entries
            .front()
            .is_some_and(|e| !e.ended && e.name == name.trim_start_matches('/'))
    }

    pub fn checkpoint(&mut self) -> Result<()> {
        self.pump()?;
        self.zip
            .as_mut()
            .ok_or_else(|| invalid("archive already finalized".into()))?
            .checkpoint()
    }

    /// Stop accepting data and release the compression worker immediately.
    pub fn abort(&mut self) -> Result<()> {
        self.entries.clear();
        if let Some(zip) = self.zip.take() {
            zip.finish()?;
        }
        self.sink.take();
        Ok(())
    }

    /// Gives `f` the archive bytes produced so far, without handing over the
    /// buffer (it is reused, so a long stream does not allocate one per
    /// flush).
    pub fn take_output_with<T>(&mut self, f: impl FnOnce(&[u8]) -> T) -> T {
        self.sink.take_with(f)
    }

    /// `sharedStrings.count` (unique strings).
    pub fn shared_strings_count(&self) -> usize {
        self.shared_strings.values.len()
    }

    /// `_writeRow(row)` for committed rows: prepares each row model (styles,
    /// shared strings, shared formulae) and writes it to the sheet entry.
    /// Returns, for each row that has any, the `{hyperlinks, comments}` its
    /// prepare gathered, in row order, for the JS rels and comments writers.
    pub fn write_rows(
        &mut self,
        entry: &str,
        rows: &[u8],
        new_styles_json: &str,
        use_shared_strings: bool,
    ) -> Result<Value> {
        let rows: Vec<RowModel> = crate::rows::decode_sheets(rows)
            .map_err(invalid)?
            .into_iter()
            .flatten()
            .collect();
        let new_styles: Vec<Value> =
            serde_json::from_str(new_styles_json).map_err(|e| invalid(e.to_string()))?;
        self.table.styles.extend(new_styles);
        let mut events = Vec::new();
        let mut out = XmlStream::new();
        let rich = RichTextXform::new();
        let state = self
            .sheets
            .entry(entry.trim_start_matches('/').to_owned())
            .or_default();
        // The streaming writer's merges.add is a no-op.
        let mut merges = Merges::ignored();
        for mut row in rows {
            prepare_row(&mut row, &mut self.styles, &mut self.table);
            // options.siFormulae is copied into each row's options.
            let mut si_formulae = 0i64;
            let mut hyperlinks = Vec::new();
            let mut comments = Vec::new();
            // cells of this row registered as masters
            let mut row_masters = Vec::new();
            let ri = state.rows;
            state.rows += 1;
            for ci in 0..row.cells.len() {
                let Some(cell) = row.cells.get_mut(ci) else {
                    continue;
                };
                let mut ctx = CellContext {
                    styles: &mut self.styles,
                    table: &mut self.table,
                    shared_strings: if use_shared_strings {
                        Some(&mut self.shared_strings)
                    } else {
                        None
                    },
                    date1904: false,
                };
                let prep = prepare_cell(
                    cell,
                    &mut ctx,
                    &mut hyperlinks,
                    &mut comments,
                    &mut merges,
                    &state.formulae,
                    &mut si_formulae,
                )
                .map_err(invalid)?;
                match prep {
                    CellPrep::None => {}
                    CellPrep::Register(address) => {
                        row_masters.push(ci);
                        state.formulae.insert(address, (ri, ci));
                        state.masters.register((ri, ci), cell);
                    }
                    CellPrep::Clone {
                        master,
                        master_address,
                        address,
                    } => {
                        cell.si = Some(Box::new(
                            state
                                .masters
                                .share(master, &master_address, &address, &mut si_formulae)
                                .map_err(invalid)?,
                        ));
                    }
                }
            }
            // A clone later in the same row updates its master before the
            // row is rendered.
            for ci in row_masters {
                if let Some(cell) = row.cells.get_mut(ci) {
                    state.masters.apply((ri, ci), cell);
                    finish_master_range(cell).map_err(invalid)?;
                }
            }
            render_row(&mut out, &row, &rich);
            if !hyperlinks.is_empty() || !comments.is_empty() {
                events.push(json!({"hyperlinks": hyperlinks, "comments": comments}));
            }
        }
        let xml = out.into_xml();
        self.write(entry, xml.as_bytes())?;
        Ok(Value::Array(events))
    }

    /// Renders one part with ExcelJS's xform for `kind` (the `toXml` calls of
    /// the streaming writers), running the xform's `prepare` where the
    /// streaming writer does.
    pub fn render(&mut self, kind: &str, model: Option<Value>) -> Result<String> {
        let mut model = model;
        let m = model.as_ref();
        let xml = match kind {
            "sheetProperties" => to_xml(&SheetPropertiesXform::new(), m),
            "sheetViews" => list_xml(
                ListXform::new("sheetViews", Box::new(SheetViewXform::default())),
                m,
            ),
            "sheetFormatProperties" => to_xml(&sheet_format_properties(), m),
            "cols" => {
                if let Some(cols) = model.as_mut() {
                    prepare_cols(cols, &mut self.styles);
                }
                list_xml(ListXform::new("cols", Box::new(col())), model.as_ref())
            }
            "hyperlinks" => list_xml(ListXform::new("hyperlinks", Box::new(hyperlink())), m),
            "conditionalFormattings" => {
                prepare_cfs(model.as_mut(), &mut self.styles).map_err(invalid)?;
                to_xml(&ConditionalFormattingsXform::new(), model.as_ref())
            }
            "dataValidations" => {
                let mut xml = XmlStream::new();
                render_data_validations(&mut xml, m).map_err(invalid)?;
                xml.into_xml()
            }
            "sheetProtection" => to_xml(&sheet_protection(), m),
            "pageMargins" => to_xml(&page_margins(), m),
            "pageSetup" => to_xml(&page_setup(), m),
            "headerFooter" => to_xml(&HeaderFooterXform::default(), m),
            "autoFilter" => to_xml(&auto_filter(), m),
            "picture" => to_xml(&picture(), m),
            "rowBreaks" => to_xml(&RowBreaksXform::new(), m),
            "relationships" => {
                render_relationships(m.and_then(Value::as_array).map_or(&[], Vec::as_slice))
            }
            "contentTypes" => {
                let mut ct = model.unwrap_or(Value::Null);
                crate::jsv::set(
                    &mut ct,
                    "sharedStrings",
                    json!({"count": self.shared_strings_count()}),
                );
                render_content_types(&ct)
            }
            "app" => render_app(m.unwrap_or(&Value::Null)),
            "core" => render_core(m.unwrap_or(&Value::Null)).map_err(invalid)?,
            "sharedStrings" => render_shared_strings(
                &self.shared_strings.values,
                self.shared_strings.values.len(),
            ),
            "styles" => self.styles.render(),
            "workbook" => {
                let mut wb = model.unwrap_or(Value::Null);
                prepare_workbook(&mut wb).map_err(invalid)?;
                to_xml(&WorkbookXform::new(), Some(&wb))
            }
            "comment" => {
                let index = get(m, "index").and_then(Value::as_u64).unwrap_or(0);
                let comment = get(m, "comment").cloned().unwrap_or(Value::Null);
                let (c, v) = render_comment_pair(&comment, usize::try_from(index).unwrap_or(0));
                serde_json::to_string(&json!([c, v])).map_err(|e| invalid(e.to_string()))?
            }
            other => return Err(invalid(format!("unknown part kind {other}"))),
        };
        Ok(xml)
    }
}

fn list_xml(list: ListXform, model: Option<&Value>) -> String {
    let mut xml = XmlStream::new();
    list.render_list(&mut xml, model);
    xml.into_xml()
}
