//! Node.js bindings for `excel-rs`.
//!
//! The JS object model (`lib/doc`, from ExcelJS) talks to the Rust engine
//! through ExcelJS's workbook model, serialised as JSON (see
//! `lib/xlsx/codec.js`). Parsing and serialisation run on the libuv thread
//! pool.

use napi::bindgen_prelude::{AsyncTask, Buffer};
use napi::{Env, Error, Result, Task};
use napi_derive::napi;

/// Result of `readXlsx`: the shared-object table and the model, as JSON.
#[napi(object)]
pub struct ReadResult {
    pub refs: String,
    pub model: String,
    /// Each worksheet's rows (binary, see `lib/xlsx/rows.js`).
    pub rows: Buffer,
}

fn to_napi(e: impl std::fmt::Display) -> Error {
    Error::from_reason(e.to_string())
}

/// Parses `.xlsx` bytes off the main thread.
pub struct ReadTask {
    data: Vec<u8>,
    options: excel_rs::ReadOptions,
}

impl Task for ReadTask {
    type Output = ReadResult;
    type JsValue = ReadResult;

    fn compute(&mut self) -> Result<Self::Output> {
        let out = excel_rs::read_xlsx_model(&self.data, &self.options)
            .map_err(|e| Error::from_reason(e.exceljs_message()))?;
        Ok(ReadResult {
            refs: serde_json::to_string(&out.refs).map_err(to_napi)?,
            model: serde_json::to_string(&out.model).map_err(to_napi)?,
            rows: out.rows.encode().into(),
        })
    }

    fn resolve(&mut self, _env: Env, output: Self::Output) -> Result<Self::JsValue> {
        Ok(output)
    }
}

fn read_options(json: &str) -> Result<excel_rs::ReadOptions> {
    let v: serde_json::Value = serde_json::from_str(json).map_err(to_napi)?;
    let count = |k: &str| {
        v.get(k)
            .and_then(serde_json::Value::as_f64)
            .filter(|n| *n > 0.0)
            .map(|n| {
                #[allow(clippy::cast_possible_truncation, clippy::cast_sign_loss)]
                let n = n as usize;
                n
            })
    };
    let ignore_nodes = v
        .get("ignoreNodes")
        .and_then(serde_json::Value::as_array)
        .map(|a| {
            a.iter()
                .filter_map(|n| n.as_str().map(str::to_owned))
                .collect()
        })
        .unwrap_or_default();
    Ok(excel_rs::ReadOptions {
        max_rows: count("maxRows"),
        max_cols: count("maxCols"),
        ignore_nodes,
    })
}

/// `readXlsx(buffer, optionsJson): Promise<{refs, model}>`.
#[napi(ts_return_type = "Promise<ReadResult>")]
pub fn read_xlsx(data: Buffer, options: String) -> Result<AsyncTask<ReadTask>> {
    Ok(AsyncTask::new(ReadTask {
        data: data.to_vec(),
        options: read_options(&options)?,
    }))
}

/// Serialises a workbook model off the main thread.
pub struct WriteTask {
    model: String,
    /// JS's buffer itself (not copied): JS does not touch it meanwhile.
    rows: Buffer,
    styles: String,
    options: excel_rs::WriteOptions,
}

impl Task for WriteTask {
    type Output = Vec<u8>;
    type JsValue = Buffer;

    fn compute(&mut self) -> Result<Self::Output> {
        excel_rs::write_xlsx_parts(&self.model, &self.rows, &self.styles, &self.options)
            .map_err(|e| Error::from_reason(e.exceljs_message()))
    }

    fn resolve(&mut self, _env: Env, output: Self::Output) -> Result<Self::JsValue> {
        Ok(output.into())
    }
}

fn write_options(json: &str) -> Result<excel_rs::WriteOptions> {
    let v: serde_json::Value = serde_json::from_str(json).map_err(to_napi)?;
    let flag = |k: &str| {
        v.get(k).map(|b| match b {
            serde_json::Value::Bool(b) => *b,
            serde_json::Value::Null => false,
            serde_json::Value::Number(n) => n.as_f64().is_some_and(|n| n != 0.0),
            serde_json::Value::String(s) => !s.is_empty(),
            _ => true,
        })
    };
    let level = v
        .get("zip")
        .and_then(|z| z.get("compressionOptions"))
        .and_then(|c| c.get("level"))
        .and_then(serde_json::Value::as_i64);
    #[allow(clippy::cast_precision_loss)]
    let now_ms = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map_or(0.0, |d| d.as_millis() as f64);
    Ok(excel_rs::WriteOptions {
        use_styles: flag("useStyles"),
        use_shared_strings: flag("useSharedStrings"),
        compression_level: level,
        now_ms,
    })
}

/// `writeXlsx(modelJson, rows, stylesJson, optionsJson): Promise<Buffer>`:
/// the model without worksheet rows, each worksheet's rows, and the style
/// objects the rows refer to (see `lib/xlsx/rows.js`).
#[napi(ts_return_type = "Promise<Buffer>")]
pub fn write_xlsx(
    model: String,
    rows: Buffer,
    styles: String,
    options: String,
) -> Result<AsyncTask<WriteTask>> {
    Ok(AsyncTask::new(WriteTask {
        model,
        rows,
        styles,
        options: write_options(&options)?,
    }))
}

/// The engine of the streaming `WorkbookWriter` (see
/// `lib/stream/xlsx/workbook-writer.js`).
#[napi]
pub struct StreamWorkbook {
    inner: excel_rs::StreamWorkbook,
}

fn engine_err(e: excel_rs::Error) -> Error {
    Error::from_reason(e.exceljs_message())
}

fn parse_json(json: &str) -> Result<serde_json::Value> {
    serde_json::from_str(json).map_err(to_napi)
}

#[napi]
impl StreamWorkbook {
    /// `new StreamWorkbook(useStyles, compressionLevel)`.
    #[napi(constructor)]
    pub fn new(use_styles: bool, compression_level: Option<i64>) -> Self {
        StreamWorkbook {
            inner: excel_rs::StreamWorkbook::new(&excel_rs::StreamOptions {
                use_styles,
                compression_level,
            }),
        }
    }

    #[napi]
    pub fn open(&mut self, name: String) {
        self.inner.open(&name);
    }

    #[napi]
    pub fn write(&mut self, name: String, data: String) -> Result<()> {
        self.inner.write(&name, data.as_bytes()).map_err(engine_err)
    }

    #[napi]
    pub fn write_bytes(&mut self, name: String, data: Buffer) -> Result<()> {
        self.inner.write(&name, &data).map_err(engine_err)
    }

    #[napi]
    pub fn end(&mut self, name: String) -> Result<()> {
        self.inner.end(&name).map_err(engine_err)
    }

    #[napi]
    pub fn append(&mut self, name: String, data: String) -> Result<()> {
        self.inner
            .append(&name, data.as_bytes())
            .map_err(engine_err)
    }

    #[napi]
    pub fn append_bytes(&mut self, name: String, data: Buffer) -> Result<()> {
        self.inner.append(&name, &data).map_err(engine_err)
    }

    /// Prepares and writes committed rows (`rows`: row models as in
    /// `writeXlsx`, `styles`: style objects new since the last call); returns
    /// the per-row `{hyperlinks, comments}` events as JSON.
    #[napi]
    pub fn write_rows(
        &mut self,
        name: String,
        rows: Buffer,
        styles: String,
        use_shared_strings: bool,
    ) -> Result<String> {
        let out = self
            .inner
            .write_rows(&name, &rows, &styles, use_shared_strings)
            .map_err(engine_err)?;
        serde_json::to_string(&out).map_err(to_napi)
    }

    /// Renders a part; `model` is JSON or undefined.
    #[napi]
    pub fn render(&mut self, kind: String, model: Option<String>) -> Result<String> {
        let model = model.map(|m| parse_json(&m)).transpose()?;
        self.inner.render(&kind, model).map_err(engine_err)
    }

    #[napi]
    pub fn shared_strings_count(&self) -> u32 {
        u32::try_from(self.inner.shared_strings_count()).unwrap_or(u32::MAX)
    }

    #[napi]
    pub fn finalize(&mut self) -> Result<()> {
        self.inner.finalize().map_err(engine_err)
    }

    /// Archive bytes produced since the last call, copied into a buffer V8
    /// owns (and accounts for), so the engine reuses its own.
    #[napi(ts_return_type = "Buffer")]
    pub fn take_output<'env>(
        &mut self,
        env: &'env Env,
    ) -> Result<napi::bindgen_prelude::BufferSlice<'env>> {
        self.inner
            .take_output_with(|bytes| napi::bindgen_prelude::BufferSlice::copy_from(env, bytes))
    }
}

/// The engine of the streaming `WorkbookReader` (see
/// `lib/stream/xlsx/workbook-reader.js`).
#[napi]
pub struct StreamReader {
    archive: excel_rs::StreamArchive,
}

/// Parses one entry off the main thread.
pub struct EntryTask {
    archive: excel_rs::StreamArchive,
    index: usize,
    name: String,
    kind: String,
}

impl Task for EntryTask {
    type Output = String;
    type JsValue = String;

    fn compute(&mut self) -> Result<Self::Output> {
        let bytes = self.archive.entry(self.index).map_err(engine_err)?;
        let name = self.name.as_str();
        let value = match self.kind.as_str() {
            "rels" => excel_rs::read_relationships(name, &bytes),
            "workbook" => excel_rs::read_workbook(name, &bytes),
            "sharedStrings" => {
                excel_rs::read_shared_strings(name, &bytes).map(serde_json::Value::Array)
            }
            "hyperlinks" => excel_rs::read_hyperlinks(name, &bytes).map(serde_json::Value::Array),
            // [refsJson, modelJson], for codec.decodeRead
            "styles" => {
                let o = excel_rs::read_style_models(name, &bytes).map_err(engine_err)?;
                let pair = (
                    serde_json::to_string(&o.refs).map_err(to_napi)?,
                    serde_json::to_string(&o.model).map_err(to_napi)?,
                );
                return serde_json::to_string(&pair).map_err(to_napi);
            }
            other => return Err(Error::from_reason(format!("unknown entry kind {other}"))),
        }
        .map_err(engine_err)?;
        serde_json::to_string(&value).map_err(to_napi)
    }

    fn resolve(&mut self, _env: Env, output: Self::Output) -> Result<Self::JsValue> {
        Ok(output)
    }
}

type Batches = std::sync::mpsc::Receiver<std::result::Result<Vec<u8>, String>>;

/// Batches of worksheet events, parsed on a worker thread.
#[napi]
pub struct SheetCursor {
    batches: std::sync::Arc<std::sync::Mutex<Batches>>,
}

/// Waits for the next batch off the main thread.
pub struct NextBatchTask {
    batches: std::sync::Arc<std::sync::Mutex<Batches>>,
}

impl Task for NextBatchTask {
    type Output = Option<Vec<u8>>;
    type JsValue = Option<Buffer>;

    fn compute(&mut self) -> Result<Self::Output> {
        let rx = self.batches.lock().map_err(to_napi)?;
        match rx.recv() {
            Ok(Ok(batch)) => Ok(Some(batch)),
            Ok(Err(message)) => Err(Error::from_reason(message)),
            Err(_) => Ok(None),
        }
    }

    fn resolve(&mut self, _env: Env, output: Self::Output) -> Result<Self::JsValue> {
        Ok(output.map(Buffer::from))
    }
}

#[napi]
impl SheetCursor {
    /// The next batch of events (binary, see `excel_rs::EVENT_ROW`), or
    /// null at the end.
    #[napi(ts_return_type = "Promise<Buffer | null>")]
    pub fn next(&self) -> AsyncTask<NextBatchTask> {
        AsyncTask::new(NextBatchTask {
            batches: self.batches.clone(),
        })
    }
}

#[napi]
impl StreamReader {
    #[napi(constructor)]
    pub fn new(data: Buffer) -> Result<Self> {
        Ok(StreamReader {
            archive: excel_rs::StreamArchive::new(data.to_vec()).map_err(engine_err)?,
        })
    }

    /// Entry paths in archive order.
    #[napi]
    pub fn entry_names(&mut self) -> Result<Vec<String>> {
        self.archive.entry_names().map_err(engine_err)
    }

    /// Parses entry `index` as `kind` (rels, workbook, sharedStrings,
    /// hyperlinks, styles); resolves to JSON.
    #[napi(ts_return_type = "Promise<string>")]
    pub fn read_entry(&self, index: u32, name: String, kind: String) -> AsyncTask<EntryTask> {
        AsyncTask::new(EntryTask {
            archive: self.archive.clone(),
            index: index as usize,
            name,
            kind,
        })
    }

    /// Starts parsing worksheet entry `index` on a worker thread.
    #[napi]
    pub fn open_worksheet(
        &self,
        index: u32,
        name: String,
        emit_sheet: bool,
        hyperlinks: bool,
        batch_rows: u32,
    ) -> SheetCursor {
        let (tx, rx) = std::sync::mpsc::sync_channel(4);
        let mut archive = self.archive.clone();
        let options = excel_rs::SheetReadOptions {
            emit_sheet,
            hyperlinks,
            batch_rows: (batch_rows as usize).max(1),
        };
        std::thread::spawn(move || {
            let result = excel_rs::read_worksheet_events(
                &name,
                &mut archive,
                index as usize,
                options,
                |events| tx.send(Ok(events)).is_ok(),
            );
            if let Err(e) = result {
                let _ = tx.send(Err(e.exceljs_message()));
            }
        });
        SheetCursor {
            batches: std::sync::Arc::new(std::sync::Mutex::new(rx)),
        }
    }
}

// ---------------------------------------------------------------------------
// xlsx reading with rows in batches

type ReadMsg = std::result::Result<excel_rs::ReadEvent, String>;
type ReadRx = std::sync::Arc<std::sync::Mutex<Option<std::sync::mpsc::Receiver<ReadMsg>>>>;

/// The model part of an [`XlsxReader`]: the shared-object table and the
/// model as JSON, and the plain shared strings (u32 count, then each as u32
/// byte length and UTF-8).
#[napi(object)]
pub struct ReadHead {
    pub refs: String,
    pub model: String,
    pub strings: Buffer,
}

/// Reads an `.xlsx` on a thread of its own: `model()` resolves when the
/// model is ready, and the rows follow in batches (`nextRows()`) while the
/// caller builds its objects.
#[napi]
pub struct XlsxReader {
    rx: ReadRx,
}

fn recv(rx: &ReadRx) -> Result<Option<excel_rs::ReadEvent>> {
    let guard = rx.lock().map_err(to_napi)?;
    let Some(rx) = guard.as_ref() else {
        return Ok(None);
    };
    match rx.recv() {
        Ok(Ok(event)) => Ok(Some(event)),
        Ok(Err(message)) => Err(Error::from_reason(message)),
        // the reader thread is done
        Err(_) => Ok(None),
    }
}

/// Waits for the model off the main thread.
pub struct ModelTask {
    rx: ReadRx,
}

impl Task for ModelTask {
    type Output = ReadHead;
    type JsValue = ReadHead;

    fn compute(&mut self) -> Result<Self::Output> {
        match recv(&self.rx)? {
            Some(excel_rs::ReadEvent::Model {
                refs,
                model,
                shared_strings,
            }) => {
                let mut strings = Vec::new();
                put_u32(&mut strings, shared_strings.len());
                for s in &shared_strings {
                    put_u32(&mut strings, s.len());
                    strings.extend_from_slice(s.as_bytes());
                }
                Ok(ReadHead {
                    refs: serde_json::to_string(&refs).map_err(to_napi)?,
                    model: serde_json::to_string(&model).map_err(to_napi)?,
                    strings: strings.into(),
                })
            }
            _ => Err(Error::from_reason("excel-rs: no model")),
        }
    }

    fn resolve(&mut self, _env: Env, output: Self::Output) -> Result<Self::JsValue> {
        Ok(output)
    }
}

fn put_u32(out: &mut Vec<u8>, n: usize) {
    out.extend_from_slice(&u32::try_from(n).unwrap_or(u32::MAX).to_le_bytes());
}

#[napi]
impl XlsxReader {
    #[napi(constructor)]
    pub fn new(data: Buffer, options: String) -> Result<Self> {
        let options = read_options(&options)?;
        // JS's buffer itself, not a copy (napi4 releases it on the JS
        // thread when the reader thread drops it).
        let (tx, rx) = std::sync::mpsc::sync_channel::<ReadMsg>(4);
        std::thread::Builder::new()
            .name("excel-rs read".into())
            .spawn(move || {
                let result = excel_rs::read_xlsx_events(&data, &options, &mut |event| {
                    tx.send(Ok(event)).is_ok()
                });
                if let Err(e) = result {
                    let _ = tx.send(Err(e.exceljs_message()));
                }
            })
            .map_err(to_napi)?;
        Ok(XlsxReader {
            rx: std::sync::Arc::new(std::sync::Mutex::new(Some(rx))),
        })
    }

    /// Resolves to the model once it is read.
    #[napi(ts_return_type = "Promise<ReadHead>")]
    pub fn model(&self) -> AsyncTask<ModelTask> {
        AsyncTask::new(ModelTask {
            rx: self.rx.clone(),
        })
    }

    /// The next batch of rows, waiting for it: u32 worksheet index, u8 1 if
    /// it is the worksheet's last, u32 byte length and JSON of the shared
    /// objects added, u32 row count, then the rows (`lib/xlsx/rows.js`).
    /// Null at the end; throws the reader's error.
    #[napi]
    pub fn next_rows(&self) -> Result<Option<Buffer>> {
        match recv(&self.rx)? {
            Some(excel_rs::ReadEvent::Rows {
                sheet,
                refs,
                count,
                bytes,
                last,
            }) => {
                let refs = serde_json::to_string(&refs).map_err(to_napi)?;
                let mut out = Vec::with_capacity(bytes.len() + refs.len() + 13);
                put_u32(&mut out, sheet);
                out.push(u8::from(last));
                put_u32(&mut out, refs.len());
                out.extend_from_slice(refs.as_bytes());
                put_u32(&mut out, count);
                out.extend_from_slice(&bytes);
                Ok(Some(out.into()))
            }
            Some(excel_rs::ReadEvent::Model { .. }) => {
                Err(Error::from_reason("excel-rs: model read twice"))
            }
            None => Ok(None),
        }
    }

    /// Stops reading (the rest is not needed).
    #[napi]
    pub fn close(&self) -> Result<()> {
        self.rx.lock().map_err(to_napi)?.take();
        Ok(())
    }
}

/// Result of `parseCsv`: complete rows (JSON array of string arrays) and the
/// unparsed rest of the input.
#[napi(object)]
pub struct CsvChunk {
    pub rows: String,
    pub rest: String,
}

/// `Parser.parse(line, hasMoreData)` of `@fast-csv/parse`.
#[napi]
pub fn parse_csv(text: String, options: String, has_more_data: bool) -> Result<CsvChunk> {
    let v = parse_json(&options)?;
    let s = |k: &str| {
        v.get(k)
            .and_then(serde_json::Value::as_str)
            .map(str::to_owned)
    };
    let b = |k: &str| {
        v.get(k)
            .and_then(serde_json::Value::as_bool)
            .unwrap_or(false)
    };
    let opts = excel_rs::CsvParseOptions {
        delimiter: s("delimiter").unwrap_or_else(|| ",".into()),
        quote: s("quote"),
        escape_char: s("escapeChar"),
        comment: s("comment"),
        ltrim: b("ltrim"),
        rtrim: b("rtrim"),
        trim: b("trim"),
        ignore_empty: b("ignoreEmpty"),
        carriage_return: s("carriageReturn").unwrap_or_default(),
    };
    let (rows, rest) =
        excel_rs::parse_csv(&text, &opts, has_more_data).map_err(Error::from_reason)?;
    Ok(CsvChunk {
        rows: serde_json::to_string(&rows).map_err(to_napi)?,
        rest,
    })
}

/// `"release"` or `"debug"`: how the addon was built (the benchmark refuses
/// a debug build).
#[napi]
pub fn build_profile() -> &'static str {
    if cfg!(debug_assertions) {
        "debug"
    } else {
        "release"
    }
}
