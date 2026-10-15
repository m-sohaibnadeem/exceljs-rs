//! The engine behind ExcelJS's streaming `WorkbookReader`
//! (`lib/stream/xlsx/workbook-reader.js`, `worksheet-reader.js`,
//! `hyperlink-reader.js`).
//!
//! The JS side keeps ExcelJS's orchestration (entry order, deferred
//! worksheets, events) and its cell-value rules. This side reads the archive
//! and runs the SAX state machines of those readers, which differ from the
//! xforms the document reader uses (for example the shared strings reader
//! keeps only the last `<t>` of a string). Worksheets are parsed into batches
//! of events: `["cols", [...]]`, `["row", {...}]` and `["hyperlink", {...}]`.

use std::io::Cursor;
use std::sync::Arc;

use serde_json::{json, Map, Value};
use zip::ZipArchive;

use crate::error::Result;
use crate::js::{decode_x_escapes, parse_float, parse_int};
use crate::jsv::{num, undefined};
use crate::xform::book::WorkbookXform;
use crate::xform::core::RelationshipsXform;
use crate::xform::parse_part;
use crate::xml::{self, Sax};

/// Shared archive bytes.
#[derive(Debug, Clone)]
struct Bytes(Arc<Vec<u8>>);

impl AsRef<[u8]> for Bytes {
    fn as_ref(&self) -> &[u8] {
        self.0.as_slice()
    }
}

/// An `.xlsx` archive read entry by entry.
#[derive(Debug, Clone)]
pub struct StreamArchive {
    zip: ZipArchive<Cursor<Bytes>>,
    bytes: Bytes,
}

impl StreamArchive {
    /// Opens the archive.
    pub fn new(data: Vec<u8>) -> Result<Self> {
        let bytes = Bytes(Arc::new(data));
        Ok(StreamArchive {
            zip: ZipArchive::new(Cursor::new(bytes.clone()))?,
            bytes,
        })
    }

    /// Validate budgets without retaining inflated entries.
    pub fn validate(&self, control: &crate::ReadControl) -> Result<()> {
        control.preflight(self.bytes.as_ref())
    }

    /// Entry paths in archive order (directories included, as unzipper
    /// reports them).
    pub fn entry_names(&mut self) -> Result<Vec<String>> {
        let mut names = Vec::with_capacity(self.zip.len());
        for i in 0..self.zip.len() {
            let e = self.zip.by_index_raw(i)?;
            names.push(e.name().map(|n| n.into_owned()).unwrap_or_default());
        }
        Ok(names)
    }

    /// The uncompressed content of entry `index`.
    pub fn entry(&mut self, index: usize) -> Result<Vec<u8>> {
        crate::read_control::read_all(self.zip.by_index(index)?)
    }
}

/// `bufferToString(chunk)`: UTF-8 with replacement characters.
fn text(bytes: &[u8]) -> std::borrow::Cow<'_, [u8]> {
    match String::from_utf8_lossy(bytes) {
        std::borrow::Cow::Borrowed(_) => std::borrow::Cow::Borrowed(bytes),
        std::borrow::Cow::Owned(s) => std::borrow::Cow::Owned(s.into_bytes()),
    }
}

/// `RelationshipsXform.parseStream`.
pub fn read_relationships(part: &str, bytes: &[u8]) -> Result<Value> {
    Ok(parse_part(&mut RelationshipsXform::default(), part, &text(bytes))?.unwrap_or(Value::Null))
}

/// `WorkbookXform.parseStream`: `{model, properties}` where `properties` is
/// the `workbookPr` model.
pub fn read_workbook(part: &str, bytes: &[u8]) -> Result<Value> {
    let model = parse_part(&mut WorkbookXform::new(), part, &text(bytes))?.unwrap_or(Value::Null);
    let properties = model.get("properties").cloned().unwrap_or(Value::Null);
    Ok(json!({"model": model, "properties": properties}))
}

/// `WorkbookReader._parseSharedStrings`: the strings in order, plain text or
/// `{richText: [{font, text}]}` with the stream reader's font shape.
///
/// Fixed from ExcelJS (B12 in COMPATIBILITY.md): `<color rgb>` gives `argb`,
/// `<rFont val>` gives `name`, `theme` is a number, only `<t>` text counts
/// (so whitespace between elements and phonetic `<rPh>` runs are not part of
/// the string), `_xHHHH_` escapes are decoded and `<si/>` is `""`.
pub fn read_shared_strings(part: &str, bytes: &[u8]) -> Result<Vec<Value>> {
    let mut out = Vec::new();
    let mut text_acc = String::new();
    let mut in_t = false;
    let mut skip = 0usize;
    let mut rich: Vec<Value> = Vec::new();
    let mut font: Option<Map<String, Value>> = None;
    let int = |v: Option<&str>| num(v.and_then(parse_int).unwrap_or(f64::NAN));
    xml::parse(part, &text(bytes), |e| {
        match e {
            Sax::Open(_) if skip > 0 => skip += 1,
            Sax::Open(node) => {
                fn f(font: &mut Option<Map<String, Value>>) -> &mut Map<String, Value> {
                    font.get_or_insert_with(Map::new)
                }
                match node.name {
                    "b" => {
                        f(&mut font).insert("bold".into(), json!(true));
                    }
                    "charset" => {
                        f(&mut font).insert("charset".into(), int(node.attr("charset")));
                    }
                    "color" => {
                        let mut color = Map::new();
                        if let Some(v) = node.attr("rgb").filter(|s| !s.is_empty()) {
                            color.insert("argb".into(), json!(v));
                        }
                        if let Some(v) = node.attr("val").filter(|s| !s.is_empty()) {
                            color.insert("argb".into(), json!(v));
                        }
                        if let Some(v) = node.attr("theme").filter(|s| !s.is_empty()) {
                            color.insert("theme".into(), int(Some(v)));
                        }
                        f(&mut font).insert("color".into(), Value::Object(color));
                    }
                    "family" => {
                        f(&mut font).insert("family".into(), int(node.attr("val")));
                    }
                    "i" => {
                        f(&mut font).insert("italic".into(), json!(true));
                    }
                    "outline" => {
                        f(&mut font).insert("outline".into(), json!(true));
                    }
                    "rFont" => {
                        f(&mut font).insert(
                            "name".into(),
                            node.attr("val").map_or_else(undefined, |s| json!(s)),
                        );
                    }
                    "rPh" | "phoneticPr" => skip = 1,
                    "si" => {
                        font = None;
                        rich.clear();
                        text_acc.clear();
                    }
                    "sz" => {
                        f(&mut font).insert("size".into(), int(node.attr("val")));
                    }
                    "t" => in_t = true,
                    "u" => {
                        f(&mut font).insert("underline".into(), json!(true));
                    }
                    "vertAlign" => {
                        f(&mut font).insert(
                            "vertAlign".into(),
                            node.attr("val").map_or_else(undefined, |s| json!(s)),
                        );
                    }
                    _ => {}
                }
            }
            Sax::Text(t) => {
                if in_t && skip == 0 {
                    text_acc.push_str(t);
                }
            }
            Sax::Close(_) if skip > 0 => skip -= 1,
            Sax::Close(name) => match name {
                "t" => in_t = false,
                "r" => {
                    rich.push(json!({
                        "font": font.take().map_or(Value::Null, Value::Object),
                        "text": decode_x_escapes(&std::mem::take(&mut text_acc)),
                    }));
                }
                "si" => {
                    out.push(if rich.is_empty() {
                        Value::String(decode_x_escapes(&std::mem::take(&mut text_acc)))
                    } else {
                        json!({"richText": std::mem::take(&mut rich)})
                    });
                    font = None;
                    text_acc.clear();
                }
                _ => {}
            },
        }
        Ok(true)
    })?;
    Ok(out)
}

/// `HyperlinkReader.read`: the hyperlink relationships.
pub fn read_hyperlinks(part: &str, bytes: &[u8]) -> Result<Vec<Value>> {
    const HYPERLINK: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    let mut out = Vec::new();
    xml::parse(part, &text(bytes), |e| {
        if let Sax::Open(node) = e {
            if node.name == "Relationship" && node.attr("Type") == Some(HYPERLINK) {
                let s = |k: &str| node.attr(k).map_or_else(undefined, |v| json!(v));
                // `type: Enums.RelationshipType.Styles` (3)
                out.push(json!({"type": 3, "rId": s("Id"), "target": s("Target"), "targetMode": s("TargetMode")}));
            }
        }
        Ok(true)
    })?;
    Ok(out)
}

/// What `WorksheetReader.parse` looks at.
#[derive(Debug, Clone, Copy)]
pub struct SheetReadOptions {
    /// `worksheets === 'emit'`.
    pub emit_sheet: bool,
    /// `hyperlinks` is `'emit'` or `'cache'`.
    pub hyperlinks: bool,
    /// Rows per batch.
    pub batch_rows: usize,
}

/// A cell as `WorksheetReader.parse` collects it before `</c>`.
#[derive(Debug, Default)]
struct Cell {
    r: Option<String>,
    s: f64,
    t: Option<String>,
    f: Option<String>,
    v: Option<String>,
}

#[derive(Clone, Copy, PartialEq)]
enum Current {
    None,
    F,
    V,
}

/// Event batches of `read_worksheet_events`, read by
/// `lib/stream/xlsx/worksheet-reader.js`: a sequence of records
///   ROW <f64 r> <u8 has ht> [<f64 ht>] <u8 has s> [<f64 s>] <u32 n> n cells
///   NULL_ROW                         (</row> with no row open)
///   HYPERLINK <opt str ref> <opt str rId>
///   COLS <str JSON>                  (codec JSON, non-finite numbers tagged)
/// where a cell is <opt str r> <f64 s> <opt str t> <opt str f> <opt str v>,
/// an opt str is 0 (undefined) or 1 <u32 len> <utf8>, and numbers are
/// little-endian f64 (NaN for JS NaN).
pub const EVENT_ROW: u8 = 1;
/// See [`EVENT_ROW`].
pub const EVENT_NULL_ROW: u8 = 2;
/// See [`EVENT_ROW`].
pub const EVENT_HYPERLINK: u8 = 3;
/// See [`EVENT_ROW`].
pub const EVENT_COLS: u8 = 4;

/// A `<row>` being read: (r, ht, s) and its cells.
type OpenRow = ((f64, Option<f64>, Option<f64>), Vec<Cell>);

struct Batch {
    out: Vec<u8>,
}

impl Batch {
    fn f64(&mut self, x: f64) {
        self.out.extend_from_slice(&x.to_le_bytes());
    }
    fn str(&mut self, s: &str) {
        let n = u32::try_from(s.len()).unwrap_or(u32::MAX);
        self.out.extend_from_slice(&n.to_le_bytes());
        self.out.extend_from_slice(s.as_bytes());
    }
    fn opt(&mut self, s: Option<&str>) {
        match s {
            None => self.out.push(0),
            Some(s) => {
                self.out.push(1);
                self.str(s);
            }
        }
    }
}

/// `WorksheetReader.parse`: calls `emit` with each batch of events (see
/// [`EVENT_ROW`]); `emit` returns false to stop.
///
/// The entry is inflated as it is parsed, so a worksheet is never whole in
/// memory.
pub fn read_worksheet_events(
    part: &str,
    archive: &mut StreamArchive,
    index: usize,
    options: SheetReadOptions,
    mut emit: impl FnMut(Vec<u8>) -> bool,
) -> Result<()> {
    if !options.emit_sheet && !options.hyperlinks {
        return Ok(());
    }
    let mut batch = Batch {
        out: Vec::with_capacity(1 << 16),
    };
    let mut rows_in_batch = 0usize;
    let mut stopped = false;
    let (mut in_cols, mut in_rows, mut in_hyperlinks) = (false, false, false);
    let mut cols: Vec<Value> = Vec::new();
    // the open row: (r, ht, s) and its cells
    let mut row: Option<OpenRow> = None;
    let mut c: Option<Cell> = None;
    // `current` outlives its cell in ExcelJS: text after a cell is appended
    // to the last `<f>`/`<v>` it created, which nothing reads again. Keeping
    // the target kind is enough to reproduce what is observable.
    let mut current = Current::None;
    let mut skip = 0usize;
    // the position of a parse error, which needs the text before it
    let again = archive.clone();
    let line_col = move |pos: usize| {
        let mut a = again.clone();
        a.entry(index)
            .map(|b| xml::line_col_of(&String::from_utf8_lossy(&b), pos))
            .unwrap_or((0, 0))
    };
    let entry = archive.zip.by_index(index)?;
    xml::parse_stream(part, entry, &line_col, |e| {
        match e {
            Sax::Open(node) => {
                if options.emit_sheet {
                    match node.name {
                        "cols" => {
                            in_cols = true;
                            cols.clear();
                        }
                        "sheetData" => in_rows = true,
                        "col" if in_cols => {
                            cols.push(json!({
                                "min": num(node.attr("min").and_then(parse_int).unwrap_or(f64::NAN)),
                                "max": num(node.attr("max").and_then(parse_int).unwrap_or(f64::NAN)),
                                "width": num(node.attr("width").map_or(f64::NAN, parse_float)),
                                "styleId": num(parse_int(node.attr("style").filter(|s| !s.is_empty()).unwrap_or("0")).unwrap_or(f64::NAN)),
                            }));
                        }
                        "row" if in_rows => {
                            let r = node.attr("r").and_then(parse_int).unwrap_or(f64::NAN);
                            let ht = node.attr("ht").filter(|s| !s.is_empty()).map(parse_float);
                            let st = node
                                .attr("s")
                                .filter(|s| !s.is_empty())
                                .map(|st| parse_int(st).unwrap_or(f64::NAN));
                            row = Some(((r, ht, st), Vec::new()));
                        }
                        "c" if row.is_some() => {
                            c = Some(Cell {
                                r: node.attr("r").map(str::to_owned),
                                s: node.attr("s").and_then(parse_int).unwrap_or(f64::NAN),
                                t: node.attr("t").map(str::to_owned),
                                f: None,
                                v: None,
                            });
                        }
                        "f" => {
                            if let Some(c) = c.as_mut() {
                                c.f = Some(String::new());
                                current = Current::F;
                            }
                        }
                        // Skipped whole: phonetic runs and their settings.
                        "rPh" | "phoneticPr" if c.is_some() => skip += 1,
                        // An inline string's text is that of all its runs.
                        // (ExcelJS restarts it at each <t>, so a rich inline
                        // string reads as its last run: B4 in COMPATIBILITY.md.)
                        "t" if skip == 0 => {
                            if let Some(c) = c.as_mut() {
                                c.v.get_or_insert_with(String::new);
                                current = Current::V;
                            }
                        }
                        "v" | "is" => {
                            if let Some(c) = c.as_mut() {
                                c.v = Some(String::new());
                                current = Current::V;
                            }
                        }
                        _ => {}
                    }
                }
                if options.hyperlinks {
                    match node.name {
                        "hyperlinks" => in_hyperlinks = true,
                        "hyperlink" if in_hyperlinks => {
                            batch.out.push(EVENT_HYPERLINK);
                            batch.opt(node.attr("ref"));
                            batch.opt(node.attr("r:id"));
                        }
                        _ => {}
                    }
                }
            }
            Sax::Text(t) => {
                if options.emit_sheet && skip == 0 {
                    if let Some(c) = c.as_mut() {
                        match current {
                            Current::F => c.f.get_or_insert_with(String::new).push_str(t),
                            Current::V => c.v.get_or_insert_with(String::new).push_str(t),
                            Current::None => {}
                        }
                    }
                }
            }
            Sax::Close(name) => {
                if options.emit_sheet {
                    match name {
                        "cols" => {
                            in_cols = false;
                            batch.out.push(EVENT_COLS);
                            batch.str(&Value::Array(std::mem::take(&mut cols)).to_string());
                        }
                        "sheetData" => in_rows = false,
                        "row" => {
                            match row.take() {
                                Some(((r, ht, st), cells)) => {
                                    batch.out.push(EVENT_ROW);
                                    batch.f64(r);
                                    for x in [ht, st] {
                                        match x {
                                            Some(x) => {
                                                batch.out.push(1);
                                                batch.f64(x);
                                            }
                                            None => batch.out.push(0),
                                        }
                                    }
                                    let n = u32::try_from(cells.len()).unwrap_or(u32::MAX);
                                    batch.out.extend_from_slice(&n.to_le_bytes());
                                    for cell in &cells {
                                        batch.opt(cell.r.as_deref());
                                        batch.f64(cell.s);
                                        batch.opt(cell.t.as_deref());
                                        batch.opt(cell.f.as_deref());
                                        batch.opt(cell.v.as_deref());
                                    }
                                }
                                None => batch.out.push(EVENT_NULL_ROW),
                            }
                            rows_in_batch += 1;
                            crate::read_control::row();
                        }
                        "rPh" | "phoneticPr" if skip > 0 => skip -= 1,
                        "c" => {
                            if let (Some((_, cells)), Some(mut cell)) = (row.as_mut(), c.take()) {
                                // `_xHHHH_` escapes (ExcelJS decodes none
                                // here: B1 in COMPATIBILITY.md).
                                if matches!(cell.t.as_deref(), Some("str" | "inlineStr")) {
                                    cell.v = cell.v.map(|v| decode_x_escapes(&v));
                                }
                                cells.push(cell);
                                // `current` still points at this cell's
                                // last part; later text goes nowhere visible.
                                current = Current::None;
                            }
                        }
                        _ => {}
                    }
                }
                if options.hyperlinks && name == "hyperlinks" {
                    in_hyperlinks = false;
                }
            }
        }
        if rows_in_batch >= options.batch_rows {
            rows_in_batch = 0;
            let out = std::mem::replace(&mut batch.out, Vec::with_capacity(1 << 16));
            if !emit(out) {
                stopped = true;
                return Ok(false);
            }
        }
        Ok(true)
    })?;
    if !stopped && !batch.out.is_empty() {
        emit(batch.out);
    }
    Ok(())
}

/// Style models for `getStyleModel(id)` of every cell style in `styles.xml`
/// (`styles: 'cache'`), as a shared-object table and a list.
pub fn read_style_models(part: &str, bytes: &[u8]) -> Result<crate::ReadOutput> {
    crate::reader::style_models(part, &text(bytes))
}
