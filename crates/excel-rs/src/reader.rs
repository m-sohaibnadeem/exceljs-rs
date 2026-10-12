//! Reading `.xlsx` into ExcelJS's workbook model: the port of
//! `XLSX.load` and `XLSX.reconcile` (`lib/xlsx/xlsx.js`), plus the
//! `reconcile` steps of the workbook, worksheet, row, cell, column, drawing,
//! table and conditional-formatting xforms.
//!
//! Objects ExcelJS shares between several places (a cell style model used
//! by every cell with that style id, the font inside it, a data validation
//! covering a range, a rich shared string, a dxf) are emitted once in
//! [`ReadOutput::refs`] and referenced as `{"$ref": n}`; the JS side restores
//! the sharing (`crates/excel-rs-node/lib/xlsx/codec.js`).

use std::collections::HashMap;
use std::io::{Cursor, Read};

use serde_json::{json, Map, Value};
use zip::ZipArchive;

use crate::error::{Error, Result};
use crate::jsv::{as_f64, as_str, get, to_js_string, truthy};
use crate::model_json::{buffer, excel_to_date};
use crate::rows::{CellModel, RowModel, Rows};
use crate::xform::book::{reconcile_defined_names, strip_target_prefix, WorkbookXform};
use crate::xform::cell::vt;
use crate::xform::comment::{CommentsXform, VmlNotesXform};
use crate::xform::core::{AppXform, CoreXform, RelationshipsXform};
use crate::xform::drawing::drawing as drawing_xform;
use crate::xform::parse_part;
use crate::xform::strings::SharedStringsXform;
use crate::xform::style::{StylePart, Styles, StylesParser};
use crate::xform::table::TableXform;
use crate::xform::worksheet::{SheetOptions, WorksheetXform};

const REL_COMMENTS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const REL_VML: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";

/// Options of `workbook.xlsx.load(data, options)` used by the reader.
#[derive(Debug, Clone, Default)]
pub struct ReadOptions {
    /// `maxRows`: more rows than this is an error.
    pub max_rows: Option<usize>,
    /// `maxCols`: more cells in a row than this is an error.
    pub max_cols: Option<usize>,
    /// `ignoreNodes`: worksheet elements to skip.
    pub ignore_nodes: Vec<String>,
}

/// The reader's result: the shared-object table and the model.
#[derive(Debug)]
pub struct ReadOutput {
    /// Each worksheet's `rows`, typed (the model's worksheets have none).
    pub rows: crate::rows::SheetRows,
    /// Shared objects; entry `n` only refers to entries `< n`.
    pub refs: Vec<Value>,
    /// The model ExcelJS hands to `workbook.model = ...`, with the top-level
    /// `comments` and `tables` part maps left out (the setter ignores them).
    pub model: Value,
}

/// The shared-object table.
#[derive(Debug, Default)]
struct Refs {
    entries: Vec<Value>,
}

impl Refs {
    fn add(&mut self, v: Value) -> Value {
        self.entries.push(v);
        json!({"$ref": self.entries.len() - 1})
    }
}

/// `n` of a `{"$ref": n}` value.
fn ref_index(v: &Value) -> Option<u32> {
    let o = v.as_object().filter(|o| o.len() == 1)?;
    o.get("$ref")
        .and_then(Value::as_u64)
        .and_then(|n| u32::try_from(n).ok())
}

/// `utils.isDateFmt(fmt)`.
pub(crate) fn is_date_fmt(fmt: Option<&Value>) -> bool {
    if !truthy(fmt) {
        return false;
    }
    let fmt = to_js_string(fmt);
    let mut cleaned = String::new();
    let mut chars = fmt.chars().peekable();
    // remove [..] then "..."
    let mut tmp = String::new();
    while let Some(c) = chars.next() {
        if c == '[' {
            let rest: String = chars.clone().collect();
            if let Some(end) = rest.find(']') {
                for _ in 0..=rest.get(..end).map_or(0, |s| s.chars().count()) {
                    chars.next();
                }
                continue;
            }
        }
        tmp.push(c);
    }
    let mut chars = tmp.chars().peekable();
    while let Some(c) = chars.next() {
        if c == '"' {
            let rest: String = chars.clone().collect();
            if let Some(end) = rest.find('"') {
                for _ in 0..=rest.get(..end).map_or(0, |s| s.chars().count()) {
                    chars.next();
                }
                continue;
            }
        }
        cleaned.push(c);
    }
    cleaned
        .chars()
        .any(|c| matches!(c, 'y' | 'm' | 'd' | 'h' | 'M' | 's' | 'b'))
}

/// Style lookups with shared-object references.
/// What a cell takes from its style id (cached per id by
/// [`StyleRefs::cell_style`]).
#[derive(Clone)]
struct CellStyle {
    /// The style as a shared-object reference, or else its value (`None`:
    /// no style).
    ref_idx: Option<u32>,
    value: Option<Value>,
    is_date: bool,
}

struct StyleRefs<'a> {
    styles: Option<&'a Styles>,
    by_number: HashMap<u64, CellStyle>,
    by_id: HashMap<String, (Value, Option<Value>)>,
    parts: HashMap<(String, String), Value>,
    dxfs: HashMap<String, Option<Value>>,
}

impl<'a> StyleRefs<'a> {
    fn new(styles: Option<&'a Styles>) -> Self {
        StyleRefs {
            styles,
            by_number: HashMap::new(),
            by_id: HashMap::new(),
            parts: HashMap::new(),
            dxfs: HashMap::new(),
        }
    }

    /// `options.styles.getStyleModel(id)`: (`null` or a reference, numFmt).
    /// [`style`](Self::style) as a cell uses it; `id` is neither 0 nor NaN.
    fn cell_style(&mut self, id: f64, refs: &mut Refs) -> std::result::Result<CellStyle, String> {
        if let Some(hit) = self.by_number.get(&id.to_bits()) {
            return Ok(hit.clone());
        }
        let (style, num_fmt) = self.style(id, refs)?;
        let has_style = truthy(Some(&style));
        let cs = CellStyle {
            ref_idx: if has_style { ref_index(&style) } else { None },
            is_date: has_style && is_date_fmt(num_fmt.as_ref()),
            value: (has_style && ref_index(&style).is_none()).then_some(style),
        };
        self.by_number.insert(id.to_bits(), cs.clone());
        Ok(cs)
    }

    fn style(
        &mut self,
        id: f64,
        refs: &mut Refs,
    ) -> std::result::Result<(Value, Option<Value>), String> {
        let styles = self.styles.ok_or_else(|| {
            "Cannot read properties of undefined (reading 'getStyleModel')".to_owned()
        })?;
        let key = crate::jsv::number_to_string(id);
        if let Some(hit) = self.by_id.get(&key) {
            return Ok(hit.clone());
        }
        let Some(parts) = styles.style_parts(id)? else {
            // `style` is undefined: null, and not cached by ExcelJS either.
            return Ok((Value::Null, None));
        };
        let mut model = Map::new();
        let mut num_fmt = None;
        for (name, part) in parts {
            let v = match part {
                StylePart::Value(v) => {
                    if name == "numFmt" {
                        num_fmt = Some(v.clone());
                    }
                    v
                }
                StylePart::Shared(group, idx) => {
                    let k = (group.to_owned(), idx.clone());
                    if let Some(r) = self.parts.get(&k) {
                        r.clone()
                    } else {
                        let v = styles.part(group, &idx).cloned().unwrap_or(Value::Null);
                        let r = refs.add(v);
                        self.parts.insert(k, r.clone());
                        r
                    }
                }
            };
            model.insert(name.to_owned(), v);
        }
        let r = refs.add(Value::Object(model));
        self.by_id.insert(key, (r.clone(), num_fmt.clone()));
        Ok((r, num_fmt))
    }

    /// `options.styles.getDxfStyle(id)`; `None` is undefined.
    fn dxf(&mut self, id: &Value, refs: &mut Refs) -> std::result::Result<Option<Value>, String> {
        let styles = self.styles.ok_or_else(|| {
            "Cannot read properties of undefined (reading 'getDxfStyle')".to_owned()
        })?;
        if get(Some(&styles.model), "dxfs").is_none() {
            return Err("Cannot read properties of undefined (reading 'dxfs')".into());
        }
        let key = to_js_string(Some(id));
        if let Some(hit) = self.dxfs.get(&key) {
            return Ok(hit.clone());
        }
        let r = styles.dxf(id).cloned().map(|d| refs.add(d));
        self.dxfs.insert(key, r.clone());
        Ok(r)
    }
}

/// `getStyleModel(id)` for every cell style of a `styles.xml` part, with
/// shared parts as references (the streaming reader's `styles: 'cache'`).
pub(crate) fn style_models(part: &str, bytes: &[u8]) -> Result<ReadOutput> {
    let m = parse_part(&mut StylesParser::new(), part, bytes)?.unwrap_or(Value::Null);
    let count = get(Some(&m), "styles")
        .and_then(Value::as_array)
        .map_or(0, Vec::len);
    let styles = Styles::new(m);
    let mut style_refs = StyleRefs::new(Some(&styles));
    let mut refs = Refs::default();
    let mut list = Vec::with_capacity(count);
    for id in 0..count {
        #[allow(clippy::cast_precision_loss)]
        let (r, _) = style_refs
            .style(id as f64, &mut refs)
            .map_err(|e| Error::invalid(part, e))?;
        list.push(r);
    }
    Ok(ReadOutput {
        rows: crate::rows::SheetRows::default(),
        refs: refs.entries,
        model: Value::Array(list),
    })
}

/// Part names and contents, in ZIP order (JSZip's `zip.files` order).
fn entries(data: &[u8]) -> Result<Vec<(usize, String)>> {
    let mut zip = ZipArchive::new(Cursor::new(data))?;
    let mut out: Vec<(usize, String)> = Vec::with_capacity(zip.len());
    for i in 0..zip.len() {
        let entry = zip.by_index_raw(i)?;
        if entry.is_dir() {
            continue;
        }
        let Ok(raw) = entry.name() else { continue };
        let name = raw.strip_prefix('/').unwrap_or(&raw).to_owned();
        // JSZip's `files` is keyed by name: a later duplicate replaces.
        if let Some(pos) = out.iter().position(|(_, n)| *n == name) {
            out.remove(pos);
        }
        out.push((i, name));
    }
    Ok(out)
}

/// The content of entry `index`.
fn entry_bytes(zip: &mut ZipArchive<Cursor<&[u8]>>, index: usize) -> Result<Vec<u8>> {
    let mut bytes = Vec::new();
    zip.by_index(index)?.read_to_end(&mut bytes)?;
    Ok(bytes)
}

/// `entry.async('string')`: UTF-8 with replacement characters; and whether
/// the bytes were UTF-8 already.
fn text(bytes: Vec<u8>) -> (Vec<u8>, bool) {
    match String::from_utf8(bytes) {
        Ok(s) => (s.into_bytes(), true),
        Err(e) => (
            String::from_utf8_lossy(e.as_bytes())
                .into_owned()
                .into_bytes(),
            false,
        ),
    }
}

/// `s.match(/<prefix>(<class>+)<suffix>/)` style helpers.
fn capture_after(
    name: &str,
    prefix: &str,
    class: fn(char) -> bool,
    suffix: &str,
    allow_any_dot: bool,
) -> Option<String> {
    for (at, _) in name.match_indices(prefix) {
        let rest = name.get(at + prefix.len()..)?;
        let n = rest.chars().take_while(|c| class(*c)).count();
        if n == 0 {
            continue;
        }
        let cap: String = rest.chars().take(n).collect();
        let after = rest.get(cap.len()..)?;
        let ok = if allow_any_dot {
            // suffix like ".xml.rels" where the second '.' is any char
            suffix_matches_any_dot(after, suffix)
        } else {
            after.starts_with(suffix)
        };
        if ok {
            return Some(cap);
        }
    }
    None
}

fn suffix_matches_any_dot(s: &str, pattern: &str) -> bool {
    // pattern: literal chars, where '?' means any single char
    let mut sc = s.chars();
    for p in pattern.chars() {
        match sc.next() {
            Some(c) if p == '?' || c == p => {}
            _ => return false,
        }
    }
    true
}

fn is_alnum(c: char) -> bool {
    c.is_ascii_alphanumeric()
}

fn is_digit(c: char) -> bool {
    c.is_ascii_digit()
}

/// `/xl\/media\/([a-zA-Z0-9]+[.][a-zA-Z0-9]{3,4})$/`.
fn media_name(name: &str) -> Option<String> {
    let at = name.rfind("xl/media/")?;
    let rest = name.get(at + "xl/media/".len()..)?;
    let (base, ext) = rest.rsplit_once('.')?;
    let ok = !base.is_empty()
        && base.chars().all(is_alnum)
        && (3..=4).contains(&ext.len())
        && ext.chars().all(is_alnum);
    ok.then(|| rest.to_owned())
}

/// `rel.Target.match(/.*\/media\/(.+[.][a-zA-Z]{3,4})/)[1]`.
fn media_target(target: &str) -> Option<String> {
    let mut starts: Vec<usize> = target
        .match_indices("/media/")
        .map(|(i, _)| i + "/media/".len())
        .collect();
    starts.reverse();
    for start in starts {
        let rest = target.get(start..)?;
        let dots: Vec<usize> = rest
            .match_indices('.')
            .map(|(i, _)| i)
            .filter(|i| *i >= 1)
            .collect();
        for d in dots.into_iter().rev() {
            let letters = rest
                .get(d + 1..)?
                .chars()
                .take_while(char::is_ascii_alphabetic)
                .count();
            if letters >= 3 {
                return rest.get(..d + 1 + letters.min(4)).map(str::to_owned);
            }
        }
    }
    None
}

fn part_error(part: &str) -> impl Fn(String) -> Error + '_ {
    move |m| Error::invalid(part, m)
}

/// Reads an `.xlsx` package.
pub fn read(data: &[u8], options: &ReadOptions) -> Result<ReadOutput> {
    let mut head = None;
    let mut refs = Vec::new();
    let mut sheets: Vec<Rows> = Vec::new();
    read_events(data, options, &mut |event| {
        match event {
            ReadEvent::Model {
                refs: r,
                model,
                shared_strings,
            } => {
                refs = r;
                head = Some((model, shared_strings));
            }
            ReadEvent::Rows {
                sheet,
                refs: r,
                count: n,
                bytes: b,
                ..
            } => {
                refs.extend(r);
                if sheets.len() <= sheet {
                    sheets.resize_with(sheet + 1, Rows::default);
                }
                if let Some(rows) = sheets.get_mut(sheet) {
                    rows.count += n;
                    rows.bytes.extend_from_slice(&b);
                }
            }
        }
        true
    })?;
    let (model, shared_strings) = head.ok_or_else(|| Error::invalid("workbook", "no model"))?;
    Ok(ReadOutput {
        rows: crate::rows::SheetRows(sheets, shared_strings),
        refs,
        model,
    })
}

/// What [`read_events`] produces, in this order: the model (without the
/// worksheets' rows), then each worksheet's rows in batches.
#[derive(Debug)]
pub enum ReadEvent {
    /// The shared-object table so far, the model, and the plain shared
    /// strings that rows refer to by index.
    Model {
        refs: Vec<Value>,
        model: Value,
        shared_strings: Vec<String>,
    },
    /// Rows of worksheet `sheet` (index in `model.worksheets`), encoded as
    /// `SheetRows::encode` lays out a sheet's rows, after the shared objects
    /// `refs` added since the last event; `last` ends the worksheet.
    Rows {
        sheet: usize,
        refs: Vec<Value>,
        count: usize,
        bytes: Vec<u8>,
        last: bool,
    },
}

/// Bytes of encoded rows per [`ReadEvent::Rows`] batch.
const ROWS_BATCH: usize = 1 << 20;

/// Reads `data`, giving the model as soon as it is complete and then the
/// rows as they are read, so that the caller can build its objects
/// meanwhile. `emit` returns false to stop.
pub fn read_events(
    data: &[u8],
    options: &ReadOptions,
    emit: &mut dyn FnMut(ReadEvent) -> bool,
) -> Result<()> {
    let sheet_options = SheetOptions {
        max_rows: options.max_rows,
        max_cols: options.max_cols,
        ignore_nodes: options.ignore_nodes.clone(),
    };
    let mut worksheets: Vec<Value> = Vec::new();
    // `rows` of each worksheet model (see `rows.rs`)
    let mut sheet_rows: Vec<SheetSource> = Vec::new();
    let mut worksheet_hash: HashMap<String, usize> = HashMap::new();
    let mut worksheet_rels: HashMap<String, Value> = HashMap::new();
    let mut themes = Map::new();
    let mut media: Vec<Value> = Vec::new();
    let mut media_index: HashMap<String, usize> = HashMap::new();
    let mut drawings: Vec<(String, Value)> = Vec::new();
    let mut drawing_rels: HashMap<String, Value> = HashMap::new();
    let mut comments: HashMap<String, Value> = HashMap::new();
    let mut tables: Vec<(String, Value)> = Vec::new();
    let mut vml_drawings: HashMap<String, Value> = HashMap::new();
    let mut top = Map::new();
    let mut workbook_rels: Option<Value> = None;
    let mut sheets: Option<Value> = None;
    let mut defined_names: Option<Value> = None;
    let mut shared_strings: Option<Vec<Value>> = None;
    let mut styles: Option<Styles> = None;

    let mut zip = ZipArchive::new(Cursor::new(data))?;
    for (index, name) in entries(data)? {
        let bytes = entry_bytes(&mut zip, index)?;
        let is_media = name.contains("xl/media/");
        let theme = capture_after(&name, "xl/theme/", is_alnum, ".xml", false);
        if is_media || theme.is_some() {
            if let Some(t) = theme {
                themes.insert(
                    t,
                    Value::String(String::from_utf8_lossy(&bytes).into_owned()),
                );
            } else if let Some(file) = media_name(&name) {
                if let Some(dot) = file.rfind('.').filter(|d| *d >= 1) {
                    let (base, ext) = (
                        file.get(..dot).unwrap_or(""),
                        file.get(dot + 1..).unwrap_or(""),
                    );
                    media_index.insert(file.clone(), media.len());
                    media_index.insert(base.to_owned(), media.len());
                    media.push(json!({"type": "image", "name": base, "extension": ext, "buffer": buffer(&bytes)}));
                }
            }
            continue;
        }
        let (bytes, utf8) = text(bytes);
        match name.as_str() {
            "_rels/.rels" => {
                parse_part(&mut RelationshipsXform::default(), &name, &bytes)?;
            }
            "xl/workbook.xml" => {
                let wb =
                    parse_part(&mut WorkbookXform::new(), &name, &bytes)?.unwrap_or(Value::Null);
                sheets = get(Some(&wb), "sheets").cloned();
                defined_names = get(Some(&wb), "definedNames").cloned();
                for key in ["views", "properties", "calcProperties"] {
                    match get(Some(&wb), key) {
                        Some(v) => top.insert(key.into(), v.clone()),
                        None => top.remove(key),
                    };
                }
            }
            "xl/_rels/workbook.xml.rels" => {
                workbook_rels = parse_part(&mut RelationshipsXform::default(), &name, &bytes)?
            }
            "xl/sharedStrings.xml" => {
                let mut x = SharedStringsXform::new();
                shared_strings = Some(match parse_part(&mut x, &name, &bytes)? {
                    Some(Value::Array(a)) => a,
                    _ => Vec::new(),
                });
            }
            "xl/styles.xml" => {
                let m = parse_part(&mut StylesParser::new(), &name, &bytes)?.unwrap_or(Value::Null);
                styles = Some(Styles::new(m));
            }
            "docProps/app.xml" => {
                let m = parse_part(&mut AppXform::new(), &name, &bytes)?;
                for key in ["company", "manager"] {
                    match get(m.as_ref(), key) {
                        Some(v) => top.insert(key.into(), v.clone()),
                        None => top.remove(key),
                    };
                }
            }
            "docProps/core.xml" => {
                if let Some(Value::Object(m)) = parse_part(&mut CoreXform::new(), &name, &bytes)? {
                    for (k, v) in m {
                        top.insert(k, v);
                    }
                }
            }
            _ => {
                if let Some(n) =
                    capture_after(&name, "xl/worksheets/sheet", is_digit, ".xml", false)
                {
                    let mut xform = WorksheetXform::new(&sheet_options);
                    let fast = if sheet_options.is_default() {
                        crate::sheet_data::fast_sheet_data(&bytes)
                    } else {
                        None
                    };
                    let parsed = match &fast {
                        Some(f) => parse_part(&mut xform, &name, &f.rest),
                        None => parse_part(&mut xform, &name, &bytes),
                    };
                    // An error is reported from the whole part, so that its
                    // position is the general parser's.
                    let (parsed, fast) = match parsed {
                        Err(_) if fast.is_some() => {
                            xform = WorksheetXform::new(&sheet_options);
                            (parse_part(&mut xform, &name, &bytes), None)
                        }
                        other => (other, fast),
                    };
                    let mut ws = parsed?.unwrap_or(json!({}));
                    if let Value::Object(o) = &mut ws {
                        o.insert("sheetNo".into(), json!(n));
                    }
                    worksheet_hash.insert(name.clone(), worksheets.len());
                    worksheets.push(ws);
                    sheet_rows.push(match fast {
                        // read again from the archive when its rows are
                        // needed, rather than kept
                        Some(f) if utf8 => SheetSource::FastStream {
                            index,
                            rows_at: f.rows_at,
                        },
                        Some(f) => SheetSource::Fast {
                            doc: bytes,
                            rows_at: f.rows_at,
                        },
                        None => SheetSource::Typed(xform.take_rows()),
                    });
                } else if let Some(n) = capture_after(
                    &name,
                    "xl/worksheets/_rels/sheet",
                    is_digit,
                    ".xml?rels",
                    true,
                ) {
                    let rels = parse_part(&mut RelationshipsXform::default(), &name, &bytes)?
                        .unwrap_or(Value::Null);
                    worksheet_rels.insert(n, rels);
                } else if let Some(n) =
                    capture_after(&name, "xl/drawings/", is_alnum, ".xml", false)
                {
                    let d = parse_part(&mut drawing_xform(), &name, &bytes)?.unwrap_or(Value::Null);
                    drawings.retain(|(k, _)| *k != n);
                    drawings.push((n, d));
                } else if let Some(n) = capture_after(&name, "xl/comments", is_digit, ".xml", false)
                {
                    let c = parse_part(&mut CommentsXform::new(), &name, &bytes)?
                        .unwrap_or(Value::Null);
                    comments.insert(format!("../comments{n}.xml"), c);
                } else if let Some(n) =
                    capture_after(&name, "xl/tables/table", is_digit, ".xml", false)
                {
                    let t =
                        parse_part(&mut TableXform::new(), &name, &bytes)?.unwrap_or(Value::Null);
                    let key = format!("../tables/table{n}.xml");
                    tables.retain(|(k, _)| *k != key);
                    tables.push((key, t));
                } else if let Some(n) =
                    capture_after(&name, "xl/drawings/_rels/", is_alnum, ".xml.rels", false)
                {
                    let r = parse_part(&mut RelationshipsXform::default(), &name, &bytes)?
                        .unwrap_or(Value::Null);
                    drawing_rels.insert(n, r);
                } else if let Some(n) =
                    capture_after(&name, "xl/drawings/vmlDrawing", is_digit, ".vml", false)
                {
                    let v = parse_part(&mut VmlNotesXform::new(), &name, &bytes)?
                        .unwrap_or(Value::Null);
                    vml_drawings.insert(format!("../drawings/vmlDrawing{n}.vml"), v);
                }
            }
        }
    }

    let mut refs = Refs::default();
    let mut style_refs = StyleRefs::new(styles.as_ref());
    let date1904 = truthy(get(top.get("properties"), "date1904"));

    // ---- WorkbookXform.reconcile
    let rels: HashMap<String, Value> = workbook_rels
        .as_ref()
        .and_then(Value::as_array)
        .into_iter()
        .flatten()
        .map(|r| (to_js_string(get(Some(r), "Id")), r.clone()))
        .collect();
    let mut referenced: Vec<usize> = Vec::new();
    for sheet in sheets
        .as_ref()
        .and_then(Value::as_array)
        .into_iter()
        .flatten()
    {
        let Some(rel) = rels.get(&to_js_string(get(Some(sheet), "rId"))) else {
            continue;
        };
        let target = as_str(get(Some(rel), "Target")).ok_or_else(|| {
            Error::invalid(
                "xl/_rels/workbook.xml.rels",
                "Cannot read properties of undefined (reading 'replace')",
            )
        })?;
        let key = format!("xl/{}", strip_target_prefix(target));
        if let Some(&idx) = worksheet_hash.get(&key) {
            if let Some(Value::Object(ws)) = worksheets.get_mut(idx) {
                for (k, src) in [("name", "name"), ("id", "id"), ("state", "state")] {
                    match get(Some(sheet), src) {
                        Some(v) => ws.insert(k.into(), v.clone()),
                        None => ws.remove(k),
                    };
                }
            }
            referenced.push(idx);
        }
    }
    let kept_names = reconcile_defined_names(defined_names.as_ref(), &mut worksheets, &referenced)
        .map_err(part_error("xl/workbook.xml"))?;
    for (i, m) in media.iter_mut().enumerate() {
        if let Value::Object(o) = m {
            o.insert("index".into(), json!(i));
        }
    }

    // ---- drawings
    let mut drawing_media: HashMap<String, Vec<Value>> = HashMap::new();
    for (name, drawing) in &mut drawings {
        drawing_media.insert(name.clone(), Vec::new());
        if let Some(drawing_rel) = drawing_rels.get(name.as_str()) {
            let drels: HashMap<String, &Value> = drawing_rel
                .as_array()
                .into_iter()
                .flatten()
                .map(|r| (to_js_string(get(Some(r), "Id")), r))
                .collect();
            let mut images = Vec::new();
            for anchor in drawing
                .get_mut("anchors")
                .and_then(Value::as_array_mut)
                .into_iter()
                .flatten()
            {
                if let Some(Value::Object(h)) = anchor.pointer_mut("/picture/hyperlinks") {
                    let rid = to_js_string(h.get("rId"));
                    if let Some(rel) = drels.get(&rid) {
                        h.insert(
                            "hyperlink".into(),
                            get(Some(rel), "Target").cloned().unwrap_or(Value::Null),
                        );
                        h.remove("rId");
                    }
                }
                // reconcilePicture
                let picture = get(Some(anchor), "picture");
                let mut medium: Option<usize> = None;
                if truthy(picture) && truthy(get(picture, "rId")) {
                    let rid = to_js_string(get(picture, "rId"));
                    let rel = drels.get(&rid).ok_or_else(|| {
                        Error::invalid(
                            &format!("xl/drawings/{name}.xml"),
                            "Cannot read properties of undefined (reading 'Target')",
                        )
                    })?;
                    if let Some(file) = media_target(&to_js_string(get(Some(rel), "Target"))) {
                        medium = media_index.get(&file).copied().filter(|i| *i < media.len());
                    }
                }
                if let Some(index) = medium {
                    let mut image = Map::new();
                    image.insert("type".into(), json!("image"));
                    image.insert("imageId".into(), json!(index));
                    if let Some(r) = anchor.get("range") {
                        image.insert("range".into(), r.clone());
                    }
                    if let Some(h) = anchor.pointer("/picture/hyperlinks") {
                        image.insert("hyperlinks".into(), h.clone());
                    }
                    images.push(Value::Object(image));
                }
            }
            drawing_media.insert(name.clone(), images);
        }
    }

    // ---- tables
    for (key, table) in &mut tables {
        for column in table
            .get_mut("columns")
            .and_then(Value::as_array_mut)
            .into_iter()
            .flatten()
        {
            if let Some(dxf_id) = column.get("dxfId").cloned() {
                match style_refs
                    .dxf(&dxf_id, &mut refs)
                    .map_err(part_error(key))?
                {
                    Some(r) => {
                        column["style"] = r;
                    }
                    None => {
                        if let Value::Object(o) = column {
                            o.remove("style");
                        }
                    }
                }
            }
        }
    }
    let table_map: HashMap<String, Value> = tables.into_iter().collect();

    // ---- worksheets
    let sst_refs: Vec<Option<Value>> = match &shared_strings {
        Some(values) => values
            .iter()
            .map(|v| matches!(v, Value::Object(_)).then(|| refs.add(v.clone())))
            .collect(),
        None => Vec::new(),
    };
    let ctx = SheetCtx {
        shared_strings: shared_strings.as_deref(),
        sst_refs: &sst_refs,
        date1904,
        drawing_media: &drawing_media,
        media_index: &media_index,
        comments: &comments,
        vml_drawings: &vml_drawings,
        tables: &table_map,
    };
    let mut row_jobs = Vec::with_capacity(sheet_rows.len());
    for (ws, source) in worksheets.iter_mut().zip(sheet_rows) {
        let sheet_no = to_js_string(get(Some(ws), "sheetNo"));
        let rels = worksheet_rels.get(&sheet_no).cloned();
        let part = format!("xl/worksheets/sheet{sheet_no}.xml");
        let row_ctx = reconcile_worksheet(ws, rels.as_ref(), &ctx, &mut style_refs, &mut refs)
            .map_err(part_error(&part))?;
        row_jobs.push((part, source, row_ctx));
    }
    let mut model = Map::new();
    model.insert("worksheets".into(), Value::Array(worksheets));
    model.insert("themes".into(), Value::Object(themes));
    model.insert("media".into(), Value::Array(media));
    if let Some(s) = sheets {
        model.insert("sheets".into(), s);
    }
    model.insert("definedNames".into(), Value::Array(kept_names));
    for (k, v) in top {
        model.insert(k, v);
    }
    let shared_strings_plain = shared_strings
        .iter()
        .flatten()
        .map(|v| match v {
            Value::String(s) => s.clone(),
            _ => String::new(),
        })
        .collect();
    let mut sent = refs.entries.len();
    if !emit(ReadEvent::Model {
        refs: refs.entries.clone(),
        model: Value::Object(model),
        shared_strings: shared_strings_plain,
    }) {
        return Ok(());
    }

    // ---- rows, in batches
    for (sheet, (part, source, row_ctx)) in row_jobs.into_iter().enumerate() {
        let mut batch = Vec::new();
        let mut count = 0;
        let mut stopped = false;
        let mut send = |batch: &mut Vec<u8>, count: &mut usize, refs: &Refs, last: bool| {
            let new_refs = refs.entries.get(sent..).unwrap_or_default().to_vec();
            sent = refs.entries.len();
            emit(ReadEvent::Rows {
                sheet,
                refs: new_refs,
                count: std::mem::take(count),
                bytes: std::mem::take(batch),
                last,
            })
        };
        let mut formulae: HashMap<String, Value> = HashMap::new();
        let mut each =
            |row: &mut RowModel, refs: &mut Refs, styles: &mut StyleRefs<'_>| -> R<bool> {
                reconcile_row(row, &row_ctx, &ctx, styles, refs, &mut formulae)?;
                crate::rows::encode_row(&mut batch, row);
                count += 1;
                if batch.len() >= ROWS_BATCH {
                    return Ok(send(&mut batch, &mut count, refs, false));
                }
                Ok(true)
            };
        let result: R<()> = match source {
            SheetSource::Typed(rows) => {
                let mut result = Ok(());
                for mut row in rows {
                    match each(&mut row, &mut refs, &mut style_refs) {
                        Ok(true) => {}
                        Ok(false) => {
                            stopped = true;
                            break;
                        }
                        Err(e) => {
                            result = Err(e);
                            break;
                        }
                    }
                }
                result
            }
            SheetSource::Fast { doc, rows_at } => {
                crate::sheet_data::fast_rows(&doc, rows_at, &mut |mut row| {
                    if stopped {
                        return Ok(());
                    }
                    if !each(&mut row, &mut refs, &mut style_refs)? {
                        stopped = true;
                    }
                    Ok(())
                })
                .ok_or_else(|| "excel-rs: <sheetData> changed between reads".to_owned())
                .and_then(|r| r)
            }
            SheetSource::FastStream { index, rows_at } => match zip.by_index(index) {
                Ok(entry) => crate::sheet_data::fast_rows_stream(entry, rows_at, &mut |mut row| {
                    if stopped {
                        return Ok(());
                    }
                    if !each(&mut row, &mut refs, &mut style_refs)? {
                        stopped = true;
                    }
                    Ok(())
                })
                .ok_or_else(|| "excel-rs: <sheetData> changed between reads".to_owned())
                .and_then(|r| r),
                Err(e) => Err(e.to_string()),
            },
        };
        result.map_err(part_error(&part))?;
        if stopped || !send(&mut batch, &mut count, &refs, true) {
            return Ok(());
        }
    }
    Ok(())
}

/// A worksheet's rows between the parse and the reconcile: typed, or a part
/// whose `<sheetData>` the fast path reads (row by row, as it reconciles).
enum SheetSource {
    Typed(Vec<RowModel>),
    Fast {
        doc: Vec<u8>,
        rows_at: usize,
    },
    /// The part is entry `index`, inflated again for its rows.
    FastStream {
        index: usize,
        rows_at: usize,
    },
}

struct SheetCtx<'a> {
    shared_strings: Option<&'a [Value]>,
    sst_refs: &'a [Option<Value>],
    date1904: bool,
    drawing_media: &'a HashMap<String, Vec<Value>>,
    media_index: &'a HashMap<String, usize>,
    comments: &'a HashMap<String, Value>,
    vml_drawings: &'a HashMap<String, Value>,
    tables: &'a HashMap<String, Value>,
}

type R<T> = std::result::Result<T, String>;

/// `WorkSheetXform.reconcile(model, options)`.
fn reconcile_worksheet(
    ws: &mut Value,
    relationships: Option<&Value>,
    ctx: &SheetCtx<'_>,
    styles: &mut StyleRefs<'_>,
    refs: &mut Refs,
) -> R<RowCtx> {
    let Value::Object(m) = ws else {
        return Ok(RowCtx::default());
    };
    let mut rels: HashMap<String, Value> = HashMap::new();
    for rel in relationships
        .and_then(Value::as_array)
        .into_iter()
        .flatten()
    {
        rels.insert(to_js_string(get(Some(rel), "Id")), rel.clone());
        let ty = as_str(get(Some(rel), "Type"));
        let target = to_js_string(get(Some(rel), "Target"));
        if ty == Some(REL_COMMENTS) {
            let c = ctx.comments.get(&target).ok_or_else(|| {
                "Cannot read properties of undefined (reading 'comments')".to_owned()
            })?;
            match get(Some(c), "comments") {
                Some(v) => m.insert("comments".into(), v.clone()),
                None => m.remove("comments"),
            };
        }
        if ty == Some(REL_VML)
            && m.get("comments")
                .and_then(Value::as_array)
                .is_some_and(|a| !a.is_empty())
        {
            let vml = ctx.vml_drawings.get(&target).ok_or_else(|| {
                "Cannot read properties of undefined (reading 'comments')".to_owned()
            })?;
            let vml_comments = get(Some(vml), "comments").cloned();
            if let Some(Value::Array(list)) = m.get_mut("comments") {
                for (i, comment) in list.iter_mut().enumerate() {
                    let mut note = match comment.get("note") {
                        Some(Value::Object(o)) => o.clone(),
                        _ => Map::new(),
                    };
                    if let Some(Value::Object(v)) = get(vml_comments.as_ref(), &i.to_string()) {
                        for (k, val) in v {
                            note.insert(k.clone(), val.clone());
                        }
                    }
                    if let Value::Object(c) = comment {
                        c.insert("note".into(), Value::Object(note));
                    }
                }
            }
        }
    }
    let mut comments_map: HashMap<String, Value> = HashMap::new();
    for c in m
        .get("comments")
        .and_then(Value::as_array)
        .into_iter()
        .flatten()
    {
        if truthy(get(Some(c), "ref")) {
            comments_map.insert(to_js_string(get(Some(c), "ref")), c.clone());
        }
    }
    let mut hyperlink_map: HashMap<String, Value> = HashMap::new();
    for h in m
        .get("hyperlinks")
        .and_then(Value::as_array)
        .into_iter()
        .flatten()
    {
        if truthy(get(Some(h), "rId")) {
            let rel = rels
                .get(&to_js_string(get(Some(h), "rId")))
                .ok_or_else(|| {
                    "Cannot read properties of undefined (reading 'Target')".to_owned()
                })?;
            hyperlink_map.insert(
                to_js_string(get(Some(h), "address")),
                get(Some(rel), "Target").cloned().unwrap_or(Value::Null),
            );
        }
    }
    // compact the rows and cells: null rows and cells are not kept by the
    // typed parser; the model's `rows` placeholder goes.
    m.remove("rows");

    // cols
    if let Some(Value::Array(cols)) = m.get_mut("cols") {
        for col in cols {
            if let Some(id) = as_f64(get(Some(col), "styleId")).filter(|v| *v != 0.0 && !v.is_nan())
            {
                col["style"] = styles.style(id, refs)?.0;
            }
        }
    }

    // conditional formattings
    if let Some(Value::Array(cfs)) = m.get_mut("conditionalFormattings") {
        for cf in cfs {
            for rule in cf
                .get_mut("rules")
                .and_then(Value::as_array_mut)
                .into_iter()
                .flatten()
            {
                if let Some(dxf_id) = rule.get("dxfId").cloned() {
                    let style = styles.dxf(&dxf_id, refs)?;
                    if let Value::Object(r) = rule {
                        match style {
                            Some(s) => r.insert("style".into(), s),
                            None => r.remove("style"),
                        };
                        r.remove("dxfId");
                    }
                }
            }
        }
    }

    // data validations: shared objects
    if let Some(Value::Object(dv)) = m.get("dataValidations").cloned() {
        let dvs: Vec<Value> = dv
            .get("$dvs")
            .and_then(Value::as_array)
            .cloned()
            .unwrap_or_default();
        let ids: Vec<Value> = dvs.into_iter().map(|d| refs.add(d)).collect();
        let mut out = Map::new();
        for (addr, idx) in dv
            .get("map")
            .and_then(Value::as_object)
            .into_iter()
            .flatten()
        {
            let i = idx
                .as_u64()
                .and_then(|i| usize::try_from(i).ok())
                .unwrap_or(0);
            out.insert(addr.clone(), ids.get(i).cloned().unwrap_or(Value::Null));
        }
        m.insert("dataValidations".into(), Value::Object(out));
    }

    // media (drawing images and background)
    let mut media = Vec::new();
    if let Some(drawing) = m.get("drawing").filter(|d| truthy(Some(d))) {
        let rel = rels
            .get(&to_js_string(get(Some(drawing), "rId")))
            .ok_or_else(|| "Cannot read properties of undefined (reading 'Target')".to_owned())?;
        let target = to_js_string(get(Some(rel), "Target"));
        if let Some(name) = drawing_name(&target) {
            let images = ctx.drawing_media.get(&name).ok_or_else(|| {
                "Cannot read properties of undefined (reading 'anchors')".to_owned()
            })?;
            media.extend(images.iter().cloned());
        }
    }
    if let Some(bg) = m.get("background").filter(|b| truthy(Some(b))) {
        if let Some(rel) = rels.get(&to_js_string(get(Some(bg), "rId"))) {
            let target = to_js_string(get(Some(rel), "Target"));
            let file = target
                .split("/media/")
                .nth(1)
                .unwrap_or("undefined")
                .to_owned();
            if let Some(&id) = ctx.media_index.get(&file) {
                media.push(json!({"type": "background", "imageId": id}));
            }
        }
    }
    m.insert("media".into(), Value::Array(media));

    // tables
    let mut tables = Vec::new();
    for part in m
        .get("tables")
        .and_then(Value::as_array)
        .into_iter()
        .flatten()
    {
        let rel = rels
            .get(&to_js_string(get(Some(part), "rId")))
            .ok_or_else(|| "Cannot read properties of undefined (reading 'Target')".to_owned())?;
        tables.push(
            ctx.tables
                .get(&to_js_string(get(Some(rel), "Target")))
                .cloned()
                .unwrap_or(Value::Null),
        );
    }
    m.insert("tables".into(), Value::Array(tables));

    m.remove("relationships");
    m.remove("hyperlinks");
    m.remove("comments");
    Ok(RowCtx {
        hyperlinks: hyperlink_map,
        comments: comments_map,
    })
}

/// What a worksheet's cells need from its reconcile: hyperlink targets and
/// comments by address.
#[derive(Default)]
struct RowCtx {
    hyperlinks: HashMap<String, Value>,
    comments: HashMap<String, Value>,
}

/// The rows part of `WorkSheetXform.reconcile`, for one row.
fn reconcile_row(
    row: &mut RowModel,
    row_ctx: &RowCtx,
    ctx: &SheetCtx<'_>,
    styles: &mut StyleRefs<'_>,
    refs: &mut Refs,
    formulae: &mut HashMap<String, Value>,
) -> R<()> {
    let style = match as_f64(row.style_id.as_ref()).filter(|v| *v != 0.0 && !v.is_nan()) {
        Some(id) => styles.style(id, refs)?.0,
        None => json!({}),
    };
    row.style_value = Some(style);
    row.style_id = None;
    for cell in &mut row.cells {
        reconcile_cell(
            cell,
            ctx,
            styles,
            refs,
            &row_ctx.hyperlinks,
            &row_ctx.comments,
            formulae,
        )?;
    }
    Ok(())
}

/// `/\/drawings\/([a-zA-Z0-9]+)[.][a-zA-Z]{3,4}$/`.
fn drawing_name(target: &str) -> Option<String> {
    let at = target.rfind("/drawings/")?;
    let rest = target.get(at + "/drawings/".len()..)?;
    let (base, ext) = rest.rsplit_once('.')?;
    (!base.is_empty()
        && base.chars().all(is_alnum)
        && (3..=4).contains(&ext.len())
        && ext.chars().all(|c| c.is_ascii_alphabetic()))
    .then(|| base.to_owned())
}

/// `CellXform.reconcile(model, options)`.
fn reconcile_cell(
    c: &mut CellModel,
    ctx: &SheetCtx<'_>,
    styles: &mut StyleRefs<'_>,
    refs: &mut Refs,
    hyperlinks: &HashMap<String, Value>,
    comments: &HashMap<String, Value>,
    formulae: &mut HashMap<String, Value>,
) -> R<()> {
    let bx = |v: Value| Some(Box::new(v));
    let style_id = as_f64(c.style_id.as_deref()).filter(|v| *v != 0.0 && !v.is_nan());
    let mut is_date = false;
    if let (Some(id), Some(_)) = (style_id, styles.styles) {
        let cs = styles.cell_style(id, refs)?;
        c.style_ref = cs.ref_idx;
        if let Some(v) = cs.value {
            c.style_value = bx(v);
        }
        is_date = cs.is_date;
    }
    c.style_id = None;
    #[allow(clippy::cast_possible_truncation)]
    let ty = c.ty.map(|t| t as i64);
    match ty {
        Some(vt::STRING) => {
            if crate::jsv::is_number(c.value.as_deref()) {
                if let Some(values) = ctx.shared_strings {
                    let idx = as_f64(c.value.as_deref()).unwrap_or(f64::NAN);
                    let entry = (idx >= 0.0 && idx.fract() == 0.0)
                        .then(|| {
                            #[allow(clippy::cast_possible_truncation, clippy::cast_sign_loss)]
                            let i = idx as usize;
                            i
                        })
                        .and_then(|i| values.get(i).map(|v| (i, v)));
                    let Some((i, v)) = entry else {
                        return Err(
                            "Cannot read properties of undefined (reading 'richText')".into()
                        );
                    };
                    if truthy(get(Some(v), "richText")) {
                        c.ty = Some(8.0);
                    }
                    match (ctx.sst_refs.get(i).cloned().flatten(), v) {
                        (Some(shared), _) => c.value = bx(shared),
                        // a plain string: sent once, by index
                        (None, Value::String(_)) => {
                            c.value = None;
                            c.sst = u32::try_from(i).ok();
                        }
                        (None, v) => c.value = bx(v.clone()),
                    }
                } else {
                    // No xl/sharedStrings.xml: the string is not in the
                    // package, so the cell is empty. (ExcelJS keeps the
                    // index as the cell's "string": B6.)
                    c.ty = Some(0.0);
                    c.value = None;
                }
            } else if truthy(get(c.value.as_deref(), "richText")) {
                c.ty = Some(8.0);
            } else if c.value.is_none() {
                return Err("Cannot read properties of undefined (reading 'richText')".into());
            }
        }
        Some(vt::NUMBER) => {
            if is_date {
                let v = as_f64(c.value.as_deref()).unwrap_or(f64::NAN);
                c.ty = Some(4.0);
                c.value = bx(excel_to_date(v, ctx.date1904));
            }
        }
        Some(vt::FORMULA) => {
            if is_date {
                if let Some(r) = c.result.as_deref() {
                    let v = crate::jsv::to_number(Some(r));
                    c.result = bx(excel_to_date(v, ctx.date1904));
                }
            }
            if as_str(c.share_type.as_deref()) == Some("shared") {
                let si = to_js_string(c.si.as_deref());
                if truthy(c.ref_.as_deref()) {
                    formulae.insert(si, c.address.as_deref().cloned().unwrap_or(Value::Null));
                } else {
                    c.shared_formula = formulae.get(&si).map(|a| Box::new(a.clone()));
                    c.share_type = None;
                }
                c.si = None;
            }
        }
        _ => {}
    }
    if hyperlinks.is_empty() && comments.is_empty() {
        return Ok(());
    }
    let address = to_js_string(c.address.as_deref());
    if let Some(target) = hyperlinks.get(&address).filter(|t| truthy(Some(t))) {
        if c.ty == Some(6.0) {
            c.text = c.result.take();
        } else {
            c.text = c.value.take().or_else(|| {
                let i = usize::try_from(c.sst.take()?).ok()?;
                ctx.shared_strings?.get(i).cloned().map(Box::new)
            });
        }
        c.ty = Some(5.0);
        c.hyperlink = bx(target.clone());
    }
    if let Some(cm) = comments.get(&address).filter(|v| truthy(Some(v))) {
        c.comment = bx(cm.clone());
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn date_formats() {
        for (f, d) in [
            ("mm-dd-yy", true),
            ("0.00%", false),
            ("[Red]0.00", false),
            ("\"day\" 0", false),
            ("General", false),
            ("h:mm", true),
        ] {
            assert_eq!(is_date_fmt(Some(&json!(f))), d, "{f}");
        }
    }

    #[test]
    fn regexes() {
        assert_eq!(
            media_target("../media/image1.png").as_deref(),
            Some("image1.png")
        );
        assert_eq!(
            media_target("../media/a.b.jpeg").as_deref(),
            Some("a.b.jpeg")
        );
        assert_eq!(
            media_name("xl/media/image1.jpeg").as_deref(),
            Some("image1.jpeg")
        );
        assert_eq!(media_name("xl/media/image1.x"), None);
        assert_eq!(
            drawing_name("../drawings/drawing1.xml").as_deref(),
            Some("drawing1")
        );
        assert_eq!(
            capture_after(
                "xl/worksheets/_rels/sheet3.xml.rels",
                "xl/worksheets/_rels/sheet",
                is_digit,
                ".xml?rels",
                true
            )
            .as_deref(),
            Some("3")
        );
    }
}
