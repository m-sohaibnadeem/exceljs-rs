//! `<sheetData>` parsing into typed rows: the parse side of
//! `ListXform({tag: 'sheetData', childXform: RowXform})`, `RowXform` and
//! `CellXform`, with their exact state machines (see `xform/cell.rs` for the
//! same rules over `Value` models, which the reader no longer uses for
//! cells).

use std::cell::Cell;
use std::rc::Rc;

use serde_json::{json, Value};

use crate::js::{decode_x_escapes, parse_float, parse_int};
use crate::jsv::{num, to_js_string, truthy};
use crate::rows::{CellModel, Field, RowModel};
use crate::xform::cell::vt;
use crate::xform::simple::{float_value, int_value};
use crate::xform::strings::RichTextXform;
use crate::xform::style::parse_boolean;
use crate::xform::{XResult, Xform};
use crate::xml::Tag;

/// Rows handed from the sheetData xform to the worksheet reader.
#[derive(Clone, Default)]
pub(crate) struct SharedRows(Rc<Cell<Vec<RowModel>>>);

impl SharedRows {
    fn set(&self, rows: Vec<RowModel>) {
        self.0.set(rows);
    }
    pub fn take(&self) -> Vec<RowModel> {
        self.0.take()
    }
}

impl std::fmt::Debug for SharedRows {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str("SharedRows")
    }
}

fn bx(v: Value) -> Field {
    Some(Box::new(v))
}

#[derive(Clone, Copy, PartialEq)]
enum Node {
    F,
    V,
    T,
}

/// `CellXform` (parse role) over a typed model.
struct CellParser {
    rich: RichTextXform,
    in_rich: bool,
    t: Option<String>,
    current: Option<Node>,
    model: Option<CellModel>,
    /// Depth inside elements `CellXform` does not know (`<rPh>`,
    /// `<phoneticPr>`, ...), which are skipped whole (B4).
    skip: usize,
    /// A `<v>` or `<t>` was seen: an empty one is the empty string (B3).
    seen_value: bool,
}

impl CellParser {
    fn model(&mut self) -> XResult<&mut CellModel> {
        self.model
            .as_mut()
            .ok_or_else(|| "Cannot set properties of undefined".to_owned())
    }

    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if self.in_rich {
            self.rich.parse_open(node)?;
            return Ok(true);
        }
        if self.skip > 0 {
            self.skip += 1;
            return Ok(true);
        }
        match node.name {
            "c" => {
                self.skip = 0;
                self.seen_value = false;
                self.t = node.attr("t").map(str::to_owned);
                self.model = Some(open_cell(|n| node.attr(n)));
                Ok(true)
            }
            "f" => {
                self.current = Some(Node::F);
                open_formula(self.model()?, |n| node.attr(n));
                Ok(true)
            }
            "v" => {
                self.current = Some(Node::V);
                self.seen_value = true;
                Ok(true)
            }
            "t" => {
                self.current = Some(Node::T);
                self.seen_value = true;
                Ok(true)
            }
            "r" => {
                self.in_rich = true;
                self.rich.parse_open(node)?;
                Ok(true)
            }
            "is" => Ok(false),
            _ => {
                self.skip = 1;
                Ok(true)
            }
        }
    }

    fn parse_text(&mut self, text: &str) {
        if self.in_rich {
            self.rich.parse_text(text);
            return;
        }
        if self.skip > 0 {
            return;
        }
        let Some(m) = self.model.as_mut() else { return };
        match self.current {
            Some(Node::F) => {
                let f = match m.formula.as_deref().filter(|f| truthy(Some(f))) {
                    Some(prev) => format!("{}{text}", to_js_string(Some(prev))),
                    None => text.to_owned(),
                };
                m.formula = bx(Value::String(f));
            }
            Some(Node::V | Node::T) => {
                let is_rich = matches!(m.value.as_deref(), Some(v) if truthy(v.get("richText")));
                if !is_rich {
                    let v = match m.value.as_deref().filter(|v| truthy(Some(v))) {
                        Some(Value::String(prev)) => {
                            let mut s = String::with_capacity(prev.len() + text.len());
                            s.push_str(prev);
                            s.push_str(text);
                            s
                        }
                        Some(prev) => format!("{}{text}", to_js_string(Some(prev))),
                        None => text.to_owned(),
                    };
                    m.value = bx(Value::String(v));
                }
            }
            None => {}
        }
    }

    fn parse_close(&mut self, name: &str) -> XResult<bool> {
        if self.skip > 0 && !self.in_rich {
            self.skip -= 1;
            return Ok(true);
        }
        match name {
            "c" => {
                self.finish()?;
                Ok(false)
            }
            "f" | "v" | "is" => {
                self.current = None;
                Ok(true)
            }
            "t" => {
                if self.in_rich {
                    self.rich.parse_close(name)?;
                } else {
                    self.current = None;
                }
                Ok(true)
            }
            "r" => {
                if !self.in_rich {
                    return Err("Cannot read properties of undefined (reading 'model')".into());
                }
                let run = self.rich.model().cloned().unwrap_or(Value::Null);
                let m = self.model()?;
                match m.value.as_deref_mut() {
                    Some(v) if truthy(Some(v)) => match v {
                        Value::Object(o) => {
                            let rt = o.entry("richText").or_insert_with(|| json!([]));
                            if !truthy(Some(rt)) {
                                *rt = json!([]);
                            }
                            match rt {
                                Value::Array(a) => a.push(run),
                                _ => {
                                    return Err("model.value.richText.push is not a function".into())
                                }
                            }
                        }
                        // A string value: JS assigns `richText` to a
                        // primitive (ignored) and then reads it: TypeError.
                        _ => {
                            return Err(
                                "Cannot read properties of undefined (reading 'push')".into()
                            )
                        }
                    },
                    _ => m.value = bx(json!({ "richText": [run] })),
                }
                self.in_rich = false;
                self.current = None;
                Ok(true)
            }
            _ => {
                if self.in_rich {
                    self.rich.parse_close(name)?;
                    return Ok(true);
                }
                Ok(false)
            }
        }
    }

    /// `parseClose('c')`: the cell's first-guess type.
    fn finish(&mut self) -> XResult<()> {
        let t = self.t.take();
        let seen_value = self.seen_value;
        finish_cell(self.model()?, t.as_deref(), seen_value)
    }
}

/// `CellXform.parseOpen` of `<c>`: the new cell model.
fn open_cell<'a>(attr: impl Fn(&str) -> Option<&'a str>) -> CellModel {
    let mut m = CellModel::default();
    if let Some(r) = attr("r") {
        m.address = bx(json!(r));
    }
    if let Some(s) = attr("s").filter(|v| !v.is_empty()) {
        m.style_id = bx(int_value(Some(s)));
    }
    m
}

/// `CellXform.parseOpen` of `<f>`.
fn open_formula<'a>(m: &mut CellModel, attr: impl Fn(&str) -> Option<&'a str>) {
    m.si = attr("si").and_then(|v| bx(json!(v)));
    m.share_type = attr("t").and_then(|v| bx(json!(v)));
    m.ref_ = attr("ref").and_then(|v| bx(json!(v)));
}

/// `CellXform.parseClose('c')`: the cell's first-guess type, from its `t`
/// attribute and whether a `<v>` or `<t>` was seen.
fn finish_cell(m: &mut CellModel, t: Option<&str>, seen_value: bool) -> XResult<()> {
    match m.value.take() {
        None => finish_text(m, t, seen_value, None),
        Some(v) => match *v {
            Value::String(s) => finish_text(m, t, seen_value, Some(&s)),
            other => finish_object(m, t, other),
        },
    }
}

/// [`finish_cell`] for a cell whose value is text (or absent).
fn finish_text(
    m: &mut CellModel,
    t: Option<&str>,
    seen_value: bool,
    text: Option<&str>,
) -> XResult<()> {
    let is_formula = truthy(m.formula.as_deref()) || truthy(m.share_type.as_deref());
    // `<v></v>` in a string cell is the empty string; ExcelJS drops the
    // value (B3).
    let is_text = matches!(t, Some("str" | "inlineStr"));
    let text = match text {
        None if is_text && seen_value => {
            if is_formula {
                m.result = bx(json!(""));
                None
            } else {
                Some("")
            }
        }
        other => other,
    };
    if is_formula {
        m.ty = Some(6.0);
        if let Some(v) = text {
            if v.is_empty() {
                m.value = bx(json!(""));
            } else {
                m.result = bx(match t {
                    Some("str") => Value::String(decode_x_escapes(v)),
                    Some("b") => Value::Bool(boolean_text(v)),
                    Some("e") => json!({ "error": v }),
                    _ => num(parse_float(v)),
                });
            }
        }
    } else if let Some(v) = text {
        let (ty, value) = match t {
            Some("s") => (vt::STRING, num(parse_int(v).unwrap_or(f64::NAN))),
            Some("str" | "inlineStr") => (vt::STRING, Value::String(decode_x_escapes(v))),
            Some("b") => (vt::BOOLEAN, Value::Bool(boolean_text(v))),
            Some("e") => (vt::ERROR, json!({ "error": v })),
            _ => (vt::NUMBER, num(parse_float(v))),
        };
        #[allow(clippy::cast_precision_loss)]
        {
            m.ty = Some(ty as f64);
        }
        m.value = bx(value);
    } else if truthy(m.style_id.as_deref()) {
        m.ty = Some(0.0);
    } else {
        m.ty = Some(1.0);
    }
    Ok(())
}

/// [`finish_cell`] for a cell whose value is an object (a rich inline
/// string, then `<v>` text appended to it, as JS does).
fn finish_object(m: &mut CellModel, t: Option<&str>, v: Value) -> XResult<()> {
    let as_text = || to_js_string(Some(&v));
    let is_formula = truthy(m.formula.as_deref()) || truthy(m.share_type.as_deref());
    if is_formula {
        m.ty = Some(6.0);
        if !truthy(Some(&v)) {
            m.value = bx(v);
            return Ok(());
        }
        let result = match t {
            Some("str") => return Err("text.replace is not a function".into()),
            Some("b") => json!(boolean_text(&as_text())),
            Some("e") => json!({ "error": v }),
            _ => num(parse_float(&as_text())),
        };
        m.result = bx(result);
        return Ok(());
    }
    let (ty, value) = match t {
        Some("s") => (vt::STRING, num(parse_int(&as_text()).unwrap_or(f64::NAN))),
        Some("str") => return Err("text.replace is not a function".into()),
        Some("inlineStr") => (vt::STRING, v),
        Some("b") => (vt::BOOLEAN, json!(boolean_text(&as_text()))),
        Some("e") => (vt::ERROR, json!({ "error": v })),
        _ => (vt::NUMBER, num(parse_float(&as_text()))),
    };
    #[allow(clippy::cast_precision_loss)]
    {
        m.ty = Some(ty as f64);
    }
    m.value = bx(value);
    Ok(())
}

/// The value of a `t="b"` cell: a non-zero integer (`parseInt`) or the word
/// `true`. ExcelJS's `parseInt(text) !== 0` makes any non-numeric text,
/// `false` included, true (B10).
fn boolean_text(text: &str) -> bool {
    match parse_int(text) {
        Some(n) => n != 0.0,
        None => text.trim().eq_ignore_ascii_case("true"),
    }
}

/// `RowXform.parseOpen` of `<row>`: the new row model.
fn open_row<'a>(attr: impl Fn(&str) -> Option<&'a str>) -> RowModel {
    let (min, max) = match attr("spans").filter(|v| !v.is_empty()) {
        Some(spans) => {
            let mut parts = spans.split(':').map(|p| int_value(Some(p)));
            (parts.next(), parts.next())
        }
        None => (None, None),
    };
    let mut m = RowModel {
        number: Some(int_value(attr("r"))),
        min,
        max,
        ..RowModel::default()
    };
    if let Some(s) = attr("s").filter(|v| !v.is_empty()) {
        m.style_id = Some(int_value(Some(s)));
    }
    if parse_boolean(attr("hidden")) {
        m.hidden = Some(json!(true));
    }
    if parse_boolean(attr("bestFit")) {
        m.best_fit = Some(json!(true));
    }
    if let Some(ht) = attr("ht").filter(|v| !v.is_empty()) {
        m.height = Some(float_value(Some(ht)));
    }
    if let Some(o) = attr("outlineLevel").filter(|v| !v.is_empty()) {
        m.outline_level = Some(int_value(Some(o)));
    }
    if parse_boolean(attr("collapsed")) {
        m.collapsed = Some(json!(true));
    }
    m
}

/// What `RowXform.model` holds: nothing yet, the row being read, or a row
/// already handed to the list (ExcelJS hands over a copy and keeps
/// appending stray cells to its own; only their count is observable).
#[allow(clippy::large_enum_variant)] // one per worksheet parse
enum RowState {
    Undefined,
    Null,
    Open(RowModel),
    Phantom(usize),
}

/// `RowXform` (parse role).
struct RowParser {
    max_cols: Option<usize>,
    cell: CellParser,
    in_cell: bool,
    model: RowState,
}

impl RowParser {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if self.in_cell {
            self.cell.parse_open(node)?;
            return Ok(true);
        }
        if node.name == "row" {
            self.model = RowState::Open(open_row(|n| node.attr(n)));
            return Ok(true);
        }
        if node.name == "c" {
            self.in_cell = true;
            self.cell.parse_open(node)?;
            return Ok(true);
        }
        Ok(false)
    }

    fn parse_text(&mut self, text: &str) {
        if self.in_cell {
            self.cell.parse_text(text);
        }
    }

    fn parse_close(&mut self, name: &str) -> XResult<bool> {
        if self.in_cell {
            if !self.cell.parse_close(name)? {
                let cell = self.cell.model.take();
                let count = match &mut self.model {
                    RowState::Open(row) => {
                        // A null cell (no model) is dropped later by the
                        // reconcile; it still counts here.
                        match cell {
                            Some(c) => row.cells.push(c),
                            None => row.null_cells += 1,
                        }
                        row.cells.len() + row.null_cells
                    }
                    RowState::Phantom(n) => {
                        *n += 1;
                        *n
                    }
                    RowState::Undefined => {
                        return Err("Cannot read properties of undefined (reading 'cells')".into())
                    }
                    RowState::Null => {
                        return Err("Cannot read properties of null (reading 'cells')".into())
                    }
                };
                if let Some(max) = self.max_cols {
                    if count > max {
                        return Err(format!("Max column count ({max}) exceeded"));
                    }
                }
                self.in_cell = false;
            }
            return Ok(true);
        }
        Ok(false)
    }

    /// The row the list receives (`this.child.model`), leaving a phantom.
    fn hand_over(&mut self) -> Option<RowModel> {
        match std::mem::replace(&mut self.model, RowState::Undefined) {
            RowState::Open(row) => {
                self.model = RowState::Phantom(row.cells.len() + row.null_cells);
                Some(row)
            }
            other => {
                self.model = other;
                None
            }
        }
    }
}

/// `this.model` of the sheetData `ListXform`.
enum ListState {
    Undefined,
    Null,
    List,
}

/// `ListXform({tag: 'sheetData', childXform: RowXform})` (parse role).
pub(crate) struct SheetDataXform {
    max_rows: Option<usize>,
    row: RowParser,
    in_child: bool,
    state: ListState,
    rows: Vec<RowModel>,
    /// Items in the list, null rows included.
    count: usize,
    out: SharedRows,
}

impl std::fmt::Debug for SheetDataXform {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("SheetDataXform")
            .field("rows", &self.rows.len())
            .finish_non_exhaustive()
    }
}

impl SheetDataXform {
    pub fn new(max_rows: Option<usize>, max_cols: Option<usize>, out: SharedRows) -> Self {
        SheetDataXform {
            max_rows,
            row: RowParser {
                max_cols,
                cell: CellParser {
                    rich: RichTextXform::new(),
                    in_rich: false,
                    t: None,
                    current: None,
                    model: None,
                    skip: 0,
                    seen_value: false,
                },
                in_cell: false,
                model: RowState::Undefined,
            },
            in_child: false,
            state: ListState::Undefined,
            rows: Vec::new(),
            count: 0,
            out,
        }
    }
}

impl Xform for SheetDataXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if self.in_child {
            self.row.parse_open(node)?;
            return Ok(true);
        }
        if node.name == "sheetData" {
            self.state = ListState::List;
            self.rows.clear();
            self.count = 0;
            return Ok(true);
        }
        if self.row.parse_open(node)? {
            self.in_child = true;
            return Ok(true);
        }
        Ok(false)
    }
    fn parse_text(&mut self, text: &str) {
        if self.in_child {
            self.row.parse_text(text);
        }
    }
    fn parse_close(&mut self, name: &str) -> XResult<bool> {
        if self.in_child {
            if !self.row.parse_close(name)? {
                match self.state {
                    ListState::List => {}
                    ListState::Null => {
                        return Err("Cannot read properties of null (reading 'push')".into())
                    }
                    ListState::Undefined => {
                        return Err("Cannot read properties of undefined (reading 'push')".into())
                    }
                }
                if let Some(row) = self.row.hand_over() {
                    self.rows.push(row);
                }
                self.count += 1;
                self.in_child = false;
                if let Some(max) = self.max_rows {
                    if self.count > max {
                        return Err(format!("Max row count ({max}) exceeded"));
                    }
                }
            }
            return Ok(true);
        }
        Ok(false)
    }
    fn model(&self) -> Option<&Value> {
        None
    }
    fn take_model(&mut self) -> Option<Value> {
        // The rows go through `out`; the worksheet model gets `rows: []`
        // (or null before any <sheetData>), which the reconcile replaces.
        self.out.set(std::mem::take(&mut self.rows));
        match self.state {
            ListState::List => Some(json!([])),
            ListState::Null => Some(Value::Null),
            ListState::Undefined => None,
        }
    }
    fn render(&self, _xml: &mut crate::xml_stream::XmlStream, _model: Option<&Value>) {}
    // ListXform.reset: `this.model = null` and the child's reset
    // (RowXform: `this.model = null`).
    fn reset(&mut self) {
        self.state = ListState::Null;
        self.rows.clear();
        self.count = 0;
        self.row.model = RowState::Null;
    }
}

// ---------------------------------------------------------------------------
// Fast path

/// A worksheet part whose `<sheetData>` [`fast_sheet_data`] accepted: the
/// part with `<sheetData/>` in its place (for the general parser), and where
/// the rows start, for [`fast_rows`].
pub(crate) struct FastSheet {
    pub rest: Vec<u8>,
    pub rows_at: usize,
}

/// Checks whether `<sheetData>` can be read without the general XML layer,
/// which costs most of the read time of a large sheet. Only the common
/// well-formed subset qualifies: `<row>`, `<c>`, `<f>` and `<v>` elements (no
/// prefixes), attribute values and text without references or line breaks to
/// normalise, and no comments, CDATA or other markup. Otherwise `None`: the
/// general parser reads the part, so the rows are always the ones it would
/// read (`fast_path_matches_general_parser`). Nothing is built here: the
/// rows are built one at a time by [`fast_rows`] when they are reconciled.
pub(crate) fn fast_sheet_data(doc: &[u8]) -> Option<FastSheet> {
    let text = std::str::from_utf8(doc).ok()?;
    let open = text.find("<sheetData")?;
    let head = text.get(..open)?;
    // Markup before it that could hide or fake the element
    if head.contains("<!") {
        return None;
    }
    let mut sc = Scanner::new(text, open + "<sheetData".len());
    if !sc.eat(">") {
        return None;
    }
    let rows_at = sc.pos;
    sc.rows(None)?.ok()?;
    let tail = text.get(sc.pos..)?;
    if tail.contains("<sheetData") {
        return None;
    }
    let mut rest = Vec::with_capacity(open + "<sheetData/>".len() + tail.len());
    rest.extend_from_slice(head.as_bytes());
    rest.extend_from_slice(b"<sheetData/>");
    rest.extend_from_slice(tail.as_bytes());
    Some(FastSheet { rest, rows_at })
}

/// The rows of a `<sheetData>` that [`fast_sheet_data`] accepted, given to
/// `each` one at a time. `None` if the part is not the one checked.
pub(crate) fn fast_rows(
    doc: &[u8],
    rows_at: usize,
    each: &mut dyn FnMut(RowModel) -> XResult<()>,
) -> Option<XResult<()>> {
    let text = std::str::from_utf8(doc).ok()?;
    let result = Scanner::new(text, rows_at).rows(Some(each));
    if let Err(e) = crate::read_control::check() {
        return Some(Err(e.exceljs_message()));
    }
    result
}

/// [`fast_rows`] over the part as `input` gives it (the entry inflated
/// again), so that the part is not held in memory while its rows are read.
/// `input` must give the bytes [`fast_sheet_data`] accepted. The rows are
/// scanned in windows that end after a `</row>` (which only occurs as a
/// close tag: text and attribute values hold no raw `<`).
pub(crate) fn fast_rows_stream(
    mut input: impl std::io::Read,
    rows_at: usize,
    each: &mut dyn FnMut(RowModel) -> XResult<()>,
) -> Option<XResult<()>> {
    let mut skip = rows_at;
    let mut buf: Vec<u8> = Vec::new();
    let mut eof = false;
    let mut chunk = vec![0u8; 1 << 16];
    loop {
        if let Err(e) = crate::read_control::check() {
            return Some(Err(e.exceljs_message()));
        }
        let n = input.read(&mut chunk).ok()?;
        if n == 0 {
            eof = true;
        }
        let mut got = chunk.get(..n)?;
        if skip > 0 {
            let s = skip.min(got.len());
            skip -= s;
            got = got.get(s..)?;
        }
        let search_at = buf.len().saturating_sub("</sheetData>".len());
        buf.extend_from_slice(got);
        // the end of the last complete row, or the end of <sheetData>
        let search = buf.get(search_at..)?;
        let end = find_last(search, b"</sheetData>")
            .map(|i| search_at + i + "</sheetData>".len())
            .or_else(|| find_last(search, b"</row>").map(|i| search_at + i + "</row>".len()));
        let Some(cut) = end else {
            if eof {
                return None;
            }
            continue;
        };
        let window = std::str::from_utf8(buf.get(..cut)?).ok()?;
        let mut sc = Scanner::new(window, 0);
        let result = sc.rows_in_window(each);
        if let Err(e) = crate::read_control::check() {
            return Some(Err(e.exceljs_message()));
        }
        match result? {
            Ok(true) => return Some(Ok(())),
            Ok(false) => {}
            Err(e) => return Some(Err(e)),
        }
        if sc.pos != cut {
            return None;
        }
        buf.drain(..cut);
        if eof {
            return None;
        }
    }
}

fn find_last(hay: &[u8], needle: &[u8]) -> Option<usize> {
    hay.windows(needle.len()).rposition(|w| w == needle)
}

struct Scanner<'a> {
    s: &'a str,
    b: &'a [u8],
    pos: usize,
    attrs: Vec<(&'a str, &'a str)>,
    control: Option<std::sync::Arc<crate::read_control::ReadControl>>,
    checked_at: usize,
}

impl<'a> Scanner<'a> {
    fn new(s: &'a str, pos: usize) -> Self {
        Scanner {
            s,
            b: s.as_bytes(),
            pos,
            attrs: Vec::new(),
            control: crate::read_control::active(),
            checked_at: pos,
        }
    }

    /// The rows up to `</sheetData>`, built and given to `each` if it is
    /// there, else only checked. `None` at anything outside the subset.
    fn rows(
        &mut self,
        mut each: Option<&mut dyn FnMut(RowModel) -> XResult<()>>,
    ) -> Option<XResult<()>> {
        loop {
            self.checkpoint()?;
            self.skip_ws()?;
            if self.eat("</sheetData>") {
                return Some(Ok(()));
            }
            if !self.eat_tag("row") {
                return None;
            }
            let closed = self.attributes()?;
            let mut row = each.as_ref().map(|_| open_row(|n| self.attr(n)));
            if !closed {
                loop {
                    self.checkpoint()?;
                    self.skip_ws()?;
                    if self.eat("</row>") {
                        break;
                    }
                    if !self.eat_tag("c") {
                        return None;
                    }
                    let cell = self.cell(row.is_some())?;
                    if let (Some(row), Some(cell)) = (row.as_mut(), cell) {
                        row.cells.push(cell);
                    }
                }
            }
            if let (Some(each), Some(row)) = (each.as_mut(), row) {
                if let Err(e) = each(row) {
                    return Some(Err(e));
                }
            }
        }
    }

    /// [`rows`](Self::rows), building them, over a window of complete
    /// rows: `Ok(true)` at `</sheetData>`, `Ok(false)` at the window's end.
    fn rows_in_window(
        &mut self,
        each: &mut dyn FnMut(RowModel) -> XResult<()>,
    ) -> Option<XResult<bool>> {
        loop {
            self.checkpoint()?;
            self.skip_ws()?;
            if self.pos >= self.b.len() {
                return Some(Ok(false));
            }
            if self.eat("</sheetData>") {
                return Some(Ok(true));
            }
            if !self.eat_tag("row") {
                return None;
            }
            let closed = self.attributes()?;
            let mut row = open_row(|n| self.attr(n));
            if !closed {
                loop {
                    self.checkpoint()?;
                    self.skip_ws()?;
                    if self.eat("</row>") {
                        break;
                    }
                    if !self.eat_tag("c") {
                        return None;
                    }
                    if let Some(cell) = self.cell(true)? {
                        row.cells.push(cell);
                    }
                }
            }
            if let Err(e) = each(row) {
                return Some(Err(e));
            }
        }
    }

    fn eat(&mut self, lit: &str) -> bool {
        let found = self
            .b
            .get(self.pos..)
            .is_some_and(|r| r.starts_with(lit.as_bytes()));
        if found {
            self.pos += lit.len();
        }
        found
    }
    /// `<name` followed by whitespace, `>` or `/`.
    fn eat_tag(&mut self, name: &str) -> bool {
        let start = self.pos;
        if self.eat("<")
            && self.eat(name)
            && matches!(
                self.b.get(self.pos),
                Some(b' ' | b'\t' | b'\n' | b'\r' | b'>' | b'/')
            )
        {
            return true;
        }
        self.pos = start;
        false
    }

    // The default scanner retains its byte loop. A controlled scan checks
    // cancellation at most every 8 KiB, including one very long XML token.
    fn checkpoint(&mut self) -> Option<()> {
        if self.pos.saturating_sub(self.checked_at) >= 8192 {
            if let Some(control) = &self.control {
                control.check().ok()?;
            }
            self.checked_at = self.pos;
        }
        Some(())
    }

    fn skip_ws(&mut self) -> Option<()> {
        while matches!(self.b.get(self.pos), Some(b' ' | b'\t' | b'\n' | b'\r')) {
            self.pos += 1;
            self.checkpoint()?;
        }
        Some(())
    }

    fn attr(&self, name: &str) -> Option<&'a str> {
        self.attrs.iter().find(|(k, _)| *k == name).map(|(_, v)| *v)
    }

    /// The attributes up to `>` (false) or `/>` (true), into `self.attrs`.
    fn attributes(&mut self) -> Option<bool> {
        self.attrs.clear();
        loop {
            self.skip_ws()?;
            if self.eat(">") {
                return Some(false);
            }
            if self.eat("/>") {
                return Some(true);
            }
            let start = self.pos;
            while self.b.get(self.pos).is_some_and(|c| {
                c.is_ascii_alphanumeric() || matches!(c, b':' | b'_' | b'-' | b'.')
            }) {
                self.pos += 1;
                self.checkpoint()?;
            }
            let name = self.s.get(start..self.pos)?;
            if name.is_empty() || !self.b.get(start)?.is_ascii_alphabetic() {
                return None;
            }
            self.skip_ws()?;
            if !self.eat("=") {
                return None;
            }
            self.skip_ws()?;
            let quote = *self.b.get(self.pos)?;
            if quote != b'"' && quote != b'\'' {
                return None;
            }
            self.pos += 1;
            let vstart = self.pos;
            loop {
                match *self.b.get(self.pos)? {
                    c if c == quote => break,
                    // references and whitespace that attribute-value
                    // normalisation would change
                    b'&' | b'<' | b'\t' | b'\n' | b'\r' => return None,
                    _ => {
                        self.pos += 1;
                        self.checkpoint()?;
                    }
                }
            }
            let value = self.s.get(vstart..self.pos)?;
            self.pos += 1;
            if self.attr(name).is_some() {
                return None;
            }
            self.attrs.push((name, value));
        }
    }

    /// Text up to the next `<`, without references or carriage returns.
    fn text(&mut self) -> Option<&'a str> {
        let start = self.pos;
        loop {
            match *self.b.get(self.pos)? {
                b'<' => break,
                b'&' | b'\r' => return None,
                _ => {
                    self.pos += 1;
                    self.checkpoint()?;
                }
            }
        }
        self.s.get(start..self.pos)
    }

    /// The rest of a `<c ...` element: `<f>` then `<v>`, each optional. The
    /// cell model if `build`.
    fn cell(&mut self, build: bool) -> Option<Option<CellModel>> {
        let closed = self.attributes()?;
        let mut m = build.then(|| open_cell(|n| self.attr(n)));
        let t = self.attr("t");
        let mut seen_value = false;
        let mut value = None;
        if !closed {
            self.skip_ws()?;
            if self.eat_tag("f") {
                let closed = self.attributes()?;
                if let Some(m) = m.as_mut() {
                    open_formula(m, |n| self.attr(n));
                }
                if !closed {
                    let text = self.text()?;
                    if let Some(m) = m.as_mut().filter(|_| !text.is_empty()) {
                        m.formula = bx(Value::String(text.to_owned()));
                    }
                    if !self.eat("</f>") {
                        return None;
                    }
                }
                self.skip_ws()?;
            }
            if self.eat_tag("v") {
                seen_value = true;
                if !self.attributes()? {
                    let text = self.text()?;
                    if !text.is_empty() {
                        value = Some(text);
                    }
                    if !self.eat("</v>") {
                        return None;
                    }
                }
                self.skip_ws()?;
            }
            if !self.eat("</c>") {
                return None;
            }
        }
        if let Some(m) = m.as_mut() {
            finish_text(m, t, seen_value, value).ok()?;
        }
        Some(m)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xform::parse_part;

    fn row(xml: &str) -> Value {
        let rows = SharedRows::default();
        let mut x = SheetDataXform::new(None, None, rows.clone());
        parse_part(
            &mut x,
            "x",
            format!("<sheetData>{xml}</sheetData>").as_bytes(),
        )
        .unwrap();
        rows.take().first().map(crate::rows::row_value).unwrap()
    }

    #[test]
    fn cancelled_fast_metadata_scan_stops_inside_a_long_token() {
        let doc = format!(
            "<worksheet><sheetData><row><c><v>{}</v></c></row></sheetData></worksheet>",
            "1".repeat(100_000)
        );
        let control = crate::ReadControl::new(crate::ReadLimits::default());
        let _scope = crate::read_control::ControlScope::enter(Some(control.clone()));
        control.cancel();
        let mut scanner = Scanner::new(&doc, "<worksheet><sheetData>".len());
        assert!(scanner.rows(None).is_none());
        assert!(scanner.pos < 20_000);
        assert!(fast_sheet_data(doc.as_bytes()).is_none());
    }

    #[test]
    fn cancelled_fast_stream_stops_an_unfinished_long_row() {
        struct CancellingInput {
            control: std::sync::Arc<crate::ReadControl>,
            reads: usize,
        }
        impl std::io::Read for CancellingInput {
            fn read(&mut self, buf: &mut [u8]) -> std::io::Result<usize> {
                self.reads += 1;
                if self.reads == 2 {
                    self.control.cancel();
                }
                buf.fill(b'1');
                Ok(buf.len())
            }
        }
        let control = crate::ReadControl::new(crate::ReadLimits::default());
        let _scope = crate::read_control::ControlScope::enter(Some(control.clone()));
        let mut input = CancellingInput { control, reads: 0 };
        assert!(matches!(
            fast_rows_stream(&mut input, 0, &mut |_| Ok(())),
            Some(Err(_))
        ));
        assert_eq!(input.reads, 2);
    }

    #[test]
    fn first_guess_types() {
        let r = row(
            r#"<row r="1" spans="1:3"><c r="A1" t="s"><v>2</v></c><c r="B1" t="b"><v>true</v></c><c r="C1"><f>1+1</f><v>2</v></c><c r="D1" s="3"/><c r="E1"/></row>"#,
        );
        assert_eq!(
            r,
            json!({"number": 1, "min": 1, "max": 3, "cells": [
                {"address": "A1", "value": 2, "type": 3},
                {"address": "B1", "value": true, "type": 9},
                {"address": "C1", "formula": "1+1", "type": 6, "result": 2},
                {"address": "D1", "styleId": 3, "type": 0},
                {"address": "E1", "type": 1},
            ]})
        );
    }

    #[test]
    fn unknown_children_are_skipped() {
        // ExcelJS ends the cell at </rPh> (COMPATIBILITY.md B4).
        let r = row(
            r#"<row r="1"><c r="B1" t="inlineStr"><is><t>base</t><rPh sb="0" eb="1"><t>P</t></rPh><phoneticPr fontId="1"/></is></c><c r="C1"><v>3</v></c></row>"#,
        );
        assert_eq!(
            r["cells"][0],
            json!({"address": "B1", "value": "base", "type": 3})
        );
        assert_eq!(
            r["cells"][1],
            json!({"address": "C1", "value": 3, "type": 2})
        );
    }

    #[test]
    fn string_cells_fixed_from_exceljs() {
        let r = row(concat!(
            r#"<row r="1"><c r="A1" t="str"><v></v></c><c r="B1" t="inlineStr"><is><t/></is></c>"#,
            r#"<c r="C1" t="str"><v>a &amp;amp; _x0062_</v></c><c r="D1" t="inlineStr"><is><t>_x0063_</t></is></c>"#,
            r#"<c r="E1" t="b"><v>false</v></c><c r="F1" t="b"><v>TRUE</v></c><c r="G1" t="str"><f>"x"</f><v></v></c></row>"#,
        ));
        let values: Vec<_> = r["cells"]
            .as_array()
            .unwrap()
            .iter()
            .map(|c| c.get("value").or(c.get("result")).cloned())
            .collect();
        assert_eq!(
            values,
            [
                json!(""),
                json!(""),
                json!("a &amp; b"),
                json!("c"),
                json!(false),
                json!(true),
                json!("")
            ]
            .map(Some)
        );
    }

    /// A worksheet model and its rows.
    type Paths = (Value, Vec<Value>);

    /// The worksheet model and rows of `doc` read by the general parser,
    /// and by the fast path if it takes the part.
    fn both_paths(doc: &[u8]) -> (Result<Paths, String>, Option<Paths>) {
        use crate::xform::worksheet::{SheetOptions, WorksheetXform};
        let read = |bytes: &[u8]| -> Result<(Value, WorksheetXform), String> {
            let mut x = WorksheetXform::new(&SheetOptions::default());
            let m = parse_part(&mut x, "sheet.xml", bytes).map_err(|e| e.to_string())?;
            Ok((m.unwrap_or(Value::Null), x))
        };
        let values =
            |rows: Vec<RowModel>| rows.iter().map(crate::rows::row_value).collect::<Vec<_>>();
        let general = read(doc).map(|(m, x)| (m, values(x.take_rows())));
        // (the reader re-reads the whole part when the rest is rejected)
        let fast = fast_sheet_data(doc).and_then(|f| {
            let mut rows = Vec::new();
            fast_rows(doc, f.rows_at, &mut |r| {
                rows.push(r);
                Ok(())
            })
            .expect("fast_rows declined what fast_sheet_data accepted")
            .expect("fast_rows failed");
            read(&f.rest).ok().map(|(m, _)| (m, values(rows)))
        });
        (general, fast)
    }

    fn worksheet_parts() -> Vec<(String, Vec<u8>)> {
        use std::io::Read;
        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
        let mut files = Vec::new();
        let mut dirs = vec![
            root.join("fixtures"),
            root.join("compat/upstream/spec/integration/data"),
        ];
        while let Some(d) = dirs.pop() {
            for e in std::fs::read_dir(&d).into_iter().flatten().flatten() {
                let p = e.path();
                if p.is_dir() {
                    dirs.push(p);
                } else if p.extension().is_some_and(|x| x == "xlsx")
                    // huge.xlsx: 14 MB; issue-1842: a whole-sheet data
                    // validation, expanded cell by cell (minutes)
                    && !p.ends_with("huge.xlsx")
                    && !p.to_string_lossy().contains("1842")
                {
                    files.push(p);
                }
            }
        }
        let mut out = Vec::new();
        for f in files {
            let Ok(data) = std::fs::read(&f) else {
                continue;
            };
            let Ok(mut z) = zip::ZipArchive::new(std::io::Cursor::new(data)) else {
                continue;
            };
            for i in 0..z.len() {
                let Ok(mut e) = z.by_index(i) else { continue };
                let Ok(name) = e.name().map(|n| n.into_owned()) else {
                    continue;
                };
                if name.contains("worksheets/sheet") && name.ends_with(".xml") {
                    let mut b = Vec::new();
                    if e.read_to_end(&mut b).is_ok() {
                        out.push((format!("{}:{name}", f.display()), b));
                    }
                }
            }
        }
        out
    }

    #[test]
    fn fast_path_matches_general_parser() {
        let wrap = |data: &str| {
            format!(
                r#"<worksheet xmlns="x"><dimension ref="A1"/><sheetData>{data}</sheetData><pageMargins left="0.7"/></worksheet>"#
            )
        };
        let snippets = [
            r#"<row r="1" spans="1:3" s="2" customFormat="1" ht="20" customHeight="1" hidden="1" outlineLevel="1" collapsed="1"><c r="A1"><v>1</v></c></row>"#,
            r#"<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="str"><v>a _x0041_</v></c><c r="C1" t="b"><v>false</v></c><c r="D1" t="e"><v>#N/A</v></c></row>"#,
            r#"<row r="1"><c r="A1"><f>B1*2</f><v>4</v></c><c r="B1"><f t="shared" ref="B1:B3" si="0">C1</f><v>0</v></c><c r="C1"><f t="shared" si="0"/></c></row>"#,
            r#"<row r="1"><c r="A1" t="str"><f>"x"</f><v></v></c><c r="B1" t="str"><v/></c><c r="C1" s="3"/><c r="D1"/><c r="E1"><v></v></c></row>"#,
            r#"<row r="2"/><row><c><v>1</v></c><c><v>2</v></c></row><row r="5"></row>"#,
            "\n  <row r=\"1\" >\n    <c r=\"A1\" s='1' t=\"n\" >\n      <v> 1e3 </v>\n    </c>\n  </row>\n",
            r#"<row r="1"><c r="A1"><v>12abc</v></c><c r="B1"><v>-Infinity</v></c><c r="C1" t="s"><v> 1.9 </v></c></row>"#,
            r#"<row r="1" x14ac:dyDescent="0.25"><c r="A1" s="" t=""><v>1</v></c></row>"#,
            // left to the general parser
            r#"<row r="1"><c r="A1" t="inlineStr"><is><t>x</t></is></c></row>"#,
            r#"<row r="1"><c r="A1"><v>a &amp; b</v></c></row>"#,
            r#"<row r="1"><c r="A1"><v><![CDATA[1]]></v></c></row>"#,
            r#"<row r="1"><!-- c --><c r="A1"><v>1</v></c></row>"#,
            r#"<row r="1" r="2"><c r="A1"><v>1</v></c></row>"#,
            r#"<row r="1"><c r="A1"><v>1</v><extLst/></c></row>"#,
            r#"<row r="1"><x:c r="A1"><v>1</v></x:c></row>"#,
            "<row r=\"1\"><c r=\"A1\" t=\"str\"><v>a\r\nb</v></c></row>",
            r#"<c r="A1"><v>1</v></c>"#,
        ];
        let mut docs: Vec<(String, Vec<u8>)> = snippets
            .iter()
            .enumerate()
            .map(|(i, s)| (format!("snippet {i}"), wrap(s).into_bytes()))
            .collect();
        docs.push((
            "no sheetData".into(),
            b"<worksheet><dimension ref=\"A1\"/></worksheet>".to_vec(),
        ));
        docs.push((
            "empty sheetData".into(),
            b"<worksheet><sheetData/></worksheet>".to_vec(),
        ));
        docs.push((
            "two sheetData".into(),
            wrap("</sheetData><sheetData><row r=\"1\"/>").into_bytes(),
        ));
        docs.push((
            "comment before".into(),
            b"<worksheet><!-- <sheetData> --><sheetData><row r=\"1\"/></sheetData></worksheet>"
                .to_vec(),
        ));
        let parts = worksheet_parts();
        assert!(parts.len() >= 20, "fixtures missing");
        docs.extend(parts);
        let mut fast_taken = 0;
        for (name, doc) in &docs {
            let (general, fast) = both_paths(doc);
            if let Some(fast) = fast {
                fast_taken += 1;
                assert_eq!(general.as_ref().ok(), Some(&fast), "{name}");
                // and read again in windows, as the reader does
                let rows_at = fast_sheet_data(doc).unwrap().rows_at;
                for chunk in [1, 7, 100, 1 << 16] {
                    let mut rows = Vec::new();
                    fast_rows_stream(Chunks(doc, chunk), rows_at, &mut |r| {
                        rows.push(crate::rows::row_value(&r));
                        Ok(())
                    })
                    .unwrap_or_else(|| panic!("{name}: windows declined (chunks of {chunk})"))
                    .unwrap();
                    assert_eq!(rows, fast.1, "{name}: windows of {chunk}");
                }
            }
        }
        // The fast path must actually be used: the ExcelJS-written fixtures.
        assert!(
            fast_taken * 2 > docs.len(),
            "fast path taken for {fast_taken} of {}",
            docs.len()
        );
        // and declines what it does not handle
        for (i, s) in snippets.iter().enumerate().skip(8) {
            assert!(fast_sheet_data(wrap(s).as_bytes()).is_none(), "snippet {i}");
        }
    }

    /// Reads at most `.1` bytes at a time.
    struct Chunks<'a>(&'a [u8], usize);

    impl std::io::Read for Chunks<'_> {
        fn read(&mut self, buf: &mut [u8]) -> std::io::Result<usize> {
            let n = self.1.min(buf.len()).min(self.0.len());
            let (head, tail) = self.0.split_at(n);
            buf.get_mut(..n).unwrap().copy_from_slice(head);
            self.0 = tail;
            Ok(n)
        }
    }
}
