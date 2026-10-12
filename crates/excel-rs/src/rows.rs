//! Typed row and cell models for the writers (`RowXform`/`CellXform`
//! `prepare` and `render`).
//!
//! The JS side sends each worksheet's rows separately from the rest of the
//! model: cell models carry only the keys `CellXform` reads, and styles are
//! indexes into a table of distinct style objects. Each field keeps JS's
//! distinction between a missing key (`None`) and `null` (`Some(Null)`), so
//! the rules below read exactly like the `Value`-based xforms.

use std::collections::HashMap;

use serde::{Deserialize, Deserializer};
use serde_json::{json, Map, Value};

use crate::js::encode_x_escapes;
use crate::jsv::{as_f64, as_str, get, number_to_string, to_js_string, truthy};
use crate::model_json::{date_to_excel, value_as_date};
use crate::writer::{RangeBox, SharedStrings, StyleManager};
use crate::xform::cell::{value_type, vt};
use crate::xform::strings::RichTextXform;
use crate::xform::Xform;
use crate::xml_stream::XmlStream;

type R<T> = std::result::Result<T, String>;

/// A present key, `null` included (`#[serde(default)]` makes absent `None`).
fn present<'de, D: Deserializer<'de>>(d: D) -> Result<Option<Value>, D::Error> {
    Value::deserialize(d).map(Some)
}

/// A cell model key: boxed, since most keys are absent from most cells.
pub(crate) type Field = Option<Box<Value>>;

fn present_boxed<'de, D: Deserializer<'de>>(d: D) -> Result<Field, D::Error> {
    Value::deserialize(d).map(|v| Some(Box::new(v)))
}

/// `type`, read as JS reads it here (`model.type` compared with numbers).
fn present_number<'de, D: Deserializer<'de>>(d: D) -> Result<Option<f64>, D::Error> {
    Value::deserialize(d).map(|v| as_f64(Some(&v)))
}

fn bx(v: Value) -> Field {
    Some(Box::new(v))
}

/// A cell model (`cell.model`), limited to what `CellXform` reads.
#[derive(Debug, Default, Clone, Deserialize)]
#[serde(default)]
pub(crate) struct CellModel {
    #[serde(deserialize_with = "present_boxed")]
    pub address: Field,
    #[serde(rename = "type", deserialize_with = "present_number")]
    pub ty: Option<f64>,
    #[serde(deserialize_with = "present_boxed")]
    pub value: Field,
    /// Index into the style table (`cell.model.style`).
    #[serde(rename = "$s")]
    pub style: Option<usize>,
    #[serde(deserialize_with = "present_boxed")]
    pub comment: Field,
    #[serde(deserialize_with = "present_boxed")]
    pub text: Field,
    #[serde(deserialize_with = "present_boxed")]
    pub hyperlink: Field,
    #[serde(deserialize_with = "present_boxed")]
    pub tooltip: Field,
    #[serde(deserialize_with = "present_boxed")]
    pub master: Field,
    #[serde(rename = "shareType", deserialize_with = "present_boxed")]
    pub share_type: Field,
    #[serde(deserialize_with = "present_boxed")]
    pub formula: Field,
    #[serde(rename = "sharedFormula", deserialize_with = "present_boxed")]
    pub shared_formula: Field,
    #[serde(deserialize_with = "present_boxed")]
    pub result: Field,
    #[serde(rename = "ref", deserialize_with = "present_boxed")]
    pub ref_: Field,
    #[serde(deserialize_with = "present_boxed")]
    pub si: Field,
    #[serde(deserialize_with = "present_boxed")]
    pub range: Field,
    #[serde(deserialize_with = "present_boxed")]
    pub date1904: Field,
    #[serde(rename = "ssId", deserialize_with = "present_boxed")]
    pub ss_id: Field,
    #[serde(rename = "styleId", deserialize_with = "present_boxed")]
    pub style_id: Field,
    /// The reader's `style` (a shared style reference); the writers use
    /// `style`, an index into the style table.
    #[serde(skip)]
    pub style_value: Field,
    /// The reader's `style` when it is shared object `n` (`{"$ref": n}`).
    #[serde(skip)]
    pub style_ref: Option<u32>,
    /// The reader's `value` when it is shared string `n`, a plain string.
    #[serde(skip)]
    pub sst: Option<u32>,
}

/// A row model (`row.model`).
#[derive(Debug, Default, Deserialize)]
#[serde(default)]
pub(crate) struct RowModel {
    #[serde(deserialize_with = "present")]
    pub number: Option<Value>,
    #[serde(deserialize_with = "present")]
    pub min: Option<Value>,
    #[serde(deserialize_with = "present")]
    pub max: Option<Value>,
    #[serde(deserialize_with = "present")]
    pub height: Option<Value>,
    #[serde(deserialize_with = "present")]
    pub hidden: Option<Value>,
    #[serde(rename = "outlineLevel", deserialize_with = "present")]
    pub outline_level: Option<Value>,
    #[serde(deserialize_with = "present")]
    pub collapsed: Option<Value>,
    #[serde(rename = "$s")]
    pub style: Option<usize>,
    #[serde(rename = "styleId", deserialize_with = "present")]
    pub style_id: Option<Value>,
    #[serde(rename = "bestFit", deserialize_with = "present")]
    pub best_fit: Option<Value>,
    /// The reader's `style` (see `CellModel::style_value`).
    #[serde(skip)]
    pub style_value: Option<Value>,
    pub cells: Vec<CellModel>,
    /// Null cells read (dropped by the reconcile; counted by the parser).
    #[serde(skip)]
    pub null_cells: usize,
}

/// The distinct style objects cells and rows refer to, and the style ids
/// they resolve to (`addStyleModel` is deterministic per style and cell
/// type, so each pair is computed once).
#[derive(Debug, Default)]
pub(crate) struct StyleTable {
    pub styles: Vec<Value>,
    ids: HashMap<(Option<usize>, i64), usize>,
}

impl StyleTable {
    pub fn new(styles: Vec<Value>) -> Self {
        StyleTable {
            styles,
            ids: HashMap::new(),
        }
    }

    /// `styles.addStyleModel(style || {}, cellType)` for a cell.
    fn cell_style_id(
        &mut self,
        manager: &mut StyleManager,
        style: Option<usize>,
        cell_type: i64,
    ) -> usize {
        if let Some(&id) = self.ids.get(&(style, cell_type)) {
            return id;
        }
        let model = style
            .and_then(|i| self.styles.get(i))
            .filter(|v| truthy(Some(v)));
        let empty = json!({});
        let id = manager.add_style_model(Some(model.unwrap_or(&empty)), cell_type);
        self.ids.insert((style, cell_type), id);
        id
    }

    /// `styles.addStyleModel(model.style)` for a row (no `|| {}`).
    fn row_style_id(&mut self, manager: &mut StyleManager, style: Option<usize>) -> usize {
        const ROW: i64 = i64::MIN;
        if let Some(&id) = self.ids.get(&(style, ROW)) {
            return id;
        }
        let id = manager.add_style_model(style.and_then(|i| self.styles.get(i)), -1);
        self.ids.insert((style, ROW), id);
        id
    }
}

/// Where a shared formula master is: (row index, cell index) for the
/// document writer, or an index into the streaming writer's master list.
pub(crate) type MasterRef = (usize, usize);

/// What `CellXform.prepare` asks of the formula registry.
pub(crate) enum CellPrep {
    None,
    /// `options.formulae[address] = model`
    Register(String),
    /// A clone of the master at this position, whose address is
    /// `master_address`.
    Clone {
        master: MasterRef,
        master_address: String,
        address: String,
    },
}

/// The shared-formula side of formula cells (`options.formulae` holds every
/// formula cell, any of which a later clone can make a master). Only what
/// `share_master` reads and writes (`address`, `shareType`, `si`, `range`) is
/// kept, and only for cells that have some of it or got a clone, so a sheet
/// of formulas does not keep a copy of each cell.
#[derive(Debug, Default)]
pub(crate) struct Masters {
    by_ref: HashMap<MasterRef, CellModel>,
}

impl Masters {
    /// `options.formulae[address] = cell`, for the cell at `at`.
    pub fn register(&mut self, at: MasterRef, cell: &CellModel) {
        if cell.share_type.is_some() || cell.si.is_some() || cell.range.is_some() {
            self.by_ref.insert(
                at,
                CellModel {
                    address: cell.address.clone(),
                    share_type: cell.share_type.clone(),
                    si: cell.si.clone(),
                    range: cell.range.clone(),
                    ..CellModel::default()
                },
            );
        } else {
            self.by_ref.remove(&at);
        }
    }

    /// [`share_master`] on the master at `at` for the clone at `address`.
    pub fn share(
        &mut self,
        at: MasterRef,
        master_address: &str,
        address: &str,
        si_formulae: &mut i64,
    ) -> R<Value> {
        let m = self.by_ref.entry(at).or_insert_with(|| CellModel {
            address: bx(json!(master_address)),
            ..CellModel::default()
        });
        share_master(m, address, si_formulae)
    }

    /// Gives the master cell at `at`, as it is rendered, what its clones
    /// changed.
    pub fn apply(&self, at: MasterRef, cell: &mut CellModel) {
        if let Some(m) = self.by_ref.get(&at) {
            cell.share_type.clone_from(&m.share_type);
            cell.si.clone_from(&m.si);
            cell.range.clone_from(&m.range);
        }
    }
}

/// The workbook-wide parts of `CellXform.prepare`'s options.
pub(crate) struct CellContext<'a> {
    pub styles: &'a mut StyleManager,
    pub table: &'a mut StyleTable,
    pub shared_strings: Option<&'a mut SharedStrings>,
    pub date1904: bool,
}

#[allow(clippy::cast_possible_truncation)]
fn trunc(t: f64) -> i64 {
    t as i64
}

/// `getEffectiveCellType(cell)`.
fn effective_cell_type(c: &CellModel) -> R<i64> {
    match c.ty.map(trunc) {
        Some(vt::FORMULA) => value_type(c.result.as_deref()),
        Some(t) => Ok(t),
        None => Ok(-1),
    }
}

/// `RowXform.prepare(model, options)` without the cells.
pub(crate) fn prepare_row(row: &mut RowModel, styles: &mut StyleManager, table: &mut StyleTable) {
    let id = table.row_style_id(styles, row.style);
    if id != 0 {
        row.style_id = Some(json!(id));
    }
}

/// `CellXform.prepare(model, options)`.
#[allow(clippy::too_many_arguments)]
pub(crate) fn prepare_cell(
    c: &mut CellModel,
    o: &mut CellContext<'_>,
    hyperlinks: &mut Vec<Value>,
    comments: &mut Vec<Value>,
    merges: &mut Merges,
    formulae: &HashMap<String, MasterRef>,
    si_formulae: &mut i64,
) -> R<CellPrep> {
    let cell_type = effective_cell_type(c)?;
    let style_id = o.table.cell_style_id(o.styles, c.style, cell_type);
    if style_id != 0 {
        c.style_id = bx(json!(style_id));
    }
    if let Some(comment) = c.comment.as_deref().filter(|v| truthy(Some(v))) {
        let mut m = match comment {
            Value::Object(o) => o.clone(),
            _ => Map::new(),
        };
        m.insert(
            "ref".into(),
            c.address.as_deref().cloned().unwrap_or(Value::Null),
        );
        comments.push(Value::Object(m));
    }
    let address = to_js_string(c.address.as_deref());
    let mut result = CellPrep::None;
    match c.ty.map(trunc) {
        Some(vt::STRING | vt::RICH_TEXT) => {
            if let Some(ss) = o.shared_strings.as_deref_mut() {
                let v = c.value.as_deref().unwrap_or(&Value::Null);
                c.ss_id = bx(json!(ss.add(v)));
            }
        }
        Some(vt::DATE) => {
            if o.date1904 {
                c.date1904 = bx(json!(true));
            }
        }
        Some(vt::HYPERLINK) => {
            if let (Some(ss), Some(t)) = (
                o.shared_strings.as_deref_mut(),
                c.text.as_deref().filter(|t| !t.is_null()),
            ) {
                c.ss_id = bx(json!(ss.add(t)));
            }
            let mut h = Map::new();
            h.insert(
                "address".into(),
                c.address.as_deref().cloned().unwrap_or(Value::Null),
            );
            if let Some(t) = c.hyperlink.as_deref() {
                h.insert("target".into(), t.clone());
            }
            if let Some(t) = c.tooltip.as_deref() {
                h.insert("tooltip".into(), t.clone());
            }
            hyperlinks.push(Value::Object(h));
        }
        Some(vt::MERGE) => merges.add(&to_js_string(c.master.as_deref()), &address)?,
        Some(vt::FORMULA) => {
            if o.date1904 {
                c.date1904 = bx(json!(true));
            }
            if as_str(c.share_type.as_deref()) == Some("shared") {
                c.si = bx(json!(*si_formulae));
                *si_formulae += 1;
            }
            if truthy(c.formula.as_deref()) {
                result = CellPrep::Register(address);
            } else if truthy(c.shared_formula.as_deref()) {
                let key = to_js_string(c.shared_formula.as_deref());
                let master = formulae.get(&key).copied().ok_or_else(|| {
                    format!("Shared Formula master must exist above and or left of clone for cell {address}")
                })?;
                result = CellPrep::Clone {
                    master,
                    master_address: key,
                    address,
                };
            }
        }
        _ => {}
    }
    Ok(result)
}

/// The master side of a shared formula clone: gives the master an `si` and
/// a range on first use, else widens its range. Returns the clone's `si`.
pub(crate) fn share_master(m: &mut CellModel, address: &str, si_formulae: &mut i64) -> R<Value> {
    if m.si.is_none() {
        m.share_type = bx(json!("shared"));
        m.si = bx(json!(*si_formulae));
        *si_formulae += 1;
        let master_address = to_js_string(m.address.as_deref());
        let r = RangeBox::from_addresses(&master_address, address)?;
        m.range = bx(r.to_model());
    } else if let Some(range) = m.range.as_deref().filter(|v| truthy(Some(v))) {
        let mut r = RangeBox::from_model(Some(range));
        r.expand_to_address(address)?;
        m.range = bx(r.to_model());
    }
    Ok(m.si.as_deref().cloned().unwrap_or(Value::Null))
}

/// `options.merges` (`Merges` in worksheet-xform.js): one range per master,
/// in first-seen order.
#[derive(Debug, Default)]
pub(crate) struct Merges {
    order: Vec<(String, RangeBox)>,
    index: HashMap<String, usize>,
    ignore: bool,
}

impl Merges {
    /// The streaming writer's `merges` (`add` is a no-op).
    pub fn ignored() -> Self {
        Merges {
            ignore: true,
            ..Merges::default()
        }
    }

    fn add(&mut self, master: &str, address: &str) -> R<()> {
        if self.ignore {
            return Ok(());
        }
        match self.index.get(master) {
            Some(&i) => {
                if let Some((_, r)) = self.order.get_mut(i) {
                    r.expand_to_address(address)?;
                }
            }
            None => {
                self.index.insert(master.to_owned(), self.order.len());
                self.order.push((
                    master.to_owned(),
                    RangeBox::from_addresses(master, address)?,
                ));
            }
        }
        Ok(())
    }

    /// `merges.mergeCells`: the ranges as strings.
    pub fn ranges(&self) -> R<Vec<Value>> {
        self.order
            .iter()
            .map(|(_, r)| r.range().map(Value::String))
            .collect()
    }
}

/// `RowXform.render(xmlStream, model)`.
pub(crate) fn render_row(xml: &mut XmlStream, m: &RowModel, rich: &RichTextXform) {
    xml.open_node("row", &[]);
    xml.add_attribute("r", m.number.as_ref());
    if truthy(m.height.as_ref()) {
        xml.add_attribute("ht", m.height.as_ref());
        xml.add_attribute("customHeight", Some("1"));
    }
    if truthy(m.hidden.as_ref()) {
        xml.add_attribute("hidden", Some("1"));
    }
    if let (Some(a), Some(b)) = (as_f64(m.min.as_ref()), as_f64(m.max.as_ref())) {
        if a > 0.0 && b > 0.0 && a <= b {
            xml.add_attribute(
                "spans",
                Some(format!(
                    "{}:{}",
                    to_js_string(m.min.as_ref()),
                    to_js_string(m.max.as_ref())
                )),
            );
        }
    }
    if truthy(m.style_id.as_ref()) {
        xml.add_attribute("s", m.style_id.as_ref());
        xml.add_attribute("customFormat", Some("1"));
    }
    xml.add_attribute("x14ac:dyDescent", Some("0.25"));
    if truthy(m.outline_level.as_ref()) {
        xml.add_attribute("outlineLevel", m.outline_level.as_ref());
    }
    if truthy(m.collapsed.as_ref()) {
        xml.add_attribute("collapsed", Some("1"));
    }
    for cell in &m.cells {
        render_cell(xml, cell, rich);
    }
    xml.close_node();
}

/// `CellXform.render(xmlStream, model)`.
pub(crate) fn render_cell(xml: &mut XmlStream, m: &CellModel, rich: &RichTextXform) {
    let ty = m.ty.map(trunc);
    if ty == Some(vt::NULL) && !truthy(m.style_id.as_deref()) {
        return;
    }
    xml.open_node("c", &[]);
    xml.add_attribute("r", m.address.as_deref());
    if truthy(m.style_id.as_deref()) {
        xml.add_attribute("s", m.style_id.as_deref());
    }
    match ty {
        Some(vt::NUMBER) => match &m.value {
            Some(v) => xml.leaf_node("v", &[], Some(&to_js_string(Some(v)))),
            None => xml.leaf_node("v", &[], None),
        },
        Some(vt::BOOLEAN) => {
            xml.add_attribute("t", Some("b"));
            xml.leaf_node(
                "v",
                &[],
                Some(if truthy(m.value.as_deref()) { "1" } else { "0" }),
            );
        }
        Some(vt::ERROR) => {
            xml.add_attribute("t", Some("e"));
            let e = get(m.value.as_deref(), "error").map(|e| to_js_string(Some(e)));
            xml.leaf_node("v", &[], e.as_deref());
        }
        Some(vt::STRING | vt::RICH_TEXT) => {
            if m.ss_id.is_some() {
                xml.add_attribute("t", Some("s"));
                xml.leaf_node("v", &[], Some(&to_js_string(m.ss_id.as_deref())));
            } else if truthy(m.value.as_deref()) && truthy(get(m.value.as_deref(), "richText")) {
                xml.add_attribute("t", Some("inlineStr"));
                xml.open_node("is", &[]);
                for run in get(m.value.as_deref(), "richText")
                    .and_then(Value::as_array)
                    .into_iter()
                    .flatten()
                {
                    rich.render(xml, Some(run));
                }
                xml.close_node();
            } else {
                xml.add_attribute("t", Some("str"));
                let v = m.value.as_deref().map(|v| to_js_string(Some(v)));
                xml.leaf_node("v", &[], v.as_deref().map(encode_x_escapes).as_deref());
            }
        }
        Some(vt::DATE) => {
            let ms = value_as_date(m.value.as_deref()).unwrap_or(f64::NAN);
            let v = date_to_excel(ms, truthy(m.date1904.as_deref()));
            xml.leaf_node("v", &[], Some(&number_to_string(v)));
        }
        Some(vt::HYPERLINK) => {
            if m.ss_id.is_some() {
                xml.add_attribute("t", Some("s"));
                xml.leaf_node("v", &[], Some(&to_js_string(m.ss_id.as_deref())));
            } else {
                xml.add_attribute("t", Some("str"));
                let v = m.text.as_deref().map(|v| to_js_string(Some(v)));
                xml.leaf_node("v", &[], v.as_deref().map(encode_x_escapes).as_deref());
            }
        }
        Some(vt::FORMULA) => render_formula(xml, m),
        _ => {}
    }
    xml.close_node();
}

/// `CellXform.renderFormula`.
fn render_formula(xml: &mut XmlStream, m: &CellModel) {
    let attrs = match as_str(m.share_type.as_deref()) {
        Some("shared") => {
            let r = if truthy(m.ref_.as_deref()) {
                m.ref_.as_deref().map(|v| to_js_string(Some(v)))
            } else {
                get(m.range.as_deref(), "range").map(|v| to_js_string(Some(v)))
            };
            Some(crate::attrs![
                ("t", Some("shared")),
                ("ref", r),
                ("si", m.si.as_deref())
            ])
        }
        Some("array") => Some(crate::attrs![
            ("t", Some("array")),
            ("ref", m.ref_.as_deref())
        ]),
        _ => {
            m.si.as_deref()
                .map(|si| crate::attrs![("t", Some("shared")), ("si", Some(si))])
        }
    };
    let attrs = attrs.unwrap_or_default();
    let formula = m.formula.as_deref().map(|f| to_js_string(Some(f)));
    let result = m.result.as_deref();
    match value_type(result) {
        Ok(vt::NULL) => xml.leaf_node("f", &attrs, formula.as_deref()),
        Ok(vt::STRING) => {
            xml.add_attribute("t", Some("str"));
            xml.leaf_node("f", &attrs, formula.as_deref());
            xml.leaf_node("v", &[], Some(&encode_x_escapes(&to_js_string(result))));
        }
        Ok(vt::NUMBER) => {
            xml.leaf_node("f", &attrs, formula.as_deref());
            xml.leaf_node("v", &[], Some(&to_js_string(result)));
        }
        Ok(vt::BOOLEAN) => {
            xml.add_attribute("t", Some("b"));
            xml.leaf_node("f", &attrs, formula.as_deref());
            xml.leaf_node("v", &[], Some(if truthy(result) { "1" } else { "0" }));
        }
        Ok(vt::ERROR) => {
            xml.add_attribute("t", Some("e"));
            xml.leaf_node("f", &attrs, formula.as_deref());
            xml.leaf_node("v", &[], Some(&to_js_string(get(result, "error"))));
        }
        Ok(vt::DATE) => {
            xml.leaf_node("f", &attrs, formula.as_deref());
            let ms = value_as_date(result).unwrap_or(f64::NAN);
            let v = date_to_excel(ms, truthy(m.date1904.as_deref()));
            xml.leaf_node("v", &[], Some(&number_to_string(v)));
        }
        _ => xml.fail("I could not understand type of value"),
    }
}

/// Shared formula masters render `ref: model.ref || model.range.range`.
pub(crate) fn finish_master_range(cell: &mut CellModel) -> R<()> {
    if let Some(range) = cell.range.as_deref().filter(|r| r.get("top").is_some()) {
        let r = RangeBox::from_model(Some(range)).range()?;
        cell.range = bx(json!({ "range": r }));
    }
    Ok(())
}

// ---------------------------------------------------------------------------
// Binary row encoding (lib/xlsx/rows.js)
//
// A sheet is a u32 row count and that many rows. A row or cell is a list of
// fields, each a key byte and a value, ended by END. A row's `CELLS` field is
// a u32 count followed by the cells. Values:
//   F64 <f64 le>            number (non-finite as the codec's $num tag)
//   STR <u32 len> <utf8>    string
//   TRUE, FALSE, NULL
//   DATE <f64 le>           Date (NaN: invalid date)
//   JSON <u32 len> <utf8>   any other value, as codec JSON
//   INDEX <u32 le>          style index (key STYLE)

const END: u8 = 0xFF;
const CELLS: u8 = 31;
const STYLE: u8 = 30;
const F64: u8 = 1;
const STR: u8 = 2;
const TRUE: u8 = 3;
const FALSE: u8 = 4;
const NULL: u8 = 5;
const DATE: u8 = 6;
const JSON_VALUE: u8 = 7;
const INDEX: u8 = 8;

struct Bin<'a> {
    data: &'a [u8],
    pos: usize,
    /// The current row's number as text (for `ADDR` values)
    row_number: String,
}

impl Bin<'_> {
    fn u8(&mut self) -> R<u8> {
        let b = self
            .data
            .get(self.pos)
            .copied()
            .ok_or_else(|| "rows: truncated".to_owned())?;
        self.pos += 1;
        Ok(b)
    }
    fn bytes(&mut self, n: usize) -> R<&[u8]> {
        let s = self
            .data
            .get(self.pos..self.pos + n)
            .ok_or_else(|| "rows: truncated".to_owned())?;
        self.pos += n;
        Ok(s)
    }
    fn u32(&mut self) -> R<usize> {
        let b = self.bytes(4)?;
        let mut a = [0u8; 4];
        a.copy_from_slice(b);
        usize::try_from(u32::from_le_bytes(a)).map_err(|e| e.to_string())
    }
    fn f64(&mut self) -> R<f64> {
        let b = self.bytes(8)?;
        let mut a = [0u8; 8];
        a.copy_from_slice(b);
        Ok(f64::from_le_bytes(a))
    }
    fn str(&mut self) -> R<&str> {
        let n = self.u32()?;
        std::str::from_utf8(self.bytes(n)?).map_err(|e| e.to_string())
    }
    fn value(&mut self) -> R<Value> {
        Ok(match self.u8()? {
            F64 => crate::jsv::num(self.f64()?),
            STR => Value::String(self.str()?.to_owned()),
            TRUE => Value::Bool(true),
            FALSE => Value::Bool(false),
            NULL => Value::Null,
            DATE => {
                let ms = self.f64()?;
                if ms.is_nan() {
                    json!({ "$date": null })
                } else {
                    json!({ "$date": crate::jsv::num(ms) })
                }
            }
            JSON_VALUE => serde_json::from_str(self.str()?).map_err(|e| e.to_string())?,
            REF => json!({ "$ref": self.u32()? }),
            ADDR => {
                let col = self.u32()?;
                #[allow(clippy::cast_precision_loss)]
                let letters = crate::col_cache::n2l_str(col as f64)?;
                Value::String(format!("{letters}{}", self.row_number))
            }
            t => return Err(format!("rows: bad value tag {t}")),
        })
    }
    fn index(&mut self) -> R<usize> {
        match self.u8()? {
            INDEX => self.u32(),
            t => Err(format!("rows: bad style tag {t}")),
        }
    }

    /// The next tag, not consumed.
    fn peek(&self) -> Option<u8> {
        self.data.get(self.pos).copied()
    }

    fn cell(&mut self) -> R<CellModel> {
        let mut c = CellModel::default();
        loop {
            let key = self.u8()?;
            if key == END {
                return Ok(c);
            }
            if key == STYLE {
                c.style = Some(self.index()?);
                continue;
            }
            // the reader's tags (`SheetRows::encode`)
            if key == 2 && self.peek() == Some(SST) {
                self.pos += 1;
                c.sst = Some(u32::try_from(self.u32()?).map_err(|e| e.to_string())?);
                continue;
            }
            if key == STYLE_VALUE {
                if self.peek() == Some(REF) {
                    self.pos += 1;
                    c.style_ref = Some(u32::try_from(self.u32()?).map_err(|e| e.to_string())?);
                } else {
                    c.style_value = bx(self.value()?);
                }
                continue;
            }
            let v = bx(self.value()?);
            match key {
                0 => c.address = v,
                1 => c.ty = as_f64(v.as_deref()),
                2 => c.value = v,
                3 => c.comment = v,
                4 => c.text = v,
                5 => c.hyperlink = v,
                6 => c.tooltip = v,
                7 => c.master = v,
                8 => c.share_type = v,
                9 => c.formula = v,
                10 => c.shared_formula = v,
                11 => c.result = v,
                12 => c.ref_ = v,
                13 => c.si = v,
                14 => c.range = v,
                15 => c.date1904 = v,
                16 => c.ss_id = v,
                17 => c.style_id = v,
                k => return Err(format!("rows: bad cell key {k}")),
            }
        }
    }

    fn row(&mut self) -> R<RowModel> {
        let mut r = RowModel::default();
        loop {
            let key = self.u8()?;
            match key {
                END => return Ok(r),
                STYLE => r.style = Some(self.index()?),
                CELLS => {
                    let n = self.u32()?;
                    r.cells.reserve(n);
                    for _ in 0..n {
                        r.cells.push(self.cell()?);
                    }
                }
                k => {
                    let v = Some(self.value()?);
                    if k == 0 {
                        self.row_number = to_js_string(v.as_ref());
                    }
                    match k {
                        0 => r.number = v,
                        1 => r.min = v,
                        2 => r.max = v,
                        3 => r.height = v,
                        4 => r.hidden = v,
                        5 => r.outline_level = v,
                        6 => r.collapsed = v,
                        7 => r.style_id = v,
                        BEST_FIT => r.best_fit = v,
                        STYLE_VALUE => r.style_value = v,
                        k => return Err(format!("rows: bad row key {k}")),
                    }
                }
            }
        }
    }
}

/// Rows decoded one at a time from the layout of [`decode_sheets`], or from
/// one sheet's rows encoded by [`encode_row`].
pub(crate) struct RowCursor<'a> {
    b: Bin<'a>,
    /// rows left in the current sheet
    left: usize,
    /// sheets not started yet
    sheets: usize,
}

impl<'a> RowCursor<'a> {
    /// Over `data` in the layout of [`decode_sheets`].
    pub fn sheets(data: &'a [u8]) -> R<Self> {
        let mut b = Bin {
            data,
            pos: 0,
            row_number: String::new(),
        };
        let sheets = b.u32()?;
        Ok(RowCursor { b, left: 0, sheets })
    }

    /// Over `count` rows encoded by [`encode_row`].
    pub fn rows(data: &'a [u8], count: usize) -> Self {
        RowCursor {
            b: Bin {
                data,
                pos: 0,
                row_number: String::new(),
            },
            left: count,
            sheets: 0,
        }
    }

    /// Moves to the next sheet (skipping what is left of this one); false if
    /// there is none.
    pub fn next_sheet(&mut self) -> R<bool> {
        while self.left > 0 {
            self.next_row()?;
        }
        if self.sheets == 0 {
            return Ok(false);
        }
        self.sheets -= 1;
        self.left = self.b.u32()?;
        Ok(true)
    }

    /// The next row of the current sheet.
    pub fn next_row(&mut self) -> R<Option<RowModel>> {
        if self.left == 0 {
            return Ok(None);
        }
        self.left -= 1;
        self.b.row().map(Some)
    }
}

/// Encodes `sheets` in the layout of [`decode_sheets`].
pub(crate) fn encode_sheets(sheets: &[Vec<RowModel>]) -> Vec<u8> {
    let mut b = BinOut(Vec::new());
    b.u32(sheets.len());
    for rows in sheets {
        b.u32(rows.len());
        for r in rows {
            b.row(r);
        }
    }
    b.0
}

/// Decodes a u32 sheet count and that many sheets of rows.
pub(crate) fn decode_sheets(data: &[u8]) -> R<Vec<Vec<RowModel>>> {
    let mut b = Bin {
        data,
        pos: 0,
        row_number: String::new(),
    };
    let sheets = b.u32()?;
    let mut out = Vec::with_capacity(sheets);
    for _ in 0..sheets {
        let n = b.u32()?;
        let mut rows = Vec::with_capacity(n);
        for _ in 0..n {
            rows.push(b.row()?);
        }
        out.push(rows);
    }
    if b.pos != data.len() {
        return Err("rows: trailing data".into());
    }
    Ok(out)
}

// ---------------------------------------------------------------------------
// Read results: typed rows to JS (binary, same layout) and to model JSON

const REF: u8 = 9;
/// Value tag: shared string `u32` (a plain string, sent once in the header).
const SST: u8 = 10;
/// Value tag of a cell `address` that is its column's letters and its row's
/// number: `u32` column (`lib/xlsx/rows.js` rebuilds the string).
const ADDR: u8 = 11;
/// Key of the reader's `style` (a value, unlike the writers' `STYLE` index).
const STYLE_VALUE: u8 = 29;
/// Row key of `bestFit`.
const BEST_FIT: u8 = 8;

/// The `rows` of a read workbook's worksheets, in `model.worksheets` order,
/// and the plain shared strings their `sst` indexes refer to.
#[derive(Debug, Default)]
pub struct SheetRows(pub(crate) Vec<Rows>, pub(crate) Vec<String>);

/// One worksheet's rows, encoded by [`encode_row`] as they are reconciled.
#[derive(Debug, Default)]
pub(crate) struct Rows {
    pub count: usize,
    pub bytes: Vec<u8>,
}

/// The column of `address` if it is canonical (`[A-Z]{1,3}` letters of a
/// column up to `XFD`, then exactly `row`).
fn address_column(address: Option<&Value>, row: Option<&str>) -> Option<u32> {
    let a = address?.as_str()?;
    let letters = a.bytes().take_while(u8::is_ascii_uppercase).count();
    if !(1..=3).contains(&letters) || a.get(letters..)? != row? {
        return None;
    }
    let col = a
        .bytes()
        .take(letters)
        .fold(0u32, |n, b| n * 26 + u32::from(b - b'A' + 1));
    (col <= 16384).then_some(col)
}

/// Appends `r` to `out` in the layout of [`SheetRows::encode`].
pub(crate) fn encode_row(out: &mut Vec<u8>, r: &RowModel) {
    let mut b = BinOut(std::mem::take(out));
    b.row(r);
    *out = b.0;
}

struct BinOut(Vec<u8>);

impl BinOut {
    /// The writers' `STYLE` index (the readers have none).
    fn style_index(&mut self, style: Option<usize>) {
        if let Some(i) = style {
            self.0.push(STYLE);
            self.0.push(INDEX);
            self.u32(i);
        }
    }

    fn row(&mut self, r: &RowModel) {
        let number = r
            .number
            .as_ref()
            .and_then(Value::as_f64)
            .filter(|n| *n >= 1.0 && n.fract() == 0.0)
            .map(crate::jsv::number_to_string);
        self.field(0, r.number.as_ref());
        self.field(1, r.min.as_ref());
        self.field(2, r.max.as_ref());
        self.field(3, r.height.as_ref());
        self.field(4, r.hidden.as_ref());
        self.field(5, r.outline_level.as_ref());
        self.field(6, r.collapsed.as_ref());
        self.field(7, r.style_id.as_ref());
        self.field(BEST_FIT, r.best_fit.as_ref());
        self.field(STYLE_VALUE, r.style_value.as_ref());
        self.style_index(r.style);
        self.0.push(CELLS);
        self.u32(r.cells.len());
        for c in &r.cells {
            match (
                address_column(c.address.as_deref(), number.as_deref()),
                &c.address,
            ) {
                (Some(col), _) => {
                    self.0.push(0);
                    self.0.push(ADDR);
                    self.u32(col as usize);
                }
                (None, a) => self.field(0, a.as_deref()),
            }
            self.number(1, c.ty);
            if let Some(i) = c.sst {
                self.0.push(2);
                self.0.push(SST);
                self.u32(i as usize);
            }
            self.field(2, c.value.as_deref());
            self.field(3, c.comment.as_deref());
            self.field(4, c.text.as_deref());
            self.field(5, c.hyperlink.as_deref());
            self.field(6, c.tooltip.as_deref());
            self.field(7, c.master.as_deref());
            self.field(8, c.share_type.as_deref());
            self.field(9, c.formula.as_deref());
            self.field(10, c.shared_formula.as_deref());
            self.field(11, c.result.as_deref());
            self.field(12, c.ref_.as_deref());
            self.field(13, c.si.as_deref());
            self.field(14, c.range.as_deref());
            self.field(15, c.date1904.as_deref());
            self.field(16, c.ss_id.as_deref());
            self.field(17, c.style_id.as_deref());
            self.field(STYLE_VALUE, c.style_value.as_deref());
            self.style_index(c.style);
            if let Some(i) = c.style_ref {
                self.0.push(STYLE_VALUE);
                self.0.push(REF);
                self.u32(i as usize);
            }
            self.0.push(END);
        }
        self.0.push(END);
    }

    fn u32(&mut self, n: usize) {
        let n = u32::try_from(n).unwrap_or(u32::MAX);
        self.0.extend_from_slice(&n.to_le_bytes());
    }
    fn f64(&mut self, x: f64) {
        self.0.extend_from_slice(&x.to_le_bytes());
    }
    fn str(&mut self, s: &str) {
        self.u32(s.len());
        self.0.extend_from_slice(s.as_bytes());
    }
    fn value(&mut self, v: &Value) {
        match v {
            Value::Null => self.0.push(NULL),
            Value::Bool(true) => self.0.push(TRUE),
            Value::Bool(false) => self.0.push(FALSE),
            Value::Number(n) => {
                self.0.push(F64);
                self.f64(n.as_f64().unwrap_or(f64::NAN));
            }
            Value::String(s) => {
                self.0.push(STR);
                self.str(s);
            }
            Value::Object(o) => {
                if o.len() == 1 {
                    if let Some(n) = o.get("$ref").and_then(Value::as_u64) {
                        self.0.push(REF);
                        self.u32(usize::try_from(n).unwrap_or(usize::MAX));
                        return;
                    }
                    if let Some(d) = o.get("$date") {
                        self.0.push(DATE);
                        self.f64(d.as_f64().unwrap_or(f64::NAN));
                        return;
                    }
                    if o.contains_key("$num") {
                        self.0.push(F64);
                        self.f64(as_f64(Some(v)).unwrap_or(f64::NAN));
                        return;
                    }
                }
                self.json(v);
            }
            Value::Array(_) => self.json(v),
        }
    }
    fn json(&mut self, v: &Value) {
        self.0.push(JSON_VALUE);
        if v.as_object().is_some_and(serde_json::Map::is_empty) {
            // every row's default style
            self.str("{}");
        } else {
            self.str(&v.to_string());
        }
    }
    fn field(&mut self, key: u8, v: Option<&Value>) {
        if let Some(v) = v {
            self.0.push(key);
            self.value(v);
        }
    }
    fn number(&mut self, key: u8, v: Option<f64>) {
        if let Some(x) = v {
            self.0.push(key);
            self.0.push(F64);
            self.f64(x);
        }
    }
}

impl SheetRows {
    /// The rows for `lib/xlsx/rows.js` `RowsReader`: the plain shared
    /// strings, then per sheet its byte length and the writers' binary
    /// layout of its rows, plus `STYLE_VALUE`, row `BEST_FIT`, and `REF` and
    /// `SST` values (shared-object table entries and shared strings).
    pub fn encode(&self) -> Vec<u8> {
        let mut b = BinOut(Vec::new());
        b.u32(self.1.len());
        for s in &self.1 {
            b.str(s);
        }
        b.u32(self.0.len());
        for rows in &self.0 {
            // byte length of the sheet, patched below
            let start = b.0.len();
            b.u32(0);
            b.u32(rows.count);
            b.0.extend_from_slice(&rows.bytes);
            let len = u32::try_from(b.0.len() - start - 4).unwrap_or(u32::MAX);
            if let Some(slot) = b.0.get_mut(start..start + 4) {
                slot.copy_from_slice(&len.to_le_bytes());
            }
        }
        b.0
    }

    /// Each worksheet's `rows` as model JSON (encoded rows are decoded, so
    /// this also checks the encoding).
    pub fn to_values(&self) -> Vec<Value> {
        self.0
            .iter()
            .map(|rows| {
                let mut b = Bin {
                    data: &rows.bytes,
                    pos: 0,
                    row_number: String::new(),
                };
                let typed = (0..rows.count)
                    .map(|_| b.row())
                    .collect::<R<Vec<_>>>()
                    .unwrap_or_default();
                Value::Array(typed.iter().map(|r| row_value_with(r, &self.1)).collect())
            })
            .collect()
    }
}

#[cfg(test)]
pub(crate) fn row_value(r: &RowModel) -> Value {
    row_value_with(r, &[])
}

fn row_value_with(r: &RowModel, sst: &[String]) -> Value {
    let mut m = Map::new();
    let mut put = |k: &str, v: Option<&Value>| {
        if let Some(v) = v {
            m.insert(k.to_owned(), v.clone());
        }
    };
    put("number", r.number.as_ref());
    put("min", r.min.as_ref());
    put("max", r.max.as_ref());
    put("styleId", r.style_id.as_ref());
    put("hidden", r.hidden.as_ref());
    put("bestFit", r.best_fit.as_ref());
    put("height", r.height.as_ref());
    put("outlineLevel", r.outline_level.as_ref());
    put("collapsed", r.collapsed.as_ref());
    put("style", r.style_value.as_ref());
    m.insert(
        "cells".into(),
        Value::Array(r.cells.iter().map(|c| cell_value(c, sst)).collect()),
    );
    Value::Object(m)
}

fn cell_value(c: &CellModel, sst: &[String]) -> Value {
    let mut m = Map::new();
    let mut put = |k: &str, v: Option<&Value>| {
        if let Some(v) = v {
            m.insert(k.to_owned(), v.clone());
        }
    };
    put("address", c.address.as_deref());
    put("type", c.ty.map(crate::jsv::num).as_ref());
    let shared = c
        .sst
        .and_then(|i| sst.get(i as usize))
        .map(|s| Value::String(s.clone()));
    put("value", shared.as_ref().or(c.value.as_deref()));
    put("comment", c.comment.as_deref());
    put("text", c.text.as_deref());
    put("hyperlink", c.hyperlink.as_deref());
    put("tooltip", c.tooltip.as_deref());
    put("master", c.master.as_deref());
    put("shareType", c.share_type.as_deref());
    put("formula", c.formula.as_deref());
    put("sharedFormula", c.shared_formula.as_deref());
    put("result", c.result.as_deref());
    put("ref", c.ref_.as_deref());
    put("si", c.si.as_deref());
    put("range", c.range.as_deref());
    put("date1904", c.date1904.as_deref());
    put("ssId", c.ss_id.as_deref());
    put("styleId", c.style_id.as_deref());
    put("style", c.style_value.as_deref());
    let style_ref = c.style_ref.map(|i| json!({ "$ref": i }));
    put("style", style_ref.as_ref());
    Value::Object(m)
}
