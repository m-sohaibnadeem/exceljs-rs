//! Writing ExcelJS's workbook model as `.xlsx`: the port of `XLSX.write`
//! (`prepareModel` and the `add*` part writers in `lib/xlsx/xlsx.js`), the
//! style manager role of `StylesXform`, the writer role of
//! `SharedStringsXform`, and the `prepare`/`render` of the worksheet xforms.

use std::collections::HashMap;

use serde_json::{json, Map, Value};
use zip::write::SimpleFileOptions;
use zip::ZipWriter;

use crate::attrs;
use crate::col_cache::{decode_address, decode_ex, n2l_str, DecodedEx};
use crate::error::{Error, Result};
use crate::jsv::{as_f64, as_str, get, number_to_string, to_js_string, truthy};
use crate::model_json::{as_buffer, base64_decode};
use crate::rows::{
    finish_master_range, prepare_cell, prepare_row, render_row, CellContext, CellPrep, MasterRef,
    Masters, Merges, RowCursor, RowModel, StyleTable,
};
use crate::xform::base::ListXform;
use crate::xform::book::{prepare as prepare_workbook, sheet_summary, WorkbookXform};
use crate::xform::cell::vt;
use crate::xform::cf::{ext_lst, is_ext, uuid_v4, ConditionalFormattingsXform};
use crate::xform::comment::{render_comments, render_vml_notes};
use crate::xform::core::{render_app, render_content_types, render_core, render_relationships};
use crate::xform::drawing::{anchor_type, drawing as drawing_xform};
use crate::xform::sheet::{
    auto_filter, col, dimension, drawing as drawing_ref, hyperlink, merge_cell, page_margins,
    page_setup, picture, prepare_sheet_view, print_options, sheet_format_properties,
    sheet_protection, table_part, HeaderFooterXform, RowBreaksXform, SheetPropertiesXform,
    SheetViewXform,
};
use crate::xform::strings::RichTextXform;
use crate::xform::strings::{render_shared_strings, shared_string_xml};
use crate::xform::style::{
    cell_styles_static, default_fmt_id, ext_lst_static, table_styles_static, BorderXform, DxfXform,
    FillXform, FontXform, NumFmtXform, StyleXform, FONT_OPTIONS,
};
use crate::xform::table::{prepare_table, TableXform};
use crate::xform::{to_xml, Xform};
use crate::xml_stream::{std_doc_attributes, XmlStream};
use crate::zip_thread::{ZipArchiveOut, ZipThread};

const REL_OFFICE_DOCUMENT: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const REL_CORE: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
const REL_EXTENDED: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
const REL_STYLES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const REL_THEME: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
const REL_SHARED_STRINGS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
const REL_WORKSHEET: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const REL_HYPERLINK: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const REL_COMMENTS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const REL_VML: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
const REL_IMAGE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const REL_TABLE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
const REL_DRAWING: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";

/// ExcelJS's default theme (`lib/xlsx/xml/theme1.js`, which the writers use).
const THEME1: &str = include_str!("theme1.xml");

/// Options of `workbook.xlsx.writeBuffer(options)` the writer uses.
#[derive(Debug, Clone, Default)]
pub struct WriteOptions {
    /// `useStyles` (default true).
    pub use_styles: Option<bool>,
    /// `useSharedStrings` (default true): false writes strings in the cells.
    pub use_shared_strings: Option<bool>,
    /// `zip.compression` level (deflate), when given.
    pub compression_level: Option<i64>,
    /// Current time in ms, for `created`/`modified` defaults.
    pub now_ms: f64,
}

type R<T> = std::result::Result<T, String>;

// ---------------------------------------------------------------------------
// Style manager (StylesXform in its writer role, and StylesXform.Mock)

#[derive(Debug)]
pub(crate) struct StyleManager {
    mock: bool,
    styles: Vec<String>,
    style_index: HashMap<String, usize>,
    num_fmts: Vec<String>,
    num_fmt_index: HashMap<String, f64>,
    fonts: Vec<String>,
    font_index: HashMap<String, usize>,
    borders: Vec<String>,
    border_index: HashMap<String, usize>,
    fills: Vec<String>,
    fill_index: HashMap<String, usize>,
    dxfs: Vec<Value>,
    mock_styles: Vec<Value>,
    mock_date_style: Option<usize>,
    font_xform: FontXform,
    border_xform: BorderXform,
    fill_xform: FillXform,
    style_xform: StyleXform,
}

impl StyleManager {
    pub fn new(use_styles: bool) -> Self {
        let mut m = StyleManager {
            mock: !use_styles,
            styles: Vec::new(),
            style_index: HashMap::new(),
            num_fmts: Vec::new(),
            num_fmt_index: HashMap::new(),
            fonts: Vec::new(),
            font_index: HashMap::new(),
            borders: Vec::new(),
            border_index: HashMap::new(),
            fills: Vec::new(),
            fill_index: HashMap::new(),
            dxfs: Vec::new(),
            mock_styles: vec![
                json!({"numFmtId": 0, "fontId": 0, "fillId": 0, "borderId": 0, "xfId": 0}),
            ],
            mock_date_style: None,
            font_xform: FontXform::new(FONT_OPTIONS),
            border_xform: BorderXform::new(),
            fill_xform: FillXform::new(),
            style_xform: StyleXform::new(true),
        };
        if use_styles {
            m.add_border(&json!({}));
            m.add_style(
                &json!({"numFmtId": 0, "fontId": 0, "fillId": 0, "borderId": 0, "xfId": 0}),
            );
            m.add_fill(&json!({"type": "pattern", "pattern": "none"}));
            m.add_fill(&json!({"type": "pattern", "pattern": "gray125"}));
        }
        m
    }

    fn intern(xml: String, list: &mut Vec<String>, index: &mut HashMap<String, usize>) -> usize {
        if let Some(&i) = index.get(&xml) {
            return i;
        }
        let i = list.len();
        index.insert(xml.clone(), i);
        list.push(xml);
        i
    }

    fn add_style(&mut self, style: &Value) -> usize {
        let xml = to_xml(&self.style_xform, Some(style));
        Self::intern(xml, &mut self.styles, &mut self.style_index)
    }

    fn add_font(&mut self, font: &Value) -> usize {
        let xml = to_xml(&self.font_xform, Some(font));
        Self::intern(xml, &mut self.fonts, &mut self.font_index)
    }

    fn add_border(&mut self, border: &Value) -> usize {
        let xml = to_xml(&self.border_xform, Some(border));
        Self::intern(xml, &mut self.borders, &mut self.border_index)
    }

    fn add_fill(&mut self, fill: &Value) -> usize {
        let xml = to_xml(&self.fill_xform, Some(fill));
        Self::intern(xml, &mut self.fills, &mut self.fill_index)
    }

    fn add_num_fmt_str(&mut self, code: &str) -> f64 {
        if let Some(id) = default_fmt_id(code) {
            return f64::from(id);
        }
        if let Some(&id) = self.num_fmt_index.get(code) {
            return id;
        }
        #[allow(clippy::cast_precision_loss)]
        let id = 164.0 + self.num_fmts.len() as f64;
        self.num_fmt_index.insert(code.to_owned(), id);
        let xml = to_xml(
            &NumFmtXform::default(),
            Some(&json!({"id": crate::jsv::num(id), "formatCode": code})),
        );
        self.num_fmts.push(xml);
        id
    }

    fn default_font(&mut self) {
        if self.fonts.is_empty() {
            self.add_font(&json!({"size": 11, "color": {"theme": 1}, "name": "Calibri", "family": 2, "scheme": "minor"}));
        }
    }

    /// `addStyleModel(model, cellType)`.
    pub fn add_style_model(&mut self, model: Option<&Value>, cell_type: i64) -> usize {
        if self.mock {
            if cell_type == vt::DATE {
                return *self.mock_date_style.get_or_insert_with(|| {
                    self.mock_styles.push(json!({"numFmtId": 14}));
                    self.mock_styles.len() - 1
                });
            }
            return 0;
        }
        if !truthy(model) {
            return 0;
        }
        self.default_font();
        let mut style = Map::new();
        let cell_type = if cell_type <= 0 {
            vt::NUMBER
        } else {
            cell_type
        };
        if truthy(get(model, "numFmt")) {
            let id = self.add_num_fmt_str(&to_js_string(get(model, "numFmt")));
            style.insert("numFmtId".into(), crate::jsv::num(id));
        } else if cell_type == vt::NUMBER {
            style.insert(
                "numFmtId".into(),
                crate::jsv::num(self.add_num_fmt_str("General")),
            );
        } else if cell_type == vt::DATE {
            style.insert(
                "numFmtId".into(),
                crate::jsv::num(self.add_num_fmt_str("mm-dd-yy")),
            );
        }
        if let Some(f) = get(model, "font").filter(|v| truthy(Some(v))) {
            style.insert("fontId".into(), json!(self.add_font(f)));
        }
        if let Some(b) = get(model, "border").filter(|v| truthy(Some(v))) {
            style.insert("borderId".into(), json!(self.add_border(b)));
        }
        if let Some(f) = get(model, "fill").filter(|v| truthy(Some(v))) {
            style.insert("fillId".into(), json!(self.add_fill(f)));
        }
        for k in ["alignment", "protection"] {
            if let Some(v) = get(model, k).filter(|v| truthy(Some(v))) {
                style.insert(k.into(), v.clone());
            }
        }
        self.add_style(&Value::Object(style))
    }

    /// `addDxfStyle(style)`.
    pub fn add_dxf_style(&mut self, style: &Value) -> R<usize> {
        if self.mock {
            return Err("Cannot read properties of undefined (reading 'push')".into());
        }
        let mut style = style.clone();
        if let Some(code) = get(Some(&style), "numFmt")
            .filter(|v| truthy(Some(v)))
            .map(|v| to_js_string(Some(v)))
        {
            let id = self.add_num_fmt_str(&code);
            if let Value::Object(o) = &mut style {
                o.insert("numFmtId".into(), crate::jsv::num(id));
            }
        }
        self.dxfs.push(style);
        Ok(self.dxfs.len() - 1)
    }

    /// `StylesXform.render` (or the Mock's).
    pub fn render(&mut self) -> String {
        let mut xml = XmlStream::new();
        xml.open_xml(&std_doc_attributes());
        xml.open_node(
            "styleSheet",
            &attrs![
                (
                    "xmlns",
                    Some("http://schemas.openxmlformats.org/spreadsheetml/2006/main")
                ),
                (
                    "xmlns:mc",
                    Some("http://schemas.openxmlformats.org/markup-compatibility/2006")
                ),
                ("mc:Ignorable", Some("x14ac x16r2")),
                (
                    "xmlns:x14ac",
                    Some("http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac")
                ),
                (
                    "xmlns:x16r2",
                    Some("http://schemas.microsoft.com/office/spreadsheetml/2015/02/main")
                ),
            ],
        );
        let cell_style_xfs =
            json!([{"numFmtId": 0, "fontId": 0, "fillId": 0, "borderId": 0, "xfId": 0}]);
        let style_xfs_list =
            ListXform::new("cellStyleXfs", Box::new(StyleXform::new(false))).count();
        if self.mock {
            // Plain JSON model, rendered through the list xforms.
            let lists: [(&str, Box<dyn Xform>, Value); 4] = [
                ("numFmts", Box::new(NumFmtXform::default()), json!([])),
                (
                    "fonts",
                    Box::new(FontXform::new(FONT_OPTIONS)),
                    json!([{"size": 11, "color": {"theme": 1}, "name": "Calibri", "family": 2, "scheme": "minor"}]),
                ),
                (
                    "fills",
                    Box::new(FillXform::new()),
                    json!([{"type": "pattern", "pattern": "none"}, {"type": "pattern", "pattern": "gray125"}]),
                ),
                ("borders", Box::new(BorderXform::new()), json!([{}])),
            ];
            for (tag, child, model) in lists {
                let mut l = ListXform::new(tag, child).count();
                if tag == "fonts" {
                    l = l.attrs(attrs![("x14ac:knownFonts", Some(1i64))]);
                }
                l.render_list(&mut xml, Some(&model));
            }
            style_xfs_list.render_list(&mut xml, Some(&cell_style_xfs));
            ListXform::new("cellXfs", Box::new(StyleXform::new(true)))
                .count()
                .render_list(&mut xml, Some(&Value::Array(self.mock_styles.clone())));
        } else {
            if !self.num_fmts.is_empty() {
                xml.open_node("numFmts", &attrs![("count", Some(self.num_fmts.len()))]);
                for x in &self.num_fmts {
                    xml.write_xml(x);
                }
                xml.close_node();
            }
            self.default_font();
            for (tag, list, extra) in [
                ("fonts", &self.fonts, true),
                ("fills", &self.fills, false),
                ("borders", &self.borders, false),
            ] {
                let mut a = attrs![("count", Some(list.len()))];
                if extra {
                    a.push(("x14ac:knownFonts".into(), Some("1".into())));
                }
                xml.open_node(tag, &a);
                for x in list {
                    xml.write_xml(x);
                }
                xml.close_node();
            }
            style_xfs_list.render_list(&mut xml, Some(&cell_style_xfs));
            xml.open_node("cellXfs", &attrs![("count", Some(self.styles.len()))]);
            for x in &self.styles {
                xml.write_xml(x);
            }
            xml.close_node();
        }
        cell_styles_static().render(&mut xml, None);
        let dxfs = if self.mock {
            None
        } else {
            Some(Value::Array(self.dxfs.clone()))
        };
        ListXform::new("dxfs", Box::new(DxfXform::new()))
            .always()
            .count()
            .render_list(&mut xml, dxfs.as_ref());
        table_styles_static().render(&mut xml, None);
        ext_lst_static().render(&mut xml, None);
        xml.close_node();
        xml.into_xml()
    }
}

// ---------------------------------------------------------------------------
// Shared strings (writer role)

#[derive(Debug, Default)]
pub(crate) struct SharedStrings {
    pub values: Vec<Value>,
    pub count: usize,
    hash: HashMap<String, usize>,
    rich: HashMap<String, usize>,
}

impl SharedStrings {
    /// `add(value)`. Rich text is keyed by its XML, plain values by
    /// `String(value)`. (The streaming writer's `utils/shared-strings.js`
    /// keys rich text as `"[object Object]"` too: B11 in COMPATIBILITY.md.)
    pub fn add(&mut self, value: &Value) -> usize {
        let (key, map) = if truthy(get(Some(value), "richText")) {
            (shared_string_xml(value), &mut self.rich)
        } else {
            (to_js_string(Some(value)), &mut self.hash)
        };
        let index = match map.get(&key) {
            Some(&i) => i,
            None => {
                let i = self.values.len();
                map.insert(key, i);
                self.values.push(value.clone());
                i
            }
        };
        self.count += 1;
        index
    }
}

// ---------------------------------------------------------------------------
// Ranges (doc/range.js) for merges and shared formulae

#[derive(Debug, Clone, Copy, Default)]
pub(crate) struct RangeBox {
    top: f64,
    left: f64,
    bottom: f64,
    right: f64,
}

fn or1(v: f64) -> f64 {
    if v == 0.0 || v.is_nan() {
        1.0
    } else {
        v
    }
}

impl RangeBox {
    pub(crate) fn from_model(m: Option<&Value>) -> Self {
        let n = |k: &str| as_f64(get(m, k)).unwrap_or(0.0);
        RangeBox {
            top: n("top"),
            left: n("left"),
            bottom: n("bottom"),
            right: n("right"),
        }
    }

    /// `new Range(tl, br)` / `new Range("tl:br")`.
    pub(crate) fn from_addresses(tl: &str, br: &str) -> R<Self> {
        let a = decode_address(tl)?;
        let b = decode_address(br)?;
        let min = |x: Option<f64>, y: Option<f64>| match (x, y) {
            (Some(x), Some(y)) => x.min(y),
            _ => f64::NAN,
        };
        let max = |x: Option<f64>, y: Option<f64>| match (x, y) {
            (Some(x), Some(y)) => x.max(y),
            _ => f64::NAN,
        };
        Ok(RangeBox {
            top: min(a.row, b.row),
            left: min(a.col, b.col),
            bottom: max(a.row, b.row),
            right: max(a.col, b.col),
        })
    }

    fn expand(&mut self, top: f64, left: f64, bottom: f64, right: f64) {
        let set = |cur: &mut f64, v: f64, better: bool| {
            if *cur == 0.0 || cur.is_nan() || better {
                *cur = v;
            }
        };
        let (t, l, b, r) = (
            or1(self.top),
            or1(self.left),
            or1(self.bottom),
            or1(self.right),
        );
        set(&mut self.top, top, top < t);
        set(&mut self.left, left, left < l);
        set(&mut self.bottom, bottom, bottom > b);
        set(&mut self.right, right, right > r);
    }

    pub(crate) fn expand_to_address(&mut self, address: &str) -> R<()> {
        let (row, col) = match decode_ex(address)? {
            DecodedEx::Address { address, .. } => (
                address.row.unwrap_or(f64::NAN),
                address.col.unwrap_or(f64::NAN),
            ),
            _ => (f64::NAN, f64::NAN),
        };
        self.expand(row, col, row, col);
        Ok(())
    }

    /// `range` getter (no sheet name).
    /// `{top, left, bottom, right}` (a Range's model).
    pub(crate) fn to_model(self) -> Value {
        json!({"top": self.top, "left": self.left, "bottom": self.bottom, "right": self.right})
    }

    pub(crate) fn range(&self) -> R<String> {
        Ok(format!(
            "{}{}:{}{}",
            n2l_str(or1(self.left))?,
            number_to_string(or1(self.top)),
            n2l_str(or1(self.right))?,
            number_to_string(or1(self.bottom))
        ))
    }
}

// ---------------------------------------------------------------------------
// Worksheet prepare

struct SheetPrepare<'a> {
    styles: &'a mut StyleManager,
    table: &'a mut StyleTable,
    shared_strings: Option<&'a mut SharedStrings>,
    date1904: bool,
    media: &'a [Value],
    drawings_count: &'a mut usize,
    drawings: &'a mut Vec<Value>,
    comment_refs: &'a mut Vec<Value>,
    /// `this.preImageId` of the shared WorksheetXform.
    pre_image_id: &'a mut Option<Value>,
}

fn obj_mut(v: &mut Value) -> R<&mut Map<String, Value>> {
    v.as_object_mut()
        .ok_or_else(|| "model is not an object".to_owned())
}

/// A worksheet's rows after `prepare`, encoded again (`rows::encode_row`) so
/// that only one row is decoded at a time, and what clones changed in their
/// masters, by (row, cell) position.
#[derive(Default)]
struct PreparedRows {
    count: usize,
    bytes: Vec<u8>,
    masters: Masters,
}

/// `WorkSheetXform.prepare(model, options)`, reading the rows from `rows`
/// one at a time.
fn prepare_worksheet(
    ws: &mut Value,
    rows: &mut RowCursor<'_>,
    o: &mut SheetPrepare<'_>,
) -> R<PreparedRows> {
    let mut hyperlinks: Vec<Value> = Vec::new();
    let mut comments: Vec<Value> = Vec::new();
    let mut merges = Merges::default();
    let mut formulae: HashMap<String, MasterRef> = HashMap::new();
    let mut si_formulae: i64 = 0;
    let mut prepared = PreparedRows::default();

    // cols
    if let Some(cols) = ws.get_mut("cols") {
        prepare_cols(cols, o.styles);
    }

    // rows and cells
    let mut ri = 0;
    while let Some(mut row) = rows.next_row()? {
        prepare_row(&mut row, o.styles, o.table);
        for ci in 0..row.cells.len() {
            let Some(cell) = row.cells.get_mut(ci) else {
                continue;
            };
            let mut ctx = CellContext {
                styles: o.styles,
                table: o.table,
                shared_strings: o.shared_strings.as_deref_mut(),
                date1904: o.date1904,
            };
            let prep = prepare_cell(
                cell,
                &mut ctx,
                &mut hyperlinks,
                &mut comments,
                &mut merges,
                &formulae,
                &mut si_formulae,
            )?;
            match prep {
                CellPrep::None => {}
                CellPrep::Register(address) => {
                    formulae.insert(address, (ri, ci));
                    prepared.masters.register((ri, ci), cell);
                }
                CellPrep::Clone {
                    master,
                    master_address,
                    address,
                } => {
                    let clone_si = prepared.masters.share(
                        master,
                        &master_address,
                        &address,
                        &mut si_formulae,
                    )?;
                    if let Some(c) = row.cells.get_mut(ci) {
                        c.si = Some(Box::new(clone_si));
                    }
                }
            }
        }
        crate::rows::encode_row(&mut prepared.bytes, &row);
        prepared.count += 1;
        ri += 1;
    }
    prepare_cfs(ws.get_mut("conditionalFormattings"), o.styles)?;

    let m = obj_mut(ws)?;
    m.insert("mergeCells".into(), Value::Array(merges.ranges()?));

    // relationships
    let mut rels: Vec<Value> = Vec::new();
    let next_rid = |r: &Vec<Value>| format!("rId{}", r.len() + 1);
    for h in &mut hyperlinks {
        let rid = next_rid(&rels);
        crate::jsv::set(h, "rId", json!(rid));
        rels.push(json!({"Id": rid, "Type": REL_HYPERLINK, "Target": h.get("target").cloned().unwrap_or(Value::Null), "TargetMode": "External"}));
    }
    let id = to_js_string(m.get("id"));
    if !comments.is_empty() {
        rels.push(json!({"Id": next_rid(&rels), "Type": REL_COMMENTS, "Target": format!("../comments{id}.xml")}));
        rels.push(json!({"Id": next_rid(&rels), "Type": REL_VML, "Target": format!("../drawings/vmlDrawing{id}.vml")}));
        for c in &mut comments {
            let r = to_js_string(c.get("ref"));
            let a = decode_address(&r)?;
            crate::jsv::set(
                c,
                "refAddress",
                json!({"address": a.address, "col": a.col, "row": a.row}),
            );
        }
        o.comment_refs.push(json!({"commentName": format!("comments{id}"), "vmlDrawing": format!("vmlDrawing{id}")}));
    }

    let mut drawing_rels_hash: HashMap<usize, String> = HashMap::new();
    let media_list: Vec<Value> = m
        .get("media")
        .and_then(Value::as_array)
        .cloned()
        .unwrap_or_default();
    for medium in &media_list {
        let image_id = get(Some(medium), "imageId");
        let book_image = image_id
            .and_then(Value::as_u64)
            .and_then(|i| usize::try_from(i).ok())
            .and_then(|i| o.media.get(i))
            .cloned()
            .ok_or_else(|| "Cannot read properties of undefined (reading 'name')".to_owned())?;
        let media_target = format!(
            "../media/{}.{}",
            to_js_string(get(Some(&book_image), "name")),
            to_js_string(get(Some(&book_image), "extension"))
        );
        match as_str(get(Some(medium), "type")) {
            Some("background") => {
                let rid = next_rid(&rels);
                rels.push(json!({"Id": rid, "Type": REL_IMAGE, "Target": media_target}));
                m.insert("background".into(), json!({"rId": rid}));
                m.insert("image".into(), book_image.clone());
            }
            Some("image") => {
                if !truthy(m.get("drawing")) {
                    *o.drawings_count += 1;
                    let d = json!({"rId": next_rid(&rels), "name": format!("drawing{}", o.drawings_count), "anchors": [], "rels": []});
                    rels.push(json!({"Id": d.get("rId"), "Type": REL_DRAWING, "Target": format!("../drawings/{}.xml", d.get("name").and_then(Value::as_str).unwrap_or(""))}));
                    m.insert("drawing".into(), d);
                }
                let drawing = m.get_mut("drawing").ok_or_else(|| "drawing".to_owned())?;
                let drels_len = drawing
                    .get("rels")
                    .and_then(Value::as_array)
                    .map_or(0, Vec::len);
                let image_key = image_id
                    .and_then(Value::as_u64)
                    .and_then(|i| usize::try_from(i).ok())
                    .unwrap_or(usize::MAX);
                let mut rid_image = if o.pre_image_id.as_ref() == image_id {
                    drawing_rels_hash.get(&image_key).cloned()
                } else {
                    drawing_rels_hash.get(&drels_len).cloned()
                };
                if rid_image.is_none() {
                    let rid = format!("rId{}", drels_len + 1);
                    drawing_rels_hash.insert(drels_len, rid.clone());
                    if let Some(Value::Array(dr)) = drawing.get_mut("rels") {
                        dr.push(json!({"Id": rid, "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", "Target": media_target}));
                    }
                    rid_image = Some(rid);
                }
                let mut anchor = json!({"picture": {"rId": rid_image}, "range": medium.get("range").cloned().unwrap_or(Value::Null)});
                let hl = get(Some(medium), "hyperlinks");
                if truthy(hl) && truthy(get(hl, "hyperlink")) {
                    let dlen = drawing
                        .get("rels")
                        .and_then(Value::as_array)
                        .map_or(0, Vec::len);
                    let rid = format!("rId{}", dlen + 1);
                    drawing_rels_hash.insert(dlen, rid.clone());
                    let mut h = Map::new();
                    if let Some(t) = get(hl, "tooltip") {
                        h.insert("tooltip".into(), t.clone());
                    }
                    h.insert("rId".into(), json!(rid));
                    if let Some(p) = anchor.get_mut("picture") {
                        crate::jsv::set(p, "hyperlinks", Value::Object(h));
                    }
                    if let Some(Value::Array(dr)) = drawing.get_mut("rels") {
                        dr.push(json!({"Id": rid, "Type": REL_HYPERLINK, "Target": get(hl, "hyperlink").cloned().unwrap_or(Value::Null), "TargetMode": "External"}));
                    }
                }
                *o.pre_image_id = image_id.cloned();
                if let Some(Value::Array(a)) = drawing.get_mut("anchors") {
                    a.push(anchor);
                }
            }
            _ => {}
        }
    }
    if let Some(d) = m.get("drawing").filter(|d| truthy(Some(d))) {
        if !o.drawings.iter().any(|x| x.get("name") == d.get("name")) {
            o.drawings.push(d.clone());
        }
    }

    // tables
    for table in m
        .get_mut("tables")
        .and_then(Value::as_array_mut)
        .into_iter()
        .flatten()
    {
        let rid = format!("rId{}", rels.len() + 1);
        crate::jsv::set(table, "rId", json!(rid));
        rels.push(json!({"Id": rid, "Type": REL_TABLE, "Target": format!("../tables/{}", to_js_string(table.get("target")))}));
        for column in table
            .get_mut("columns")
            .and_then(Value::as_array_mut)
            .into_iter()
            .flatten()
        {
            if let Some(style) = column.get("style").filter(|s| truthy(Some(s))).cloned() {
                crate::jsv::set(column, "dxfId", json!(o.styles.add_dxf_style(&style)?));
            }
        }
    }

    // ext items (x14 conditional formatting ids)
    for cf in m
        .get_mut("conditionalFormattings")
        .and_then(Value::as_array_mut)
        .into_iter()
        .flatten()
    {
        for rule in cf
            .get_mut("rules")
            .and_then(Value::as_array_mut)
            .into_iter()
            .flatten()
        {
            if is_ext(Some(rule)) {
                crate::jsv::set(
                    rule,
                    "x14Id",
                    json!(format!("{{{}}}", uuid_v4()).to_uppercase()),
                );
            }
        }
    }

    m.insert("hyperlinks".into(), Value::Array(hyperlinks));
    m.insert("comments".into(), Value::Array(comments));
    m.insert("rels".into(), Value::Array(rels));
    Ok(prepared)
}

/// `ListXform(cols).prepare`: style ids for columns.
pub(crate) fn prepare_cols(cols: &mut Value, styles: &mut StyleManager) {
    if let Value::Array(cols) = cols {
        for c in cols {
            let style = get(Some(c), "style")
                .filter(|v| truthy(Some(v)))
                .cloned()
                .unwrap_or_else(|| json!({}));
            let id = styles.add_style_model(Some(&style), -1);
            if id != 0 {
                crate::jsv::set(c, "styleId", json!(id));
            }
        }
    }
}

/// `ConditionalFormattingsXform.prepare`: priorities and dxf ids.
pub(crate) fn prepare_cfs(cfs: Option<&mut Value>, styles: &mut StyleManager) -> R<()> {
    let Some(Value::Array(cfs)) = cfs else {
        return Err("Cannot read properties of undefined (reading 'reduce')".into());
    };
    let mut next = 1.0f64;
    for cf in cfs.iter() {
        for rule in cf
            .get("rules")
            .and_then(Value::as_array)
            .into_iter()
            .flatten()
        {
            let p = get(Some(rule), "priority")
                .filter(|v| truthy(Some(v)))
                .map_or(0.0, |v| crate::jsv::to_number(Some(v)));
            next = next.max(p);
        }
    }
    for cf in cfs.iter_mut() {
        for rule in cf
            .get_mut("rules")
            .and_then(Value::as_array_mut)
            .into_iter()
            .flatten()
        {
            if !truthy(rule.get("priority")) {
                crate::jsv::set(rule, "priority", crate::jsv::num(next));
                next += 1.0;
            }
            if let Some(style) = rule.get("style").filter(|s| truthy(Some(s))).cloned() {
                crate::jsv::set(rule, "dxfId", json!(styles.add_dxf_style(&style)?));
            }
        }
    }
    Ok(())
}

// ---------------------------------------------------------------------------
// Worksheet render

/// Renders the worksheet part. Rows are given to `out` in chunks as they are
/// rendered (so they can be compressed meanwhile); the rest is returned.
fn render_worksheet(
    ws: &Value,
    rows: &PreparedRows,
    out: &mut dyn FnMut(String) -> Result<()>,
) -> std::result::Result<String, WriteFailure> {
    let m = Some(ws);
    let mut xml = XmlStream::new();
    xml.open_xml(&std_doc_attributes());
    xml.open_node(
        "worksheet",
        &attrs![
            (
                "xmlns",
                Some("http://schemas.openxmlformats.org/spreadsheetml/2006/main")
            ),
            (
                "xmlns:r",
                Some("http://schemas.openxmlformats.org/officeDocument/2006/relationships")
            ),
            (
                "xmlns:mc",
                Some("http://schemas.openxmlformats.org/markup-compatibility/2006")
            ),
            ("mc:Ignorable", Some("x14ac")),
            (
                "xmlns:x14ac",
                Some("http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac")
            ),
        ],
    );
    let props = get(m, "properties");
    let sheet_format = if truthy(props) {
        let mut f = Map::new();
        for (k, src) in [
            ("defaultRowHeight", "defaultRowHeight"),
            ("dyDescent", "dyDescent"),
            ("outlineLevelCol", "outlineLevelCol"),
            ("outlineLevelRow", "outlineLevelRow"),
        ] {
            if let Some(v) = get(props, src) {
                f.insert(k.into(), v.clone());
            }
        }
        if truthy(get(props, "defaultColWidth")) {
            f.insert(
                "defaultColWidth".into(),
                get(props, "defaultColWidth")
                    .cloned()
                    .unwrap_or(Value::Null),
            );
        }
        Some(Value::Object(f))
    } else {
        None
    };
    let ps = get(m, "pageSetup");
    let mut sheet_props = Map::new();
    if truthy(props) {
        if let Some(v) = get(props, "outlineProperties") {
            sheet_props.insert("outlineProperties".into(), v.clone());
        }
        if let Some(v) = get(props, "tabColor") {
            sheet_props.insert("tabColor".into(), v.clone());
        }
    }
    if truthy(ps) && truthy(get(ps, "fitToPage")) {
        sheet_props.insert(
            "pageSetup".into(),
            json!({"fitToPage": get(ps, "fitToPage")}),
        );
    }
    let margins = if truthy(ps) { get(ps, "margins") } else { None };
    let mut print = Map::new();
    if truthy(ps) {
        for k in [
            "showRowColHeaders",
            "showGridLines",
            "horizontalCentered",
            "verticalCentered",
        ] {
            if let Some(v) = get(ps, k) {
                print.insert(k.into(), v.clone());
            }
        }
    }

    SheetPropertiesXform::new().render(&mut xml, Some(&Value::Object(sheet_props)));
    // dimensions: a Range ({model: {...}}) or a string
    let dims = match get(m, "dimensions") {
        Some(Value::String(s)) => Some(json!(s)),
        Some(d) if truthy(Some(d)) => Some(json!(RangeBox::from_model(
            get(Some(d), "model").or(Some(d))
        )
        .range()?)),
        _ => None,
    };
    dimension().render(&mut xml, dims.as_ref());
    let mut views = get(m, "views").cloned();
    if let Some(Value::Array(vs)) = &mut views {
        for v in vs.iter_mut() {
            prepare_sheet_view(v);
        }
    }
    ListXform::new("sheetViews", Box::new(SheetViewXform::default()))
        .render_list(&mut xml, views.as_ref());
    sheet_format_properties().render(&mut xml, sheet_format.as_ref());
    ListXform::new("cols", Box::new(col())).render_list(&mut xml, get(m, "cols"));
    // ListXform({tag: 'sheetData', empty: true, childXform: RowXform})
    if rows.count == 0 {
        xml.leaf_node("sheetData", &[], None);
    } else {
        xml.open_node("sheetData", &[]);
        let rich = RichTextXform::new();
        let mut cursor = RowCursor::rows(&rows.bytes, rows.count);
        let mut ri = 0;
        while let Some(mut row) = cursor.next_row()? {
            for (ci, cell) in row.cells.iter_mut().enumerate() {
                rows.masters.apply((ri, ci), cell);
                finish_master_range(cell)?;
            }
            ri += 1;
            render_row(&mut xml, &row, &rich);
            if let Some(chunk) = xml.drain(CHUNK) {
                out(chunk).map_err(WriteFailure::Zip)?;
            }
        }
        xml.close_node();
    }
    sheet_protection().render(&mut xml, get(m, "sheetProtection"));
    auto_filter().render(&mut xml, get(m, "autoFilter"));
    ListXform::new("mergeCells", Box::new(merge_cell()))
        .count()
        .render_list(&mut xml, get(m, "mergeCells"));
    ConditionalFormattingsXform::new().render(&mut xml, get(m, "conditionalFormattings"));
    render_data_validations(&mut xml, get(m, "dataValidations")).map_err(WriteFailure::Render)?;
    ListXform::new("hyperlinks", Box::new(hyperlink())).render_list(&mut xml, get(m, "hyperlinks"));
    print_options().render(&mut xml, Some(&Value::Object(print)));
    page_margins().render(&mut xml, margins);
    page_setup().render(&mut xml, ps);
    HeaderFooterXform::default().render(&mut xml, get(m, "headerFooter"));
    RowBreaksXform::new().render(&mut xml, get(m, "rowBreaks"));
    drawing_ref().render(&mut xml, get(m, "drawing"));
    picture().render(&mut xml, get(m, "background"));
    ListXform::new("tableParts", Box::new(table_part()))
        .count()
        .render_list(&mut xml, get(m, "tables"));
    ext_lst().render(&mut xml, m);
    for rel in get(m, "rels")
        .and_then(Value::as_array)
        .into_iter()
        .flatten()
    {
        if as_str(get(Some(rel), "Type")) == Some(REL_VML) {
            xml.leaf_node(
                "legacyDrawing",
                &attrs![("r:id", get(Some(rel), "Id"))],
                None,
            );
        }
    }
    xml.close_node();
    xml.finish().map_err(WriteFailure::Render)
}

/// Size of the pieces of a worksheet handed to the compressor.
const CHUNK: usize = 1 << 20;

/// Why [`render_worksheet`] stopped.
enum WriteFailure {
    /// ExcelJS's render would throw here.
    Render(String),
    /// The archive failed.
    Zip(Error),
}

impl From<String> for WriteFailure {
    fn from(m: String) -> Self {
        WriteFailure::Render(m)
    }
}

/// `lodash isEqual` as used by `optimiseDataValidations`.
fn js_is_equal(a: &Value, b: &Value) -> bool {
    a == b
}

/// `DataValidationsXform.render(xmlStream, model)`.
pub(crate) fn render_data_validations(xml: &mut XmlStream, model: Option<&Value>) -> R<()> {
    let Some(Value::Object(map)) = model else {
        return Ok(());
    };
    let mut list: Vec<(&String, &Value)> = map.iter().collect();
    list.sort_by(|a, b| a.0.cmp(b.0));
    let mut marked: HashMap<String, bool> =
        list.iter().map(|(k, _)| ((*k).clone(), false)).collect();
    let mut optimized = Vec::new();
    for (address, dv) in &list {
        if marked.get(address.as_str()).copied().unwrap_or(false) {
            continue;
        }
        let with_sqref = |sqref: String| -> Value {
            let mut o = match dv {
                Value::Object(o) => o.clone(),
                _ => Map::new(),
            };
            o.insert("sqref".into(), json!(sqref));
            Value::Object(o)
        };
        let decoded = decode_ex(address)?;
        let (row, col) = match decoded {
            DecodedEx::Range(r) => {
                marked.insert(r.dimensions.clone(), true);
                optimized.push(with_sqref((*address).clone()));
                continue;
            }
            DecodedEx::Address { address: a, .. } => {
                (a.row.unwrap_or(f64::NAN), a.col.unwrap_or(f64::NAN))
            }
            DecodedEx::Error { .. } => (f64::NAN, f64::NAN),
        };
        let enc =
            |r: f64, c: f64| -> R<String> { Ok(format!("{}{}", n2l_str(c)?, number_to_string(r))) };
        let mut height = 1.0;
        loop {
            let other = enc(row + height, col)?;
            match map.get(&other) {
                Some(v) if js_is_equal(dv, v) => height += 1.0,
                _ => break,
            }
        }
        let match_col = |c: f64| -> R<bool> {
            let mut i = 0.0;
            while i < height {
                let other = enc(row + i, c)?;
                match map.get(&other) {
                    Some(v)
                        if js_is_equal(map.get(address.as_str()).unwrap_or(&Value::Null), v) => {}
                    _ => return Ok(false),
                }
                i += 1.0;
            }
            Ok(true)
        };
        let mut width = 1.0;
        while match_col(col + width)? {
            width += 1.0;
        }
        let mut i = 0.0;
        while i < height {
            let mut j = 0.0;
            while j < width {
                let a = enc(row + i, col + j)?;
                if let Some(k) = map.keys().find(|k| **k == a) {
                    marked.insert(k.clone(), true);
                }
                j += 1.0;
            }
            i += 1.0;
        }
        if height > 1.0 || width > 1.0 {
            optimized.push(with_sqref(format!(
                "{address}:{}",
                enc(row + height - 1.0, col + width - 1.0)?
            )));
        } else {
            optimized.push(with_sqref((*address).clone()));
        }
    }
    if optimized.is_empty() {
        return Ok(());
    }
    xml.open_node("dataValidations", &attrs![("count", Some(optimized.len()))]);
    for v in &optimized {
        let v = Some(v);
        xml.open_node("dataValidation", &[]);
        let ty = as_str(get(v, "type"));
        if ty != Some("any") {
            xml.add_attribute("type", get(v, "type"));
            let op = get(v, "operator");
            if truthy(op) && ty != Some("list") && as_str(op) != Some("between") {
                xml.add_attribute("operator", op);
            }
            if truthy(get(v, "allowBlank")) {
                xml.add_attribute("allowBlank", Some("1"));
            }
        }
        if truthy(get(v, "showInputMessage")) {
            xml.add_attribute("showInputMessage", Some("1"));
        }
        for k in ["promptTitle", "prompt"] {
            if truthy(get(v, k)) {
                xml.add_attribute(k, get(v, k));
            }
        }
        if truthy(get(v, "showErrorMessage")) {
            xml.add_attribute("showErrorMessage", Some("1"));
        }
        for k in ["errorStyle", "errorTitle", "error"] {
            if truthy(get(v, k)) {
                xml.add_attribute(k, get(v, k));
            }
        }
        xml.add_attribute("sqref", get(v, "sqref"));
        for (i, f) in get(v, "formulae")
            .and_then(Value::as_array)
            .into_iter()
            .flatten()
            .enumerate()
        {
            xml.open_node(&format!("formula{}", i + 1), &[]);
            if ty == Some("date") {
                let ms = match f {
                    Value::Object(o) => crate::model_json::as_date(o).unwrap_or(f64::NAN),
                    Value::String(s) => crate::model_json::parse_iso_date(s),
                    other => as_f64(Some(other)).unwrap_or(f64::NAN),
                };
                xml.write_text(&number_to_string(crate::model_json::date_to_excel(
                    ms, false,
                )));
            } else {
                xml.write_text(&to_js_string(Some(f)));
            }
            xml.close_node();
        }
        xml.close_node();
    }
    xml.close_node();
    Ok(())
}

// ---------------------------------------------------------------------------
// XLSX.write

fn rels_json(list: &[(&str, &str, &str)]) -> Vec<Value> {
    list.iter()
        .map(|(id, ty, target)| json!({"Id": id, "Type": ty, "Target": target}))
        .collect()
}

/// Writes ExcelJS's workbook model (`workbook.model`) as `.xlsx`.
pub fn write(model: Value, options: &WriteOptions) -> Result<Vec<u8>> {
    let mut model = model;
    let mut styles = Vec::new();
    let mut rows = Vec::new();
    for ws in model
        .get_mut("worksheets")
        .and_then(Value::as_array_mut)
        .into_iter()
        .flatten()
    {
        rows.push(split_rows(ws, &mut styles)?);
    }
    write_rows(model, &crate::rows::encode_sheets(&rows), styles, options)
}

/// Writes a workbook model whose worksheets' rows come separately:
/// `rows` holds each worksheet's row models in the binary form of
/// `lib/xlsx/rows.js`, with styles as indexes into `styles_json`.
pub fn write_xlsx_parts(
    model: &str,
    rows: &[u8],
    styles_json: &str,
    options: &WriteOptions,
) -> Result<Vec<u8>> {
    let bad = |e: serde_json::Error| Error::invalid("workbook", e.to_string());
    let model: Value = serde_json::from_str(model).map_err(bad)?;
    let styles: Vec<Value> = serde_json::from_str(styles_json).map_err(bad)?;
    write_rows(model, rows, styles, options)
}

/// Takes `ws.rows` out of a worksheet model as typed rows, moving each
/// style object into `styles`.
fn split_rows(ws: &mut Value, styles: &mut Vec<Value>) -> Result<Vec<RowModel>> {
    let mut rows = match ws.get_mut("rows").map(Value::take) {
        Some(Value::Array(a)) => a,
        _ => return Ok(Vec::new()),
    };
    let mut move_style = |o: &mut Value| {
        if let Some(m) = o.as_object_mut() {
            if let Some(st) = m.remove("style") {
                m.insert("$s".into(), json!(styles.len()));
                styles.push(st);
            }
        }
    };
    for row in &mut rows {
        move_style(row);
        for cell in row
            .get_mut("cells")
            .and_then(Value::as_array_mut)
            .into_iter()
            .flatten()
        {
            move_style(cell);
        }
    }
    serde_json::from_value(Value::Array(rows))
        .map_err(|e| Error::invalid("workbook", e.to_string()))
}

fn write_rows(
    model: Value,
    rows: &[u8],
    style_table: Vec<Value>,
    options: &WriteOptions,
) -> Result<Vec<u8>> {
    let mut table = StyleTable::new(style_table);
    let mut model = model;
    let fail = |m: String| Error::invalid("workbook", m);
    {
        let m = model
            .as_object_mut()
            .ok_or_else(|| fail("model is not an object".into()))?;
        for (k, d) in [("creator", "ExcelJS"), ("lastModifiedBy", "ExcelJS")] {
            if !truthy(m.get(k)) {
                m.insert(k.into(), json!(d));
            }
        }
        for k in ["created", "modified"] {
            if !truthy(m.get(k)) {
                m.insert(k.into(), crate::model_json::date(options.now_ms));
            }
        }
    }
    let use_styles = options.use_styles.unwrap_or(true);
    let mut styles = StyleManager::new(use_styles);
    let mut shared_strings = SharedStrings::default();
    // ExcelJS sets `model.useSharedStrings` and then ignores it (B3).
    let use_shared_strings = options.use_shared_strings.unwrap_or(true);
    prepare_workbook(&mut model).map_err(fail)?;
    let date1904 = match get(Some(&model), "properties") {
        None => {
            return Err(fail(
                "Cannot read properties of undefined (reading 'date1904')".into(),
            ))
        }
        Some(Value::Null) => {
            return Err(fail(
                "Cannot read properties of null (reading 'date1904')".into(),
            ))
        }
        p => truthy(get(p, "date1904")),
    };
    let media: Vec<Value> = get(Some(&model), "media")
        .and_then(Value::as_array)
        .cloned()
        .unwrap_or_default();
    let mut drawings: Vec<Value> = Vec::new();
    let mut comment_refs: Vec<Value> = Vec::new();
    let mut drawings_count = 0usize;
    let mut pre_image_id: Option<Value> = None;
    let mut tables: Vec<Value> = Vec::new();
    let mut table_count = 0usize;
    let mut worksheets: Vec<Value> = match model.get_mut("worksheets").map(Value::take) {
        Some(Value::Array(a)) => a,
        _ => Vec::new(),
    };
    // Each worksheet's rows are decoded one at a time, prepared and encoded
    // again; rendering decodes them one at a time too.
    let mut input = RowCursor::sheets(rows).map_err(|e| Error::invalid("workbook", e))?;
    let mut prepared_rows: Vec<PreparedRows> = Vec::with_capacity(worksheets.len());
    for ws in &mut worksheets {
        let has_rows = input
            .next_sheet()
            .map_err(|e| Error::invalid("workbook", e))?;
        let mut none = RowCursor::rows(&[], 0);
        for table in ws
            .get_mut("tables")
            .and_then(Value::as_array_mut)
            .into_iter()
            .flatten()
        {
            table_count += 1;
            crate::jsv::set(table, "target", json!(format!("table{table_count}.xml")));
            crate::jsv::set(table, "id", json!(table_count));
        }
        let mut ctx = SheetPrepare {
            styles: &mut styles,
            table: &mut table,
            shared_strings: use_shared_strings.then_some(&mut shared_strings),
            date1904,
            media: &media,
            drawings_count: &mut drawings_count,
            drawings: &mut drawings,
            comment_refs: &mut comment_refs,
            pre_image_id: &mut pre_image_id,
        };
        let sheet_rows = if has_rows { &mut input } else { &mut none };
        let prepared = prepare_worksheet(ws, sheet_rows, &mut ctx).map_err(|e| {
            Error::invalid(
                &format!("worksheet {}", to_js_string(get(Some(ws), "name"))),
                e,
            )
        })?;
        prepared_rows.push(prepared);
        tables.extend(
            ws.get("tables")
                .and_then(Value::as_array)
                .into_iter()
                .flatten()
                .cloned(),
        );
    }

    // Parts are compressed in order, on another thread, while the next
    // ones are rendered.
    let mut zip = Zipper::new(options);

    // [Content_Types].xml
    let sheet_summaries = |ws: &[Value]| Value::Array(ws.iter().map(sheet_summary).collect());
    let ct_model = json!({
        "media": media,
        "worksheets": sheet_summaries(&worksheets),
        "sharedStrings": {"count": shared_strings.count},
        "tables": tables,
        "drawings": drawings,
        "commentRefs": comment_refs,
    });
    zip.part(
        "[Content_Types].xml".into(),
        render_content_types(&ct_model).into_bytes(),
    )?;
    zip.part(
        "_rels/.rels".into(),
        render_relationships(&rels_json(&[
            ("rId1", REL_OFFICE_DOCUMENT, "xl/workbook.xml"),
            ("rId2", REL_CORE, "docProps/core.xml"),
            ("rId3", REL_EXTENDED, "docProps/app.xml"),
        ]))
        .into_bytes(),
    )?;
    // xl/_rels/workbook.xml.rels
    let mut count = 1;
    let mut wb_rels = Vec::new();
    wb_rels.push(json!({"Id": format!("rId{count}"), "Type": REL_STYLES, "Target": "styles.xml"}));
    count += 1;
    wb_rels.push(
        json!({"Id": format!("rId{count}"), "Type": REL_THEME, "Target": "theme/theme1.xml"}),
    );
    count += 1;
    if shared_strings.count > 0 {
        wb_rels.push(json!({"Id": format!("rId{count}"), "Type": REL_SHARED_STRINGS, "Target": "sharedStrings.xml"}));
        count += 1;
    }
    for ws in &mut worksheets {
        let rid = format!("rId{count}");
        count += 1;
        crate::jsv::set(ws, "rId", json!(rid));
        wb_rels.push(json!({"Id": rid, "Type": REL_WORKSHEET, "Target": format!("worksheets/sheet{}.xml", to_js_string(ws.get("id")))}));
    }
    zip.part(
        "xl/_rels/workbook.xml.rels".into(),
        render_relationships(&wb_rels).into_bytes(),
    )?;
    // worksheets
    for (ws, ws_rows) in worksheets.iter().zip(&prepared_rows) {
        let id = to_js_string(ws.get("id"));
        let name = format!("xl/worksheets/sheet{id}.xml");
        zip.start(name)?;
        let rest = render_worksheet(ws, ws_rows, &mut |chunk| zip.data(chunk.into_bytes()))
            .map_err(|e| match e {
                WriteFailure::Render(m) => fail(m),
                WriteFailure::Zip(e) => e,
            })?;
        zip.data(rest.into_bytes())?;
        if let Some(rels) = ws
            .get("rels")
            .and_then(Value::as_array)
            .filter(|r| !r.is_empty())
        {
            zip.part(
                format!("xl/worksheets/_rels/sheet{id}.xml.rels"),
                render_relationships(rels).into_bytes(),
            )?;
        }
        if ws
            .get("comments")
            .and_then(Value::as_array)
            .is_some_and(|c| !c.is_empty())
        {
            zip.part(
                format!("xl/comments{id}.xml"),
                render_comments(ws).into_bytes(),
            )?;
            zip.part(
                format!("xl/drawings/vmlDrawing{id}.vml"),
                render_vml_notes(ws).into_bytes(),
            )?;
        }
    }
    if shared_strings.count > 0 {
        zip.part(
            "xl/sharedStrings.xml".into(),
            render_shared_strings(&shared_strings.values, shared_strings.count).into_bytes(),
        )?;
    }
    // drawings
    for ws in &worksheets {
        if let Some(drawing) = ws.get("drawing").filter(|d| truthy(Some(d))) {
            let mut d = drawing.clone();
            for (index, anchor) in d
                .get_mut("anchors")
                .and_then(Value::as_array_mut)
                .into_iter()
                .flatten()
                .enumerate()
            {
                crate::jsv::set(
                    anchor,
                    "anchorType",
                    json!(anchor_type(anchor).map_err(fail)?),
                );
                if let Some(Value::Object(p)) = anchor.get_mut("picture") {
                    p.insert("index".into(), json!(index + 1));
                }
            }
            let name = to_js_string(d.get("name"));
            zip.part(
                format!("xl/drawings/{name}.xml"),
                to_xml(&drawing_xform(), Some(&d)).into_bytes(),
            )?;
            let rels: Vec<Value> = d
                .get("rels")
                .and_then(Value::as_array)
                .cloned()
                .unwrap_or_default();
            zip.part(
                format!("xl/drawings/_rels/{name}.xml.rels"),
                render_relationships(&rels).into_bytes(),
            )?;
        }
    }
    // tables
    for ws in &worksheets {
        for table in ws
            .get("tables")
            .and_then(Value::as_array)
            .into_iter()
            .flatten()
        {
            let mut t = table.clone();
            prepare_table(&mut t);
            let target = to_js_string(t.get("target"));
            zip.part(
                format!("xl/tables/{target}"),
                to_xml(&TableXform::new(), Some(&t)).into_bytes(),
            )?;
        }
    }
    // themes
    match get(Some(&model), "themes").filter(|t| truthy(Some(t))) {
        Some(Value::Object(themes)) => {
            for (name, xml) in themes {
                zip.part(
                    format!("xl/theme/{name}.xml"),
                    to_js_string(Some(xml)).into_bytes(),
                )?;
            }
        }
        Some(_) => {}
        None => zip.part("xl/theme/theme1.xml".into(), THEME1.as_bytes().to_vec())?,
    }
    zip.part("xl/styles.xml".into(), styles.render().into_bytes())?;
    // media
    for medium in &media {
        if as_str(get(Some(medium), "type")) != Some("image") {
            return Err(fail("Unsupported media".into()));
        }
        let filename = format!(
            "xl/media/{}.{}",
            to_js_string(get(Some(medium), "name")),
            to_js_string(get(Some(medium), "extension"))
        );
        let data = if truthy(get(Some(medium), "filename")) {
            std::fs::read(to_js_string(get(Some(medium), "filename")))?
        } else if let Some(b) = as_buffer(get(Some(medium), "buffer")) {
            b
        } else if truthy(get(Some(medium), "base64")) {
            let s = to_js_string(get(Some(medium), "base64"));
            let content = s
                .find(',')
                .map_or(s.as_str(), |i| s.get(i + 1..).unwrap_or(""));
            base64_decode(content).unwrap_or_default()
        } else {
            return Err(fail("Unsupported media".into()));
        };
        zip.part(filename, data)?;
    }
    // docProps
    let summaries = sheet_summaries(&worksheets);
    let mut app_model = Map::new();
    app_model.insert("worksheets".into(), summaries.clone());
    for k in ["company", "manager"] {
        if let Some(v) = get(Some(&model), k) {
            app_model.insert(k.into(), v.clone());
        }
    }
    let app_model = Value::Object(app_model);
    zip.part(
        "docProps/app.xml".into(),
        render_app(&app_model).into_bytes(),
    )?;
    zip.part(
        "docProps/core.xml".into(),
        render_core(&model).map_err(fail)?.into_bytes(),
    )?;
    // workbook
    crate::jsv::set(&mut model, "sheets", summaries);
    zip.part(
        "xl/workbook.xml".into(),
        to_xml(&WorkbookXform::new(), Some(&model)).into_bytes(),
    )?;

    zip.finish()
}

/// The archive being written: parts are compressed on another thread
/// (`zip_thread`), in order. A later part with a name already written is
/// left out.
struct Zipper {
    thread: ZipThread<ZipArchiveOut<std::io::Cursor<Vec<u8>>>>,
    seen: std::collections::HashSet<String>,
    skip: bool,
}

impl Zipper {
    fn new(options: &WriteOptions) -> Self {
        Zipper {
            thread: ZipThread::new(ZipArchiveOut {
                zip: ZipWriter::new(std::io::Cursor::new(Vec::new())),
                options: SimpleFileOptions::default()
                    .compression_method(zip::CompressionMethod::Deflated)
                    .compression_level(options.compression_level),
            }),
            seen: std::collections::HashSet::new(),
            skip: false,
        }
    }

    fn start(&mut self, name: String) -> Result<()> {
        self.skip = !self.seen.insert(name.clone());
        if self.skip {
            return Ok(());
        }
        self.thread.start(name)
    }

    fn data(&mut self, data: Vec<u8>) -> Result<()> {
        if self.skip {
            return Ok(());
        }
        self.thread.data(data)
    }

    fn part(&mut self, name: String, data: Vec<u8>) -> Result<()> {
        self.start(name)?;
        self.data(data)
    }

    fn finish(self) -> Result<Vec<u8>> {
        Ok(self.thread.finish()?.into_inner())
    }
}
