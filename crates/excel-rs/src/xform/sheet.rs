//! ExcelJS `lib/xlsx/xform/sheet/*.js` except the cell/row/worksheet
//! xforms (see `cell.rs` and `worksheet.rs`) and conditional formatting
//! (`cf.rs`).

use serde_json::{json, Map, Value};

use super::base::ListXform;
use super::composite::{object, Leaf, LeafMatch};
use super::simple::{float_value, int_value};
use super::style::{parse_boolean, ColorXform};
use super::{XResult, Xform};
use crate::attrs;
use crate::col_cache::{decode_ex, DecodedEx};
use crate::jsv::{as_str, get, to_js_string, truthy};
use crate::xml::Tag;
use crate::xml_stream::{Attrs, XmlStream};

fn s(v: Option<&str>) -> Option<Value> {
    v.map(|x| json!(x))
}

/// `{ref: model}` leaf rendering for string models.
fn ref_leaf(tag: &'static str) -> fn(&mut XmlStream, Option<&Value>) {
    match tag {
        "mergeCell" => |xml, m| xml.leaf_node("mergeCell", &attrs![("ref", m)], None),
        _ => |xml, m| {
            if truthy(m) {
                xml.leaf_node("dimension", &attrs![("ref", m)], None);
            }
        },
    }
}

pub(crate) fn merge_cell() -> Leaf {
    Leaf::new(
        "mergeCell",
        LeafMatch::Tag,
        |n| s(n.attr("ref")),
        ref_leaf("mergeCell"),
    )
}

pub(crate) fn dimension() -> Leaf {
    Leaf::new(
        "dimension",
        LeafMatch::Tag,
        |n| s(n.attr("ref")),
        ref_leaf("dimension"),
    )
}

fn obj_attrs(m: Option<&Value>) -> Attrs {
    match m {
        Some(Value::Object(o)) => o
            .iter()
            .map(|(k, v)| (k.clone(), Some(to_js_string(Some(v)))))
            .collect(),
        _ => Vec::new(),
    }
}

/// `PageBreaksXform` (`<brk>`): parses `ref`, renders the model's keys.
pub(crate) fn page_break() -> Leaf {
    Leaf::new(
        "brk",
        LeafMatch::Tag,
        |n| s(n.attr("ref")),
        |xml, m| xml.leaf_node("brk", &obj_attrs(m), None),
    )
}

/// `RowBreaksXform`: a ListXform with its own render.
#[derive(Debug)]
pub(crate) struct RowBreaksXform {
    list: ListXform,
}

impl RowBreaksXform {
    pub fn new() -> Self {
        RowBreaksXform {
            list: ListXform::new("rowBreaks", Box::new(page_break())).count(),
        }
    }
}

impl Xform for RowBreaksXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        self.list.parse_open(node)
    }
    fn parse_text(&mut self, text: &str) {
        self.list.parse_text(text);
    }
    fn parse_close(&mut self, name: &str) -> XResult<bool> {
        self.list.parse_close(name)
    }
    fn model(&self) -> Option<&Value> {
        self.list.model()
    }
    fn take_model(&mut self) -> Option<Value> {
        self.list.take_model()
    }
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        let items = model.and_then(Value::as_array).filter(|a| !a.is_empty());
        if let Some(items) = items {
            xml.open_node("rowBreaks", &[]);
            xml.add_attribute("count", Some(items.len()));
            xml.add_attribute("manualBreakCount", Some(items.len()));
            for i in items {
                page_break().render(xml, Some(i));
            }
            xml.close_node();
        }
    }
    fn reset(&mut self) {
        self.list.reset();
    }
}

fn rid_leaf(tag: &'static str) -> Leaf {
    let render: fn(&mut XmlStream, Option<&Value>) = match tag {
        "drawing" => |xml, m| {
            if truthy(m) {
                xml.leaf_node("drawing", &attrs![("r:id", get(m, "rId"))], None);
            }
        },
        "picture" => |xml, m| {
            if truthy(m) {
                xml.leaf_node("picture", &attrs![("r:id", get(m, "rId"))], None);
            }
        },
        _ => |xml, m| {
            if truthy(m) {
                xml.leaf_node("tablePart", &attrs![("r:id", get(m, "rId"))], None);
            }
        },
    };
    Leaf::new(
        tag,
        LeafMatch::Tag,
        |n| Some(object(vec![("rId", s(n.attr("r:id")))])),
        render,
    )
}

pub(crate) fn drawing() -> Leaf {
    rid_leaf("drawing")
}
pub(crate) fn picture() -> Leaf {
    rid_leaf("picture")
}
pub(crate) fn table_part() -> Leaf {
    rid_leaf("tablePart")
}

/// `AutoFilterXform`.
pub(crate) fn auto_filter() -> Leaf {
    Leaf::new(
        "autoFilter",
        LeafMatch::TagNoReturn,
        |n| s(n.attr("ref")),
        |xml, m| {
            if !truthy(m) {
                return;
            }
            if let Some(Value::String(r)) = m {
                xml.leaf_node("autoFilter", &attrs![("ref", Some(r))], None);
                return;
            }
            let address = |a: Option<&Value>| -> Result<String, String> {
                match a {
                    Some(Value::String(s)) => Ok(s.clone()),
                    _ => {
                        let col = crate::jsv::to_number(get(a, "column"));
                        let row = get(a, "row");
                        let letters = crate::col_cache::n2l_str(col)?;
                        Ok(crate::col_cache::decode_address(&format!(
                            "{letters}{}",
                            to_js_string(row)
                        ))?
                        .address)
                    }
                }
            };
            match (address(get(m, "from")), address(get(m, "to"))) {
                (Ok(a), Ok(b)) => {
                    if !a.is_empty() && !b.is_empty() {
                        xml.leaf_node(
                            "autoFilter",
                            &attrs![("ref", Some(format!("{a}:{b}")))],
                            None,
                        );
                    }
                }
                (Err(e), _) | (_, Err(e)) => xml.fail(e),
            }
        },
    )
}

fn page_setup_properties() -> Leaf {
    Leaf::new(
        "pageSetUpPr",
        LeafMatch::Tag,
        |n| Some(json!({"fitToPage": n.attr("fitToPage") == Some("1")})),
        |xml, m| {
            if truthy(m) && truthy(get(m, "fitToPage")) {
                xml.leaf_node("pageSetUpPr", &attrs![("fitToPage", Some("1"))], None);
            }
        },
    )
}

fn outline_properties() -> Leaf {
    Leaf::new(
        "outlinePr",
        LeafMatch::Tag,
        |n| {
            let b = |a: &str| {
                n.attr(a).map(|v| {
                    json!(
                        crate::jsv::string_to_number(v) != 0.0
                            && !crate::jsv::string_to_number(v).is_nan()
                    )
                })
            };
            Some(object(vec![
                ("summaryBelow", b("summaryBelow")),
                ("summaryRight", b("summaryRight")),
            ]))
        },
        |xml, m| {
            render_outline_properties(xml, m);
        },
    )
}

fn render_outline_properties(xml: &mut XmlStream, m: Option<&Value>) -> bool {
    let below = get(m, "summaryBelow");
    let right = get(m, "summaryRight");
    if truthy(m) && (below.is_some() || right.is_some()) {
        let n = |v: Option<&Value>| {
            v.map(|v| crate::jsv::number_to_string(crate::jsv::to_number(Some(v))))
        };
        xml.leaf_node(
            "outlinePr",
            &attrs![("summaryBelow", n(below)), ("summaryRight", n(right))],
            None,
        );
        return true;
    }
    false
}

/// `PageMarginsXform`.
pub(crate) fn page_margins() -> Leaf {
    Leaf::new(
        "pageMargins",
        LeafMatch::Tag,
        |n| {
            let f = |a: &str, d: &str| {
                float_value(Some(n.attr(a).filter(|v| !v.is_empty()).unwrap_or(d)))
            };
            Some(json!({
                "left": f("left", "0.7"),
                "right": f("right", "0.7"),
                "top": f("top", "0.75"),
                "bottom": f("bottom", "0.75"),
                "header": f("header", "0.3"),
                "footer": f("footer", "0.3"),
            }))
        },
        |xml, m| {
            if !truthy(m) {
                return;
            }
            let a = attrs![
                ("left", get(m, "left")),
                ("right", get(m, "right")),
                ("top", get(m, "top")),
                ("bottom", get(m, "bottom")),
                ("header", get(m, "header")),
                ("footer", get(m, "footer")),
            ];
            if a.iter().any(|(_, v)| v.is_some()) {
                xml.leaf_node("pageMargins", &a, None);
            }
        },
    )
}

fn bool_xml(v: Option<&Value>) -> Option<&'static str> {
    truthy(v).then_some("1")
}

/// `PrintOptionsXform`.
pub(crate) fn print_options() -> Leaf {
    Leaf::new(
        "printOptions",
        LeafMatch::Tag,
        |n| {
            Some(json!({
                "showRowColHeaders": n.attr("headings") == Some("1"),
                "showGridLines": n.attr("gridLines") == Some("1"),
                "horizontalCentered": n.attr("horizontalCentered") == Some("1"),
                "verticalCentered": n.attr("verticalCentered") == Some("1"),
            }))
        },
        |xml, m| {
            if !truthy(m) {
                return;
            }
            let a = attrs![
                ("headings", bool_xml(get(m, "showRowColHeaders"))),
                ("gridLines", bool_xml(get(m, "showGridLines"))),
                ("horizontalCentered", bool_xml(get(m, "horizontalCentered"))),
                ("verticalCentered", bool_xml(get(m, "verticalCentered"))),
            ];
            if a.iter().any(|(_, v)| v.is_some()) {
                xml.leaf_node("printOptions", &a, None);
            }
        },
    )
}

/// `/^[^!]+![a-zA-Z]+[\d]+$/`.
fn is_internal_link(target: &str) -> bool {
    let Some((sheet, cell)) = target.split_once('!') else {
        return false;
    };
    let letters = cell.bytes().take_while(u8::is_ascii_alphabetic).count();
    let digits = cell.get(letters..).unwrap_or("");
    !sheet.is_empty()
        && letters > 0
        && !digits.is_empty()
        && digits.bytes().all(|b| b.is_ascii_digit())
}

/// `HyperlinkXform`.
pub(crate) fn hyperlink() -> Leaf {
    Leaf::new(
        "hyperlink",
        LeafMatch::Tag,
        |n| {
            Some(object(vec![
                ("address", s(n.attr("ref"))),
                ("rId", s(n.attr("r:id"))),
                ("tooltip", s(n.attr("tooltip"))),
                ("target", s(n.attr("location").filter(|l| !l.is_empty()))),
            ]))
        },
        |xml, m| {
            let target = get(m, "target");
            let internal = truthy(target) && is_internal_link(&to_js_string(target));
            let mut a = attrs![
                ("ref", get(m, "address")),
                ("r:id", get(m, "rId")),
                ("tooltip", get(m, "tooltip"))
            ];
            if internal {
                a.push(("location".into(), target.map(|t| to_js_string(Some(t)))));
            }
            xml.leaf_node("hyperlink", &a, None);
        },
    )
}

/// `SheetFormatPropertiesXform`.
pub(crate) fn sheet_format_properties() -> Leaf {
    Leaf::new(
        "sheetFormatPr",
        LeafMatch::Tag,
        |n| {
            let or = |a: &str, d: &'static str| n.attr(a).filter(|v| !v.is_empty()).unwrap_or(d);
            let mut m = Map::new();
            m.insert(
                "defaultRowHeight".into(),
                float_value(Some(or("defaultRowHeight", "0"))),
            );
            m.insert(
                "dyDescent".into(),
                float_value(Some(or("x14ac:dyDescent", "0"))),
            );
            m.insert(
                "outlineLevelRow".into(),
                int_value(Some(or("outlineLevelRow", "0"))),
            );
            m.insert(
                "outlineLevelCol".into(),
                int_value(Some(or("outlineLevelCol", "0"))),
            );
            if let Some(w) = n.attr("defaultColWidth").filter(|v| !v.is_empty()) {
                m.insert("defaultColWidth".into(), float_value(Some(w)));
            }
            Some(Value::Object(m))
        },
        |xml, m| {
            if !truthy(m) {
                return;
            }
            let mut a = attrs![
                ("defaultRowHeight", get(m, "defaultRowHeight")),
                ("outlineLevelRow", get(m, "outlineLevelRow")),
                ("outlineLevelCol", get(m, "outlineLevelCol")),
                ("x14ac:dyDescent", get(m, "dyDescent")),
            ];
            if truthy(get(m, "defaultColWidth")) {
                a.push((
                    "defaultColWidth".into(),
                    get(m, "defaultColWidth").map(|v| to_js_string(Some(v))),
                ));
            }
            let drh = get(m, "defaultRowHeight");
            if !truthy(drh) || crate::jsv::as_f64(drh) != Some(15.0) {
                a.push(("customHeight".into(), Some("1".into())));
            }
            if a.iter().any(|(_, v)| v.is_some()) {
                xml.leaf_node("sheetFormatPr", &a, None);
            }
        },
    )
}

/// `ColXform` (parse and render; prepare/reconcile live in the worksheet).
pub(crate) fn col() -> Leaf {
    Leaf::new(
        "col",
        LeafMatch::Tag,
        |n| {
            let or0 = |a: &str| n.attr(a).filter(|v| !v.is_empty()).unwrap_or("0");
            let mut m = Map::new();
            m.insert("min".into(), int_value(Some(or0("min"))));
            m.insert("max".into(), int_value(Some(or0("max"))));
            if n.attr("width").is_some() {
                m.insert("width".into(), float_value(Some(or0("width"))));
            }
            if let Some(st) = n.attr("style").filter(|v| !v.is_empty()) {
                m.insert("styleId".into(), int_value(Some(st)));
            }
            if parse_boolean(n.attr("hidden")) {
                m.insert("hidden".into(), json!(true));
            }
            if parse_boolean(n.attr("bestFit")) {
                m.insert("bestFit".into(), json!(true));
            }
            if let Some(o) = n.attr("outlineLevel").filter(|v| !v.is_empty()) {
                m.insert("outlineLevel".into(), int_value(Some(o)));
            }
            if parse_boolean(n.attr("collapsed")) {
                m.insert("collapsed".into(), json!(true));
            }
            Some(Value::Object(m))
        },
        |xml, m| {
            xml.open_node("col", &[]);
            xml.add_attribute("min", get(m, "min"));
            xml.add_attribute("max", get(m, "max"));
            if truthy(get(m, "width")) {
                xml.add_attribute("width", get(m, "width"));
            }
            if truthy(get(m, "styleId")) {
                xml.add_attribute("style", get(m, "styleId"));
            }
            if truthy(get(m, "hidden")) {
                xml.add_attribute("hidden", Some("1"));
            }
            if truthy(get(m, "bestFit")) {
                xml.add_attribute("bestFit", Some("1"));
            }
            if truthy(get(m, "outlineLevel")) {
                xml.add_attribute("outlineLevel", get(m, "outlineLevel"));
            }
            if truthy(get(m, "collapsed")) {
                xml.add_attribute("collapsed", Some("1"));
            }
            xml.add_attribute("customWidth", Some("1"));
            xml.close_node();
        },
    )
}

/// `SheetProtectionXform`.
pub(crate) fn sheet_protection() -> Leaf {
    Leaf::new(
        "sheetProtection",
        LeafMatch::Tag,
        |n| {
            let to_bool = |a: &str, eq: &str| (n.attr(a) == Some(eq)).then_some(json!(true));
            let to_false = |a: &str| (n.attr(a) == Some("1")).then_some(json!(false));
            let mut pairs = vec![
                ("sheet", to_bool("sheet", "1")),
                ("objects", to_false("objects")),
                ("scenarios", to_false("scenarios")),
                ("selectLockedCells", to_false("selectLockedCells")),
                ("selectUnlockedCells", to_false("selectUnlockedCells")),
            ];
            for k in [
                "formatCells",
                "formatColumns",
                "formatRows",
                "insertColumns",
                "insertRows",
                "insertHyperlinks",
                "deleteColumns",
                "deleteRows",
                "sort",
                "autoFilter",
                "pivotTables",
            ] {
                pairs.push((k, to_bool(k, "0")));
            }
            let mut m = match object(pairs) {
                Value::Object(m) => m,
                _ => Map::new(),
            };
            if let Some(alg) = n.attr("algorithmName").filter(|v| !v.is_empty()) {
                m.insert("algorithmName".into(), json!(alg));
                if let Some(h) = n.attr("hashValue") {
                    m.insert("hashValue".into(), json!(h));
                }
                if let Some(sv) = n.attr("saltValue") {
                    m.insert("saltValue".into(), json!(sv));
                }
                m.insert("spinCount".into(), int_value(n.attr("spinCount")));
            }
            Some(Value::Object(m))
        },
        |xml, m| {
            if !truthy(m) {
                return;
            }
            let b = |k: &str, v: &'static str| truthy(get(m, k)).then_some(v);
            let is_false = |k: &str| get(m, k) == Some(&Value::Bool(false));
            let mut a = attrs![
                ("sheet", b("sheet", "1")),
                (
                    "selectLockedCells",
                    is_false("selectLockedCells").then_some("1")
                ),
                (
                    "selectUnlockedCells",
                    is_false("selectUnlockedCells").then_some("1")
                ),
            ];
            for k in [
                "formatCells",
                "formatColumns",
                "formatRows",
                "insertColumns",
                "insertRows",
                "insertHyperlinks",
                "deleteColumns",
                "deleteRows",
                "sort",
                "autoFilter",
                "pivotTables",
            ] {
                a.push((k.into(), b(k, "0").map(str::to_owned)));
            }
            if truthy(get(m, "sheet")) {
                for k in ["algorithmName", "hashValue", "saltValue", "spinCount"] {
                    a.push((k.into(), get(m, k).map(|v| to_js_string(Some(v)))));
                }
                a.push((
                    "objects".into(),
                    is_false("objects").then(|| "1".to_owned()),
                ));
                a.push((
                    "scenarios".into(),
                    is_false("scenarios").then(|| "1".to_owned()),
                ));
            }
            if a.iter().any(|(_, v)| v.is_some()) {
                xml.leaf_node("sheetProtection", &a, None);
            }
        },
    )
}

/// `PageSetupXform`.
pub(crate) fn page_setup() -> Leaf {
    Leaf::new(
        "pageSetup",
        LeafMatch::Tag,
        |n| {
            let or = |a: &str, d: &'static str| n.attr(a).filter(|v| !v.is_empty()).unwrap_or(d);
            Some(object(vec![
                ("paperSize", n.attr("paperSize").map(|v| int_value(Some(v)))),
                ("orientation", Some(json!(or("orientation", "portrait")))),
                (
                    "horizontalDpi",
                    Some(int_value(Some(or("horizontalDpi", "4294967295")))),
                ),
                (
                    "verticalDpi",
                    Some(int_value(Some(or("verticalDpi", "4294967295")))),
                ),
                ("pageOrder", Some(json!(or("pageOrder", "downThenOver")))),
                (
                    "blackAndWhite",
                    Some(json!(n.attr("blackAndWhite") == Some("1"))),
                ),
                ("draft", Some(json!(n.attr("draft") == Some("1")))),
                ("cellComments", Some(json!(or("cellComments", "None")))),
                ("errors", Some(json!(or("errors", "displayed")))),
                ("scale", Some(int_value(Some(or("scale", "100"))))),
                ("fitToWidth", Some(int_value(Some(or("fitToWidth", "1"))))),
                ("fitToHeight", Some(int_value(Some(or("fitToHeight", "1"))))),
                (
                    "firstPageNumber",
                    Some(int_value(Some(or("firstPageNumber", "1")))),
                ),
                (
                    "useFirstPageNumber",
                    Some(json!(n.attr("useFirstPageNumber") == Some("1"))),
                ),
                (
                    "usePrinterDefaults",
                    Some(json!(n.attr("usePrinterDefaults") == Some("1"))),
                ),
                ("copies", Some(int_value(Some(or("copies", "1"))))),
            ]))
        },
        |xml, m| {
            if !truthy(m) {
                return;
            }
            let v = |k: &str| get(m, k).map(|x| to_js_string(Some(x)));
            let only = |k: &str, allowed: &[&str]| {
                as_str(get(m, k))
                    .filter(|x| allowed.contains(x))
                    .map(str::to_owned)
            };
            let a = attrs![
                ("paperSize", v("paperSize")),
                ("orientation", v("orientation")),
                ("horizontalDpi", v("horizontalDpi")),
                ("verticalDpi", v("verticalDpi")),
                ("pageOrder", only("pageOrder", &["overThenDown"])),
                ("blackAndWhite", bool_xml(get(m, "blackAndWhite"))),
                ("draft", bool_xml(get(m, "draft"))),
                (
                    "cellComments",
                    only("cellComments", &["atEnd", "asDisplyed"])
                ),
                ("errors", only("errors", &["dash", "blank", "NA"])),
                ("scale", v("scale")),
                ("fitToWidth", v("fitToWidth")),
                ("fitToHeight", v("fitToHeight")),
                ("firstPageNumber", v("firstPageNumber")),
                ("useFirstPageNumber", bool_xml(get(m, "firstPageNumber"))),
                ("usePrinterDefaults", bool_xml(get(m, "usePrinterDefaults"))),
                ("copies", v("copies")),
            ];
            if a.iter().any(|(_, v)| v.is_some()) {
                xml.leaf_node("pageSetup", &a, None);
            }
        },
    )
}

/// `SheetPropertiesXform` (`<sheetPr>`).
#[derive(Debug)]
pub(crate) struct SheetPropertiesXform {
    tab_color: ColorXform,
    page_setup: Leaf,
    outline: Leaf,
    active: Option<u8>,
    model: Option<Value>,
}

impl SheetPropertiesXform {
    pub fn new() -> Self {
        SheetPropertiesXform {
            tab_color: ColorXform::new("tabColor"),
            page_setup: page_setup_properties(),
            outline: outline_properties(),
            active: None,
            model: None,
        }
    }

    fn child(&mut self, i: u8) -> &mut dyn Xform {
        match i {
            0 => &mut self.tab_color,
            1 => &mut self.page_setup,
            _ => &mut self.outline,
        }
    }
}

impl Xform for SheetPropertiesXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if let Some(i) = self.active {
            self.child(i).parse_open(node)?;
            return Ok(true);
        }
        if node.name == "sheetPr" {
            self.reset();
            return Ok(true);
        }
        self.active = match node.name {
            "tabColor" => Some(0),
            "pageSetUpPr" => Some(1),
            "outlinePr" => Some(2),
            _ => None,
        };
        match self.active {
            Some(i) => {
                self.child(i).parse_open(node)?;
                Ok(true)
            }
            None => Ok(false),
        }
    }
    fn parse_text(&mut self, text: &str) {
        if let Some(i) = self.active {
            self.child(i).parse_text(text);
        }
    }
    fn parse_close(&mut self, name: &str) -> XResult<bool> {
        if let Some(i) = self.active {
            if !self.child(i).parse_close(name)? {
                self.active = None;
            }
            return Ok(true);
        }
        let tab = self.tab_color.model().filter(|v| truthy(Some(v))).cloned();
        let ps = self.page_setup.model().filter(|v| truthy(Some(v))).cloned();
        let ol = self.outline.model().filter(|v| truthy(Some(v))).cloned();
        self.model = Some(if tab.is_some() || ps.is_some() || ol.is_some() {
            object(vec![
                ("tabColor", tab),
                ("pageSetup", ps),
                ("outlineProperties", ol),
            ])
        } else {
            Value::Null
        });
        Ok(false)
    }
    crate::model_accessors!();
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        if !truthy(model) {
            return;
        }
        xml.add_rollback();
        xml.open_node("sheetPr", &[]);
        let mut inner = false;
        let tab = get(model, "tabColor");
        if truthy(tab) {
            self.tab_color.render(xml, tab);
            inner = true;
        }
        let ps = get(model, "pageSetup");
        if truthy(ps) && truthy(get(ps, "fitToPage")) {
            self.page_setup.render(xml, ps);
            inner = true;
        }
        inner = render_outline_properties(xml, get(model, "outlineProperties")) || inner;
        if inner {
            xml.close_node();
            xml.commit();
        } else {
            xml.rollback();
        }
    }
    fn reset(&mut self) {
        self.model = Some(Value::Null);
        self.tab_color.reset();
        self.page_setup.reset();
        self.outline.reset();
    }
}

const HF_NODES: [&str; 6] = [
    "oddHeader",
    "oddFooter",
    "evenHeader",
    "evenFooter",
    "firstHeader",
    "firstFooter",
];

/// `HeaderFooterXform`.
#[derive(Debug, Default)]
pub(crate) struct HeaderFooterXform {
    current: Option<&'static str>,
    model: Option<Value>,
}

impl Xform for HeaderFooterXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if node.name == "headerFooter" {
            let mut m = Map::new();
            for a in ["differentFirst", "differentOddEven"] {
                if let Some(v) = node.attr(a).filter(|v| !v.is_empty()) {
                    m.insert(a.into(), json!(crate::js::parse_int(v) == Some(1.0)));
                }
            }
            self.model = Some(Value::Object(m));
            return Ok(true);
        }
        if let Some(n) = HF_NODES.iter().find(|n| **n == node.name) {
            self.current = Some(n);
            return Ok(true);
        }
        Ok(false)
    }
    fn parse_text(&mut self, text: &str) {
        if let (Some(cur), Some(Value::Object(m))) = (self.current, self.model.as_mut()) {
            m.insert(cur.into(), json!(text));
        }
    }
    fn parse_close(&mut self, _name: &str) -> XResult<bool> {
        if self.current.is_some() {
            self.current = None;
            return Ok(true);
        }
        Ok(false)
    }
    crate::model_accessors!();
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        if !truthy(model) {
            return;
        }
        xml.add_rollback();
        let mut create = false;
        xml.open_node("headerFooter", &[]);
        for a in ["differentFirst", "differentOddEven"] {
            if truthy(get(model, a)) {
                xml.add_attribute(a, Some("1"));
                create = true;
            }
        }
        for n in HF_NODES {
            if let Some(Value::String(t)) = get(model, n).filter(|v| truthy(Some(v))) {
                xml.leaf_node(n, &[], Some(t));
                create = true;
            }
        }
        if create {
            xml.close_node();
            xml.commit();
        } else {
            xml.rollback();
        }
    }
    fn reset(&mut self) {
        self.model = Some(Value::Null);
    }
}

/// `SheetViewXform`.
#[derive(Debug, Default)]
pub(crate) struct SheetViewXform {
    sheet_view: Option<Map<String, Value>>,
    pane: Option<Map<String, Value>>,
    selections: Map<String, Value>,
    model: Option<Value>,
}

impl Xform for SheetViewXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        match node.name {
            "sheetView" => {
                let or = |a: &str, d: &'static str| n_or(node, a, d);
                let m = object(vec![
                    (
                        "workbookViewId",
                        Some(int_value(node.attr("workbookViewId"))),
                    ),
                    (
                        "rightToLeft",
                        Some(json!(node.attr("rightToLeft") == Some("1"))),
                    ),
                    (
                        "tabSelected",
                        Some(json!(node.attr("tabSelected") == Some("1"))),
                    ),
                    (
                        "showRuler",
                        Some(json!(node.attr("showRuler") != Some("0"))),
                    ),
                    (
                        "showRowColHeaders",
                        Some(json!(node.attr("showRowColHeaders") != Some("0"))),
                    ),
                    (
                        "showGridLines",
                        Some(json!(node.attr("showGridLines") != Some("0"))),
                    ),
                    ("zoomScale", Some(int_value(Some(or("zoomScale", "100"))))),
                    (
                        "zoomScaleNormal",
                        Some(int_value(Some(or("zoomScaleNormal", "100")))),
                    ),
                    ("style", s(node.attr("view"))),
                ]);
                self.sheet_view = m.as_object().cloned();
                self.pane = None;
                self.selections = Map::new();
                Ok(true)
            }
            "pane" => {
                let m = object(vec![
                    ("xSplit", Some(int_value(Some(n_or(node, "xSplit", "0"))))),
                    ("ySplit", Some(int_value(Some(n_or(node, "ySplit", "0"))))),
                    ("topLeftCell", s(node.attr("topLeftCell"))),
                    (
                        "activePane",
                        Some(json!(n_or(node, "activePane", "topLeft"))),
                    ),
                    ("state", s(node.attr("state"))),
                ]);
                self.pane = m.as_object().cloned();
                Ok(true)
            }
            "selection" => {
                let name = n_or(node, "pane", "topLeft").to_owned();
                let sel = object(vec![
                    ("pane", Some(json!(name))),
                    ("activeCell", s(node.attr("activeCell"))),
                ]);
                self.selections.insert(name, sel);
                Ok(true)
            }
            _ => Ok(false),
        }
    }
    fn parse_close(&mut self, name: &str) -> XResult<bool> {
        if name != "sheetView" {
            return Ok(true);
        }
        let sv = self.sheet_view.as_ref().ok_or_else(|| {
            "Cannot read properties of undefined (reading 'workbookViewId')".to_owned()
        })?;
        let g = |k: &str| sv.get(k).cloned();
        let mut m = Map::new();
        let put = |m: &mut Map<String, Value>, k: &str, v: Option<Value>| {
            if let Some(v) = v {
                m.insert(k.into(), v);
            }
        };
        put(&mut m, "workbookViewId", g("workbookViewId"));
        put(&mut m, "rightToLeft", g("rightToLeft"));
        let selection;
        if let Some(pane) = &self.pane {
            let state = match pane.get("state").and_then(Value::as_str) {
                Some("frozen" | "frozenSplit") => "frozen",
                Some("split") => "split",
                _ => "split",
            };
            m.insert("state".into(), json!(state));
            put(&mut m, "xSplit", pane.get("xSplit").cloned());
            put(&mut m, "ySplit", pane.get("ySplit").cloned());
            // ExcelJS keeps the key even when the pane has no topLeftCell.
            m.insert(
                "topLeftCell".into(),
                pane.get("topLeftCell")
                    .cloned()
                    .unwrap_or_else(crate::jsv::undefined),
            );
            for k in [
                "showRuler",
                "showRowColHeaders",
                "showGridLines",
                "zoomScale",
                "zoomScaleNormal",
            ] {
                put(&mut m, k, g(k));
            }
            if state == "split" {
                put(&mut m, "activePane", pane.get("activePane").cloned());
            }
            let active = to_js_string(pane.get("activePane"));
            selection = self.selections.get(&active).cloned();
        } else {
            m.insert("state".into(), json!("normal"));
            for k in [
                "showRuler",
                "showRowColHeaders",
                "showGridLines",
                "zoomScale",
                "zoomScaleNormal",
            ] {
                put(&mut m, k, g(k));
            }
            selection = self.selections.get("topLeft").cloned();
        }
        if let Some(ac) = get(selection.as_ref(), "activeCell").filter(|v| truthy(Some(v))) {
            m.insert("activeCell".into(), ac.clone());
        }
        if let Some(st) = sv.get("style").filter(|v| truthy(Some(v))) {
            m.insert("style".into(), st.clone());
        }
        self.model = Some(Value::Object(m));
        Ok(false)
    }
    crate::model_accessors!();
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        let wbv = get(model, "workbookViewId")
            .filter(|v| truthy(Some(v)))
            .cloned()
            .unwrap_or(json!(0));
        xml.open_node("sheetView", &attrs![("workbookViewId", Some(wbv))]);
        let add = |xml: &mut XmlStream, name: &str, value: Option<String>, included: bool| {
            if included {
                xml.add_attribute(name, value);
            }
        };
        let v = |k: &str| get(model, k).map(|x| to_js_string(Some(x)));
        add(
            xml,
            "rightToLeft",
            Some("1".into()),
            get(model, "rightToLeft") == Some(&Value::Bool(true)),
        );
        add(
            xml,
            "tabSelected",
            Some("1".into()),
            truthy(get(model, "tabSelected")),
        );
        add(
            xml,
            "showRuler",
            Some("0".into()),
            get(model, "showRuler") == Some(&Value::Bool(false)),
        );
        add(
            xml,
            "showRowColHeaders",
            Some("0".into()),
            get(model, "showRowColHeaders") == Some(&Value::Bool(false)),
        );
        add(
            xml,
            "showGridLines",
            Some("0".into()),
            get(model, "showGridLines") == Some(&Value::Bool(false)),
        );
        add(
            xml,
            "zoomScale",
            v("zoomScale"),
            truthy(get(model, "zoomScale")),
        );
        add(
            xml,
            "zoomScaleNormal",
            v("zoomScaleNormal"),
            truthy(get(model, "zoomScaleNormal")),
        );
        add(xml, "view", v("style"), truthy(get(model, "style")));
        let or_undef = |k: &str| if truthy(get(model, k)) { v(k) } else { None };
        match as_str(get(model, "state")) {
            Some("frozen") => {
                let num = |k: &str| {
                    if truthy(get(model, k)) {
                        crate::jsv::to_number(get(model, k))
                    } else {
                        0.0
                    }
                };
                let (xs, ys) = (num("xSplit"), num("ySplit"));
                let top_left = if truthy(get(model, "topLeftCell")) {
                    v("topLeftCell")
                } else {
                    match crate::col_cache::n2l_str(xs + 1.0) {
                        Ok(l) => crate::col_cache::decode_address(&format!(
                            "{l}{}",
                            crate::jsv::number_to_string(ys + 1.0)
                        ))
                        .ok()
                        .map(|a| a.address),
                        Err(e) => {
                            xml.fail(e);
                            None
                        }
                    }
                };
                let pane = if truthy(get(model, "xSplit")) && truthy(get(model, "ySplit")) {
                    "bottomRight"
                } else if truthy(get(model, "xSplit")) {
                    "topRight"
                } else {
                    "bottomLeft"
                };
                xml.leaf_node(
                    "pane",
                    &attrs![
                        ("xSplit", or_undef("xSplit")),
                        ("ySplit", or_undef("ySplit")),
                        ("topLeftCell", top_left),
                        ("activePane", Some(pane)),
                        ("state", Some("frozen")),
                    ],
                    None,
                );
                xml.leaf_node(
                    "selection",
                    &attrs![
                        ("pane", Some(pane)),
                        ("activeCell", v("activeCell")),
                        ("sqref", v("activeCell"))
                    ],
                    None,
                );
            }
            Some("split") => {
                let active = if as_str(get(model, "activePane")) == Some("topLeft") {
                    None
                } else {
                    v("activePane")
                };
                xml.leaf_node(
                    "pane",
                    &attrs![
                        ("xSplit", or_undef("xSplit")),
                        ("ySplit", or_undef("ySplit")),
                        ("topLeftCell", v("topLeftCell")),
                        ("activePane", active.clone()),
                    ],
                    None,
                );
                xml.leaf_node(
                    "selection",
                    &attrs![
                        ("pane", active),
                        ("activeCell", v("activeCell")),
                        ("sqref", v("activeCell"))
                    ],
                    None,
                );
            }
            Some("normal") if truthy(get(model, "activeCell")) => {
                xml.leaf_node(
                    "selection",
                    &attrs![("activeCell", v("activeCell")), ("sqref", v("activeCell"))],
                    None,
                );
            }
            _ => {}
        }
        xml.close_node();
    }
    fn reset(&mut self) {
        self.model = Some(Value::Null);
    }
}

fn n_or<'a>(node: &'a Tag<'_>, a: &str, d: &'static str) -> &'a str {
    node.attr(a).filter(|v| !v.is_empty()).unwrap_or(d)
}

/// `SheetViewXform.prepare(model)`.
pub(crate) fn prepare_sheet_view(view: &mut Value) {
    if let Value::Object(o) = view {
        let state = o.get("state").and_then(Value::as_str);
        if !matches!(state, Some("frozen" | "split")) {
            o.insert("state".into(), json!("normal"));
        }
    }
}

// ---------------------------------------------------------------------------
// DataValidationsXform

/// `DataValidationsXform`. The model maps each address to a data validation;
/// every address of one `<dataValidation>` shares the same object, so the
/// model is `{"$dvs": [dv...], "map": {address: index}}` and the reader turns
/// the indices into shared references.
#[derive(Debug, Default)]
pub(crate) struct DataValidationsXform {
    address: Option<String>,
    current: Option<Map<String, Value>>,
    formula: Option<String>,
    dvs: Vec<Value>,
    map: Option<Map<String, Value>>,
    model: Option<Value>,
}

impl DataValidationsXform {
    fn sync_model(&mut self) {
        self.model = self
            .map
            .as_ref()
            .map(|m| json!({"$dvs": self.dvs.clone(), "map": m}));
    }
}

impl Xform for DataValidationsXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        match node.name {
            "dataValidations" => {
                self.map = Some(Map::new());
                self.dvs.clear();
                self.sync_model();
                Ok(true)
            }
            "dataValidation" => {
                self.address = node.attr("sqref").map(str::to_owned);
                let mut dv = Map::new();
                dv.insert("type".into(), json!(node.attr("type").unwrap_or("any")));
                dv.insert("formulae".into(), json!([]));
                let bool_attr = |dv: &mut Map<String, Value>, a: &str| {
                    if let Some(v) = node.attr(a) {
                        dv.insert(a.into(), json!(parse_boolean(Some(v))));
                    }
                };
                if node.attr("type").is_some() {
                    bool_attr(&mut dv, "allowBlank");
                }
                bool_attr(&mut dv, "showInputMessage");
                bool_attr(&mut dv, "showErrorMessage");
                let ty = dv
                    .get("type")
                    .and_then(Value::as_str)
                    .unwrap_or("any")
                    .to_owned();
                if !matches!(ty.as_str(), "any" | "list" | "custom") {
                    dv.insert(
                        "operator".into(),
                        json!(node.attr("operator").unwrap_or("between")),
                    );
                }
                for a in ["promptTitle", "prompt", "errorStyle", "errorTitle", "error"] {
                    if let Some(v) = node.attr(a) {
                        dv.insert(a.into(), json!(v));
                    }
                }
                self.current = Some(dv);
                Ok(true)
            }
            "formula1" | "formula2" => {
                self.formula = Some(String::new());
                Ok(true)
            }
            _ => Ok(false),
        }
    }
    fn parse_text(&mut self, text: &str) {
        if let Some(f) = self.formula.as_mut() {
            f.push_str(text);
        }
    }
    fn parse_close(&mut self, name: &str) -> XResult<bool> {
        match name {
            "dataValidations" => Ok(false),
            "dataValidation" => {
                let mut dv = self.current.take().ok_or_else(|| {
                    "Cannot read properties of undefined (reading 'formulae')".to_owned()
                })?;
                if !dv
                    .get("formulae")
                    .and_then(Value::as_array)
                    .is_some_and(|a| !a.is_empty())
                {
                    dv.remove("formulae");
                    dv.remove("operator");
                }
                let address = self.address.clone().ok_or_else(|| {
                    "Cannot read properties of undefined (reading 'split')".to_owned()
                })?;
                let index = self.dvs.len();
                self.dvs.push(Value::Object(dv.clone()));
                // Keep the open element for following formula elements (JS
                // keeps `this._dataValidation`).
                self.current = Some(dv);
                let map = self
                    .map
                    .as_mut()
                    .ok_or_else(|| "Cannot set properties of undefined".to_owned())?;
                for addr in split_whitespace_js(&address) {
                    if addr.contains(':') {
                        if let DecodedEx::Range(r) = decode_ex(&addr)? {
                            if r.top != 0.0 && !r.top.is_nan() {
                                let mut c = r.left;
                                while c <= r.right {
                                    let mut row = r.top;
                                    while row <= r.bottom {
                                        let a = format!(
                                            "{}{}",
                                            crate::col_cache::n2l_str(c)?,
                                            crate::jsv::number_to_string(row)
                                        );
                                        map.insert(a, json!(index));
                                        row += 1.0;
                                    }
                                    c += 1.0;
                                }
                            }
                        }
                    } else {
                        map.insert(addr, json!(index));
                    }
                }
                self.sync_model();
                Ok(true)
            }
            "formula1" | "formula2" => {
                let text = self.formula.take().unwrap_or_default();
                let dv = self.current.as_mut().ok_or_else(|| {
                    "Cannot read properties of undefined (reading 'type')".to_owned()
                })?;
                let value = match dv.get("type").and_then(Value::as_str) {
                    Some("whole" | "textLength") => int_value(Some(&text)),
                    Some("decimal") => float_value(Some(&text)),
                    Some("date") => {
                        crate::model_json::excel_to_date(crate::js::parse_float(&text), false)
                    }
                    _ => json!(text),
                };
                match dv.get_mut("formulae") {
                    Some(Value::Array(a)) => a.push(value),
                    _ => return Err("Cannot read properties of undefined (reading 'push')".into()),
                }
                Ok(true)
            }
            _ => Ok(true),
        }
    }
    crate::model_accessors!();
    fn reset(&mut self) {
        self.model = Some(Value::Null);
        self.map = None;
        self.dvs.clear();
    }
}

/// `str.split(/\s+/g)` (JS whitespace).
fn split_whitespace_js(s: &str) -> Vec<String> {
    let mut out = vec![String::new()];
    let mut in_space = false;
    for c in s.chars() {
        if c.is_whitespace() || c == '\u{feff}' {
            if !in_space {
                out.push(String::new());
                in_space = true;
            }
        } else {
            in_space = false;
            if let Some(last) = out.last_mut() {
                last.push(c);
            }
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xform::parse_part;

    #[test]
    fn data_validations_share_one_object_per_element() {
        let xml = r#"<dataValidations><dataValidation type="whole" sqref="A1:A2 C3"><formula1>5</formula1></dataValidation></dataValidations>"#;
        let m = parse_part(&mut DataValidationsXform::default(), "x", xml.as_bytes())
            .unwrap()
            .unwrap();
        assert_eq!(
            m,
            json!({"$dvs": [{"type": "whole", "formulae": [5], "operator": "between"}], "map": {"A1": 0, "A2": 0, "C3": 0}})
        );
    }

    #[test]
    fn internal_links() {
        assert!(is_internal_link("Sheet2!D3"));
        assert!(!is_internal_link("Sheet2!$D$3"));
        assert!(!is_internal_link("http://x"));
    }
}
