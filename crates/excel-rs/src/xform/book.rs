//! ExcelJS `lib/xlsx/xform/book/*.js`: `xl/workbook.xml`.

use serde_json::{json, Map, Value};

use super::base::{Children, ListXform, StaticNode, StaticXform};
use super::simple::int_value;
use super::{XResult, Xform};
use crate::attrs;
use crate::col_cache::decode_ex;
use crate::jsv::{get, to_js_string, truthy};
use crate::xml::Tag;
use crate::xml_stream::{std_doc_attributes, XmlStream};

/// `extractRanges(parsedText)` from defined-name-xform.js.
fn extract_ranges(text: &str) -> Vec<String> {
    let valid = |r: &str| decode_ex(r).is_ok();
    let mut ranges = Vec::new();
    let mut quotes_opened = false;
    let mut last = String::new();
    for item in text.split(',') {
        if item.is_empty() {
            continue;
        }
        let quotes = item.matches('\'').count();
        if quotes == 0 {
            if quotes_opened {
                last.push_str(item);
                last.push(',');
            } else if valid(item) {
                ranges.push(item.to_owned());
            }
            continue;
        }
        let even = quotes % 2 == 0;
        if !quotes_opened && even && valid(item) {
            ranges.push(item.to_owned());
        } else if quotes_opened && !even {
            quotes_opened = false;
            let joined = format!("{last}{item}");
            if valid(&joined) {
                ranges.push(joined);
            }
            last.clear();
        } else {
            quotes_opened = true;
            last.push_str(item);
            last.push(',');
        }
    }
    ranges
}

/// `DefinedNamesXform` (one `<definedName>`).
#[derive(Debug, Default)]
pub(crate) struct DefinedNameXform {
    name: Option<String>,
    local_sheet_id: Option<String>,
    text: Option<String>,
    model: Option<Value>,
}

impl Xform for DefinedNameXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if node.name == "definedName" {
            self.name = node.attr("name").map(str::to_owned);
            self.local_sheet_id = node.attr("localSheetId").map(str::to_owned);
            self.text = Some(String::new());
            return Ok(true);
        }
        Ok(false)
    }
    fn parse_text(&mut self, text: &str) {
        if let Some(t) = self.text.as_mut() {
            t.push_str(text);
        }
    }
    fn parse_close(&mut self, _name: &str) -> XResult<bool> {
        let text = self
            .text
            .as_deref()
            .ok_or_else(|| "Cannot read properties of undefined (reading 'join')".to_owned())?;
        let mut m = Map::new();
        if let Some(n) = &self.name {
            m.insert("name".into(), json!(n));
        }
        m.insert("ranges".into(), json!(extract_ranges(text)));
        if let Some(id) = &self.local_sheet_id {
            m.insert("localSheetId".into(), int_value(Some(id)));
        }
        self.model = Some(Value::Object(m));
        Ok(false)
    }
    crate::model_accessors!();
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        xml.open_node(
            "definedName",
            &attrs![
                ("name", get(model, "name")),
                ("localSheetId", get(model, "localSheetId"))
            ],
        );
        let ranges = get(model, "ranges");
        match ranges {
            Some(Value::Array(_)) => xml.write_text(&to_js_string(ranges)),
            _ => xml.fail("Cannot read properties of undefined (reading 'join')"),
        }
        xml.close_node();
    }
}

/// `sheet-xform.js` (one `<sheet>` in `<sheets>`).
#[derive(Debug, Default)]
pub(crate) struct SheetXform {
    model: Option<Value>,
}

impl Xform for SheetXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if node.name != "sheet" {
            return Ok(false);
        }
        let name = node
            .attr("name")
            .ok_or_else(|| "Cannot read properties of undefined (reading 'replace')".to_owned())?;
        let mut m = Map::new();
        m.insert("name".into(), json!(name));
        m.insert("id".into(), int_value(node.attr("sheetId")));
        if let Some(s) = node.attr("state") {
            m.insert("state".into(), json!(s));
        }
        if let Some(r) = node.attr("r:id") {
            m.insert("rId".into(), json!(r));
        }
        self.model = Some(Value::Object(m));
        Ok(true)
    }
    fn parse_close(&mut self, _name: &str) -> XResult<bool> {
        Ok(false)
    }
    crate::model_accessors!();
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        xml.leaf_node(
            "sheet",
            &attrs![
                ("sheetId", get(model, "id")),
                ("name", get(model, "name")),
                ("state", get(model, "state")),
                ("r:id", get(model, "rId")),
            ],
            None,
        );
    }
}

/// `workbook-view-xform.js`.
#[derive(Debug, Default)]
pub(crate) struct WorkbookViewXform {
    model: Option<Value>,
}

fn or_default(v: Option<&Value>, dflt: i64) -> Value {
    if truthy(v) {
        v.cloned().unwrap_or(Value::Null)
    } else {
        json!(dflt)
    }
}

impl Xform for WorkbookViewXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if node.name != "workbookView" {
            return Ok(false);
        }
        let mut m = Map::new();
        let mut add_n = |name: &str, attr: &str, dflt: Option<i64>| match node.attr(attr) {
            Some(v) => {
                m.insert(name.into(), int_value(Some(v)));
            }
            None => {
                if let Some(d) = dflt {
                    m.insert(name.into(), json!(d));
                }
            }
        };
        add_n("x", "xWindow", Some(0));
        add_n("y", "yWindow", Some(0));
        add_n("width", "windowWidth", Some(25000));
        add_n("height", "windowHeight", Some(10000));
        m.insert(
            "visibility".into(),
            json!(node.attr("visibility").unwrap_or("visible")),
        );
        let mut add_n = |name: &str, attr: &str| {
            if let Some(v) = node.attr(attr) {
                m.insert(name.into(), int_value(Some(v)));
            }
        };
        add_n("activeTab", "activeTab");
        add_n("firstSheet", "firstSheet");
        self.model = Some(Value::Object(m));
        Ok(true)
    }
    fn parse_close(&mut self, _name: &str) -> XResult<bool> {
        Ok(false)
    }
    crate::model_accessors!();
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        let mut a = attrs![
            ("xWindow", Some(or_default(get(model, "x"), 0))),
            ("yWindow", Some(or_default(get(model, "y"), 0))),
            ("windowWidth", Some(or_default(get(model, "width"), 12000))),
            (
                "windowHeight",
                Some(or_default(get(model, "height"), 24000))
            ),
            ("firstSheet", get(model, "firstSheet")),
            ("activeTab", get(model, "activeTab")),
        ];
        let vis = get(model, "visibility");
        if truthy(vis) && crate::jsv::as_str(vis) != Some("visible") {
            a.push(("visibility".into(), Some(to_js_string(vis))));
        }
        xml.leaf_node("workbookView", &a, None);
    }
}

/// `workbook-properties-xform.js` (`<workbookPr>`).
#[derive(Debug, Default)]
pub(crate) struct WorkbookPropertiesXform {
    model: Option<Value>,
}

impl Xform for WorkbookPropertiesXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if node.name == "workbookPr" {
            self.model = Some(json!({"date1904": node.attr("date1904") == Some("1")}));
            return Ok(true);
        }
        Ok(false)
    }
    fn parse_close(&mut self, _name: &str) -> XResult<bool> {
        Ok(false)
    }
    crate::model_accessors!();
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        if model.is_none() {
            xml.fail("Cannot read properties of undefined (reading 'date1904')");
        }
        xml.leaf_node(
            "workbookPr",
            &attrs![
                ("date1904", truthy(get(model, "date1904")).then_some(1i64)),
                ("defaultThemeVersion", Some(164_011i64)),
                ("filterPrivacy", Some(1i64)),
            ],
            None,
        );
    }
}

/// `workbook-calc-properties-xform.js` (`<calcPr>`).
#[derive(Debug, Default)]
pub(crate) struct CalcPropertiesXform {
    model: Option<Value>,
}

impl Xform for CalcPropertiesXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if node.name == "calcPr" {
            self.model = Some(json!({}));
            return Ok(true);
        }
        Ok(false)
    }
    fn parse_close(&mut self, _name: &str) -> XResult<bool> {
        Ok(false)
    }
    crate::model_accessors!();
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        if model.is_none() {
            xml.fail("Cannot read properties of undefined (reading 'fullCalcOnLoad')");
        }
        xml.leaf_node(
            "calcPr",
            &attrs![
                ("calcId", Some(171_027i64)),
                (
                    "fullCalcOnLoad",
                    truthy(get(model, "fullCalcOnLoad")).then_some(1i64)
                )
            ],
            None,
        );
    }
}

fn file_version_static() -> StaticXform {
    StaticXform::new(StaticNode::new(
        "fileVersion",
        attrs![
            ("appName", Some("xl")),
            ("lastEdited", Some(5i64)),
            ("lowestEdited", Some(5i64)),
            ("rupBuild", Some(9303i64)),
        ],
        vec![],
    ))
}

/// `WorkbookXform`.
#[derive(Debug)]
pub(crate) struct WorkbookXform {
    map: Children,
    model: Option<Value>,
}

impl WorkbookXform {
    pub fn new() -> Self {
        WorkbookXform {
            map: Children::new(vec![
                ("fileVersion", Box::new(file_version_static())),
                ("workbookPr", Box::new(WorkbookPropertiesXform::default())),
                (
                    "bookViews",
                    Box::new(ListXform::new(
                        "bookViews",
                        Box::new(WorkbookViewXform::default()),
                    )),
                ),
                (
                    "sheets",
                    Box::new(ListXform::new("sheets", Box::new(SheetXform::default()))),
                ),
                (
                    "definedNames",
                    Box::new(ListXform::new(
                        "definedNames",
                        Box::new(DefinedNameXform::default()),
                    )),
                ),
                ("calcPr", Box::new(CalcPropertiesXform::default())),
            ]),
            model: None,
        }
    }
}

impl Xform for WorkbookXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if self.map.forward_open(node)? {
            return Ok(true);
        }
        if node.name == "workbook" {
            return Ok(true);
        }
        if self.map.activate(node.name) {
            if let Some(p) = self.map.parser() {
                p.parse_open(node)?;
            }
        }
        Ok(true)
    }
    fn parse_text(&mut self, text: &str) {
        self.map.forward_text(text);
    }
    fn parse_close(&mut self, name: &str) -> XResult<bool> {
        if let Some(p) = self.map.parser() {
            if !p.parse_close(name)? {
                self.map.active = None;
            }
            return Ok(true);
        }
        if name == "workbook" {
            let mut m = Map::new();
            if let Some(s) = self.map.model_of("sheets") {
                m.insert("sheets".into(), s.clone());
            }
            let props = self
                .map
                .model_of("workbookPr")
                .filter(|p| truthy(Some(p)))
                .cloned()
                .unwrap_or_else(|| json!({}));
            m.insert("properties".into(), props);
            if let Some(v) = self.map.model_of("bookViews") {
                m.insert("views".into(), v.clone());
            }
            m.insert("calcProperties".into(), json!({}));
            if let Some(d) = self
                .map
                .model_of("definedNames")
                .filter(|d| truthy(Some(d)))
            {
                m.insert("definedNames".into(), d.clone());
            }
            self.model = Some(Value::Object(m));
            return Ok(false);
        }
        Ok(true)
    }
    crate::model_accessors!();
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        xml.open_xml(&std_doc_attributes());
        xml.open_node(
            "workbook",
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
                ("mc:Ignorable", Some("x15")),
                (
                    "xmlns:x15",
                    Some("http://schemas.microsoft.com/office/spreadsheetml/2010/11/main")
                ),
            ],
        );
        for (child, key) in [
            ("fileVersion", None),
            ("workbookPr", Some("properties")),
            ("bookViews", Some("views")),
            ("sheets", Some("sheets")),
            ("definedNames", Some("definedNames")),
            ("calcPr", Some("calcProperties")),
        ] {
            if let Some(x) = self.map.get(child) {
                x.render(xml, key.and_then(|k| get(model, k)));
            }
        }
        xml.close_node();
    }
}

/// What the workbook-level parts read of a worksheet model (`<sheet>`,
/// app.xml titles, content types): its id, name, state and rId.
pub(crate) fn sheet_summary(ws: &Value) -> Value {
    let mut m = serde_json::Map::new();
    for k in ["id", "name", "state", "rId"] {
        if let Some(v) = get(Some(ws), k) {
            m.insert(k.into(), v.clone());
        }
    }
    Value::Object(m)
}

/// `WorkbookXform.prepare(model)` (writer). `model.sheets` gets the sheet
/// summaries (the render reads nothing else of them).
pub(crate) fn prepare(model: &mut Value) -> XResult<()> {
    let empty = Vec::new();
    let worksheets = match get(Some(model), "worksheets") {
        Some(Value::Array(a)) => a,
        _ => &empty,
    };
    let sheets = match get(Some(model), "worksheets") {
        Some(Value::Array(_)) => Value::Array(worksheets.iter().map(sheet_summary).collect()),
        Some(other) => other.clone(),
        None => Value::Null,
    };
    let mut print_areas = Vec::new();
    for (index, sheet) in worksheets.iter().enumerate() {
        let page_setup = get(Some(sheet), "pageSetup");
        let name = to_js_string(get(Some(sheet), "name"));
        if truthy(page_setup) && truthy(get(page_setup, "printArea")) {
            for area in to_js_string(get(page_setup, "printArea")).split("&&") {
                let c: Vec<&str> = area.split(':').collect();
                let a = c.first().copied().unwrap_or("undefined");
                let b = c.get(1).copied().unwrap_or("undefined");
                print_areas.push(json!({
                    "name": "_xlnm.Print_Area",
                    "ranges": [format!("'{name}'!${a}:${b}")],
                    "localSheetId": index,
                }));
            }
        }
        if truthy(page_setup)
            && (truthy(get(page_setup, "printTitlesRow"))
                || truthy(get(page_setup, "printTitlesColumn")))
        {
            let mut ranges = Vec::new();
            for key in ["printTitlesColumn", "printTitlesRow"] {
                if truthy(get(page_setup, key)) {
                    let s = to_js_string(get(page_setup, key));
                    let c: Vec<&str> = s.split(':').collect();
                    let a = c.first().copied().unwrap_or("undefined");
                    let b = c.get(1).copied().unwrap_or("undefined");
                    ranges.push(json!(format!("'{name}'!${a}:${b}")));
                }
            }
            print_areas.push(
                json!({"name": "_xlnm.Print_Titles", "ranges": ranges, "localSheetId": index}),
            );
        }
    }
    let Value::Object(m) = model else {
        return Err("model is not an object".into());
    };
    m.insert("sheets".into(), sheets);
    if !print_areas.is_empty() {
        let mut names = match m.get("definedNames") {
            Some(Value::Array(a)) => a.clone(),
            _ => return Err("Cannot read properties of undefined (reading 'concat')".into()),
        };
        names.extend(print_areas);
        m.insert("definedNames".into(), Value::Array(names));
    }
    if let Some(Value::Array(media)) = m.get_mut("media") {
        for (i, medium) in media.iter_mut().enumerate() {
            if let Value::Object(o) = medium {
                let t = to_js_string(o.get("type"));
                o.insert("name".into(), json!(format!("{t}{}", i + 1)));
            }
        }
    }
    Ok(())
}

/// `rel.Target.replace(/^(\s|\/xl\/)+/, '')`.
pub(crate) fn strip_target_prefix(mut target: &str) -> &str {
    loop {
        let trimmed = target.strip_prefix("/xl/").or_else(|| {
            let mut chars = target.chars();
            chars
                .next()
                .filter(|c| crate::js::is_js_whitespace(*c))
                .map(|_| chars.as_str())
        });
        match trimmed {
            Some(rest) => target = rest,
            None => return target,
        }
    }
}

/// The print-area / print-titles part of `WorkbookXform.reconcile`.
/// `order[i]` is the index in `worksheets` of the i-th referenced sheet in
/// `<sheets>` order (ExcelJS's local `worksheets` array).
pub(crate) fn reconcile_defined_names(
    defined_names: Option<&Value>,
    worksheets: &mut [Value],
    order: &[usize],
) -> XResult<Vec<Value>> {
    let mut kept = Vec::new();
    for dn in defined_names
        .and_then(Value::as_array)
        .into_iter()
        .flatten()
    {
        let name = crate::jsv::as_str(get(Some(dn), "name"));
        let sheet_index = crate::jsv::as_f64(get(Some(dn), "localSheetId"));
        let target = sheet_index
            .filter(|i| *i >= 0.0 && i.fract() == 0.0)
            .and_then(|i| {
                #[allow(clippy::cast_possible_truncation, clippy::cast_sign_loss)]
                let i = i as usize;
                order.get(i).copied()
            })
            .and_then(|idx| worksheets.get_mut(idx));
        match name {
            Some("_xlnm.Print_Area") => {
                if let Some(ws) = target {
                    let first = get(get(Some(dn), "ranges"), "0");
                    let first = crate::jsv::as_str(first).ok_or_else(|| {
                        "Cannot read properties of undefined (reading 'match')".to_owned()
                    })?;
                    let dims = match decode_ex(first)? {
                        crate::col_cache::DecodedEx::Range(r) => r.dimensions,
                        _ => "undefined".to_owned(),
                    };
                    if let Some(ps) = page_setup_mut(ws) {
                        let new = match ps.get("printArea").filter(|v| truthy(Some(v))) {
                            Some(existing) => format!("{}&&{dims}", to_js_string(Some(existing))),
                            None => dims,
                        };
                        ps.insert("printArea".into(), json!(new));
                    }
                }
            }
            Some("_xlnm.Print_Titles") => {
                if let Some(ws) = target {
                    let ranges = to_js_string(get(Some(dn), "ranges"));
                    if let Some(ps) = page_setup_mut(ws) {
                        if let Some(r) = find_row_range(&ranges) {
                            ps.insert("printTitlesRow".into(), json!(r.replace('$', "")));
                        }
                        if let Some(r) = find_col_range(&ranges) {
                            ps.insert("printTitlesColumn".into(), json!(r.replace('$', "")));
                        }
                    }
                }
            }
            _ => kept.push(dn.clone()),
        }
    }
    Ok(kept)
}

/// `if (!worksheet.pageSetup) worksheet.pageSetup = {}` then the object.
fn page_setup_mut(ws: &mut Value) -> Option<&mut Map<String, Value>> {
    let o = ws.as_object_mut()?;
    if !truthy(o.get("pageSetup")) {
        o.insert("pageSetup".into(), json!({}));
    }
    o.get_mut("pageSetup")?.as_object_mut()
}

/// `/\$\d+:\$\d+/`.
fn find_row_range(s: &str) -> Option<&str> {
    find_pattern(s, |c| c.is_ascii_digit())
}

/// `/\$[A-Z]+:\$[A-Z]+/`.
fn find_col_range(s: &str) -> Option<&str> {
    find_pattern(s, |c| c.is_ascii_uppercase())
}

fn find_pattern(s: &str, class: fn(u8) -> bool) -> Option<&str> {
    let b = s.as_bytes();
    for start in 0..b.len() {
        if b.get(start) != Some(&b'$') {
            continue;
        }
        let run = |from: usize| {
            b.get(from..)
                .map_or(0, |r| r.iter().take_while(|c| class(**c)).count())
        };
        let n1 = run(start + 1);
        if n1 == 0 {
            continue;
        }
        let colon = start + 1 + n1;
        if b.get(colon) != Some(&b':') || b.get(colon + 1) != Some(&b'$') {
            continue;
        }
        let n2 = run(colon + 2);
        if n2 == 0 {
            continue;
        }
        return s.get(start..colon + 2 + n2);
    }
    None
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xform::parse_part;

    #[test]
    fn parses_workbook() {
        let xml = r#"<workbook><workbookPr date1904="1"/><bookViews><workbookView xWindow="10" activeTab="1"/></bookViews><sheets><sheet name="a &amp;lt; b" sheetId="2" r:id="rId1"/></sheets><definedNames><definedName name="x" localSheetId="0">'S 1'!$A$1:$B$2,Sheet2!C3</definedName></definedNames><calcPr calcId="1"/></workbook>"#;
        let m = parse_part(&mut WorkbookXform::new(), "wb", xml.as_bytes())
            .unwrap()
            .unwrap();
        assert_eq!(
            m,
            json!({
                "sheets": [{"name": "a &lt; b", "id": 2, "rId": "rId1"}],
                "properties": {"date1904": true},
                "views": [{"x": 10, "y": 0, "width": 25000, "height": 10000, "visibility": "visible", "activeTab": 1}],
                "calcProperties": {},
                "definedNames": [{"name": "x", "ranges": ["'S 1'!$A$1:$B$2", "Sheet2!C3"], "localSheetId": 0}],
            })
        );
    }

    #[test]
    fn ranges_with_commas_in_quotes() {
        assert_eq!(extract_ranges("'a,b'!A1,C1"), vec!["'a,b'!A1", "C1"]);
        assert_eq!(find_row_range("'s'!$1:$2,'s'!$A:$B"), Some("$1:$2"));
        assert_eq!(find_col_range("'s'!$1:$2,'s'!$A:$B"), Some("$A:$B"));
    }
}
