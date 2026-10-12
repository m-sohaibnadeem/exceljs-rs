//! ExcelJS `lib/xlsx/xform/table/*.js`: `xl/tables/tableN.xml`.

use serde_json::{json, Value};

use super::base::ListXform;
use super::composite::{object, Leaf, LeafMatch};
use super::{XResult, Xform};
use crate::attrs;
use crate::jsv::{get, truthy};
use crate::xml::Tag;
use crate::xml_stream::{std_doc_attributes, XmlStream};

fn s(v: Option<&str>) -> Option<Value> {
    v.map(|x| json!(x))
}

fn custom_filter() -> Leaf {
    Leaf::new(
        "customFilter",
        LeafMatch::Tag,
        |n| {
            Some(object(vec![
                ("val", s(n.attr("val"))),
                ("operator", s(n.attr("operator"))),
            ]))
        },
        |xml, m| {
            xml.leaf_node(
                "customFilter",
                &attrs![("val", get(m, "val")), ("operator", get(m, "operator"))],
                None,
            )
        },
    )
}

fn filter() -> Leaf {
    Leaf::new(
        "filter",
        LeafMatch::Tag,
        |n| Some(object(vec![("val", s(n.attr("val")))])),
        |xml, m| xml.leaf_node("filter", &attrs![("val", get(m, "val"))], None),
    )
}

/// `FilterColumnXform`.
#[derive(Debug)]
struct FilterColumnXform {
    custom_filters: ListXform,
    filters: ListXform,
    active: Option<bool>, // true = customFilters
    model: Option<Value>,
}

impl FilterColumnXform {
    fn new() -> Self {
        FilterColumnXform {
            custom_filters: ListXform::new("customFilters", Box::new(custom_filter())).empty(),
            filters: ListXform::new("filters", Box::new(filter())).empty(),
            active: None,
            model: None,
        }
    }
    fn child(&mut self, custom: bool) -> &mut ListXform {
        if custom {
            &mut self.custom_filters
        } else {
            &mut self.filters
        }
    }
}

impl Xform for FilterColumnXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if let Some(c) = self.active {
            self.child(c).parse_open(node)?;
            return Ok(true);
        }
        match node.name {
            "filterColumn" => {
                self.model = Some(json!({"filterButton": node.attr("hiddenButton") == Some("0")}));
                Ok(true)
            }
            "customFilters" | "filters" => {
                let c = node.name == "customFilters";
                self.active = Some(c);
                self.child(c).parse_open(node)?;
                Ok(true)
            }
            other => Err(format!("Unexpected xml node in parseOpen: {other}")),
        }
    }
    fn parse_close(&mut self, name: &str) -> XResult<bool> {
        if let Some(c) = self.active {
            if !self.child(c).parse_close(name)? {
                self.active = None;
            }
            return Ok(true);
        }
        if name == "filterColumn" {
            let cf = self.custom_filters.model().cloned();
            if let (Some(Value::Object(m)), Some(cf)) = (self.model.as_mut(), cf) {
                m.insert("customFilters".into(), cf);
            }
            return Ok(false);
        }
        Ok(true)
    }
    crate::model_accessors!();
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        let a = attrs![
            ("colId", get(model, "colId")),
            (
                "hiddenButton",
                Some(if truthy(get(model, "filterButton")) {
                    "0"
                } else {
                    "1"
                })
            ),
        ];
        if truthy(get(model, "customFilters")) {
            xml.open_node("filterColumn", &a);
            self.custom_filters.render(xml, get(model, "customFilters"));
            xml.close_node();
        } else {
            xml.leaf_node("filterColumn", &a, None);
        }
    }
    fn reset(&mut self) {
        self.model = Some(Value::Null);
        self.custom_filters.reset();
        self.filters.reset();
    }
}

/// The table's `AutoFilterXform`.
#[derive(Debug)]
struct TableAutoFilterXform {
    column: FilterColumnXform,
    in_column: bool,
    model: Option<Value>,
}

impl Xform for TableAutoFilterXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if self.in_column {
            self.column.parse_open(node)?;
            return Ok(true);
        }
        match node.name {
            "autoFilter" => {
                self.model = Some(object(vec![
                    ("autoFilterRef", s(node.attr("ref"))),
                    ("columns", Some(json!([]))),
                ]));
                Ok(true)
            }
            "filterColumn" => {
                self.in_column = true;
                self.column.parse_open(node)?;
                Ok(true)
            }
            other => Err(format!("Unexpected xml node in parseOpen: {other}")),
        }
    }
    fn parse_close(&mut self, name: &str) -> XResult<bool> {
        if self.in_column {
            if !self.column.parse_close(name)? {
                let c = self.column.model().cloned().unwrap_or(Value::Null);
                match self
                    .model
                    .as_mut()
                    .and_then(|m| m.get_mut("columns"))
                    .and_then(Value::as_array_mut)
                {
                    Some(cols) => cols.push(c),
                    None => {
                        return Err("Cannot read properties of undefined (reading 'columns')".into())
                    }
                }
                self.in_column = false;
            }
            return Ok(true);
        }
        match name {
            "autoFilter" => Ok(false),
            other => Err(format!("Unexpected xml node in parseClose: {other}")),
        }
    }
    crate::model_accessors!();
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        xml.open_node("autoFilter", &attrs![("ref", get(model, "autoFilterRef"))]);
        for c in get(model, "columns")
            .and_then(Value::as_array)
            .into_iter()
            .flatten()
        {
            self.column.render(xml, Some(c));
        }
        xml.close_node();
    }
    fn reset(&mut self) {
        self.model = Some(Value::Null);
        self.column.reset();
    }
}

fn table_column() -> Leaf {
    Leaf::new(
        "tableColumn",
        LeafMatch::Tag,
        |n| {
            Some(object(vec![
                ("name", s(n.attr("name"))),
                ("totalsRowLabel", s(n.attr("totalsRowLabel"))),
                ("totalsRowFunction", s(n.attr("totalsRowFunction"))),
                ("dxfId", s(n.attr("dxfId"))),
            ]))
        },
        |xml, m| {
            xml.leaf_node(
                "tableColumn",
                &attrs![
                    ("id", get(m, "id")),
                    ("name", get(m, "name")),
                    ("totalsRowLabel", get(m, "totalsRowLabel")),
                    ("totalsRowFunction", get(m, "totalsRowFunction")),
                    ("dxfId", get(m, "dxfId")),
                ],
                None,
            );
        },
    )
}

fn table_style_info() -> Leaf {
    Leaf::new(
        "tableStyleInfo",
        LeafMatch::Tag,
        |n| {
            let flag = |a: &str| json!(n.attr(a) == Some("1"));
            Some(json!({
                "theme": n.attr("name").filter(|v| !v.is_empty()).map_or(Value::Null, |v| json!(v)),
                "showFirstColumn": flag("showFirstColumn"),
                "showLastColumn": flag("showLastColumn"),
                "showRowStripes": flag("showRowStripes"),
                "showColumnStripes": flag("showColumnStripes"),
            }))
        },
        |xml, m| {
            let flag = |k: &str| Some(if truthy(get(m, k)) { "1" } else { "0" });
            xml.leaf_node(
                "tableStyleInfo",
                &attrs![
                    (
                        "name",
                        if truthy(get(m, "theme")) {
                            get(m, "theme").cloned()
                        } else {
                            None
                        }
                    ),
                    ("showFirstColumn", flag("showFirstColumn")),
                    ("showLastColumn", flag("showLastColumn")),
                    ("showRowStripes", flag("showRowStripes")),
                    ("showColumnStripes", flag("showColumnStripes")),
                ],
                None,
            );
        },
    )
}

/// `TableXform`.
#[derive(Debug)]
pub(crate) struct TableXform {
    auto_filter: TableAutoFilterXform,
    columns: ListXform,
    style_info: Leaf,
    active: Option<u8>,
    model: Option<Value>,
}

impl TableXform {
    pub fn new() -> Self {
        TableXform {
            auto_filter: TableAutoFilterXform {
                column: FilterColumnXform::new(),
                in_column: false,
                model: None,
            },
            columns: ListXform::new("tableColumns", Box::new(table_column()))
                .count()
                .empty(),
            style_info: table_style_info(),
            active: None,
            model: None,
        }
    }
    fn child(&mut self, i: u8) -> &mut dyn Xform {
        match i {
            0 => &mut self.auto_filter,
            1 => &mut self.columns,
            _ => &mut self.style_info,
        }
    }
}

impl Xform for TableXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if let Some(i) = self.active {
            self.child(i).parse_open(node)?;
            return Ok(true);
        }
        if node.name == "table" {
            self.reset();
            let name = node.attr("name");
            let display = node.attr("displayName").filter(|v| !v.is_empty()).or(name);
            self.model = Some(object(vec![
                ("name", s(name)),
                ("displayName", s(display)),
                ("tableRef", s(node.attr("ref"))),
                (
                    "totalsRow",
                    Some(json!(node.attr("totalsRowCount") == Some("1"))),
                ),
                (
                    "headerRow",
                    Some(json!(node.attr("headerRowCount") == Some("1"))),
                ),
            ]));
            return Ok(true);
        }
        self.active = match node.name {
            "autoFilter" => Some(0),
            "tableColumns" => Some(1),
            "tableStyleInfo" => Some(2),
            _ => None,
        };
        if let Some(i) = self.active {
            self.child(i).parse_open(node)?;
        }
        Ok(true)
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
        if name != "table" {
            return Ok(true);
        }
        let mut columns = self.columns.model().cloned();
        let af = self
            .auto_filter
            .model()
            .filter(|m| truthy(Some(m)))
            .cloned();
        let m = match self.model.as_mut() {
            Some(Value::Object(m)) => m,
            _ => return Err("Cannot set properties of undefined (setting 'columns')".into()),
        };
        if let Some(af) = &af {
            for (i, col) in get(Some(af), "columns")
                .and_then(Value::as_array)
                .into_iter()
                .flatten()
                .enumerate()
            {
                match columns
                    .as_mut()
                    .and_then(|c| c.get_mut(i))
                    .and_then(Value::as_object_mut)
                {
                    Some(c) => {
                        match col.get("filterButton") {
                            Some(fb) => c.insert("filterButton".into(), fb.clone()),
                            None => c.remove("filterButton"),
                        };
                    }
                    None => {
                        return Err(
                            "Cannot set properties of undefined (setting 'filterButton')".into(),
                        )
                    }
                }
            }
        }
        if let Some(c) = columns {
            m.insert("columns".into(), c);
        }
        if let Some(af) = &af {
            match af.get("autoFilterRef") {
                Some(r) => m.insert("autoFilterRef".into(), r.clone()),
                None => m.remove("autoFilterRef"),
            };
        }
        match self.style_info.model() {
            Some(st) => m.insert("style".into(), st.clone()),
            None => m.remove("style"),
        };
        Ok(false)
    }
    crate::model_accessors!();
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        xml.open_xml(&std_doc_attributes());
        let display = if truthy(get(model, "displayName")) {
            get(model, "displayName")
        } else {
            get(model, "name")
        };
        xml.open_node(
            "table",
            &attrs![
                (
                    "xmlns",
                    Some("http://schemas.openxmlformats.org/spreadsheetml/2006/main")
                ),
                (
                    "xmlns:mc",
                    Some("http://schemas.openxmlformats.org/markup-compatibility/2006")
                ),
                ("mc:Ignorable", Some("xr xr3")),
                (
                    "xmlns:xr",
                    Some("http://schemas.microsoft.com/office/spreadsheetml/2014/revision")
                ),
                (
                    "xmlns:xr3",
                    Some("http://schemas.microsoft.com/office/spreadsheetml/2016/revision3")
                ),
                ("id", get(model, "id")),
                ("name", get(model, "name")),
                ("displayName", display),
                ("ref", get(model, "tableRef")),
                (
                    "totalsRowCount",
                    truthy(get(model, "totalsRow")).then_some("1")
                ),
                (
                    "totalsRowShown",
                    (!truthy(get(model, "totalsRow"))).then_some("1")
                ),
                (
                    "headerRowCount",
                    Some(if truthy(get(model, "headerRow")) {
                        "1"
                    } else {
                        "0"
                    })
                ),
            ],
        );
        self.auto_filter.render(xml, model);
        self.columns.render(xml, get(model, "columns"));
        self.style_info.render(xml, get(model, "style"));
        xml.close_node();
    }
    fn reset(&mut self) {
        self.model = Some(Value::Null);
        self.auto_filter.reset();
        self.columns.reset();
        self.style_info.reset();
    }
}

/// `TableColumnXform.prepare` / `FilterColumnXform.prepare` (writer):
/// `column.id = index + 1`, `column.colId = index.toString()`.
pub(crate) fn prepare_table(table: &mut Value) {
    if let Some(Value::Array(cols)) = table.get_mut("columns") {
        for (i, c) in cols.iter_mut().enumerate() {
            if let Value::Object(o) = c {
                o.insert("colId".into(), json!(i.to_string()));
                o.insert("id".into(), json!(i + 1));
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xform::parse_part;

    #[test]
    fn parses_table() {
        let xml = r#"<table id="1" name="T" displayName="T" ref="A1:B3" headerRowCount="1" totalsRowCount="1"><autoFilter ref="A1:B2"><filterColumn colId="0" hiddenButton="0"/></autoFilter><tableColumns count="2"><tableColumn id="1" name="a" dxfId="3"/><tableColumn id="2" name="b" totalsRowFunction="sum"/></tableColumns><tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/></table>"#;
        let m = parse_part(&mut TableXform::new(), "t", xml.as_bytes())
            .unwrap()
            .unwrap();
        assert_eq!(
            m,
            json!({
                "name": "T", "displayName": "T", "tableRef": "A1:B3", "totalsRow": true, "headerRow": true,
                "columns": [{"name": "a", "dxfId": "3", "filterButton": true}, {"name": "b", "totalsRowFunction": "sum"}],
                "autoFilterRef": "A1:B2",
                "style": {"theme": "TableStyleMedium2", "showFirstColumn": false, "showLastColumn": false, "showRowStripes": true, "showColumnStripes": false},
            })
        );
    }
}
