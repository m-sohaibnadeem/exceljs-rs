//! ExcelJS `lib/xlsx/xform/sheet/worksheet-xform.js` (parse; the writer's
//! prepare/render live in `crate::writer`).

use serde_json::{json, Map, Value};

use super::base::{Children, ListXform};
use super::cf::{ext_lst, merge_conditional_formattings, ConditionalFormattingsXform};
use super::sheet::{
    auto_filter, col, dimension, drawing, hyperlink, merge_cell, page_margins, page_setup, picture,
    print_options, sheet_format_properties, sheet_protection, table_part, DataValidationsXform,
    HeaderFooterXform, RowBreaksXform, SheetPropertiesXform, SheetViewXform,
};
use super::{XResult, Xform};
use crate::jsv::{get, truthy};
use crate::sheet_data::{SharedRows, SheetDataXform};
use crate::xml::Tag;

/// Reader options that reach the worksheet xform (`options` of
/// `workbook.xlsx.load(data, options)`).
#[derive(Debug, Clone, Default)]
pub(crate) struct SheetOptions {
    pub max_rows: Option<usize>,
    pub max_cols: Option<usize>,
    pub ignore_nodes: Vec<String>,
}

impl SheetOptions {
    /// No limits and no ignored nodes.
    pub fn is_default(&self) -> bool {
        self.max_rows.is_none() && self.max_cols.is_none() && self.ignore_nodes.is_empty()
    }
}

/// `WorkSheetXform` (parse role).
#[derive(Debug)]
pub(crate) struct WorksheetXform {
    map: Children,
    rows: SharedRows,
    ignore_nodes: Vec<String>,
    model: Option<Value>,
}

impl WorksheetXform {
    pub fn new(options: &SheetOptions) -> Self {
        let rows = SharedRows::default();
        WorksheetXform {
            map: Children::new(vec![
                ("sheetPr", Box::new(SheetPropertiesXform::new())),
                ("dimension", Box::new(dimension())),
                (
                    "sheetViews",
                    Box::new(ListXform::new(
                        "sheetViews",
                        Box::new(SheetViewXform::default()),
                    )),
                ),
                ("sheetFormatPr", Box::new(sheet_format_properties())),
                ("cols", Box::new(ListXform::new("cols", Box::new(col())))),
                (
                    "sheetData",
                    Box::new(SheetDataXform::new(
                        options.max_rows,
                        options.max_cols,
                        rows.clone(),
                    )),
                ),
                ("autoFilter", Box::new(auto_filter())),
                (
                    "mergeCells",
                    Box::new(ListXform::new("mergeCells", Box::new(merge_cell())).count()),
                ),
                ("rowBreaks", Box::new(RowBreaksXform::new())),
                (
                    "hyperlinks",
                    Box::new(ListXform::new("hyperlinks", Box::new(hyperlink()))),
                ),
                ("pageMargins", Box::new(page_margins())),
                ("dataValidations", Box::new(DataValidationsXform::default())),
                ("pageSetup", Box::new(page_setup())),
                ("headerFooter", Box::new(HeaderFooterXform::default())),
                ("printOptions", Box::new(print_options())),
                ("picture", Box::new(picture())),
                ("drawing", Box::new(drawing())),
                ("sheetProtection", Box::new(sheet_protection())),
                (
                    "tableParts",
                    Box::new(ListXform::new("tableParts", Box::new(table_part())).count()),
                ),
                (
                    "conditionalFormatting",
                    Box::new(ConditionalFormattingsXform::new()),
                ),
                ("extLst", Box::new(ext_lst())),
            ]),
            ignore_nodes: options.ignore_nodes.clone(),
            rows,
            model: None,
        }
    }

    /// The rows of `<sheetData>` (taken with the model).
    pub fn take_rows(&self) -> Vec<crate::rows::RowModel> {
        self.rows.take()
    }

    fn take(&mut self, name: &str) -> Option<Value> {
        self.map.get_mut(name).and_then(|x| x.take_model())
    }
}

impl Xform for WorksheetXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if self.map.forward_open(node)? {
            return Ok(true);
        }
        if node.name == "worksheet" {
            self.map.reset_all();
            return Ok(true);
        }
        if self.map.index(node.name).is_some() && !self.ignore_nodes.iter().any(|n| n == node.name)
        {
            self.map.activate(node.name);
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
        if name != "worksheet" {
            return Ok(true);
        }
        let sheet_pr = self.map.model_of("sheetPr").cloned();
        let mut properties = match self
            .map
            .model_of("sheetFormatPr")
            .filter(|v| truthy(Some(v)))
        {
            Some(Value::Object(o)) => o.clone(),
            _ => Map::new(),
        };
        if let Some(tc) = get(sheet_pr.as_ref(), "tabColor").filter(|v| truthy(Some(v))) {
            properties.insert("tabColor".into(), tc.clone());
        }
        if let Some(op) = get(sheet_pr.as_ref(), "outlineProperties").filter(|v| truthy(Some(v))) {
            properties.insert("outlineProperties".into(), op.clone());
        }
        let fit = get(get(sheet_pr.as_ref(), "pageSetup"), "fitToPage")
            .filter(|v| truthy(Some(v)))
            .cloned()
            .unwrap_or(json!(false));
        let mut page_setup = Map::new();
        page_setup.insert("fitToPage".into(), fit);
        if let Some(m) = self.map.model_of("pageMargins") {
            page_setup.insert("margins".into(), m.clone());
        }
        for src in ["pageSetup", "printOptions"] {
            if let Some(Value::Object(o)) = self.map.model_of(src) {
                for (k, v) in o {
                    page_setup.insert(k.clone(), v.clone());
                }
            }
        }
        let ext = self
            .map
            .model_of("extLst")
            .and_then(|e| get(Some(e), "x14:conditionalFormattings"))
            .cloned();
        let ext = if truthy(self.map.model_of("extLst")) {
            ext
        } else {
            None
        };
        let cfs = merge_conditional_formattings(self.take("conditionalFormatting"), ext);

        let mut m = Map::new();
        let put = |m: &mut Map<String, Value>, key: &str, v: Option<Value>| {
            if let Some(v) = v {
                m.insert(key.into(), v);
            }
        };
        put(&mut m, "dimensions", self.take("dimension"));
        put(&mut m, "cols", self.take("cols"));
        put(&mut m, "rows", self.take("sheetData"));
        put(&mut m, "mergeCells", self.take("mergeCells"));
        put(&mut m, "hyperlinks", self.take("hyperlinks"));
        put(&mut m, "dataValidations", self.take("dataValidations"));
        m.insert("properties".into(), Value::Object(properties));
        put(&mut m, "views", self.take("sheetViews"));
        m.insert("pageSetup".into(), Value::Object(page_setup));
        put(&mut m, "headerFooter", self.take("headerFooter"));
        put(&mut m, "background", self.take("picture"));
        put(&mut m, "drawing", self.take("drawing"));
        put(&mut m, "tables", self.take("tableParts"));
        put(&mut m, "conditionalFormattings", cfs);
        if let Some(af) = self
            .map
            .model_of("autoFilter")
            .filter(|v| truthy(Some(v)))
            .cloned()
        {
            m.insert("autoFilter".into(), af);
        }
        if let Some(sp) = self
            .map
            .model_of("sheetProtection")
            .filter(|v| truthy(Some(v)))
            .cloned()
        {
            m.insert("sheetProtection".into(), sp);
        }
        self.model = Some(Value::Object(m));
        Ok(false)
    }
    crate::model_accessors!();
}
