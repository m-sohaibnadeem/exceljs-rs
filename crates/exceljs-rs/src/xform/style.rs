//! ExcelJS `lib/xlsx/xform/style/*.js` (parsing and rendering) and
//! `lib/xlsx/defaultnumformats.js`.

use serde_json::{json, Map, Value};

use super::base::{BoxX, Children, ListXform, StaticNode, StaticXform};
use super::simple::{BooleanXform, SimpleOptions, ValueXform};
use super::{XResult, Xform};
use crate::attrs;
use crate::js::parse_int;
use crate::jsv::{as_f64, defined, get, num, to_js_string, truthy};
use crate::xml::Tag;
use crate::xml_stream::XmlStream;

/// `defaultnumformats.js` entries that have a format code (`dnf.f`).
pub(crate) const DEFAULT_NUM_FORMATS: [(u32, &str); 28] = [
    (0, "General"),
    (1, "0"),
    (2, "0.00"),
    (3, "#,##0"),
    (4, "#,##0.00"),
    (9, "0%"),
    (10, "0.00%"),
    (11, "0.00E+00"),
    (12, "# ?/?"),
    (13, "# ??/??"),
    (14, "mm-dd-yy"),
    (15, "d-mmm-yy"),
    (16, "d-mmm"),
    (17, "mmm-yy"),
    (18, "h:mm AM/PM"),
    (19, "h:mm:ss AM/PM"),
    (20, "h:mm"),
    (21, "h:mm:ss"),
    (22, "m/d/yy \"h\":mm"),
    (37, "#,##0 ;(#,##0)"),
    (38, "#,##0 ;[Red](#,##0)"),
    (39, "#,##0.00 ;(#,##0.00)"),
    (40, "#,##0.00 ;[Red](#,##0.00)"),
    (45, "mm:ss"),
    (46, "[h]:mm:ss"),
    (47, "mmss.0"),
    (48, "##0.0E+0"),
    (49, "@"),
];

/// `NumFmtXform.getDefaultFmtId(formatCode)`.
pub(crate) fn default_fmt_id(code: &str) -> Option<u32> {
    DEFAULT_NUM_FORMATS
        .iter()
        .find(|(_, f)| *f == code)
        .map(|(id, _)| *id)
}

/// `NumFmtXform.getDefaultFmtCode(numFmtId)`.
pub(crate) fn default_fmt_code(id: f64) -> Option<&'static str> {
    DEFAULT_NUM_FORMATS
        .iter()
        .find(|(i, _)| f64::from(*i) == id)
        .map(|(_, f)| *f)
}

/// ExcelJS `utils.parseBoolean` for an attribute string.
pub(crate) fn parse_boolean(v: Option<&str>) -> bool {
    matches!(v, Some("true" | "1"))
}

/// ExcelJS `utils.validInt(value)`: `parseInt(value, 10)`, NaN -> 0.
pub(crate) fn valid_int(v: Option<&Value>) -> f64 {
    let s = match v {
        None => return 0.0,
        Some(v) => to_js_string(Some(v)),
    };
    parse_int(&s).unwrap_or(0.0)
}

fn obj() -> Value {
    Value::Object(Map::new())
}

fn set(model: &mut Option<Value>, key: &str, value: Value) -> XResult<()> {
    match model {
        Some(Value::Object(o)) => {
            o.insert(key.to_owned(), value);
            Ok(())
        }
        other => Err(format!(
            "Cannot set properties of {} (setting '{key}')",
            if other.is_some() { "null" } else { "undefined" }
        )),
    }
}

// ---------------------------------------------------------------------------
// ColorXform

#[derive(Debug)]
pub(crate) struct ColorXform {
    pub name: &'static str,
    model: Option<Value>,
}

impl ColorXform {
    pub fn new(name: &'static str) -> Self {
        ColorXform { name, model: None }
    }
}

impl Xform for ColorXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if node.name != self.name {
            return Ok(false);
        }
        let nonempty = |a: &str| node.attr(a).filter(|s| !s.is_empty());
        self.model = if let Some(rgb) = nonempty("rgb") {
            Some(json!({"argb": rgb}))
        } else if let Some(theme) = nonempty("theme") {
            let mut m = Map::new();
            m.insert("theme".into(), num(parse_int(theme).unwrap_or(f64::NAN)));
            if let Some(tint) = nonempty("tint") {
                m.insert("tint".into(), num(crate::js::parse_float(tint)));
            }
            Some(Value::Object(m))
        } else {
            nonempty("indexed").map(|i| json!({"indexed": num(parse_int(i).unwrap_or(f64::NAN))}))
        };
        Ok(true)
    }
    fn parse_close(&mut self, _name: &str) -> XResult<bool> {
        Ok(false)
    }
    crate::model_accessors!();
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        if !truthy(model) {
            return;
        }
        xml.open_node(self.name, &[]);
        if truthy(get(model, "argb")) {
            xml.add_attribute("rgb", get(model, "argb"));
        } else if defined(get(model, "theme")) {
            xml.add_attribute("theme", get(model, "theme"));
            if defined(get(model, "tint")) {
                xml.add_attribute("tint", get(model, "tint"));
            }
        } else if defined(get(model, "indexed")) {
            xml.add_attribute("indexed", get(model, "indexed"));
        } else {
            xml.add_attribute("auto", Some("1"));
        }
        xml.close_node();
    }
    fn reset(&mut self) {
        self.model = Some(Value::Null);
    }
}

// ---------------------------------------------------------------------------
// UnderlineXform

#[derive(Debug, Default)]
pub(crate) struct UnderlineXform {
    model: Option<Value>,
}

impl Xform for UnderlineXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if node.name == "u" {
            self.model = Some(match node.attr("val").filter(|v| !v.is_empty()) {
                Some(v) => Value::String(v.to_owned()),
                None => Value::Bool(true),
            });
        }
        Ok(false)
    }
    fn parse_close(&mut self, _name: &str) -> XResult<bool> {
        Ok(false)
    }
    crate::model_accessors!();
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        // model = model || this.model
        let model = if truthy(model) {
            model
        } else {
            self.model.as_ref()
        };
        if model == Some(&Value::Bool(true)) {
            xml.leaf_node("u", &[], None);
            return;
        }
        let attrs = match to_js_string(model).as_str() {
            "single" => Some(None),
            "double" => Some(Some("double")),
            "singleAccounting" => Some(Some("singleAccounting")),
            "doubleAccounting" => Some(Some("doubleAccounting")),
            _ => None,
        };
        if let Some(val) = attrs {
            xml.leaf_node("u", &attrs![("val", val)], None);
        }
    }
    fn reset(&mut self) {
        self.model = Some(Value::Null);
    }
}

// ---------------------------------------------------------------------------
// FontXform

/// `FontXform` options: `{tagName, fontNameTag}`.
#[derive(Debug, Clone, Copy)]
pub(crate) struct FontOptions {
    pub tag_name: &'static str,
    pub font_name_tag: &'static str,
}

pub(crate) const FONT_OPTIONS: FontOptions = FontOptions {
    tag_name: "font",
    font_name_tag: "name",
};
/// `RichTextXform.FONT_OPTIONS`.
pub(crate) const RICH_TEXT_FONT_OPTIONS: FontOptions = FontOptions {
    tag_name: "rPr",
    font_name_tag: "rFont",
};

#[derive(Debug)]
pub(crate) struct FontXform {
    pub options: FontOptions,
    map: Children,
    props: Vec<(&'static str, &'static str)>,
    model: Option<Value>,
}

impl FontXform {
    pub fn new(options: FontOptions) -> Self {
        let b = |tag| -> BoxX { Box::new(BooleanXform::new(tag)) };
        let i = |tag| -> BoxX { Box::new(ValueXform::integer(SimpleOptions::attr(tag, "val"))) };
        let s = |tag| -> BoxX { Box::new(ValueXform::string(SimpleOptions::attr(tag, "val"))) };
        let items: Vec<(&'static str, &'static str, BoxX)> = vec![
            ("b", "bold", b("b")),
            ("i", "italic", b("i")),
            ("u", "underline", Box::new(UnderlineXform::default())),
            ("charset", "charset", i("charset")),
            ("color", "color", Box::new(ColorXform::new("color"))),
            ("condense", "condense", b("condense")),
            ("extend", "extend", b("extend")),
            ("family", "family", i("family")),
            ("outline", "outline", b("outline")),
            ("vertAlign", "vertAlign", s("vertAlign")),
            ("scheme", "scheme", s("scheme")),
            ("shadow", "shadow", b("shadow")),
            ("strike", "strike", b("strike")),
            ("sz", "size", i("sz")),
            (options.font_name_tag, "name", s(options.font_name_tag)),
        ];
        let props = items.iter().map(|(t, p, _)| (*t, *p)).collect();
        let map = Children::new(items.into_iter().map(|(t, _, x)| (t, x)).collect());
        FontXform {
            options,
            map,
            props,
            model: None,
        }
    }

    fn prop(&self, tag: &str) -> Option<&'static str> {
        self.props.iter().find(|(t, _)| *t == tag).map(|(_, p)| *p)
    }
}

impl Xform for FontXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if self.map.forward_open(node)? {
            return Ok(true);
        }
        if self.map.activate(node.name) {
            return match self.map.parser() {
                Some(p) => p.parse_open(node),
                None => Ok(false),
            };
        }
        if node.name == self.options.tag_name {
            self.model = Some(obj());
            return Ok(true);
        }
        Ok(false)
    }
    fn parse_text(&mut self, text: &str) {
        self.map.forward_text(text);
    }
    fn parse_close(&mut self, name: &str) -> XResult<bool> {
        let mut fell_through = true;
        if let Some(p) = self.map.parser() {
            if !p.parse_close(name)? {
                let prop = self.prop(name).ok_or_else(|| {
                    "Cannot read properties of undefined (reading 'prop')".to_owned()
                })?;
                if let Some(m) = self.map.parser_model().filter(|m| truthy(Some(m))).cloned() {
                    set(&mut self.model, prop, m)?;
                }
                self.map.active = None;
                return Ok(true);
            }
            fell_through = true;
        }
        let _ = fell_through;
        Ok(name != self.options.tag_name)
    }
    crate::model_accessors!();
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        xml.open_node(self.options.tag_name, &[]);
        for (i, (_, prop)) in self.props.iter().enumerate() {
            if let Some(child) = self.map.at(i) {
                child.render(xml, get(model, prop));
            }
        }
        xml.close_node();
    }
    fn reset(&mut self) {
        self.model = Some(Value::Null);
        self.map.reset_all();
    }
}

// ---------------------------------------------------------------------------
// Fills

#[derive(Debug)]
struct StopXform {
    color: ColorXform,
    in_color: bool,
    model: Option<Value>,
}

impl StopXform {
    fn new() -> Self {
        StopXform {
            color: ColorXform::new("color"),
            in_color: false,
            model: None,
        }
    }
}

impl Xform for StopXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if self.in_color {
            self.color.parse_open(node)?;
            return Ok(true);
        }
        match node.name {
            "stop" => {
                self.model =
                    Some(json!({"position": super::simple::float_value(node.attr("position"))}));
                Ok(true)
            }
            "color" => {
                self.in_color = true;
                self.color.parse_open(node)?;
                Ok(true)
            }
            _ => Ok(false),
        }
    }
    fn parse_close(&mut self, name: &str) -> XResult<bool> {
        if self.in_color {
            if !self.color.parse_close(name)? {
                if let Some(c) = self.color.model().cloned() {
                    set(&mut self.model, "color", c)?;
                }
                self.in_color = false;
            }
            return Ok(true);
        }
        Ok(false)
    }
    crate::model_accessors!();
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        xml.open_node("stop", &[]);
        xml.add_attribute("position", get(model, "position"));
        self.color.render(xml, get(model, "color"));
        xml.close_node();
    }
    fn reset(&mut self) {
        self.model = Some(Value::Null);
        self.color.reset();
    }
}

#[derive(Debug)]
struct PatternFillXform {
    map: Children,
    model: Option<Value>,
}

impl PatternFillXform {
    fn new() -> Self {
        PatternFillXform {
            map: Children::new(vec![
                ("fgColor", Box::new(ColorXform::new("fgColor"))),
                ("bgColor", Box::new(ColorXform::new("bgColor"))),
            ]),
            model: None,
        }
    }
}

impl Xform for PatternFillXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if self.map.forward_open(node)? {
            return Ok(true);
        }
        if node.name == "patternFill" {
            let mut m = Map::new();
            m.insert("type".into(), json!("pattern"));
            if let Some(p) = node.attr("patternType") {
                m.insert("pattern".into(), json!(p));
            }
            self.model = Some(Value::Object(m));
            return Ok(true);
        }
        if self.map.activate(node.name) {
            if let Some(p) = self.map.parser() {
                p.parse_open(node)?;
            }
            return Ok(true);
        }
        Ok(false)
    }
    fn parse_text(&mut self, text: &str) {
        self.map.forward_text(text);
    }
    fn parse_close(&mut self, name: &str) -> XResult<bool> {
        if let Some(p) = self.map.parser() {
            if !p.parse_close(name)? {
                if let Some(m) = self.map.parser_model().filter(|m| truthy(Some(m))).cloned() {
                    set(&mut self.model, name, m)?;
                }
                self.map.active = None;
            }
            return Ok(true);
        }
        Ok(false)
    }
    crate::model_accessors!();
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        xml.open_node("patternFill", &[]);
        xml.add_attribute("patternType", get(model, "pattern"));
        for name in ["fgColor", "bgColor"] {
            if truthy(get(model, name)) {
                if let Some(x) = self.map.get(name) {
                    x.render(xml, get(model, name));
                }
            }
        }
        xml.close_node();
    }
    fn reset(&mut self) {
        self.model = Some(Value::Null);
        self.map.reset_all();
    }
}

#[derive(Debug)]
struct GradientFillXform {
    stop: StopXform,
    in_stop: bool,
    model: Option<Value>,
}

impl GradientFillXform {
    fn new() -> Self {
        GradientFillXform {
            stop: StopXform::new(),
            in_stop: false,
            model: None,
        }
    }
}

impl Xform for GradientFillXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if self.in_stop {
            self.stop.parse_open(node)?;
            return Ok(true);
        }
        match node.name {
            "gradientFill" => {
                let mut m = Map::new();
                m.insert("stops".into(), Value::Array(Vec::new()));
                let attr = |a: &str| node.attr(a);
                let truthy_attr = |a: &str| attr(a).is_some_and(|s| !s.is_empty());
                let float_or_zero = |a: &str| {
                    if truthy_attr(a) {
                        num(crate::js::parse_float(attr(a).unwrap_or("")))
                    } else {
                        num(0.0)
                    }
                };
                if truthy_attr("degree") {
                    m.insert("gradient".into(), json!("angle"));
                    m.insert("degree".into(), super::simple::int_value(attr("degree")));
                } else if attr("type") == Some("path") {
                    m.insert("gradient".into(), json!("path"));
                    let mut center = Map::new();
                    center.insert("left".into(), float_or_zero("left"));
                    center.insert("top".into(), float_or_zero("top"));
                    if attr("right") != attr("left") {
                        center.insert("right".into(), float_or_zero("right"));
                    }
                    if attr("bottom") != attr("top") {
                        center.insert("bottom".into(), float_or_zero("bottom"));
                    }
                    m.insert("center".into(), Value::Object(center));
                }
                self.model = Some(Value::Object(m));
                Ok(true)
            }
            "stop" => {
                self.in_stop = true;
                self.stop.parse_open(node)?;
                Ok(true)
            }
            _ => Ok(false),
        }
    }
    fn parse_text(&mut self, text: &str) {
        if self.in_stop {
            self.stop.parse_text(text);
        }
    }
    fn parse_close(&mut self, name: &str) -> XResult<bool> {
        if self.in_stop {
            if !self.stop.parse_close(name)? {
                let item = self.stop.model().cloned().unwrap_or(Value::Null);
                match self
                    .model
                    .as_mut()
                    .and_then(|m| m.get_mut("stops"))
                    .and_then(Value::as_array_mut)
                {
                    Some(stops) => stops.push(item),
                    None => {
                        return Err("Cannot read properties of undefined (reading 'stops')".into())
                    }
                }
                self.in_stop = false;
            }
            return Ok(true);
        }
        Ok(false)
    }
    crate::model_accessors!();
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        xml.open_node("gradientFill", &[]);
        match crate::jsv::as_str(get(model, "gradient")) {
            Some("angle") => xml.add_attribute("degree", get(model, "degree")),
            Some("path") => {
                xml.add_attribute("type", Some("path"));
                let center = get(model, "center");
                let c = |k: &str| get(center, k);
                if truthy(c("left")) {
                    xml.add_attribute("left", c("left"));
                    if !defined(c("right")) {
                        xml.add_attribute("right", c("left"));
                    }
                }
                if truthy(c("right")) {
                    xml.add_attribute("right", c("right"));
                }
                if truthy(c("top")) {
                    xml.add_attribute("top", c("top"));
                    if !defined(c("bottom")) {
                        xml.add_attribute("bottom", c("top"));
                    }
                }
                if truthy(c("bottom")) {
                    xml.add_attribute("bottom", c("bottom"));
                }
            }
            _ => {}
        }
        for stop in get(model, "stops")
            .and_then(Value::as_array)
            .into_iter()
            .flatten()
        {
            self.stop.render(xml, Some(stop));
        }
        xml.close_node();
    }
    fn reset(&mut self) {
        self.model = Some(Value::Null);
        self.stop.reset();
    }
}

#[derive(Debug)]
pub(crate) struct FillXform {
    map: Children,
    model: Option<Value>,
}

impl FillXform {
    pub fn new() -> Self {
        FillXform {
            map: Children::new(vec![
                ("patternFill", Box::new(PatternFillXform::new())),
                ("gradientFill", Box::new(GradientFillXform::new())),
            ]),
            model: None,
        }
    }
}

impl Xform for FillXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if self.map.forward_open(node)? {
            return Ok(true);
        }
        if node.name == "fill" {
            self.model = Some(obj());
            return Ok(true);
        }
        if self.map.activate(node.name) {
            if let Some(p) = self.map.parser() {
                p.parse_open(node)?;
            }
            return Ok(true);
        }
        Ok(false)
    }
    fn parse_text(&mut self, text: &str) {
        self.map.forward_text(text);
    }
    fn parse_close(&mut self, name: &str) -> XResult<bool> {
        if let Some(p) = self.map.parser() {
            if !p.parse_close(name)? {
                let kind = if self.map.active_name() == Some("patternFill") {
                    "pattern"
                } else {
                    "gradient"
                };
                let mut m = self.map.parser_model().cloned();
                set(&mut m, "type", json!(kind))?;
                self.model = m;
                self.map.active = None;
            }
            return Ok(true);
        }
        Ok(false)
    }
    crate::model_accessors!();
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        xml.add_rollback();
        xml.open_node("fill", &[]);
        let child = match crate::jsv::as_str(get(model, "type")) {
            Some("pattern") => self.map.get("patternFill"),
            Some("gradient") => self.map.get("gradientFill"),
            _ => None,
        };
        match child {
            Some(c) => {
                c.render(xml, model);
                xml.close_node();
                xml.commit();
            }
            None => xml.rollback(),
        }
    }
    fn reset(&mut self) {
        self.model = Some(Value::Null);
        self.map.reset_all();
    }
}

// ---------------------------------------------------------------------------
// Borders

#[derive(Debug)]
struct EdgeXform {
    name: &'static str,
    color: ColorXform,
    in_color: bool,
    model: Option<Value>,
}

impl EdgeXform {
    fn new(name: &'static str) -> Self {
        EdgeXform {
            name,
            color: ColorXform::new("color"),
            in_color: false,
            model: None,
        }
    }

    fn render_edge(
        &self,
        xml: &mut XmlStream,
        model: Option<&Value>,
        default_color: Option<&Value>,
    ) {
        let color = if truthy(get(model, "color")) {
            get(model, "color")
        } else if truthy(default_color) {
            default_color
        } else {
            None
        };
        xml.open_node(self.name, &[]);
        if truthy(model) && truthy(get(model, "style")) {
            xml.add_attribute("style", get(model, "style"));
            if truthy(color) {
                self.color.render(xml, color);
            }
        }
        xml.close_node();
    }
}

impl Xform for EdgeXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if self.in_color {
            self.color.parse_open(node)?;
            return Ok(true);
        }
        if node.name == self.name {
            self.model = node
                .attr("style")
                .filter(|s| !s.is_empty())
                .map(|s| json!({"style": s}));
            return Ok(true);
        }
        if node.name == "color" {
            self.in_color = true;
            self.color.parse_open(node)?;
            return Ok(true);
        }
        Ok(false)
    }
    fn parse_close(&mut self, name: &str) -> XResult<bool> {
        if self.in_color {
            if !self.color.parse_close(name)? {
                self.in_color = false;
            }
            return Ok(true);
        }
        if name == self.name {
            if let Some(c) = self.color.model().filter(|c| truthy(Some(c))).cloned() {
                if !truthy(self.model.as_ref()) {
                    self.model = Some(obj());
                }
                set(&mut self.model, "color", c)?;
            }
        }
        Ok(false)
    }
    crate::model_accessors!();
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        self.render_edge(xml, model, None);
    }
    fn reset(&mut self) {
        self.model = Some(Value::Null);
        self.color.reset();
    }
}

const EDGES: [&str; 5] = ["top", "left", "bottom", "right", "diagonal"];

#[derive(Debug)]
pub(crate) struct BorderXform {
    edges: Vec<EdgeXform>,
    active: Option<usize>,
    diagonal_up: bool,
    diagonal_down: bool,
    model: Option<Value>,
}

impl BorderXform {
    pub fn new() -> Self {
        BorderXform {
            edges: EDGES.iter().map(|n| EdgeXform::new(n)).collect(),
            active: None,
            diagonal_up: false,
            diagonal_down: false,
            model: None,
        }
    }

    fn edge(&self, name: &str) -> Option<&EdgeXform> {
        self.edges.iter().find(|e| e.name == name)
    }
}

impl Xform for BorderXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if let Some(e) = self.active.and_then(|i| self.edges.get_mut(i)) {
            e.parse_open(node)?;
            return Ok(true);
        }
        if node.name == "border" {
            self.reset();
            self.diagonal_up = parse_boolean(node.attr("diagonalUp"));
            self.diagonal_down = parse_boolean(node.attr("diagonalDown"));
            return Ok(true);
        }
        self.active = self.edges.iter().position(|e| e.name == node.name);
        match self.active.and_then(|i| self.edges.get_mut(i)) {
            Some(e) => {
                e.parse_open(node)?;
                Ok(true)
            }
            None => Ok(false),
        }
    }
    fn parse_text(&mut self, text: &str) {
        if let Some(e) = self.active.and_then(|i| self.edges.get_mut(i)) {
            e.parse_text(text);
        }
    }
    fn parse_close(&mut self, name: &str) -> XResult<bool> {
        if let Some(e) = self.active.and_then(|i| self.edges.get_mut(i)) {
            if !e.parse_close(name)? {
                self.active = None;
            }
            return Ok(true);
        }
        if name == "border" {
            let mut m = Map::new();
            for key in ["left", "right", "top", "bottom", "diagonal"] {
                if let Some(edge) = self
                    .edge(key)
                    .and_then(|e| e.model())
                    .filter(|v| truthy(Some(v)))
                {
                    let mut edge = edge.clone();
                    if key == "diagonal" {
                        if let Value::Object(o) = &mut edge {
                            o.insert("up".into(), Value::Bool(self.diagonal_up));
                            o.insert("down".into(), Value::Bool(self.diagonal_down));
                        }
                    }
                    m.insert(key.into(), edge);
                }
            }
            self.model = Some(Value::Object(m));
        }
        Ok(false)
    }
    crate::model_accessors!();
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        let color = get(model, "color");
        xml.open_node("border", &[]);
        let diagonal = get(model, "diagonal");
        if truthy(diagonal) && truthy(get(diagonal, "style")) {
            if truthy(get(diagonal, "up")) {
                xml.add_attribute("diagonalUp", Some("1"));
            }
            if truthy(get(diagonal, "down")) {
                xml.add_attribute("diagonalDown", Some("1"));
            }
        }
        for key in ["left", "right", "top", "bottom", "diagonal"] {
            let edge_model = get(model, key);
            let patched;
            let edge_model =
                if truthy(edge_model) && !truthy(get(edge_model, "color")) && truthy(color) {
                    let mut o = match edge_model {
                        Some(Value::Object(o)) => o.clone(),
                        _ => Map::new(),
                    };
                    o.insert("color".into(), color.cloned().unwrap_or(Value::Null));
                    patched = Value::Object(o);
                    Some(&patched)
                } else {
                    edge_model
                };
            if let Some(x) = self.edge(key) {
                x.render_edge(xml, edge_model, color);
            }
        }
        xml.close_node();
    }
    fn reset(&mut self) {
        self.model = Some(Value::Null);
        for e in &mut self.edges {
            e.reset();
        }
    }
}

// ---------------------------------------------------------------------------
// AlignmentXform

fn valid_horizontal(v: Option<&Value>) -> Option<String> {
    let s = to_js_string(v);
    matches!(
        s.as_str(),
        "left" | "center" | "right" | "fill" | "centerContinuous" | "distributed" | "justify"
    )
    .then_some(s)
}

fn valid_vertical(v: Option<&Value>) -> Option<String> {
    let s = to_js_string(v);
    if s == "middle" {
        return Some("center".into());
    }
    matches!(s.as_str(), "top" | "bottom" | "distributed" | "justify").then_some(s)
}

/// `validation.textRotation` then `textRotationXform.toXml`.
fn text_rotation_to_xml(v: Option<&Value>) -> Option<f64> {
    if crate::jsv::as_str(v) == Some("vertical") {
        return Some(255.0);
    }
    let value = valid_int(v);
    if !(-90.0..=90.0).contains(&value) || value == 0.0 {
        return None;
    }
    let tr = crate::model_json::js_round(value);
    if (0.0..=90.0).contains(&tr) {
        Some(tr)
    } else if (-90.0..0.0).contains(&tr) {
        Some(90.0 - tr)
    } else {
        None
    }
}

fn text_rotation_to_model(attr: Option<&str>) -> Option<Value> {
    let tr = attr.and_then(parse_int).unwrap_or(0.0);
    if tr == 255.0 {
        return Some(json!("vertical"));
    }
    if (0.0..=90.0).contains(&tr) {
        return Some(num(tr));
    }
    if tr > 90.0 && tr <= 180.0 {
        return Some(num(90.0 - tr));
    }
    None
}

#[derive(Debug, Default)]
pub(crate) struct AlignmentXform {
    model: Option<Value>,
}

impl Xform for AlignmentXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        let mut m = Map::new();
        let mut valid = false;
        let truthy_attr = |a: &str| node.attr(a).is_some_and(|s| !s.is_empty());
        let mut add = |cond: bool, name: &str, value: Option<Value>| {
            if cond {
                if let Some(v) = value {
                    m.insert(name.to_owned(), v);
                }
                valid = true;
            }
        };
        add(
            truthy_attr("horizontal"),
            "horizontal",
            node.attr("horizontal").map(|s| json!(s)),
        );
        add(
            truthy_attr("vertical"),
            "vertical",
            node.attr("vertical")
                .map(|s| json!(if s == "center" { "middle" } else { s })),
        );
        add(
            truthy_attr("wrapText"),
            "wrapText",
            Some(json!(parse_boolean(node.attr("wrapText")))),
        );
        add(
            truthy_attr("shrinkToFit"),
            "shrinkToFit",
            Some(json!(parse_boolean(node.attr("shrinkToFit")))),
        );
        add(
            truthy_attr("indent"),
            "indent",
            Some(super::simple::int_value(node.attr("indent"))),
        );
        add(
            truthy_attr("textRotation"),
            "textRotation",
            text_rotation_to_model(node.attr("textRotation")),
        );
        add(
            truthy_attr("readingOrder"),
            "readingOrder",
            Some(json!(if node.attr("readingOrder") == Some("2") {
                "rtl"
            } else {
                "ltr"
            })),
        );
        self.model = Some(if valid { Value::Object(m) } else { Value::Null });
        Ok(false)
    }
    fn parse_close(&mut self, _name: &str) -> XResult<bool> {
        Ok(false)
    }
    crate::model_accessors!();
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        xml.add_rollback();
        xml.open_node("alignment", &[]);
        let mut valid = false;
        let mut add = |xml: &mut XmlStream, name: &str, value: Option<String>| {
            if let Some(v) = value.filter(|v| !v.is_empty()) {
                xml.add_attribute(name, Some(v));
                valid = true;
            }
        };
        add(
            xml,
            "horizontal",
            valid_horizontal(get(model, "horizontal")),
        );
        add(xml, "vertical", valid_vertical(get(model, "vertical")));
        add(
            xml,
            "wrapText",
            truthy(get(model, "wrapText")).then(|| "1".to_owned()),
        );
        add(
            xml,
            "shrinkToFit",
            truthy(get(model, "shrinkToFit")).then(|| "1".to_owned()),
        );
        let indent = valid_int(get(model, "indent")).max(0.0);
        add(
            xml,
            "indent",
            (indent != 0.0).then(|| crate::jsv::number_to_string(indent)),
        );
        add(
            xml,
            "textRotation",
            text_rotation_to_xml(get(model, "textRotation")).map(crate::jsv::number_to_string),
        );
        let ro = match crate::jsv::as_str(get(model, "readingOrder")) {
            Some("ltr") => Some("1".to_owned()),
            Some("rtl") => Some("2".to_owned()),
            _ => None,
        };
        add(xml, "readingOrder", ro);
        xml.close_node();
        if valid {
            xml.commit();
        } else {
            xml.rollback();
        }
    }
    fn reset(&mut self) {
        self.model = Some(Value::Null);
    }
}

// ---------------------------------------------------------------------------
// ProtectionXform

#[derive(Debug, Default)]
pub(crate) struct ProtectionXform {
    model: Option<Value>,
}

impl Xform for ProtectionXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        let locked = node.attr("locked") != Some("0");
        let hidden = node.attr("hidden") == Some("1");
        self.model = Some(if !locked || hidden {
            json!({"locked": locked, "hidden": hidden})
        } else {
            Value::Null
        });
        Ok(false)
    }
    fn parse_close(&mut self, _name: &str) -> XResult<bool> {
        Ok(false)
    }
    crate::model_accessors!();
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        xml.add_rollback();
        xml.open_node("protection", &[]);
        let mut valid = false;
        let locked = get(model, "locked").is_none_or(|v| truthy(Some(v)));
        if !locked {
            xml.add_attribute("locked", Some("0"));
            valid = true;
        }
        let hidden = get(model, "hidden").is_some_and(|v| truthy(Some(v)));
        if hidden {
            xml.add_attribute("hidden", Some("1"));
            valid = true;
        }
        xml.close_node();
        if valid {
            xml.commit();
        } else {
            xml.rollback();
        }
    }
    fn reset(&mut self) {
        self.model = Some(Value::Null);
    }
}

// ---------------------------------------------------------------------------
// NumFmtXform

/// `formatCode.replace(/[\\](.)/g, '$1')`.
fn unescape_format_code(code: &str) -> String {
    let mut out = String::with_capacity(code.len());
    let mut chars = code.chars().peekable();
    while let Some(c) = chars.next() {
        if c == '\\' {
            match chars.peek() {
                Some(&next) if !matches!(next, '\n' | '\r' | '\u{2028}' | '\u{2029}') => {
                    out.push(next);
                    chars.next();
                }
                _ => out.push(c),
            }
        } else {
            out.push(c);
        }
    }
    out
}

#[derive(Debug, Default)]
pub(crate) struct NumFmtXform {
    model: Option<Value>,
}

impl Xform for NumFmtXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if node.name != "numFmt" {
            return Ok(false);
        }
        let code = node
            .attr("formatCode")
            .ok_or_else(|| "Cannot read properties of undefined (reading 'replace')".to_owned())?;
        self.model = Some(json!({
            "id": super::simple::int_value(node.attr("numFmtId")),
            "formatCode": unescape_format_code(code),
        }));
        Ok(true)
    }
    fn parse_close(&mut self, _name: &str) -> XResult<bool> {
        Ok(false)
    }
    crate::model_accessors!();
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        xml.leaf_node(
            "numFmt",
            &attrs![
                ("numFmtId", get(model, "id")),
                ("formatCode", get(model, "formatCode"))
            ],
            None,
        );
    }
    fn reset(&mut self) {
        self.model = Some(Value::Null);
    }
}

// ---------------------------------------------------------------------------
// StyleXform (<xf>)

#[derive(Debug)]
pub(crate) struct StyleXform {
    xf_id: bool,
    alignment: AlignmentXform,
    protection: ProtectionXform,
    active: Option<bool>, // Some(true) = protection, Some(false) = alignment
    model: Option<Value>,
}

impl StyleXform {
    pub fn new(xf_id: bool) -> Self {
        StyleXform {
            xf_id,
            alignment: AlignmentXform::default(),
            protection: ProtectionXform::default(),
            active: None,
            model: None,
        }
    }
}

fn or_zero(v: Option<&Value>) -> Value {
    if truthy(v) {
        v.cloned().unwrap_or(Value::Null)
    } else {
        json!(0)
    }
}

impl Xform for StyleXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        match self.active {
            Some(true) => {
                self.protection.parse_open(node)?;
                return Ok(true);
            }
            Some(false) => {
                self.alignment.parse_open(node)?;
                return Ok(true);
            }
            None => {}
        }
        match node.name {
            "xf" => {
                let int = |a| super::simple::int_value(node.attr(a));
                let mut m = Map::new();
                m.insert("numFmtId".into(), int("numFmtId"));
                m.insert("fontId".into(), int("fontId"));
                m.insert("fillId".into(), int("fillId"));
                m.insert("borderId".into(), int("borderId"));
                if self.xf_id {
                    m.insert("xfId".into(), int("xfId"));
                }
                self.model = Some(Value::Object(m));
                Ok(true)
            }
            "alignment" => {
                self.active = Some(false);
                self.alignment.parse_open(node)?;
                Ok(true)
            }
            "protection" => {
                self.active = Some(true);
                self.protection.parse_open(node)?;
                Ok(true)
            }
            _ => Ok(false),
        }
    }
    fn parse_close(&mut self, name: &str) -> XResult<bool> {
        if let Some(is_protection) = self.active {
            let done = if is_protection {
                !self.protection.parse_close(name)?
            } else {
                !self.alignment.parse_close(name)?
            };
            if done {
                let (key, m) = if is_protection {
                    ("protection", self.protection.model().cloned())
                } else {
                    ("alignment", self.alignment.model().cloned())
                };
                if let Some(m) = m {
                    set(&mut self.model, key, m)?;
                }
                self.active = None;
            }
            return Ok(true);
        }
        Ok(name != "xf")
    }
    crate::model_accessors!();
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        xml.open_node(
            "xf",
            &attrs![
                ("numFmtId", Some(or_zero(get(model, "numFmtId")))),
                ("fontId", Some(or_zero(get(model, "fontId")))),
                ("fillId", Some(or_zero(get(model, "fillId")))),
                ("borderId", Some(or_zero(get(model, "borderId")))),
            ],
        );
        if self.xf_id {
            xml.add_attribute("xfId", Some(or_zero(get(model, "xfId"))));
        }
        for (key, attr) in [
            ("numFmtId", "applyNumberFormat"),
            ("fontId", "applyFont"),
            ("fillId", "applyFill"),
            ("borderId", "applyBorder"),
            ("alignment", "applyAlignment"),
            ("protection", "applyProtection"),
        ] {
            if truthy(get(model, key)) {
                xml.add_attribute(attr, Some("1"));
            }
        }
        if truthy(get(model, "alignment")) {
            self.alignment.render(xml, get(model, "alignment"));
        }
        if truthy(get(model, "protection")) {
            self.protection.render(xml, get(model, "protection"));
        }
        xml.close_node();
    }
    fn reset(&mut self) {
        self.model = Some(Value::Null);
        self.alignment.reset();
        self.protection.reset();
    }
}

// ---------------------------------------------------------------------------
// DxfXform

#[derive(Debug)]
pub(crate) struct DxfXform {
    map: Children,
    model: Option<Value>,
}

impl DxfXform {
    pub fn new() -> Self {
        DxfXform {
            map: Children::new(vec![
                ("alignment", Box::new(AlignmentXform::default())),
                ("border", Box::new(BorderXform::new())),
                ("fill", Box::new(FillXform::new())),
                ("font", Box::new(FontXform::new(FONT_OPTIONS))),
                ("numFmt", Box::new(NumFmtXform::default())),
                ("protection", Box::new(ProtectionXform::default())),
            ]),
            model: None,
        }
    }
}

impl Xform for DxfXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if self.map.forward_open(node)? {
            return Ok(true);
        }
        if node.name == "dxf" {
            self.reset();
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
        if name == "dxf" {
            let mut m = Map::new();
            for key in [
                "alignment",
                "border",
                "fill",
                "font",
                "numFmt",
                "protection",
            ] {
                if let Some(v) = self.map.model_of(key) {
                    m.insert(key.into(), v.clone());
                }
            }
            self.model = Some(Value::Object(m));
            return Ok(false);
        }
        Ok(true)
    }
    crate::model_accessors!();
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        xml.open_node("dxf", &[]);
        if truthy(get(model, "font")) {
            if let Some(x) = self.map.get("font") {
                x.render(xml, get(model, "font"));
            }
        }
        if truthy(get(model, "numFmt")) && truthy(get(model, "numFmtId")) {
            let nf = json!({"id": get(model, "numFmtId"), "formatCode": get(model, "numFmt")});
            if let Some(x) = self.map.get("numFmt") {
                x.render(xml, Some(&nf));
            }
        }
        for key in ["fill", "alignment", "border", "protection"] {
            if truthy(get(model, key)) {
                if let Some(x) = self.map.get(key) {
                    x.render(xml, get(model, key));
                }
            }
        }
        xml.close_node();
    }
    fn reset(&mut self) {
        self.model = Some(Value::Null);
        self.map.reset_all();
    }
}

// ---------------------------------------------------------------------------
// StylesXform: parsing styles.xml

pub(crate) fn cell_styles_static() -> StaticXform {
    StaticXform::new(StaticNode::new(
        "cellStyles",
        attrs![("count", Some(1i64))],
        vec![StaticNode::new(
            "cellStyle",
            attrs![
                ("name", Some("Normal")),
                ("xfId", Some(0i64)),
                ("builtinId", Some(0i64))
            ],
            vec![],
        )],
    ))
}

pub(crate) fn table_styles_static() -> StaticXform {
    StaticXform::new(StaticNode::new(
        "tableStyles",
        attrs![
            ("count", Some(0i64)),
            ("defaultTableStyle", Some("TableStyleMedium2")),
            ("defaultPivotStyle", Some("PivotStyleLight16")),
        ],
        vec![],
    ))
}

pub(crate) fn ext_lst_static() -> StaticXform {
    StaticXform::new(StaticNode::new(
        "extLst",
        vec![],
        vec![
            StaticNode::new(
                "ext",
                attrs![
                    ("uri", Some("{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}")),
                    (
                        "xmlns:x14",
                        Some("http://schemas.microsoft.com/office/spreadsheetml/2009/9/main")
                    ),
                ],
                vec![StaticNode::new(
                    "x14:slicerStyles",
                    attrs![("defaultSlicerStyle", Some("SlicerStyleLight1"))],
                    vec![],
                )],
            ),
            StaticNode::new(
                "ext",
                attrs![
                    ("uri", Some("{9260A510-F301-46a8-8635-F512D64BE5F5}")),
                    (
                        "xmlns:x15",
                        Some("http://schemas.microsoft.com/office/spreadsheetml/2010/11/main")
                    ),
                ],
                vec![StaticNode::new(
                    "x15:timelineStyles",
                    attrs![("defaultTimelineStyle", Some("TimeSlicerStyleLight1"))],
                    vec![],
                )],
            ),
        ],
    ))
}

/// Parses `xl/styles.xml` (StylesXform's parse role).
pub(crate) struct StylesParser {
    map: Children,
    model: Option<Value>,
}

impl std::fmt::Debug for StylesParser {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str("StylesParser")
    }
}

impl StylesParser {
    pub fn new() -> Self {
        StylesParser {
            map: Children::new(vec![
                (
                    "numFmts",
                    Box::new(ListXform::new("numFmts", Box::new(NumFmtXform::default())).count()),
                ),
                (
                    "fonts",
                    Box::new(
                        ListXform::new("fonts", Box::new(FontXform::new(FONT_OPTIONS)))
                            .count()
                            .attrs(attrs![("x14ac:knownFonts", Some(1i64))]),
                    ),
                ),
                (
                    "fills",
                    Box::new(ListXform::new("fills", Box::new(FillXform::new())).count()),
                ),
                (
                    "borders",
                    Box::new(ListXform::new("borders", Box::new(BorderXform::new())).count()),
                ),
                (
                    "cellStyleXfs",
                    Box::new(
                        ListXform::new("cellStyleXfs", Box::new(StyleXform::new(false))).count(),
                    ),
                ),
                (
                    "cellXfs",
                    Box::new(ListXform::new("cellXfs", Box::new(StyleXform::new(true))).count()),
                ),
                (
                    "dxfs",
                    Box::new(
                        ListXform::new("dxfs", Box::new(DxfXform::new()))
                            .always()
                            .count(),
                    ),
                ),
                ("numFmt", Box::new(NumFmtXform::default())),
                ("font", Box::new(FontXform::new(FONT_OPTIONS))),
                ("fill", Box::new(FillXform::new())),
                ("border", Box::new(BorderXform::new())),
                ("style", Box::new(StyleXform::new(true))),
                ("cellStyles", Box::new(cell_styles_static())),
                ("tableStyles", Box::new(table_styles_static())),
                ("extLst", Box::new(ext_lst_static())),
            ]),
            model: None,
        }
    }
}

impl Xform for StylesParser {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if self.map.forward_open(node)? {
            return Ok(true);
        }
        if node.name == "styleSheet" {
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
        if name == "styleSheet" {
            let mut m = Map::new();
            for (prop, xform) in [
                ("numFmts", "numFmts"),
                ("fonts", "fonts"),
                ("fills", "fills"),
                ("borders", "borders"),
                ("styles", "cellXfs"),
                ("dxfs", "dxfs"),
            ] {
                if let Some(Value::Array(a)) = self.map.model_of(xform) {
                    if !a.is_empty() {
                        m.insert(prop.into(), Value::Array(a.clone()));
                    }
                }
            }
            self.model = Some(Value::Object(m));
            return Ok(false);
        }
        Ok(true)
    }
    crate::model_accessors!();
}

/// The parsed styles.xml model with ExcelJS's `getStyleModel` lookup.
#[derive(Debug, Default)]
pub(crate) struct Styles {
    pub model: Value,
    /// `index.numFmt`: custom number formats by id.
    num_fmts: std::collections::HashMap<String, Value>,
}

impl Styles {
    pub fn new(model: Value) -> Self {
        let mut num_fmts = std::collections::HashMap::new();
        for nf in get(Some(&model), "numFmts")
            .and_then(Value::as_array)
            .into_iter()
            .flatten()
        {
            num_fmts.insert(
                to_js_string(get(Some(nf), "id")),
                get(Some(nf), "formatCode").cloned().unwrap_or(Value::Null),
            );
        }
        Styles { model, num_fmts }
    }

    pub fn dxf(&self, id: &Value) -> Option<&Value> {
        get(get(Some(&self.model), "dxfs"), &to_js_string(Some(id)))
    }

    /// The parts of `getStyleModel(id)`, each with a part reference so the
    /// caller can share objects: returns `None` for `null` (no such style),
    /// or the list of (key, value) where part values are `Part`.
    pub fn style_parts(&self, id: f64) -> XResult<Option<Vec<(&'static str, StylePart)>>> {
        let styles = get(Some(&self.model), "styles");
        let Some(styles) = styles else {
            return Err("Cannot read properties of undefined (reading 'NaN')".into());
        };
        let key = crate::jsv::number_to_string(id);
        let style = get(Some(styles), &key);
        if !truthy(style) {
            return Ok(None);
        }
        let mut parts = Vec::new();
        let num_fmt_id = get(style, "numFmtId");
        if truthy(num_fmt_id) {
            let custom = self
                .num_fmts
                .get(&to_js_string(num_fmt_id))
                .filter(|v| truthy(Some(v)))
                .cloned();
            let code = custom.or_else(|| {
                as_f64(num_fmt_id)
                    .and_then(default_fmt_code)
                    .map(|c| json!(c))
            });
            if let Some(code) = code.filter(|c| truthy(Some(c))) {
                parts.push(("numFmt", StylePart::Value(code)));
            }
        }
        for (name, group, id_key) in [
            ("font", "fonts", "fontId"),
            ("border", "borders", "borderId"),
            ("fill", "fills", "fillId"),
        ] {
            let part_id = get(style, id_key);
            let use_it = truthy(part_id) || as_f64(part_id) == Some(0.0);
            if use_it {
                let group_v = get(Some(&self.model), group).ok_or_else(|| {
                    format!(
                        "Cannot read properties of undefined (reading '{}')",
                        to_js_string(part_id)
                    )
                })?;
                let part_key = to_js_string(part_id);
                if truthy(get(Some(group_v), &part_key)) {
                    parts.push((name, StylePart::Shared(group, part_key)));
                }
            }
        }
        for key in ["alignment", "protection"] {
            if let Some(v) = get(style, key).filter(|v| truthy(Some(v))) {
                parts.push((key, StylePart::Value(v.clone())));
            }
        }
        Ok(Some(parts))
    }

    pub fn part(&self, group: &str, key: &str) -> Option<&Value> {
        get(get(Some(&self.model), group), key)
    }
}

/// A piece of a style model: a value of its own, or a font/border/fill
/// object shared by every style that uses it.
#[derive(Debug, Clone)]
pub(crate) enum StylePart {
    Value(Value),
    Shared(&'static str, String),
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xform::{parse_part, to_xml};

    #[test]
    fn font_round_trip() {
        let mut f = FontXform::new(FONT_OPTIONS);
        let xml = r#"<font><b/><sz val="12"/><color rgb="FF0000FF"/><name val="Calibri"/><u val="double"/></font>"#;
        let m = parse_part(&mut f, "x", xml.as_bytes()).unwrap().unwrap();
        assert_eq!(
            m,
            json!({"bold": true, "size": 12, "color": {"argb": "FF0000FF"}, "name": "Calibri", "underline": "double"})
        );
        assert_eq!(
            to_xml(&f, Some(&m)),
            r#"<font><b/><u val="double"/><color rgb="FF0000FF"/><sz val="12"/><name val="Calibri"/></font>"#
        );
    }

    #[test]
    fn border_and_fill() {
        let mut b = BorderXform::new();
        let xml = r#"<border diagonalUp="1"><left style="thin"><color indexed="64"/></left><right/><diagonal style="thin"/></border>"#;
        let m = parse_part(&mut b, "x", xml.as_bytes()).unwrap().unwrap();
        assert_eq!(
            m,
            json!({"left": {"style": "thin", "color": {"indexed": 64}}, "diagonal": {"style": "thin", "up": true, "down": false}})
        );
        let mut f = FillXform::new();
        let xml = r#"<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill>"#;
        let m = parse_part(&mut f, "x", xml.as_bytes()).unwrap().unwrap();
        assert_eq!(
            m,
            json!({"type": "pattern", "pattern": "solid", "fgColor": {"argb": "FFFFFF00"}, "bgColor": {"indexed": 64}})
        );
        assert_eq!(to_xml(&f, Some(&json!({"type": "nope"}))), "");
    }

    #[test]
    fn num_fmt_unescape() {
        assert_eq!(unescape_format_code(r#"0.00\ "x"\\"#), r#"0.00 "x"\"#);
    }
}
