//! ExcelJS `lib/xlsx/xform/strings/*.js`.

use serde_json::{json, Map, Value};

use super::style::{FontXform, RICH_TEXT_FONT_OPTIONS};
use super::{XResult, Xform};
use crate::js::{decode_x_escapes, encode_x_escapes};
use crate::jsv::{get, to_js_string, truthy};
use crate::xml::Tag;
use crate::xml_stream::XmlStream;

/// `TextXform` (`<t>`).
#[derive(Debug, Default)]
pub(crate) struct TextXform {
    text: Vec<String>,
}

impl TextXform {
    /// The `model` getter: joined text with `_xHHHH_` decoded.
    pub fn text_model(&self) -> String {
        decode_x_escapes(&self.text.concat())
    }
}

/// `/^\s|\n|\s$/.test(model)`.
fn needs_preserve(s: &str) -> bool {
    let is_space = |c: char| c.is_whitespace() || c == '\u{feff}';
    s.starts_with(is_space) || s.contains('\n') || s.ends_with(is_space)
}

impl Xform for TextXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if node.name == "t" {
            self.text.clear();
            return Ok(true);
        }
        Ok(false)
    }
    fn parse_text(&mut self, text: &str) {
        self.text.push(text.to_owned());
    }
    fn parse_close(&mut self, _name: &str) -> XResult<bool> {
        Ok(false)
    }
    fn model(&self) -> Option<&Value> {
        None
    }
    fn take_model(&mut self) -> Option<Value> {
        Some(Value::String(self.text_model()))
    }
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        let text = to_js_string(model);
        xml.open_node("t", &[]);
        if needs_preserve(&text) {
            xml.add_attribute("xml:space", Some("preserve"));
        }
        xml.write_text(&encode_x_escapes(&text));
        xml.close_node();
    }
}

/// `RichTextXform` (`<r>`).
#[derive(Debug)]
pub(crate) struct RichTextXform {
    text: TextXform,
    font: FontXform,
    active: Option<bool>, // Some(true) = text, Some(false) = font
    model: Option<Value>,
}

impl RichTextXform {
    pub fn new() -> Self {
        RichTextXform {
            text: TextXform::default(),
            font: FontXform::new(RICH_TEXT_FONT_OPTIONS),
            active: None,
            model: None,
        }
    }
}

impl Xform for RichTextXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        match self.active {
            Some(true) => {
                self.text.parse_open(node)?;
                return Ok(true);
            }
            Some(false) => {
                self.font.parse_open(node)?;
                return Ok(true);
            }
            None => {}
        }
        match node.name {
            "r" => {
                self.model = Some(Value::Object(Map::new()));
                Ok(true)
            }
            "t" => {
                self.active = Some(true);
                self.text.parse_open(node)?;
                Ok(true)
            }
            "rPr" => {
                self.active = Some(false);
                self.font.parse_open(node)?;
                Ok(true)
            }
            _ => Ok(false),
        }
    }
    fn parse_text(&mut self, text: &str) {
        match self.active {
            Some(true) => self.text.parse_text(text),
            Some(false) => self.font.parse_text(text),
            None => {}
        }
    }
    fn parse_close(&mut self, name: &str) -> XResult<bool> {
        let set = |model: &mut Option<Value>, key: &str, v: Value| -> XResult<()> {
            match model {
                Some(Value::Object(o)) => {
                    o.insert(key.into(), v);
                    Ok(())
                }
                _ => Err(format!(
                    "Cannot set properties of undefined (setting '{key}')"
                )),
            }
        };
        match name {
            "r" => Ok(false),
            "t" => {
                let v = match self.active {
                    Some(true) => Some(Value::String(self.text.text_model())),
                    Some(false) => self.font.model().cloned(),
                    None => {
                        return Err("Cannot read properties of undefined (reading 'model')".into())
                    }
                };
                if let Some(v) = v {
                    set(&mut self.model, "text", v)?;
                }
                self.active = None;
                Ok(true)
            }
            "rPr" => {
                let v = match self.active {
                    Some(true) => Some(Value::String(self.text.text_model())),
                    Some(false) => self.font.model().cloned(),
                    None => {
                        return Err("Cannot read properties of undefined (reading 'model')".into())
                    }
                };
                if let Some(v) = v {
                    set(&mut self.model, "font", v)?;
                }
                self.active = None;
                Ok(true)
            }
            _ => {
                match self.active {
                    Some(true) => {
                        self.text.parse_close(name)?;
                    }
                    Some(false) => {
                        self.font.parse_close(name)?;
                    }
                    None => {}
                }
                Ok(true)
            }
        }
    }
    crate::model_accessors!();
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        let model = if truthy(model) {
            model
        } else {
            self.model.as_ref()
        };
        xml.open_node("r", &[]);
        if truthy(get(model, "font")) {
            self.font.render(xml, get(model, "font"));
        }
        self.text.render(xml, get(model, "text"));
        xml.close_node();
    }
    fn reset(&mut self) {
        self.model = Some(Value::Null);
    }
}

/// `PhoneticTextXform` (`<rPh>`).
#[derive(Debug)]
pub(crate) struct PhoneticTextXform {
    r: RichTextXform,
    t: TextXform,
    active: Option<bool>, // Some(true) = r, Some(false) = t
    model: Option<Value>,
}

impl PhoneticTextXform {
    pub fn new() -> Self {
        PhoneticTextXform {
            r: RichTextXform::new(),
            t: TextXform::default(),
            active: None,
            model: None,
        }
    }
}

impl Xform for PhoneticTextXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        match self.active {
            Some(true) => {
                self.r.parse_open(node)?;
                return Ok(true);
            }
            Some(false) => {
                self.t.parse_open(node)?;
                return Ok(true);
            }
            None => {}
        }
        match node.name {
            "rPh" => {
                self.model = Some(json!({
                    "sb": super::simple::int_value(node.attr("sb")),
                    "eb": super::simple::int_value(node.attr("eb")),
                }));
                Ok(true)
            }
            "r" => {
                self.active = Some(true);
                self.r.parse_open(node)?;
                Ok(true)
            }
            "t" => {
                self.active = Some(false);
                self.t.parse_open(node)?;
                Ok(true)
            }
            _ => Ok(false),
        }
    }
    fn parse_text(&mut self, text: &str) {
        match self.active {
            Some(true) => self.r.parse_text(text),
            Some(false) => self.t.parse_text(text),
            None => {}
        }
    }
    fn parse_close(&mut self, name: &str) -> XResult<bool> {
        if let Some(is_r) = self.active {
            let done = if is_r {
                !self.r.parse_close(name)?
            } else {
                !self.t.parse_close(name)?
            };
            if done {
                if let Some(Value::Object(m)) = self.model.as_mut() {
                    match name {
                        "r" => {
                            let run = self.r.model().cloned().unwrap_or(Value::Null);
                            if let Some(Value::Array(rt)) = m.get_mut("richText") {
                                rt.push(run);
                            } else {
                                m.insert("richText".into(), Value::Array(vec![run]));
                            }
                        }
                        "t" => {
                            let v = if is_r {
                                self.r.model().cloned()
                            } else {
                                Some(Value::String(self.t.text_model()))
                            };
                            if let Some(v) = v {
                                m.insert("text".into(), v);
                            }
                        }
                        _ => {}
                    }
                }
                self.active = None;
            }
            return Ok(true);
        }
        Ok(name != "rPh")
    }
    crate::model_accessors!();
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        let or_zero = |k| {
            if truthy(get(model, k)) {
                get(model, k).cloned()
            } else {
                Some(json!(0))
            }
        };
        xml.open_node(
            "rPh",
            &crate::attrs![("sb", or_zero("sb")), ("eb", or_zero("eb"))],
        );
        if let Some(rt) = get(model, "richText").filter(|v| truthy(Some(v))) {
            for run in rt.as_array().into_iter().flatten() {
                self.r.render(xml, Some(run));
            }
        } else if model.is_some() {
            self.t.render(xml, get(model, "text"));
        }
        xml.close_node();
    }
    fn reset(&mut self) {
        self.model = Some(Value::Null);
    }
}

/// What `<si>` is parsing.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum SiChild {
    R,
    T,
    RPh,
}

/// `SharedStringXform` (`<si>`).
#[derive(Debug)]
pub(crate) struct SharedStringXform {
    r: RichTextXform,
    t: TextXform,
    rph: PhoneticTextXform,
    active: Option<SiChild>,
    model: Option<Value>,
}

impl SharedStringXform {
    pub fn new() -> Self {
        SharedStringXform {
            r: RichTextXform::new(),
            t: TextXform::default(),
            rph: PhoneticTextXform::new(),
            active: None,
            model: None,
        }
    }

    fn child(&mut self, c: SiChild) -> &mut dyn Xform {
        match c {
            SiChild::R => &mut self.r,
            SiChild::T => &mut self.t,
            SiChild::RPh => &mut self.rph,
        }
    }
}

impl Xform for SharedStringXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if let Some(c) = self.active {
            self.child(c).parse_open(node)?;
            return Ok(true);
        }
        if node.name == "si" {
            self.model = Some(Value::Object(Map::new()));
            return Ok(true);
        }
        self.active = match node.name {
            "r" => Some(SiChild::R),
            "t" => Some(SiChild::T),
            "rPh" => Some(SiChild::RPh),
            _ => None,
        };
        match self.active {
            Some(c) => {
                self.child(c).parse_open(node)?;
                Ok(true)
            }
            None => Ok(false),
        }
    }
    fn parse_text(&mut self, text: &str) {
        if let Some(c) = self.active {
            self.child(c).parse_text(text);
        }
    }
    fn parse_close(&mut self, name: &str) -> XResult<bool> {
        if let Some(c) = self.active {
            if !self.child(c).parse_close(name)? {
                match name {
                    "r" => {
                        let run = match c {
                            SiChild::R => self.r.model().cloned(),
                            SiChild::T => Some(Value::String(self.t.text_model())),
                            SiChild::RPh => self.rph.model().cloned(),
                        }
                        .unwrap_or(Value::Null);
                        match self.model.as_mut() {
                            Some(Value::Object(m)) => match m.get_mut("richText") {
                                Some(Value::Array(rt)) => rt.push(run),
                                _ => {
                                    m.insert("richText".into(), Value::Array(vec![run]));
                                }
                            },
                            // A string model: ExcelJS assigns a property to a
                            // primitive (ignored) and pushes to a temporary.
                            Some(Value::String(_)) => {}
                            _ => {
                                return Err(
                                    "Cannot read properties of undefined (reading 'richText')"
                                        .into(),
                                )
                            }
                        }
                    }
                    "t" => {
                        self.model = match c {
                            SiChild::T => Some(Value::String(self.t.text_model())),
                            SiChild::R => self.r.model().cloned(),
                            SiChild::RPh => self.rph.model().cloned(),
                        };
                    }
                    _ => {}
                }
                self.active = None;
            }
            return Ok(true);
        }
        if name == "si" {
            // `<si/>` is the empty string; ExcelJS keeps the `{}` it started
            // with (B5).
            if matches!(&self.model, Some(Value::Object(m)) if m.is_empty()) {
                self.model = Some(Value::String(String::new()));
            }
            return Ok(false);
        }
        Ok(true)
    }
    crate::model_accessors!();
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        xml.open_node("si", &[]);
        match model {
            Some(Value::Object(o)) if o.contains_key("richText") && truthy(o.get("richText")) => {
                match o.get("richText").and_then(Value::as_array) {
                    Some(runs) if !runs.is_empty() => {
                        for run in runs {
                            self.r.render(xml, Some(run));
                        }
                    }
                    _ => self.t.render(xml, Some(&json!(""))),
                }
            }
            None | Some(Value::Null) => {}
            Some(m) => self.t.render(xml, Some(m)),
        }
        xml.close_node();
    }
    fn reset(&mut self) {
        self.model = Some(Value::Null);
    }
}

/// `SharedStringsXform` (`xl/sharedStrings.xml`), parse role.
#[derive(Debug)]
pub(crate) struct SharedStringsXform {
    si: SharedStringXform,
    in_si: bool,
    pub values: Vec<Value>,
}

impl SharedStringsXform {
    pub fn new() -> Self {
        SharedStringsXform {
            si: SharedStringXform::new(),
            in_si: false,
            values: Vec::new(),
        }
    }
}

impl Xform for SharedStringsXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if self.in_si {
            self.si.parse_open(node)?;
            return Ok(true);
        }
        match node.name {
            "sst" => Ok(true),
            "si" => {
                self.in_si = true;
                self.si.parse_open(node)?;
                Ok(true)
            }
            other => Err(format!("Unexpected xml node in parseOpen: {other}")),
        }
    }
    fn parse_text(&mut self, text: &str) {
        if self.in_si {
            self.si.parse_text(text);
        }
    }
    fn parse_close(&mut self, name: &str) -> XResult<bool> {
        if self.in_si {
            if !self.si.parse_close(name)? {
                self.values
                    .push(self.si.model().cloned().unwrap_or(Value::Null));
                self.in_si = false;
            }
            return Ok(true);
        }
        match name {
            "sst" => Ok(false),
            other => Err(format!("Unexpected xml node in parseClose: {other}")),
        }
    }
    fn model(&self) -> Option<&Value> {
        None
    }
    fn take_model(&mut self) -> Option<Value> {
        Some(Value::Array(std::mem::take(&mut self.values)))
    }
}

/// Renders `xl/sharedStrings.xml` (`SharedStringsXform.render`).
pub(crate) fn render_shared_strings(values: &[Value], count: usize) -> String {
    let mut xml = XmlStream::new();
    xml.open_xml(&crate::xml_stream::std_doc_attributes());
    xml.open_node(
        "sst",
        &crate::attrs![
            (
                "xmlns",
                Some("http://schemas.openxmlformats.org/spreadsheetml/2006/main")
            ),
            ("count", Some(count)),
            ("uniqueCount", Some(values.len())),
        ],
    );
    let sx = SharedStringXform::new();
    for v in values {
        sx.render(&mut xml, Some(v));
    }
    xml.close_node();
    xml.into_xml()
}

/// `SharedStringXform.toXml(value)`: the key for rich text de-duplication.
pub(crate) fn shared_string_xml(value: &Value) -> String {
    super::to_xml(&SharedStringXform::new(), Some(value))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xform::parse_part;

    #[test]
    fn parses_like_exceljs() {
        let mut x = SharedStringsXform::new();
        let xml = concat!(
            "<sst>",
            "<si><t>a_x0042_</t></si>",
            "<si><r><rPr><b/></rPr><t>x</t></r><r><t>y</t></r></si>",
            "<si><t>base</t><rPh sb=\"0\" eb=\"1\"><t>P</t></rPh></si>",
            "<si/>",
            "<si><r><t>lost</t></r><t>text wins</t><r><t>lost too</t></r></si>",
            "</sst>"
        );
        let m = parse_part(&mut x, "sst", xml.as_bytes()).unwrap().unwrap();
        assert_eq!(
            m,
            json!([
                "aB",
                {"richText": [{"font": {"bold": true}, "text": "x"}, {"text": "y"}]},
                "base",
                "", // ExcelJS: {} (B5)
                "text wins",
            ])
        );
        assert!(parse_part(&mut SharedStringsXform::new(), "sst", b"<sst><x/></sst>").is_err());
    }

    #[test]
    fn renders_like_exceljs() {
        assert_eq!(
            render_shared_strings(&[json!(" a "), json!({"richText": [{"font": {"bold": true}, "text": "b"}]})], 3),
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"3\" uniqueCount=\"2\"><si><t xml:space=\"preserve\"> a </t></si><si><r><rPr><b/></rPr><t>b</t></r></si></sst>"
        );
    }
}
