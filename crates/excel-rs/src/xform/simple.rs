//! ExcelJS `lib/xlsx/xform/simple/*.js`.

use serde_json::Value;

use super::{XResult, Xform};
use crate::js::{parse_float, parse_int};
use crate::jsv::{num, to_js_string, truthy};
use crate::model_json::{date, parse_iso_date, to_iso_string, value_as_date};
use crate::xml::Tag;
use crate::xml_stream::{Attrs, XmlStream};

/// `parseInt(x, 10)` as a model value (`NaN` tagged).
pub(crate) fn int_value(s: Option<&str>) -> Value {
    num(s.and_then(parse_int).unwrap_or(f64::NAN))
}

/// `parseFloat(x)` as a model value; `parseFloat(undefined)` is `NaN`.
pub(crate) fn float_value(s: Option<&str>) -> Value {
    num(s.map_or(f64::NAN, parse_float))
}

/// Shared shape of the simple xforms: `{tag, attr, attrs}`.
#[derive(Debug, Clone, Default)]
pub(crate) struct SimpleOptions {
    pub tag: &'static str,
    pub attr: Option<&'static str>,
    pub attrs: Attrs,
}

impl SimpleOptions {
    pub fn tag(tag: &'static str) -> Self {
        SimpleOptions {
            tag,
            ..Self::default()
        }
    }

    pub fn attr(tag: &'static str, attr: &'static str) -> Self {
        SimpleOptions {
            tag,
            attr: Some(attr),
            ..Self::default()
        }
    }

    fn render_value(&self, xml: &mut XmlStream, text: &str) {
        xml.open_node(self.tag, &[]);
        if !self.attrs.is_empty() {
            xml.add_attributes(&self.attrs);
        }
        match self.attr {
            Some(attr) => xml.add_attribute(attr, Some(text)),
            None => xml.write_text(text),
        }
        xml.close_node();
    }
}

/// `BooleanXform`: an element whose presence means `true`.
#[derive(Debug)]
pub(crate) struct BooleanXform {
    pub tag: &'static str,
    model: Option<Value>,
}

impl BooleanXform {
    pub fn new(tag: &'static str) -> Self {
        BooleanXform { tag, model: None }
    }
}

impl Xform for BooleanXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if node.name == self.tag {
            self.model = Some(Value::Bool(true));
        }
        Ok(false)
    }
    fn parse_close(&mut self, _name: &str) -> XResult<bool> {
        Ok(false)
    }
    fn model(&self) -> Option<&Value> {
        self.model.as_ref()
    }
    fn take_model(&mut self) -> Option<Value> {
        self.model.take()
    }
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        if truthy(model) {
            xml.open_node(self.tag, &[]);
            xml.close_node();
        }
    }
    fn reset(&mut self) {
        self.model = Some(Value::Null);
    }
}

/// How a `DateXform` formats and parses (ExcelJS passes functions).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum DateFormat {
    /// core.xml (`CoreXform.DateFormat`): `toISOString()` with the first
    /// `.ddd` removed; throws for anything but a valid Date.
    Core,
}

/// `DateXform`.
#[derive(Debug)]
pub(crate) struct DateXform {
    pub opts: SimpleOptions,
    pub format: DateFormat,
    text: String,
    model: Option<Value>,
}

impl DateXform {
    pub fn new(opts: SimpleOptions, format: DateFormat) -> Self {
        DateXform {
            opts,
            format,
            text: String::new(),
            model: None,
        }
    }

    fn format_date(&self, model: &Value) -> Result<String, String> {
        let iso = value_as_date(Some(model)).and_then(to_iso_string);
        match (self.format, iso) {
            (DateFormat::Core, Some(iso)) => Ok(remove_first_millis(&iso)),
            (DateFormat::Core, None) => Err(if value_as_date(Some(model)).is_some() {
                "Invalid time value".to_owned()
            } else {
                "dt.toISOString is not a function".to_owned()
            }),
        }
    }
}

/// `.replace(/[.]\d{3}/, '')`.
fn remove_first_millis(iso: &str) -> String {
    let b = iso.as_bytes();
    for i in 0..b.len() {
        if b.get(i) == Some(&b'.')
            && b.get(i + 1..i + 4)
                .is_some_and(|d| d.iter().all(u8::is_ascii_digit))
        {
            return format!(
                "{}{}",
                iso.get(..i).unwrap_or(""),
                iso.get(i + 4..).unwrap_or("")
            );
        }
    }
    iso.to_owned()
}

fn parse_date(s: &str) -> Value {
    date(parse_iso_date(s))
}

impl Xform for DateXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if node.name == self.opts.tag {
            match self.opts.attr {
                Some(a) => self.model = Some(parse_date(node.attr(a).unwrap_or("undefined"))),
                None => self.text.clear(),
            }
        }
        Ok(false)
    }
    fn parse_text(&mut self, text: &str) {
        if self.opts.attr.is_none() {
            self.text.push_str(text);
        }
    }
    fn parse_close(&mut self, _name: &str) -> XResult<bool> {
        if self.opts.attr.is_none() {
            self.model = Some(parse_date(&self.text));
        }
        Ok(false)
    }
    fn model(&self) -> Option<&Value> {
        self.model.as_ref()
    }
    fn take_model(&mut self) -> Option<Value> {
        self.model.take()
    }
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        if let Some(m) = model.filter(|m| truthy(Some(m))) {
            match self.format_date(m) {
                Ok(text) => self.opts.render_value(xml, &text),
                Err(e) => xml.fail(e),
            }
        }
    }
    fn reset(&mut self) {
        self.model = Some(Value::Null);
    }
}

/// Which simple numeric/string kind.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Kind {
    Integer { zero: bool },
    String,
}

/// `FloatXform`, `IntegerXform` and `StringXform`.
#[derive(Debug)]
pub(crate) struct ValueXform {
    pub opts: SimpleOptions,
    pub kind: Kind,
    text: String,
    model: Option<Value>,
}

impl ValueXform {
    pub fn new(opts: SimpleOptions, kind: Kind) -> Self {
        ValueXform {
            opts,
            kind,
            text: String::new(),
            model: None,
        }
    }
    pub fn integer(opts: SimpleOptions) -> Self {
        Self::new(opts, Kind::Integer { zero: false })
    }
    pub fn integer_zero(opts: SimpleOptions) -> Self {
        Self::new(opts, Kind::Integer { zero: true })
    }
    pub fn string(opts: SimpleOptions) -> Self {
        Self::new(opts, Kind::String)
    }

    fn convert(&self, raw: Option<&str>) -> Value {
        match self.kind {
            Kind::Integer { .. } => int_value(raw),
            Kind::String => raw.map_or(Value::Null, |s| Value::String(s.to_owned())),
        }
    }
}

impl Xform for ValueXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if node.name == self.opts.tag {
            match self.opts.attr {
                Some(a) => {
                    let raw = node.attr(a);
                    self.model = match (self.kind, raw) {
                        // `node.attributes[attr]` is undefined.
                        (Kind::String, None) => None,
                        _ => Some(self.convert(raw)),
                    };
                }
                None => self.text.clear(),
            }
            return Ok(matches!(self.kind, Kind::Integer { .. }));
        }
        Ok(false)
    }
    fn parse_text(&mut self, text: &str) {
        if self.opts.attr.is_none() {
            self.text.push_str(text);
        }
    }
    fn parse_close(&mut self, _name: &str) -> XResult<bool> {
        if self.opts.attr.is_none() {
            self.model = Some(match self.kind {
                // parseInt(this.text.join('') || 0, 10)
                Kind::Integer { .. } if self.text.is_empty() => num(0.0),
                _ => self.convert(Some(&self.text)),
            });
        }
        Ok(false)
    }
    fn model(&self) -> Option<&Value> {
        self.model.as_ref()
    }
    fn take_model(&mut self) -> Option<Value> {
        self.model.take()
    }
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        let render = match self.kind {
            Kind::Integer { zero } => truthy(model) || zero,
            Kind::String => model.is_some(),
        };
        if render {
            self.opts.render_value(xml, &to_js_string(model));
        }
    }
    fn reset(&mut self) {
        self.model = Some(Value::Null);
    }
}
