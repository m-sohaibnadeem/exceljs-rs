//! Port of ExcelJS's xforms (`lib/xlsx/xform/**`): each one parses one kind
//! of XML element into ExcelJS's model and renders it back.
//!
//! The parse protocol is ExcelJS's: a parent forwards `parse_open`,
//! `parse_text` and `parse_close` to the active child and incorporates the
//! child's model when the child's `parse_close` returns `false`. Keeping
//! that protocol keeps ExcelJS's results on unusual XML.

pub(crate) mod base;
pub(crate) mod book;
pub(crate) mod cell;
pub(crate) mod cf;
pub(crate) mod comment;
pub(crate) mod composite;
pub(crate) mod core;
pub(crate) mod drawing;
pub(crate) mod sheet;
pub(crate) mod simple;
pub(crate) mod strings;
pub(crate) mod style;
pub(crate) mod table;
pub(crate) mod worksheet;

use serde_json::Value;

use crate::error::{Error, Result};
use crate::xml::{self, Sax, Tag};
use crate::xml_stream::XmlStream;

/// Errors inside xforms carry a message; the driver adds the part name.
pub(crate) type XResult<T> = std::result::Result<T, String>;

pub(crate) trait Xform {
    /// `parseOpen(node)`: the JS truthiness of its return value.
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool>;
    /// `parseText(text)`.
    fn parse_text(&mut self, _text: &str) {}
    /// `parseClose(name)`: `false` when this xform's element is complete.
    fn parse_close(&mut self, name: &str) -> XResult<bool>;
    /// `this.model` (`None` is `undefined`).
    fn model(&self) -> Option<&Value>;
    /// Moves `this.model` out (the parent takes ownership).
    fn take_model(&mut self) -> Option<Value>;
    /// `render(xmlStream, model)`.
    fn render(&self, _xml: &mut XmlStream, _model: Option<&Value>) {}
    /// `reset()`.
    fn reset(&mut self) {}
}

/// `xform.parseStream(stream)`: feeds a whole part to `xform` and returns its
/// model.
pub(crate) fn parse_part(xform: &mut dyn Xform, part: &str, bytes: &[u8]) -> Result<Option<Value>> {
    let mut failure = None;
    xml::parse(part, bytes, |event| {
        let step = match event {
            Sax::Open(tag) => xform.parse_open(tag).map(|_| true),
            Sax::Text(text) => {
                xform.parse_text(text);
                Ok(true)
            }
            Sax::Close(name) => xform.parse_close(name),
        };
        match step {
            Ok(go_on) => Ok(go_on),
            Err(message) => {
                failure = Some(message);
                Ok(false)
            }
        }
    })?;
    if let Some(message) = failure {
        return Err(Error::invalid(part, message));
    }
    Ok(xform.take_model())
}

/// Renders `model` with `xform` into a standalone XML string (`toXml`).
pub(crate) fn to_xml(xform: &dyn Xform, model: Option<&Value>) -> String {
    let mut xml = XmlStream::new();
    xform.render(&mut xml, model);
    xml.into_xml()
}
