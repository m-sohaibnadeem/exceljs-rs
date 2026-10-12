//! Generic forms of three recurring ExcelJS xform shapes:
//!
//! * [`Composite`]: `CompositeXform` subclasses (a tag, `createNewModel`,
//!   `onParserClose`, and a render function);
//! * [`TextAccum`]: `parseOpen() { model = '' }`, `parseText` appends,
//!   `parseClose(name) { return name !== tag }`;
//! * [`Leaf`]: one element whose attributes make the model.

use serde_json::Value;

use super::base::{BoxX, Children};
use super::{XResult, Xform};
use crate::xml::Tag;
use crate::xml_stream::XmlStream;

pub(crate) type CreateFn = fn(&Tag<'_>) -> Value;
pub(crate) type OnCloseFn = fn(&mut Value, &str, Option<Value>) -> XResult<()>;
pub(crate) type RenderFn = fn(&Composite, &mut XmlStream, Option<&Value>);

/// `CompositeXform`.
pub(crate) struct Composite {
    pub tag: &'static str,
    pub map: Children,
    create: CreateFn,
    on_close: OnCloseFn,
    render_fn: RenderFn,
    model: Option<Value>,
}

impl std::fmt::Debug for Composite {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "Composite({})", self.tag)
    }
}

impl Composite {
    pub fn new(
        tag: &'static str,
        children: Vec<(&'static str, BoxX)>,
        create: CreateFn,
        on_close: OnCloseFn,
        render_fn: RenderFn,
    ) -> Self {
        Composite {
            tag,
            map: Children::new(children),
            create,
            on_close,
            render_fn,
            model: None,
        }
    }

    /// Renders child `name` (`this.map[name].render(...)`).
    pub fn render_child(&self, name: &str, xml: &mut XmlStream, model: Option<&Value>) {
        if let Some(x) = self.map.get(name) {
            x.render(xml, model);
        }
    }
}

impl Xform for Composite {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        // this.parser = this.parser || this.map[node.name]
        if self.map.active.is_none() {
            self.map.active = self.map.index(node.name);
        }
        if let Some(p) = self.map.parser() {
            p.parse_open(node)?;
            return Ok(true);
        }
        if node.name == self.tag {
            self.model = Some((self.create)(node));
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
                let child_model = self.map.parser_model().cloned();
                match self.model.as_mut() {
                    Some(m) => (self.on_close)(m, name, child_model)?,
                    None => {
                        return Err(format!(
                            "Cannot set properties of undefined (onParserClose {name})"
                        ))
                    }
                }
                self.map.active = None;
            }
            return Ok(true);
        }
        Ok(name != self.tag)
    }
    crate::model_accessors!();
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        (self.render_fn)(self, xml, model);
    }
    fn reset(&mut self) {
        self.model = Some(Value::Null);
        self.map.reset_all();
    }
}

/// `onParserClose`: `Object.assign(this.model, parser.model)`.
pub(crate) fn merge_into(model: &mut Value, _name: &str, child: Option<Value>) -> XResult<()> {
    if let (Value::Object(o), Some(Value::Object(c))) = (model, child) {
        for (k, v) in c {
            o.insert(k, v);
        }
    }
    Ok(())
}

/// `this.model[key].push(parser.model)`.
pub(crate) fn push_to(model: &mut Value, key: &str, child: Option<Value>) -> XResult<()> {
    match model.get_mut(key).and_then(Value::as_array_mut) {
        Some(a) => {
            a.push(child.unwrap_or(Value::Null));
            Ok(())
        }
        None => Err(format!(
            "Cannot read properties of undefined (reading 'push') [{key}]"
        )),
    }
}

/// Text accumulator elements (`formula`, `x14:id`, `xm:sqref`, `xm:f`).
#[derive(Debug)]
pub(crate) struct TextAccum {
    pub tag: &'static str,
    model: Option<Value>,
}

impl TextAccum {
    pub fn new(tag: &'static str) -> Self {
        TextAccum { tag, model: None }
    }
}

impl Xform for TextAccum {
    fn parse_open(&mut self, _node: &Tag<'_>) -> XResult<bool> {
        self.model = Some(Value::String(String::new()));
        Ok(false)
    }
    fn parse_text(&mut self, text: &str) {
        // this.model += text (undefined + text is "undefined..." in JS)
        match self.model.as_mut() {
            Some(Value::String(s)) => s.push_str(text),
            Some(other) => {
                let s = format!("{}{text}", crate::jsv::to_js_string(Some(other)));
                *other = Value::String(s);
            }
            None => self.model = Some(Value::String(format!("undefined{text}"))),
        }
    }
    fn parse_close(&mut self, name: &str) -> XResult<bool> {
        Ok(name != self.tag)
    }
    crate::model_accessors!();
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        let text = model.map(|m| crate::jsv::to_js_string(Some(m)));
        xml.leaf_node(self.tag, &[], text.as_deref());
    }
    fn reset(&mut self) {
        self.model = Some(Value::Null);
    }
}

pub(crate) type LeafParseFn = fn(&Tag<'_>) -> Option<Value>;
pub(crate) type LeafRenderFn = fn(&mut XmlStream, Option<&Value>);

/// How a [`Leaf`] decides which elements are its own.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum LeafMatch {
    /// `if (node.name === tag) { model = ...; return true } return false`;
    /// `parseClose() { return false }`.
    Tag,
    /// `parseOpen(node) { model = ... }` for any node (returns undefined);
    /// `parseClose(name) { return name !== tag }`.
    AnyUntilClose,
    /// AutoFilter: `if (node.name === tag) model = ...` returning undefined,
    /// no parseClose (undefined: done).
    TagNoReturn,
}

/// One element whose attributes make the model.
pub(crate) struct Leaf {
    pub tag: &'static str,
    kind: LeafMatch,
    parse: LeafParseFn,
    render_fn: LeafRenderFn,
    model: Option<Value>,
}

impl std::fmt::Debug for Leaf {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "Leaf({})", self.tag)
    }
}

impl Leaf {
    pub fn new(
        tag: &'static str,
        kind: LeafMatch,
        parse: LeafParseFn,
        render_fn: LeafRenderFn,
    ) -> Self {
        Leaf {
            tag,
            kind,
            parse,
            render_fn,
            model: None,
        }
    }
}

impl Xform for Leaf {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        match self.kind {
            LeafMatch::Tag => {
                if node.name == self.tag {
                    self.model = (self.parse)(node);
                    return Ok(true);
                }
                Ok(false)
            }
            LeafMatch::TagNoReturn => {
                if node.name == self.tag {
                    self.model = (self.parse)(node);
                }
                Ok(false)
            }
            LeafMatch::AnyUntilClose => {
                self.model = (self.parse)(node);
                Ok(false)
            }
        }
    }
    fn parse_close(&mut self, name: &str) -> XResult<bool> {
        Ok(match self.kind {
            LeafMatch::Tag | LeafMatch::TagNoReturn => false,
            LeafMatch::AnyUntilClose => name != self.tag,
        })
    }
    crate::model_accessors!();
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        (self.render_fn)(xml, model);
    }
    fn reset(&mut self) {
        self.model = Some(Value::Null);
    }
}

// ---------------------------------------------------------------------------
// BaseXform static helpers

/// JS `a !== b` for model values against a default.
fn strictly_differs(value: &Value, dflt: &Value) -> bool {
    match (
        crate::jsv::as_f64(Some(value)),
        crate::jsv::as_f64(Some(dflt)),
    ) {
        (Some(a), Some(b)) => a != b || a.is_nan(),
        _ => match (value, dflt) {
            (Value::String(a), Value::String(b)) => a != b,
            (Value::Bool(a), Value::Bool(b)) => a != b,
            (Value::Null, Value::Null) => false,
            _ => true,
        },
    }
}

/// `BaseXform.toAttribute(value, dflt, always)`.
pub(crate) fn to_attribute(
    value: Option<&Value>,
    dflt: Option<Value>,
    always: bool,
) -> Option<String> {
    match value {
        None => {
            if always {
                dflt.map(|d| crate::jsv::to_js_string(Some(&d)))
            } else {
                None
            }
        }
        Some(v) => {
            if always || dflt.as_ref().is_none_or(|d| strictly_differs(v, d)) {
                Some(crate::jsv::to_js_string(Some(v)))
            } else {
                None
            }
        }
    }
}

/// `BaseXform.toBoolAttribute(value, dflt, always)`.
pub(crate) fn to_bool_attribute(value: Option<&Value>, dflt: bool, always: bool) -> Option<String> {
    match value {
        None => always.then(|| if dflt { "1" } else { "0" }.to_owned()),
        Some(v) => {
            if always || strictly_differs(v, &Value::Bool(dflt)) {
                Some(
                    if crate::jsv::truthy(Some(v)) {
                        "1"
                    } else {
                        "0"
                    }
                    .to_owned(),
                )
            } else {
                None
            }
        }
    }
}

/// `BaseXform.toIntValue(attr, dflt)`.
pub(crate) fn to_int_value(attr: Option<&str>, dflt: Option<Value>) -> Option<Value> {
    match attr {
        None => dflt,
        Some(a) => Some(super::simple::int_value(Some(a))),
    }
}

/// `BaseXform.toBoolValue(attr, dflt)`.
pub(crate) fn to_bool_value(attr: Option<&str>, dflt: Option<bool>) -> Option<Value> {
    match attr {
        None => dflt.map(Value::Bool),
        Some(a) => Some(Value::Bool(a == "1")),
    }
}

/// `BaseXform.toStringValue(attr, dflt)`.
pub(crate) fn to_string_value(attr: Option<&str>, dflt: Option<&str>) -> Option<Value> {
    attr.or(dflt).map(|s| Value::String(s.to_owned()))
}

/// `BaseXform.toFloatValue(attr, dflt)`.
pub(crate) fn to_float_value(attr: Option<&str>, dflt: Option<Value>) -> Option<Value> {
    match attr {
        None => dflt,
        Some(a) => Some(super::simple::float_value(Some(a))),
    }
}

/// Builds a JSON object from `(key, Option<Value>)` pairs, skipping `None`
/// (JS properties whose value is `undefined`).
pub(crate) fn object(pairs: Vec<(&str, Option<Value>)>) -> Value {
    let mut m = serde_json::Map::new();
    for (k, v) in pairs {
        if let Some(v) = v {
            m.insert(k.to_owned(), v);
        }
    }
    Value::Object(m)
}
