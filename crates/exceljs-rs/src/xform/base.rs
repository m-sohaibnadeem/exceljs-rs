//! ExcelJS `list-xform.js`, `static-xform.js` and helpers shared by the
//! composite xforms (the `this.map` / `this.parser` pattern).

use serde_json::Value;

use super::{XResult, Xform};
use crate::xml::Tag;
use crate::xml_stream::{Attrs, XmlStream};

pub(crate) type BoxX = Box<dyn Xform>;

/// `this.map` plus `this.parser`: named child xforms and the active one.
pub(crate) struct Children {
    items: Vec<(&'static str, BoxX)>,
    pub active: Option<usize>,
}

impl std::fmt::Debug for Children {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_list()
            .entries(self.items.iter().map(|(n, _)| n))
            .finish()
    }
}

impl Children {
    pub fn new(items: Vec<(&'static str, BoxX)>) -> Self {
        Children {
            items,
            active: None,
        }
    }

    pub fn index(&self, name: &str) -> Option<usize> {
        self.items.iter().position(|(n, _)| *n == name)
    }

    pub fn get(&self, name: &str) -> Option<&dyn Xform> {
        self.items
            .iter()
            .find(|(n, _)| *n == name)
            .map(|(_, x)| x.as_ref())
    }

    pub fn get_mut(&mut self, name: &str) -> Option<&mut BoxX> {
        self.items
            .iter_mut()
            .find(|(n, _)| *n == name)
            .map(|(_, x)| x)
    }

    pub fn at(&self, i: usize) -> Option<&dyn Xform> {
        self.items.get(i).map(|(_, x)| x.as_ref())
    }

    pub fn at_mut(&mut self, i: usize) -> Option<&mut BoxX> {
        self.items.get_mut(i).map(|(_, x)| x)
    }

    pub fn name_at(&self, i: usize) -> Option<&'static str> {
        self.items.get(i).map(|(n, _)| *n)
    }

    pub fn parser(&mut self) -> Option<&mut BoxX> {
        let i = self.active?;
        self.at_mut(i)
    }

    /// `this.parser = this.map[name]`; returns whether one was found.
    pub fn activate(&mut self, name: &str) -> bool {
        self.active = self.index(name);
        self.active.is_some()
    }

    /// The active child's name.
    pub fn active_name(&self) -> Option<&'static str> {
        self.name_at(self.active?)
    }

    /// `this.parser.model` of the active child.
    pub fn parser_model(&self) -> Option<&Value> {
        self.at(self.active?)?.model()
    }

    /// `this.map[name].model`.
    pub fn model_of(&self, name: &str) -> Option<&Value> {
        self.get(name)?.model()
    }

    /// Resets every child (BaseXform.reset's `this.map` loop).
    pub fn reset_all(&mut self) {
        for (_, x) in &mut self.items {
            x.reset();
        }
    }

    /// Forwards an open to the active parser if any (`if (this.parser)
    /// { this.parser.parseOpen(node); return true; }`).
    pub fn forward_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        match self.parser() {
            Some(p) => {
                p.parse_open(node)?;
                Ok(true)
            }
            None => Ok(false),
        }
    }

    pub fn forward_text(&mut self, text: &str) {
        if let Some(p) = self.parser() {
            p.parse_text(text);
        }
    }
}

/// Implements `model`, `take_model` for a struct with `model: Option<Value>`.
#[macro_export]
#[doc(hidden)]
macro_rules! model_accessors {
    () => {
        fn model(&self) -> Option<&serde_json::Value> {
            self.model.as_ref()
        }
        fn take_model(&mut self) -> Option<serde_json::Value> {
            self.model.take()
        }
    };
}

/// `ListXform`.
pub(crate) struct ListXform {
    pub tag: &'static str,
    pub always: bool,
    pub count: bool,
    pub empty: bool,
    pub count_attr: &'static str,
    pub attrs: Attrs,
    pub child: BoxX,
    pub max_items: Option<usize>,
    /// The child xform's tag, for the "Max <tag> count" error.
    pub child_tag: &'static str,
    in_child: bool,
    model: Option<Value>,
}

impl std::fmt::Debug for ListXform {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "ListXform({})", self.tag)
    }
}

impl ListXform {
    pub fn new(tag: &'static str, child: BoxX) -> Self {
        ListXform {
            tag,
            always: false,
            count: false,
            empty: false,
            count_attr: "count",
            attrs: Vec::new(),
            child,
            max_items: None,
            child_tag: "undefined",
            in_child: false,
            model: None,
        }
    }
    pub fn count(mut self) -> Self {
        self.count = true;
        self
    }
    pub fn always(mut self) -> Self {
        self.always = true;
        self
    }
    pub fn empty(mut self) -> Self {
        self.empty = true;
        self
    }
    pub fn attrs(mut self, attrs: Attrs) -> Self {
        self.attrs = attrs;
        self
    }
    pub fn render_list(&self, xml: &mut XmlStream, model: Option<&Value>) {
        let items = model.and_then(Value::as_array);
        let len = items.map_or(0, Vec::len);
        if self.always || len > 0 {
            xml.open_node(self.tag, &self.attrs);
            if self.count {
                xml.add_attribute(self.count_attr, Some(len));
            }
            for item in items.into_iter().flatten() {
                self.child.render(xml, Some(item));
            }
            xml.close_node();
        } else if self.empty {
            xml.leaf_node(self.tag, &[], None);
        }
    }
}

impl Xform for ListXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if self.in_child {
            self.child.parse_open(node)?;
            return Ok(true);
        }
        if node.name == self.tag {
            self.model = Some(Value::Array(Vec::new()));
            return Ok(true);
        }
        if self.child.parse_open(node)? {
            self.in_child = true;
            return Ok(true);
        }
        Ok(false)
    }
    fn parse_text(&mut self, text: &str) {
        if self.in_child {
            self.child.parse_text(text);
        }
    }
    fn parse_close(&mut self, name: &str) -> XResult<bool> {
        if self.in_child {
            if !self.child.parse_close(name)? {
                let item = self.child.model().cloned().unwrap_or(Value::Null);
                match self.model.as_mut() {
                    Some(Value::Array(a)) => a.push(item),
                    // `this.model.push` on null/undefined: TypeError.
                    _ => {
                        return Err(format!(
                            "Cannot read properties of {} (reading 'push')",
                            if self.model.is_some() {
                                "null"
                            } else {
                                "undefined"
                            }
                        ))
                    }
                }
                self.in_child = false;
                if let (Some(max), Some(Value::Array(a))) = (self.max_items, self.model.as_ref()) {
                    if a.len() > max {
                        return Err(format!("Max {} count ({max}) exceeded", self.child_tag));
                    }
                }
            }
            return Ok(true);
        }
        Ok(false)
    }
    crate::model_accessors!();
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        self.render_list(xml, model);
    }
    fn reset(&mut self) {
        self.model = Some(Value::Null);
        self.child.reset();
    }
}

/// A node of a `StaticXform` model: `{tag, $, c, t}`.
#[derive(Debug, Clone)]
pub(crate) struct StaticNode {
    pub tag: &'static str,
    pub attrs: Attrs,
    pub children: Vec<StaticNode>,
    pub text: Option<&'static str>,
}

impl StaticNode {
    pub fn new(tag: &'static str, attrs: Attrs, children: Vec<StaticNode>) -> Self {
        StaticNode {
            tag,
            attrs,
            children,
            text: None,
        }
    }

    fn build(&self, xml: &mut XmlStream) {
        xml.open_node(self.tag, &self.attrs);
        for c in &self.children {
            c.build(xml);
        }
        if let Some(t) = self.text.filter(|t| !t.is_empty()) {
            xml.write_text(t);
        }
        xml.close_node();
    }
}

/// `StaticXform`: renders fixed XML; while parsing it only tracks its own
/// element.
#[derive(Debug)]
pub(crate) struct StaticXform {
    pub node: StaticNode,
    model: Option<Value>,
}

impl StaticXform {
    pub fn new(node: StaticNode) -> Self {
        StaticXform { node, model: None }
    }

    pub fn xml(&self) -> String {
        let mut x = XmlStream::new();
        self.node.build(&mut x);
        x.into_xml()
    }
}

impl Xform for StaticXform {
    fn parse_open(&mut self, _node: &Tag<'_>) -> XResult<bool> {
        Ok(true)
    }
    fn parse_close(&mut self, name: &str) -> XResult<bool> {
        Ok(name != self.node.tag)
    }
    crate::model_accessors!();
    fn render(&self, xml: &mut XmlStream, _model: Option<&Value>) {
        xml.write_xml(&self.xml());
    }
    fn reset(&mut self) {
        self.model = Some(Value::Null);
    }
}
