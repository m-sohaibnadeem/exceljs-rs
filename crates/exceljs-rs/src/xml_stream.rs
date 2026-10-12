//! Port of ExcelJS `lib/utils/xml-stream.js`: the XML builder every xform
//! renders through. Output is byte-for-byte what ExcelJS produces.

use serde_json::Value;

use crate::jsv::{escape_into, number_to_string, to_js_string};

/// A value that can be written as an XML attribute (JS `value.toString()`).
pub(crate) trait AttrValue {
    fn attr_string(&self) -> String;
    /// Appends the escaped attribute value to `out`.
    fn write_attr(&self, out: &mut String) {
        escape_into(out, &self.attr_string());
    }
}

impl AttrValue for &str {
    fn attr_string(&self) -> String {
        (*self).to_owned()
    }
    fn write_attr(&self, out: &mut String) {
        escape_into(out, self);
    }
}
impl AttrValue for String {
    fn attr_string(&self) -> String {
        self.clone()
    }
    fn write_attr(&self, out: &mut String) {
        escape_into(out, self);
    }
}
impl AttrValue for &String {
    fn attr_string(&self) -> String {
        (*self).clone()
    }
    fn write_attr(&self, out: &mut String) {
        escape_into(out, self);
    }
}
impl AttrValue for f64 {
    fn attr_string(&self) -> String {
        number_to_string(*self)
    }
}
impl AttrValue for i64 {
    fn attr_string(&self) -> String {
        self.to_string()
    }
}
impl AttrValue for u32 {
    fn attr_string(&self) -> String {
        self.to_string()
    }
}
impl AttrValue for usize {
    fn attr_string(&self) -> String {
        self.to_string()
    }
}
impl AttrValue for bool {
    fn attr_string(&self) -> String {
        self.to_string()
    }
}
impl AttrValue for &Value {
    fn attr_string(&self) -> String {
        to_js_string(Some(self))
    }
    fn write_attr(&self, out: &mut String) {
        match self {
            Value::String(s) => escape_into(out, s),
            Value::Number(n) => out.push_str(&crate::jsv::number_value_to_string(n)),
            other => escape_into(out, &to_js_string(Some(other))),
        }
    }
}
impl AttrValue for Value {
    fn attr_string(&self) -> String {
        to_js_string(Some(self))
    }
}

/// An ordered attribute list; `None` values are skipped (JS `undefined`).
pub(crate) type Attrs = Vec<(String, Option<String>)>;

/// Builds an [`Attrs`] list: `attrs![("r", Some(x)), ("t", None::<&str>)]`.
#[macro_export]
#[doc(hidden)]
macro_rules! attrs {
    ($(($name:expr, $value:expr)),* $(,)?) => {{
        #[allow(clippy::useless_vec)]
        let attrs: $crate::xml_stream::Attrs =
            vec![$(($name.to_string(), ($value).map(|v| $crate::xml_stream::AttrValue::attr_string(&v)))),*];
        attrs
    }};
}

#[derive(Debug, Clone, Copy)]
struct Rollback {
    xml: usize,
    stack: usize,
    leaf: bool,
    open: bool,
}

#[derive(Debug, Default)]
pub(crate) struct XmlStream {
    xml: String,
    /// Names of the open nodes, one after the other; `stack` holds where
    /// each starts.
    names: String,
    stack: Vec<usize>,
    rollbacks: Vec<Rollback>,
    leaf: bool,
    open: bool,
    /// Set where ExcelJS's render would throw; the writer reports it.
    error: Option<String>,
    /// Bytes already taken by [`drain`](Self::drain).
    drained: usize,
}

pub(crate) fn std_doc_attributes() -> Attrs {
    vec![
        ("version".into(), Some("1.0".into())),
        ("encoding".into(), Some("UTF-8".into())),
        ("standalone".into(), Some("yes".into())),
    ]
}

impl XmlStream {
    pub fn new() -> Self {
        Self::default()
    }

    fn push_attributes(&mut self, attrs: &[(String, Option<String>)]) {
        for (name, value) in attrs {
            if let Some(v) = value {
                self.push_attribute(name, v);
            }
        }
    }

    fn push_attribute(&mut self, name: &str, value: &str) {
        self.xml.push(' ');
        self.xml.push_str(name);
        self.xml.push_str("=\"");
        escape_into(&mut self.xml, value);
        self.xml.push('"');
    }

    /// Byte length so far; ExcelJS's `cursor` is only compared for change.
    pub fn cursor(&self) -> usize {
        self.drained + self.xml.len()
    }

    /// Takes what is written so far if it is at least `min` bytes and no
    /// rollback can still need it (the rest of the document follows it).
    pub fn drain(&mut self, min: usize) -> Option<String> {
        if self.xml.len() < min || !self.rollbacks.is_empty() {
            return None;
        }
        self.drained += self.xml.len();
        Some(std::mem::replace(
            &mut self.xml,
            String::with_capacity(min + min / 4),
        ))
    }

    pub fn open_xml(&mut self, attrs: &[(String, Option<String>)]) {
        self.xml.push_str("<?xml");
        self.push_attributes(attrs);
        self.xml.push_str("?>\n");
    }

    pub fn open_node(&mut self, name: &str, attrs: &[(String, Option<String>)]) {
        if !self.stack.is_empty() && self.open {
            self.xml.push('>');
        }
        self.stack.push(self.names.len());
        self.names.push_str(name);
        self.xml.push('<');
        self.xml.push_str(name);
        self.push_attributes(attrs);
        self.leaf = true;
        self.open = true;
    }

    /// `addAttribute(name, value)`; `None` is JS `undefined` (skipped).
    /// ExcelJS throws if no node is open; that would be a bug here, so the
    /// attribute is dropped instead.
    pub fn add_attribute(&mut self, name: &str, value: Option<impl AttrValue>) {
        if self.open {
            if let Some(v) = value {
                self.xml.push(' ');
                self.xml.push_str(name);
                self.xml.push_str("=\"");
                v.write_attr(&mut self.xml);
                self.xml.push('"');
            }
        }
    }

    pub fn add_attributes(&mut self, attrs: &[(String, Option<String>)]) {
        if self.open {
            self.push_attributes(attrs);
        }
    }

    pub fn write_text(&mut self, text: &str) {
        if self.open {
            self.xml.push('>');
            self.open = false;
        }
        self.leaf = false;
        escape_into(&mut self.xml, text);
    }

    pub fn write_xml(&mut self, xml: &str) {
        if self.open {
            self.xml.push('>');
            self.open = false;
        }
        self.leaf = false;
        self.xml.push_str(xml);
    }

    pub fn close_node(&mut self) {
        let start = self.stack.pop().unwrap_or(self.names.len());
        if self.leaf {
            self.xml.push_str("/>");
        } else {
            self.xml.push_str("</");
            self.xml.push_str(self.names.get(start..).unwrap_or(""));
            self.xml.push('>');
        }
        self.names.truncate(start);
        self.open = false;
        self.leaf = false;
    }

    /// `leafNode(name, attributes, text)`; `text` `None` is `undefined`.
    pub fn leaf_node(
        &mut self,
        name: &str,
        attrs: &[(String, Option<String>)],
        text: Option<&str>,
    ) {
        self.open_node(name, attrs);
        if let Some(t) = text {
            self.write_text(t);
        }
        self.close_node();
    }

    pub fn close_all(&mut self) {
        while !self.stack.is_empty() {
            self.close_node();
        }
    }

    pub fn add_rollback(&mut self) -> usize {
        self.rollbacks.push(Rollback {
            xml: self.xml.len(),
            stack: self.stack.len(),
            leaf: self.leaf,
            open: self.open,
        });
        self.cursor()
    }

    pub fn commit(&mut self) {
        self.rollbacks.pop();
    }

    pub fn rollback(&mut self) {
        if let Some(r) = self.rollbacks.pop() {
            self.xml.truncate(r.xml);
            if let Some(&start) = self.stack.get(r.stack) {
                self.names.truncate(start);
            }
            self.stack.truncate(r.stack);
            self.leaf = r.leaf;
            self.open = r.open;
        }
    }

    /// `xml` getter: closes every open node and returns the document.
    pub fn into_xml(mut self) -> String {
        self.close_all();
        self.xml
    }

    /// Records a failure where ExcelJS's render would throw (first wins).
    pub fn fail(&mut self, message: impl Into<String>) {
        if self.error.is_none() {
            self.error = Some(message.into());
        }
    }

    /// `xml` getter, or the recorded failure.
    pub fn finish(self) -> Result<String, String> {
        match self.error.clone() {
            Some(e) => Err(e),
            None => Ok(self.into_xml()),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn matches_exceljs_output() {
        let mut x = XmlStream::new();
        x.open_xml(&std_doc_attributes());
        x.open_node("a", &crate::attrs![("x", Some("1&2")), ("y", None::<&str>)]);
        x.leaf_node("b", &[], None);
        x.leaf_node("c", &[], Some("0"));
        x.add_rollback();
        x.open_node("d", &[]);
        x.rollback();
        assert_eq!(
            x.into_xml(),
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<a x=\"1&amp;2\"><b/><c>0</c></a>"
        );
    }
}
