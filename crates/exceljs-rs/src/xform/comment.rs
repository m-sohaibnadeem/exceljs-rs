//! ExcelJS `lib/xlsx/xform/comment/*.js`: `xl/commentsN.xml` and the VML
//! drawing that positions each note.

use serde_json::{json, Map, Value};

use super::strings::RichTextXform;
use super::{XResult, Xform};
use crate::attrs;
use crate::jsv::{get, to_fixed, to_js_string, truthy};
use crate::xml::Tag;
use crate::xml_stream::{std_doc_attributes, XmlStream};

/// `CommentXform` (`<comment>`).
#[derive(Debug)]
struct CommentXform {
    rich: RichTextXform,
    in_rich: bool,
    model: Option<Value>,
}

impl Xform for CommentXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if self.in_rich {
            self.rich.parse_open(node)?;
            return Ok(true);
        }
        match node.name {
            "comment" => {
                let mut m = Map::new();
                m.insert("type".into(), json!("note"));
                m.insert("note".into(), json!({"texts": []}));
                for (k, v) in node.attributes() {
                    m.insert(k.to_owned(), json!(v));
                }
                self.model = Some(Value::Object(m));
                Ok(true)
            }
            "r" => {
                self.in_rich = true;
                self.rich.parse_open(node)?;
                Ok(true)
            }
            _ => Ok(false),
        }
    }
    fn parse_text(&mut self, text: &str) {
        if self.in_rich {
            self.rich.parse_text(text);
        }
    }
    fn parse_close(&mut self, name: &str) -> XResult<bool> {
        match name {
            "comment" => Ok(false),
            "r" => {
                let run = self.rich.model().cloned().unwrap_or(Value::Null);
                match self
                    .model
                    .as_mut()
                    .and_then(|m| m.pointer_mut("/note/texts"))
                    .and_then(Value::as_array_mut)
                {
                    Some(t) => t.push(run),
                    None => {
                        return Err("Cannot read properties of undefined (reading 'texts')".into())
                    }
                }
                self.in_rich = false;
                Ok(true)
            }
            _ => {
                if self.in_rich {
                    self.rich.parse_close(name)?;
                }
                Ok(true)
            }
        }
    }
    crate::model_accessors!();
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        xml.open_node(
            "comment",
            &attrs![("ref", get(model, "ref")), ("authorId", Some(0i64))],
        );
        xml.open_node("text", &[]);
        let texts = get(get(model, "note"), "texts");
        if truthy(model) && truthy(get(model, "note")) && truthy(texts) {
            for t in texts.and_then(Value::as_array).into_iter().flatten() {
                self.rich.render(xml, Some(t));
            }
        }
        xml.close_node();
        xml.close_node();
    }
}

/// `CommentsXform`.
#[derive(Debug)]
pub(crate) struct CommentsXform {
    comment: CommentXform,
    in_comment: bool,
    model: Option<Value>,
}

impl CommentsXform {
    pub fn new() -> Self {
        CommentsXform {
            comment: CommentXform {
                rich: RichTextXform::new(),
                in_rich: false,
                model: None,
            },
            in_comment: false,
            model: None,
        }
    }
}

impl Xform for CommentsXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if self.in_comment {
            self.comment.parse_open(node)?;
            return Ok(true);
        }
        match node.name {
            "commentList" => {
                self.model = Some(json!({"comments": []}));
                Ok(true)
            }
            "comment" => {
                self.in_comment = true;
                self.comment.parse_open(node)?;
                Ok(true)
            }
            _ => Ok(false),
        }
    }
    fn parse_text(&mut self, text: &str) {
        if self.in_comment {
            self.comment.parse_text(text);
        }
    }
    fn parse_close(&mut self, name: &str) -> XResult<bool> {
        match name {
            "commentList" => Ok(false),
            "comment" => {
                if !self.in_comment {
                    return Err("Cannot read properties of undefined (reading 'model')".into());
                }
                let c = self.comment.model().cloned().unwrap_or(Value::Null);
                match self
                    .model
                    .as_mut()
                    .and_then(|m| m.get_mut("comments"))
                    .and_then(Value::as_array_mut)
                {
                    Some(list) => list.push(c),
                    None => {
                        return Err(
                            "Cannot read properties of undefined (reading 'comments')".into()
                        )
                    }
                }
                self.in_comment = false;
                Ok(true)
            }
            _ => {
                if self.in_comment {
                    self.comment.parse_close(name)?;
                }
                Ok(true)
            }
        }
    }
    crate::model_accessors!();
}

/// `CommentsXform.render`.
pub(crate) fn render_comments(model: &Value) -> String {
    let mut xml = XmlStream::new();
    xml.open_xml(&std_doc_attributes());
    xml.open_node(
        "comments",
        &attrs![(
            "xmlns",
            Some("http://schemas.openxmlformats.org/spreadsheetml/2006/main")
        )],
    );
    xml.open_node("authors", &[]);
    xml.leaf_node("author", &[], Some("Author"));
    xml.close_node();
    xml.open_node("commentList", &[]);
    let cx = CommentsXform::new();
    for c in get(Some(model), "comments")
        .and_then(Value::as_array)
        .into_iter()
        .flatten()
    {
        cx.comment.render(&mut xml, Some(c));
    }
    xml.close_node();
    xml.close_node();
    xml.into_xml()
}

// ---------------------------------------------------------------------------
// VML

/// `VmlTextboxXform.reverseConversionUnit(inset)`.
fn reverse_conversion_unit(inset: Option<&str>) -> Value {
    let parts: Vec<Value> = inset
        .unwrap_or("")
        .split(',')
        .map(|margin| {
            let v = crate::js::parse_float(margin);
            // `${parseFloat(value) * 0.1.toFixed(2)}` -> parseFloat -> toFixed(2) -> Number
            let product = v * 0.1;
            let s = crate::jsv::number_to_string(product);
            let reparsed = crate::js::parse_float(&s);
            crate::jsv::num(crate::jsv::string_to_number(&to_fixed(reparsed, 2)))
        })
        .collect();
    Value::Array(parts)
}

/// `conversionUnit(value, 10, 'mm')`.
fn conversion_unit_mm(v: &Value) -> String {
    let x = crate::js::parse_float(&to_js_string(Some(v)));
    format!("{}mm", crate::jsv::number_to_string(x * 10.0))
}

#[derive(Debug, Default)]
struct VmlTextbox {
    model: Option<Value>,
}

impl Xform for VmlTextbox {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if node.name == "v:textbox" {
            self.model = Some(json!({"inset": reverse_conversion_unit(node.attr("inset"))}));
        }
        Ok(true)
    }
    fn parse_close(&mut self, name: &str) -> XResult<bool> {
        Ok(name != "v:textbox")
    }
    crate::model_accessors!();
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        let mut a = attrs![("style", Some("mso-direction-alt:auto"))];
        if truthy(model) && truthy(get(model, "note")) {
            let inset = get(get(get(model, "note"), "margins"), "inset");
            let inset = match inset {
                Some(Value::Array(items)) => Some(
                    items
                        .iter()
                        .map(conversion_unit_mm)
                        .collect::<Vec<_>>()
                        .join(","),
                ),
                Some(v) if truthy(Some(v)) => Some(to_js_string(Some(v))),
                _ => None,
            };
            if let Some(i) = inset.filter(|i| !i.is_empty()) {
                a.push(("inset".into(), Some(i)));
            }
        }
        xml.open_node("v:textbox", &a);
        xml.leaf_node("div", &attrs![("style", Some("text-align:left"))], None);
        xml.close_node();
    }
    fn reset(&mut self) {
        self.model = Some(Value::Null);
    }
}

/// `x:Anchor`, `x:Locked`, `x:LockText`: text in `.text`, which `reset()`
/// does not clear.
#[derive(Debug)]
struct VmlText {
    tag: &'static str,
    text: Option<String>,
    model: Option<Value>,
}

impl Xform for VmlText {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if node.name == self.tag {
            self.text = Some(String::new());
            return Ok(true);
        }
        Ok(false)
    }
    fn parse_text(&mut self, text: &str) {
        self.text = Some(text.to_owned());
    }
    fn parse_close(&mut self, _name: &str) -> XResult<bool> {
        Ok(false)
    }
    crate::model_accessors!();
    fn reset(&mut self) {
        self.model = Some(Value::Null);
    }
}

/// `x:SizeWithCells` / `x:MoveWithCells`.
#[derive(Debug)]
struct VmlPosition {
    tag: &'static str,
    model: Option<Value>,
}

impl Xform for VmlPosition {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if node.name == self.tag {
            let mut m = Map::new();
            m.insert(self.tag.into(), json!(true));
            self.model = Some(Value::Object(m));
            return Ok(true);
        }
        Ok(false)
    }
    fn parse_close(&mut self, _name: &str) -> XResult<bool> {
        Ok(false)
    }
    crate::model_accessors!();
    fn reset(&mut self) {
        self.model = Some(Value::Null);
    }
}

const POSITION_TYPE: [&str; 3] = ["twoCells", "oneCells", "absolute"];

#[derive(Debug)]
struct VmlClientData {
    anchor: VmlText,
    locked: VmlText,
    lock_text: VmlText,
    size: VmlPosition,
    moves: VmlPosition,
    active: Option<u8>,
    model: Option<Value>,
}

impl VmlClientData {
    fn new() -> Self {
        VmlClientData {
            anchor: VmlText {
                tag: "x:Anchor",
                text: None,
                model: None,
            },
            locked: VmlText {
                tag: "x:Locked",
                text: None,
                model: None,
            },
            lock_text: VmlText {
                tag: "x:LockText",
                text: None,
                model: None,
            },
            size: VmlPosition {
                tag: "x:SizeWithCells",
                model: None,
            },
            moves: VmlPosition {
                tag: "x:MoveWithCells",
                model: None,
            },
            active: None,
            model: None,
        }
    }
    fn child(&mut self, i: u8) -> &mut dyn Xform {
        match i {
            0 => &mut self.anchor,
            1 => &mut self.locked,
            2 => &mut self.lock_text,
            3 => &mut self.size,
            _ => &mut self.moves,
        }
    }
}

impl Xform for VmlClientData {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if node.name == "x:ClientData" {
            self.reset();
            self.model = Some(json!({"anchor": [], "protection": {}, "editAs": ""}));
            return Ok(true);
        }
        self.active = match node.name {
            "x:Anchor" => Some(0),
            "x:Locked" => Some(1),
            "x:LockText" => Some(2),
            "x:SizeWithCells" => Some(3),
            "x:MoveWithCells" => Some(4),
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
        if name != "x:ClientData" {
            return Ok(true);
        }
        // normalizeModel
        let mut position = Map::new();
        for m in [self.moves.model(), self.size.model()]
            .into_iter()
            .flatten()
        {
            if let Value::Object(o) = m {
                for (k, v) in o {
                    position.insert(k.clone(), v.clone());
                }
            }
        }
        let edit_as = POSITION_TYPE
            .get(position.len())
            .map_or(Value::Null, |s| json!(s));
        let text = |t: &Option<String>| t.as_ref().map(|s| json!(s));
        if let Some(Value::Object(m)) = self.model.as_mut() {
            if POSITION_TYPE.get(position.len()).is_some() {
                m.insert("editAs".into(), edit_as);
            } else {
                m.remove("editAs");
            }
            match text(&self.anchor.text) {
                Some(a) => m.insert("anchor".into(), a),
                None => m.remove("anchor"),
            };
            if let Some(Value::Object(p)) = m.get_mut("protection") {
                match text(&self.locked.text) {
                    Some(v) => p.insert("locked".into(), v),
                    None => p.remove("locked"),
                };
                match text(&self.lock_text.text) {
                    Some(v) => p.insert("lockText".into(), v),
                    None => p.remove("lockText"),
                };
            }
        }
        Ok(false)
    }
    crate::model_accessors!();
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        let note = get(model, "note");
        let protection = get(note, "protection");
        let edit_as = get(note, "editAs");
        xml.open_node("x:ClientData", &attrs![("ObjectType", Some("Note"))]);
        let is = |v: Option<&Value>, s: &str| crate::jsv::as_str(v) == Some(s);
        if is(edit_as, POSITION_TYPE[2]) {
            xml.leaf_node("x:MoveWithCells", &[], None);
        }
        if is(edit_as, POSITION_TYPE[2]) || is(edit_as, POSITION_TYPE[1]) {
            xml.leaf_node("x:SizeWithCells", &[], None);
        }
        // x:Anchor
        let rect: Vec<f64> = match get(model, "anchor").filter(|a| truthy(Some(a))) {
            Some(anchor) => {
                let n = |k: &str| crate::jsv::to_number(get(Some(anchor), k));
                let (l, t, r, b) = (n("left"), n("top"), n("right"), n("bottom"));
                vec![
                    l.floor(),
                    ((l - l.floor()) * 68.0).floor(),
                    t.floor(),
                    ((t - t.floor()) * 18.0).floor(),
                    r.floor(),
                    ((r - r.floor()) * 68.0).floor(),
                    b.floor(),
                    ((b - b.floor()) * 18.0).floor(),
                ]
            }
            None => {
                let ref_addr = get(model, "refAddress");
                let col = crate::jsv::to_number(get(ref_addr, "col"));
                let row = crate::jsv::to_number(get(ref_addr, "row"));
                let t = (row - 2.0).max(0.0);
                vec![col, 6.0, t, 14.0, col + 2.0, 2.0, t + 4.0, 16.0]
            }
        };
        let rect_s = rect
            .iter()
            .map(|v| crate::jsv::number_to_string(*v))
            .collect::<Vec<_>>()
            .join(", ");
        xml.leaf_node("x:Anchor", &[], Some(&rect_s));
        let text_of = |v: Option<&Value>| v.map(|v| to_js_string(Some(v)));
        xml.leaf_node(
            "x:Locked",
            &[],
            text_of(get(protection, "locked")).as_deref(),
        );
        xml.leaf_node("x:AutoFill", &[], Some("False"));
        xml.leaf_node(
            "x:LockText",
            &[],
            text_of(get(protection, "lockText")).as_deref(),
        );
        let ref_addr = get(model, "refAddress");
        let row = crate::jsv::to_number(get(ref_addr, "row")) - 1.0;
        let col = crate::jsv::to_number(get(ref_addr, "col")) - 1.0;
        xml.leaf_node("x:Row", &[], Some(&crate::jsv::number_to_string(row)));
        xml.leaf_node("x:Column", &[], Some(&crate::jsv::number_to_string(col)));
        xml.close_node();
    }
    fn reset(&mut self) {
        self.model = Some(Value::Null);
        self.anchor.reset();
        self.locked.reset();
        self.lock_text.reset();
        self.size.reset();
        self.moves.reset();
    }
}

#[derive(Debug)]
struct VmlShape {
    textbox: VmlTextbox,
    client: VmlClientData,
    active: Option<bool>, // true = textbox
    model: Option<Value>,
}

impl Xform for VmlShape {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        match self.active {
            Some(true) => {
                self.textbox.parse_open(node)?;
                return Ok(true);
            }
            Some(false) => {
                self.client.parse_open(node)?;
                return Ok(true);
            }
            None => {}
        }
        match node.name {
            "v:shape" => {
                self.reset();
                let mut margins = Map::new();
                if let Some(i) = node.attr("o:insetmode") {
                    margins.insert("insetmode".into(), json!(i));
                }
                self.model =
                    Some(json!({"margins": margins, "anchor": "", "editAs": "", "protection": {}}));
            }
            "v:textbox" => {
                self.active = Some(true);
                self.textbox.parse_open(node)?;
            }
            "x:ClientData" => {
                self.active = Some(false);
                self.client.parse_open(node)?;
            }
            _ => {}
        }
        Ok(true)
    }
    fn parse_text(&mut self, text: &str) {
        match self.active {
            Some(true) => self.textbox.parse_text(text),
            Some(false) => self.client.parse_text(text),
            None => {}
        }
    }
    fn parse_close(&mut self, name: &str) -> XResult<bool> {
        if let Some(is_tb) = self.active {
            let done = if is_tb {
                !self.textbox.parse_close(name)?
            } else {
                !self.client.parse_close(name)?
            };
            if done {
                self.active = None;
            }
            return Ok(true);
        }
        if name != "v:shape" {
            return Ok(true);
        }
        let tb = self.textbox.model().filter(|v| truthy(Some(v))).cloned();
        let cd = self.client.model().filter(|v| truthy(Some(v))).cloned();
        let m = match self.model.as_mut() {
            Some(Value::Object(m)) => m,
            _ => return Err("Cannot read properties of undefined (reading 'margins')".into()),
        };
        // `x && x.prop` yields x itself when x is falsy (null).
        let pick = |src: &Option<Value>, raw: Option<&Value>, key: &str| -> Option<Value> {
            match src {
                Some(v) => v.get(key).cloned(),
                None => raw.cloned(),
            }
        };
        let inset = pick(&tb, self.textbox.model(), "inset");
        if let Some(Value::Object(margins)) = m.get_mut("margins") {
            match inset {
                Some(v) => margins.insert("inset".into(), v),
                None => margins.remove("inset"),
            };
        }
        for (key, src) in [
            ("protection", "protection"),
            ("anchor", "anchor"),
            ("editAs", "editAs"),
        ] {
            match pick(&cd, self.client.model(), src) {
                Some(v) => m.insert(key.into(), v),
                None => m.remove(key),
            };
        }
        Ok(false)
    }
    crate::model_accessors!();
    fn reset(&mut self) {
        self.model = Some(Value::Null);
        self.textbox.reset();
        self.client.reset();
    }
}

/// `VmlNotesXform` (`xl/drawings/vmlDrawingN.vml`), parse role.
#[derive(Debug)]
pub(crate) struct VmlNotesXform {
    shape: VmlShape,
    in_shape: bool,
    model: Option<Value>,
}

impl VmlNotesXform {
    pub fn new() -> Self {
        VmlNotesXform {
            shape: VmlShape {
                textbox: VmlTextbox::default(),
                client: VmlClientData::new(),
                active: None,
                model: None,
            },
            in_shape: false,
            model: None,
        }
    }
}

impl Xform for VmlNotesXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if self.in_shape {
            self.shape.parse_open(node)?;
            return Ok(true);
        }
        if node.name == "xml" {
            self.model = Some(json!({"comments": []}));
            self.shape.reset();
        } else if node.name == "v:shape" {
            self.in_shape = true;
            self.shape.parse_open(node)?;
        }
        Ok(true)
    }
    fn parse_text(&mut self, text: &str) {
        if self.in_shape {
            self.shape.parse_text(text);
        }
    }
    fn parse_close(&mut self, name: &str) -> XResult<bool> {
        if self.in_shape {
            if !self.shape.parse_close(name)? {
                let s = self.shape.model().cloned().unwrap_or(Value::Null);
                match self
                    .model
                    .as_mut()
                    .and_then(|m| m.get_mut("comments"))
                    .and_then(Value::as_array_mut)
                {
                    Some(c) => c.push(s),
                    None => {
                        return Err(
                            "Cannot read properties of undefined (reading 'comments')".into()
                        )
                    }
                }
                self.in_shape = false;
            }
            return Ok(true);
        }
        Ok(name != "xml")
    }
    crate::model_accessors!();
}

/// `VmlNotesXform.render`.
pub(crate) fn render_vml_notes(model: &Value) -> String {
    let mut xml = XmlStream::new();
    xml.open_xml(&std_doc_attributes());
    xml.open_node(
        "xml",
        &attrs![
            ("xmlns:v", Some("urn:schemas-microsoft-com:vml")),
            ("xmlns:o", Some("urn:schemas-microsoft-com:office:office")),
            ("xmlns:x", Some("urn:schemas-microsoft-com:office:excel")),
        ],
    );
    xml.open_node("o:shapelayout", &attrs![("v:ext", Some("edit"))]);
    xml.leaf_node(
        "o:idmap",
        &attrs![("v:ext", Some("edit")), ("data", Some(1i64))],
        None,
    );
    xml.close_node();
    xml.open_node(
        "v:shapetype",
        &attrs![
            ("id", Some("_x0000_t202")),
            ("coordsize", Some("21600,21600")),
            ("o:spt", Some(202i64)),
            ("path", Some("m,l,21600r21600,l21600,xe")),
        ],
    );
    xml.leaf_node("v:stroke", &attrs![("joinstyle", Some("miter"))], None);
    xml.leaf_node(
        "v:path",
        &attrs![
            ("gradientshapeok", Some("t")),
            ("o:connecttype", Some("rect"))
        ],
        None,
    );
    xml.close_node();
    let textbox = VmlTextbox::default();
    let client = VmlClientData::new();
    for (index, item) in get(Some(model), "comments")
        .and_then(Value::as_array)
        .into_iter()
        .flatten()
        .enumerate()
    {
        render_vml_shape(&mut xml, item, index, &textbox, &client);
    }
    xml.close_node();
    xml.into_xml()
}

/// `VmlShapeXform.render(xmlStream, model, index)`.
fn render_vml_shape(
    xml: &mut XmlStream,
    item: &Value,
    index: usize,
    textbox: &VmlTextbox,
    client: &VmlClientData,
) {
    let margins = get(get(Some(item), "note"), "margins");
    xml.open_node(
        "v:shape",
        &attrs![
            ("id", Some(format!("_x0000_s{}", 1025 + index))),
            ("type", Some("#_x0000_t202")),
            (
                "style",
                Some("position:absolute; margin-left:105.3pt;margin-top:10.5pt;width:97.8pt;height:59.1pt;z-index:1;visibility:hidden")
            ),
            ("fillcolor", Some("infoBackground [80]")),
            ("strokecolor", Some("none [81]")),
            ("o:insetmode", if truthy(margins) { get(margins, "insetmode").cloned() } else { None }),
        ],
    );
    xml.leaf_node(
        "v:fill",
        &attrs![("color2", Some("infoBackground [80]"))],
        None,
    );
    xml.leaf_node(
        "v:shadow",
        &attrs![("color", Some("none [81]")), ("obscured", Some("t"))],
        None,
    );
    xml.leaf_node("v:path", &attrs![("o:connecttype", Some("none"))], None);
    textbox.render(xml, Some(item));
    client.render(xml, Some(item));
    xml.close_node();
}

/// One `<comment>` and its `<v:shape>` (the streaming writer's
/// `SheetCommentsWriter._writeComment`).
pub(crate) fn render_comment_pair(comment: &Value, index: usize) -> (String, String) {
    let mut c = XmlStream::new();
    CommentsXform::new().comment.render(&mut c, Some(comment));
    let mut v = XmlStream::new();
    render_vml_shape(
        &mut v,
        comment,
        index,
        &VmlTextbox::default(),
        &VmlClientData::new(),
    );
    (c.into_xml(), v.into_xml())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xform::parse_part;

    #[test]
    fn insets_match_exceljs() {
        assert_eq!(
            reverse_conversion_unit(Some("0.13cm,0.13cm,0.25cm,0.25cm")),
            json!([0.01, 0.01, 0.03, 0.03])
        );
        assert_eq!(reverse_conversion_unit(None), json!([{"$num": "NaN"}]));
    }

    #[test]
    fn comments() {
        let xml = r#"<comments><authors><author>A</author></authors><commentList><comment ref="B1" authorId="0"><text><r><rPr><b/></rPr><t>hi</t></r></text></comment></commentList></comments>"#;
        let m = parse_part(&mut CommentsXform::new(), "c", xml.as_bytes())
            .unwrap()
            .unwrap();
        assert_eq!(
            m,
            json!({"comments": [{"type": "note", "note": {"texts": [{"font": {"bold": true}, "text": "hi"}]}, "ref": "B1", "authorId": "0"}]})
        );
    }
}
