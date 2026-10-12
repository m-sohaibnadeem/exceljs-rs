//! ExcelJS `lib/xlsx/xform/drawing/*.js`: `xl/drawings/drawingN.xml`.

use serde_json::{json, Map, Value};

use super::base::{BoxX, Children, StaticNode, StaticXform};
use super::simple::{SimpleOptions, ValueXform};
use super::{XResult, Xform};
use crate::attrs;
use crate::jsv::{get, to_js_string, truthy};
use crate::xml::Tag;
use crate::xml_stream::{std_doc_attributes, XmlStream};

type OpenFn = fn(&Tag<'_>) -> Option<Value>;
type ChildCloseFn = fn(&mut Option<Value>, &str, Option<Value>) -> XResult<()>;
type TagCloseFn = fn(&Children, &mut Option<Value>);
type RenderFn = fn(&Container, &mut XmlStream, Option<&Value>);

/// The common drawing-xform shape: on its own tag it resets (model `null`,
/// children reset) and optionally sets a model; other elements go to the
/// child of that name; `parseOpen` always returns true; on its own close it
/// builds the model from the children.
pub(crate) struct Container {
    tag: &'static str,
    map: Children,
    open: OpenFn,
    child_close: ChildCloseFn,
    tag_close: TagCloseFn,
    render_fn: RenderFn,
    model: Option<Value>,
}

impl std::fmt::Debug for Container {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "Container({})", self.tag)
    }
}

fn no_child_close(_: &mut Option<Value>, _: &str, _: Option<Value>) -> XResult<()> {
    Ok(())
}

fn no_tag_close(_: &Children, _: &mut Option<Value>) {}

fn no_open(_: &Tag<'_>) -> Option<Value> {
    None
}

impl Container {
    fn new(tag: &'static str, children: Vec<(&'static str, BoxX)>, render_fn: RenderFn) -> Self {
        Container {
            tag,
            map: Children::new(children),
            open: no_open,
            child_close: no_child_close,
            tag_close: no_tag_close,
            render_fn,
            model: None,
        }
    }
    fn on_open(mut self, f: OpenFn) -> Self {
        self.open = f;
        self
    }
    fn on_child_close(mut self, f: ChildCloseFn) -> Self {
        self.child_close = f;
        self
    }
    fn on_tag_close(mut self, f: TagCloseFn) -> Self {
        self.tag_close = f;
        self
    }
    fn render_child(&self, name: &str, xml: &mut XmlStream, model: Option<&Value>) {
        if let Some(x) = self.map.get(name) {
            x.render(xml, model);
        }
    }
}

impl Xform for Container {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if self.map.forward_open(node)? {
            return Ok(true);
        }
        if node.name == self.tag {
            self.reset();
            if let Some(m) = (self.open)(node) {
                self.model = Some(m);
            }
        } else if self.map.activate(node.name) {
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
                let child = self.map.parser_model().cloned();
                (self.child_close)(&mut self.model, name, child)?;
                self.map.active = None;
            }
            return Ok(true);
        }
        if name == self.tag {
            (self.tag_close)(&self.map, &mut self.model);
            return Ok(false);
        }
        Ok(true)
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

/// Elements whose content is ignored: `parseOpen` always true, close of
/// their own tag ends them (`CNvPicPrXform`, `ExtLstXform`).
#[derive(Debug)]
struct Opaque {
    tag: &'static str,
    render_fn: fn(&mut XmlStream),
    model: Option<Value>,
}

impl Xform for Opaque {
    fn parse_open(&mut self, _node: &Tag<'_>) -> XResult<bool> {
        Ok(true)
    }
    fn parse_close(&mut self, name: &str) -> XResult<bool> {
        Ok(name != self.tag)
    }
    crate::model_accessors!();
    fn render(&self, xml: &mut XmlStream, _model: Option<&Value>) {
        (self.render_fn)(xml);
    }
    fn reset(&mut self) {
        self.model = Some(Value::Null);
    }
}

/// `BlipXform` (`a:blip`).
#[derive(Debug, Default)]
struct BlipXform {
    model: Option<Value>,
}

impl Xform for BlipXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if node.name == "a:blip" {
            let mut m = Map::new();
            if let Some(r) = node.attr("r:embed") {
                m.insert("rId".into(), json!(r));
            }
            self.model = Some(Value::Object(m));
        }
        Ok(true)
    }
    fn parse_close(&mut self, name: &str) -> XResult<bool> {
        Ok(name != "a:blip")
    }
    crate::model_accessors!();
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        xml.leaf_node(
            "a:blip",
            &attrs![
                (
                    "xmlns:r",
                    Some("http://schemas.openxmlformats.org/officeDocument/2006/relationships")
                ),
                ("r:embed", get(model, "rId")),
                ("cstate", Some("print")),
            ],
            None,
        );
    }
    fn reset(&mut self) {
        self.model = Some(Value::Null);
    }
}

/// `HLinkClickXform` (`a:hlinkClick`).
#[derive(Debug, Default)]
struct HLinkClickXform {
    model: Option<Value>,
}

impl Xform for HLinkClickXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if node.name == "a:hlinkClick" {
            let mut h = Map::new();
            if let Some(r) = node.attr("r:id") {
                h.insert("rId".into(), json!(r));
            }
            if let Some(t) = node.attr("tooltip") {
                h.insert("tooltip".into(), json!(t));
            }
            self.model = Some(json!({"hyperlinks": h}));
        }
        Ok(true)
    }
    fn parse_close(&mut self, _name: &str) -> XResult<bool> {
        Ok(false)
    }
    crate::model_accessors!();
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        let h = get(model, "hyperlinks");
        if !(truthy(h) && truthy(get(h, "rId"))) {
            return;
        }
        xml.leaf_node(
            "a:hlinkClick",
            &attrs![
                (
                    "xmlns:r",
                    Some("http://schemas.openxmlformats.org/officeDocument/2006/relationships")
                ),
                ("r:id", get(h, "rId")),
                ("tooltip", get(h, "tooltip")),
            ],
            None,
        );
    }
    fn reset(&mut self) {
        self.model = Some(Value::Null);
    }
}

/// `xdr:ext`.
#[derive(Debug, Default)]
struct ExtXform {
    model: Option<Value>,
}

const EMU_PER_PIXEL_AT_96_DPI: f64 = 9525.0;

impl Xform for ExtXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if node.name == "xdr:ext" {
            let n = |a: &str| {
                crate::js::parse_int(node.attr(a).filter(|v| !v.is_empty()).unwrap_or("0"))
                    .unwrap_or(f64::NAN)
            };
            self.model = Some(json!({
                "width": crate::jsv::num(n("cx") / EMU_PER_PIXEL_AT_96_DPI),
                "height": crate::jsv::num(n("cy") / EMU_PER_PIXEL_AT_96_DPI),
            }));
            return Ok(true);
        }
        Ok(false)
    }
    fn parse_close(&mut self, _name: &str) -> XResult<bool> {
        Ok(false)
    }
    crate::model_accessors!();
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        xml.open_node("xdr:ext", &[]);
        let w = (crate::jsv::to_number(get(model, "width")) * EMU_PER_PIXEL_AT_96_DPI).floor();
        let h = (crate::jsv::to_number(get(model, "height")) * EMU_PER_PIXEL_AT_96_DPI).floor();
        xml.add_attribute("cx", Some(w));
        xml.add_attribute("cy", Some(h));
        xml.close_node();
    }
    fn reset(&mut self) {
        self.model = Some(Value::Null);
    }
}

fn int_zero(tag: &'static str) -> BoxX {
    Box::new(ValueXform::integer_zero(SimpleOptions::tag(tag)))
}

fn cell_position(tag: &'static str) -> Container {
    Container::new(
        tag,
        vec![
            ("xdr:col", int_zero("xdr:col")),
            ("xdr:colOff", int_zero("xdr:colOff")),
            ("xdr:row", int_zero("xdr:row")),
            ("xdr:rowOff", int_zero("xdr:rowOff")),
        ],
        |c, xml, m| {
            xml.open_node(c.tag, &[]);
            c.render_child("xdr:col", xml, get(m, "nativeCol"));
            c.render_child("xdr:colOff", xml, get(m, "nativeColOff"));
            c.render_child("xdr:row", xml, get(m, "nativeRow"));
            c.render_child("xdr:rowOff", xml, get(m, "nativeRowOff"));
            xml.close_node();
        },
    )
    .on_tag_close(|map, model| {
        let mut m = Map::new();
        for (key, child) in [
            ("nativeCol", "xdr:col"),
            ("nativeColOff", "xdr:colOff"),
            ("nativeRow", "xdr:row"),
            ("nativeRowOff", "xdr:rowOff"),
        ] {
            if let Some(v) = map.model_of(child) {
                m.insert(key.into(), v.clone());
            }
        }
        *model = Some(Value::Object(m));
    })
}

fn sp_pr() -> StaticXform {
    StaticXform::new(StaticNode::new(
        "xdr:spPr",
        vec![],
        vec![
            StaticNode::new(
                "a:xfrm",
                vec![],
                vec![
                    StaticNode::new("a:off", attrs![("x", Some("0")), ("y", Some("0"))], vec![]),
                    StaticNode::new(
                        "a:ext",
                        attrs![("cx", Some("0")), ("cy", Some("0"))],
                        vec![],
                    ),
                ],
            ),
            StaticNode::new(
                "a:prstGeom",
                attrs![("prst", Some("rect"))],
                vec![StaticNode::new("a:avLst", vec![], vec![])],
            ),
        ],
    ))
}

fn c_nv_pr() -> Container {
    Container::new(
        "xdr:cNvPr",
        vec![
            ("a:hlinkClick", Box::new(HLinkClickXform::default())),
            (
                "a:extLst",
                Box::new(Opaque {
                    tag: "a:extLst",
                    render_fn: |xml| {
                        xml.open_node("a:extLst", &[]);
                        xml.open_node(
                            "a:ext",
                            &attrs![("uri", Some("{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"))],
                        );
                        xml.leaf_node(
                            "a16:creationId",
                            &attrs![
                                (
                                    "xmlns:a16",
                                    Some("http://schemas.microsoft.com/office/drawing/2014/main")
                                ),
                                ("id", Some("{00000000-0008-0000-0000-000002000000}")),
                            ],
                            None,
                        );
                        xml.close_node();
                        xml.close_node();
                    },
                    model: None,
                }),
            ),
        ],
        |c, xml, m| {
            let index = get(m, "index");
            xml.open_node(
                "xdr:cNvPr",
                &attrs![
                    ("id", index),
                    ("name", Some(format!("Picture {}", to_js_string(index))))
                ],
            );
            c.render_child("a:hlinkClick", xml, m);
            c.render_child("a:extLst", xml, m);
            xml.close_node();
        },
    )
    .on_tag_close(|map, model| *model = map.model_of("a:hlinkClick").cloned())
}

fn nv_pic_pr() -> Container {
    Container::new(
        "xdr:nvPicPr",
        vec![
            ("xdr:cNvPr", Box::new(c_nv_pr())),
            (
                "xdr:cNvPicPr",
                Box::new(Opaque {
                    tag: "xdr:cNvPicPr",
                    render_fn: |xml| {
                        xml.open_node("xdr:cNvPicPr", &[]);
                        xml.leaf_node("a:picLocks", &attrs![("noChangeAspect", Some("1"))], None);
                        xml.close_node();
                    },
                    model: None,
                }),
            ),
        ],
        |c, xml, m| {
            xml.open_node("xdr:nvPicPr", &[]);
            c.render_child("xdr:cNvPr", xml, m);
            c.render_child("xdr:cNvPicPr", xml, m);
            xml.close_node();
        },
    )
    .on_tag_close(|map, model| *model = map.model_of("xdr:cNvPr").cloned())
}

fn blip_fill() -> Container {
    Container::new(
        "xdr:blipFill",
        vec![("a:blip", Box::new(BlipXform::default()))],
        |c, xml, m| {
            xml.open_node("xdr:blipFill", &[]);
            c.render_child("a:blip", xml, m);
            xml.open_node("a:stretch", &[]);
            xml.leaf_node("a:fillRect", &[], None);
            xml.close_node();
            xml.close_node();
        },
    )
    .on_tag_close(|map, model| *model = map.model_of("a:blip").cloned())
}

fn pic() -> Container {
    Container::new(
        "xdr:pic",
        vec![
            ("xdr:nvPicPr", Box::new(nv_pic_pr())),
            ("xdr:blipFill", Box::new(blip_fill())),
            ("xdr:spPr", Box::new(sp_pr())),
        ],
        |c, xml, m| {
            xml.open_node("xdr:pic", &[]);
            c.render_child("xdr:nvPicPr", xml, m);
            c.render_child("xdr:blipFill", xml, m);
            c.render_child("xdr:spPr", xml, m);
            xml.close_node();
        },
    )
    .on_child_close(|model, _name, child| {
        // mergeModel: this.model = Object.assign(this.model || {}, obj)
        if !truthy(model.as_ref()) {
            *model = Some(Value::Object(Map::new()));
        }
        if let (Some(Value::Object(m)), Some(Value::Object(c))) = (model.as_mut(), child) {
            for (k, v) in c {
                m.insert(k, v);
            }
        }
        Ok(())
    })
}

fn client_data() -> StaticXform {
    StaticXform::new(StaticNode::new("xdr:clientData", vec![], vec![]))
}

fn anchor_open(n: &Tag<'_>) -> Option<Value> {
    Some(
        json!({"range": {"editAs": n.attr("editAs").filter(|v| !v.is_empty()).unwrap_or("oneCell")}}),
    )
}

fn set_range(model: &mut Option<Value>, key: &str, v: Option<&Value>) {
    if let (Some(range), Some(v)) = (
        model
            .as_mut()
            .and_then(|m| m.get_mut("range"))
            .and_then(Value::as_object_mut),
        v,
    ) {
        range.insert(key.into(), v.clone());
    }
}

fn two_cell_anchor() -> Container {
    Container::new(
        "xdr:twoCellAnchor",
        vec![
            ("xdr:from", Box::new(cell_position("xdr:from"))),
            ("xdr:to", Box::new(cell_position("xdr:to"))),
            ("xdr:pic", Box::new(pic())),
            ("xdr:clientData", Box::new(client_data())),
        ],
        |c, xml, m| {
            let range = get(m, "range");
            let edit = if truthy(get(range, "editAs")) {
                to_js_string(get(range, "editAs"))
            } else {
                "oneCell".into()
            };
            xml.open_node("xdr:twoCellAnchor", &attrs![("editAs", Some(edit))]);
            c.render_child("xdr:from", xml, get(range, "tl"));
            c.render_child("xdr:to", xml, get(range, "br"));
            c.render_child("xdr:pic", xml, get(m, "picture"));
            c.render_child("xdr:clientData", xml, Some(&json!({})));
            xml.close_node();
        },
    )
    .on_open(anchor_open)
    .on_tag_close(|map, model| {
        set_range(model, "tl", map.model_of("xdr:from"));
        set_range(model, "br", map.model_of("xdr:to"));
        if let (Some(Value::Object(m)), Some(p)) = (model.as_mut(), map.model_of("xdr:pic")) {
            m.insert("picture".into(), p.clone());
        }
    })
}

fn one_cell_anchor() -> Container {
    Container::new(
        "xdr:oneCellAnchor",
        vec![
            ("xdr:from", Box::new(cell_position("xdr:from"))),
            ("xdr:ext", Box::new(ExtXform::default())),
            ("xdr:pic", Box::new(pic())),
            ("xdr:clientData", Box::new(client_data())),
        ],
        |c, xml, m| {
            let range = get(m, "range");
            let edit = if truthy(get(range, "editAs")) {
                to_js_string(get(range, "editAs"))
            } else {
                "oneCell".into()
            };
            xml.open_node("xdr:oneCellAnchor", &attrs![("editAs", Some(edit))]);
            c.render_child("xdr:from", xml, get(range, "tl"));
            c.render_child("xdr:ext", xml, get(range, "ext"));
            c.render_child("xdr:pic", xml, get(m, "picture"));
            c.render_child("xdr:clientData", xml, Some(&json!({})));
            xml.close_node();
        },
    )
    .on_open(anchor_open)
    .on_tag_close(|map, model| {
        set_range(model, "tl", map.model_of("xdr:from"));
        set_range(model, "ext", map.model_of("xdr:ext"));
        if let (Some(Value::Object(m)), Some(p)) = (model.as_mut(), map.model_of("xdr:pic")) {
            m.insert("picture".into(), p.clone());
        }
    })
}

/// `DrawingXform` (`xdr:wsDr`).
pub(crate) fn drawing() -> Container {
    Container::new(
        "xdr:wsDr",
        vec![
            ("xdr:twoCellAnchor", Box::new(two_cell_anchor())),
            ("xdr:oneCellAnchor", Box::new(one_cell_anchor())),
        ],
        |c, xml, m| {
            xml.open_xml(&std_doc_attributes());
            xml.open_node(
                "xdr:wsDr",
                &attrs![
                    (
                        "xmlns:xdr",
                        Some("http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")
                    ),
                    (
                        "xmlns:a",
                        Some("http://schemas.openxmlformats.org/drawingml/2006/main")
                    ),
                ],
            );
            for anchor in get(m, "anchors")
                .and_then(Value::as_array)
                .into_iter()
                .flatten()
            {
                let kind = to_js_string(get(Some(anchor), "anchorType"));
                c.render_child(&kind, xml, Some(anchor));
            }
            xml.close_node();
        },
    )
    .on_open(|_| Some(json!({"anchors": []})))
    .on_child_close(|model, _name, child| {
        match model
            .as_mut()
            .and_then(|m| m.get_mut("anchors"))
            .and_then(Value::as_array_mut)
        {
            Some(a) => a.push(child.unwrap_or(Value::Null)),
            None => return Err("Cannot read properties of undefined (reading 'anchors')".into()),
        }
        Ok(())
    })
}

/// `getAnchorType(model)` (writer).
pub(crate) fn anchor_type(anchor: &Value) -> Result<&'static str, String> {
    let range = get(Some(anchor), "range");
    let has_br = match range {
        Some(Value::String(s)) => crate::col_cache::decode_range(s)?.is_some(),
        other => truthy(get(other, "br")),
    };
    Ok(if has_br {
        "xdr:twoCellAnchor"
    } else {
        "xdr:oneCellAnchor"
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xform::{parse_part, to_xml};

    #[test]
    fn parses_two_cell_anchor() {
        let xml = r#"<xdr:wsDr><xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Picture 1"><a:hlinkClick r:id="rId2" tooltip="t"/></xdr:cNvPr><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill></xdr:pic><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"#;
        let m = parse_part(&mut drawing(), "d", xml.as_bytes())
            .unwrap()
            .unwrap();
        assert_eq!(
            m,
            json!({"anchors": [{
                "range": {"editAs": "oneCell", "tl": {"nativeCol": 1, "nativeColOff": 0, "nativeRow": 2, "nativeRowOff": 0},
                          "br": {"nativeCol": 3, "nativeColOff": 0, "nativeRow": 4, "nativeRowOff": 0}},
                "picture": {"hyperlinks": {"rId": "rId2", "tooltip": "t"}, "rId": "rId1"},
            }]})
        );
        let mut anchor = m["anchors"][0].clone();
        anchor["anchorType"] = json!("xdr:twoCellAnchor");
        anchor["picture"]["index"] = json!(1);
        let out = to_xml(&drawing(), Some(&json!({"anchors": [anchor]})));
        assert!(
            out.contains("<xdr:cNvPr id=\"1\" name=\"Picture 1\"><a:hlinkClick"),
            "{out}"
        );
    }
}
