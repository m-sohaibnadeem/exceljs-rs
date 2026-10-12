//! ExcelJS `lib/xlsx/xform/sheet/cf/*.js` and `cf-ext/*.js`: conditional
//! formatting, both the plain `<conditionalFormatting>` elements and the x14
//! extension inside `<extLst>`.

use serde_json::{json, Value};

use super::composite::{
    merge_into, object, push_to, to_attribute, to_bool_attribute, to_bool_value, to_float_value,
    to_int_value, to_string_value, Composite, Leaf, LeafMatch, TextAccum,
};
use super::style::ColorXform;
use super::{XResult, Xform};
use crate::attrs;
use crate::col_cache::range_tl;
use crate::jsv::{as_str, get, to_js_string, truthy};
use crate::xml::Tag;
use crate::xml_stream::XmlStream;

const EXT_ICONS: [&str; 3] = ["3Triangles", "3Stars", "5Boxes"];

fn is_ext_icon(rule: Option<&Value>) -> bool {
    truthy(get(rule, "custom"))
        || as_str(get(rule, "iconSet")).is_some_and(|s| EXT_ICONS.contains(&s))
}

/// `CfRuleXform.isPrimitive(rule)`.
pub(crate) fn is_primitive(rule: Option<&Value>) -> bool {
    !(as_str(get(rule, "type")) == Some("iconSet") && is_ext_icon(rule))
}

/// `CfRuleExtXform.isExt(rule)`.
pub(crate) fn is_ext(rule: Option<&Value>) -> bool {
    match as_str(get(rule, "type")) {
        Some("dataBar") => !truthy(get(rule, "gradient")),
        Some("iconSet") => is_ext_icon(rule),
        _ => false,
    }
}

fn arr(v: Option<&Value>) -> impl Iterator<Item = &Value> {
    v.and_then(Value::as_array).into_iter().flatten()
}

// ---------------------------------------------------------------------------
// Plain conditional formatting

fn cfvo() -> Leaf {
    Leaf::new(
        "cfvo",
        LeafMatch::AnyUntilClose,
        |n| {
            Some(object(vec![
                ("type", n.attr("type").map(|t| json!(t))),
                ("value", to_float_value(n.attr("val"), None)),
            ]))
        },
        |xml, m| {
            xml.leaf_node(
                "cfvo",
                &attrs![("type", get(m, "type")), ("val", get(m, "value"))],
                None,
            )
        },
    )
}

fn color_scale() -> Composite {
    Composite::new(
        "colorScale",
        vec![
            ("cfvo", Box::new(cfvo())),
            ("color", Box::new(ColorXform::new("color"))),
        ],
        |_| json!({"cfvo": [], "color": []}),
        push_to,
        |x, xml, m| {
            xml.open_node("colorScale", &[]);
            for c in arr(get(m, "cfvo")) {
                x.render_child("cfvo", xml, Some(c));
            }
            for c in arr(get(m, "color")) {
                x.render_child("color", xml, Some(c));
            }
            xml.close_node();
        },
    )
}

fn icon_set() -> Composite {
    Composite::new(
        "iconSet",
        vec![("cfvo", Box::new(cfvo()))],
        |n| {
            object(vec![
                (
                    "iconSet",
                    to_string_value(n.attr("iconSet"), Some("3TrafficLights")),
                ),
                ("reverse", to_bool_value(n.attr("reverse"), None)),
                ("showValue", to_bool_value(n.attr("showValue"), None)),
                ("cfvo", Some(json!([]))),
            ])
        },
        push_to,
        |x, xml, m| {
            xml.open_node(
                "iconSet",
                &attrs![
                    (
                        "iconSet",
                        to_attribute(get(m, "iconSet"), Some(json!("3TrafficLights")), false)
                    ),
                    (
                        "reverse",
                        to_bool_attribute(get(m, "reverse"), false, false)
                    ),
                    (
                        "showValue",
                        to_bool_attribute(get(m, "showValue"), true, false)
                    ),
                ],
            );
            for c in arr(get(m, "cfvo")) {
                x.render_child("cfvo", xml, Some(c));
            }
            xml.close_node();
        },
    )
}

fn databar() -> Composite {
    Composite::new(
        "dataBar",
        vec![
            ("cfvo", Box::new(cfvo())),
            ("color", Box::new(ColorXform::new("color"))),
        ],
        |_| json!({"cfvo": []}),
        |m, name, c| match name {
            "cfvo" => push_to(m, "cfvo", c),
            "color" => {
                if let (Value::Object(o), Some(c)) = (m, c) {
                    o.insert("color".into(), c);
                }
                Ok(())
            }
            _ => Ok(()),
        },
        |x, xml, m| {
            xml.open_node("dataBar", &[]);
            for c in arr(get(m, "cfvo")) {
                x.render_child("cfvo", xml, Some(c));
            }
            x.render_child("color", xml, get(m, "color"));
            xml.close_node();
        },
    )
}

fn ext_lst_ref() -> Composite {
    let ext = Composite::new(
        "ext",
        vec![("x14:id", Box::new(TextAccum::new("x14:id")))],
        |_| json!({}),
        |m, _name, c| {
            if let (Value::Object(o), Some(c)) = (m, c) {
                o.insert("x14Id".into(), c);
            }
            Ok(())
        },
        |x, xml, m| {
            xml.open_node(
                "ext",
                &attrs![
                    ("uri", Some("{B025F937-C7B1-47D3-B67F-A62EFF666E3E}")),
                    (
                        "xmlns:x14",
                        Some("http://schemas.microsoft.com/office/spreadsheetml/2009/9/main")
                    ),
                ],
            );
            x.render_child("x14:id", xml, get(m, "x14Id"));
            xml.close_node();
        },
    );
    Composite::new(
        "extLst",
        vec![("ext", Box::new(ext))],
        |_| json!({}),
        merge_into,
        |x, xml, m| {
            xml.open_node("extLst", &[]);
            x.render_child("ext", xml, m);
            xml.close_node();
        },
    )
}

/// `getTextFormula(model)`.
fn text_formula(m: Option<&Value>) -> Result<Option<String>, String> {
    if let Some(f) = get(get(m, "formulae"), "0").filter(|f| truthy(Some(f))) {
        return Ok(Some(to_js_string(Some(f))));
    }
    let tl = range_tl(&to_js_string(get(m, "ref")))?;
    let text = to_js_string(get(m, "text"));
    Ok(match as_str(get(m, "operator")) {
        Some("containsText") => Some(format!("NOT(ISERROR(SEARCH(\"{text}\",{tl})))")),
        Some("containsBlanks") => Some(format!("LEN(TRIM({tl}))=0")),
        Some("notContainsBlanks") => Some(format!("LEN(TRIM({tl}))>0")),
        Some("containsErrors") => Some(format!("ISERROR({tl})")),
        Some("notContainsErrors") => Some(format!("NOT(ISERROR({tl}))")),
        _ => None,
    })
}

/// `getTimePeriodFormula(model)`.
fn time_period_formula(m: Option<&Value>) -> Result<Option<String>, String> {
    if let Some(f) = get(get(m, "formulae"), "0").filter(|f| truthy(Some(f))) {
        return Ok(Some(to_js_string(Some(f))));
    }
    let tl = range_tl(&to_js_string(get(m, "ref")))?;
    Ok(match as_str(get(m, "timePeriod")) {
        Some("thisWeek") => Some(format!("AND(TODAY()-ROUNDDOWN({tl},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN({tl},0)-TODAY()<=7-WEEKDAY(TODAY()))")),
        Some("lastWeek") => Some(format!("AND(TODAY()-ROUNDDOWN({tl},0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN({tl},0)<(WEEKDAY(TODAY())+7))")),
        Some("nextWeek") => Some(format!("AND(ROUNDDOWN({tl},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN({tl},0)-TODAY()<(15-WEEKDAY(TODAY())))")),
        Some("yesterday") => Some(format!("FLOOR({tl},1)=TODAY()-1")),
        Some("today") => Some(format!("FLOOR({tl},1)=TODAY()")),
        Some("tomorrow") => Some(format!("FLOOR({tl},1)=TODAY()+1")),
        Some("last7Days") => Some(format!("AND(TODAY()-FLOOR({tl},1)<=6,FLOOR({tl},1)<=TODAY())")),
        Some("lastMonth") => Some(format!("AND(MONTH({tl})=MONTH(EDATE(TODAY(),0-1)),YEAR({tl})=YEAR(EDATE(TODAY(),0-1)))")),
        Some("thisMonth") => Some(format!("AND(MONTH({tl})=MONTH(TODAY()),YEAR({tl})=YEAR(TODAY()))")),
        Some("nextMonth") => Some(format!("AND(MONTH({tl})=MONTH(EDATE(TODAY(),0+1)),YEAR({tl})=YEAR(EDATE(TODAY(),0+1)))")),
        _ => None,
    })
}

fn cf_rule() -> Composite {
    Composite::new(
        "cfRule",
        vec![
            ("dataBar", Box::new(databar())),
            ("extLst", Box::new(ext_lst_ref())),
            ("formula", Box::new(TextAccum::new("formula"))),
            ("colorScale", Box::new(color_scale())),
            ("iconSet", Box::new(icon_set())),
        ],
        |n| {
            let t = n.attr("type");
            let (ty, op) = match t {
                Some(
                    t @ ("containsText" | "containsBlanks" | "notContainsBlanks" | "containsErrors"
                    | "notContainsErrors"),
                ) => (Some(json!("containsText")), Some(json!(t))),
                _ => (t.map(|t| json!(t)), n.attr("operator").map(|o| json!(o))),
            };
            object(vec![
                ("type", ty),
                ("operator", op),
                ("dxfId", to_int_value(n.attr("dxfId"), None)),
                ("priority", to_int_value(n.attr("priority"), None)),
                ("timePeriod", n.attr("timePeriod").map(|t| json!(t))),
                ("percent", to_bool_value(n.attr("percent"), None)),
                ("bottom", to_bool_value(n.attr("bottom"), None)),
                ("rank", to_int_value(n.attr("rank"), None)),
                ("aboveAverage", to_bool_value(n.attr("aboveAverage"), None)),
            ])
        },
        |m, name, c| match name {
            "dataBar" | "extLst" | "colorScale" | "iconSet" => merge_into(m, name, c),
            "formula" => {
                if let Value::Object(o) = m {
                    let entry = o.entry("formulae").or_insert_with(|| json!([]));
                    if !truthy(Some(entry)) {
                        *entry = json!([]);
                    }
                    if let Value::Array(a) = entry {
                        a.push(c.unwrap_or(Value::Null));
                    }
                }
                Ok(())
            }
            _ => Ok(()),
        },
        render_cf_rule,
    )
}

fn render_cf_rule(x: &Composite, xml: &mut XmlStream, m: Option<&Value>) {
    let common = |ty: &str, with_dxf: bool| {
        let mut a = attrs![("type", Some(ty))];
        if with_dxf {
            a.push((
                "dxfId".into(),
                get(m, "dxfId").map(|v| to_js_string(Some(v))),
            ));
        }
        a.push((
            "priority".into(),
            get(m, "priority").map(|v| to_js_string(Some(v))),
        ));
        a
    };
    match as_str(get(m, "type")) {
        Some("expression") => {
            xml.open_node("cfRule", &common("expression", true));
            x.render_child("formula", xml, get(get(m, "formulae"), "0"));
            xml.close_node();
        }
        Some("cellIs") => {
            let mut a = common("cellIs", true);
            a.push((
                "operator".into(),
                get(m, "operator").map(|v| to_js_string(Some(v))),
            ));
            xml.open_node("cfRule", &a);
            for f in arr(get(m, "formulae")) {
                x.render_child("formula", xml, Some(f));
            }
            xml.close_node();
        }
        Some("top10") => {
            let mut a = common("top10", true);
            a.push((
                "percent".into(),
                to_bool_attribute(get(m, "percent"), false, false),
            ));
            a.push((
                "bottom".into(),
                to_bool_attribute(get(m, "bottom"), false, false),
            ));
            // BaseXform.toIntValue(model.rank, 10, true): rank === undefined ? 10 : parseInt(rank)
            let rank = match get(m, "rank") {
                None => Some("10".to_owned()),
                Some(r) => Some(crate::jsv::number_to_string(
                    crate::js::parse_int(&to_js_string(Some(r))).unwrap_or(f64::NAN),
                )),
            };
            a.push(("rank".into(), rank));
            xml.leaf_node("cfRule", &a, None);
        }
        Some("aboveAverage") => {
            let mut a = common("aboveAverage", true);
            a.push((
                "aboveAverage".into(),
                to_bool_attribute(get(m, "aboveAverage"), true, false),
            ));
            xml.leaf_node("cfRule", &a, None);
        }
        Some("dataBar") => {
            xml.open_node("cfRule", &common("dataBar", false));
            x.render_child("dataBar", xml, m);
            x.render_child("extLst", xml, m);
            xml.close_node();
        }
        Some("colorScale") => {
            xml.open_node("cfRule", &common("colorScale", false));
            x.render_child("colorScale", xml, m);
            xml.close_node();
        }
        Some("iconSet") => {
            if !is_primitive(m) {
                return;
            }
            xml.open_node("cfRule", &common("iconSet", false));
            x.render_child("iconSet", xml, m);
            xml.close_node();
        }
        Some("containsText") => {
            let op = get(m, "operator").map(|v| to_js_string(Some(v)));
            let mut a = attrs![("type", op.clone())];
            a.push((
                "dxfId".into(),
                get(m, "dxfId").map(|v| to_js_string(Some(v))),
            ));
            a.push((
                "priority".into(),
                get(m, "priority").map(|v| to_js_string(Some(v))),
            ));
            a.push((
                "operator".into(),
                to_attribute(get(m, "operator"), Some(json!("containsText")), false),
            ));
            xml.open_node("cfRule", &a);
            match text_formula(m) {
                Ok(Some(f)) => x.render_child("formula", xml, Some(&json!(f))),
                Ok(None) => {}
                Err(e) => xml.fail(e),
            }
            xml.close_node();
        }
        Some("timePeriod") => {
            let mut a = common("timePeriod", true);
            a.push((
                "timePeriod".into(),
                get(m, "timePeriod").map(|v| to_js_string(Some(v))),
            ));
            xml.open_node("cfRule", &a);
            match time_period_formula(m) {
                Ok(Some(f)) => x.render_child("formula", xml, Some(&json!(f))),
                Ok(None) => {}
                Err(e) => xml.fail(e),
            }
            xml.close_node();
        }
        _ => {}
    }
}

fn conditional_formatting() -> Composite {
    Composite::new(
        "conditionalFormatting",
        vec![("cfRule", Box::new(cf_rule()))],
        |n| {
            object(vec![
                ("ref", n.attr("sqref").map(|s| json!(s))),
                ("rules", Some(json!([]))),
            ])
        },
        |m, _name, c| push_to(m, "rules", c),
        |x, xml, m| {
            if !arr(get(m, "rules")).any(|r| is_primitive(Some(r))) {
                return;
            }
            xml.open_node("conditionalFormatting", &attrs![("sqref", get(m, "ref"))]);
            for rule in arr(get(m, "rules")) {
                if is_primitive(Some(rule)) {
                    let mut rule = rule.clone();
                    if let (Value::Object(o), Some(r)) = (&mut rule, get(m, "ref")) {
                        o.insert("ref".into(), r.clone());
                    }
                    x.render_child("cfRule", xml, Some(&rule));
                }
            }
            xml.close_node();
        },
    )
}

/// `ConditionalFormattingsXform`: `model` is the array of every
/// `<conditionalFormatting>` in the sheet (it accumulates across elements).
#[derive(Debug)]
pub(crate) struct ConditionalFormattingsXform {
    cf: Composite,
    in_cf: bool,
    model: Option<Value>,
}

impl ConditionalFormattingsXform {
    pub fn new() -> Self {
        ConditionalFormattingsXform {
            cf: conditional_formatting(),
            in_cf: false,
            model: None,
        }
    }
}

impl Xform for ConditionalFormattingsXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if self.in_cf {
            self.cf.parse_open(node)?;
            return Ok(true);
        }
        if node.name == "conditionalFormatting" {
            self.in_cf = true;
            self.cf.parse_open(node)?;
            return Ok(true);
        }
        Ok(false)
    }
    fn parse_text(&mut self, text: &str) {
        if self.in_cf {
            self.cf.parse_text(text);
        }
    }
    fn parse_close(&mut self, name: &str) -> XResult<bool> {
        if self.in_cf {
            if !self.cf.parse_close(name)? {
                let item = self.cf.model().cloned().unwrap_or(Value::Null);
                match self.model.as_mut() {
                    Some(Value::Array(a)) => a.push(item),
                    _ => return Err("Cannot read properties of undefined (reading 'push')".into()),
                }
                self.in_cf = false;
                return Ok(false);
            }
            return Ok(true);
        }
        Ok(false)
    }
    crate::model_accessors!();
    fn render(&self, xml: &mut XmlStream, model: Option<&Value>) {
        for cf in arr(model) {
            self.cf.render(xml, Some(cf));
        }
    }
    fn reset(&mut self) {
        self.model = Some(json!([]));
    }
}

// ---------------------------------------------------------------------------
// x14 extension

fn cfvo_ext() -> Composite {
    Composite::new(
        "x14:cfvo",
        vec![("xm:f", Box::new(TextAccum::new("xm:f")))],
        |n| object(vec![("type", n.attr("type").map(|t| json!(t)))]),
        |m, name, c| {
            if name == "xm:f" {
                if let Value::Object(o) = m {
                    let v = if truthy(c.as_ref()) {
                        super::simple::float_value(Some(&to_js_string(c.as_ref())))
                    } else {
                        json!(0)
                    };
                    o.insert("value".into(), v);
                }
            }
            Ok(())
        },
        |x, xml, m| {
            xml.open_node("x14:cfvo", &attrs![("type", get(m, "type"))]);
            if get(m, "value").is_some() {
                x.render_child("xm:f", xml, get(m, "value"));
            }
            xml.close_node();
        },
    )
}

fn cf_icon_ext() -> Leaf {
    Leaf::new(
        "x14:cfIcon",
        LeafMatch::AnyUntilClose,
        |n| {
            Some(object(vec![
                ("iconSet", n.attr("iconSet").map(|s| json!(s))),
                ("iconId", to_int_value(n.attr("iconId"), None)),
            ]))
        },
        |xml, m| {
            xml.leaf_node(
                "x14:cfIcon",
                &attrs![("iconSet", get(m, "iconSet")), ("iconId", get(m, "iconId"))],
                None,
            )
        },
    )
}

/// `onParserClose` of the ext dataBar/iconSet: `name.split(':')[1]`.
fn ext_prop_close(m: &mut Value, name: &str, c: Option<Value>) -> XResult<()> {
    let prop = name.split(':').nth(1).unwrap_or("undefined");
    match prop {
        "cfvo" => push_to(m, "cfvo", c),
        "cfIcon" => {
            if let Value::Object(o) = m {
                let icons = o.entry("icons").or_insert_with(|| json!([]));
                if let Value::Array(a) = icons {
                    a.push(c.unwrap_or(Value::Null));
                }
            }
            Ok(())
        }
        _ => {
            if let (Value::Object(o), Some(c)) = (m, c) {
                o.insert(prop.to_owned(), c);
            }
            Ok(())
        }
    }
}

fn databar_ext() -> Composite {
    Composite::new(
        "x14:dataBar",
        vec![
            ("x14:cfvo", Box::new(cfvo_ext())),
            (
                "x14:borderColor",
                Box::new(ColorXform::new("x14:borderColor")),
            ),
            (
                "x14:negativeBorderColor",
                Box::new(ColorXform::new("x14:negativeBorderColor")),
            ),
            (
                "x14:negativeFillColor",
                Box::new(ColorXform::new("x14:negativeFillColor")),
            ),
            ("x14:axisColor", Box::new(ColorXform::new("x14:axisColor"))),
        ],
        |n| {
            object(vec![
                ("cfvo", Some(json!([]))),
                (
                    "minLength",
                    to_int_value(n.attr("minLength"), Some(json!(0))),
                ),
                (
                    "maxLength",
                    to_int_value(n.attr("maxLength"), Some(json!(100))),
                ),
                ("border", to_bool_value(n.attr("border"), Some(false))),
                ("gradient", to_bool_value(n.attr("gradient"), Some(true))),
                (
                    "negativeBarColorSameAsPositive",
                    to_bool_value(n.attr("negativeBarColorSameAsPositive"), Some(true)),
                ),
                (
                    "negativeBarBorderColorSameAsPositive",
                    to_bool_value(n.attr("negativeBarBorderColorSameAsPositive"), Some(true)),
                ),
                (
                    "axisPosition",
                    to_string_value(n.attr("axisPosition"), Some("auto")),
                ),
                (
                    "direction",
                    to_string_value(n.attr("direction"), Some("leftToRight")),
                ),
            ])
        },
        ext_prop_close,
        |x, xml, m| {
            xml.open_node(
                "x14:dataBar",
                &attrs![
                    (
                        "minLength",
                        to_attribute(get(m, "minLength"), Some(json!(0)), true)
                    ),
                    (
                        "maxLength",
                        to_attribute(get(m, "maxLength"), Some(json!(100)), true)
                    ),
                    ("border", to_bool_attribute(get(m, "border"), false, false)),
                    (
                        "gradient",
                        to_bool_attribute(get(m, "gradient"), true, false)
                    ),
                    (
                        "negativeBarColorSameAsPositive",
                        to_bool_attribute(get(m, "negativeBarColorSameAsPositive"), true, false)
                    ),
                    (
                        "negativeBarBorderColorSameAsPositive",
                        to_bool_attribute(
                            get(m, "negativeBarBorderColorSameAsPositive"),
                            true,
                            false
                        )
                    ),
                    (
                        "axisPosition",
                        to_attribute(get(m, "axisPosition"), Some(json!("auto")), false)
                    ),
                    (
                        "direction",
                        to_attribute(get(m, "direction"), Some(json!("leftToRight")), false)
                    ),
                ],
            );
            for c in arr(get(m, "cfvo")) {
                x.render_child("x14:cfvo", xml, Some(c));
            }
            x.render_child("x14:borderColor", xml, get(m, "borderColor"));
            x.render_child(
                "x14:negativeBorderColor",
                xml,
                get(m, "negativeBorderColor"),
            );
            x.render_child("x14:negativeFillColor", xml, get(m, "negativeFillColor"));
            x.render_child("x14:axisColor", xml, get(m, "axisColor"));
            xml.close_node();
        },
    )
}

fn icon_set_ext() -> Composite {
    Composite::new(
        "x14:iconSet",
        vec![
            ("x14:cfvo", Box::new(cfvo_ext())),
            ("x14:cfIcon", Box::new(cf_icon_ext())),
        ],
        |n| {
            object(vec![
                ("cfvo", Some(json!([]))),
                (
                    "iconSet",
                    to_string_value(n.attr("iconSet"), Some("3TrafficLights")),
                ),
                ("reverse", to_bool_value(n.attr("reverse"), Some(false))),
                ("showValue", to_bool_value(n.attr("showValue"), Some(true))),
            ])
        },
        ext_prop_close,
        |x, xml, m| {
            xml.open_node(
                "x14:iconSet",
                &attrs![
                    ("iconSet", to_attribute(get(m, "iconSet"), None, false)),
                    (
                        "reverse",
                        to_bool_attribute(get(m, "reverse"), false, false)
                    ),
                    (
                        "showValue",
                        to_bool_attribute(get(m, "showValue"), true, false)
                    ),
                    ("custom", to_bool_attribute(get(m, "icons"), false, false)),
                ],
            );
            for c in arr(get(m, "cfvo")) {
                x.render_child("x14:cfvo", xml, Some(c));
            }
            if truthy(get(m, "icons")) {
                for (i, icon) in arr(get(m, "icons")).enumerate() {
                    let mut icon = icon.clone();
                    if let Value::Object(o) = &mut icon {
                        o.insert("iconId".into(), json!(i));
                    }
                    x.render_child("x14:cfIcon", xml, Some(&icon));
                }
            }
            xml.close_node();
        },
    )
}

fn cf_rule_ext() -> Composite {
    Composite::new(
        "x14:cfRule",
        vec![
            ("x14:dataBar", Box::new(databar_ext())),
            ("x14:iconSet", Box::new(icon_set_ext())),
        ],
        |n| {
            object(vec![
                ("type", n.attr("type").map(|t| json!(t))),
                ("x14Id", n.attr("id").map(|t| json!(t))),
                ("priority", to_int_value(n.attr("priority"), None)),
            ])
        },
        merge_into,
        |x, xml, m| {
            if !is_ext(m) {
                return;
            }
            match as_str(get(m, "type")) {
                Some("dataBar") => {
                    xml.open_node(
                        "x14:cfRule",
                        &attrs![("type", Some("dataBar")), ("id", get(m, "x14Id"))],
                    );
                    x.render_child("x14:dataBar", xml, m);
                    xml.close_node();
                }
                Some("iconSet") => {
                    let id = get(m, "x14Id")
                        .filter(|v| truthy(Some(v)))
                        .map_or_else(|| format!("{{{}}}", uuid_v4()), |v| to_js_string(Some(v)));
                    xml.open_node(
                        "x14:cfRule",
                        &attrs![
                            ("type", Some("iconSet")),
                            ("priority", get(m, "priority")),
                            ("id", Some(id))
                        ],
                    );
                    x.render_child("x14:iconSet", xml, m);
                    xml.close_node();
                }
                _ => {}
            }
        },
    )
}

fn conditional_formatting_ext() -> Composite {
    Composite::new(
        "x14:conditionalFormatting",
        vec![
            ("xm:sqref", Box::new(TextAccum::new("xm:sqref"))),
            ("x14:cfRule", Box::new(cf_rule_ext())),
        ],
        |_| json!({"rules": []}),
        |m, name, c| match name {
            "xm:sqref" => {
                if let (Value::Object(o), Some(c)) = (m, c) {
                    o.insert("ref".into(), c);
                }
                Ok(())
            }
            "x14:cfRule" => push_to(m, "rules", c),
            _ => Ok(()),
        },
        |x, xml, m| {
            if !arr(get(m, "rules")).any(|r| is_ext(Some(r))) {
                return;
            }
            xml.open_node(
                "x14:conditionalFormatting",
                &attrs![(
                    "xmlns:xm",
                    Some("http://schemas.microsoft.com/office/excel/2006/main")
                )],
            );
            for rule in arr(get(m, "rules")).filter(|r| is_ext(Some(r))) {
                x.render_child("x14:cfRule", xml, Some(rule));
            }
            x.render_child("xm:sqref", xml, get(m, "ref"));
            xml.close_node();
        },
    )
}

/// `ConditionalFormattingsExt.hasContent(model)`.
pub(crate) fn ext_has_content(model: Option<&Value>) -> bool {
    arr(model).any(|cf| arr(get(Some(cf), "rules")).any(|r| is_ext(Some(r))))
}

fn conditional_formattings_ext() -> Composite {
    Composite::new(
        "x14:conditionalFormattings",
        vec![(
            "x14:conditionalFormatting",
            Box::new(conditional_formatting_ext()),
        )],
        |_| json!([]),
        |m, _name, c| {
            if let Value::Array(a) = m {
                a.push(c.unwrap_or(Value::Null));
            }
            Ok(())
        },
        |x, xml, m| {
            if ext_has_content(m) {
                xml.open_node("x14:conditionalFormattings", &[]);
                for cf in arr(m) {
                    x.render_child("x14:conditionalFormatting", xml, Some(cf));
                }
                xml.close_node();
            }
        },
    )
}

/// The worksheet's `ExtLstXform`.
pub(crate) fn ext_lst() -> Composite {
    let ext = Composite::new(
        "ext",
        vec![(
            "x14:conditionalFormattings",
            Box::new(conditional_formattings_ext()),
        )],
        |_| json!({}),
        |m, name, c| {
            if let (Value::Object(o), Some(c)) = (m, c) {
                o.insert(name.to_owned(), c);
            }
            Ok(())
        },
        |x, xml, m| {
            xml.open_node(
                "ext",
                &attrs![
                    ("uri", Some("{78C0D931-6437-407d-A8EE-F0AAD7539E65}")),
                    (
                        "xmlns:x14",
                        Some("http://schemas.microsoft.com/office/spreadsheetml/2009/9/main")
                    ),
                ],
            );
            x.render_child(
                "x14:conditionalFormattings",
                xml,
                get(m, "conditionalFormattings"),
            );
            xml.close_node();
        },
    );
    Composite::new(
        "extLst",
        vec![("ext", Box::new(ext))],
        |_| json!({}),
        merge_into,
        |x, xml, m| {
            if !ext_has_content(get(m, "conditionalFormattings")) {
                return;
            }
            xml.open_node("extLst", &[]);
            x.render_child("ext", xml, m);
            xml.close_node();
        },
    )
}

/// `mergeConditionalFormattings(model, extModel)` from worksheet-xform.js.
pub(crate) fn merge_conditional_formattings(
    model: Option<Value>,
    ext: Option<Value>,
) -> Option<Value> {
    let ext_list = match &ext {
        Some(Value::Array(a)) if !a.is_empty() => a.clone(),
        _ => return model,
    };
    let mut list = match model {
        Some(Value::Array(a)) if !a.is_empty() => a,
        _ => return ext,
    };
    // cfMap and ruleMap are built once, from the original model.
    let mut cf_map: Vec<(String, usize)> = Vec::new();
    let mut rule_map: Vec<(Value, (usize, usize))> = Vec::new();
    for (ci, cf) in list.iter().enumerate() {
        let key = to_js_string(get(Some(cf), "ref"));
        cf_map.retain(|(k, _)| *k != key);
        cf_map.push((key, ci));
        for (ri, rule) in arr(get(Some(cf), "rules")).enumerate() {
            if let Some(id) = get(Some(rule), "x14Id").filter(|v| truthy(Some(v))) {
                rule_map.retain(|(k, _)| k != id);
                rule_map.push((id.clone(), (ci, ri)));
            }
        }
    }
    for ext_cf in &ext_list {
        for ext_rule in arr(get(Some(ext_cf), "rules")) {
            let id = get(Some(ext_rule), "x14Id").cloned().unwrap_or(Value::Null);
            if let Some((_, (ci, ri))) = rule_map.iter().find(|(k, _)| *k == id) {
                let rule = list
                    .get_mut(*ci)
                    .and_then(|cf| cf.get_mut("rules"))
                    .and_then(|r| r.get_mut(*ri));
                if let (Some(Value::Object(r)), Value::Object(e)) = (rule, ext_rule) {
                    for (k, v) in e {
                        if !r.contains_key(k) {
                            r.insert(k.clone(), v.clone());
                        }
                    }
                }
                continue;
            }
            let ext_ref = get(Some(ext_cf), "ref");
            let key = to_js_string(ext_ref);
            match cf_map.iter().find(|(k, _)| *k == key).map(|(_, ci)| *ci) {
                Some(ci) => {
                    if let Some(Value::Array(rules)) =
                        list.get_mut(ci).and_then(|cf| cf.get_mut("rules"))
                    {
                        rules.push(ext_rule.clone());
                    }
                }
                None => list.push(object(vec![
                    ("ref", ext_ref.cloned()),
                    ("rules", Some(json!([ext_rule]))),
                ])),
            }
        }
    }
    Some(Value::Array(list))
}

/// A version-4 UUID string in upper case (`uuidv4().toUpperCase()` callers
/// upper-case it themselves).
pub(crate) fn uuid_v4() -> String {
    use std::collections::hash_map::RandomState;
    use std::hash::{BuildHasher, Hasher};
    let mut bytes = [0u8; 16];
    for chunk in bytes.chunks_mut(8) {
        let mut h = RandomState::new().build_hasher();
        h.write_u128(
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .map_or(0, |d| d.as_nanos()),
        );
        for (b, v) in chunk.iter_mut().zip(h.finish().to_le_bytes()) {
            *b = v;
        }
    }
    if let Some(b) = bytes.get_mut(6) {
        *b = (*b & 0x0f) | 0x40;
    }
    if let Some(b) = bytes.get_mut(8) {
        *b = (*b & 0x3f) | 0x80;
    }
    let hex: String = bytes.iter().map(|b| format!("{b:02x}")).collect();
    format!(
        "{}-{}-{}-{}-{}",
        hex.get(0..8).unwrap_or(""),
        hex.get(8..12).unwrap_or(""),
        hex.get(12..16).unwrap_or(""),
        hex.get(16..20).unwrap_or(""),
        hex.get(20..32).unwrap_or("")
    )
}
