//! ExcelJS `lib/xlsx/xform/core/*.js`: relationships, content types,
//! docProps/app.xml and docProps/core.xml.

use serde_json::{Map, Value};

use super::simple::{DateFormat, DateXform, SimpleOptions, ValueXform};
use super::{XResult, Xform};
use crate::attrs;
use crate::jsv::{get, to_js_string, truthy};
use crate::xml::Tag;
use crate::xml_stream::{std_doc_attributes, Attrs, XmlStream};

fn obj_attrs(model: Option<&Value>) -> Attrs {
    match model {
        Some(Value::Object(o)) => o
            .iter()
            .map(|(k, v)| (k.clone(), Some(to_js_string(Some(v)))))
            .collect(),
        _ => Vec::new(),
    }
}

/// `RelationshipsXform` (parse role): returns the list of attribute objects.
#[derive(Debug, Default)]
pub(crate) struct RelationshipsXform {
    rels: Option<Vec<Value>>,
    depth: u32,
}

impl Xform for RelationshipsXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if self.depth > 0 {
            // RelationshipXform.parseOpen for a nested node: returns false.
            return Ok(true);
        }
        match node.name {
            "Relationships" => {
                self.rels = Some(Vec::new());
                Ok(true)
            }
            "Relationship" => {
                self.depth = 1;
                let rel = node.attributes_object();
                match self.rels.as_mut() {
                    Some(r) => r.push(rel),
                    None => {
                        return Err("Cannot read properties of undefined (reading 'push')".into())
                    }
                }
                Ok(true)
            }
            other => Err(format!("Unexpected xml node in parseOpen: {other}")),
        }
    }
    fn parse_close(&mut self, name: &str) -> XResult<bool> {
        if self.depth > 0 {
            // RelationshipXform.parseClose always returns false.
            self.depth = 0;
            return Ok(true);
        }
        match name {
            "Relationships" => Ok(false),
            other => Err(format!("Unexpected xml node in parseClose: {other}")),
        }
    }
    fn model(&self) -> Option<&Value> {
        None
    }
    fn take_model(&mut self) -> Option<Value> {
        self.rels.take().map(Value::Array)
    }
}

/// `RelationshipsXform.render`.
pub(crate) fn render_relationships(rels: &[Value]) -> String {
    let mut xml = XmlStream::new();
    xml.open_xml(&std_doc_attributes());
    xml.open_node(
        "Relationships",
        &attrs![(
            "xmlns",
            Some("http://schemas.openxmlformats.org/package/2006/relationships")
        )],
    );
    for rel in rels {
        xml.leaf_node("Relationship", &obj_attrs(Some(rel)), None);
    }
    xml.close_node();
    xml.into_xml()
}

/// `ContentTypesXform.render`.
pub(crate) fn render_content_types(model: &Value) -> String {
    let m = Some(model);
    let mut xml = XmlStream::new();
    xml.open_xml(&std_doc_attributes());
    xml.open_node(
        "Types",
        &attrs![(
            "xmlns",
            Some("http://schemas.openxmlformats.org/package/2006/content-types")
        )],
    );
    let mut seen = std::collections::HashSet::new();
    for medium in get(m, "media")
        .and_then(Value::as_array)
        .into_iter()
        .flatten()
    {
        if crate::jsv::as_str(get(Some(medium), "type")) == Some("image") {
            let ext = to_js_string(get(Some(medium), "extension"));
            // `if (!mediaHash[imageType])`
            if seen.insert(ext.clone()) {
                xml.leaf_node(
                    "Default",
                    &attrs![
                        ("Extension", Some(&ext)),
                        ("ContentType", Some(format!("image/{ext}")))
                    ],
                    None,
                );
            }
        }
    }
    let leaf = |xml: &mut XmlStream, tag: &str, a: (&str, String), ct: &str| {
        xml.leaf_node(
            tag,
            &attrs![(a.0, Some(a.1)), ("ContentType", Some(ct))],
            None,
        );
    };
    leaf(
        &mut xml,
        "Default",
        ("Extension", "rels".into()),
        "application/vnd.openxmlformats-package.relationships+xml",
    );
    leaf(
        &mut xml,
        "Default",
        ("Extension", "xml".into()),
        "application/xml",
    );
    leaf(
        &mut xml,
        "Override",
        ("PartName", "/xl/workbook.xml".into()),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    );
    for ws in get(m, "worksheets")
        .and_then(Value::as_array)
        .into_iter()
        .flatten()
    {
        leaf(
            &mut xml,
            "Override",
            (
                "PartName",
                format!(
                    "/xl/worksheets/sheet{}.xml",
                    to_js_string(get(Some(ws), "id"))
                ),
            ),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
        );
    }
    leaf(
        &mut xml,
        "Override",
        ("PartName", "/xl/theme/theme1.xml".into()),
        "application/vnd.openxmlformats-officedocument.theme+xml",
    );
    leaf(
        &mut xml,
        "Override",
        ("PartName", "/xl/styles.xml".into()),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
    );
    if truthy(get(m, "sharedStrings")) && truthy(get(get(m, "sharedStrings"), "count")) {
        leaf(
            &mut xml,
            "Override",
            ("PartName", "/xl/sharedStrings.xml".into()),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
        );
    }
    if truthy(get(m, "tables")) {
        for t in get(m, "tables")
            .and_then(Value::as_array)
            .into_iter()
            .flatten()
        {
            leaf(
                &mut xml,
                "Override",
                (
                    "PartName",
                    format!("/xl/tables/{}", to_js_string(get(Some(t), "target"))),
                ),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml",
            );
        }
    }
    if truthy(get(m, "drawings")) {
        for d in get(m, "drawings")
            .and_then(Value::as_array)
            .into_iter()
            .flatten()
        {
            leaf(
                &mut xml,
                "Override",
                (
                    "PartName",
                    format!("/xl/drawings/{}.xml", to_js_string(get(Some(d), "name"))),
                ),
                "application/vnd.openxmlformats-officedocument.drawing+xml",
            );
        }
    }
    if truthy(get(m, "commentRefs")) {
        leaf(
            &mut xml,
            "Default",
            ("Extension", "vml".into()),
            "application/vnd.openxmlformats-officedocument.vmlDrawing",
        );
        for c in get(m, "commentRefs")
            .and_then(Value::as_array)
            .into_iter()
            .flatten()
        {
            leaf(
                &mut xml,
                "Override",
                (
                    "PartName",
                    format!("/xl/{}.xml", to_js_string(get(Some(c), "commentName"))),
                ),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml",
            );
        }
    }
    leaf(
        &mut xml,
        "Override",
        ("PartName", "/docProps/core.xml".into()),
        "application/vnd.openxmlformats-package.core-properties+xml",
    );
    leaf(
        &mut xml,
        "Override",
        ("PartName", "/docProps/app.xml".into()),
        "application/vnd.openxmlformats-officedocument.extended-properties+xml",
    );
    xml.close_node();
    xml.into_xml()
}

/// `AppXform` (docProps/app.xml).
#[derive(Debug)]
pub(crate) struct AppXform {
    company: ValueXform,
    manager: ValueXform,
    /// Inside HeadingPairs / TitlesOfParts (ignored; ExcelJS's map key
    /// `TitleOfParts` never matches the element `TitlesOfParts`).
    active: Option<&'static str>,
    model: Option<Value>,
}

impl AppXform {
    pub fn new() -> Self {
        AppXform {
            company: ValueXform::string(SimpleOptions::tag("Company")),
            manager: ValueXform::string(SimpleOptions::tag("Manager")),
            active: None,
            model: None,
        }
    }
}

impl Xform for AppXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        match self.active {
            Some("Company") => {
                self.company.parse_open(node)?;
                return Ok(true);
            }
            Some("Manager") => {
                self.manager.parse_open(node)?;
                return Ok(true);
            }
            Some(_) => return Ok(true),
            None => {}
        }
        match node.name {
            "Properties" => Ok(true),
            "Company" => {
                self.active = Some("Company");
                self.company.parse_open(node)?;
                Ok(true)
            }
            "Manager" => {
                self.active = Some("Manager");
                self.manager.parse_open(node)?;
                Ok(true)
            }
            "HeadingPairs" => {
                self.active = Some("HeadingPairs");
                Ok(true)
            }
            _ => Ok(false),
        }
    }
    fn parse_text(&mut self, text: &str) {
        match self.active {
            Some("Company") => self.company.parse_text(text),
            Some("Manager") => self.manager.parse_text(text),
            _ => {}
        }
    }
    fn parse_close(&mut self, name: &str) -> XResult<bool> {
        if let Some(a) = self.active {
            let go_on = match a {
                "Company" => self.company.parse_close(name)?,
                "Manager" => self.manager.parse_close(name)?,
                _ => name != "HeadingPairs",
            };
            if !go_on {
                self.active = None;
            }
            return Ok(true);
        }
        if name == "Properties" {
            let mut m = Map::new();
            if let Some(c) = self.company.model() {
                m.insert("company".into(), c.clone());
            }
            if let Some(c) = self.manager.model() {
                m.insert("manager".into(), c.clone());
            }
            self.model = Some(Value::Object(m));
            return Ok(false);
        }
        Ok(true)
    }
    crate::model_accessors!();
}

/// `AppXform.render`.
pub(crate) fn render_app(model: &Value) -> String {
    let m = Some(model);
    let mut xml = XmlStream::new();
    xml.open_xml(&std_doc_attributes());
    xml.open_node(
        "Properties",
        &attrs![
            (
                "xmlns",
                Some("http://schemas.openxmlformats.org/officeDocument/2006/extended-properties")
            ),
            (
                "xmlns:vt",
                Some("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")
            ),
        ],
    );
    xml.leaf_node("Application", &[], Some("Microsoft Excel"));
    xml.leaf_node("DocSecurity", &[], Some("0"));
    xml.leaf_node("ScaleCrop", &[], Some("false"));
    let sheets: Vec<&Value> = get(m, "worksheets")
        .and_then(Value::as_array)
        .into_iter()
        .flatten()
        .collect();
    // HeadingPairs
    xml.open_node("HeadingPairs", &[]);
    xml.open_node(
        "vt:vector",
        &attrs![("size", Some(2i64)), ("baseType", Some("variant"))],
    );
    xml.open_node("vt:variant", &[]);
    xml.leaf_node("vt:lpstr", &[], Some("Worksheets"));
    xml.close_node();
    xml.open_node("vt:variant", &[]);
    xml.leaf_node("vt:i4", &[], Some(&sheets.len().to_string()));
    xml.close_node();
    xml.close_node();
    xml.close_node();
    // TitlesOfParts
    xml.open_node("TitlesOfParts", &[]);
    xml.open_node(
        "vt:vector",
        &attrs![("size", Some(sheets.len())), ("baseType", Some("lpstr"))],
    );
    for s in &sheets {
        xml.leaf_node("vt:lpstr", &[], Some(&to_js_string(get(Some(s), "name"))));
    }
    xml.close_node();
    xml.close_node();
    let company = if truthy(get(m, "company")) {
        get(m, "company").cloned()
    } else {
        Some(Value::String(String::new()))
    };
    ValueXform::string(SimpleOptions::tag("Company")).render(&mut xml, company.as_ref());
    ValueXform::string(SimpleOptions::tag("Manager")).render(&mut xml, get(m, "manager"));
    xml.leaf_node("LinksUpToDate", &[], Some("false"));
    xml.leaf_node("SharedDoc", &[], Some("false"));
    xml.leaf_node("HyperlinksChanged", &[], Some("false"));
    xml.leaf_node("AppVersion", &[], Some("16.0300"));
    xml.close_node();
    xml.into_xml()
}

/// The children of `cp:coreProperties`, in ExcelJS's map order, with the
/// model property each one fills.
fn core_children() -> Vec<(&'static str, Option<&'static str>, Box<dyn Xform>)> {
    let s = |tag| -> Box<dyn Xform> { Box::new(ValueXform::string(SimpleOptions::tag(tag))) };
    let date_attrs = || SimpleOptions {
        tag: "",
        attr: None,
        attrs: attrs![("xsi:type", Some("dcterms:W3CDTF"))],
    };
    vec![
        ("dc:creator", Some("creator"), s("dc:creator")),
        ("dc:title", Some("title"), s("dc:title")),
        ("dc:subject", Some("subject"), s("dc:subject")),
        ("dc:description", Some("description"), s("dc:description")),
        ("dc:identifier", Some("identifier"), s("dc:identifier")),
        ("dc:language", Some("language"), s("dc:language")),
        ("cp:keywords", Some("keywords"), s("cp:keywords")),
        ("cp:category", Some("category"), s("cp:category")),
        (
            "cp:lastModifiedBy",
            Some("lastModifiedBy"),
            s("cp:lastModifiedBy"),
        ),
        (
            "cp:lastPrinted",
            Some("lastPrinted"),
            Box::new(DateXform::new(
                SimpleOptions::tag("cp:lastPrinted"),
                DateFormat::Core,
            )),
        ),
        (
            "cp:revision",
            Some("revision"),
            Box::new(ValueXform::integer(SimpleOptions::tag("cp:revision"))),
        ),
        // Parsed but not part of the model.
        ("cp:version", None, s("cp:version")),
        (
            "cp:contentStatus",
            Some("contentStatus"),
            s("cp:contentStatus"),
        ),
        ("cp:contentType", Some("contentType"), s("cp:contentType")),
        (
            "dcterms:created",
            Some("created"),
            Box::new(DateXform::new(
                SimpleOptions {
                    tag: "dcterms:created",
                    ..date_attrs()
                },
                DateFormat::Core,
            )),
        ),
        (
            "dcterms:modified",
            Some("modified"),
            Box::new(DateXform::new(
                SimpleOptions {
                    tag: "dcterms:modified",
                    ..date_attrs()
                },
                DateFormat::Core,
            )),
        ),
    ]
}

/// `CoreXform` (docProps/core.xml).
pub(crate) struct CoreXform {
    children: Vec<(&'static str, Option<&'static str>, Box<dyn Xform>)>,
    active: Option<usize>,
    model: Option<Value>,
}

impl std::fmt::Debug for CoreXform {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str("CoreXform")
    }
}

impl CoreXform {
    pub fn new() -> Self {
        CoreXform {
            children: core_children(),
            active: None,
            model: None,
        }
    }
}

impl Xform for CoreXform {
    fn parse_open(&mut self, node: &Tag<'_>) -> XResult<bool> {
        if let Some((_, _, x)) = self.active.and_then(|i| self.children.get_mut(i)) {
            x.parse_open(node)?;
            return Ok(true);
        }
        match node.name {
            "cp:coreProperties" | "coreProperties" => Ok(true),
            name => {
                self.active = self.children.iter().position(|(t, _, _)| *t == name);
                match self.active.and_then(|i| self.children.get_mut(i)) {
                    Some((_, _, x)) => {
                        x.parse_open(node)?;
                        Ok(true)
                    }
                    None => Err(format!("Unexpected xml node in parseOpen: {name}")),
                }
            }
        }
    }
    fn parse_text(&mut self, text: &str) {
        if let Some((_, _, x)) = self.active.and_then(|i| self.children.get_mut(i)) {
            x.parse_text(text);
        }
    }
    fn parse_close(&mut self, name: &str) -> XResult<bool> {
        if let Some((_, _, x)) = self.active.and_then(|i| self.children.get_mut(i)) {
            if !x.parse_close(name)? {
                self.active = None;
            }
            return Ok(true);
        }
        match name {
            "cp:coreProperties" | "coreProperties" => {
                let mut m = Map::new();
                // ExcelJS's model order differs from the map order.
                for prop in [
                    "creator",
                    "title",
                    "subject",
                    "description",
                    "identifier",
                    "language",
                    "keywords",
                    "category",
                    "lastModifiedBy",
                    "lastPrinted",
                    "revision",
                    "contentStatus",
                    "contentType",
                    "created",
                    "modified",
                ] {
                    if let Some((_, _, x)) = self.children.iter().find(|(_, p, _)| *p == Some(prop))
                    {
                        if let Some(v) = x.model() {
                            m.insert(prop.into(), v.clone());
                        }
                    }
                }
                self.model = Some(Value::Object(m));
                Ok(false)
            }
            other => Err(format!("Unexpected xml node in parseClose: {other}")),
        }
    }
    crate::model_accessors!();
}

/// `CoreXform.render`.
pub(crate) fn render_core(model: &Value) -> Result<String, String> {
    let m = Some(model);
    let mut xml = XmlStream::new();
    xml.open_xml(&std_doc_attributes());
    xml.open_node(
        "cp:coreProperties",
        &attrs![
            (
                "xmlns:cp",
                Some("http://schemas.openxmlformats.org/package/2006/metadata/core-properties")
            ),
            ("xmlns:dc", Some("http://purl.org/dc/elements/1.1/")),
            ("xmlns:dcterms", Some("http://purl.org/dc/terms/")),
            ("xmlns:dcmitype", Some("http://purl.org/dc/dcmitype/")),
            (
                "xmlns:xsi",
                Some("http://www.w3.org/2001/XMLSchema-instance")
            ),
        ],
    );
    let children = core_children();
    for (tag, prop) in [
        ("dc:creator", "creator"),
        ("dc:title", "title"),
        ("dc:subject", "subject"),
        ("dc:description", "description"),
        ("dc:identifier", "identifier"),
        ("dc:language", "language"),
        ("cp:keywords", "keywords"),
        ("cp:category", "category"),
        ("cp:lastModifiedBy", "lastModifiedBy"),
        ("cp:lastPrinted", "lastPrinted"),
        ("cp:revision", "revision"),
        ("cp:version", "version"),
        ("cp:contentStatus", "contentStatus"),
        ("cp:contentType", "contentType"),
        ("dcterms:created", "created"),
        ("dcterms:modified", "modified"),
    ] {
        if let Some((_, _, x)) = children.iter().find(|(t, _, _)| *t == tag) {
            x.render(&mut xml, get(m, prop));
        }
    }
    xml.close_node();
    xml.finish()
}
