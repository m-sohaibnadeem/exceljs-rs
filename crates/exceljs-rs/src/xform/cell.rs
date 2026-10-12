//! ExcelJS `lib/xlsx/xform/sheet/cell-xform.js` and `row-xform.js`: the
//! value types and `getValueType`. Cells are parsed by `crate::sheet_data`
//! and prepared/rendered by `crate::rows`, over typed models.

use serde_json::Value;

use crate::jsv::{get, truthy};

/// `Enums.ValueType`.
pub(crate) mod vt {
    pub const NULL: i64 = 0;
    pub const MERGE: i64 = 1;
    pub const NUMBER: i64 = 2;
    pub const STRING: i64 = 3;
    pub const DATE: i64 = 4;
    pub const HYPERLINK: i64 = 5;
    pub const FORMULA: i64 = 6;
    #[allow(dead_code)] // listed for completeness, as in Enums.ValueType
    pub const SHARED_STRING: i64 = 7;
    pub const RICH_TEXT: i64 = 8;
    pub const BOOLEAN: i64 = 9;
    pub const ERROR: i64 = 10;
}

/// `getValueType(v)` from cell-xform.js.
pub(crate) fn value_type(v: Option<&Value>) -> Result<i64, String> {
    match v {
        None | Some(Value::Null) => Ok(vt::NULL),
        Some(Value::String(_)) => Ok(vt::STRING),
        Some(Value::Number(_)) => Ok(vt::NUMBER),
        Some(Value::Bool(_)) => Ok(vt::BOOLEAN),
        Some(Value::Object(o)) if crate::jsv::num_tag(o).is_some() => Ok(vt::NUMBER),
        Some(Value::Object(o)) if crate::model_json::as_date(o).is_some() => Ok(vt::DATE),
        Some(o) => {
            if truthy(get(Some(o), "text")) && truthy(get(Some(o), "hyperlink")) {
                Ok(vt::HYPERLINK)
            } else if truthy(get(Some(o), "formula")) {
                Ok(vt::FORMULA)
            } else if truthy(get(Some(o), "error")) {
                Ok(vt::ERROR)
            } else {
                Err("I could not understand type of value".into())
            }
        }
    }
}
