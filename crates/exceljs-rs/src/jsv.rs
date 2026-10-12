//! JavaScript semantics over `serde_json::Value`.
//!
//! The workbook model exchanged with the Node object model is ExcelJS's own
//! JS object graph. ExcelJS's xforms decide what to write with JS rules
//! (truthiness, `String(x)`, `x === undefined`), so exceljs-rs applies the same
//! rules to the same data. In this module a missing key is JS `undefined` and
//! `Value::Null` is JS `null`.

use serde_json::{Map, Value};

pub(crate) type Obj = Map<String, Value>;

/// JS truthiness (`if (x)`). `None` is `undefined`.
pub(crate) fn truthy(v: Option<&Value>) -> bool {
    match v {
        None | Some(Value::Null) => false,
        Some(Value::Bool(b)) => *b,
        Some(Value::Number(n)) => n.as_f64().is_some_and(|f| f != 0.0 && !f.is_nan()),
        Some(Value::String(s)) => !s.is_empty(),
        Some(Value::Array(_)) => true,
        Some(Value::Object(o)) => num_tag(o).is_none_or(|f| f != 0.0 && !f.is_nan()),
    }
}

/// `{"$num": "NaN" | "Infinity" | "-Infinity"}`: a JS number JSON cannot hold.
pub(crate) fn num_tag(o: &Obj) -> Option<f64> {
    if o.len() != 1 {
        return None;
    }
    match o.get("$num")?.as_str()? {
        "Infinity" => Some(f64::INFINITY),
        "-Infinity" => Some(f64::NEG_INFINITY),
        _ => Some(f64::NAN),
    }
}

/// Whether `v` is a JS number (`typeof v === 'number'`).
pub(crate) fn is_number(v: Option<&Value>) -> bool {
    match v {
        Some(Value::Number(_)) => true,
        Some(Value::Object(o)) => num_tag(o).is_some(),
        _ => false,
    }
}

/// `obj[key]` for an object value; `None` (undefined) otherwise.
pub(crate) fn get<'a>(v: Option<&'a Value>, key: &str) -> Option<&'a Value> {
    match v {
        Some(Value::Object(o)) => o.get(key),
        Some(Value::Array(a)) => key.parse::<usize>().ok().and_then(|i| a.get(i)),
        _ => None,
    }
}

/// `x !== undefined`.
pub(crate) fn defined(v: Option<&Value>) -> bool {
    v.is_some()
}

/// JS `Number(x)` coercion for the value kinds that occur in models.
pub(crate) fn to_number(v: Option<&Value>) -> f64 {
    match v {
        None => f64::NAN,
        Some(Value::Null) => 0.0,
        Some(Value::Bool(b)) => f64::from(u8::from(*b)),
        Some(Value::Number(n)) => n.as_f64().unwrap_or(f64::NAN),
        Some(Value::String(s)) => string_to_number(s),
        Some(Value::Array(a)) => match a.as_slice() {
            [] => 0.0,
            [one] => to_number(Some(one)),
            _ => f64::NAN,
        },
        Some(Value::Object(o)) => num_tag(o).unwrap_or(f64::NAN),
    }
}

/// ECMAScript `StringToNumber` (decimal, hex/octal/binary literals, Infinity).
pub(crate) fn string_to_number(s: &str) -> f64 {
    let t = s.trim_matches(crate::js::is_js_whitespace);
    if t.is_empty() {
        return 0.0;
    }
    for (prefix, radix) in [
        ("0x", 16),
        ("0X", 16),
        ("0o", 8),
        ("0O", 8),
        ("0b", 2),
        ("0B", 2),
    ] {
        if let Some(digits) = t.strip_prefix(prefix) {
            return digits
                .chars()
                .try_fold(0.0f64, |acc, c| {
                    c.to_digit(radix)
                        .map(|d| acc * f64::from(radix) + f64::from(d))
                })
                .filter(|_| !digits.is_empty())
                .unwrap_or(f64::NAN);
        }
    }
    match t {
        "Infinity" | "+Infinity" => return f64::INFINITY,
        "-Infinity" => return f64::NEG_INFINITY,
        _ => {}
    }
    let valid = t
        .bytes()
        .all(|b| b.is_ascii_digit() || matches!(b, b'.' | b'e' | b'E' | b'+' | b'-'));
    if !valid || t.contains("inf") {
        return f64::NAN;
    }
    t.parse::<f64>().unwrap_or(f64::NAN)
}

/// JS `String(x)` / template-literal conversion.
pub(crate) fn to_js_string(v: Option<&Value>) -> String {
    match v {
        None => "undefined".to_owned(),
        Some(Value::Null) => "null".to_owned(),
        Some(Value::Bool(b)) => b.to_string(),
        Some(Value::Number(n)) => number_value_to_string(n),
        Some(Value::String(s)) => s.clone(),
        Some(Value::Array(a)) => a
            .iter()
            .map(|e| match e {
                Value::Null => String::new(),
                other => to_js_string(Some(other)),
            })
            .collect::<Vec<_>>()
            .join(","),
        Some(Value::Object(o)) => {
            if let Some(n) = num_tag(o) {
                number_to_string(n)
            } else if let Some(d) = crate::model_json::as_date(o) {
                crate::model_json::date_to_string(d)
            } else {
                "[object Object]".to_owned()
            }
        }
    }
}

pub(crate) fn number_value_to_string(n: &serde_json::Number) -> String {
    if let Some(i) = n.as_i64() {
        return i.to_string();
    }
    if let Some(u) = n.as_u64() {
        return number_to_string(u as f64);
    }
    number_to_string(n.as_f64().unwrap_or(f64::NAN))
}

/// ECMAScript `Number::toString(x)` (radix 10).
pub(crate) fn number_to_string(x: f64) -> String {
    // Integers below 2^53: their digits are the shortest round-trip ones.
    if x.fract() == 0.0 && x.abs() < 9_007_199_254_740_992.0 && x != 0.0 {
        #[allow(clippy::cast_possible_truncation)]
        return (x as i64).to_string();
    }
    if x.is_nan() {
        return "NaN".to_owned();
    }
    if x == 0.0 {
        return "0".to_owned();
    }
    if x.is_infinite() {
        return if x > 0.0 { "Infinity" } else { "-Infinity" }.to_owned();
    }
    if x < 0.0 {
        return format!("-{}", number_to_string(-x));
    }
    // Shortest round-trip digits and exponent from Rust's `{:e}`:
    // "d.ddddde<exp>" -> digits "dddddd", n = exp + 1 (ECMA-262 terms).
    let sci = format!("{x:e}");
    let (mantissa, exp) = sci.split_once('e').unwrap_or((sci.as_str(), "0"));
    let digits: String = mantissa.chars().filter(char::is_ascii_digit).collect();
    let k = i64::try_from(digits.len()).unwrap_or(i64::MAX);
    let n = exp.parse::<i64>().unwrap_or(0) + 1;
    let pad = |count: i64| "0".repeat(usize::try_from(count).unwrap_or(0));
    if k <= n && n <= 21 {
        format!("{digits}{}", pad(n - k))
    } else if 0 < n && n <= 21 {
        let split = usize::try_from(n).unwrap_or(0);
        let (int, frac) = digits.split_at(split.min(digits.len()));
        format!("{int}.{frac}")
    } else if -6 < n && n <= 0 {
        format!("0.{}{digits}", pad(-n))
    } else {
        let e = n - 1;
        let sign = if e < 0 { '-' } else { '+' };
        let (first, rest) = digits.split_at(1.min(digits.len()));
        if rest.is_empty() {
            format!("{first}e{sign}{}", e.abs())
        } else {
            format!("{first}.{rest}e{sign}{}", e.abs())
        }
    }
}

/// A `serde_json::Value` for a JS number (integers stay integers so they
/// print as JS prints them; NaN and infinities use the `$num` tag).
pub(crate) fn num(x: f64) -> Value {
    if !x.is_finite() {
        return serde_json::json!({"$num": number_to_string(x)});
    }
    if x.fract() == 0.0 && x.abs() < 9_007_199_254_740_992.0 && !(x == 0.0 && x.is_sign_negative())
    {
        #[allow(clippy::cast_possible_truncation)]
        return Value::from(x as i64);
    }
    serde_json::Number::from_f64(x).map_or(Value::Null, Value::Number)
}

/// ExcelJS `utils.xmlEncode`: escapes `" & ' < >` and drops the control
/// characters XML 1.0 cannot hold (0-8, 11, 12, 14-31, 127).
#[cfg(test)]
pub(crate) fn xml_encode(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    escape_into(&mut out, text);
    out
}

/// Appends `xml_encode(text)` to `out`.
pub(crate) fn escape_into(out: &mut String, text: &str) {
    let special = |b: u8| {
        matches!(b, b'"' | b'&' | b'\'' | b'<' | b'>' | 0x7f)
            || (b < b' ' && !matches!(b, b'\t' | b'\n' | b'\r'))
    };
    if !text.bytes().any(special) {
        out.push_str(text);
        return;
    }
    for c in text.chars() {
        match c {
            '"' => out.push_str("&quot;"),
            '&' => out.push_str("&amp;"),
            '\'' => out.push_str("&apos;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '\u{7f}' => {}
            c if c < ' ' && !matches!(c, '\t' | '\n' | '\r') => {}
            c => out.push(c),
        }
    }
}

pub(crate) fn as_str(v: Option<&Value>) -> Option<&str> {
    v.and_then(Value::as_str)
}

/// The JS number in `v`, if it is one (including `$num` tags).
pub(crate) fn as_f64(v: Option<&Value>) -> Option<f64> {
    match v? {
        Value::Number(n) => n.as_f64(),
        Value::Object(o) => num_tag(o),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn number_to_string_matches_javascript() {
        for (x, s) in [
            (1.0, "1"),
            (-2.5, "-2.5"),
            (0.1 + 0.2, "0.30000000000000004"),
            (1e21, "1e+21"),
            (1e20, "100000000000000000000"),
            (123_456_789.123_456_79, "123456789.12345679"),
            (1e-7, "1e-7"),
            (1.5e-7, "1.5e-7"),
            (0.000_001, "0.000001"),
            (5e-324, "5e-324"),
            (1.797_693_134_862_315_7e308, "1.7976931348623157e+308"),
            (f64::NAN, "NaN"),
            (f64::NEG_INFINITY, "-Infinity"),
            (-0.0, "0"),
            (45351.0, "45351"),
        ] {
            assert_eq!(number_to_string(x), s, "{x}");
        }
    }

    #[test]
    fn truthiness_and_strings() {
        let v: Value = serde_json::json!({"a": 0, "b": "", "c": [1, null, "x"], "d": {}});
        assert!(!truthy(get(Some(&v), "a")));
        assert!(!truthy(get(Some(&v), "b")));
        assert!(truthy(get(Some(&v), "d")));
        assert!(!truthy(get(Some(&v), "zzz")));
        assert_eq!(to_js_string(get(Some(&v), "c")), "1,,x");
        assert_eq!(to_js_string(get(Some(&v), "d")), "[object Object]");
        assert_eq!(
            xml_encode("a<b>\u{1}&\"'\t"),
            "a&lt;b&gt;&amp;&quot;&apos;\t"
        );
    }

    #[test]
    fn string_to_number_matches_javascript() {
        assert_eq!(string_to_number(" 12 "), 12.0);
        assert_eq!(string_to_number(""), 0.0);
        assert_eq!(string_to_number("0x10"), 16.0);
        assert!(string_to_number("12abc").is_nan());
        assert_eq!(string_to_number("-Infinity"), f64::NEG_INFINITY);
        assert!(string_to_number("inf").is_nan());
    }
}

/// JS `Number.prototype.toFixed(digits)` (exact-tie rounding picks the
/// larger n, unlike Rust's `{:.N}` which rounds half to even).
pub(crate) fn to_fixed(x: f64, digits: usize) -> String {
    if !x.is_finite() {
        return number_to_string(x);
    }
    if x.abs() >= 1e21 {
        return number_to_string(x);
    }
    // Exact decimal expansion of the binary value, with enough digits to see
    // whether the rest is exactly a half.
    let exact = format!("{:.*}", digits + 30, x.abs());
    let (int_part, frac) = exact.split_once('.').unwrap_or((exact.as_str(), ""));
    let keep = frac.get(..digits).unwrap_or("");
    let rest = frac.get(digits..).unwrap_or("");
    let mut digits_vec: Vec<u8> = format!("{int_part}{keep}")
        .bytes()
        .map(|b| b - b'0')
        .collect();
    let round_up = match rest.as_bytes().first() {
        // At or above half: JS picks the larger n (on the magnitude, since
        // toFixed negates negative numbers first).
        Some(b'5'..=b'9') => true,
        _ => false,
    };
    if round_up {
        let mut i = digits_vec.len();
        loop {
            if i == 0 {
                digits_vec.insert(0, 1);
                break;
            }
            i -= 1;
            if let Some(d) = digits_vec.get_mut(i) {
                if *d == 9 {
                    *d = 0;
                } else {
                    *d += 1;
                    break;
                }
            }
        }
    }
    let s: String = digits_vec.iter().map(|d| char::from(b'0' + d)).collect();
    let (ip, fp) = s.split_at(s.len().saturating_sub(digits));
    let ip = if ip.is_empty() { "0" } else { ip };
    let body = if digits == 0 {
        ip.to_owned()
    } else {
        format!("{ip}.{fp}")
    };
    if x < 0.0 {
        format!("-{body}")
    } else {
        body
    }
}

#[cfg(test)]
mod to_fixed_tests {
    use super::to_fixed;

    #[test]
    fn matches_javascript() {
        assert_eq!(to_fixed(0.125, 2), "0.13");
        assert_eq!(to_fixed(1.005, 2), "1.00");
        assert_eq!(to_fixed(0.1, 2), "0.10");
        assert_eq!(to_fixed(-0.125, 2), "-0.13");
        assert_eq!(to_fixed(-0.0013, 2), "-0.00");
        assert_eq!(to_fixed(2.5, 0), "3");
        assert_eq!(to_fixed(0.0013, 2), "0.00");
        assert_eq!(to_fixed(99.996, 2), "100.00");
    }
}

/// A property that is present with the value `undefined` (`{"$undef": true}`
/// in the JSON codec).
pub(crate) fn undefined() -> Value {
    serde_json::json!({"$undef": true})
}

/// `obj[key] = value` for an object; other values are left alone (JS would
/// set a property on a primitive wrapper and lose it).
pub(crate) fn set(target: &mut Value, key: &str, value: Value) {
    if let Value::Object(o) = target {
        o.insert(key.to_owned(), value);
    }
}
