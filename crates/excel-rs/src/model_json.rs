//! Tagged values in the JSON workbook model (see
//! `crates/excel-rs-node/lib/xlsx/codec.js`):
//!
//! * a JS `Date` is `{"$date": <ms since epoch>}` (`null` for an invalid date);
//! * a `Buffer` is `{"$buffer": "<base64>"}`.
//!
//! Also ExcelJS's date arithmetic (`utils.dateToExcel` / `excelToDate`).

use serde_json::{json, Value};

use crate::jsv::Obj;

/// Milliseconds since the epoch, `NaN` for an invalid date.
pub(crate) type JsDate = f64;

/// ECMAScript's time value limit: `new Date(x)` is invalid beyond it.
const MAX_TIME: f64 = 8.64e15;

pub(crate) fn as_date(o: &Obj) -> Option<JsDate> {
    if o.len() != 1 {
        return None;
    }
    match o.get("$date")? {
        Value::Null => Some(f64::NAN),
        v => Some(v.as_f64().unwrap_or(f64::NAN)),
    }
}

pub(crate) fn value_as_date(v: Option<&Value>) -> Option<JsDate> {
    match v {
        Some(Value::Object(o)) => as_date(o),
        _ => None,
    }
}

/// `new Date(ms)`: applies JS's TimeClip.
pub(crate) fn date(ms: f64) -> Value {
    if ms.is_nan() || ms.abs() > MAX_TIME {
        return json!({"$date": null});
    }
    json!({"$date": crate::jsv::num(ms.trunc() + 0.0)})
}

pub(crate) fn as_buffer(v: Option<&Value>) -> Option<Vec<u8>> {
    match v {
        Some(Value::Object(o)) if o.len() == 1 => base64_decode(o.get("$buffer")?.as_str()?),
        _ => None,
    }
}

pub(crate) fn buffer(bytes: &[u8]) -> Value {
    json!({"$buffer": base64_encode(bytes)})
}

/// JS `Math.round`: halves round toward +Infinity.
pub(crate) fn js_round(x: f64) -> f64 {
    if !x.is_finite() {
        return x;
    }
    let f = x.floor();
    if x - f >= 0.5 {
        f + 1.0
    } else {
        f
    }
}

/// ExcelJS `utils.dateToExcel(d, date1904)`.
pub(crate) fn date_to_excel(ms: JsDate, date1904: bool) -> f64 {
    25569.0 + (ms / (24.0 * 3600.0 * 1000.0)) - if date1904 { 1462.0 } else { 0.0 }
}

/// ExcelJS `utils.excelToDate(v, date1904)`.
pub(crate) fn excel_to_date(v: f64, date1904: bool) -> Value {
    let offset = if date1904 { 1462.0 } else { 0.0 };
    date(js_round((v - 25569.0 + offset) * 24.0 * 3600.0 * 1000.0))
}

/// JS `Date.prototype.toISOString()`; `None` for an invalid date (where JS
/// throws a RangeError).
pub(crate) fn to_iso_string(ms: JsDate) -> Option<String> {
    if ms.is_nan() || ms.abs() > MAX_TIME {
        return None;
    }
    #[allow(clippy::cast_possible_truncation)]
    let total_ms = ms as i64;
    let days = total_ms.div_euclid(86_400_000);
    let rem = total_ms.rem_euclid(86_400_000);
    let (y, m, d) = civil_from_days(days);
    let (h, mi, s, milli) = (
        rem / 3_600_000,
        rem / 60_000 % 60,
        rem / 1000 % 60,
        rem % 1000,
    );
    let year = if (0..=9999).contains(&y) {
        format!("{y:04}")
    } else if y < 0 {
        format!("-{:06}", -y)
    } else {
        format!("+{y:06}")
    };
    Some(format!(
        "{year}-{m:02}-{d:02}T{h:02}:{mi:02}:{s:02}.{milli:03}Z"
    ))
}

/// `String(date)` is time-zone dependent in JS; excel-rs uses UTC. Only used
/// where ExcelJS would stringify a date by accident.
pub(crate) fn date_to_string(ms: JsDate) -> String {
    to_iso_string(ms).unwrap_or_else(|| "Invalid Date".to_owned())
}

/// Parses the date forms `new Date(string)` accepts in ExcelJS's core.xml
/// reader (W3CDTF / ISO-8601). Returns milliseconds or `NaN`.
pub(crate) fn parse_iso_date(s: &str) -> JsDate {
    parse_iso(s.trim()).unwrap_or(f64::NAN)
}

fn parse_iso(s: &str) -> Option<f64> {
    let b = s.as_bytes();
    let num = |from: usize, len: usize| -> Option<i64> {
        let part = s.get(from..from + len)?;
        part.bytes()
            .all(|c| c.is_ascii_digit())
            .then(|| part.parse().ok())?
    };
    let year = num(0, 4)?;
    let (mut month, mut day) = (1, 1);
    let mut i = 4;
    if b.get(i) == Some(&b'-') {
        month = num(i + 1, 2)?;
        i += 3;
        if b.get(i) == Some(&b'-') {
            day = num(i + 1, 2)?;
            i += 3;
        }
    }
    let (mut h, mut mi, mut sec, mut ms) = (0, 0, 0, 0.0);
    let mut date_only = true;
    if matches!(b.get(i), Some(b'T' | b't' | b' ')) {
        date_only = false;
        h = num(i + 1, 2)?;
        if b.get(i + 3) != Some(&b':') {
            return None;
        }
        mi = num(i + 4, 2)?;
        i += 6;
        if b.get(i) == Some(&b':') {
            sec = num(i + 1, 2)?;
            i += 3;
            if b.get(i) == Some(&b'.') {
                let digits = b
                    .get(i + 1..)?
                    .iter()
                    .take_while(|c| c.is_ascii_digit())
                    .count();
                let frac: String = s.get(i + 1..i + 1 + digits)?.chars().take(3).collect();
                ms = format!("0.{frac}").parse::<f64>().ok()? * 1000.0;
                i += 1 + digits;
            }
        }
    }
    let mut offset_min = 0i64;
    match b.get(i) {
        None => {
            // Date-only forms are UTC; date-time forms without an offset are
            // local time in JS. excel-rs treats them as UTC.
            let _ = date_only;
        }
        Some(b'Z' | b'z') if i + 1 == b.len() => {}
        Some(sign @ (b'+' | b'-')) => {
            let oh = num(i + 1, 2)?;
            let om = if b.get(i + 3) == Some(&b':') {
                num(i + 4, 2)?
            } else {
                num(i + 3, 2)?
            };
            offset_min = (oh * 60 + om) * if *sign == b'-' { -1 } else { 1 };
        }
        _ => return None,
    }
    if !(1..=12).contains(&month) || !(1..=31).contains(&day) || h > 24 || mi > 59 || sec > 59 {
        return None;
    }
    let days = days_from_civil(year, month, day);
    #[allow(clippy::cast_precision_loss)]
    let total =
        ((days * 24 + h) * 60 + mi - offset_min) as f64 * 60_000.0 + sec as f64 * 1000.0 + ms;
    Some(total)
}

fn days_from_civil(y: i64, m: i64, d: i64) -> i64 {
    let y = if m <= 2 { y - 1 } else { y };
    let era = y.div_euclid(400);
    let yoe = y - era * 400;
    let mp = (m + 9) % 12;
    let doy = (153 * mp + 2) / 5 + d - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    era * 146_097 + doe - 719_468
}

fn civil_from_days(z: i64) -> (i64, i64, i64) {
    let z = z + 719_468;
    let era = z.div_euclid(146_097);
    let doe = z - era * 146_097;
    let yoe = (doe - doe / 1460 + doe / 36_524 - doe / 146_096) / 365;
    let y = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = doy - (153 * mp + 2) / 5 + 1;
    let m = if mp < 10 { mp + 3 } else { mp - 9 };
    (if m <= 2 { y + 1 } else { y }, m, d)
}

const B64: &[u8; 64] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

pub(crate) fn base64_encode(bytes: &[u8]) -> String {
    let mut out = String::with_capacity(bytes.len().div_ceil(3) * 4);
    for chunk in bytes.chunks(3) {
        let n = chunk
            .iter()
            .enumerate()
            .fold(0u32, |acc, (i, &b)| acc | u32::from(b) << (16 - 8 * i));
        for i in 0..4 {
            if i <= chunk.len() {
                let idx = (n >> (18 - 6 * i)) & 63;
                out.push(char::from(*B64.get(idx as usize).unwrap_or(&b'A')));
            } else {
                out.push('=');
            }
        }
    }
    out
}

/// Lenient like Node's `Buffer.from(s, 'base64')`: skips characters outside
/// the alphabet, accepts URL-safe digits, stops at padding.
pub(crate) fn base64_decode(s: &str) -> Option<Vec<u8>> {
    let mut out = Vec::with_capacity(s.len() / 4 * 3);
    let mut acc = 0u32;
    let mut bits = 0;
    for c in s.bytes() {
        let v = match c {
            b'A'..=b'Z' => c - b'A',
            b'a'..=b'z' => c - b'a' + 26,
            b'0'..=b'9' => c - b'0' + 52,
            b'+' | b'-' => 62,
            b'/' | b'_' => 63,
            b'=' => break,
            _ => continue,
        };
        acc = (acc << 6) | u32::from(v);
        bits += 6;
        if bits >= 8 {
            bits -= 8;
            out.push(u8::try_from((acc >> bits) & 0xff).ok()?);
        }
    }
    Some(out)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn base64_round_trip() {
        for s in [&b""[..], b"f", b"fo", b"foo", b"foob", b"fooba", b"foobar"] {
            assert_eq!(base64_decode(&base64_encode(s)).unwrap(), s);
        }
        assert_eq!(base64_encode(b"foobar"), "Zm9vYmFy");
        assert_eq!(base64_encode(b"fo"), "Zm8=");
    }

    #[test]
    fn dates() {
        assert_eq!(to_iso_string(0.0).unwrap(), "1970-01-01T00:00:00.000Z");
        let ms = parse_iso_date("2024-02-29T08:30:00Z");
        assert_eq!(to_iso_string(ms).unwrap(), "2024-02-29T08:30:00.000Z");
        assert_eq!(
            parse_iso_date("2020-01-01T00:00:00.123+01:00"),
            parse_iso_date("2019-12-31T23:00:00.123Z")
        );
        assert!(parse_iso_date("nope").is_nan());
        assert_eq!(date_to_excel(0.0, false), 25569.0);
        assert_eq!(js_round(-0.5), 0.0);
        assert_eq!(js_round(2.5), 3.0);
        assert_eq!(js_round(0.499_999_999_999_999_94), 0.0);
        assert_eq!(excel_to_date(25569.0, false), json!({"$date": 0}));
    }
}
