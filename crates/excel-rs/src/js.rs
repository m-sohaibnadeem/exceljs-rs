//! JavaScript primitives that ExcelJS relies on, reproduced exactly.
//!
//! ExcelJS turns XML text into values with `parseInt(x, 10)`, `parseFloat(x)`
//! and a few regex replacements. Those functions accept odd input ("12abc"
//! parses as 12) and excel-rs must accept the same input the same way, so
//! these are behaviour ports, not general-purpose parsers.

/// `StrWhiteSpaceChar` from ECMA-262: what `parseInt`/`parseFloat` skip.
/// (Differs from `char::is_whitespace`: includes U+FEFF, excludes U+0085.)
pub(crate) fn is_js_whitespace(c: char) -> bool {
    matches!(
        c,
        '\t' | '\n' | '\u{0B}' | '\u{0C}' | '\r' | ' ' | '\u{A0}' | '\u{1680}' | '\u{2000}'
            ..='\u{200A}'
                | '\u{2028}'
                | '\u{2029}'
                | '\u{202F}'
                | '\u{205F}'
                | '\u{3000}'
                | '\u{FEFF}'
    )
}

fn trim_js_start(s: &str) -> &str {
    s.trim_start_matches(is_js_whitespace)
}

/// `parseInt(s, 10)`. `None` stands for `NaN`.
pub(crate) fn parse_int(s: &str) -> Option<f64> {
    // Common case: a short run of digits only.
    if !s.is_empty() && s.len() <= 15 && s.bytes().all(|b| b.is_ascii_digit()) {
        #[allow(clippy::cast_precision_loss)] // at most 15 digits: exact
        return Some(s.bytes().fold(0u64, |n, b| n * 10 + u64::from(b - b'0')) as f64);
    }
    let s = trim_js_start(s);
    let (negative, rest) = match s.as_bytes().first() {
        Some(b'-') => (true, s.get(1..).unwrap_or("")),
        Some(b'+') => (false, s.get(1..).unwrap_or("")),
        _ => (false, s),
    };
    let digits_len = rest.bytes().take_while(u8::is_ascii_digit).count();
    let digits = rest.get(..digits_len).filter(|d| !d.is_empty())?;
    let value: f64 = digits.parse().ok()?;
    Some(if negative { -value } else { value })
}

/// `parseFloat(s)`. Returns `NaN` where JavaScript would.
pub(crate) fn parse_float(s: &str) -> f64 {
    let s = trim_js_start(s);
    let bytes = s.as_bytes();
    let mut i = 0;
    let negative = match bytes.first() {
        Some(b'-') => {
            i = 1;
            true
        }
        Some(b'+') => {
            i = 1;
            false
        }
        _ => false,
    };
    let after_sign = s.get(i..).unwrap_or("");
    if after_sign.starts_with("Infinity") {
        return if negative {
            f64::NEG_INFINITY
        } else {
            f64::INFINITY
        };
    }
    let count_digits = |from: usize| {
        bytes
            .get(from..)
            .map_or(0, |b| b.iter().take_while(|c| c.is_ascii_digit()).count())
    };
    let int_digits = count_digits(i);
    let mut end = i + int_digits;
    let mut frac_digits = 0;
    if bytes.get(end) == Some(&b'.') {
        frac_digits = count_digits(end + 1);
        if int_digits > 0 || frac_digits > 0 {
            end += 1 + frac_digits;
        }
    }
    if int_digits == 0 && frac_digits == 0 {
        return f64::NAN;
    }
    if matches!(bytes.get(end), Some(b'e' | b'E')) {
        let mut j = end + 1;
        if matches!(bytes.get(j), Some(b'+' | b'-')) {
            j += 1;
        }
        let exp_digits = count_digits(j);
        if exp_digits > 0 {
            end = j + exp_digits;
        }
    }
    s.get(..end)
        .and_then(|literal| literal.parse::<f64>().ok())
        .unwrap_or(f64::NAN)
}

/// Decodes the `_xHHHH_` escapes of OOXML text (`ST_Xstring`): each names a
/// UTF-16 code unit, in upper- or lower-case hex. ExcelJS's `TextXform`
/// matches only `/_x([0-9A-F]{4})_/g` and only in shared strings (B1).
///
/// `fromCharCode` produces UTF-16 code units, so `_xD83D__xDE00_` becomes one
/// emoji. A lone surrogate, which a Rust `String` cannot hold, becomes U+FFFD.
pub(crate) fn decode_x_escapes(text: &str) -> String {
    if !text.contains("_x") {
        return text.to_owned();
    }
    let units: Vec<u16> = text.encode_utf16().collect();
    let mut out: Vec<u16> = Vec::with_capacity(units.len());
    let mut i = 0;
    while let Some(&unit) = units.get(i) {
        if let Some(code) = x_escape_at(&units, i) {
            out.push(code);
            i += 7;
        } else {
            out.push(unit);
            i += 1;
        }
    }
    String::from_utf16_lossy(&out)
}

/// If `units[i..i + 7]` is `_xHHHH_`, the code unit it names.
fn x_escape_at(units: &[u16], i: usize) -> Option<u16> {
    let window = units.get(i..i + 7)?;
    let (head, rest) = window.split_at(2);
    let (hex, tail) = rest.split_at(4);
    if head != [u16::from(b'_'), u16::from(b'x')] || tail != [u16::from(b'_')] {
        return None;
    }
    hex.iter().try_fold(0u16, |acc, &u| {
        let digit = match u8::try_from(u).ok()? {
            d @ b'0'..=b'9' => d - b'0',
            d @ b'A'..=b'F' => d - b'A' + 10,
            d @ b'a'..=b'f' => d - b'a' + 10,
            _ => return None,
        };
        Some(acc * 16 + u16::from(digit))
    })
}

/// Escapes text for an `ST_Xstring` element, so that [`decode_x_escapes`]
/// gives it back: the `_` of literal text shaped like an escape becomes
/// `_x005F_` (B1). Control characters are left to `xmlEncode`, which drops
/// them (ExcelJS's issue-234 spec requires that).
pub(crate) fn encode_x_escapes(text: &str) -> std::borrow::Cow<'_, str> {
    if !text.contains("_x") {
        return std::borrow::Cow::Borrowed(text);
    }
    let units: Vec<u16> = text.encode_utf16().collect();
    let mut out = String::with_capacity(text.len() + 14);
    let mut at = 0; // UTF-16 index of `c`
    for c in text.chars() {
        if c == '_' && x_escape_at(&units, at).is_some() {
            out.push_str("_x005F_");
        } else {
            out.push(c);
        }
        at += c.len_utf16();
    }
    std::borrow::Cow::Owned(out)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_int_matches_javascript() {
        assert_eq!(parse_int("12"), Some(12.0));
        assert_eq!(parse_int("  1.9 "), Some(1.0));
        assert_eq!(parse_int("-3x"), Some(-3.0));
        assert_eq!(parse_int("+7"), Some(7.0));
        assert_eq!(parse_int("\u{FEFF}5"), Some(5.0));
        assert_eq!(parse_int("true"), None);
        assert_eq!(parse_int(""), None);
        assert_eq!(parse_int("-"), None);
        assert_eq!(parse_int("\u{85}5"), None);
    }

    #[test]
    fn parse_float_matches_javascript() {
        assert_eq!(parse_float("12abc"), 12.0);
        assert_eq!(parse_float(" 1e3 "), 1000.0);
        assert_eq!(parse_float(".5"), 0.5);
        assert_eq!(parse_float("5."), 5.0);
        assert_eq!(parse_float("1e"), 1.0);
        assert_eq!(parse_float("1e+"), 1.0);
        assert_eq!(parse_float("-1.5E-2x"), -0.015);
        assert_eq!(parse_float("0x10"), 0.0);
        assert_eq!(parse_float("-Infinity"), f64::NEG_INFINITY);
        assert_eq!(parse_float("Infinityx"), f64::INFINITY);
        assert!(parse_float("abc").is_nan());
        assert!(parse_float(".").is_nan());
        assert!(parse_float("").is_nan());
        assert!(parse_float("-0").is_sign_negative());
    }

    #[test]
    fn x_escapes_match_exceljs() {
        assert_eq!(decode_x_escapes("_x0041__x0062_"), "Ab");
        assert_eq!(decode_x_escapes("_x006a_"), "j");
        assert_eq!(decode_x_escapes("_xD83D__xDE00_"), "😀");
        assert_eq!(decode_x_escapes("_xD83D_"), "\u{FFFD}");
        assert_eq!(decode_x_escapes("_x004_"), "_x004_");
        assert_eq!(decode_x_escapes("é_x0041_"), "éA");
    }

    #[test]
    fn x_escapes_round_trip() {
        for s in [
            "_x0041_ literal",
            "a_x005F_b",
            "_x_ _x12_ é_xABCD_",
            "plain",
        ] {
            assert_eq!(decode_x_escapes(&encode_x_escapes(s)), s, "{s:?}");
        }
        assert_eq!(encode_x_escapes("_x0041_"), "_x005F_x0041_");
        assert_eq!(encode_x_escapes("_x12_"), "_x12_");
    }
}
