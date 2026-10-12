//! Port of ExcelJS `lib/utils/col-cache.js` (address and range decoding),
//! with JS number semantics: a missing row or column is `undefined`, and
//! `Math.min(undefined, x)` is `NaN`.

/// `colCache.decodeAddress(value)`.
#[derive(Debug, Clone, PartialEq)]
pub(crate) struct Address {
    pub address: String,
    pub col: Option<f64>,
    pub row: Option<f64>,
    pub col_letters: String,
    pub row_digits: String,
}

/// `colCache.decodeAddress`. `Err` where ExcelJS throws (column past XFD).
pub(crate) fn decode_address(value: &str) -> Result<Address, String> {
    let mut col = String::new();
    let mut col_number: f64 = 0.0;
    let mut row = String::new();
    let mut row_number: f64 = 0.0;
    for u in value.encode_utf16() {
        let has_row = !row.is_empty();
        if !has_row && (65..=90).contains(&u) {
            col.push(char::from(u8::try_from(u).unwrap_or(b'A')));
            col_number = col_number * 26.0 + f64::from(u) - 64.0;
        } else if (48..=57).contains(&u) {
            row.push(char::from(u8::try_from(u).unwrap_or(b'0')));
            row_number = row_number * 10.0 + f64::from(u) - 48.0;
        } else if has_row && !col.is_empty() && u != 36 {
            break;
        }
    }
    let col_v = if col.is_empty() {
        None
    } else if col_number > 16384.0 {
        return Err(format!("Out of bounds. Invalid column letter: {col}"));
    } else {
        Some(col_number)
    };
    let row_v = (!row.is_empty()).then_some(row_number);
    Ok(Address {
        address: format!("{col}{row}"),
        col: col_v,
        row: row_v,
        col_letters: col,
        row_digits: row,
    })
}

/// `colCache.n2l(n)`: `Err` outside 1..=16384, `None` (JS `undefined`) for
/// values with no entry (NaN, fractions).
pub(crate) fn n2l(n: f64) -> Result<Option<String>, String> {
    // Not `!(1.0..=16384.0).contains(&n)`: NaN must pass this check, as in JS.
    #[allow(clippy::manual_range_contains)]
    if n < 1.0 || n > 16384.0 {
        return Err(format!(
            "{} is out of bounds. Excel supports columns from 1 to 16384",
            crate::jsv::number_to_string(n)
        ));
    }
    if n.fract() != 0.0 || n.is_nan() {
        return Ok(None);
    }
    #[allow(clippy::cast_possible_truncation, clippy::cast_sign_loss)]
    Ok(column_letters(n as u32))
}

/// `String(n2l(n))`.
pub(crate) fn n2l_str(n: f64) -> Result<String, String> {
    Ok(n2l(n)?.unwrap_or_else(|| "undefined".to_owned()))
}

/// JS `String(x)` for an optional number.
pub(crate) fn num_str(x: Option<f64>) -> String {
    x.map_or_else(|| "undefined".to_owned(), crate::jsv::number_to_string)
}

fn js_min(a: Option<f64>, b: Option<f64>) -> f64 {
    match (a, b) {
        (Some(a), Some(b)) if !a.is_nan() && !b.is_nan() => a.min(b),
        _ => f64::NAN,
    }
}

fn js_max(a: Option<f64>, b: Option<f64>) -> f64 {
    match (a, b) {
        (Some(a), Some(b)) if !a.is_nan() && !b.is_nan() => a.max(b),
        _ => f64::NAN,
    }
}

/// A decoded range (the `parts.length > 1` result of `decode`/`decodeEx`).
#[derive(Debug, Clone, PartialEq)]
pub(crate) struct RangeEx {
    pub top: f64,
    pub left: f64,
    pub bottom: f64,
    pub right: f64,
    pub sheet_name: Option<String>,
    pub tl: String,
    pub br: String,
    pub dimensions: String,
}

/// Result of `colCache.decodeEx`.
#[derive(Debug, Clone, PartialEq)]
pub(crate) enum DecodedEx {
    Range(RangeEx),
    Error {
        sheet_name: Option<String>,
        error: String,
    },
    Address {
        sheet_name: Option<String>,
        address: Address,
    },
}

/// `/(?:(?:(?:'((?:[^']|'')*)')|([^'^ !]*))!)?(.*)/` -> (sheet name, rest).
fn split_sheet(value: &str) -> (Option<String>, &str) {
    // Quoted: '...'! where '' is an escaped quote.
    if let Some(rest) = value.strip_prefix('\'') {
        let bytes = rest.as_bytes();
        let mut i = 0;
        while i < bytes.len() {
            if bytes.get(i) == Some(&b'\'') {
                if bytes.get(i + 1) == Some(&b'\'') {
                    i += 2;
                    continue;
                }
                break;
            }
            i += 1;
        }
        if bytes.get(i) == Some(&b'\'') && bytes.get(i + 1) == Some(&b'!') {
            let name = rest.get(..i).unwrap_or("");
            let after = rest.get(i + 2..).unwrap_or("");
            return ((!name.is_empty()).then(|| name.to_owned()), after);
        }
    }
    // Unquoted: [^'^ !]* followed by '!'.
    let end = value.find(['\'', '^', ' ', '!']).unwrap_or(value.len());
    if value.as_bytes().get(end) == Some(&b'!') {
        let name = value.get(..end).unwrap_or("");
        return (
            (!name.is_empty()).then(|| name.to_owned()),
            value.get(end + 1..).unwrap_or(""),
        );
    }
    (None, value)
}

fn range_from(tl: &Address, br: &Address, sheet_name: Option<String>) -> Result<RangeEx, String> {
    let top = js_min(tl.row, br.row);
    let left = js_min(tl.col, br.col);
    let bottom = js_max(tl.row, br.row);
    let right = js_max(tl.col, br.col);
    let tl_s = format!("{}{}", n2l_str(left)?, num_str(Some(top)));
    let br_s = format!("{}{}", n2l_str(right)?, num_str(Some(bottom)));
    let dimensions = format!("{tl_s}:{br_s}");
    Ok(RangeEx {
        top,
        left,
        bottom,
        right,
        sheet_name,
        tl: tl_s,
        br: br_s,
        dimensions,
    })
}

/// `colCache.decodeEx(value)`.
pub(crate) fn decode_ex(value: &str) -> Result<DecodedEx, String> {
    let (sheet_name, reference) = split_sheet(value);
    let parts: Vec<&str> = reference.split(':').collect();
    if let (Some(a), Some(b)) = (parts.first(), parts.get(1)) {
        let tl = decode_address(a)?;
        let br = decode_address(b)?;
        return Ok(DecodedEx::Range(range_from(&tl, &br, sheet_name)?));
    }
    if reference.starts_with('#') {
        return Ok(DecodedEx::Error {
            sheet_name,
            error: reference.to_owned(),
        });
    }
    Ok(DecodedEx::Address {
        sheet_name,
        address: decode_address(reference)?,
    })
}

/// `colCache.decode(value)` for the two-part case; `None` for a single
/// address.
pub(crate) fn decode_range(value: &str) -> Result<Option<RangeEx>, String> {
    let parts: Vec<&str> = value.split(':').collect();
    if parts.len() == 2 {
        if let (Some(a), Some(b)) = (parts.first(), parts.get(1)) {
            let tl = decode_address(a)?;
            let br = decode_address(b)?;
            return Ok(Some(range_from(&tl, &br, None)?));
        }
    }
    Ok(None)
}

/// `new Range(value).tl` for a string value.
pub(crate) fn range_tl(value: &str) -> Result<String, String> {
    let (top, left) = match decode_ex(value)? {
        DecodedEx::Range(r) if r.top != 0.0 && !r.top.is_nan() => (Some(r.top), Some(r.left)),
        DecodedEx::Range(_) | DecodedEx::Error { .. } => (None, None),
        DecodedEx::Address { address, .. } => (address.row, address.col),
    };
    Ok(format!(
        "{}{}",
        n2l_str(left.unwrap_or(f64::NAN))?,
        num_str(top)
    ))
}

/// Column number to letters (`1` -> `"A"`, `16384` -> `"XFD"`).
/// `None` outside `1..=16384`, where ExcelJS's `n2l` throws.
fn column_letters(mut col: u32) -> Option<String> {
    if !(1..=16_384).contains(&col) {
        return None;
    }
    let mut letters = Vec::with_capacity(3);
    while col > 0 {
        let rem = (col - 1) % 26;
        letters.push(char::from(b'A' + u8::try_from(rem).ok()?));
        col = (col - 1) / 26;
    }
    Some(letters.iter().rev().collect())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn n2l_lets_nan_through_like_js() {
        // `n < 1 || n > 16384` is false for NaN: no throw, no letters.
        assert_eq!(n2l(f64::NAN), Ok(None));
        assert!(n2l(0.0).is_err());
        assert!(n2l(16385.0).is_err());
    }

    #[test]
    fn decode_ex_matches_exceljs() {
        match decode_ex("'My ''Sheet'''!$A$1:$C$3").unwrap() {
            DecodedEx::Range(r) => {
                assert_eq!(r.sheet_name.as_deref(), Some("My ''Sheet''"));
                assert_eq!(r.dimensions, "A1:C3");
            }
            other => panic!("{other:?}"),
        }
        match decode_ex("Sheet1!B2").unwrap() {
            DecodedEx::Address {
                sheet_name,
                address,
            } => {
                assert_eq!(sheet_name.as_deref(), Some("Sheet1"));
                assert_eq!(address.address, "B2");
            }
            other => panic!("{other:?}"),
        }
        assert!(matches!(
            decode_ex("Sheet1!#REF!").unwrap(),
            DecodedEx::Error { .. }
        ));
        match decode_ex("A:B").unwrap() {
            DecodedEx::Range(r) => assert_eq!(r.dimensions, "ANaN:BNaN"),
            other => panic!("{other:?}"),
        }
        assert!(decode_ex("XFE1").is_err());
    }
}
