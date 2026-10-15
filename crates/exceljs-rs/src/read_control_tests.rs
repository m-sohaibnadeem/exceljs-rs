//! Boundary tests independent of Node and ZIP size headers.
use crate::{ReadControl, ReadLimits};
use std::io::{Cursor, Write};
use std::sync::atomic::Ordering;

fn package(parts: &[(&str, &str)]) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for (name, text) in parts {
        assert!(zip
            .start_file(*name, zip::write::SimpleFileOptions::default())
            .is_ok());
        assert!(zip.write_all(text.as_bytes()).is_ok());
    }
    zip.finish()
        .map_or_else(|e| unreachable!("test ZIP: {e}"), Cursor::into_inner)
}

#[test]
fn inflated_bytes_count_all_entries_once() {
    let data = package(&[("unknown", "123"), ("xl/media/a.png", "4567")]);
    let exact = ReadControl::new(ReadLimits {
        max_entry_bytes: Some(4),
        max_inflated_bytes: Some(7),
        max_zip_entries: Some(2),
        ..ReadLimits::default()
    });
    assert!(exact.preflight(&data).is_ok());
    assert_eq!(exact.inflated_bytes.load(Ordering::Relaxed), 7);
    for limit in 0..7 {
        let control = ReadControl::new(ReadLimits {
            max_inflated_bytes: Some(limit),
            ..ReadLimits::default()
        });
        assert!(control.preflight(&data).is_err());
    }
}

#[test]
fn raw_tokens_are_bounded_before_parser_allocation() {
    for token in [
        "<a/>",
        "<!--quoted > -->",
        "<![CDATA[text >]]>",
        "<a v='>'/>",
    ] {
        let data = package(&[("unknown.xml", token)]);
        let size = u64::try_from(token.len()).unwrap_or(u64::MAX);
        let exact = ReadControl::new(ReadLimits {
            max_token_bytes: Some(size),
            ..ReadLimits::default()
        });
        assert!(exact.preflight(&data).is_ok());
        let small = ReadControl::new(ReadLimits {
            max_token_bytes: Some(size - 1),
            ..ReadLimits::default()
        });
        assert!(small.preflight(&data).is_err());
    }
}

#[test]
fn cells_strings_and_expansions_have_inclusive_bounds() {
    let data = package(&[
        (
            "xl/worksheets/sheet1.xml",
            "<worksheet><c/><c/><dataValidation sqref='A1:B2&#xFEFF;D1'/></worksheet>",
        ),
        ("xl/sharedStrings.xml", "<sst><si/><si/></sst>"),
    ]);
    let exact = ReadControl::new(ReadLimits {
        max_cells: Some(2),
        max_shared_strings: Some(2),
        max_validation_cells: Some(5),
        ..ReadLimits::default()
    });
    assert!(exact.preflight(&data).is_ok());
    for limits in [
        ReadLimits {
            max_cells: Some(1),
            ..ReadLimits::default()
        },
        ReadLimits {
            max_shared_strings: Some(1),
            ..ReadLimits::default()
        },
        ReadLimits {
            max_validation_cells: Some(4),
            ..ReadLimits::default()
        },
    ] {
        assert!(ReadControl::new(limits).preflight(&data).is_err());
    }
}

#[test]
fn cancellation_stops_preflight() {
    let control = ReadControl::new(ReadLimits::default());
    control.cancel();
    assert!(control.preflight(&package(&[("unknown", "123")])).is_err());
    assert_eq!(control.inflated_bytes.load(Ordering::Relaxed), 0);
}
