//! Snapshot test of the reader: the engine's workbook model for every
//! `fixtures/*.xlsx` must equal `fixtures/<name>.model.json` (the model
//! `workbook.model = ...` receives, as ExcelJS's reader shapes it). The
//! snapshots began as the real ExcelJS 4.4.0's models; intended changes
//! (COMPATIBILITY.md) are written back with `UPDATE_SNAPSHOTS=1 cargo test`.
//! The cells a user sees (`fixtures/<name>.json`) are checked by
//! `crates/excel-rs-node/test/snapshots.test.js`.

// Test code: failing loudly is the point. (The library itself denies these.)
#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::panic,
    clippy::indexing_slicing
)]

use std::path::{Path, PathBuf};

use excel_rs::{read_xlsx_model, ReadOptions};
use serde_json::Value;

fn fixtures_dir() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures")
}

/// The read model with each worksheet's typed rows put back as `rows`.
fn full_model(out: &excel_rs::ReadOutput) -> Value {
    let mut model = out.model.clone();
    if let Some(Value::Array(sheets)) = model.get_mut("worksheets") {
        for (ws, rows) in sheets.iter_mut().zip(out.rows.to_values()) {
            if let Value::Object(m) = ws {
                m.insert("rows".into(), rows);
            }
        }
    }
    model
}

/// Replaces `{"$ref": n}` with the shared object (snapshots hold shared
/// objects in full) and drops `{"$undef": true}` keys (JSON has no
/// undefined).
fn normalise(v: &Value, refs: &[Value]) -> Value {
    match v {
        Value::Object(o) => {
            if let (1, Some(n)) = (o.len(), o.get("$ref").and_then(Value::as_u64)) {
                return normalise(&refs[usize::try_from(n).unwrap()], refs);
            }
            Value::Object(
                o.iter()
                    .filter(|(_, v)| v.get("$undef") != Some(&Value::Bool(true)))
                    .map(|(k, v)| (k.clone(), normalise(v, refs)))
                    .collect(),
            )
        }
        Value::Array(a) => Value::Array(a.iter().map(|v| normalise(v, refs)).collect()),
        other => other.clone(),
    }
}

/// One difference: the path and each side's value (`None`: absent).
type Diff = (String, Option<Value>, Option<Value>);

/// Every difference between two models.
fn all_diffs(a: &Value, b: &Value, at: &str, out: &mut Vec<Diff>) {
    match (a, b) {
        (Value::Object(x), Value::Object(y)) => {
            let mut keys: Vec<&String> = x.keys().collect();
            keys.extend(y.keys().filter(|k| !x.contains_key(*k)));
            for k in keys {
                let path = format!("{at}.{k}");
                match (x.get(k), y.get(k)) {
                    (Some(p), Some(q)) => all_diffs(p, q, &path, out),
                    (p, q) => out.push((path, p.cloned(), q.cloned())),
                }
            }
        }
        (Value::Array(x), Value::Array(y)) => {
            if x.len() != y.len() {
                out.push((
                    format!("{at}.length"),
                    Some(x.len().into()),
                    Some(y.len().into()),
                ));
            }
            for (i, (p, q)) in x.iter().zip(y).enumerate() {
                all_diffs(p, q, &format!("{at}[{i}]"), out);
            }
        }
        // 1 and 1.0 are the same JS number
        (Value::Number(x), Value::Number(y)) if x.as_f64() == y.as_f64() => {}
        _ if a == b => {}
        _ => out.push((at.to_owned(), Some(a.clone()), Some(b.clone()))),
    }
}

/// The snapshot's JSON: one-space indent, a final newline.
fn to_snapshot(v: &Value) -> String {
    let mut out = Vec::new();
    let mut ser = serde_json::Serializer::with_formatter(
        &mut out,
        serde_json::ser::PrettyFormatter::with_indent(b" "),
    );
    serde::Serialize::serialize(v, &mut ser).unwrap();
    out.push(b'\n');
    String::from_utf8(out).unwrap()
}

fn check_fixture(path: &Path, update: bool) -> String {
    let expected: Value = serde_json::from_str(&std::fs::read_to_string(path).unwrap()).unwrap();
    let name = path
        .file_name()
        .unwrap()
        .to_str()
        .unwrap()
        .replace(".model.json", ".xlsx");
    let data = std::fs::read(fixtures_dir().join(&name)).unwrap();
    let ours = match read_xlsx_model(&data, &ReadOptions::default()) {
        Ok(out) => normalise(&full_model(&out), &out.refs),
        Err(err) => serde_json::json!({ "error": err.exceljs_message() }),
    };
    let mut diffs = Vec::new();
    all_diffs(&expected, &ours, "$", &mut diffs);
    if diffs.is_empty() {
        return match ours.get("error") {
            Some(e) => format!("{name}: rejected as expected ({e})"),
            None => format!("{name}: model equals the snapshot"),
        };
    }
    if update {
        std::fs::write(path, to_snapshot(&ours)).unwrap();
        return format!("{name}: snapshot updated ({} differences)", diffs.len());
    }
    panic!(
        "{name}: model differs from {} (UPDATE_SNAPSHOTS=1 writes it if the change is intended):\n{}",
        path.display(),
        diffs
            .iter()
            .take(20)
            .map(|(p, a, b)| format!("  {p}: snapshot {a:?}, engine {b:?}"))
            .collect::<Vec<_>>()
            .join("\n")
    )
}

#[test]
fn every_fixture_model_matches_its_snapshot() {
    let mut paths: Vec<PathBuf> = std::fs::read_dir(fixtures_dir())
        .unwrap()
        .map(|e| e.unwrap().path())
        .filter(|p| p.to_str().is_some_and(|s| s.ends_with(".model.json")))
        .collect();
    paths.sort();
    assert!(paths.len() >= 20, "fixtures missing");
    let update = std::env::var_os("UPDATE_SNAPSHOTS").is_some();
    let mut failures = Vec::new();
    for path in &paths {
        match std::panic::catch_unwind(|| check_fixture(path, update)) {
            Ok(line) => println!("{line}"),
            Err(e) => failures.push(
                e.downcast_ref::<String>()
                    .cloned()
                    .unwrap_or_else(|| format!("{}: failed", path.display())),
            ),
        }
    }
    assert!(failures.is_empty(), "\n{}", failures.join("\n\n"));
}

/// Truncated and corrupted copies of every fixture must give `Err` or a
/// model, never a panic; so must writing back what was read.
#[test]
fn damaged_files_never_panic() {
    for entry in std::fs::read_dir(fixtures_dir()).unwrap() {
        let path = entry.unwrap().path();
        if path.extension().is_none_or(|e| e != "xlsx") {
            continue;
        }
        let data = std::fs::read(&path).unwrap();
        let options = ReadOptions::default();
        for cut in [
            0,
            1,
            4,
            22,
            100,
            data.len() / 2,
            data.len().saturating_sub(1),
        ] {
            let _ = read_xlsx_model(data.get(..cut).unwrap_or(&data), &options);
        }
        let mut flipped = data.clone();
        for i in (0..flipped.len()).step_by(97) {
            flipped[i] ^= 0x5A;
            if let Ok(out) = read_xlsx_model(&flipped, &options) {
                let _ = excel_rs::write_xlsx_model(
                    normalise(&full_model(&out), &out.refs),
                    &excel_rs::WriteOptions::default(),
                );
            }
        }
    }
}
