//! Times the engine's write alone, on the arguments the Node package passes
//! to `writeXlsx` (dump them with `encodeWorkbook(wb._modelForWrite())`):
//! `cargo run --release --example write_parts -- <dir> [n]` reads
//! `<dir>/w-model.json`, `<dir>/w-rows.bin` and `<dir>/w-styles.json` and
//! writes the workbook n times (default 1).

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut args = std::env::args().skip(1);
    let dir = std::path::PathBuf::from(args.next().ok_or("usage: write_parts <dir> [n]")?);
    let n: usize = args.next().map_or(Ok(1), |s| s.parse())?;
    let model = std::fs::read_to_string(dir.join("w-model.json"))?;
    let rows = std::fs::read(dir.join("w-rows.bin"))?;
    let styles = std::fs::read_to_string(dir.join("w-styles.json"))?;
    for _ in 0..n {
        let t = std::time::Instant::now();
        let out =
            excel_rs::write_xlsx_parts(&model, &rows, &styles, &excel_rs::WriteOptions::default())?;
        println!("write {:?} ({} bytes)", t.elapsed(), out.len());
    }
    Ok(())
}
