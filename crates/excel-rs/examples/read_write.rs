//! Times the engine alone: `cargo run --release --example read_write -- book.xlsx [n]`
//! reads the file n times (default 1) and writes the read model back.

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut args = std::env::args().skip(1);
    let path = args.next().ok_or("usage: read_write <file.xlsx> [n]")?;
    let n: usize = args.next().map_or(Ok(1), |s| s.parse())?;
    let data = std::fs::read(&path)?;
    let threaded = std::env::var("THREAD").is_ok();
    for _ in 0..n {
        if threaded {
            let d = data.clone();
            let t = std::time::Instant::now();
            let h = std::thread::spawn(move || {
                let out = excel_rs::read_xlsx_model(&d, &excel_rs::ReadOptions::default())
                    .map(|o| o.rows.encode().len());
                out.map_err(|e| e.to_string())
            });
            let len = h.join().map_err(|_| "thread panicked")??;
            println!("thread read+encode {:?} ({len} bytes)", t.elapsed());
            continue;
        }
        let t = std::time::Instant::now();
        let out = excel_rs::read_xlsx_model(&data, &excel_rs::ReadOptions::default())?;
        let read = t.elapsed();
        let t = std::time::Instant::now();
        let rows = out.rows.encode();
        let encode = t.elapsed();
        println!(
            "read {read:?}  encode rows {encode:?} ({} bytes)",
            rows.len()
        );
    }
    Ok(())
}
