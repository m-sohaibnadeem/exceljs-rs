//! Opt-in archive budgets and cooperative cancellation. Budgets count a
//! preflight inflation of every ZIP entry once; subsequent parsing is uncharged.
use std::cell::RefCell;
use std::io::{Cursor, Read};
use std::sync::{
    atomic::{AtomicBool, AtomicU64, Ordering},
    Arc,
};

use crate::error::{Error, Result};
use serde::Deserialize;

/// Inclusive limits, in bytes or counts. Zero is a meaningful limit.
#[derive(Debug, Clone, Default, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct ReadLimits {
    pub max_input_bytes: Option<u64>,
    pub max_zip_entries: Option<u64>,
    pub max_entry_bytes: Option<u64>,
    pub max_inflated_bytes: Option<u64>,
    pub max_cells: Option<u64>,
    pub max_shared_strings: Option<u64>,
    pub max_token_bytes: Option<u64>,
    pub max_validation_cells: Option<u64>,
    pub max_merge_cells: Option<u64>,
}

/// Shared by the parser thread and its owner. Cancellation never needs a
/// receiver lock, so a pending read can be interrupted safely.
#[derive(Debug, Default)]
pub struct ReadControl {
    pub limits: ReadLimits,
    cancelled: AtomicBool,
    pub inflated_bytes: AtomicU64,
    pub entries: AtomicU64,
    pub cells: AtomicU64,
    pub shared_strings: AtomicU64,
    pub validation_cells: AtomicU64,
    pub merge_cells: AtomicU64,
    pub rows: AtomicU64,
    pub finished: AtomicBool,
}

impl ReadControl {
    pub fn new(limits: ReadLimits) -> Arc<Self> {
        Arc::new(Self {
            limits,
            ..Self::default()
        })
    }
    pub fn cancel(&self) {
        self.cancelled.store(true, Ordering::Relaxed);
    }
    pub fn check(&self) -> Result<()> {
        if self.cancelled.load(Ordering::Relaxed) {
            Err(Error::invalid("workbook", "XLSX read cancelled"))
        } else {
            Ok(())
        }
    }
    fn add(&self, counter: &AtomicU64, n: u64, limit: Option<u64>, name: &str) -> Result<()> {
        let old = counter.load(Ordering::Relaxed);
        let next = old.checked_add(n).ok_or_else(|| exceeded(name))?;
        bound(next, limit, name)?;
        counter.store(next, Ordering::Relaxed);
        Ok(())
    }
    /// Scan every entry, including directories, duplicates, media and ignored
    /// parts. A fixed-size read prevents trusting attacker-controlled ZIP sizes.
    pub fn preflight(&self, data: &[u8]) -> Result<()> {
        self.check()?;
        bound(
            u64::try_from(data.len()).unwrap_or(u64::MAX),
            self.limits.max_input_bytes,
            "maxInputBytes",
        )?;
        let mut zip = zip::ZipArchive::new(Cursor::new(data))?;
        bound(
            u64::try_from(zip.len()).unwrap_or(u64::MAX),
            self.limits.max_zip_entries,
            "maxZipEntries",
        )?;
        for i in 0..zip.len() {
            self.check()?;
            let mut entry = zip.by_index(i)?;
            let name = entry.name().map(|s| s.into_owned()).unwrap_or_default();
            let xml = name.ends_with(".xml") || name.ends_with(".rels") || name.ends_with(".vml");
            let mut scan = TokenScan::default();
            let mut entry_size = 0u64;
            let mut buf = [0u8; 8192];
            loop {
                self.check()?;
                let n = entry.read(&mut buf)?;
                if n == 0 {
                    break;
                }
                let n64 = u64::try_from(n).unwrap_or(u64::MAX);
                entry_size = entry_size
                    .checked_add(n64)
                    .ok_or_else(|| exceeded("maxEntryBytes"))?;
                bound(entry_size, self.limits.max_entry_bytes, "maxEntryBytes")?;
                self.add(
                    &self.inflated_bytes,
                    n64,
                    self.limits.max_inflated_bytes,
                    "maxInflatedBytes",
                )?;
                if xml {
                    scan.feed(buf.get(..n).unwrap_or_default(), &name, self)?;
                }
            }
            self.entries.fetch_add(1, Ordering::Relaxed);
        }
        Ok(())
    }
}

fn exceeded(name: &str) -> Error {
    Error::invalid("workbook", format!("XLSX limit exceeded: {name}"))
}
fn bound(n: u64, limit: Option<u64>, name: &str) -> Result<()> {
    if limit.is_some_and(|limit| n > limit) {
        Err(exceeded(name))
    } else {
        Ok(())
    }
}

// Raw XML tokens: a tag/PI/comment/CDATA/declaration, or a contiguous run
// of text. Check each byte before retaining it; quick-xml never sees an
// over-budget token. Text is counted but not retained here.
#[derive(Default)]
struct TokenScan {
    tag: Vec<u8>,
    size: u64,
    in_tag: bool,
    quote: u8,
}
impl TokenScan {
    fn feed(&mut self, bytes: &[u8], part: &str, c: &ReadControl) -> Result<()> {
        for &b in bytes {
            if !self.in_tag && b == b'<' {
                self.size = 0;
                self.in_tag = true;
                self.quote = 0;
            }
            self.size = self.size.saturating_add(1);
            bound(self.size, c.limits.max_token_bytes, "maxTokenBytes")?;
            if !self.in_tag {
                continue;
            }
            self.tag.push(b);
            let special = self.tag.starts_with(b"<!--") || self.tag.starts_with(b"<![CDATA[");
            if !special && (b == b'\'' || b == b'"') {
                if self.quote == 0 {
                    self.quote = b;
                } else if self.quote == b {
                    self.quote = 0;
                }
            }
            let end = if self.tag.starts_with(b"<!--") {
                self.tag.ends_with(b"-->")
            } else if self.tag.starts_with(b"<![CDATA[") {
                self.tag.ends_with(b"]]>")
            } else {
                b == b'>' && self.quote == 0
            };
            if end {
                self.count(part, c)?;
                self.tag.clear();
                self.in_tag = false;
                self.size = 0;
            }
        }
        Ok(())
    }
    fn count(&self, part: &str, c: &ReadControl) -> Result<()> {
        use quick_xml::events::Event;
        let mut reader = quick_xml::Reader::from_reader(self.tag.as_slice());
        let Ok(Event::Start(tag) | Event::Empty(tag)) = reader.read_event() else {
            return Ok(());
        };
        let name = tag.name();
        if part.contains("xl/worksheets/") && name.as_ref() == "c" {
            c.add(&c.cells, 1, c.limits.max_cells, "maxCells")?;
        }
        if part.ends_with("xl/sharedStrings.xml") && name.as_ref() == "si" {
            c.add(
                &c.shared_strings,
                1,
                c.limits.max_shared_strings,
                "maxSharedStrings",
            )?;
        }
        if part.contains("xl/worksheets/") && name.as_ref() == "dataValidation" {
            for a in tag.attributes().flatten() {
                if a.key.as_ref() != "sqref" {
                    continue;
                }
                let text = a
                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                    .map_err(|e| Error::invalid(part, e.to_string()))?;
                for addr in crate::xform::sheet::split_whitespace_js(&text) {
                    let n = if addr.contains(':') {
                        match crate::col_cache::decode_ex(&addr)
                            .map_err(|e| Error::invalid(part, e))?
                        {
                            crate::col_cache::DecodedEx::Range(r)
                                if r.top != 0.0 && !r.top.is_nan() =>
                            {
                                let n = (r.right - r.left + 1.0).max(0.0)
                                    * (r.bottom - r.top + 1.0).max(0.0);
                                #[allow(clippy::cast_precision_loss)]
                                let max = u64::MAX as f64;
                                if !n.is_finite() || n > max {
                                    return Err(exceeded("maxValidationCells"));
                                }
                                #[allow(clippy::cast_possible_truncation, clippy::cast_sign_loss)]
                                {
                                    n.max(0.0) as u64
                                }
                            }
                            _ => 0,
                        }
                    } else {
                        1
                    };
                    c.add(
                        &c.validation_cells,
                        n,
                        c.limits.max_validation_cells,
                        "maxValidationCells",
                    )?;
                }
            }
        }
        if part.contains("xl/worksheets/") && name.as_ref() == "mergeCell" {
            for a in tag.attributes().flatten() {
                if a.key.as_ref() != "ref" {
                    continue;
                }
                let text = a
                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                    .map_err(|e| Error::invalid(part, e.to_string()))?;
                let n = match crate::col_cache::decode_ex(&text)
                    .map_err(|e| Error::invalid(part, e))?
                {
                    crate::col_cache::DecodedEx::Range(r) => {
                        (r.right - r.left + 1.0).max(0.0) * (r.bottom - r.top + 1.0).max(0.0)
                    }
                    _ => 1.0,
                };
                #[allow(clippy::cast_precision_loss)]
                let max = u64::MAX as f64;
                if !n.is_finite() || n > max {
                    return Err(exceeded("maxMergeCells"));
                }
                #[allow(clippy::cast_possible_truncation, clippy::cast_sign_loss)]
                let n = n as u64;
                c.add(&c.merge_cells, n, c.limits.max_merge_cells, "maxMergeCells")?;
            }
        }
        Ok(())
    }
}

thread_local! { static ACTIVE: RefCell<Option<Arc<ReadControl>>> = const { RefCell::new(None) }; }
#[derive(Debug)]
pub struct ControlScope(Option<Arc<ReadControl>>);
impl ControlScope {
    pub fn enter(control: Option<Arc<ReadControl>>) -> Self {
        Self(ACTIVE.with(|a| a.replace(control)))
    }
}
impl Drop for ControlScope {
    fn drop(&mut self) {
        ACTIVE.with(|a| a.replace(self.0.take()));
    }
}
pub(crate) fn check() -> Result<()> {
    ACTIVE.with(|a| a.borrow().as_ref().map_or(Ok(()), |c| c.check()))
}
pub(crate) fn active() -> Option<Arc<ReadControl>> {
    ACTIVE.with(|a| a.borrow().clone())
}
pub(crate) fn row() {
    ACTIVE.with(|a| {
        if let Some(c) = a.borrow().as_ref() {
            c.rows.fetch_add(1, Ordering::Relaxed);
        }
    });
}
/// Chunked, cancellation-aware decompression used by all parsing passes.
pub(crate) fn read_all(mut input: impl Read) -> Result<Vec<u8>> {
    let mut out = Vec::new();
    if ACTIVE.with(|a| a.borrow().is_none()) {
        input.read_to_end(&mut out)?;
        return Ok(out);
    }
    let mut buf = [0u8; 8192];
    loop {
        check()?;
        let n = input.read(&mut buf)?;
        if n == 0 {
            return Ok(out);
        }
        out.extend_from_slice(buf.get(..n).unwrap_or_default());
    }
}
