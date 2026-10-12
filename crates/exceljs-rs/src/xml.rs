//! A small SAX-style layer over quick-xml that produces the same event stream
//! ExcelJS gets from `saxes` (see ExcelJS `lib/utils/parse-sax.js`):
//!
//! * element names are qualified names exactly as written (`x:c` is not `c`);
//! * `<a/>` is reported as an open followed by a close;
//! * text has entity and character references resolved and line endings
//!   normalised; adjacent pieces may arrive as several `Text` events;
//! * CDATA sections are text. (ExcelJS registers no CDATA handler with saxes
//!   and drops them: B8 in COMPATIBILITY.md.)

use std::borrow::Cow;

use quick_xml::escape::resolve_predefined_entity;
use quick_xml::events::Event;
use quick_xml::{Reader, XmlVersion};

use crate::error::{Error, Result};

/// An opening tag and its attributes (values unescaped and normalised).
#[derive(Debug)]
pub(crate) struct Tag<'a> {
    pub name: &'a str,
    attrs: Vec<(&'a str, Cow<'a, str>)>,
}

impl Tag<'_> {
    /// All attributes in document order (`node.attributes`).
    pub fn attributes(&self) -> impl Iterator<Item = (&str, &str)> {
        self.attrs.iter().map(|(k, v)| (*k, v.as_ref()))
    }

    /// `node.attributes` as a JSON object (later duplicates win, as in JS).
    pub fn attributes_object(&self) -> serde_json::Value {
        let mut m = serde_json::Map::new();
        for (k, v) in self.attributes() {
            m.insert(k.to_owned(), serde_json::Value::String(v.to_owned()));
        }
        serde_json::Value::Object(m)
    }

    /// Value of the attribute with this qualified name, if present.
    pub fn attr(&self, name: &str) -> Option<&str> {
        self.attrs
            .iter()
            .find(|(key, _)| *key == name)
            .map(|(_, value)| value.as_ref())
    }
}

/// One parser event.
#[derive(Debug)]
pub(crate) enum Sax<'e, 'a> {
    Open(&'e Tag<'a>),
    Text(&'e str),
    Close(&'e str),
}

/// Parses `bytes` (the content of ZIP entry `part`) and feeds every event to
/// `handler`. `handler` returns `Ok(false)` to stop early, as an ExcelJS
/// xform does when its root element closes.
pub(crate) fn parse(
    part: &str,
    bytes: &[u8],
    handler: impl FnMut(Sax<'_, '_>) -> Result<bool>,
) -> Result<()> {
    let doc = std::str::from_utf8(bytes).map_err(|e| xml_error(part, e))?;
    let doc = doc.strip_prefix('\u{FEFF}').unwrap_or(doc);
    let mut reader = Reader::from_str(doc);
    reader.config_mut().expand_empty_elements = true;
    let line_col = |pos: usize| -> (usize, usize) {
        let before = doc.get(..pos.min(doc.len())).unwrap_or("");
        let line = before.matches('\n').count() + 1;
        let col = before.rsplit('\n').next().map_or(0, |l| l.chars().count()) + 1;
        (line, col)
    };
    run(part, &mut SliceSource(reader), &line_col, handler)
}

/// [`parse`] of a part read from `input` as it is parsed (so it is never
/// whole in memory), with invalid UTF-8 replaced as [`parse`]'s callers do
/// for a whole part (`String::from_utf8_lossy`). `line_col` gives the
/// line and column of a byte offset of the decoded text, for the one error
/// that reports a position (it may read the part again).
pub(crate) fn parse_stream(
    part: &str,
    input: impl std::io::Read,
    line_col: &dyn Fn(usize) -> (usize, usize),
    handler: impl FnMut(Sax<'_, '_>) -> Result<bool>,
) -> Result<()> {
    let mut reader = Reader::from_reader(std::io::BufReader::with_capacity(
        1 << 16,
        LossyUtf8::new(input),
    ));
    reader.config_mut().expand_empty_elements = true;
    run(
        part,
        &mut BufSource {
            reader,
            buf: Vec::new(),
        },
        line_col,
        handler,
    )
}

/// Line and column (saxes-style, 1-based) of byte offset `pos` of `doc`,
/// after a leading BOM is dropped.
pub(crate) fn line_col_of(doc: &str, pos: usize) -> (usize, usize) {
    let doc = doc.strip_prefix('\u{FEFF}').unwrap_or(doc);
    let before = doc.get(..pos.min(doc.len())).unwrap_or("");
    let line = before.matches('\n').count() + 1;
    let col = before.rsplit('\n').next().map_or(0, |l| l.chars().count()) + 1;
    (line, col)
}

/// Where [`run`] gets its events from.
trait Source {
    fn next_event(&mut self) -> quick_xml::Result<Event<'_>>;
    fn position(&self) -> u64;
}

struct SliceSource<'d>(Reader<&'d [u8]>);

impl Source for SliceSource<'_> {
    fn next_event(&mut self) -> quick_xml::Result<Event<'_>> {
        self.0.read_event()
    }
    fn position(&self) -> u64 {
        self.0.buffer_position()
    }
}

struct BufSource<R: std::io::BufRead> {
    reader: Reader<R>,
    buf: Vec<u8>,
}

impl<R: std::io::BufRead> Source for BufSource<R> {
    fn next_event(&mut self) -> quick_xml::Result<Event<'_>> {
        self.buf.clear();
        self.reader.read_event_into(&mut self.buf)
    }
    fn position(&self) -> u64 {
        self.reader.buffer_position()
    }
}

/// UTF-8 from `R` with invalid sequences replaced by U+FFFD (as
/// `String::from_utf8_lossy` on the whole input), a leading BOM dropped.
struct LossyUtf8<R> {
    inner: R,
    /// Bytes read but not given out yet (an incomplete character at the end
    /// of a read is kept for the next).
    pending: Vec<u8>,
    out: Vec<u8>,
    out_at: usize,
    start: bool,
    eof: bool,
}

impl<R: std::io::Read> LossyUtf8<R> {
    fn new(inner: R) -> Self {
        LossyUtf8 {
            inner,
            pending: Vec::new(),
            out: Vec::new(),
            out_at: 0,
            start: true,
            eof: false,
        }
    }

    /// Fills `out` with the next decoded piece; false at the end.
    fn refill(&mut self) -> std::io::Result<bool> {
        self.out.clear();
        self.out_at = 0;
        while self.out.is_empty() {
            if !self.eof {
                let mut chunk = [0u8; 1 << 16];
                let n = self.inner.read(&mut chunk)?;
                if n == 0 {
                    self.eof = true;
                } else {
                    self.pending
                        .extend_from_slice(chunk.get(..n).unwrap_or_default());
                }
            }
            if self.eof && self.pending.is_empty() {
                return Ok(false);
            }
            self.decode_pending();
        }
        if self.start {
            self.start = false;
            if self.out.starts_with("\u{FEFF}".as_bytes()) {
                self.out_at = 3;
            }
        }
        Ok(true)
    }

    /// Moves `pending` to `out` as `from_utf8_lossy` would (valid runs, one
    /// U+FFFD per invalid sequence), keeping an incomplete character at the
    /// end unless the input is over.
    fn decode_pending(&mut self) {
        let mut at = 0;
        loop {
            let rest = self.pending.get(at..).unwrap_or_default();
            match std::str::from_utf8(rest) {
                Ok(s) => {
                    self.out.extend_from_slice(s.as_bytes());
                    at = self.pending.len();
                    break;
                }
                Err(e) => {
                    let valid = rest.get(..e.valid_up_to()).unwrap_or_default();
                    self.out.extend_from_slice(valid);
                    match e.error_len() {
                        Some(len) => {
                            self.out.extend_from_slice("\u{FFFD}".as_bytes());
                            at += e.valid_up_to() + len;
                        }
                        None if self.eof => {
                            self.out.extend_from_slice("\u{FFFD}".as_bytes());
                            at = self.pending.len();
                            break;
                        }
                        None => {
                            at += e.valid_up_to();
                            break;
                        }
                    }
                }
            }
        }
        self.pending.drain(..at);
    }
}

impl<R: std::io::Read> std::io::Read for LossyUtf8<R> {
    fn read(&mut self, buf: &mut [u8]) -> std::io::Result<usize> {
        while self.out_at >= self.out.len() {
            if !self.refill()? {
                return Ok(0);
            }
        }
        let avail = self.out.get(self.out_at..).unwrap_or_default();
        let n = avail.len().min(buf.len());
        if let (Some(dst), Some(src)) = (buf.get_mut(..n), avail.get(..n)) {
            dst.copy_from_slice(src);
        }
        self.out_at += n;
        Ok(n)
    }
}

/// The parse loop of [`parse`] and [`parse_stream`].
fn run(
    part: &str,
    reader: &mut dyn Source,
    line_col: &dyn Fn(usize) -> (usize, usize),
    mut handler: impl FnMut(Sax<'_, '_>) -> Result<bool>,
) -> Result<()> {
    // saxes reports text between two pieces of markup as one event, with
    // references decoded inline, so text is gathered here and flushed when
    // the next markup (tag, CDATA, comment, PI, ...) or the end arrives.
    let mut text = String::new();
    // saxes is a conforming parser: one root element, no text outside it.
    let mut depth: usize = 0;
    let mut seen_root = false;
    // Byte offset where pending text ended (for saxes-style line:column).
    let mut text_end: usize = 0;
    let flush = |text: &mut String,
                 depth: usize,
                 at: usize,
                 handler: &mut dyn FnMut(Sax<'_, '_>) -> Result<bool>|
     -> Result<bool> {
        if text.is_empty() {
            return Ok(true);
        }
        if depth == 0 {
            if text.chars().any(|c| !matches!(c, ' ' | '\t' | '\n' | '\r')) {
                let (line, col) = line_col(at);
                return Err(xml_error(
                    part,
                    format!("{line}:{col}: text data outside of root node."),
                ));
            }
            text.clear();
            return Ok(true);
        }
        let go_on = handler(Sax::Text(text))?;
        text.clear();
        Ok(go_on)
    };
    loop {
        let event = reader.next_event().map_err(|e| xml_error(part, e))?;
        match &event {
            Event::Text(t) => {
                text.push_str(&t.xml10_content());
                text_end = usize::try_from(reader.position()).unwrap_or(usize::MAX);
                continue;
            }
            Event::CData(c) => {
                // Line endings are normalised in CDATA too (XML 1.0 §2.11).
                let raw: &str = c;
                text.push_str(&raw.replace("\r\n", "\n").replace('\r', "\n"));
                text_end = usize::try_from(reader.position()).unwrap_or(usize::MAX);
                continue;
            }
            Event::GeneralRef(r) => {
                let name = r.xml10_content();
                match r.resolve_char_ref().map_err(|e| xml_error(part, e))? {
                    Some(c) => text.push(c),
                    None => text
                        .push_str(resolve_predefined_entity(&name).ok_or_else(|| {
                            xml_error(part, format!("undefined entity &{name};"))
                        })?),
                }
                continue;
            }
            _ => {}
        }
        if !flush(&mut text, depth, text_end, &mut handler)? {
            return Ok(());
        }
        let keep_going = match event {
            // `Empty` does not occur with `expand_empty_elements`, but if it
            // did it would mean open-then-close, as saxes reports it.
            Event::Start(ref start) | Event::Empty(ref start) => {
                if depth == 0 {
                    if seen_root {
                        return Err(xml_error(
                            part,
                            "document must contain only one root element",
                        ));
                    }
                    seen_root = true;
                }
                if matches!(event, Event::Start(_)) {
                    depth += 1;
                }
                let mut attrs = Vec::new();
                for attr in start.attributes() {
                    let attr = attr.map_err(|e| xml_error(part, e))?;
                    let value = attr
                        .normalized_value(XmlVersion::Implicit1_0)
                        .map_err(|e| xml_error(part, e))?;
                    attrs.push((attr.key.0, value));
                }
                let tag = Tag {
                    name: start.name().0,
                    attrs,
                };
                // `$` cannot start an XML name (saxes rejects it); the model
                // codec relies on no element or attribute name starting with it.
                if tag.name.starts_with('$') || tag.attrs.iter().any(|(k, _)| k.starts_with('$')) {
                    return Err(xml_error(part, "invalid name start character '$'"));
                }
                let open = handler(Sax::Open(&tag))?;
                if open && matches!(event, Event::Empty(_)) {
                    handler(Sax::Close(tag.name))?
                } else {
                    open
                }
            }
            Event::End(end) => {
                depth = depth.saturating_sub(1);
                handler(Sax::Close(end.name().0))?
            }
            Event::Eof => return Ok(()),
            Event::Text(_)
            | Event::GeneralRef(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_) => true,
        };
        if !keep_going {
            return Ok(());
        }
    }
}

fn xml_error(part: &str, e: impl std::fmt::Display) -> Error {
    Error::Xml {
        part: part.to_owned(),
        message: e.to_string(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn events(xml: &str) -> Result<Vec<String>> {
        let mut out = Vec::new();
        parse("test.xml", xml.as_bytes(), |e| {
            out.push(match e {
                Sax::Open(t) => format!("<{}{}>", t.name, t.attr("r").unwrap_or("")),
                Sax::Text(s) => format!("{s:?}"),
                Sax::Close(n) => format!("</{n}>"),
            });
            Ok(true)
        })?;
        Ok(out)
    }

    #[test]
    fn mirrors_saxes_events() {
        let got = events("<a r=\"x &amp; y\"><b/>t&lt;&#65;<![CDATA[kept]]>\r\n</a>").unwrap();
        assert_eq!(got, ["<ax & y>", "<b>", "</b>", "\"t<Akept\\n\"", "</a>"]);
    }

    #[test]
    fn malformed_xml_is_an_error_not_a_panic() {
        assert!(events("<a></b>").is_err());
        assert!(events("<a>&nope;</a>").is_err());
        assert!(parse("x", b"\xff\xfe", |_| Ok(true)).is_err());
    }

    /// Reads at most `n` bytes at a time.
    struct Trickle<'a>(&'a [u8], usize);

    impl std::io::Read for Trickle<'_> {
        fn read(&mut self, buf: &mut [u8]) -> std::io::Result<usize> {
            let n = self.1.min(buf.len()).min(self.0.len());
            let (head, tail) = self.0.split_at(n);
            buf.get_mut(..n).unwrap_or_default().copy_from_slice(head);
            self.0 = tail;
            Ok(n)
        }
    }

    fn stream_events(bytes: &[u8], chunk: usize) -> std::result::Result<Vec<String>, String> {
        let mut out = Vec::new();
        let lossy = String::from_utf8_lossy(bytes).into_owned();
        parse_stream(
            "test.xml",
            Trickle(bytes, chunk),
            &|pos| line_col_of(&lossy, pos),
            |e| {
                out.push(match e {
                    Sax::Open(t) => format!("<{}{}>", t.name, t.attr("r").unwrap_or("")),
                    Sax::Text(s) => format!("{s:?}"),
                    Sax::Close(n) => format!("</{n}>"),
                });
                Ok(true)
            },
        )
        .map_err(|e| e.to_string())?;
        Ok(out)
    }

    #[test]
    fn stream_parse_matches_parse() {
        let docs: Vec<Vec<u8>> = vec![
            "<a r=\"x &amp; y\"><b/>t&lt;&#65;<![CDATA[kept]]>\r\n</a>".into(),
            "\u{FEFF}<a>é日本語 😀 text</a>".into(),
            b"<a>bad \xff\xfe byte and cut \xe6\x97</a>".to_vec(),
            "<a>x</a>\n\n  tail".into(),
            "<a><b r='1'/><c>é</c></a>".into(),
        ];
        for doc in &docs {
            let lossy = String::from_utf8_lossy(doc).into_owned();
            let whole = events(&lossy).map_err(|e| e.to_string());
            for chunk in [1, 2, 3, 7, 1 << 16] {
                assert_eq!(
                    stream_events(doc, chunk),
                    whole,
                    "{lossy:?} in chunks of {chunk}"
                );
            }
        }
    }
}
