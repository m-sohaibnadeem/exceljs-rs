//! CSV tokenizing: a port of `@fast-csv/parse` 4.3.6 (`Parser`,
//! `RowParser`, `ColumnParser`, `Scanner`), which ExcelJS's `CSV.read` uses.
//! Row handling after tokenizing (headers, skipLines, maxRows, ...) stays in
//! JS because it calls user functions.

use crate::js::is_js_whitespace;

/// The parser options the tokenizer looks at (`ParserOptions`).
#[derive(Debug, Clone)]
pub struct CsvParseOptions {
    /// One character, or empty.
    pub delimiter: String,
    /// `quote`; `None` when null.
    pub quote: Option<String>,
    /// `escape ?? quote`.
    pub escape_char: Option<String>,
    /// `comment`; `None` when null or undefined.
    pub comment: Option<String>,
    pub ltrim: bool,
    pub rtrim: bool,
    pub trim: bool,
    pub ignore_empty: bool,
    /// `carriageReturn` (`"\r"`).
    pub carriage_return: String,
}

#[derive(Clone, Copy, Debug)]
enum Text {
    Char(char),
    Crlf,
}

#[derive(Clone, Copy, Debug)]
struct Token {
    start: usize,
    end: usize,
    text: Text,
}

impl Token {
    fn is(&self, s: Option<&str>) -> bool {
        let Some(s) = s else { return false };
        match self.text {
            Text::Char(c) => {
                let mut it = s.chars();
                it.next() == Some(c) && it.next().is_none()
            }
            Text::Crlf => s == "\r\n",
        }
    }
    fn is_row_delimiter(&self) -> bool {
        matches!(self.text, Text::Crlf | Text::Char('\r' | '\n'))
    }
    fn string(&self) -> String {
        let mut s = String::new();
        self.push_to(&mut s);
        s
    }
    fn push_to(&self, s: &mut String) {
        match self.text {
            Text::Char(c) => s.push(c),
            Text::Crlf => s.push_str("\r\n"),
        }
    }
}

struct Scanner<'a> {
    chars: &'a [char],
    /// Start of `line` (moved by `truncateToCursor`).
    base: usize,
    cursor: usize,
    has_more_data: bool,
    opts: &'a CsvParseOptions,
    delimiter: Option<char>,
}

impl Scanner<'_> {
    fn has_more_characters(&self) -> bool {
        self.chars.len() > self.cursor
    }

    /// `NEXT_TOKEN_REGEXP = /([^\s]|\r\n|\n|\r|<delimiter>)/`.
    fn next_non_space_token(&self) -> Option<Token> {
        let mut i = self.cursor;
        while let Some(&c) = self.chars.get(i) {
            if !is_js_whitespace(c) {
                return Some(Token {
                    start: i,
                    end: i,
                    text: Text::Char(c),
                });
            }
            if c == '\r' && self.chars.get(i + 1) == Some(&'\n') {
                return Some(Token {
                    start: i,
                    end: i + 1,
                    text: Text::Crlf,
                });
            }
            if c == '\n' || c == '\r' || Some(c) == self.delimiter {
                return Some(Token {
                    start: i,
                    end: i,
                    text: Text::Char(c),
                });
            }
            i += 1;
        }
        None
    }

    fn next_character_token(&self) -> Option<Token> {
        self.chars.get(self.cursor).map(|&c| Token {
            start: self.cursor,
            end: self.cursor,
            text: Text::Char(c),
        })
    }

    fn line_from_cursor(&self) -> String {
        self.chars
            .get(self.cursor..)
            .unwrap_or(&[])
            .iter()
            .collect()
    }

    fn line(&self) -> String {
        self.chars.get(self.base..).unwrap_or(&[]).iter().collect()
    }

    /// `advancePastLine`; false is its `null`.
    fn advance_past_line(&mut self) -> bool {
        let rest = self.chars.get(self.cursor..).unwrap_or(&[]);
        match rest.iter().position(|&c| c == '\r' || c == '\n') {
            None => {
                if self.has_more_data {
                    return false;
                }
                self.cursor = self.chars.len();
            }
            Some(i) => {
                let crlf = rest.get(i) == Some(&'\r') && rest.get(i + 1) == Some(&'\n');
                self.cursor += i + if crlf { 2 } else { 1 };
            }
        }
        true
    }

    fn advance_past(&mut self, t: &Token) {
        self.cursor = t.end + 1;
    }

    fn truncate(&mut self) {
        self.base = self.cursor;
    }

    fn is_delimiter(&self, t: &Token) -> bool {
        t.is(Some(self.opts.delimiter.as_str()))
    }
    fn is_quote(&self, t: &Token) -> bool {
        t.is(self.opts.quote.as_deref())
    }
    fn is_escape(&self, t: &Token) -> bool {
        t.is(self.opts.escape_char.as_deref())
    }
}

fn js_trim_start(s: &str) -> &str {
    s.trim_start_matches(is_js_whitespace)
}
fn js_trim_end(s: &str) -> &str {
    s.trim_end_matches(is_js_whitespace)
}

/// `ColumnFormatter`.
fn format_column(opts: &CsvParseOptions, col: String) -> String {
    if opts.trim {
        js_trim_end(js_trim_start(&col)).to_owned()
    } else if opts.ltrim {
        js_trim_start(&col).to_owned()
    } else if opts.rtrim {
        js_trim_end(&col).to_owned()
    } else {
        col
    }
}

/// `lodash.escapeRegExp`.
fn escape_regexp(s: &str) -> String {
    let mut out = String::new();
    for c in s.chars() {
        if "^$.*+?()[]{}|\\".contains(c) {
            out.push('\\');
        }
        out.push(c);
    }
    out
}

fn preview(s: &str) -> String {
    s.chars()
        .map(|c| {
            if c == '\r' || c == '\n' {
                "\\n'".to_owned()
            } else {
                c.to_string()
            }
        })
        .collect()
}

/// `NonQuotedColumnParser.parse`.
fn parse_non_quoted(s: &mut Scanner<'_>) -> Option<String> {
    if !s.has_more_characters() {
        return None;
    }
    let mut col = String::new();
    while let Some(t) = s.next_character_token() {
        if s.is_delimiter(&t) || t.is_row_delimiter() {
            break;
        }
        t.push_to(&mut col);
        s.advance_past(&t);
    }
    Some(format_column(s.opts, col))
}

/// `QuotedColumnParser.parse`.
fn parse_quoted(s: &mut Scanner<'_>) -> Result<Option<String>, String> {
    if !s.has_more_characters() {
        return Ok(None);
    }
    let original = s.cursor;
    let mut found_start = false;
    let mut found_close = false;
    let mut col = String::new();
    let mut next = s.next_character_token();
    while !found_close {
        let Some(mut tok) = next else { break };
        let is_quote = s.is_quote(&tok);
        if !found_start && is_quote {
            found_start = true;
        } else if found_start {
            if s.is_escape(&tok) {
                s.advance_past(&tok);
                let following = s.next_character_token();
                match following {
                    Some(f) if s.is_quote(&f) || s.is_escape(&f) => {
                        f.push_to(&mut col);
                        tok = f;
                    }
                    _ if is_quote => found_close = true,
                    _ => tok.push_to(&mut col),
                }
            } else if is_quote {
                found_close = true;
            } else {
                tok.push_to(&mut col);
            }
        }
        s.advance_past(&tok);
        next = s.next_character_token();
    }
    let col = format_column(s.opts, col);
    if !found_close {
        s.cursor = original;
        if !s.has_more_data {
            return Err(format!(
                "Parse Error: missing closing: '{}' in line: at '{}'",
                s.opts.quote.as_deref().unwrap_or(""),
                preview(&s.line_from_cursor())
            ));
        }
        return Ok(None);
    }
    // checkForMalformedColumn
    match s.next_non_space_token() {
        Some(t) => {
            if !(s.is_delimiter(&t) || t.is_row_delimiter()) {
                let p: String = s.line_from_cursor().chars().take(10).collect();
                return Err(format!(
                    "Parse Error: expected: '{}' OR new line got: '{}'. at '{}",
                    escape_regexp(&s.opts.delimiter),
                    t.string(),
                    preview(&p)
                ));
            }
            s.cursor = t.start;
        }
        None => {
            if !s.has_more_data {
                s.advance_past_line();
            }
        }
    }
    Ok(Some(col))
}

/// `ColumnParser.parse`.
fn parse_column(s: &mut Scanner<'_>) -> Result<Option<String>, String> {
    match s.next_non_space_token() {
        Some(t) if s.is_quote(&t) => {
            s.cursor = t.start;
            parse_quoted(s)
        }
        _ => Ok(parse_non_quoted(s)),
    }
}

/// `RowParser.parse`; `None` is its `null`.
fn parse_row_columns(s: &mut Scanner<'_>) -> Result<Option<Vec<String>>, String> {
    let mut columns = Vec::new();
    // getStartToken
    let mut current = s.next_non_space_token();
    if let Some(t) = current {
        if s.is_delimiter(&t) {
            columns.push(String::new());
            current = s.next_non_space_token();
        }
    }
    while let Some(t) = current {
        if t.is_row_delimiter() {
            s.advance_past(&t);
            if !s.has_more_characters()
                && t.is(Some(s.opts.carriage_return.as_str()))
                && s.has_more_data
            {
                return Ok(None);
            }
            s.truncate();
            return Ok(Some(columns));
        }
        // shouldSkipColumnParse
        let mut skip = false;
        if s.is_delimiter(&t) {
            s.advance_past(&t);
            let next = s.next_character_token();
            if !s.has_more_characters()
                || next.is_some_and(|n| n.is_row_delimiter() || s.is_delimiter(&n))
            {
                columns.push(String::new());
                skip = true;
            }
        }
        if !skip {
            match parse_column(s)? {
                Some(item) => columns.push(item),
                None => return Ok(None),
            }
        }
        current = s.next_non_space_token();
    }
    if !s.has_more_data {
        s.truncate();
        return Ok(Some(columns));
    }
    Ok(None)
}

fn is_empty_row(row: &[String]) -> bool {
    row.iter().all(|c| c.chars().all(is_js_whitespace))
}

/// `Parser.parseRow`: false stops parsing.
fn parse_row(s: &mut Scanner<'_>, rows: &mut Vec<Vec<String>>) -> Result<bool, String> {
    if s.next_non_space_token().is_none() {
        return Ok(false);
    }
    let Some(row) = parse_row_columns(s)? else {
        return Ok(false);
    };
    if s.opts.ignore_empty && is_empty_row(&row) {
        return Ok(true);
    }
    rows.push(row);
    Ok(true)
}

/// `Parser.parse(line, hasMoreData)`: the complete rows and the unparsed
/// rest of `line`.
pub fn parse_csv(
    line: &str,
    options: &CsvParseOptions,
    has_more_data: bool,
) -> Result<(Vec<Vec<String>>, String), String> {
    let line = line.strip_prefix('\u{FEFF}').unwrap_or(line);
    let chars: Vec<char> = line.chars().collect();
    let mut delims = options.delimiter.chars();
    let delimiter = match (delims.next(), delims.next()) {
        (Some(c), None) => Some(c),
        _ => None,
    };
    let mut s = Scanner {
        chars: &chars,
        base: 0,
        cursor: 0,
        has_more_data,
        opts: options,
        delimiter,
    };
    let mut rows = Vec::new();
    if options.comment.is_none() {
        while parse_row(&mut s, &mut rows)? {}
        return Ok((rows, s.line()));
    }
    while let Some(t) = s.next_character_token() {
        if t.is(options.comment.as_deref()) {
            if !s.advance_past_line() || !s.has_more_characters() {
                return Ok((rows, s.line_from_cursor()));
            }
            s.truncate();
        } else if !parse_row(&mut s, &mut rows)? {
            break;
        }
    }
    Ok((rows, s.line()))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn opts() -> CsvParseOptions {
        CsvParseOptions {
            delimiter: ",".into(),
            quote: Some("\"".into()),
            escape_char: Some("\"".into()),
            comment: None,
            ltrim: false,
            rtrim: false,
            trim: false,
            ignore_empty: false,
            carriage_return: "\r".into(),
        }
    }

    fn all(text: &str, o: &CsvParseOptions) -> Result<Vec<Vec<String>>, String> {
        let (mut rows, rest) = parse_csv(text, o, true)?;
        let (more, _) = parse_csv(&rest, o, false)?;
        rows.extend(more);
        Ok(rows)
    }

    #[test]
    fn parses_like_fast_csv() {
        let rows = all("a,b\r\n\"x \"\"q\"\"\",,c\n,d", &opts()).unwrap();
        assert_eq!(
            rows,
            vec![vec!["a", "b"], vec!["x \"q\"", "", "c"], vec!["", "d"]]
        );
        let rows = all(
            "1\t 2 \n",
            &CsvParseOptions {
                delimiter: "\t".into(),
                ..opts()
            },
        )
        .unwrap();
        assert_eq!(rows, vec![vec!["1", " 2 "]]);
        assert!(all("\"abc", &opts())
            .unwrap_err()
            .starts_with("Parse Error: missing closing"));
        assert!(all("\"a\"b,c", &opts())
            .unwrap_err()
            .starts_with("Parse Error: expected: ','"));
        let rows = all(
            "#c\na\n",
            &CsvParseOptions {
                comment: Some("#".into()),
                ..opts()
            },
        )
        .unwrap();
        assert_eq!(rows, vec![vec!["a"]]);
    }
}
