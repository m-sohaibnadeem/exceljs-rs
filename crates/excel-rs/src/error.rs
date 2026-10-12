//! The crate's error type.

use std::fmt;
use std::path::PathBuf;

/// Result alias used throughout the crate.
pub type Result<T, E = Error> = std::result::Result<T, E>;

/// Everything that can go wrong while reading a workbook.
///
/// Malformed input always becomes one of these variants; the reader never
/// panics. Where ExcelJS rejects a file, excel-rs returns an `Err` (the exact
/// ExcelJS message is not reproduced).
#[derive(Debug)]
#[non_exhaustive]
pub enum Error {
    /// `readFile` was given a path that does not exist.
    /// ExcelJS: `File not found: <filename>`.
    FileNotFound(PathBuf),
    /// Reading from the filesystem failed.
    Io(std::io::Error),
    /// The data is not a readable ZIP archive.
    Zip(zip::result::ZipError),
    /// A package part is not well-formed XML (or not UTF-8).
    Xml {
        /// ZIP entry name of the part, e.g. `xl/worksheets/sheet1.xml`.
        part: String,
        /// What the XML parser reported.
        message: String,
    },
    /// The XML is well-formed but has a shape ExcelJS cannot load.
    InvalidPart {
        /// ZIP entry name of the part.
        part: String,
        /// What is wrong.
        message: String,
    },
    /// A `t="s"` cell refers to a shared string that does not exist.
    SharedStringIndex {
        /// ZIP entry name of the worksheet part.
        part: String,
        /// Cell reference as written in the file.
        cell: String,
        /// The raw `<v>` text.
        index: String,
    },
    /// A cell reference cannot be decoded (no column letters, or past `XFD`).
    CellReference {
        /// ZIP entry name of the worksheet part.
        part: String,
        /// The reference as written in the file.
        reference: String,
    },
}

impl Error {
    /// The message ExcelJS (or the library it uses) gives for this failure,
    /// without the part name: what the Node API rejects with.
    pub fn exceljs_message(&self) -> String {
        match self {
            Error::Xml { message, .. } | Error::InvalidPart { message, .. } => message.clone(),
            Error::Zip(zip::result::ZipError::InvalidArchive(_)) => {
                "Can't find end of central directory : is this a zip file ? If it is, see https://stuk.github.io/jszip/documentation/howto/read_zip.html".to_owned()
            }
            other => other.to_string(),
        }
    }

    pub(crate) fn invalid(part: &str, message: impl Into<String>) -> Self {
        Error::InvalidPart {
            part: part.to_owned(),
            message: message.into(),
        }
    }
}

impl fmt::Display for Error {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Error::FileNotFound(path) => write!(f, "File not found: {}", path.display()),
            Error::Io(e) => write!(f, "I/O error: {e}"),
            Error::Zip(e) => write!(f, "not a readable .xlsx (ZIP) file: {e}"),
            Error::Xml { part, message } => write!(f, "malformed XML in {part}: {message}"),
            Error::InvalidPart { part, message } => write!(f, "invalid {part}: {message}"),
            Error::SharedStringIndex { part, cell, index } => write!(
                f,
                "{part}: cell {cell} refers to shared string {index:?}, which does not exist"
            ),
            Error::CellReference { part, reference } => {
                write!(f, "{part}: invalid cell reference {reference:?}")
            }
        }
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Error::Io(e) => Some(e),
            Error::Zip(e) => Some(e),
            _ => None,
        }
    }
}

impl From<std::io::Error> for Error {
    fn from(e: std::io::Error) -> Self {
        Error::Io(e)
    }
}

impl From<zip::result::ZipError> for Error {
    fn from(e: zip::result::ZipError) -> Self {
        Error::Zip(e)
    }
}
