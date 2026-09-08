use thiserror::Error;

/// Core error type for OPC/XML operations.
#[derive(Debug, Error)]
pub enum Error {
    /// ZIP archive error.
    #[error("ZIP error: {0}")]
    Zip(#[from] zip::result::ZipError),
    /// XML parse error.
    #[error("XML error: {0}")]
    Xml(#[from] quick_xml::Error),
    /// XML attribute parse error.
    #[error("XML attribute error: {0}")]
    XmlAttr(#[from] quick_xml::events::attributes::AttrError),
    /// I/O error.
    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),
    /// A required OPC part is absent from the ZIP.
    #[error("missing required part: {0}")]
    MissingPart(String),
    /// A required XML attribute is absent from an element.
    #[error("missing required attribute '{attr}' on element <{element}>")]
    MissingAttribute {
        /// The element local name.
        element: String,
        /// The attribute name.
        attr: String,
    },
    /// An OPC part name fails validation.
    #[error("invalid part name: {0}")]
    InvalidPartName(String),
    /// The content type for a part is not recognized.
    #[error("unknown content type for part: {0}")]
    UnknownContentType(String),
    /// A relationship ID or type was not found.
    #[error("relationship not found: {0}")]
    RelationshipNotFound(String),
    /// The XML structure is not as expected.
    #[error("malformed XML: {0}")]
    MalformedXml(String),
    /// A feature present in the file is not supported by this library.
    #[error("unsupported feature: {0}")]
    Unsupported(String),
    /// The package contains two entries with the same part name. Two
    /// readers that disagree about which copy wins see two different
    /// documents from the same bytes, which is how content is slipped past
    /// a scanner that reads one copy while the renderer reads the other.
    #[error("duplicate part name in package: {0}")]
    DuplicatePart(String),
    /// The package's primary part is not the type the caller asked for —
    /// e.g. an XLSX opened as a DOCX.
    #[error("format mismatch: package holds {found}, expected {expected}")]
    FormatMismatch {
        /// What the package actually contains.
        found: String,
        /// What the caller asked for.
        expected: String,
    },
    /// An XML part ends before its root element is closed.
    #[error("truncated XML part: '{0}' ends before its root element is closed")]
    TruncatedPart(String),
    /// A part expands to more than the decompression limit allows. Raised
    /// before the memory is allocated, so a malicious archive cannot be used
    /// to exhaust the host's memory.
    #[error("decompression limit exceeded: part '{part}' expands to more than {limit} bytes")]
    DecompressionLimit {
        /// The offending part name.
        part: String,
        /// The limit that was exceeded, in bytes.
        limit: u64,
    },
    /// Integer parse error.
    #[error("integer parse error: {0}")]
    ParseInt(#[from] std::num::ParseIntError),
    /// Floating-point parse error.
    #[error("float parse error: {0}")]
    ParseFloat(#[from] std::num::ParseFloatError),
    /// UTF-8 decode error.
    #[error("UTF-8 error: {0}")]
    Utf8(#[from] std::str::Utf8Error),
}

/// Convenience `Result` alias using [`enum@Error`].
pub type Result<T> = std::result::Result<T, Error>;
