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

/// Convenience `Result` alias using [`Error`].
pub type Result<T> = std::result::Result<T, Error>;
