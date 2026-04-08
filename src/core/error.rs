use thiserror::Error;

#[derive(Debug, Error)]
pub enum Error {
    #[error("ZIP error: {0}")]
    Zip(#[from] zip::result::ZipError),
    #[error("XML error: {0}")]
    Xml(#[from] quick_xml::Error),
    #[error("XML attribute error: {0}")]
    XmlAttr(#[from] quick_xml::events::attributes::AttrError),
    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),
    #[error("missing required part: {0}")]
    MissingPart(String),
    #[error("missing required attribute '{attr}' on element <{element}>")]
    MissingAttribute { element: String, attr: String },
    #[error("invalid part name: {0}")]
    InvalidPartName(String),
    #[error("unknown content type for part: {0}")]
    UnknownContentType(String),
    #[error("relationship not found: {0}")]
    RelationshipNotFound(String),
    #[error("malformed XML: {0}")]
    MalformedXml(String),
    #[error("unsupported feature: {0}")]
    Unsupported(String),
    #[error("integer parse error: {0}")]
    ParseInt(#[from] std::num::ParseIntError),
    #[error("float parse error: {0}")]
    ParseFloat(#[from] std::num::ParseFloatError),
    #[error("UTF-8 error: {0}")]
    Utf8(#[from] std::str::Utf8Error),
}

pub type Result<T> = std::result::Result<T, Error>;
