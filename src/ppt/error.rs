/// Errors when reading legacy PPT files.
#[derive(Debug, thiserror::Error)]
pub enum PptError {
    #[error("CFB error: {0}")]
    Cfb(#[from] crate::cfb::CfbError),

    #[error("invalid record: {0}")]
    InvalidRecord(String),

    #[error("missing stream: {0}")]
    MissingStream(String),

    #[error("corrupted data: {0}")]
    Corrupted(String),

    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),
}

pub type Result<T> = std::result::Result<T, PptError>;
