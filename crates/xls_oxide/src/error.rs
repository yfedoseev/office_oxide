/// Errors when reading legacy XLS files.
#[derive(Debug, thiserror::Error)]
pub enum XlsError {
    #[error("CFB error: {0}")]
    Cfb(#[from] cfb_oxide::CfbError),

    #[error("invalid BIFF record: {0}")]
    InvalidRecord(String),

    #[error("unsupported BIFF version: {0}")]
    UnsupportedVersion(u16),

    #[error("missing stream: {0}")]
    MissingStream(String),

    #[error("corrupted data: {0}")]
    Corrupted(String),

    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),
}

pub type Result<T> = std::result::Result<T, XlsError>;
