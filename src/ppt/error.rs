/// Errors when reading legacy PPT files.
#[derive(Debug, thiserror::Error)]
pub enum PptError {
    /// Error from the underlying CFB container reader.
    #[error("CFB error: {0}")]
    Cfb(#[from] crate::cfb::CfbError),

    /// A binary record has an unexpected length or content.
    #[error("invalid record: {0}")]
    InvalidRecord(String),

    /// A required CFB stream is absent from the file.
    #[error("missing stream: {0}")]
    MissingStream(String),

    /// The data is internally inconsistent or truncated.
    #[error("corrupted data: {0}")]
    Corrupted(String),

    /// Underlying I/O error.
    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),
}

/// Convenience `Result` alias using `PptError`.
pub type Result<T> = std::result::Result<T, PptError>;
