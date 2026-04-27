/// Errors when reading legacy XLS files.
#[derive(Debug, thiserror::Error)]
pub enum XlsError {
    /// Error from the underlying CFB container reader.
    #[error("CFB error: {0}")]
    Cfb(#[from] crate::cfb::CfbError),

    /// A BIFF record has an unexpected length or content.
    #[error("invalid BIFF record: {0}")]
    InvalidRecord(String),

    /// The file uses a BIFF version that is not supported.
    #[error("unsupported BIFF version: {0}")]
    UnsupportedVersion(u16),

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

/// Convenience `Result` alias using `XlsError`.
pub type Result<T> = std::result::Result<T, XlsError>;
