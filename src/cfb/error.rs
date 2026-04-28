/// Errors that can occur when reading a CFB file.
#[derive(Debug, thiserror::Error)]
pub enum CfbError {
    /// I/O error reading the underlying stream.
    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),

    /// The 512-byte CFB header is invalid or has a bad signature.
    #[error("invalid CFB header: {0}")]
    InvalidHeader(String),

    /// The File Allocation Table is corrupted or structurally invalid.
    #[error("invalid FAT: {0}")]
    InvalidFat(String),

    /// The directory tree is corrupted or unreadable.
    #[error("invalid directory: {0}")]
    InvalidDirectory(String),

    /// A requested stream name does not exist in the CFB container.
    #[error("stream not found: {0}")]
    StreamNotFound(String),

    /// A stream's data is corrupted or truncated.
    #[error("corrupted stream: {0}")]
    CorruptedStream(String),
}

/// Convenience `Result` alias using [`CfbError`].
pub type Result<T> = std::result::Result<T, CfbError>;
