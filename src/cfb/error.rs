/// Errors that can occur when reading a CFB file.
#[derive(Debug, thiserror::Error)]
pub enum CfbError {
    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),

    #[error("invalid CFB header: {0}")]
    InvalidHeader(String),

    #[error("invalid FAT: {0}")]
    InvalidFat(String),

    #[error("invalid directory: {0}")]
    InvalidDirectory(String),

    #[error("stream not found: {0}")]
    StreamNotFound(String),

    #[error("corrupted stream: {0}")]
    CorruptedStream(String),
}

pub type Result<T> = std::result::Result<T, CfbError>;
