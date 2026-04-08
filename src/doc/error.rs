/// Errors when reading legacy DOC files.
#[derive(Debug, thiserror::Error)]
pub enum DocError {
    #[error("CFB error: {0}")]
    Cfb(#[from] crate::cfb::CfbError),

    #[error("invalid FIB: {0}")]
    InvalidFib(String),

    #[error("invalid piece table: {0}")]
    InvalidPieceTable(String),

    #[error("missing stream: {0}")]
    MissingStream(String),

    #[error("corrupted data: {0}")]
    Corrupted(String),

    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),
}

pub type Result<T> = std::result::Result<T, DocError>;
