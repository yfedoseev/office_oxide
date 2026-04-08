use thiserror::Error;

/// Errors specific to DOCX processing.
#[derive(Debug, Error)]
pub enum DocxError {
    /// Error from the underlying OPC/XML layer.
    #[error(transparent)]
    Core(#[from] crate::core::Error),
    /// The document body element was not found.
    #[error("missing document body")]
    MissingBody,
    /// A style reference could not be resolved.
    #[error("invalid style reference: {0}")]
    InvalidStyleRef(String),
}

/// Convenience alias for `Result<T, DocxError>`.
pub type Result<T> = std::result::Result<T, DocxError>;
