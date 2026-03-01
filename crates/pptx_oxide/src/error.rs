use thiserror::Error;

/// Errors specific to PPTX processing.
#[derive(Debug, Error)]
pub enum PptxError {
    /// Error from the underlying OPC/XML layer.
    #[error(transparent)]
    Core(#[from] office_core::Error),
    /// The presentation part was not found.
    #[error("missing presentation")]
    MissingPresentation,
}

/// Convenience alias for `Result<T, PptxError>`.
pub type Result<T> = std::result::Result<T, PptxError>;
