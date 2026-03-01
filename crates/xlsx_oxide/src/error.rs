use thiserror::Error;

/// Errors specific to XLSX processing.
#[derive(Debug, Error)]
pub enum XlsxError {
    /// Error from the underlying OPC/XML layer.
    #[error(transparent)]
    Core(#[from] office_core::Error),
    /// The workbook part was not found.
    #[error("missing workbook")]
    MissingWorkbook,
    /// A cell reference string could not be parsed.
    #[error("invalid cell reference: {0}")]
    InvalidCellRef(String),
}

/// Convenience alias for `Result<T, XlsxError>`.
pub type Result<T> = std::result::Result<T, XlsxError>;
