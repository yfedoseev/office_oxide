/// Unified error type for office_oxide.
#[derive(Debug, thiserror::Error)]
pub enum OfficeError {
    /// Error from the core OPC/XML layer.
    #[error(transparent)]
    Core(#[from] crate::core::Error),

    /// Error from the DOCX reader/writer.
    #[error(transparent)]
    Docx(#[from] crate::docx::DocxError),

    /// Error from the XLSX reader/writer.
    #[error(transparent)]
    Xlsx(#[from] crate::xlsx::XlsxError),

    /// Error from the PPTX reader/writer.
    #[error(transparent)]
    Pptx(#[from] crate::pptx::PptxError),

    /// Error from the legacy DOC reader.
    #[error(transparent)]
    Doc(#[from] crate::doc::DocError),

    /// Error from the legacy XLS reader.
    #[error(transparent)]
    Xls(#[from] crate::xls::XlsError),

    /// Error from the legacy PPT reader.
    #[error(transparent)]
    Ppt(#[from] crate::ppt::PptError),

    /// The file extension or format is not supported.
    #[error("unsupported format: {0}")]
    UnsupportedFormat(String),
}

/// Convenience `Result` alias using `OfficeError`.
pub type Result<T> = std::result::Result<T, OfficeError>;
