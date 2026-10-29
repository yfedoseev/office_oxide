/// Unified error type for office_oxide.
#[derive(Debug, thiserror::Error)]
#[non_exhaustive]
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

    /// A parser panicked. Reported honestly rather than as "unsupported
    /// format": conflating a bug in this library with a file we simply do
    /// not read hid real defects from users and blinded the fuzz target,
    /// which could not tell a caught panic from a clean rejection.
    #[error("internal error: parsing panicked: {0}")]
    Panic(String),

    /// The document was read, but the result is known to be incomplete and
    /// the caller asked for an error in that case
    /// ([`OnIncomplete::Error`](crate::options::OnIncomplete::Error)). The
    /// message names the first losses; [`Document::damage`](crate::Document::damage)
    /// lists them all for a document opened without that option.
    #[error("the document is incomplete: {0}")]
    Incomplete(String),

    /// The file extension or format is not supported.
    #[error("unsupported format: {0}")]
    UnsupportedFormat(String),
}

/// Convenience `Result` alias using `OfficeError`.
pub type Result<T> = std::result::Result<T, OfficeError>;
