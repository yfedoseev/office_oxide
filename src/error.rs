/// Unified error type for office_oxide.
#[derive(Debug, thiserror::Error)]
pub enum OfficeError {
    #[error(transparent)]
    Core(#[from] crate::core::Error),

    #[error(transparent)]
    Docx(#[from] crate::docx::DocxError),

    #[error(transparent)]
    Xlsx(#[from] crate::xlsx::XlsxError),

    #[error(transparent)]
    Pptx(#[from] crate::pptx::PptxError),

    #[error(transparent)]
    Doc(#[from] crate::doc::DocError),

    #[error(transparent)]
    Xls(#[from] crate::xls::XlsError),

    #[error(transparent)]
    Ppt(#[from] crate::ppt::PptError),

    #[error("unsupported format: {0}")]
    UnsupportedFormat(String),
}

pub type Result<T> = std::result::Result<T, OfficeError>;
