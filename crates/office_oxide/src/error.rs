/// Unified error type for office_oxide.
#[derive(Debug, thiserror::Error)]
pub enum OfficeError {
    #[error(transparent)]
    Core(#[from] office_core::Error),

    #[cfg(feature = "docx")]
    #[error(transparent)]
    Docx(#[from] docx_oxide::DocxError),

    #[cfg(feature = "xlsx")]
    #[error(transparent)]
    Xlsx(#[from] xlsx_oxide::XlsxError),

    #[cfg(feature = "pptx")]
    #[error(transparent)]
    Pptx(#[from] pptx_oxide::PptxError),

    #[cfg(feature = "doc")]
    #[error(transparent)]
    Doc(#[from] doc_oxide::DocError),

    #[cfg(feature = "xls")]
    #[error(transparent)]
    Xls(#[from] xls_oxide::XlsError),

    #[cfg(feature = "ppt")]
    #[error(transparent)]
    Ppt(#[from] ppt_oxide::PptError),

    #[error("unsupported format: {0}")]
    UnsupportedFormat(String),
}

pub type Result<T> = std::result::Result<T, OfficeError>;
