pub mod error;
pub mod format;
pub mod ir;
mod ir_render;

#[cfg(feature = "docx")]
mod convert_docx;
#[cfg(feature = "xlsx")]
mod convert_xlsx;
#[cfg(feature = "pptx")]
mod convert_pptx;

pub use error::{OfficeError, Result};
pub use format::DocumentFormat;
pub use ir::DocumentIR;

pub mod create;
pub mod edit;

#[cfg(feature = "python")]
mod python;
#[cfg(feature = "wasm")]
mod wasm;

// Re-export format crate types
#[cfg(feature = "docx")]
pub use docx_oxide;
#[cfg(feature = "xlsx")]
pub use xlsx_oxide;
#[cfg(feature = "pptx")]
pub use pptx_oxide;
pub use office_core;

use std::io::{Read, Seek};
use std::path::Path;

use log::info;

/// A unified document handle supporting DOCX, XLSX, and PPTX formats.
pub struct Document {
    inner: DocumentInner,
}

enum DocumentInner {
    #[cfg(feature = "docx")]
    Docx(Box<docx_oxide::DocxDocument>),
    #[cfg(feature = "xlsx")]
    Xlsx(Box<xlsx_oxide::XlsxDocument>),
    #[cfg(feature = "pptx")]
    Pptx(Box<pptx_oxide::PptxDocument>),
}

impl Document {
    /// Open a document from a file path. Format is detected from the extension.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let path = path.as_ref();
        let format = DocumentFormat::from_path(path);
        info!("Document::open: {:?} format, path '{}'", format, path.display());
        let format = format.ok_or_else(|| {
            OfficeError::UnsupportedFormat(
                path.extension()
                    .and_then(|e| e.to_str())
                    .unwrap_or("(none)")
                    .to_string(),
            )
        })?;
        match format {
            #[cfg(feature = "docx")]
            DocumentFormat::Docx => {
                let doc = docx_oxide::DocxDocument::open(path)?;
                Ok(Self { inner: DocumentInner::Docx(Box::new(doc)) })
            }
            #[cfg(feature = "xlsx")]
            DocumentFormat::Xlsx => {
                let doc = xlsx_oxide::XlsxDocument::open(path)?;
                Ok(Self { inner: DocumentInner::Xlsx(Box::new(doc)) })
            }
            #[cfg(feature = "pptx")]
            DocumentFormat::Pptx => {
                let doc = pptx_oxide::PptxDocument::open(path)?;
                Ok(Self { inner: DocumentInner::Pptx(Box::new(doc)) })
            }
            #[allow(unreachable_patterns)]
            _ => Err(OfficeError::UnsupportedFormat(format!("{format:?}"))),
        }
    }

    /// Open a document from a file path using memory-mapped I/O.
    /// Format is detected from the extension.
    #[cfg(feature = "mmap")]
    pub fn open_mmap(path: impl AsRef<Path>) -> Result<Self> {
        let path = path.as_ref();
        let format = DocumentFormat::from_path(path).ok_or_else(|| {
            OfficeError::UnsupportedFormat(
                path.extension()
                    .and_then(|e| e.to_str())
                    .unwrap_or("(none)")
                    .to_string(),
            )
        })?;
        info!("Document::open_mmap: {:?} format, path '{}'", format, path.display());
        match format {
            #[cfg(feature = "docx")]
            DocumentFormat::Docx => {
                let doc = docx_oxide::DocxDocument::open_mmap(path)?;
                Ok(Self { inner: DocumentInner::Docx(Box::new(doc)) })
            }
            #[cfg(feature = "xlsx")]
            DocumentFormat::Xlsx => {
                let doc = xlsx_oxide::XlsxDocument::open_mmap(path)?;
                Ok(Self { inner: DocumentInner::Xlsx(Box::new(doc)) })
            }
            #[cfg(feature = "pptx")]
            DocumentFormat::Pptx => {
                let doc = pptx_oxide::PptxDocument::open_mmap(path)?;
                Ok(Self { inner: DocumentInner::Pptx(Box::new(doc)) })
            }
            #[allow(unreachable_patterns)]
            _ => Err(OfficeError::UnsupportedFormat(format!("{format:?}"))),
        }
    }

    /// Open a document from any `Read + Seek` source with an explicit format.
    pub fn from_reader<R: Read + Seek>(reader: R, format: DocumentFormat) -> Result<Self> {
        match format {
            #[cfg(feature = "docx")]
            DocumentFormat::Docx => {
                let doc = docx_oxide::DocxDocument::from_reader(reader)?;
                Ok(Self { inner: DocumentInner::Docx(Box::new(doc)) })
            }
            #[cfg(feature = "xlsx")]
            DocumentFormat::Xlsx => {
                let doc = xlsx_oxide::XlsxDocument::from_reader(reader)?;
                Ok(Self { inner: DocumentInner::Xlsx(Box::new(doc)) })
            }
            #[cfg(feature = "pptx")]
            DocumentFormat::Pptx => {
                let doc = pptx_oxide::PptxDocument::from_reader(reader)?;
                Ok(Self { inner: DocumentInner::Pptx(Box::new(doc)) })
            }
            #[allow(unreachable_patterns)]
            _ => Err(OfficeError::UnsupportedFormat(format!("{format:?}"))),
        }
    }

    /// Returns the document format.
    pub fn format(&self) -> DocumentFormat {
        match &self.inner {
            #[cfg(feature = "docx")]
            DocumentInner::Docx(_) => DocumentFormat::Docx,
            #[cfg(feature = "xlsx")]
            DocumentInner::Xlsx(_) => DocumentFormat::Xlsx,
            #[cfg(feature = "pptx")]
            DocumentInner::Pptx(_) => DocumentFormat::Pptx,
        }
    }

    /// Extract plain text using the format-specific implementation.
    pub fn plain_text(&self) -> String {
        match &self.inner {
            #[cfg(feature = "docx")]
            DocumentInner::Docx(doc) => doc.plain_text(),
            #[cfg(feature = "xlsx")]
            DocumentInner::Xlsx(doc) => doc.plain_text(),
            #[cfg(feature = "pptx")]
            DocumentInner::Pptx(doc) => doc.plain_text(),
        }
    }

    /// Convert to markdown using the format-specific implementation.
    pub fn to_markdown(&self) -> String {
        match &self.inner {
            #[cfg(feature = "docx")]
            DocumentInner::Docx(doc) => doc.to_markdown(),
            #[cfg(feature = "xlsx")]
            DocumentInner::Xlsx(doc) => doc.to_markdown(),
            #[cfg(feature = "pptx")]
            DocumentInner::Pptx(doc) => doc.to_markdown(),
        }
    }

    /// Convert to the format-agnostic Document IR.
    pub fn to_ir(&self) -> DocumentIR {
        match &self.inner {
            #[cfg(feature = "docx")]
            DocumentInner::Docx(doc) => convert_docx::docx_to_ir(doc),
            #[cfg(feature = "xlsx")]
            DocumentInner::Xlsx(doc) => convert_xlsx::xlsx_to_ir(doc),
            #[cfg(feature = "pptx")]
            DocumentInner::Pptx(doc) => convert_pptx::pptx_to_ir(doc),
        }
    }

    /// Access the underlying DOCX document, if this is a DOCX file.
    #[cfg(feature = "docx")]
    pub fn as_docx(&self) -> Option<&docx_oxide::DocxDocument> {
        match &self.inner {
            DocumentInner::Docx(doc) => Some(doc),
            _ => None,
        }
    }

    /// Access the underlying XLSX document, if this is an XLSX file.
    #[cfg(feature = "xlsx")]
    pub fn as_xlsx(&self) -> Option<&xlsx_oxide::XlsxDocument> {
        match &self.inner {
            DocumentInner::Xlsx(doc) => Some(doc),
            _ => None,
        }
    }

    /// Access the underlying PPTX document, if this is a PPTX file.
    #[cfg(feature = "pptx")]
    pub fn as_pptx(&self) -> Option<&pptx_oxide::PptxDocument> {
        match &self.inner {
            DocumentInner::Pptx(doc) => Some(doc),
            _ => None,
        }
    }
}

/// Extract plain text from any supported document file.
pub fn extract_text(path: impl AsRef<Path>) -> Result<String> {
    Ok(Document::open(path)?.plain_text())
}

/// Convert any supported document file to markdown.
pub fn to_markdown(path: impl AsRef<Path>) -> Result<String> {
    Ok(Document::open(path)?.to_markdown())
}
