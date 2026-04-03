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
pub use office_core::OfficeDocument;

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
#[cfg(feature = "doc")]
pub use doc_oxide;
#[cfg(feature = "xls")]
pub use xls_oxide;
#[cfg(feature = "ppt")]
pub use ppt_oxide;
pub use office_core;

use std::io::{Read, Seek};
use std::path::Path;

use log::info;

/// Dispatch a method call to the inner document type across all feature-gated variants.
macro_rules! dispatch_inner {
    ($self:expr, $method:ident) => {
        match &$self.inner {
            #[cfg(feature = "docx")]
            DocumentInner::Docx(doc) => doc.$method(),
            #[cfg(feature = "xlsx")]
            DocumentInner::Xlsx(doc) => doc.$method(),
            #[cfg(feature = "pptx")]
            DocumentInner::Pptx(doc) => doc.$method(),
            #[cfg(feature = "doc")]
            DocumentInner::Doc(doc) => doc.$method(),
            #[cfg(feature = "xls")]
            DocumentInner::Xls(doc) => doc.$method(),
            #[cfg(feature = "ppt")]
            DocumentInner::Ppt(doc) => doc.$method(),
        }
    };
}

/// A unified document handle supporting DOCX, XLSX, PPTX, DOC, XLS, and PPT formats.
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
    #[cfg(feature = "doc")]
    Doc(Box<doc_oxide::DocDocument>),
    #[cfg(feature = "xls")]
    Xls(Box<xls_oxide::XlsDocument>),
    #[cfg(feature = "ppt")]
    Ppt(Box<ppt_oxide::PptDocument>),
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
        // Sniff magic bytes to detect format mismatches (e.g., .doc that's actually OOXML, or vice versa).
        let format = sniff_format(path, format);

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
            #[cfg(feature = "doc")]
            DocumentFormat::Doc => {
                let doc = doc_oxide::DocDocument::open(path)?;
                Ok(Self { inner: DocumentInner::Doc(Box::new(doc)) })
            }
            #[cfg(feature = "xls")]
            DocumentFormat::Xls => {
                let doc = xls_oxide::XlsDocument::open(path)?;
                Ok(Self { inner: DocumentInner::Xls(Box::new(doc)) })
            }
            #[cfg(feature = "ppt")]
            DocumentFormat::Ppt => {
                let doc = ppt_oxide::PptDocument::open(path)?;
                Ok(Self { inner: DocumentInner::Ppt(Box::new(doc)) })
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
            #[cfg(feature = "doc")]
            DocumentFormat::Doc => {
                let doc = doc_oxide::DocDocument::from_reader(reader)?;
                Ok(Self { inner: DocumentInner::Doc(Box::new(doc)) })
            }
            #[cfg(feature = "xls")]
            DocumentFormat::Xls => {
                let doc = xls_oxide::XlsDocument::from_reader(reader)?;
                Ok(Self { inner: DocumentInner::Xls(Box::new(doc)) })
            }
            #[cfg(feature = "ppt")]
            DocumentFormat::Ppt => {
                let doc = ppt_oxide::PptDocument::from_reader(reader)?;
                Ok(Self { inner: DocumentInner::Ppt(Box::new(doc)) })
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
            #[cfg(feature = "doc")]
            DocumentInner::Doc(_) => DocumentFormat::Doc,
            #[cfg(feature = "xls")]
            DocumentInner::Xls(_) => DocumentFormat::Xls,
            #[cfg(feature = "ppt")]
            DocumentInner::Ppt(_) => DocumentFormat::Ppt,
        }
    }

    /// Extract plain text using the format-specific implementation.
    pub fn plain_text(&self) -> String {
        dispatch_inner!(self, plain_text)
    }

    /// Convert to markdown using the format-specific implementation.
    pub fn to_markdown(&self) -> String {
        dispatch_inner!(self, to_markdown)
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
            // Legacy formats: convert via plain text for now (no deep IR).
            #[cfg(feature = "doc")]
            DocumentInner::Doc(doc) => plain_text_to_ir(&doc.plain_text(), "Document", DocumentFormat::Doc),
            #[cfg(feature = "xls")]
            DocumentInner::Xls(doc) => plain_text_to_ir(&doc.plain_text(), "Spreadsheet", DocumentFormat::Xls),
            #[cfg(feature = "ppt")]
            DocumentInner::Ppt(doc) => plain_text_to_ir(&doc.plain_text(), "Presentation", DocumentFormat::Ppt),
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

    /// Access the underlying DOC document, if this is a legacy .doc file.
    #[cfg(feature = "doc")]
    pub fn as_doc(&self) -> Option<&doc_oxide::DocDocument> {
        match &self.inner {
            DocumentInner::Doc(doc) => Some(doc),
            _ => None,
        }
    }

    /// Access the underlying XLS document, if this is a legacy .xls file.
    #[cfg(feature = "xls")]
    pub fn as_xls(&self) -> Option<&xls_oxide::XlsDocument> {
        match &self.inner {
            DocumentInner::Xls(doc) => Some(doc),
            _ => None,
        }
    }

    /// Access the underlying PPT document, if this is a legacy .ppt file.
    #[cfg(feature = "ppt")]
    pub fn as_ppt(&self) -> Option<&ppt_oxide::PptDocument> {
        match &self.inner {
            DocumentInner::Ppt(doc) => Some(doc),
            _ => None,
        }
    }
}

/// Simple IR conversion from plain text for legacy formats.
fn plain_text_to_ir(text: &str, title: &str, format: DocumentFormat) -> DocumentIR {
    use ir::{Element, InlineContent, Metadata, Paragraph, Section, TextSpan};

    let elements: Vec<Element> = text
        .lines()
        .filter(|l| !l.trim().is_empty())
        .map(|line| {
            Element::Paragraph(Paragraph {
                content: vec![InlineContent::Text(TextSpan {
                    text: line.to_string(),
                    bold: false,
                    italic: false,
                    strikethrough: false,
                    hyperlink: None,
                })],
            })
        })
        .collect();

    DocumentIR {
        metadata: Metadata {
            format,
            title: Some(title.to_string()),
        },
        sections: vec![Section {
            title: Some(title.to_string()),
            elements,
        }],
    }
}

impl OfficeDocument for Document {
    fn plain_text(&self) -> String {
        self.plain_text()
    }

    fn to_markdown(&self) -> String {
        self.to_markdown()
    }
}

/// Sniff magic bytes to detect format mismatches.
///
/// Handles cases like:
/// - `.doc` file that's actually OOXML (ZIP) → route to DOCX parser
/// - `.docx` file that's actually OLE2 (CFB) → route to DOC parser
/// - Same for xls/xlsx and ppt/pptx
fn sniff_format(path: &Path, ext_format: DocumentFormat) -> DocumentFormat {
    let Ok(mut file) = std::fs::File::open(path) else {
        return ext_format;
    };
    let mut magic = [0u8; 4];
    if std::io::Read::read(&mut file, &mut magic).unwrap_or(0) < 4 {
        return ext_format;
    }

    let is_zip = magic == [0x50, 0x4B, 0x03, 0x04]; // PK\x03\x04
    let is_cfb = magic == [0xD0, 0xCF, 0x11, 0xE0]; // CFB signature (first 4 bytes)

    match ext_format {
        // Legacy extension but actually OOXML (ZIP)
        DocumentFormat::Doc if is_zip => DocumentFormat::Docx,
        DocumentFormat::Xls if is_zip => DocumentFormat::Xlsx,
        DocumentFormat::Ppt if is_zip => DocumentFormat::Pptx,
        // OOXML extension but actually legacy (CFB)
        DocumentFormat::Docx if is_cfb => DocumentFormat::Doc,
        DocumentFormat::Xlsx if is_cfb => DocumentFormat::Xls,
        DocumentFormat::Pptx if is_cfb => DocumentFormat::Ppt,
        // No mismatch
        _ => ext_format,
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
