// SPDX-License-Identifier: MIT OR Apache-2.0
#![warn(missing_docs)]
//! # office_oxide
//!
//! The fastest Office document processing library for Rust.
//!
//! Reads, writes, and edits **DOCX, XLSX, PPTX, DOC, XLS, PPT** — all six
//! Microsoft Office formats — with a single unified API and zero C/C++
//! dependencies.
//!
//! ## Quick start
//!
//! ```rust,no_run
//! use office_oxide::Document;
//!
//! let doc = Document::open("report.docx")?;
//! println!("{}", doc.plain_text());
//! # Ok::<(), office_oxide::OfficeError>(())
//! ```
//!
//! ## Feature flags
//!
//! | Flag | What it enables |
//! |------|-----------------|
//! | `python` | PyO3 Python bindings |
//! | `wasm` | wasm-bindgen WASM bindings |
//! | `mmap` | Memory-mapped file I/O |
//! | `parallel` | Rayon-based parallel processing |

// Sub-modules (previously separate crates)
/// Library version (matches the Cargo package version).
pub const VERSION: &str = env!("CARGO_PKG_VERSION");

/// Compound Binary File (OLE2/CFB) container reader, used by legacy formats.
pub mod cfb;
/// Shared OOXML primitives: OPC, XML utilities, relationships, theme, units.
pub mod core;
/// Legacy Word Binary (.doc) document reader.
pub mod doc;
/// Word document (.docx) reader, writer, and editor.
pub mod docx;
/// Legacy PowerPoint Binary (.ppt) presentation reader.
pub mod ppt;
/// PowerPoint presentation (.pptx) reader, writer, and editor.
pub mod pptx;
/// Legacy Excel Binary (.xls) workbook reader.
pub mod xls;
/// Excel spreadsheet (.xlsx) reader, writer, and editor.
pub mod xlsx;

// Top-level modules
mod convert_doc;
mod convert_docx;
mod convert_ppt;
mod convert_pptx;
mod convert_xls;
mod convert_xlsx;
/// Document creation API: write new DOCX/XLSX/PPTX from scratch or from IR.
pub mod create;
/// Document editing API: modify existing DOCX/XLSX/PPTX files in-place.
pub mod edit;
/// Top-level error type wrapping all format-specific errors.
pub mod error;
/// `DocumentFormat` enum and format detection utilities.
pub mod format;
/// Format-agnostic intermediate representation (IR) of a document.
pub mod ir;
mod ir_from_markdown;
/// Format-agnostic renderers over [`DocumentIR`] — plain text, markdown
/// and HTML — plus the options that steer them.
pub mod ir_render;

#[cfg(not(target_family = "wasm"))]
pub mod ffi;

#[cfg(feature = "python")]
mod python;
#[cfg(feature = "wasm")]
mod wasm;

pub use core::OfficeDocument;
pub use error::{OfficeError, Result};
pub use format::DocumentFormat;
pub use ir::DocumentIR;

use std::io::{Read, Seek};
use std::path::Path;

use log::info;

/// Stack size for parsing threads (16 MB).
const PARSE_STACK_SIZE: usize = 16 * 1024 * 1024;

/// Whether the parse must run on a thread whose stack size we control.
///
/// This used to infer the answer from `RLIMIT_STACK`, and the inference was
/// unsound: that limit describes the process's *main* thread and says nothing
/// about the stack of whichever thread is actually running. The worst case was
/// `RLIM_INFINITY`, which took the "assume enough" branch and then ran inline
/// on an ordinary spawned thread with a 2 MiB stack — a 256-deep document
/// overflowed it and aborted the process, which is the uncatchable crash
/// `MAX_NESTING_DEPTH` exists to prevent. It reproduced on both Linux and
/// Windows CI while passing on a developer machine, purely because the two
/// had different `ulimit -s` values.
///
/// So we no longer guess: wherever threads exist, the parse gets
/// `PARSE_STACK_SIZE`. The cost is one spawn per top-level parse, which is
/// microseconds against a document parse, and in exchange the depth cap is
/// calibrated against a stack we own rather than the caller's.
fn needs_stack_thread() -> bool {
    // wasm32 has no threads; the host bounds the stack itself.
    !cfg!(target_arch = "wasm32")
}

/// Run a parsing closure on a stack whose size we control.
///
/// Every caller gets `PARSE_STACK_SIZE`, so a deeply nested document meets the
/// same headroom whether it arrives from a Rust binary, a Python binding or a
/// test harness. Only wasm32, which has no threads, runs inline.
fn with_parse_stack<F, T>(f: F) -> Result<T>
where
    F: FnOnce() -> Result<T> + Send + 'static,
    T: Send + 'static,
{
    if needs_stack_thread() {
        std::thread::Builder::new()
            .stack_size(PARSE_STACK_SIZE)
            .spawn(f)
            .map_err(|e| OfficeError::UnsupportedFormat(format!("thread spawn failed: {e}")))?
            .join()
            .unwrap_or_else(|payload| {
                // Surface the panic as itself. Reporting it as
                // `UnsupportedFormat` made every internal bug look like an
                // unreadable file, so real defects went unreported and the
                // fuzz target could not distinguish a crash from a clean
                // rejection.
                let msg = payload
                    .downcast_ref::<&str>()
                    .map(|s| (*s).to_string())
                    .or_else(|| payload.downcast_ref::<String>().cloned())
                    .unwrap_or_else(|| "unknown panic payload".to_string());
                Err(OfficeError::Panic(msg))
            })
    } else {
        f()
    }
}

/// Whether a reader's first bytes are the CFB (compound file) signature.
///
/// Leaves the reader rewound to the start.
fn is_cfb_container<R: Read + Seek>(reader: &mut R) -> Result<bool> {
    use std::io::SeekFrom;
    let mut magic = [0u8; 8];
    reader.seek(SeekFrom::Start(0)).map_err(core::Error::from)?;
    let n = read_up_to(reader, &mut magic)?;
    reader.seek(SeekFrom::Start(0)).map_err(core::Error::from)?;
    Ok(n == 8 && magic == crate::cfb::CFB_SIGNATURE)
}

/// Read up to `buf.len()` bytes, tolerating short reads.
fn read_up_to<R: Read>(reader: &mut R, buf: &mut [u8]) -> Result<usize> {
    let mut filled = 0;
    while filled < buf.len() {
        match reader.read(&mut buf[filled..]) {
            Ok(0) => break,
            Ok(n) => filled += n,
            Err(ref e) if e.kind() == std::io::ErrorKind::Interrupted => {},
            Err(e) => return Err(core::Error::from(e).into()),
        }
    }
    Ok(filled)
}

/// Dispatch a method call to the inner document type across all variants.
macro_rules! dispatch_inner {
    ($self:expr, $method:ident) => {
        match &$self.inner {
            DocumentInner::Docx(doc) => doc.$method(),
            DocumentInner::Xlsx(doc) => doc.$method(),
            DocumentInner::Pptx(doc) => doc.$method(),
            DocumentInner::Doc(doc) => doc.$method(),
            DocumentInner::Xls(doc) => doc.$method(),
            DocumentInner::Ppt(doc) => doc.$method(),
        }
    };
}

/// A unified document handle supporting DOCX, XLSX, PPTX, DOC, XLS, and PPT formats.
pub struct Document {
    inner: DocumentInner,
}

enum DocumentInner {
    Docx(Box<docx::DocxDocument>),
    Xlsx(Box<xlsx::XlsxDocument>),
    Pptx(Box<pptx::PptxDocument>),
    Doc(Box<doc::DocDocument>),
    Xls(Box<xls::XlsDocument>),
    Ppt(Box<ppt::PptDocument>),
}

impl Document {
    /// Open a document from a file path. Format is detected from the extension.
    #[must_use = "opening a document allocates — use the returned handle or drop it"]
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let path = path.as_ref().to_owned();
        with_parse_stack(move || Self::open_inner(&path))
    }

    fn open_inner(path: &Path) -> Result<Self> {
        let format = DocumentFormat::from_path(path);
        // The path is deliberately not logged: it routinely carries a
        // username and a document name, and this runs at info level on every
        // open. The caller already knows which path it passed.
        info!("Document::open: {format:?} format");
        let format = format.ok_or_else(|| {
            OfficeError::UnsupportedFormat(
                path.extension()
                    .and_then(|e| e.to_str())
                    .unwrap_or("(none)")
                    .to_string(),
            )
        })?;
        let format = sniff_format(path, format);

        match format {
            DocumentFormat::Docx => {
                let doc = docx::DocxDocument::open(path)?;
                Ok(Self {
                    inner: DocumentInner::Docx(Box::new(doc)),
                })
            },
            DocumentFormat::Xlsx => {
                let doc = xlsx::XlsxDocument::open(path)?;
                Ok(Self {
                    inner: DocumentInner::Xlsx(Box::new(doc)),
                })
            },
            DocumentFormat::Pptx => {
                let doc = pptx::PptxDocument::open(path)?;
                Ok(Self {
                    inner: DocumentInner::Pptx(Box::new(doc)),
                })
            },
            DocumentFormat::Doc => {
                let doc = doc::DocDocument::open(path)?;
                Ok(Self {
                    inner: DocumentInner::Doc(Box::new(doc)),
                })
            },
            DocumentFormat::Xls => {
                let doc = xls::XlsDocument::open(path)?;
                Ok(Self {
                    inner: DocumentInner::Xls(Box::new(doc)),
                })
            },
            DocumentFormat::Ppt => {
                let doc = ppt::PptDocument::open(path)?;
                Ok(Self {
                    inner: DocumentInner::Ppt(Box::new(doc)),
                })
            },
        }
    }

    /// Open a document from a file path using memory-mapped I/O.
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
        info!("Document::open_mmap: {format:?} format");
        match format {
            DocumentFormat::Docx => {
                let doc = docx::DocxDocument::open_mmap(path)?;
                Ok(Self {
                    inner: DocumentInner::Docx(Box::new(doc)),
                })
            },
            DocumentFormat::Xlsx => {
                let doc = xlsx::XlsxDocument::open_mmap(path)?;
                Ok(Self {
                    inner: DocumentInner::Xlsx(Box::new(doc)),
                })
            },
            DocumentFormat::Pptx => {
                let doc = pptx::PptxDocument::open_mmap(path)?;
                Ok(Self {
                    inner: DocumentInner::Pptx(Box::new(doc)),
                })
            },
            _ => Err(OfficeError::UnsupportedFormat(format!("{format:?}"))),
        }
    }

    /// Open a document from any `Read + Seek` source with an explicit format.
    #[must_use = "opening a document allocates — use the returned handle or drop it"]
    pub fn from_reader<R: Read + Seek + Send + 'static>(
        reader: R,
        format: DocumentFormat,
    ) -> Result<Self> {
        with_parse_stack(move || Self::from_reader_inner(reader, format))
    }

    fn from_reader_inner<R: Read + Seek>(mut reader: R, format: DocumentFormat) -> Result<Self> {
        // A password-protected OOXML file is not a zip at all: Office wraps
        // the encrypted package in a CFB container. Opening one as a zip
        // fails with an unhelpful archive error that says nothing about the
        // real reason, so name it here.
        if matches!(format, DocumentFormat::Docx | DocumentFormat::Xlsx | DocumentFormat::Pptx)
            && is_cfb_container(&mut reader)?
        {
            return Err(OfficeError::UnsupportedFormat(
                "the file is a password-protected (encrypted) OOXML package; \
                 decryption is not supported"
                    .into(),
            ));
        }
        match format {
            DocumentFormat::Docx => {
                let doc = docx::DocxDocument::from_reader(reader)?;
                Ok(Self {
                    inner: DocumentInner::Docx(Box::new(doc)),
                })
            },
            DocumentFormat::Xlsx => {
                let doc = xlsx::XlsxDocument::from_reader(reader)?;
                Ok(Self {
                    inner: DocumentInner::Xlsx(Box::new(doc)),
                })
            },
            DocumentFormat::Pptx => {
                let doc = pptx::PptxDocument::from_reader(reader)?;
                Ok(Self {
                    inner: DocumentInner::Pptx(Box::new(doc)),
                })
            },
            DocumentFormat::Doc => {
                let doc = doc::DocDocument::from_reader(reader)?;
                Ok(Self {
                    inner: DocumentInner::Doc(Box::new(doc)),
                })
            },
            DocumentFormat::Xls => {
                let doc = xls::XlsDocument::from_reader(reader)?;
                Ok(Self {
                    inner: DocumentInner::Xls(Box::new(doc)),
                })
            },
            DocumentFormat::Ppt => {
                let doc = ppt::PptDocument::from_reader(reader)?;
                Ok(Self {
                    inner: DocumentInner::Ppt(Box::new(doc)),
                })
            },
        }
    }

    /// Returns the document format.
    pub fn format(&self) -> DocumentFormat {
        match &self.inner {
            DocumentInner::Docx(_) => DocumentFormat::Docx,
            DocumentInner::Xlsx(_) => DocumentFormat::Xlsx,
            DocumentInner::Pptx(_) => DocumentFormat::Pptx,
            DocumentInner::Doc(_) => DocumentFormat::Doc,
            DocumentInner::Xls(_) => DocumentFormat::Xls,
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

    /// Convert to markdown with explicit rendering options.
    ///
    /// Unlike [`Self::to_markdown`], which uses the format-specific
    /// renderer, this goes through the IR so that options such as
    /// [`ir_render::ImageEmbed::Base64`] apply uniformly to every format.
    pub fn to_markdown_with(&self, options: ir_render::MarkdownOptions) -> String {
        self.to_ir().to_markdown_with(options)
    }

    /// Convert to an HTML fragment.
    pub fn to_html(&self) -> String {
        self.to_ir().to_html()
    }

    /// Convert to the format-agnostic Document IR.
    pub fn to_ir(&self) -> DocumentIR {
        match &self.inner {
            DocumentInner::Docx(doc) => convert_docx::docx_to_ir(doc),
            DocumentInner::Xlsx(doc) => convert_xlsx::xlsx_to_ir(doc),
            DocumentInner::Pptx(doc) => convert_pptx::pptx_to_ir(doc),
            DocumentInner::Doc(doc) => convert_doc::doc_to_ir(doc),
            DocumentInner::Xls(doc) => convert_xls::xls_to_ir(doc),
            DocumentInner::Ppt(doc) => convert_ppt::ppt_to_ir(doc),
        }
    }

    /// Return the inner DOCX document, if this document is a DOCX.
    pub fn as_docx(&self) -> Option<&docx::DocxDocument> {
        match &self.inner {
            DocumentInner::Docx(doc) => Some(doc),
            _ => None,
        }
    }

    /// Return the inner XLSX document, if this document is an XLSX.
    pub fn as_xlsx(&self) -> Option<&xlsx::XlsxDocument> {
        match &self.inner {
            DocumentInner::Xlsx(doc) => Some(doc),
            _ => None,
        }
    }

    /// Return the inner PPTX document, if this document is a PPTX.
    pub fn as_pptx(&self) -> Option<&pptx::PptxDocument> {
        match &self.inner {
            DocumentInner::Pptx(doc) => Some(doc),
            _ => None,
        }
    }

    /// Return the inner DOC document, if this document is a legacy DOC.
    pub fn as_doc(&self) -> Option<&doc::DocDocument> {
        match &self.inner {
            DocumentInner::Doc(doc) => Some(doc),
            _ => None,
        }
    }

    /// Return the inner XLS document, if this document is a legacy XLS.
    pub fn as_xls(&self) -> Option<&xls::XlsDocument> {
        match &self.inner {
            DocumentInner::Xls(doc) => Some(doc),
            _ => None,
        }
    }

    /// Return the inner PPT document, if this document is a legacy PPT.
    pub fn as_ppt(&self) -> Option<&ppt::PptDocument> {
        match &self.inner {
            DocumentInner::Ppt(doc) => Some(doc),
            _ => None,
        }
    }

    /// Save/convert the document to a file. Format is detected from the extension.
    ///
    /// Legacy formats (DOC, XLS, PPT) are automatically converted to OOXML
    /// (DOCX, XLSX, PPTX) via the intermediate representation.
    pub fn save_as(&self, path: impl AsRef<Path>) -> Result<()> {
        let path = path.as_ref();
        let target_format = DocumentFormat::from_path(path).ok_or_else(|| {
            OfficeError::UnsupportedFormat(
                path.extension()
                    .and_then(|e| e.to_str())
                    .unwrap_or("(none)")
                    .to_string(),
            )
        })?;
        let ir = self.to_ir();
        create::create_from_ir(&ir, target_format, path)?;
        Ok(())
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
fn sniff_format(path: &Path, ext_format: DocumentFormat) -> DocumentFormat {
    let Ok(mut file) = std::fs::File::open(path) else {
        return ext_format;
    };
    let mut magic = [0u8; 4];
    if std::io::Read::read(&mut file, &mut magic).unwrap_or(0) < 4 {
        return ext_format;
    }

    let is_zip = magic == [0x50, 0x4B, 0x03, 0x04];
    let is_cfb = magic == [0xD0, 0xCF, 0x11, 0xE0];

    match ext_format {
        DocumentFormat::Doc if is_zip => DocumentFormat::Docx,
        DocumentFormat::Xls if is_zip => DocumentFormat::Xlsx,
        DocumentFormat::Ppt if is_zip => DocumentFormat::Pptx,
        DocumentFormat::Docx if is_cfb => DocumentFormat::Doc,
        DocumentFormat::Xlsx if is_cfb => DocumentFormat::Xls,
        DocumentFormat::Pptx if is_cfb => DocumentFormat::Ppt,
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

/// Convert any supported document file to an HTML fragment.
pub fn to_html(path: impl AsRef<Path>) -> Result<String> {
    Ok(Document::open(path)?.to_html())
}
