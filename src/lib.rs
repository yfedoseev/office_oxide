// Sub-modules (previously separate crates)
/// Library version (matches the Cargo package version).
pub const VERSION: &str = env!("CARGO_PKG_VERSION");

pub mod cfb;
pub mod core;
pub mod doc;
pub mod docx;
pub mod ppt;
pub mod pptx;
pub mod xls;
pub mod xlsx;

// Top-level modules
mod convert_doc;
mod convert_docx;
mod convert_ppt;
mod convert_pptx;
mod convert_xls;
mod convert_xlsx;
pub mod create;
pub mod edit;
pub mod error;
pub mod format;
pub mod ir;
mod ir_render;

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

/// Minimum stack size to run inline without spawning a thread (12 MB).
/// Below this, we spawn a thread with PARSE_STACK_SIZE.
const MIN_STACK_INLINE: usize = 12 * 1024 * 1024;

/// Check once whether the current environment has a large enough stack.
/// Caches the result for subsequent calls.
fn needs_stack_thread() -> bool {
    use std::sync::OnceLock;
    static NEEDS_THREAD: OnceLock<bool> = OnceLock::new();
    *NEEDS_THREAD.get_or_init(|| {
        // Check RLIMIT_STACK on Unix
        #[cfg(unix)]
        {
            let mut rlim = libc::rlimit {
                rlim_cur: 0,
                rlim_max: 0,
            };
            let ret = unsafe { libc::getrlimit(libc::RLIMIT_STACK, &mut rlim) };
            if ret == 0 && rlim.rlim_cur != libc::RLIM_INFINITY {
                return (rlim.rlim_cur as usize) < MIN_STACK_INLINE;
            }
            false // unlimited or error → assume enough
        }
        #[cfg(not(unix))]
        {
            false // On Windows/WASM, default is usually enough or handled differently
        }
    })
}

/// Run a parsing closure, spawning a thread with large stack only when needed.
/// On environments with sufficient stack (Rust programs, large-stack threads),
/// runs inline with zero overhead. On Python/constrained environments, spawns
/// a thread once and detects this via RLIMIT_STACK check (cached, O(1) after first call).
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
            .unwrap_or_else(|_| Err(OfficeError::UnsupportedFormat("parsing panicked".into())))
    } else {
        f()
    }
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
        info!("Document::open: {:?} format, path '{}'", format, path.display());
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
        info!("Document::open_mmap: {:?} format, path '{}'", format, path.display());
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

    fn from_reader_inner<R: Read + Seek>(reader: R, format: DocumentFormat) -> Result<Self> {
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

    pub fn as_docx(&self) -> Option<&docx::DocxDocument> {
        match &self.inner {
            DocumentInner::Docx(doc) => Some(doc),
            _ => None,
        }
    }

    pub fn as_xlsx(&self) -> Option<&xlsx::XlsxDocument> {
        match &self.inner {
            DocumentInner::Xlsx(doc) => Some(doc),
            _ => None,
        }
    }

    pub fn as_pptx(&self) -> Option<&pptx::PptxDocument> {
        match &self.inner {
            DocumentInner::Pptx(doc) => Some(doc),
            _ => None,
        }
    }

    pub fn as_doc(&self) -> Option<&doc::DocDocument> {
        match &self.inner {
            DocumentInner::Doc(doc) => Some(doc),
            _ => None,
        }
    }

    pub fn as_xls(&self) -> Option<&xls::XlsDocument> {
        match &self.inner {
            DocumentInner::Xls(doc) => Some(doc),
            _ => None,
        }
    }

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
