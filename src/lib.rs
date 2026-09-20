// SPDX-License-Identifier: MIT OR Apache-2.0
#![warn(missing_docs)]
//! # office_oxide
//!
//! The fastest Office document processing library for Rust.
//!
//! Reads all six Microsoft Office formats — **DOCX, XLSX, PPTX, DOC, XLS,
//! PPT** — through a single unified API, with zero C/C++ dependencies.
//! Writing and editing cover **DOCX, XLSX and PPTX**; the legacy binary
//! formats (DOC, XLS, PPT) are read-only, and convertible to OOXML via
//! [`Document::save_as`].
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

/// Number of parse threads that may be in flight at once.
///
/// Each parse thread reserves `PARSE_STACK_SIZE` (16 MB) of address space, so
/// an unbounded fan-out of simultaneous parses could exhaust the process's
/// thread or address-space limits. Scale with the machine but stay inside a
/// fixed ceiling, because the point is to have *some* bound.
fn max_parse_threads() -> usize {
    static MAX: std::sync::OnceLock<usize> = std::sync::OnceLock::new();
    *MAX.get_or_init(|| {
        std::thread::available_parallelism()
            .map_or(16, |n| n.get().saturating_mul(4))
            .clamp(8, 64)
    })
}

/// Parse threads currently in flight, and the signal that one has finished.
static PARSE_THREADS: std::sync::Mutex<usize> = std::sync::Mutex::new(0);
static PARSE_SLOT_FREED: std::sync::Condvar = std::sync::Condvar::new();

/// RAII reservation for one in-flight parse thread.
struct ParseSlot;

impl ParseSlot {
    /// Reserve a slot, waiting for one to free up when the cap is reached.
    ///
    /// Waiting (rather than failing) is deliberate: a parse is short-lived and
    /// always releases its slot, so a host that fans out many simultaneous
    /// parses is throttled to the cap instead of being handed a spurious
    /// "too many concurrent parses" error it cannot act on. Callers already
    /// block for the duration of their own parse, so the only visible effect
    /// is that the (cap + 1)-th concurrent parse starts slightly later.
    fn acquire() -> Self {
        let cap = max_parse_threads();
        let mut in_flight = PARSE_THREADS.lock().unwrap_or_else(|e| e.into_inner());
        while *in_flight >= cap {
            // A timeout keeps a lost notification from parking a caller
            // forever; the predicate is rechecked on every wake.
            let (guard, _) = PARSE_SLOT_FREED
                .wait_timeout(in_flight, std::time::Duration::from_millis(50))
                .unwrap_or_else(|e| e.into_inner());
            in_flight = guard;
        }
        *in_flight += 1;
        Self
    }
}

impl Drop for ParseSlot {
    fn drop(&mut self) {
        let mut in_flight = PARSE_THREADS.lock().unwrap_or_else(|e| e.into_inner());
        *in_flight = in_flight.saturating_sub(1);
        drop(in_flight);
        PARSE_SLOT_FREED.notify_one();
    }
}

/// Run a parsing closure on a stack whose size we control.
///
/// Every caller gets `PARSE_STACK_SIZE`, so a deeply nested document meets the
/// same headroom whether it arrives from a Rust binary, a Python binding or a
/// test harness. Only wasm32, which has no threads, runs inline.
///
/// At most [`max_parse_threads`] parses are in flight at once; further callers
/// wait for a slot rather than spawning an unbounded number of 16 MB-stack
/// threads. The reservation is released when this function returns, whether
/// the parse succeeded, failed or panicked.
fn with_parse_stack<F, T>(f: F) -> Result<T>
where
    F: FnOnce() -> Result<T> + Send + 'static,
    T: Send + 'static,
{
    if needs_stack_thread() {
        // Held until the join below completes, so the count tracks threads
        // that actually exist.
        let _slot = ParseSlot::acquire();
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
/// Leaves the reader rewound to the start. Thin wrapper around
/// `cfb::is_cfb_container`, which is also called directly by each OOXML
/// format's own `from_reader` — kept here too so this
/// module's existing `Result` (`OfficeError`) call sites don't change.
fn is_cfb_container<R: Read + Seek>(reader: &mut R) -> Result<bool> {
    Ok(crate::cfb::is_cfb_container(reader).map_err(core::Error::from)?)
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
        let ext_format = format.ok_or_else(|| {
            OfficeError::UnsupportedFormat(
                path.extension()
                    .and_then(|e| e.to_str())
                    .unwrap_or("(none)")
                    .to_string(),
            )
        })?;
        let format = sniff_format(path, ext_format);
        // sniff_format's CFB branch exists so a legacy .doc/.xls/.ppt file
        // saved with the wrong extension still opens — but an encrypted
        // OOXML package is ALSO a CFB container, so the same branch was
        // silently misrouting it to the legacy parser too, which then
        // failed looking for a stream that was never there ("missing
        // stream: WordDocument stream not found") instead of surfacing
        // the real cause. Distinguish the two by the MS-OFFCRYPTO streams
        // an encrypted package actually carries.
        if matches!(format, DocumentFormat::Doc | DocumentFormat::Xls | DocumentFormat::Ppt)
            && matches!(
                ext_format,
                DocumentFormat::Docx | DocumentFormat::Xlsx | DocumentFormat::Pptx
            )
            && is_encrypted_ooxml_cfb(path)
        {
            return Err(OfficeError::UnsupportedFormat(
                "the file is a password-protected (encrypted) OOXML package; \
                 decryption is not supported"
                    .into(),
            ));
        }

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

    /// Convert to an HTML fragment with explicit rendering options.
    ///
    /// With [`ir_render::ImageEmbed::Base64`] the fragment carries its
    /// images inline as `data:` URIs, so it renders standalone — the
    /// counterpart of [`Self::to_markdown_with`].
    pub fn to_html_with(&self, options: ir_render::HtmlOptions) -> String {
        self.to_ir().to_html_with(options)
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
/// `true` when the CFB container at `path` carries the MS-OFFCRYPTO
/// encryption streams (`EncryptionInfo` + `EncryptedPackage`) that mark it
/// as a password-protected OOXML package, rather than a genuine legacy
/// binary document that happens to reuse the same container format.
fn is_encrypted_ooxml_cfb(path: &Path) -> bool {
    let Ok(file) = std::fs::File::open(path) else {
        return false;
    };
    let Ok(reader) = crate::cfb::CfbReader::new(file) else {
        return false;
    };
    reader.has_stream("EncryptionInfo") && reader.has_stream("EncryptedPackage")
}

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

#[cfg(test)]
mod tests {
    use super::*;

    /// Build a minimal CFB container with two named, empty-ish streams at
    /// the root — enough for `has_stream` (a flat scan over directory
    /// entries; sibling-tree links don't matter for it) without needing a
    /// full CFB writer. Mirrors `cfb::reader::tests::build_minimal_cfb`,
    /// extended to two streams instead of one.
    fn build_two_stream_cfb(name1: &str, name2: &str) -> Vec<u8> {
        const NO_ENTRY: u32 = 0xFFFF_FFFF;
        const END_OF_CHAIN: u32 = 0xFFFF_FFFE;
        const FAT_SECT: u32 = 0xFFFF_FFFD;
        const FREE_SECT: u32 = 0xFFFF_FFFF;
        let sector_size = 512usize;
        let mut file = vec![0u8; 512 + 4 * sector_size];

        file[0..8].copy_from_slice(&[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
        file[0x18..0x1A].copy_from_slice(&0x003Eu16.to_le_bytes());
        file[0x1A..0x1C].copy_from_slice(&3u16.to_le_bytes());
        file[0x1C..0x1E].copy_from_slice(&0xFFFEu16.to_le_bytes());
        file[0x1E..0x20].copy_from_slice(&9u16.to_le_bytes());
        file[0x20..0x22].copy_from_slice(&6u16.to_le_bytes());
        file[0x2C..0x30].copy_from_slice(&1u32.to_le_bytes());
        file[0x30..0x34].copy_from_slice(&0u32.to_le_bytes());
        file[0x38..0x3C].copy_from_slice(&4096u32.to_le_bytes());
        file[0x3C..0x40].copy_from_slice(&END_OF_CHAIN.to_le_bytes());
        file[0x40..0x44].copy_from_slice(&0u32.to_le_bytes());
        file[0x44..0x48].copy_from_slice(&END_OF_CHAIN.to_le_bytes());
        file[0x48..0x4C].copy_from_slice(&0u32.to_le_bytes());
        file[0x4C..0x50].copy_from_slice(&1u32.to_le_bytes());
        for i in 1..109 {
            let off = 0x4C + i * 4;
            file[off..off + 4].copy_from_slice(&FREE_SECT.to_le_bytes());
        }

        fn write_dir_entry(
            buf: &mut [u8],
            name: &str,
            entry_type: u8,
            child: u32,
            start_sector: u32,
            stream_size: u32,
        ) {
            const NO_ENTRY: u32 = 0xFFFF_FFFF;
            let utf16: Vec<u16> = name.encode_utf16().collect();
            for (i, &ch) in utf16.iter().enumerate() {
                let bytes = ch.to_le_bytes();
                buf[i * 2] = bytes[0];
                buf[i * 2 + 1] = bytes[1];
            }
            let name_size = ((utf16.len() + 1) * 2) as u16;
            buf[0x40..0x42].copy_from_slice(&name_size.to_le_bytes());
            buf[0x42] = entry_type;
            buf[0x43] = 1;
            buf[0x44..0x48].copy_from_slice(&NO_ENTRY.to_le_bytes());
            buf[0x48..0x4C].copy_from_slice(&NO_ENTRY.to_le_bytes());
            buf[0x4C..0x50].copy_from_slice(&child.to_le_bytes());
            buf[0x74..0x78].copy_from_slice(&start_sector.to_le_bytes());
            buf[0x78..0x7C].copy_from_slice(&stream_size.to_le_bytes());
        }

        let dir_offset = 512;
        write_dir_entry(
            &mut file[dir_offset..dir_offset + 128],
            "Root Entry",
            5,
            1,
            END_OF_CHAIN,
            0,
        );
        write_dir_entry(&mut file[dir_offset + 128..dir_offset + 256], name1, 2, NO_ENTRY, 2, 4);
        write_dir_entry(&mut file[dir_offset + 256..dir_offset + 384], name2, 2, NO_ENTRY, 3, 4);
        // Sibling-link entry 1 ("name1") to entry 2 ("name2") so both are
        // reachable from Root's tree — find_entry walks the tree
        // via child/sibling pointers, not a flat directory-array scan.
        file[dir_offset + 128 + 0x48..dir_offset + 128 + 0x4C].copy_from_slice(&2u32.to_le_bytes());
        file[dir_offset + 384 + 0x42] = 0; // empty 4th entry

        let fat_offset = 512 + sector_size;
        let write_fat = |file: &mut [u8], index: usize, value: u32| {
            let off = fat_offset + index * 4;
            file[off..off + 4].copy_from_slice(&value.to_le_bytes());
        };
        write_fat(&mut file, 0, END_OF_CHAIN); // dir
        write_fat(&mut file, 1, FAT_SECT); // FAT
        write_fat(&mut file, 2, END_OF_CHAIN); // stream1
        write_fat(&mut file, 3, END_OF_CHAIN); // stream2
        for i in 4..128 {
            write_fat(&mut file, i, FREE_SECT);
        }

        file[512 + 2 * sector_size..512 + 2 * sector_size + 4].copy_from_slice(b"data");
        file[512 + 3 * sector_size..512 + 3 * sector_size + 4].copy_from_slice(b"data");

        file
    }

    /// An encrypted OOXML file has the same CFB magic bytes
    /// as a genuine legacy .doc/.xls/.ppt, so `sniff_format`'s
    /// "wrong-extension" branch silently remapped it to the legacy
    /// parser, which then failed looking for a stream that was never
    /// there ("missing stream: WordDocument stream not found") instead
    /// of naming the real cause.
    #[test]
    fn test_open_encrypted_ooxml_gives_a_friendly_error_not_a_legacy_parser_failure() {
        let data = build_two_stream_cfb("EncryptionInfo", "EncryptedPackage");
        let dir = std::env::temp_dir();
        let path = dir.join(format!("office_oxide_test_encrypted_{}.docx", std::process::id()));
        std::fs::write(&path, &data).unwrap();
        let result = Document::open(&path);
        std::fs::remove_file(&path).ok();
        let err = result.err().expect("expected an Err");
        let msg = err.to_string();
        assert!(
            msg.contains("password-protected"),
            "expected a friendly password-protected message, got: {msg}"
        );
    }

    /// An encrypted .pptx used to succeed with Ok and a
    /// silently EMPTY (zero-slide) presentation rather than erroring,
    /// because sniff_format remapped it to the legacy PPT parser, which
    /// (unlike the DOC/XLS legacy readers) doesn't detect encryption
    /// itself and just parsed whatever little structure it could find.
    /// The encrypted-OOXML check runs before any legacy-parser dispatch at all, so
    /// this is the same code path with a .pptx extension.
    #[test]
    fn test_open_encrypted_pptx_gives_a_friendly_error_not_an_empty_presentation() {
        let data = build_two_stream_cfb("EncryptionInfo", "EncryptedPackage");
        let dir = std::env::temp_dir();
        let path = dir.join(format!("office_oxide_test_encrypted_{}.pptx", std::process::id()));
        std::fs::write(&path, &data).unwrap();
        let result = Document::open(&path);
        std::fs::remove_file(&path).ok();
        let err = result
            .err()
            .expect("expected an Err, not a silently empty presentation");
        let msg = err.to_string();
        assert!(
            msg.contains("password-protected"),
            "expected a friendly password-protected message, got: {msg}"
        );
    }

    /// The same CFB-magic-on-a-.docx-path shape, but WITHOUT the
    /// MS-OFFCRYPTO streams, must still fall through to the legacy
    /// parser (a genuinely misnamed legacy file) — the new check must
    /// not misfire on the case `sniff_format` already handled
    /// correctly.
    #[test]
    fn test_open_cfb_without_offcrypto_streams_still_falls_through_to_legacy_parser() {
        let data = build_two_stream_cfb("SomeStream", "OtherStream");
        let dir = std::env::temp_dir();
        let path = dir.join(format!("office_oxide_test_not_encrypted_{}.docx", std::process::id()));
        std::fs::write(&path, &data).unwrap();
        let result = Document::open(&path);
        std::fs::remove_file(&path).ok();
        let err = result.err().expect("expected an Err");
        let msg = err.to_string();
        assert!(
            !msg.contains("password-protected"),
            "a non-OFFCRYPTO CFB must not be misreported as encrypted: {msg}"
        );
    }

    /// `with_parse_stack` used to spawn one 16 MB-stack thread per parse with
    /// no bound at all, so a host that fanned out many simultaneous parses
    /// could drive the process into its thread-creation limit. More callers
    /// than the cap must still all complete, and never more than the cap may
    /// be in flight at once.
    #[test]
    fn test_with_parse_stack_bounds_concurrent_parses() {
        use std::sync::Arc;
        use std::sync::atomic::{AtomicUsize, Ordering};

        let cap = max_parse_threads();
        assert!((8..=64).contains(&cap), "cap should be sane, got {cap}");

        let peak = Arc::new(AtomicUsize::new(0));
        let done = Arc::new(AtomicUsize::new(0));
        let callers: Vec<_> = (0..cap + 16)
            .map(|_| {
                let peak = Arc::clone(&peak);
                let done = Arc::clone(&done);
                std::thread::spawn(move || {
                    let r: Result<()> = with_parse_stack(move || {
                        let live = *PARSE_THREADS.lock().unwrap_or_else(|e| e.into_inner());
                        peak.fetch_max(live, Ordering::SeqCst);
                        std::thread::sleep(std::time::Duration::from_millis(5));
                        Ok(())
                    });
                    r.expect("parse closure should succeed");
                    done.fetch_add(1, Ordering::SeqCst);
                })
            })
            .collect();
        for c in callers {
            c.join().expect("caller thread should not panic");
        }

        assert_eq!(done.load(Ordering::SeqCst), cap + 16, "every parse must complete");
        let peak = peak.load(Ordering::SeqCst);
        assert!(peak >= 1, "the in-flight counter should have been observed");
        assert!(peak <= cap, "in-flight parses {peak} exceeded the cap {cap}");
    }

    /// A panic inside the parse closure must still release its slot, or the
    /// cap would leak and eventually deadlock every later parse.
    #[test]
    fn test_parse_slot_released_after_panic() {
        // `PARSE_THREADS` is a process-wide gauge shared with every other
        // test in this binary, and it never exceeds `max_parse_threads()`,
        // so sampling it before and after is noise, not evidence: the old
        // `after < before + 20` assertion failed whenever the rest of the
        // suite happened to hold twenty slots at once. The deterministic
        // property is what a leak would *do*: once leaked slots reach the
        // cap, the next acquire waits forever. So make more calls than the
        // cap, on a helper thread, and fail if it does not come back.
        let calls = max_parse_threads() + 8;
        let worker = std::thread::spawn(move || {
            for _ in 0..calls {
                let r: Result<()> = with_parse_stack(|| panic!("boom"));
                assert!(matches!(r, Err(OfficeError::Panic(_))), "panic should surface as itself");
            }
        });
        let deadline = std::time::Instant::now() + std::time::Duration::from_secs(120);
        while !worker.is_finished() {
            assert!(
                std::time::Instant::now() < deadline,
                "a panicking parse leaked its slot: {calls} calls did not complete \
                 (the cap is {})",
                max_parse_threads()
            );
            std::thread::sleep(std::time::Duration::from_millis(10));
        }
        worker.join().expect("worker panicked");
    }
}
