//! C Foreign Function Interface (FFI) for office_oxide.
//!
//! Provides `#[no_mangle] pub extern "C"` functions that Go (CGo), Node.js (N-API),
//! and C# (P/Invoke) bindings can link against. The compiled `liboffice_oxide.so` /
//! `.dylib` / `.dll` / `.a` exports these symbols.
//!
//! # Error Convention
//! Most functions accept an `error_code: *mut i32` out-parameter:
//! - 0 = success
//! - 1 = invalid argument / path
//! - 2 = IO error
//! - 3 = parse error
//! - 4 = extraction failed
//! - 5 = internal error (including a panic caught at the boundary)
//! - 6 = unsupported format / feature
//!
//! The code is chosen from the error's *variant* (see `classify_error`),
//! never from its message text.
//!
//! # Panic containment
//! Every exported function runs its body under `catch_unwind`. A panic in
//! the library — a bug, never an expected outcome — is reported as
//! `OFFICE_ERR_INTERNAL` (with a NULL / -1 / status return) instead of
//! unwinding into the host, which aborts the whole process (Node, the Go
//! runtime, .NET, CPython). A handle whose call panicked stays memory-safe
//! to use and free, but its contents may reflect a half-applied edit.
//! Builds with `panic = "abort"` (the `release-small` profile) cannot
//! contain panics at all.
//!
//! # Memory Convention
//! - Strings returned as `*mut c_char` are heap-allocated and must be freed with
//!   `office_oxide_free_string`. A NUL byte inside the text (which a C string
//!   cannot carry) is replaced by U+FFFD.
//! - Byte buffers returned as `*mut u8` (with an `out_len`) must be freed with
//!   `office_oxide_free_bytes(ptr, len)`.
//! - Opaque handles (`*mut OfficeDocumentHandle`, `*mut OfficeEditableHandle`)
//!   must be freed with their corresponding `*_free` function.
//!
//! # Thread-safety Convention
//! **A handle must not be used from more than one thread at a time.** The
//! contract is the same as `sqlite3*` in serialized-off mode or `FILE*`:
//!
//! - Each handle is owned by the caller and carries no internal lock. These
//!   functions reconstruct `&`/`&mut` references to the handle's contents
//!   across the FFI boundary, so two concurrent calls that touch the same
//!   handle — e.g. two threads calling
//!   `office_oxide_editable_replace_text` on one `*mut OfficeEditableHandle`,
//!   or one thread calling a `*_free` while another still uses the handle —
//!   are a data race and undefined behaviour. Rust cannot detect or prevent
//!   this across the boundary; it is the caller's responsibility.
//! - Callers that share a handle between threads (Go goroutines on different
//!   OS threads, .NET thread-pool work items, raw pthreads) must serialize
//!   every call on that handle with their own mutex. Python's GIL happens to
//!   provide that serialization for the Python binding; no other binding gets
//!   it for free.
//! - *Distinct* handles are independent: different threads may each use their
//!   own handle concurrently without synchronization, and the library's own
//!   Rust-side state is otherwise thread-safe.
#![allow(missing_docs)]
#![allow(clippy::missing_safety_doc)]
#![allow(clippy::not_unsafe_ptr_arg_deref)]
#![allow(clippy::too_many_arguments)]

use std::ffi::{CStr, CString};
use std::os::raw::c_char;
use std::panic::{AssertUnwindSafe, catch_unwind};
use std::path::PathBuf;
use std::ptr;
use std::slice;

use crate::Document;
use crate::edit::EditableDocument;
use crate::format::DocumentFormat;

// ─── Error codes ────────────────────────────────────────────────────────────

pub const OFFICE_OK: i32 = 0;
pub const OFFICE_ERR_INVALID_ARG: i32 = 1;
pub const OFFICE_ERR_IO: i32 = 2;
pub const OFFICE_ERR_PARSE: i32 = 3;
pub const OFFICE_ERR_EXTRACTION: i32 = 4;
pub const OFFICE_ERR_INTERNAL: i32 = 5;
pub const OFFICE_ERR_UNSUPPORTED: i32 = 6;

fn set_err(ptr: *mut i32, code: i32) {
    if !ptr.is_null() {
        unsafe { *ptr = code };
    }
}

/// Run an exported function's body, containing any panic.
///
/// A panic unwinding out of an `extern "C"` function aborts the host
/// process. Only the parse entry points ran on a guarded thread; every
/// render, edit, save and writer call was unprotected. On a panic this
/// sets `OFFICE_ERR_INTERNAL` and returns `on_panic`.
fn guard<T>(error_code: *mut i32, on_panic: T, body: impl FnOnce() -> T) -> T {
    match catch_unwind(AssertUnwindSafe(body)) {
        Ok(v) => v,
        Err(_) => {
            set_err(error_code, OFFICE_ERR_INTERNAL);
            on_panic
        },
    }
}

/// Map an error to its FFI code by variant.
///
/// This used to grep the lowercased `Display` string for `"io"`, which the
/// core I/O error (`"I/O error: …"`) does not contain — so permission
/// denied, disk full or a missing directory surfaced as
/// `OFFICE_ERR_INTERNAL`, and any message mentioning "invalid" became a
/// parse error.
fn classify_error(e: &crate::OfficeError) -> i32 {
    use crate::OfficeError as E;
    match e {
        E::Core(c) => classify_core(c),
        E::Docx(crate::docx::DocxError::Core(c)) => classify_core(c),
        E::Xlsx(crate::xlsx::XlsxError::Core(c)) => classify_core(c),
        E::Xlsx(crate::xlsx::XlsxError::InvalidCellRef(_)) => OFFICE_ERR_INVALID_ARG,
        E::Pptx(crate::pptx::PptxError::Core(c)) => classify_core(c),
        E::Doc(crate::doc::DocError::Io(_)) => OFFICE_ERR_IO,
        E::Doc(crate::doc::DocError::Cfb(c)) => classify_cfb(c),
        E::Doc(crate::doc::DocError::Encrypted | crate::doc::DocError::UnsupportedVersion(_)) => {
            OFFICE_ERR_UNSUPPORTED
        },
        E::Xls(crate::xls::XlsError::Io(_)) => OFFICE_ERR_IO,
        E::Xls(crate::xls::XlsError::Cfb(c)) => classify_cfb(c),
        E::Xls(crate::xls::XlsError::Encrypted | crate::xls::XlsError::UnsupportedVersion(_)) => {
            OFFICE_ERR_UNSUPPORTED
        },
        E::Ppt(crate::ppt::PptError::Io(_)) => OFFICE_ERR_IO,
        E::Ppt(crate::ppt::PptError::Cfb(c)) => classify_cfb(c),
        E::Ppt(crate::ppt::PptError::Encrypted) => OFFICE_ERR_UNSUPPORTED,
        E::UnsupportedFormat(_) => OFFICE_ERR_UNSUPPORTED,
        E::Panic(_) => OFFICE_ERR_INTERNAL,
        // Every remaining format-level variant describes content the parser
        // could not accept.
        _ => OFFICE_ERR_PARSE,
    }
}

fn classify_core(e: &crate::core::Error) -> i32 {
    use crate::core::Error as C;
    match e {
        C::Io(_) | C::Zip(zip::result::ZipError::Io(_)) => OFFICE_ERR_IO,
        C::InvalidArgument(_) => OFFICE_ERR_INVALID_ARG,
        C::Unsupported(_) => OFFICE_ERR_UNSUPPORTED,
        _ => OFFICE_ERR_PARSE,
    }
}

fn classify_cfb(e: &crate::cfb::CfbError) -> i32 {
    match e {
        crate::cfb::CfbError::Io(_) => OFFICE_ERR_IO,
        _ => OFFICE_ERR_PARSE,
    }
}

/// Convert text to a C string. A NUL byte — which a C string cannot carry,
/// and which would otherwise silently cut the text short at that point — is
/// replaced by U+FFFD, as documented in the header.
fn to_c_string(s: &str) -> *mut c_char {
    let cleaned: String = s.replace('\0', "\u{FFFD}");
    match CString::new(cleaned) {
        Ok(cs) => cs.into_raw(),
        Err(_) => ptr::null_mut(),
    }
}

fn cstr_to_str<'a>(ptr: *const c_char) -> Option<&'a str> {
    if ptr.is_null() {
        return None;
    }
    unsafe { CStr::from_ptr(ptr).to_str().ok() }
}

fn cstr_to_pathbuf(ptr: *const c_char) -> Option<PathBuf> {
    cstr_to_str(ptr).map(PathBuf::from)
}

/// Hand a byte buffer to the caller: its length goes to `out_len`, and it
/// is freed with `office_oxide_free_bytes(ptr, len)`.
///
/// Converted to a boxed slice so the allocation's size is exactly `len`.
/// The old `shrink_to_fit(); mem::forget` relied on `shrink_to_fit` making
/// capacity equal length, which it explicitly does not guarantee; freeing
/// with a `len`-sized layout after that is undefined behaviour under any
/// allocator that honours layouts.
fn into_ffi_bytes(bytes: Vec<u8>, out_len: *mut usize) -> *mut u8 {
    let boxed = bytes.into_boxed_slice();
    let len = boxed.len();
    unsafe { *out_len = len };
    Box::into_raw(boxed) as *mut u8
}

// ─── Version / memory ──────────────────────────────────────────────────────

static VERSION: &[u8] = concat!(env!("CARGO_PKG_VERSION"), "\0").as_bytes();

/// Return the library version as a NUL-terminated C string. Do not free.
#[unsafe(no_mangle)]
pub extern "C" fn office_oxide_version() -> *const c_char {
    guard(ptr::null_mut(), ptr::null(), || VERSION.as_ptr() as *const c_char)
}

/// Free a string returned by any FFI function.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn office_oxide_free_string(ptr: *mut c_char) {
    guard(ptr::null_mut(), (), || {
        if !ptr.is_null() {
            drop(unsafe { CString::from_raw(ptr) });
        }
    })
}

/// Free a byte buffer returned by an FFI function.
///
/// `len` must match the `out_len` returned alongside the pointer.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn office_oxide_free_bytes(ptr: *mut u8, len: usize) {
    guard(ptr::null_mut(), (), || {
        if !ptr.is_null() && len > 0 {
            drop(unsafe { Box::from_raw(ptr::slice_from_raw_parts_mut(ptr, len)) });
        }
    })
}

// ─── Format detection ──────────────────────────────────────────────────────

/// Detect document format from a file path. Returns the extension as a static
/// C string ("docx", "xlsx", etc.) or NULL if unsupported. Do not free.
#[unsafe(no_mangle)]
pub extern "C" fn office_oxide_detect_format(path: *const c_char) -> *const c_char {
    guard(ptr::null_mut(), ptr::null(), || {
        let Some(path) = cstr_to_pathbuf(path) else {
            return ptr::null();
        };
        match DocumentFormat::from_path(&path) {
            Some(f) => format_to_cstr(f),
            None => ptr::null(),
        }
    })
}

fn format_to_cstr(f: DocumentFormat) -> *const c_char {
    static DOCX: &[u8] = b"docx\0";
    static XLSX: &[u8] = b"xlsx\0";
    static PPTX: &[u8] = b"pptx\0";
    static DOC: &[u8] = b"doc\0";
    static XLS: &[u8] = b"xls\0";
    static PPT: &[u8] = b"ppt\0";
    let s: &[u8] = match f {
        DocumentFormat::Docx => DOCX,
        DocumentFormat::Xlsx => XLSX,
        DocumentFormat::Pptx => PPTX,
        DocumentFormat::Doc => DOC,
        DocumentFormat::Xls => XLS,
        DocumentFormat::Ppt => PPT,
    };
    s.as_ptr() as *const c_char
}

fn parse_format(s: &str) -> Option<DocumentFormat> {
    DocumentFormat::from_extension(s)
}

// ─── Document (read-only) ───────────────────────────────────────────────────

/// Opaque handle for a read-only Document.
///
/// Not safe to share across threads without external synchronization: see
/// the module-level "Thread-safety Convention". Concurrent calls on the
/// *same* handle are undefined behaviour; distinct handles are independent.
pub struct OfficeDocumentHandle {
    _doc: Document,
}

/// Open a document from a file path. Format is detected from the extension.
#[unsafe(no_mangle)]
pub extern "C" fn office_document_open(
    path: *const c_char,
    error_code: *mut i32,
) -> *mut OfficeDocumentHandle {
    guard(error_code, ptr::null_mut(), || {
        let Some(path) = cstr_to_pathbuf(path) else {
            set_err(error_code, OFFICE_ERR_INVALID_ARG);
            return ptr::null_mut();
        };
        match Document::open(&path) {
            Ok(doc) => {
                set_err(error_code, OFFICE_OK);
                Box::into_raw(Box::new(OfficeDocumentHandle { _doc: doc })) as *mut _
            },
            Err(e) => {
                set_err(error_code, classify_error(&e));
                ptr::null_mut()
            },
        }
    })
}

/// Open a document from an in-memory byte buffer.
///
/// `format` must be one of "docx", "xlsx", "pptx", "doc", "xls", "ppt".
#[unsafe(no_mangle)]
pub extern "C" fn office_document_open_from_bytes(
    data: *const u8,
    len: usize,
    format: *const c_char,
    error_code: *mut i32,
) -> *mut OfficeDocumentHandle {
    guard(error_code, ptr::null_mut(), || {
        if data.is_null() || len == 0 {
            set_err(error_code, OFFICE_ERR_INVALID_ARG);
            return ptr::null_mut();
        }
        let Some(fmt_str) = cstr_to_str(format) else {
            set_err(error_code, OFFICE_ERR_INVALID_ARG);
            return ptr::null_mut();
        };
        let Some(fmt) = parse_format(fmt_str) else {
            set_err(error_code, OFFICE_ERR_UNSUPPORTED);
            return ptr::null_mut();
        };
        let bytes = unsafe { slice::from_raw_parts(data, len) }.to_vec();
        let cursor = std::io::Cursor::new(bytes);
        match Document::from_reader(cursor, fmt) {
            Ok(doc) => {
                set_err(error_code, OFFICE_OK);
                Box::into_raw(Box::new(OfficeDocumentHandle { _doc: doc })) as *mut _
            },
            Err(e) => {
                set_err(error_code, classify_error(&e));
                ptr::null_mut()
            },
        }
    })
}

/// Free a document handle.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn office_document_free(handle: *mut OfficeDocumentHandle) {
    guard(ptr::null_mut(), (), || {
        if !handle.is_null() {
            drop(unsafe { Box::from_raw(handle) });
        }
    })
}

/// Return the document format as a static C string. Do not free. Returns NULL on invalid handle.
#[unsafe(no_mangle)]
pub extern "C" fn office_document_format(handle: *const OfficeDocumentHandle) -> *const c_char {
    guard(ptr::null_mut(), ptr::null(), || {
        if handle.is_null() {
            return ptr::null();
        }
        let h = unsafe { &*handle };
        format_to_cstr(h._doc.format())
    })
}

/// Render a document handle to a C string with `render`.
fn render_document(
    handle: *const OfficeDocumentHandle,
    error_code: *mut i32,
    render: impl FnOnce(&Document) -> String,
) -> *mut c_char {
    guard(error_code, ptr::null_mut(), || {
        if handle.is_null() {
            set_err(error_code, OFFICE_ERR_INVALID_ARG);
            return ptr::null_mut();
        }
        let h = unsafe { &*handle };
        let s = render(&h._doc);
        set_err(error_code, OFFICE_OK);
        to_c_string(&s)
    })
}

/// Extract plain text. Returns a heap-allocated C string — free with `office_oxide_free_string`.
#[unsafe(no_mangle)]
pub extern "C" fn office_document_plain_text(
    handle: *const OfficeDocumentHandle,
    error_code: *mut i32,
) -> *mut c_char {
    render_document(handle, error_code, Document::plain_text)
}

/// Convert to Markdown. Free with `office_oxide_free_string`.
#[unsafe(no_mangle)]
pub extern "C" fn office_document_to_markdown(
    handle: *const OfficeDocumentHandle,
    error_code: *mut i32,
) -> *mut c_char {
    render_document(handle, error_code, Document::to_markdown)
}

/// Convert to Markdown with every image embedded inline as
/// `[image-base64:<data>]` at its position in the document flow. Plain
/// `office_document_to_markdown` drops images entirely. Free with
/// `office_oxide_free_string`.
#[unsafe(no_mangle)]
pub extern "C" fn office_document_to_markdown_with_images(
    handle: *const OfficeDocumentHandle,
    error_code: *mut i32,
) -> *mut c_char {
    render_document(handle, error_code, |doc| {
        use crate::ir_render::{ImageEmbed, MarkdownOptions};
        doc.to_markdown_with(MarkdownOptions {
            image_embed: ImageEmbed::Base64,
        })
    })
}

/// Convert to HTML fragment. Free with `office_oxide_free_string`.
#[unsafe(no_mangle)]
pub extern "C" fn office_document_to_html(
    handle: *const OfficeDocumentHandle,
    error_code: *mut i32,
) -> *mut c_char {
    render_document(handle, error_code, Document::to_html)
}

/// Convert to the document IR, serialized as JSON. Free with `office_oxide_free_string`.
#[unsafe(no_mangle)]
pub extern "C" fn office_document_to_ir_json(
    handle: *const OfficeDocumentHandle,
    error_code: *mut i32,
) -> *mut c_char {
    guard(error_code, ptr::null_mut(), || {
        if handle.is_null() {
            set_err(error_code, OFFICE_ERR_INVALID_ARG);
            return ptr::null_mut();
        }
        let h = unsafe { &*handle };
        let ir = h._doc.to_ir();
        match serde_json::to_string(&ir) {
            Ok(s) => {
                set_err(error_code, OFFICE_OK);
                to_c_string(&s)
            },
            Err(_) => {
                // The document parsed, but its content could not be
                // produced in the requested form.
                set_err(error_code, OFFICE_ERR_EXTRACTION);
                ptr::null_mut()
            },
        }
    })
}

/// Map a unit result to a status, recording it in `error_code` too.
fn status(error_code: *mut i32, result: crate::Result<()>) -> i32 {
    let code = match result {
        Ok(()) => OFFICE_OK,
        Err(e) => classify_error(&e),
    };
    set_err(error_code, code);
    code
}

/// Save/convert the document to a file. Target format is detected from the extension.
/// Returns 0 on success, nonzero error code on failure.
#[unsafe(no_mangle)]
pub extern "C" fn office_document_save_as(
    handle: *const OfficeDocumentHandle,
    path: *const c_char,
    error_code: *mut i32,
) -> i32 {
    guard(error_code, OFFICE_ERR_INTERNAL, || {
        if handle.is_null() {
            set_err(error_code, OFFICE_ERR_INVALID_ARG);
            return OFFICE_ERR_INVALID_ARG;
        }
        let Some(path) = cstr_to_pathbuf(path) else {
            set_err(error_code, OFFICE_ERR_INVALID_ARG);
            return OFFICE_ERR_INVALID_ARG;
        };
        let h = unsafe { &*handle };
        status(error_code, h._doc.save_as(&path))
    })
}

// ─── EditableDocument ──────────────────────────────────────────────────────

/// Opaque handle for an editable document.
///
/// Not safe to share across threads without external synchronization: see
/// the module-level "Thread-safety Convention". Every mutating call takes
/// `&mut` to the handle's contents, so two concurrent calls on the same
/// handle are a data race the caller must prevent with its own mutex.
pub struct OfficeEditableHandle {
    doc: EditableDocument,
}

/// Open a document for editing. Supports DOCX, XLSX, PPTX.
#[unsafe(no_mangle)]
pub extern "C" fn office_editable_open(
    path: *const c_char,
    error_code: *mut i32,
) -> *mut OfficeEditableHandle {
    guard(error_code, ptr::null_mut(), || {
        let Some(path) = cstr_to_pathbuf(path) else {
            set_err(error_code, OFFICE_ERR_INVALID_ARG);
            return ptr::null_mut();
        };
        match EditableDocument::open(&path) {
            Ok(doc) => {
                set_err(error_code, OFFICE_OK);
                Box::into_raw(Box::new(OfficeEditableHandle { doc })) as *mut _
            },
            Err(e) => {
                set_err(error_code, classify_error(&e));
                ptr::null_mut()
            },
        }
    })
}

/// Open an editable document from a byte buffer.
#[unsafe(no_mangle)]
pub extern "C" fn office_editable_open_from_bytes(
    data: *const u8,
    len: usize,
    format: *const c_char,
    error_code: *mut i32,
) -> *mut OfficeEditableHandle {
    guard(error_code, ptr::null_mut(), || {
        if data.is_null() || len == 0 {
            set_err(error_code, OFFICE_ERR_INVALID_ARG);
            return ptr::null_mut();
        }
        let Some(fmt_str) = cstr_to_str(format) else {
            set_err(error_code, OFFICE_ERR_INVALID_ARG);
            return ptr::null_mut();
        };
        let Some(fmt) = parse_format(fmt_str) else {
            set_err(error_code, OFFICE_ERR_UNSUPPORTED);
            return ptr::null_mut();
        };
        let bytes = unsafe { slice::from_raw_parts(data, len) }.to_vec();
        let cursor = std::io::Cursor::new(bytes);
        match EditableDocument::from_reader(cursor, fmt) {
            Ok(doc) => {
                set_err(error_code, OFFICE_OK);
                Box::into_raw(Box::new(OfficeEditableHandle { doc })) as *mut _
            },
            Err(e) => {
                set_err(error_code, classify_error(&e));
                ptr::null_mut()
            },
        }
    })
}

/// Free an editable document handle.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn office_editable_free(handle: *mut OfficeEditableHandle) {
    guard(ptr::null_mut(), (), || {
        if !handle.is_null() {
            drop(unsafe { Box::from_raw(handle) });
        }
    })
}

/// Replace every occurrence of `find` with `replace` in text content.
/// Returns the number of replacements, or -1 on error.
///
/// An empty `find` is `OFFICE_ERR_INVALID_ARG`; an XLSX document (which has
/// no text replacement) is `OFFICE_ERR_UNSUPPORTED`.
#[unsafe(no_mangle)]
pub extern "C" fn office_editable_replace_text(
    handle: *mut OfficeEditableHandle,
    find: *const c_char,
    replace: *const c_char,
    error_code: *mut i32,
) -> i64 {
    guard(error_code, -1, || {
        if handle.is_null() {
            set_err(error_code, OFFICE_ERR_INVALID_ARG);
            return -1;
        }
        let Some(find_s) = cstr_to_str(find) else {
            set_err(error_code, OFFICE_ERR_INVALID_ARG);
            return -1;
        };
        let Some(replace_s) = cstr_to_str(replace) else {
            set_err(error_code, OFFICE_ERR_INVALID_ARG);
            return -1;
        };
        let h = unsafe { &mut *handle };
        match h.doc.replace_text(find_s, replace_s) {
            Ok(n) => {
                set_err(error_code, OFFICE_OK);
                n as i64
            },
            Err(e) => {
                set_err(error_code, classify_error(&e));
                -1
            },
        }
    })
}

/// Set a cell value in an XLSX document.
///
/// `value_type` is one of: 0 = empty, 1 = string, 2 = number, 3 = boolean.
/// `value_str` is used for strings (types 1) and ignored otherwise (pass NULL).
/// `value_num` is used for numbers (type 2) and booleans (type 3, nonzero = true).
#[unsafe(no_mangle)]
pub extern "C" fn office_editable_set_cell(
    handle: *mut OfficeEditableHandle,
    sheet_index: u32,
    cell_ref: *const c_char,
    value_type: i32,
    value_str: *const c_char,
    value_num: f64,
    error_code: *mut i32,
) -> i32 {
    guard(error_code, OFFICE_ERR_INTERNAL, || {
        if handle.is_null() {
            set_err(error_code, OFFICE_ERR_INVALID_ARG);
            return OFFICE_ERR_INVALID_ARG;
        }
        let Some(cell) = cstr_to_str(cell_ref) else {
            set_err(error_code, OFFICE_ERR_INVALID_ARG);
            return OFFICE_ERR_INVALID_ARG;
        };
        let value = match value_type {
            0 => crate::xlsx::edit::CellValue::Empty,
            1 => {
                let Some(s) = cstr_to_str(value_str) else {
                    set_err(error_code, OFFICE_ERR_INVALID_ARG);
                    return OFFICE_ERR_INVALID_ARG;
                };
                crate::xlsx::edit::CellValue::String(s.to_string())
            },
            2 => crate::xlsx::edit::CellValue::Number(value_num),
            3 => crate::xlsx::edit::CellValue::Boolean(value_num != 0.0),
            _ => {
                set_err(error_code, OFFICE_ERR_INVALID_ARG);
                return OFFICE_ERR_INVALID_ARG;
            },
        };
        let h = unsafe { &mut *handle };
        status(error_code, h.doc.set_cell(sheet_index as usize, cell, value))
    })
}

/// Save the edited document to a file.
#[unsafe(no_mangle)]
pub extern "C" fn office_editable_save(
    handle: *const OfficeEditableHandle,
    path: *const c_char,
    error_code: *mut i32,
) -> i32 {
    guard(error_code, OFFICE_ERR_INTERNAL, || {
        if handle.is_null() {
            set_err(error_code, OFFICE_ERR_INVALID_ARG);
            return OFFICE_ERR_INVALID_ARG;
        }
        let Some(path) = cstr_to_pathbuf(path) else {
            set_err(error_code, OFFICE_ERR_INVALID_ARG);
            return OFFICE_ERR_INVALID_ARG;
        };
        let h = unsafe { &*handle };
        status(error_code, h.doc.save(&path))
    })
}

/// Save the edited document into a heap-allocated byte buffer.
/// Returns a pointer and writes the length to `out_len`. Free with `office_oxide_free_bytes`.
#[unsafe(no_mangle)]
pub extern "C" fn office_editable_save_to_bytes(
    handle: *const OfficeEditableHandle,
    out_len: *mut usize,
    error_code: *mut i32,
) -> *mut u8 {
    guard(error_code, ptr::null_mut(), || {
        if handle.is_null() || out_len.is_null() {
            set_err(error_code, OFFICE_ERR_INVALID_ARG);
            return ptr::null_mut();
        }
        let h = unsafe { &*handle };
        let mut cursor = std::io::Cursor::new(Vec::new());
        match h.doc.write_to(&mut cursor) {
            Ok(()) => {
                set_err(error_code, OFFICE_OK);
                into_ffi_bytes(cursor.into_inner(), out_len)
            },
            Err(e) => {
                set_err(error_code, classify_error(&e));
                ptr::null_mut()
            },
        }
    })
}

// ─── Convenience one-shot helpers ───────────────────────────────────────────

/// Run a one-shot path → string conversion.
fn one_shot(
    path: *const c_char,
    error_code: *mut i32,
    convert: impl FnOnce(&std::path::Path) -> crate::Result<String>,
) -> *mut c_char {
    guard(error_code, ptr::null_mut(), || {
        let Some(path) = cstr_to_pathbuf(path) else {
            set_err(error_code, OFFICE_ERR_INVALID_ARG);
            return ptr::null_mut();
        };
        match convert(&path) {
            Ok(s) => {
                set_err(error_code, OFFICE_OK);
                to_c_string(&s)
            },
            Err(e) => {
                set_err(error_code, classify_error(&e));
                ptr::null_mut()
            },
        }
    })
}

/// One-shot: open a file, extract plain text, return. Free the result with
/// `office_oxide_free_string`.
#[unsafe(no_mangle)]
pub extern "C" fn office_extract_text(path: *const c_char, error_code: *mut i32) -> *mut c_char {
    one_shot(path, error_code, |p| crate::extract_text(p))
}

/// One-shot: open a file, convert to markdown, return. Free with `office_oxide_free_string`.
#[unsafe(no_mangle)]
pub extern "C" fn office_to_markdown(path: *const c_char, error_code: *mut i32) -> *mut c_char {
    one_shot(path, error_code, |p| crate::to_markdown(p))
}

/// One-shot: open a file, convert to HTML, return. Free with `office_oxide_free_string`.
#[unsafe(no_mangle)]
pub extern "C" fn office_to_html(path: *const c_char, error_code: *mut i32) -> *mut c_char {
    one_shot(path, error_code, |p| crate::to_html(p))
}

// ─── XlsxWriter ─────────────────────────────────────────────────────────────

/// Opaque handle wrapping an XlsxWriter.
pub struct OfficeXlsxWriterHandle {
    writer: crate::xlsx::write::XlsxWriter,
}

/// Create a new XLSX writer. Free with `office_xlsx_writer_free`.
#[unsafe(no_mangle)]
pub extern "C" fn office_xlsx_writer_new() -> *mut OfficeXlsxWriterHandle {
    guard(ptr::null_mut(), ptr::null_mut(), || {
        Box::into_raw(Box::new(OfficeXlsxWriterHandle {
            writer: crate::xlsx::write::XlsxWriter::new(),
        }))
    })
}

/// Free an XLSX writer handle.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn office_xlsx_writer_free(handle: *mut OfficeXlsxWriterHandle) {
    guard(ptr::null_mut(), (), || {
        if !handle.is_null() {
            drop(unsafe { Box::from_raw(handle) });
        }
    })
}

/// Add a sheet by name; returns its 0-based index, or u32::MAX on null handle.
#[unsafe(no_mangle)]
pub extern "C" fn office_xlsx_writer_add_sheet(
    handle: *mut OfficeXlsxWriterHandle,
    name: *const c_char,
) -> u32 {
    guard(ptr::null_mut(), u32::MAX, || {
        if handle.is_null() {
            return u32::MAX;
        }
        let name = cstr_to_str(name).unwrap_or("Sheet");
        let h = unsafe { &mut *handle };
        h.writer.add_sheet_get_index(name) as u32
    })
}

/// Decode an FFI cell value.
///
/// Returns `None` for an unrecognised `value_type`. Mapping it to `Empty`
/// meant a caller passing a bad type erased whatever was already in the cell,
/// while the function returned nothing to say so.
fn ffi_cell_data(
    value_type: i32,
    value_str: *const c_char,
    value_num: f64,
) -> Option<crate::xlsx::write::CellData> {
    use crate::xlsx::write::CellData;
    match value_type {
        0 => Some(CellData::Empty),
        1 => Some(CellData::String(cstr_to_str(value_str)?.to_string())),
        2 => Some(CellData::Number(value_num)),
        3 => Some(CellData::Boolean(value_num != 0.0)),
        4 => Some(CellData::Formula(cstr_to_str(value_str)?.to_string())),
        _ => None,
    }
}

/// Set a cell value.
///
/// `value_type`: 0=empty, 1=string (`value_str`), 2=number (`value_num`),
/// 3=boolean (`value_num` != 0), 4=formula (`value_str`; a leading `=` is
/// accepted and dropped).
///
/// Returns `OFFICE_OK`, `OFFICE_ERR_INVALID_ARG` for a bad handle or type, or
/// `OFFICE_ERR_UNSUPPORTED` when the target cell is outside Excel's grid and
/// nothing was written.
#[unsafe(no_mangle)]
pub extern "C" fn office_xlsx_sheet_set_cell(
    handle: *mut OfficeXlsxWriterHandle,
    sheet: u32,
    row: u32,
    col: u32,
    value_type: i32,
    value_str: *const c_char,
    value_num: f64,
) -> i32 {
    guard(ptr::null_mut(), OFFICE_ERR_INTERNAL, || {
        if handle.is_null() {
            return OFFICE_ERR_INVALID_ARG;
        }
        let Some(data) = ffi_cell_data(value_type, value_str, value_num) else {
            return OFFICE_ERR_INVALID_ARG;
        };
        let h = unsafe { &mut *handle };
        if h.writer
            .sheet_set_cell(sheet as usize, row as usize, col as usize, data)
        {
            OFFICE_OK
        } else {
            OFFICE_ERR_UNSUPPORTED
        }
    })
}

/// Set a cell with styling. bold applies bold weight; bg_color is a 6-char hex
/// string ("D3D3D3") or NULL for no background fill. Same `value_type`s and
/// return statuses as `office_xlsx_sheet_set_cell`.
#[unsafe(no_mangle)]
pub extern "C" fn office_xlsx_sheet_set_cell_styled(
    handle: *mut OfficeXlsxWriterHandle,
    sheet: u32,
    row: u32,
    col: u32,
    value_type: i32,
    value_str: *const c_char,
    value_num: f64,
    bold: bool,
    bg_color: *const c_char,
) -> i32 {
    guard(ptr::null_mut(), OFFICE_ERR_INTERNAL, || {
        if handle.is_null() {
            return OFFICE_ERR_INVALID_ARG;
        }
        use crate::xlsx::write::CellStyle;
        let Some(data) = ffi_cell_data(value_type, value_str, value_num) else {
            return OFFICE_ERR_INVALID_ARG;
        };
        let mut style = CellStyle::new();
        if bold {
            style = style.bold();
        }
        if let Some(bg) = cstr_to_str(bg_color) {
            if !bg.is_empty() {
                style = style.background(bg.to_string());
            }
        }
        let h = unsafe { &mut *handle };
        if h.writer
            .sheet_set_cell_styled(sheet as usize, row as usize, col as usize, data, style)
        {
            OFFICE_OK
        } else {
            OFFICE_ERR_UNSUPPORTED
        }
    })
}

/// Merge a rectangular range. row_span / col_span must be >= 1.
#[unsafe(no_mangle)]
pub extern "C" fn office_xlsx_sheet_merge_cells(
    handle: *mut OfficeXlsxWriterHandle,
    sheet: u32,
    row: u32,
    col: u32,
    row_span: u32,
    col_span: u32,
) {
    guard(ptr::null_mut(), (), || {
        if handle.is_null() {
            return;
        }
        let h = unsafe { &mut *handle };
        h.writer.sheet_merge_cells(
            sheet as usize,
            row as usize,
            col as usize,
            row_span as usize,
            col_span as usize,
        );
    })
}

/// Set column width in Excel character units (e.g. 20.0).
#[unsafe(no_mangle)]
pub extern "C" fn office_xlsx_sheet_set_column_width(
    handle: *mut OfficeXlsxWriterHandle,
    sheet: u32,
    col: u32,
    width: f64,
) {
    guard(ptr::null_mut(), (), || {
        if handle.is_null() {
            return;
        }
        let h = unsafe { &mut *handle };
        h.writer
            .sheet_set_column_width(sheet as usize, col as usize, width);
    })
}

/// Save to a file. Returns OFFICE_OK (0) on success, otherwise the error's
/// code (e.g. `OFFICE_ERR_IO` when the file cannot be created).
#[unsafe(no_mangle)]
pub extern "C" fn office_xlsx_writer_save(
    handle: *const OfficeXlsxWriterHandle,
    path: *const c_char,
    error_code: *mut i32,
) -> i32 {
    guard(error_code, OFFICE_ERR_INTERNAL, || {
        if handle.is_null() {
            set_err(error_code, OFFICE_ERR_INVALID_ARG);
            return OFFICE_ERR_INVALID_ARG;
        }
        let Some(path) = cstr_to_pathbuf(path) else {
            set_err(error_code, OFFICE_ERR_INVALID_ARG);
            return OFFICE_ERR_INVALID_ARG;
        };
        let h = unsafe { &*handle };
        status(error_code, h.writer.save(&path).map_err(Into::into))
    })
}

/// Serialize to a heap byte buffer. Free with office_oxide_free_bytes.
#[unsafe(no_mangle)]
pub extern "C" fn office_xlsx_writer_to_bytes(
    handle: *const OfficeXlsxWriterHandle,
    out_len: *mut usize,
    error_code: *mut i32,
) -> *mut u8 {
    guard(error_code, ptr::null_mut(), || {
        if handle.is_null() || out_len.is_null() {
            set_err(error_code, OFFICE_ERR_INVALID_ARG);
            return ptr::null_mut();
        }
        let h = unsafe { &*handle };
        let mut cursor = std::io::Cursor::new(Vec::new());
        match h.writer.write_to(&mut cursor) {
            Ok(()) => {
                set_err(error_code, OFFICE_OK);
                into_ffi_bytes(cursor.into_inner(), out_len)
            },
            Err(e) => {
                set_err(error_code, classify_error(&e.into()));
                ptr::null_mut()
            },
        }
    })
}

// ─── PptxWriter ──────────────────────────────────────────────────────────────

/// Opaque handle wrapping a PptxWriter.
pub struct OfficePptxWriterHandle {
    writer: crate::pptx::write::PptxWriter,
}

/// Create a new PPTX writer. Free with `office_pptx_writer_free`.
#[unsafe(no_mangle)]
pub extern "C" fn office_pptx_writer_new() -> *mut OfficePptxWriterHandle {
    guard(ptr::null_mut(), ptr::null_mut(), || {
        Box::into_raw(Box::new(OfficePptxWriterHandle {
            writer: crate::pptx::write::PptxWriter::new(),
        }))
    })
}

/// Free a PPTX writer handle.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn office_pptx_writer_free(handle: *mut OfficePptxWriterHandle) {
    guard(ptr::null_mut(), (), || {
        if !handle.is_null() {
            drop(unsafe { Box::from_raw(handle) });
        }
    })
}

/// Override presentation canvas size. 914400 EMU = 1 inch.
#[unsafe(no_mangle)]
pub extern "C" fn office_pptx_writer_set_presentation_size(
    handle: *mut OfficePptxWriterHandle,
    cx: u64,
    cy: u64,
) {
    guard(ptr::null_mut(), (), || {
        if handle.is_null() {
            return;
        }
        let h = unsafe { &mut *handle };
        h.writer.set_presentation_size(cx, cy);
    })
}

/// Add a slide; returns its 0-based index, or u32::MAX on null handle.
#[unsafe(no_mangle)]
pub extern "C" fn office_pptx_writer_add_slide(handle: *mut OfficePptxWriterHandle) -> u32 {
    guard(ptr::null_mut(), u32::MAX, || {
        if handle.is_null() {
            return u32::MAX;
        }
        let h = unsafe { &mut *handle };
        h.writer.add_slide_get_index() as u32
    })
}

/// Set the slide title. Returns `OFFICE_OK`, `OFFICE_ERR_INVALID_ARG`, or
/// `OFFICE_ERR_UNSUPPORTED` when the slide does not exist and nothing was
/// written.
#[unsafe(no_mangle)]
pub extern "C" fn office_pptx_slide_set_title(
    handle: *mut OfficePptxWriterHandle,
    slide: u32,
    title: *const c_char,
) -> i32 {
    guard(ptr::null_mut(), OFFICE_ERR_INTERNAL, || {
        if handle.is_null() {
            return OFFICE_ERR_INVALID_ARG;
        }
        let Some(title) = cstr_to_str(title) else {
            return OFFICE_ERR_INVALID_ARG;
        };
        let h = unsafe { &mut *handle };
        if h.writer.slide_set_title(slide as usize, title) {
            OFFICE_OK
        } else {
            OFFICE_ERR_UNSUPPORTED
        }
    })
}

/// Add a plain text paragraph to the slide body. Same statuses as
/// `office_pptx_slide_set_title`.
#[unsafe(no_mangle)]
pub extern "C" fn office_pptx_slide_add_text(
    handle: *mut OfficePptxWriterHandle,
    slide: u32,
    text: *const c_char,
) -> i32 {
    guard(ptr::null_mut(), OFFICE_ERR_INTERNAL, || {
        if handle.is_null() {
            return OFFICE_ERR_INVALID_ARG;
        }
        let Some(text) = cstr_to_str(text) else {
            return OFFICE_ERR_INVALID_ARG;
        };
        let h = unsafe { &mut *handle };
        if h.writer.slide_add_text(slide as usize, text) {
            OFFICE_OK
        } else {
            OFFICE_ERR_UNSUPPORTED
        }
    })
}

fn parse_image_format(s: &str) -> Option<crate::ir::ImageFormat> {
    match s.to_ascii_lowercase().as_str() {
        "png" => Some(crate::ir::ImageFormat::Png),
        "jpeg" | "jpg" => Some(crate::ir::ImageFormat::Jpeg),
        "gif" => Some(crate::ir::ImageFormat::Gif),
        _ => None,
    }
}

/// Embed an image on a slide.
///
/// `data`/`len` are the raw image bytes (PNG, JPEG, or GIF).
/// `format` is "png", "jpeg"/"jpg", or "gif".
/// `x`, `y`, `cx`, `cy` are in EMU (914400 = 1 inch).
///
/// Returns `OFFICE_OK`; `OFFICE_ERR_INVALID_ARG` for a null handle, empty
/// data or an unrecognised `format`; `OFFICE_ERR_UNSUPPORTED` when the slide
/// does not exist. It returned nothing, so every one of those failures
/// dropped the image silently.
#[unsafe(no_mangle)]
pub extern "C" fn office_pptx_slide_add_image(
    handle: *mut OfficePptxWriterHandle,
    slide: u32,
    data: *const u8,
    len: usize,
    format: *const c_char,
    x: i64,
    y: i64,
    cx: u64,
    cy: u64,
) -> i32 {
    guard(ptr::null_mut(), OFFICE_ERR_INTERNAL, || {
        if handle.is_null() || data.is_null() || len == 0 {
            return OFFICE_ERR_INVALID_ARG;
        }
        let Some(fmt) = cstr_to_str(format).and_then(parse_image_format) else {
            return OFFICE_ERR_INVALID_ARG;
        };
        let bytes = unsafe { slice::from_raw_parts(data, len) }.to_vec();
        let h = unsafe { &mut *handle };
        if h.writer
            .slide_add_image(slide as usize, bytes, fmt, x, y, cx, cy)
        {
            OFFICE_OK
        } else {
            OFFICE_ERR_UNSUPPORTED
        }
    })
}

/// Save to a file. Returns OFFICE_OK (0) on success, otherwise the error's
/// code (e.g. `OFFICE_ERR_IO` when the file cannot be created).
#[unsafe(no_mangle)]
pub extern "C" fn office_pptx_writer_save(
    handle: *const OfficePptxWriterHandle,
    path: *const c_char,
    error_code: *mut i32,
) -> i32 {
    guard(error_code, OFFICE_ERR_INTERNAL, || {
        if handle.is_null() {
            set_err(error_code, OFFICE_ERR_INVALID_ARG);
            return OFFICE_ERR_INVALID_ARG;
        }
        let Some(path) = cstr_to_pathbuf(path) else {
            set_err(error_code, OFFICE_ERR_INVALID_ARG);
            return OFFICE_ERR_INVALID_ARG;
        };
        let h = unsafe { &*handle };
        status(error_code, h.writer.save(&path).map_err(Into::into))
    })
}

/// Serialize to a heap byte buffer. Free with office_oxide_free_bytes.
#[unsafe(no_mangle)]
pub extern "C" fn office_pptx_writer_to_bytes(
    handle: *const OfficePptxWriterHandle,
    out_len: *mut usize,
    error_code: *mut i32,
) -> *mut u8 {
    guard(error_code, ptr::null_mut(), || {
        if handle.is_null() || out_len.is_null() {
            set_err(error_code, OFFICE_ERR_INVALID_ARG);
            return ptr::null_mut();
        }
        let h = unsafe { &*handle };
        let mut cursor = std::io::Cursor::new(Vec::new());
        match h.writer.write_to(&mut cursor) {
            Ok(()) => {
                set_err(error_code, OFFICE_OK);
                into_ffi_bytes(cursor.into_inner(), out_len)
            },
            Err(e) => {
                set_err(error_code, classify_error(&e.into()));
                ptr::null_mut()
            },
        }
    })
}

/// One-shot: convert a Markdown string to an Office document at `path`.
///
/// `format` must be one of `"docx"`, `"xlsx"`, or `"pptx"` (case-insensitive).
/// Returns `OFFICE_OK` (0) on success, otherwise a positive error code.
#[unsafe(no_mangle)]
pub extern "C" fn office_create_from_markdown(
    markdown: *const c_char,
    format: *const c_char,
    path: *const c_char,
    error_code: *mut i32,
) -> i32 {
    guard(error_code, OFFICE_ERR_INTERNAL, || {
        let Some(md) = cstr_to_str(markdown) else {
            set_err(error_code, OFFICE_ERR_INVALID_ARG);
            return OFFICE_ERR_INVALID_ARG;
        };
        let Some(fmt_str) = cstr_to_str(format) else {
            set_err(error_code, OFFICE_ERR_INVALID_ARG);
            return OFFICE_ERR_INVALID_ARG;
        };
        let Some(out_path) = cstr_to_pathbuf(path) else {
            set_err(error_code, OFFICE_ERR_INVALID_ARG);
            return OFFICE_ERR_INVALID_ARG;
        };
        let doc_format = match fmt_str.to_ascii_lowercase().as_str() {
            "docx" => crate::DocumentFormat::Docx,
            "xlsx" => crate::DocumentFormat::Xlsx,
            "pptx" => crate::DocumentFormat::Pptx,
            _ => {
                set_err(error_code, OFFICE_ERR_INVALID_ARG);
                return OFFICE_ERR_INVALID_ARG;
            },
        };
        status(error_code, crate::create::create_from_markdown(md, doc_format, &out_path))
    })
}

#[cfg(test)]
mod write_status_tests {
    use super::*;

    /// The writer returns `bool` so a silent no-op is detectable. Reporting
    /// nothing across the FFI reintroduced the exact regression the Rust API
    /// was fixed for: an out-of-range index discarded every value it wrote
    /// while the caller saw success.
    #[test]
    fn test_an_out_of_range_write_reports_a_status_instead_of_vanishing() {
        let h = office_xlsx_writer_new();
        assert!(!h.is_null());
        let name = std::ffi::CString::new("Data").unwrap();
        assert_eq!(office_xlsx_writer_add_sheet(h, name.as_ptr()), 0);

        let text = std::ffi::CString::new("LOST").unwrap();
        // Sheet 1 does not exist.
        assert_ne!(
            office_xlsx_sheet_set_cell(h, 1, 0, 0, 1, text.as_ptr(), 0.0),
            OFFICE_OK,
            "writing to a missing sheet must not report success"
        );
        // Row past the end of Excel's grid.
        assert_ne!(
            office_xlsx_sheet_set_cell(h, 0, 1_048_576, 0, 1, text.as_ptr(), 0.0),
            OFFICE_OK,
            "writing outside the grid must not report success"
        );
        // A valid write still succeeds.
        assert_eq!(office_xlsx_sheet_set_cell(h, 0, 0, 0, 1, text.as_ptr(), 0.0), OFFICE_OK);

        // An unrecognised value_type must be rejected, not silently written
        // as Empty over the existing value.
        assert_ne!(
            office_xlsx_sheet_set_cell(h, 0, 0, 0, 7, text.as_ptr(), 0.0),
            OFFICE_OK,
            "an unknown value_type must be an error, not an erase"
        );

        unsafe { office_xlsx_writer_free(h) };
    }

    /// `CellData::Boolean` and `CellData::Formula` were unreachable from any
    /// binding: no `value_type` mapped to them.
    #[test]
    fn test_booleans_and_formulas_are_reachable_over_the_ffi() {
        let h = office_xlsx_writer_new();
        let name = std::ffi::CString::new("S").unwrap();
        office_xlsx_writer_add_sheet(h, name.as_ptr());
        let f = std::ffi::CString::new("SUM(A1:A2)").unwrap();
        assert_eq!(office_xlsx_sheet_set_cell(h, 0, 0, 0, 3, std::ptr::null(), 1.0), OFFICE_OK);
        assert_eq!(office_xlsx_sheet_set_cell(h, 0, 1, 0, 4, f.as_ptr(), 0.0), OFFICE_OK);
        unsafe { office_xlsx_writer_free(h) };
    }

    /// `office_pptx_slide_add_image` returned nothing, so a missing slide,
    /// an unknown format or empty data dropped the image silently.
    #[test]
    fn test_add_image_reports_a_status() {
        let h = office_pptx_writer_new();
        assert_eq!(office_pptx_writer_add_slide(h), 0);
        let png = std::ffi::CString::new("png").unwrap();
        let bmp = std::ffi::CString::new("bmp").unwrap();
        let data = [0x89u8, b'P', b'N', b'G'];
        let add = |slide, fmt: &std::ffi::CString, len| {
            office_pptx_slide_add_image(h, slide, data.as_ptr(), len, fmt.as_ptr(), 0, 0, 1, 1)
        };
        assert_eq!(add(0, &png, data.len()), OFFICE_OK);
        assert_eq!(add(5, &png, data.len()), OFFICE_ERR_UNSUPPORTED, "missing slide");
        assert_eq!(add(0, &bmp, data.len()), OFFICE_ERR_INVALID_ARG, "unknown format");
        assert_eq!(add(0, &png, 0), OFFICE_ERR_INVALID_ARG, "empty data");
        let title = std::ffi::CString::new("t").unwrap();
        assert_eq!(office_pptx_slide_set_title(h, 9, title.as_ptr()), OFFICE_ERR_UNSUPPORTED);
        assert_eq!(office_pptx_slide_add_text(h, 9, title.as_ptr()), OFFICE_ERR_UNSUPPORTED);
        unsafe { office_pptx_writer_free(h) };
    }
}

#[cfg(test)]
mod error_code_tests {
    use super::*;

    fn cstring(s: &str) -> CString {
        CString::new(s).unwrap()
    }

    /// A path whose parent is a regular file: creating it fails with an
    /// I/O error ("Not a directory") that no substring rule recognised.
    fn uncreatable_path(tag: &str) -> (std::path::PathBuf, CString) {
        let blocker = std::env::temp_dir()
            .join(format!("office_oxide_ffi_blocker_{tag}_{}", std::process::id()));
        std::fs::write(&blocker, b"x").unwrap();
        let target = blocker.join("out.xlsx");
        let c = cstring(target.to_str().unwrap());
        (blocker, c)
    }

    /// Error codes were picked by grepping the lowercased message for "io";
    /// the core I/O error renders as "I/O error: …", so disk-level failures
    /// came back as OFFICE_ERR_INTERNAL.
    #[test]
    fn test_classify_error_uses_the_variant_not_the_message() {
        use crate::OfficeError as E;
        let io = || std::io::Error::new(std::io::ErrorKind::PermissionDenied, "denied");
        let cases: Vec<(E, i32)> = vec![
            (E::Core(crate::core::Error::Io(io())), OFFICE_ERR_IO),
            (E::Xls(crate::xls::XlsError::Io(io())), OFFICE_ERR_IO),
            (E::Doc(crate::doc::DocError::Io(io())), OFFICE_ERR_IO),
            (E::Ppt(crate::ppt::PptError::Cfb(crate::cfb::CfbError::Io(io()))), OFFICE_ERR_IO),
            (
                E::Docx(crate::docx::DocxError::Core(crate::core::Error::Io(io()))),
                OFFICE_ERR_IO,
            ),
            (E::Core(crate::core::Error::InvalidArgument("x".into())), OFFICE_ERR_INVALID_ARG),
            (
                E::Xlsx(crate::xlsx::XlsxError::InvalidCellRef("ZZ".into())),
                OFFICE_ERR_INVALID_ARG,
            ),
            // Messages that contain "invalid" or "io" but are not those
            // classes.
            (E::Core(crate::core::Error::MalformedXml("ratio".into())), OFFICE_ERR_PARSE),
            (E::Doc(crate::doc::DocError::InvalidFib("x".into())), OFFICE_ERR_PARSE),
            (E::Xls(crate::xls::XlsError::Encrypted), OFFICE_ERR_UNSUPPORTED),
            (E::UnsupportedFormat("rtf".into()), OFFICE_ERR_UNSUPPORTED),
            (E::Panic("invalid io".into()), OFFICE_ERR_INTERNAL),
        ];
        for (err, want) in cases {
            assert_eq!(classify_error(&err), want, "{err:?}");
        }
    }

    /// The writers' save discarded the error and always reported
    /// OFFICE_ERR_INTERNAL.
    #[test]
    fn test_writer_save_failures_report_the_io_code() {
        let (blocker, path) = uncreatable_path("xlsx");
        let x = office_xlsx_writer_new();
        office_xlsx_writer_add_sheet(x, cstring("S").as_ptr());
        let mut err = -1;
        assert_eq!(office_xlsx_writer_save(x, path.as_ptr(), &mut err), OFFICE_ERR_IO);
        assert_eq!(err, OFFICE_ERR_IO);
        unsafe { office_xlsx_writer_free(x) };

        let p = office_pptx_writer_new();
        office_pptx_writer_add_slide(p);
        let mut err = -1;
        assert_eq!(office_pptx_writer_save(p, path.as_ptr(), &mut err), OFFICE_ERR_IO);
        assert_eq!(err, OFFICE_ERR_IO);
        unsafe { office_pptx_writer_free(p) };
        std::fs::remove_file(&blocker).ok();
    }

    fn docx_bytes() -> Vec<u8> {
        let mut w = crate::docx::write::DocxWriter::new();
        w.add_paragraph("Hello world");
        let mut buf = std::io::Cursor::new(Vec::new());
        w.write_to(&mut buf).unwrap();
        buf.into_inner()
    }

    /// Every replace_text failure was reported as OFFICE_ERR_UNSUPPORTED,
    /// so an empty search string looked like an unsupported format.
    #[test]
    fn test_replace_text_errors_carry_their_own_codes() {
        let data = docx_bytes();
        let mut err = -1;
        let h = office_editable_open_from_bytes(
            data.as_ptr(),
            data.len(),
            cstring("docx").as_ptr(),
            &mut err,
        );
        assert_eq!(err, OFFICE_OK);
        let n =
            office_editable_replace_text(h, cstring("").as_ptr(), cstring("X").as_ptr(), &mut err);
        assert_eq!((n, err), (-1, OFFICE_ERR_INVALID_ARG));
        let (blocker, path) = uncreatable_path("editable");
        assert_eq!(office_editable_save(h, path.as_ptr(), &mut err), OFFICE_ERR_IO);
        std::fs::remove_file(&blocker).ok();
        unsafe { office_editable_free(h) };
    }

    /// Markdown with embedded images was reachable only from Python and
    /// WASM; C and every FFI binding got image-stripped markdown.
    #[test]
    fn test_markdown_with_images_is_exported() {
        let mut w = crate::docx::write::DocxWriter::new();
        w.add_paragraph("Before");
        w.add_ir_image(&crate::ir::Image {
            data: Some(vec![0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A]),
            format: Some(crate::ir::ImageFormat::Png),
            display_width_emu: Some(9525),
            display_height_emu: Some(9525),
            ..Default::default()
        });
        let mut buf = std::io::Cursor::new(Vec::new());
        w.write_to(&mut buf).unwrap();
        let data = buf.into_inner();
        let mut err = -1;
        let h = office_document_open_from_bytes(
            data.as_ptr(),
            data.len(),
            cstring("docx").as_ptr(),
            &mut err,
        );
        assert_eq!(err, OFFICE_OK);
        let s = office_document_to_markdown_with_images(h, &mut err);
        assert_eq!(err, OFFICE_OK);
        let md = unsafe { CStr::from_ptr(s) }.to_str().unwrap().to_string();
        unsafe { office_oxide_free_string(s) };
        assert!(md.contains("[image-base64:"), "{md}");
        assert!(md.contains("Before"), "{md}");
        unsafe { office_document_free(h) };
    }

    /// A panic inside an exported function must come back as
    /// OFFICE_ERR_INTERNAL with the function's failure value, not unwind
    /// into the host.
    #[test]
    fn test_a_panic_is_contained_and_reported_as_internal() {
        let mut err = OFFICE_OK;
        let out: *mut c_char = guard(&mut err, ptr::null_mut(), || panic!("boom"));
        assert!(out.is_null());
        assert_eq!(err, OFFICE_ERR_INTERNAL);
        // No error_code pointer: still contained.
        assert_eq!(guard(ptr::null_mut(), -1i64, || panic!("boom")), -1);
    }

    /// Every exported function must run its body under `guard` (directly,
    /// or through one of the helpers that does). A new export added without
    /// it would reopen the host-abort path, so this is checked from the
    /// source rather than trusted to review.
    #[test]
    fn test_every_export_contains_panics() {
        let src = include_str!("ffi.rs");
        let helpers = ["guard(", "render_document(", "one_shot("];
        let mut checked = 0;
        // Split so this line does not match itself.
        let marker = concat!("#[unsafe(", "no_mangle)]");
        let mut rest = src;
        while let Some(at) = rest.find(marker) {
            rest = &rest[at + 1..];
            let body_start = rest.find('{').unwrap();
            let name_at = rest.find("fn ").unwrap() + 3;
            let name: String = rest[name_at..]
                .chars()
                .take_while(|c| c.is_alphanumeric() || *c == '_')
                .collect();
            let body = rest[body_start + 1..].trim_start();
            assert!(
                helpers.iter().any(|h| body.starts_with(h)),
                "`{name}` does not run its body under guard()"
            );
            checked += 1;
        }
        assert!(checked >= 40, "only {checked} exports found");
    }

    /// The header is hand-written. Every exported symbol must be declared in
    /// it and every declared function must exist, so a binding generated or
    /// written from the header cannot link against a symbol that is missing
    /// or has been renamed.
    #[test]
    fn test_header_declares_exactly_the_exported_functions() {
        let src = include_str!("ffi.rs");
        let header = include_str!("../include/office_oxide_c/office_oxide.h");
        let mut exported = std::collections::BTreeSet::new();
        for line in src.lines() {
            let line = line.trim_start();
            if let Some(rest) = line
                .strip_prefix("pub extern \"C\" fn ")
                .or_else(|| line.strip_prefix("pub unsafe extern \"C\" fn "))
            {
                let name: String = rest
                    .chars()
                    .take_while(|c| c.is_alphanumeric() || *c == '_')
                    .collect();
                exported.insert(name);
            }
        }
        let mut declared = std::collections::BTreeSet::new();
        for token in header.split(|c: char| !(c.is_alphanumeric() || c == '_')) {
            if (token.starts_with("office_")) && header.contains(&format!("{token}(")) {
                declared.insert(token.to_string());
            }
        }
        assert_eq!(exported, declared, "header and src/ffi.rs disagree");
    }

    /// Byte buffers must be freed with the layout they were allocated with.
    /// A buffer with spare capacity is the case `shrink_to_fit` did not
    /// guarantee to handle.
    #[test]
    fn test_byte_buffers_round_trip_through_free_bytes() {
        let mut v = Vec::with_capacity(4096);
        v.extend_from_slice(b"payload");
        let mut len = 0usize;
        let ptr = into_ffi_bytes(v, &mut len);
        assert_eq!(len, 7);
        assert_eq!(unsafe { slice::from_raw_parts(ptr, len) }, b"payload");
        unsafe { office_oxide_free_bytes(ptr, len) };
    }
}
