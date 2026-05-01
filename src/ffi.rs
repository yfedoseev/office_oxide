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
//! - 5 = internal error
//! - 6 = unsupported format / feature
//!
//! # Memory Convention
//! - Strings returned as `*mut c_char` are heap-allocated and must be freed with
//!   `office_oxide_free_string`.
//! - Byte buffers returned as `*mut u8` (with an `out_len`) must be freed with
//!   `office_oxide_free_bytes(ptr, len)`.
//! - Opaque handles (`*mut OfficeDocumentHandle`, `*mut OfficeEditableHandle`)
//!   must be freed with their corresponding `*_free` function.
#![allow(missing_docs)]
#![allow(clippy::missing_safety_doc)]
#![allow(clippy::not_unsafe_ptr_arg_deref)]
#![allow(clippy::too_many_arguments)]

use std::ffi::{CStr, CString};
use std::os::raw::c_char;
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

fn classify_error(e: &crate::OfficeError) -> i32 {
    match e {
        crate::OfficeError::UnsupportedFormat(_) => OFFICE_ERR_UNSUPPORTED,
        _ => {
            let msg = format!("{e}").to_lowercase();
            if msg.contains("not found") || msg.contains("no such file") || msg.contains("io") {
                OFFICE_ERR_IO
            } else if msg.contains("parse") || msg.contains("invalid") || msg.contains("xml") {
                OFFICE_ERR_PARSE
            } else {
                OFFICE_ERR_INTERNAL
            }
        },
    }
}

fn to_c_string(s: &str) -> *mut c_char {
    // Replace NUL bytes (invalid in C strings) with replacement char.
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

// ─── Version / memory ──────────────────────────────────────────────────────

/// Return the library version as a NUL-terminated C string. Do not free.
#[unsafe(no_mangle)]
pub extern "C" fn office_oxide_version() -> *const c_char {
    static VERSION: &[u8] = concat!(env!("CARGO_PKG_VERSION"), "\0").as_bytes();
    VERSION.as_ptr() as *const c_char
}

/// Free a string returned by any FFI function.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn office_oxide_free_string(ptr: *mut c_char) {
    if !ptr.is_null() {
        drop(unsafe { CString::from_raw(ptr) });
    }
}

/// Free a byte buffer returned by an FFI function.
///
/// `len` must match the `out_len` returned alongside the pointer.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn office_oxide_free_bytes(ptr: *mut u8, len: usize) {
    if !ptr.is_null() && len > 0 {
        drop(unsafe { Vec::from_raw_parts(ptr, len, len) });
    }
}

// ─── Format detection ──────────────────────────────────────────────────────

/// Detect document format from a file path. Returns the extension as a static
/// C string ("docx", "xlsx", etc.) or NULL if unsupported. Do not free.
#[unsafe(no_mangle)]
pub extern "C" fn office_oxide_detect_format(path: *const c_char) -> *const c_char {
    let Some(path) = cstr_to_pathbuf(path) else {
        return ptr::null();
    };
    match DocumentFormat::from_path(&path) {
        Some(f) => format_to_cstr(f),
        None => ptr::null(),
    }
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
pub struct OfficeDocumentHandle {
    _doc: Document,
}

/// Open a document from a file path. Format is detected from the extension.
#[unsafe(no_mangle)]
pub extern "C" fn office_document_open(
    path: *const c_char,
    error_code: *mut i32,
) -> *mut OfficeDocumentHandle {
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
}

/// Free a document handle.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn office_document_free(handle: *mut OfficeDocumentHandle) {
    if !handle.is_null() {
        drop(unsafe { Box::from_raw(handle) });
    }
}

/// Return the document format as a static C string. Do not free. Returns NULL on invalid handle.
#[unsafe(no_mangle)]
pub extern "C" fn office_document_format(handle: *const OfficeDocumentHandle) -> *const c_char {
    if handle.is_null() {
        return ptr::null();
    }
    let h = unsafe { &*handle };
    format_to_cstr(h._doc.format())
}

/// Extract plain text. Returns a heap-allocated C string — free with `office_oxide_free_string`.
#[unsafe(no_mangle)]
pub extern "C" fn office_document_plain_text(
    handle: *const OfficeDocumentHandle,
    error_code: *mut i32,
) -> *mut c_char {
    if handle.is_null() {
        set_err(error_code, OFFICE_ERR_INVALID_ARG);
        return ptr::null_mut();
    }
    let h = unsafe { &*handle };
    let s = h._doc.plain_text();
    set_err(error_code, OFFICE_OK);
    to_c_string(&s)
}

/// Convert to Markdown. Free with `office_oxide_free_string`.
#[unsafe(no_mangle)]
pub extern "C" fn office_document_to_markdown(
    handle: *const OfficeDocumentHandle,
    error_code: *mut i32,
) -> *mut c_char {
    if handle.is_null() {
        set_err(error_code, OFFICE_ERR_INVALID_ARG);
        return ptr::null_mut();
    }
    let h = unsafe { &*handle };
    let s = h._doc.to_markdown();
    set_err(error_code, OFFICE_OK);
    to_c_string(&s)
}

/// Convert to HTML fragment. Free with `office_oxide_free_string`.
#[unsafe(no_mangle)]
pub extern "C" fn office_document_to_html(
    handle: *const OfficeDocumentHandle,
    error_code: *mut i32,
) -> *mut c_char {
    if handle.is_null() {
        set_err(error_code, OFFICE_ERR_INVALID_ARG);
        return ptr::null_mut();
    }
    let h = unsafe { &*handle };
    let s = h._doc.to_html();
    set_err(error_code, OFFICE_OK);
    to_c_string(&s)
}

/// Convert to the document IR, serialized as JSON. Free with `office_oxide_free_string`.
#[unsafe(no_mangle)]
pub extern "C" fn office_document_to_ir_json(
    handle: *const OfficeDocumentHandle,
    error_code: *mut i32,
) -> *mut c_char {
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
            set_err(error_code, OFFICE_ERR_INTERNAL);
            ptr::null_mut()
        },
    }
}

/// Save/convert the document to a file. Target format is detected from the extension.
/// Returns 0 on success, nonzero error code on failure.
#[unsafe(no_mangle)]
pub extern "C" fn office_document_save_as(
    handle: *const OfficeDocumentHandle,
    path: *const c_char,
    error_code: *mut i32,
) -> i32 {
    if handle.is_null() {
        set_err(error_code, OFFICE_ERR_INVALID_ARG);
        return OFFICE_ERR_INVALID_ARG;
    }
    let Some(path) = cstr_to_pathbuf(path) else {
        set_err(error_code, OFFICE_ERR_INVALID_ARG);
        return OFFICE_ERR_INVALID_ARG;
    };
    let h = unsafe { &*handle };
    match h._doc.save_as(&path) {
        Ok(()) => {
            set_err(error_code, OFFICE_OK);
            OFFICE_OK
        },
        Err(e) => {
            let c = classify_error(&e);
            set_err(error_code, c);
            c
        },
    }
}

// ─── EditableDocument ──────────────────────────────────────────────────────

/// Opaque handle for an editable document.
pub struct OfficeEditableHandle {
    doc: EditableDocument,
}

/// Open a document for editing. Supports DOCX, XLSX, PPTX.
#[unsafe(no_mangle)]
pub extern "C" fn office_editable_open(
    path: *const c_char,
    error_code: *mut i32,
) -> *mut OfficeEditableHandle {
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
}

/// Open an editable document from a byte buffer.
#[unsafe(no_mangle)]
pub extern "C" fn office_editable_open_from_bytes(
    data: *const u8,
    len: usize,
    format: *const c_char,
    error_code: *mut i32,
) -> *mut OfficeEditableHandle {
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
}

/// Free an editable document handle.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn office_editable_free(handle: *mut OfficeEditableHandle) {
    if !handle.is_null() {
        drop(unsafe { Box::from_raw(handle) });
    }
}

/// Replace every occurrence of `find` with `replace` in text content.
/// Returns the number of replacements, or -1 on error.
#[unsafe(no_mangle)]
pub extern "C" fn office_editable_replace_text(
    handle: *mut OfficeEditableHandle,
    find: *const c_char,
    replace: *const c_char,
    error_code: *mut i32,
) -> i64 {
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
    let n = h.doc.replace_text(find_s, replace_s);
    set_err(error_code, OFFICE_OK);
    n as i64
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
    match h.doc.set_cell(sheet_index as usize, cell, value) {
        Ok(()) => {
            set_err(error_code, OFFICE_OK);
            OFFICE_OK
        },
        Err(e) => {
            let c = classify_error(&e);
            set_err(error_code, c);
            c
        },
    }
}

/// Save the edited document to a file.
#[unsafe(no_mangle)]
pub extern "C" fn office_editable_save(
    handle: *const OfficeEditableHandle,
    path: *const c_char,
    error_code: *mut i32,
) -> i32 {
    if handle.is_null() {
        set_err(error_code, OFFICE_ERR_INVALID_ARG);
        return OFFICE_ERR_INVALID_ARG;
    }
    let Some(path) = cstr_to_pathbuf(path) else {
        set_err(error_code, OFFICE_ERR_INVALID_ARG);
        return OFFICE_ERR_INVALID_ARG;
    };
    let h = unsafe { &*handle };
    match h.doc.save(&path) {
        Ok(()) => {
            set_err(error_code, OFFICE_OK);
            OFFICE_OK
        },
        Err(e) => {
            let c = classify_error(&e);
            set_err(error_code, c);
            c
        },
    }
}

/// Save the edited document into a heap-allocated byte buffer.
/// Returns a pointer and writes the length to `out_len`. Free with `office_oxide_free_bytes`.
#[unsafe(no_mangle)]
pub extern "C" fn office_editable_save_to_bytes(
    handle: *const OfficeEditableHandle,
    out_len: *mut usize,
    error_code: *mut i32,
) -> *mut u8 {
    if handle.is_null() || out_len.is_null() {
        set_err(error_code, OFFICE_ERR_INVALID_ARG);
        return ptr::null_mut();
    }
    let h = unsafe { &*handle };
    let buf: Vec<u8> = Vec::new();
    let mut cursor = std::io::Cursor::new(buf);
    match h.doc.write_to(&mut cursor) {
        Ok(()) => {
            let mut bytes = cursor.into_inner();
            bytes.shrink_to_fit();
            let len = bytes.len();
            let ptr = bytes.as_mut_ptr();
            std::mem::forget(bytes);
            unsafe { *out_len = len };
            set_err(error_code, OFFICE_OK);
            ptr
        },
        Err(e) => {
            set_err(error_code, classify_error(&e));
            ptr::null_mut()
        },
    }
}

// ─── Convenience one-shot helpers ───────────────────────────────────────────

/// One-shot: open a file, extract plain text, return. Free the result with
/// `office_oxide_free_string`.
#[unsafe(no_mangle)]
pub extern "C" fn office_extract_text(path: *const c_char, error_code: *mut i32) -> *mut c_char {
    let Some(path) = cstr_to_pathbuf(path) else {
        set_err(error_code, OFFICE_ERR_INVALID_ARG);
        return ptr::null_mut();
    };
    match crate::extract_text(&path) {
        Ok(s) => {
            set_err(error_code, OFFICE_OK);
            to_c_string(&s)
        },
        Err(e) => {
            set_err(error_code, classify_error(&e));
            ptr::null_mut()
        },
    }
}

/// One-shot: open a file, convert to markdown, return. Free with `office_oxide_free_string`.
#[unsafe(no_mangle)]
pub extern "C" fn office_to_markdown(path: *const c_char, error_code: *mut i32) -> *mut c_char {
    let Some(path) = cstr_to_pathbuf(path) else {
        set_err(error_code, OFFICE_ERR_INVALID_ARG);
        return ptr::null_mut();
    };
    match crate::to_markdown(&path) {
        Ok(s) => {
            set_err(error_code, OFFICE_OK);
            to_c_string(&s)
        },
        Err(e) => {
            set_err(error_code, classify_error(&e));
            ptr::null_mut()
        },
    }
}

/// One-shot: open a file, convert to HTML, return. Free with `office_oxide_free_string`.
#[unsafe(no_mangle)]
pub extern "C" fn office_to_html(path: *const c_char, error_code: *mut i32) -> *mut c_char {
    let Some(path) = cstr_to_pathbuf(path) else {
        set_err(error_code, OFFICE_ERR_INVALID_ARG);
        return ptr::null_mut();
    };
    match crate::to_html(&path) {
        Ok(s) => {
            set_err(error_code, OFFICE_OK);
            to_c_string(&s)
        },
        Err(e) => {
            set_err(error_code, classify_error(&e));
            ptr::null_mut()
        },
    }
}

// ─── XlsxWriter ─────────────────────────────────────────────────────────────

/// Opaque handle wrapping an XlsxWriter.
pub struct OfficeXlsxWriterHandle {
    writer: crate::xlsx::write::XlsxWriter,
}

/// Create a new XLSX writer. Free with `office_xlsx_writer_free`.
#[unsafe(no_mangle)]
pub extern "C" fn office_xlsx_writer_new() -> *mut OfficeXlsxWriterHandle {
    Box::into_raw(Box::new(OfficeXlsxWriterHandle {
        writer: crate::xlsx::write::XlsxWriter::new(),
    }))
}

/// Free an XLSX writer handle.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn office_xlsx_writer_free(handle: *mut OfficeXlsxWriterHandle) {
    if !handle.is_null() {
        drop(unsafe { Box::from_raw(handle) });
    }
}

/// Add a sheet by name; returns its 0-based index, or u32::MAX on null handle.
#[unsafe(no_mangle)]
pub extern "C" fn office_xlsx_writer_add_sheet(
    handle: *mut OfficeXlsxWriterHandle,
    name: *const c_char,
) -> u32 {
    if handle.is_null() {
        return u32::MAX;
    }
    let name = cstr_to_str(name).unwrap_or("Sheet");
    let h = unsafe { &mut *handle };
    h.writer.add_sheet_get_index(name) as u32
}

/// Set a cell value. value_type: 0=empty, 1=string (value_str), 2=number (value_num).
#[unsafe(no_mangle)]
pub extern "C" fn office_xlsx_sheet_set_cell(
    handle: *mut OfficeXlsxWriterHandle,
    sheet: u32,
    row: u32,
    col: u32,
    value_type: i32,
    value_str: *const c_char,
    value_num: f64,
) {
    if handle.is_null() {
        return;
    }
    use crate::xlsx::write::CellData;
    let data = match value_type {
        1 => CellData::String(cstr_to_str(value_str).unwrap_or("").to_string()),
        2 => CellData::Number(value_num),
        _ => CellData::Empty,
    };
    let h = unsafe { &mut *handle };
    h.writer
        .sheet_set_cell(sheet as usize, row as usize, col as usize, data);
}

/// Set a cell with styling. bold applies bold weight; bg_color is a 6-char hex
/// string ("D3D3D3") or NULL for no background fill.
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
) {
    if handle.is_null() {
        return;
    }
    use crate::xlsx::write::{CellData, CellStyle};
    let data = match value_type {
        1 => CellData::String(cstr_to_str(value_str).unwrap_or("").to_string()),
        2 => CellData::Number(value_num),
        _ => CellData::Empty,
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
    h.writer
        .sheet_set_cell_styled(sheet as usize, row as usize, col as usize, data, style);
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
}

/// Set column width in Excel character units (e.g. 20.0).
#[unsafe(no_mangle)]
pub extern "C" fn office_xlsx_sheet_set_column_width(
    handle: *mut OfficeXlsxWriterHandle,
    sheet: u32,
    col: u32,
    width: f64,
) {
    if handle.is_null() {
        return;
    }
    let h = unsafe { &mut *handle };
    h.writer
        .sheet_set_column_width(sheet as usize, col as usize, width);
}

/// Save to a file. Returns OFFICE_OK (0) on success.
#[unsafe(no_mangle)]
pub extern "C" fn office_xlsx_writer_save(
    handle: *const OfficeXlsxWriterHandle,
    path: *const c_char,
    error_code: *mut i32,
) -> i32 {
    if handle.is_null() {
        set_err(error_code, OFFICE_ERR_INVALID_ARG);
        return OFFICE_ERR_INVALID_ARG;
    }
    let Some(path) = cstr_to_pathbuf(path) else {
        set_err(error_code, OFFICE_ERR_INVALID_ARG);
        return OFFICE_ERR_INVALID_ARG;
    };
    let h = unsafe { &*handle };
    match h.writer.save(&path) {
        Ok(()) => {
            set_err(error_code, OFFICE_OK);
            OFFICE_OK
        },
        Err(_) => {
            set_err(error_code, OFFICE_ERR_INTERNAL);
            OFFICE_ERR_INTERNAL
        },
    }
}

/// Serialize to a heap byte buffer. Free with office_oxide_free_bytes.
#[unsafe(no_mangle)]
pub extern "C" fn office_xlsx_writer_to_bytes(
    handle: *const OfficeXlsxWriterHandle,
    out_len: *mut usize,
    error_code: *mut i32,
) -> *mut u8 {
    if handle.is_null() || out_len.is_null() {
        set_err(error_code, OFFICE_ERR_INVALID_ARG);
        return ptr::null_mut();
    }
    let h = unsafe { &*handle };
    let mut cursor = std::io::Cursor::new(Vec::new());
    match h.writer.write_to(&mut cursor) {
        Ok(()) => {
            let mut bytes = cursor.into_inner();
            bytes.shrink_to_fit();
            let len = bytes.len();
            let raw_ptr = bytes.as_mut_ptr();
            std::mem::forget(bytes);
            unsafe { *out_len = len };
            set_err(error_code, OFFICE_OK);
            raw_ptr
        },
        Err(_) => {
            set_err(error_code, OFFICE_ERR_INTERNAL);
            ptr::null_mut()
        },
    }
}

// ─── PptxWriter ──────────────────────────────────────────────────────────────

/// Opaque handle wrapping a PptxWriter.
pub struct OfficePptxWriterHandle {
    writer: crate::pptx::write::PptxWriter,
}

/// Create a new PPTX writer. Free with `office_pptx_writer_free`.
#[unsafe(no_mangle)]
pub extern "C" fn office_pptx_writer_new() -> *mut OfficePptxWriterHandle {
    Box::into_raw(Box::new(OfficePptxWriterHandle {
        writer: crate::pptx::write::PptxWriter::new(),
    }))
}

/// Free a PPTX writer handle.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn office_pptx_writer_free(handle: *mut OfficePptxWriterHandle) {
    if !handle.is_null() {
        drop(unsafe { Box::from_raw(handle) });
    }
}

/// Override presentation canvas size. 914400 EMU = 1 inch.
#[unsafe(no_mangle)]
pub extern "C" fn office_pptx_writer_set_presentation_size(
    handle: *mut OfficePptxWriterHandle,
    cx: u64,
    cy: u64,
) {
    if handle.is_null() {
        return;
    }
    let h = unsafe { &mut *handle };
    h.writer.set_presentation_size(cx, cy);
}

/// Add a slide; returns its 0-based index, or u32::MAX on null handle.
#[unsafe(no_mangle)]
pub extern "C" fn office_pptx_writer_add_slide(handle: *mut OfficePptxWriterHandle) -> u32 {
    if handle.is_null() {
        return u32::MAX;
    }
    let h = unsafe { &mut *handle };
    h.writer.add_slide_get_index() as u32
}

/// Set the slide title.
#[unsafe(no_mangle)]
pub extern "C" fn office_pptx_slide_set_title(
    handle: *mut OfficePptxWriterHandle,
    slide: u32,
    title: *const c_char,
) {
    if handle.is_null() {
        return;
    }
    let title = cstr_to_str(title).unwrap_or("");
    let h = unsafe { &mut *handle };
    h.writer.slide_set_title(slide as usize, title);
}

/// Add a plain text paragraph to the slide body.
#[unsafe(no_mangle)]
pub extern "C" fn office_pptx_slide_add_text(
    handle: *mut OfficePptxWriterHandle,
    slide: u32,
    text: *const c_char,
) {
    if handle.is_null() {
        return;
    }
    let text = cstr_to_str(text).unwrap_or("");
    let h = unsafe { &mut *handle };
    h.writer.slide_add_text(slide as usize, text);
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
) {
    if handle.is_null() || data.is_null() || len == 0 {
        return;
    }
    let Some(fmt_str) = cstr_to_str(format) else {
        return;
    };
    let Some(fmt) = parse_image_format(fmt_str) else {
        return;
    };
    let bytes = unsafe { slice::from_raw_parts(data, len) }.to_vec();
    let h = unsafe { &mut *handle };
    h.writer
        .slide_add_image(slide as usize, bytes, fmt, x, y, cx, cy);
}

/// Save to a file. Returns OFFICE_OK (0) on success.
#[unsafe(no_mangle)]
pub extern "C" fn office_pptx_writer_save(
    handle: *const OfficePptxWriterHandle,
    path: *const c_char,
    error_code: *mut i32,
) -> i32 {
    if handle.is_null() {
        set_err(error_code, OFFICE_ERR_INVALID_ARG);
        return OFFICE_ERR_INVALID_ARG;
    }
    let Some(path) = cstr_to_pathbuf(path) else {
        set_err(error_code, OFFICE_ERR_INVALID_ARG);
        return OFFICE_ERR_INVALID_ARG;
    };
    let h = unsafe { &*handle };
    match h.writer.save(&path) {
        Ok(()) => {
            set_err(error_code, OFFICE_OK);
            OFFICE_OK
        },
        Err(_) => {
            set_err(error_code, OFFICE_ERR_INTERNAL);
            OFFICE_ERR_INTERNAL
        },
    }
}

/// Serialize to a heap byte buffer. Free with office_oxide_free_bytes.
#[unsafe(no_mangle)]
pub extern "C" fn office_pptx_writer_to_bytes(
    handle: *const OfficePptxWriterHandle,
    out_len: *mut usize,
    error_code: *mut i32,
) -> *mut u8 {
    if handle.is_null() || out_len.is_null() {
        set_err(error_code, OFFICE_ERR_INVALID_ARG);
        return ptr::null_mut();
    }
    let h = unsafe { &*handle };
    let mut cursor = std::io::Cursor::new(Vec::new());
    match h.writer.write_to(&mut cursor) {
        Ok(()) => {
            let mut bytes = cursor.into_inner();
            bytes.shrink_to_fit();
            let len = bytes.len();
            let raw_ptr = bytes.as_mut_ptr();
            std::mem::forget(bytes);
            unsafe { *out_len = len };
            set_err(error_code, OFFICE_OK);
            raw_ptr
        },
        Err(_) => {
            set_err(error_code, OFFICE_ERR_INTERNAL);
            ptr::null_mut()
        },
    }
}

/// One-shot: convert a Markdown string to an Office document at `path`.
///
/// `format` must be one of `"docx"`, `"xlsx"`, or `"pptx"` (case-insensitive).
/// Returns `OFFICE_OK` (0) on success, a negative error code on failure.
#[unsafe(no_mangle)]
pub extern "C" fn office_create_from_markdown(
    markdown: *const c_char,
    format: *const c_char,
    path: *const c_char,
    error_code: *mut i32,
) -> i32 {
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
    match crate::create::create_from_markdown(md, doc_format, &out_path) {
        Ok(()) => {
            set_err(error_code, OFFICE_OK);
            OFFICE_OK
        },
        Err(e) => {
            let code = classify_error(&e);
            set_err(error_code, code);
            code
        },
    }
}
