/**
 * office_oxide C API
 *
 * C-compatible Foreign Function Interface for office_oxide.
 * Consumed by Go (CGo), Node.js (N-API), C# (P/Invoke) bindings.
 *
 * Error Convention:
 *   Most functions accept an `int* error_code` out-parameter.
 *     0 = success
 *     1 = invalid argument / path
 *     2 = IO error
 *     3 = parse error
 *     4 = extraction failed
 *     5 = internal error
 *     6 = unsupported format / feature
 *
 * Memory Convention:
 *   - Strings returned as `char*` must be freed with office_oxide_free_string().
 *   - Byte buffers returned as `uint8_t*` (with an `out_len`) must be freed with
 *     office_oxide_free_bytes(ptr, len).
 *   - Opaque handles must be freed with their corresponding *_free() function.
 *   - Static C strings returned as `const char*` (e.g., version, format names)
 *     are NOT to be freed.
 */

#ifndef OFFICE_OXIDE_H
#define OFFICE_OXIDE_H

#include <stdint.h>
#include <stdbool.h>
#include <stddef.h>

#ifdef __cplusplus
extern "C" {
#endif

/* ─── Error codes ─────────────────────────────────────────────────────────── */

#define OFFICE_OK                   0
#define OFFICE_ERR_INVALID_ARG      1
#define OFFICE_ERR_IO               2
#define OFFICE_ERR_PARSE            3
#define OFFICE_ERR_EXTRACTION       4
#define OFFICE_ERR_INTERNAL         5
#define OFFICE_ERR_UNSUPPORTED      6

/* ─── Cell value types (for office_editable_set_cell) ─────────────────────── */

#define OFFICE_CELL_EMPTY           0
#define OFFICE_CELL_STRING          1
#define OFFICE_CELL_NUMBER          2
#define OFFICE_CELL_BOOLEAN         3

/* ─── Opaque handle types ─────────────────────────────────────────────────── */

typedef struct OfficeDocumentHandle OfficeDocumentHandle;
typedef struct OfficeEditableHandle OfficeEditableHandle;

/* ─── Library info / memory ───────────────────────────────────────────────── */

/** Return library version (e.g. "0.1.0"). Do not free. */
const char* office_oxide_version(void);

/** Free a string returned by any FFI function. */
void office_oxide_free_string(char* ptr);

/** Free a byte buffer. `len` must match the out_len returned with the pointer. */
void office_oxide_free_bytes(uint8_t* ptr, size_t len);

/**
 * Detect document format from a file path. Returns a static C string
 * ("docx", "xlsx", "pptx", "doc", "xls", "ppt") or NULL if unsupported.
 * Do not free.
 */
const char* office_oxide_detect_format(const char* path);

/* ─── Document (read-only) ────────────────────────────────────────────────── */

/** Open a document from a file path. Format detected from extension. */
OfficeDocumentHandle* office_document_open(const char* path, int* error_code);

/**
 * Open a document from an in-memory buffer.
 * `format` is one of: "docx", "xlsx", "pptx", "doc", "xls", "ppt".
 */
OfficeDocumentHandle* office_document_open_from_bytes(
    const uint8_t* data,
    size_t len,
    const char* format,
    int* error_code);

/** Free a document handle. */
void office_document_free(OfficeDocumentHandle* handle);

/** Return the document format as a static C string. Do not free. */
const char* office_document_format(const OfficeDocumentHandle* handle);

/** Extract plain text. Free result with office_oxide_free_string. */
char* office_document_plain_text(const OfficeDocumentHandle* handle, int* error_code);

/** Convert to Markdown. Free result with office_oxide_free_string. */
char* office_document_to_markdown(const OfficeDocumentHandle* handle, int* error_code);

/** Convert to HTML fragment. Free result with office_oxide_free_string. */
char* office_document_to_html(const OfficeDocumentHandle* handle, int* error_code);

/** Convert to the document IR as JSON. Free result with office_oxide_free_string. */
char* office_document_to_ir_json(const OfficeDocumentHandle* handle, int* error_code);

/**
 * Save/convert to a file. Target format detected from extension.
 * Returns 0 on success, nonzero error code otherwise.
 */
int32_t office_document_save_as(
    const OfficeDocumentHandle* handle,
    const char* path,
    int* error_code);

/* ─── EditableDocument ────────────────────────────────────────────────────── */

/** Open a document for editing (DOCX, XLSX, PPTX). */
OfficeEditableHandle* office_editable_open(const char* path, int* error_code);

/** Open an editable document from bytes. `format`: "docx"|"xlsx"|"pptx". */
OfficeEditableHandle* office_editable_open_from_bytes(
    const uint8_t* data,
    size_t len,
    const char* format,
    int* error_code);

/** Free an editable document handle. */
void office_editable_free(OfficeEditableHandle* handle);

/**
 * Replace every occurrence of `find` with `replace` in text content.
 * Returns the number of replacements, or -1 on error.
 */
int64_t office_editable_replace_text(
    OfficeEditableHandle* handle,
    const char* find,
    const char* replace,
    int* error_code);

/**
 * Set a cell value in an XLSX document.
 *
 * `value_type` is one of OFFICE_CELL_EMPTY / STRING / NUMBER / BOOLEAN.
 * `value_str` is used when value_type == OFFICE_CELL_STRING; pass NULL otherwise.
 * `value_num` carries the number (NUMBER) or boolean (BOOLEAN: nonzero = true).
 */
int32_t office_editable_set_cell(
    OfficeEditableHandle* handle,
    uint32_t sheet_index,
    const char* cell_ref,
    int32_t value_type,
    const char* value_str,
    double value_num,
    int* error_code);

/** Save the edited document to a file. */
int32_t office_editable_save(
    const OfficeEditableHandle* handle,
    const char* path,
    int* error_code);

/**
 * Serialize the edited document to a heap byte buffer.
 * Writes length to *out_len. Free with office_oxide_free_bytes(ptr, len).
 */
uint8_t* office_editable_save_to_bytes(
    const OfficeEditableHandle* handle,
    size_t* out_len,
    int* error_code);

/* ─── One-shot convenience helpers ────────────────────────────────────────── */

/** Open + extract plain text from a file. Free with office_oxide_free_string. */
char* office_extract_text(const char* path, int* error_code);

/** Open + convert to markdown. Free with office_oxide_free_string. */
char* office_to_markdown(const char* path, int* error_code);

/** Open + convert to HTML. Free with office_oxide_free_string. */
char* office_to_html(const char* path, int* error_code);

/**
 * Convert a Markdown string to an Office document file.
 * format must be "docx", "xlsx", or "pptx" (case-insensitive).
 * Returns OFFICE_OK (0) on success, a negative error code on failure.
 */
int32_t office_create_from_markdown(
    const char* markdown,
    const char* format,
    const char* path,
    int* error_code);

/* ─── XlsxWriter ──────────────────────────────────────────────────────────── */

typedef struct OfficeXlsxWriterHandle OfficeXlsxWriterHandle;

/** Create a new XLSX writer. Free with office_xlsx_writer_free. */
OfficeXlsxWriterHandle* office_xlsx_writer_new(void);

/** Free an XLSX writer handle. */
void office_xlsx_writer_free(OfficeXlsxWriterHandle* handle);

/**
 * Add a sheet; returns its 0-based index.
 * Returns UINT32_MAX on null handle.
 */
uint32_t office_xlsx_writer_add_sheet(OfficeXlsxWriterHandle* handle, const char* name);

/**
 * Set a cell value.
 * value_type: OFFICE_CELL_EMPTY=0, OFFICE_CELL_STRING=1, OFFICE_CELL_NUMBER=2.
 * value_str used when value_type==1; value_num used when value_type==2.
 */
void office_xlsx_sheet_set_cell(
    OfficeXlsxWriterHandle* handle,
    uint32_t sheet, uint32_t row, uint32_t col,
    int32_t value_type, const char* value_str, double value_num);

/**
 * Set a cell with styling.
 * bold: apply bold weight.
 * bg_color: 6-char hex string ("D3D3D3") or NULL for no fill.
 */
void office_xlsx_sheet_set_cell_styled(
    OfficeXlsxWriterHandle* handle,
    uint32_t sheet, uint32_t row, uint32_t col,
    int32_t value_type, const char* value_str, double value_num,
    bool bold, const char* bg_color);

/** Merge a rectangular range of cells. row_span and col_span must be >= 1. */
void office_xlsx_sheet_merge_cells(
    OfficeXlsxWriterHandle* handle,
    uint32_t sheet, uint32_t row, uint32_t col,
    uint32_t row_span, uint32_t col_span);

/** Set column width in Excel character units (e.g. 20.0). */
void office_xlsx_sheet_set_column_width(
    OfficeXlsxWriterHandle* handle,
    uint32_t sheet, uint32_t col, double width);

/** Save the workbook to a file. Returns OFFICE_OK on success. */
int32_t office_xlsx_writer_save(
    const OfficeXlsxWriterHandle* handle,
    const char* path, int* error_code);

/**
 * Serialize to a heap byte buffer.
 * Writes length to *out_len. Free with office_oxide_free_bytes(ptr, len).
 */
uint8_t* office_xlsx_writer_to_bytes(
    const OfficeXlsxWriterHandle* handle,
    size_t* out_len, int* error_code);

/* ─── PptxWriter ──────────────────────────────────────────────────────────── */

typedef struct OfficePptxWriterHandle OfficePptxWriterHandle;

/** Create a new PPTX writer. Free with office_pptx_writer_free. */
OfficePptxWriterHandle* office_pptx_writer_new(void);

/** Free a PPTX writer handle. */
void office_pptx_writer_free(OfficePptxWriterHandle* handle);

/**
 * Override the presentation canvas size.
 * 914400 EMU = 1 inch. Default: 12192000 x 6858000 (16:9).
 */
void office_pptx_writer_set_presentation_size(
    OfficePptxWriterHandle* handle,
    uint64_t cx, uint64_t cy);

/**
 * Add a slide; returns its 0-based index.
 * Returns UINT32_MAX on null handle.
 */
uint32_t office_pptx_writer_add_slide(OfficePptxWriterHandle* handle);

/** Set the slide title. */
void office_pptx_slide_set_title(
    OfficePptxWriterHandle* handle,
    uint32_t slide, const char* title);

/** Add a plain text paragraph to the slide body. */
void office_pptx_slide_add_text(
    OfficePptxWriterHandle* handle,
    uint32_t slide, const char* text);

/**
 * Embed an image on a slide.
 * data/len: raw PNG, JPEG, or GIF bytes.
 * format: "png", "jpeg"/"jpg", or "gif".
 * x, y, cx, cy: position and size in EMU (914400 = 1 inch).
 */
void office_pptx_slide_add_image(
    OfficePptxWriterHandle* handle,
    uint32_t slide,
    const uint8_t* data, size_t len,
    const char* format,
    int64_t x, int64_t y,
    uint64_t cx, uint64_t cy);

/** Save the presentation to a file. Returns OFFICE_OK on success. */
int32_t office_pptx_writer_save(
    const OfficePptxWriterHandle* handle,
    const char* path, int* error_code);

/**
 * Serialize to a heap byte buffer.
 * Writes length to *out_len. Free with office_oxide_free_bytes(ptr, len).
 */
uint8_t* office_pptx_writer_to_bytes(
    const OfficePptxWriterHandle* handle,
    size_t* out_len, int* error_code);

#ifdef __cplusplus
}  /* extern "C" */
#endif

#endif  /* OFFICE_OXIDE_H */
