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

#ifdef __cplusplus
}  /* extern "C" */
#endif

#endif  /* OFFICE_OXIDE_H */
