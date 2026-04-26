// SPDX-License-Identifier: MIT OR Apache-2.0
// Package officeoxide provides idiomatic Go bindings for the office_oxide
// Rust library, which parses, converts, and edits Microsoft Office documents
// (DOCX, XLSX, PPTX, DOC, XLS, PPT).
//
// The bindings link against the C FFI layer exposed by the Rust crate, so
// the final Go binary is self-contained (no runtime library lookups).
//
// # Build configuration
//
// You must tell cgo where to find the office_oxide headers and static
// library. Two supported modes:
//
//  1. Monorepo development — build with the `office_oxide_dev` build tag
//     after running `cargo build --release --lib` in the workspace root.
//     The `cgo_dev.go` file points cgo at `target/release`.
//
//  2. Downstream consumers — set CGO_CFLAGS and CGO_LDFLAGS to point at an
//     installed prefix that contains the header and library, e.g.:
//
//        export CGO_CFLAGS="-I/usr/local/include"
//        export CGO_LDFLAGS="-L/usr/local/lib -loffice_oxide"
//
// # Example
//
//	doc, err := officeoxide.Open("report.docx")
//	if err != nil { log.Fatal(err) }
//	defer doc.Close()
//	fmt.Println(doc.PlainText())
package officeoxide

/*
#include <stdlib.h>
#include <stdint.h>
#include <stdbool.h>
#include <stddef.h>
#include "office_oxide.h"
*/
import "C"

import (
	"errors"
	"fmt"
	"runtime"
	"unsafe"
)

// ─── Error handling ─────────────────────────────────────────────────────────

// Error wraps an office_oxide FFI error code with the originating operation.
type Error struct {
	Code int
	Op   string
}

func (e *Error) Error() string {
	var kind string
	switch e.Code {
	case int(C.OFFICE_OK):
		kind = "ok"
	case int(C.OFFICE_ERR_INVALID_ARG):
		kind = "invalid argument"
	case int(C.OFFICE_ERR_IO):
		kind = "io error"
	case int(C.OFFICE_ERR_PARSE):
		kind = "parse error"
	case int(C.OFFICE_ERR_EXTRACTION):
		kind = "extraction failed"
	case int(C.OFFICE_ERR_INTERNAL):
		kind = "internal error"
	case int(C.OFFICE_ERR_UNSUPPORTED):
		kind = "unsupported format"
	default:
		kind = fmt.Sprintf("code=%d", e.Code)
	}
	return fmt.Sprintf("office_oxide: %s: %s", e.Op, kind)
}

// ErrClosed is returned when using a handle that has already been closed.
var ErrClosed = errors.New("office_oxide: handle is closed")

// ─── Library info ───────────────────────────────────────────────────────────

// Version returns the underlying office_oxide library version.
func Version() string {
	return C.GoString(C.office_oxide_version())
}

// DetectFormat returns the detected format ("docx", "xlsx", "pptx", "doc",
// "xls", "ppt") for the given file path, or an empty string if unsupported.
func DetectFormat(path string) string {
	cpath := C.CString(path)
	defer C.free(unsafe.Pointer(cpath))
	out := C.office_oxide_detect_format(cpath)
	if out == nil {
		return ""
	}
	return C.GoString(out)
}

// ─── Document (read-only) ───────────────────────────────────────────────────

// Document is a read-only Office document handle.
type Document struct {
	handle *C.OfficeDocumentHandle
}

// Open loads an Office document from a file path. Format is detected from
// the extension and corrected via magic-byte sniffing. Remember to call
// Close (or use defer doc.Close()) when done.
func Open(path string) (*Document, error) {
	cpath := C.CString(path)
	defer C.free(unsafe.Pointer(cpath))
	var errCode C.int
	h := C.office_document_open(cpath, &errCode)
	if h == nil {
		return nil, &Error{Code: int(errCode), Op: "Open"}
	}
	d := &Document{handle: h}
	runtime.SetFinalizer(d, func(d *Document) { d.Close() })
	return d, nil
}

// OpenFromBytes loads a document from an in-memory buffer. `format` must be
// one of "docx", "xlsx", "pptx", "doc", "xls", "ppt".
func OpenFromBytes(data []byte, format string) (*Document, error) {
	if len(data) == 0 {
		return nil, &Error{Code: int(C.OFFICE_ERR_INVALID_ARG), Op: "OpenFromBytes"}
	}
	cfmt := C.CString(format)
	defer C.free(unsafe.Pointer(cfmt))
	var errCode C.int
	h := C.office_document_open_from_bytes(
		(*C.uint8_t)(unsafe.Pointer(&data[0])),
		C.size_t(len(data)),
		cfmt,
		&errCode,
	)
	if h == nil {
		return nil, &Error{Code: int(errCode), Op: "OpenFromBytes"}
	}
	d := &Document{handle: h}
	runtime.SetFinalizer(d, func(d *Document) { d.Close() })
	return d, nil
}

// Close releases the underlying document handle. Safe to call multiple times.
func (d *Document) Close() error {
	if d == nil || d.handle == nil {
		return nil
	}
	C.office_document_free(d.handle)
	d.handle = nil
	runtime.SetFinalizer(d, nil)
	return nil
}

// Format returns the detected document format ("docx", "xlsx", ...).
func (d *Document) Format() (string, error) {
	if d.handle == nil {
		return "", ErrClosed
	}
	cs := C.office_document_format(d.handle)
	if cs == nil {
		return "", &Error{Code: int(C.OFFICE_ERR_INVALID_ARG), Op: "Format"}
	}
	return C.GoString(cs), nil
}

func cStrOrErr(out *C.char, errCode C.int, op string) (string, error) {
	if out == nil {
		return "", &Error{Code: int(errCode), Op: op}
	}
	defer C.office_oxide_free_string(out)
	return C.GoString(out), nil
}

// PlainText extracts plain text from the document.
func (d *Document) PlainText() (string, error) {
	if d.handle == nil {
		return "", ErrClosed
	}
	var errCode C.int
	return cStrOrErr(C.office_document_plain_text(d.handle, &errCode), errCode, "PlainText")
}

// ToMarkdown converts the document to Markdown.
func (d *Document) ToMarkdown() (string, error) {
	if d.handle == nil {
		return "", ErrClosed
	}
	var errCode C.int
	return cStrOrErr(C.office_document_to_markdown(d.handle, &errCode), errCode, "ToMarkdown")
}

// ToHTML converts the document to an HTML fragment.
func (d *Document) ToHTML() (string, error) {
	if d.handle == nil {
		return "", ErrClosed
	}
	var errCode C.int
	return cStrOrErr(C.office_document_to_html(d.handle, &errCode), errCode, "ToHTML")
}

// ToIRJSON returns the format-agnostic document IR serialised as JSON.
// Use encoding/json to unmarshal it.
func (d *Document) ToIRJSON() (string, error) {
	if d.handle == nil {
		return "", ErrClosed
	}
	var errCode C.int
	return cStrOrErr(C.office_document_to_ir_json(d.handle, &errCode), errCode, "ToIRJSON")
}

// SaveAs writes the document to `path`. Target format is inferred from the
// extension. Legacy formats (.doc/.xls/.ppt) are converted to OOXML.
func (d *Document) SaveAs(path string) error {
	if d.handle == nil {
		return ErrClosed
	}
	cpath := C.CString(path)
	defer C.free(unsafe.Pointer(cpath))
	var errCode C.int
	rc := C.office_document_save_as(d.handle, cpath, &errCode)
	if rc != 0 {
		return &Error{Code: int(errCode), Op: "SaveAs"}
	}
	return nil
}

// ─── EditableDocument ──────────────────────────────────────────────────────

// EditableDocument is a DOCX / XLSX / PPTX document opened for editing.
type EditableDocument struct {
	handle *C.OfficeEditableHandle
}

// OpenEditable opens a document for editing.
func OpenEditable(path string) (*EditableDocument, error) {
	cpath := C.CString(path)
	defer C.free(unsafe.Pointer(cpath))
	var errCode C.int
	h := C.office_editable_open(cpath, &errCode)
	if h == nil {
		return nil, &Error{Code: int(errCode), Op: "OpenEditable"}
	}
	ed := &EditableDocument{handle: h}
	runtime.SetFinalizer(ed, func(ed *EditableDocument) { ed.Close() })
	return ed, nil
}

// Close releases the editable document handle.
func (ed *EditableDocument) Close() error {
	if ed == nil || ed.handle == nil {
		return nil
	}
	C.office_editable_free(ed.handle)
	ed.handle = nil
	runtime.SetFinalizer(ed, nil)
	return nil
}

// ReplaceText replaces every occurrence of `find` with `replace` in text
// content. Returns the number of replacements.
func (ed *EditableDocument) ReplaceText(find, replace string) (int64, error) {
	if ed.handle == nil {
		return 0, ErrClosed
	}
	cfind := C.CString(find)
	crepl := C.CString(replace)
	defer C.free(unsafe.Pointer(cfind))
	defer C.free(unsafe.Pointer(crepl))
	var errCode C.int
	n := C.office_editable_replace_text(ed.handle, cfind, crepl, &errCode)
	if n < 0 {
		return 0, &Error{Code: int(errCode), Op: "ReplaceText"}
	}
	return int64(n), nil
}

// CellValue is the value written by SetCell. Use one of NewCell* constructors.
type CellValue struct {
	kind int
	str  string
	num  float64
}

// NewEmptyCell returns an empty-cell value.
func NewEmptyCell() CellValue { return CellValue{kind: int(C.OFFICE_CELL_EMPTY)} }

// NewStringCell wraps a string value.
func NewStringCell(s string) CellValue {
	return CellValue{kind: int(C.OFFICE_CELL_STRING), str: s}
}

// NewNumberCell wraps a numeric value.
func NewNumberCell(v float64) CellValue {
	return CellValue{kind: int(C.OFFICE_CELL_NUMBER), num: v}
}

// NewBoolCell wraps a boolean value.
func NewBoolCell(b bool) CellValue {
	v := 0.0
	if b {
		v = 1.0
	}
	return CellValue{kind: int(C.OFFICE_CELL_BOOLEAN), num: v}
}

// SetCell sets a cell value in an XLSX document. `cellRef` is a spreadsheet
// reference like "A1" or "C12".
func (ed *EditableDocument) SetCell(sheetIndex int, cellRef string, value CellValue) error {
	if ed.handle == nil {
		return ErrClosed
	}
	cref := C.CString(cellRef)
	defer C.free(unsafe.Pointer(cref))
	var cstr *C.char
	if value.kind == int(C.OFFICE_CELL_STRING) {
		cstr = C.CString(value.str)
		defer C.free(unsafe.Pointer(cstr))
	}
	var errCode C.int
	rc := C.office_editable_set_cell(
		ed.handle,
		C.uint32_t(sheetIndex),
		cref,
		C.int32_t(value.kind),
		cstr,
		C.double(value.num),
		&errCode,
	)
	if rc != 0 {
		return &Error{Code: int(errCode), Op: "SetCell"}
	}
	return nil
}

// Save writes the edited document to `path`.
func (ed *EditableDocument) Save(path string) error {
	if ed.handle == nil {
		return ErrClosed
	}
	cpath := C.CString(path)
	defer C.free(unsafe.Pointer(cpath))
	var errCode C.int
	rc := C.office_editable_save(ed.handle, cpath, &errCode)
	if rc != 0 {
		return &Error{Code: int(errCode), Op: "Save"}
	}
	return nil
}

// SaveToBytes serialises the edited document to a new byte slice.
func (ed *EditableDocument) SaveToBytes() ([]byte, error) {
	if ed.handle == nil {
		return nil, ErrClosed
	}
	var outLen C.size_t
	var errCode C.int
	ptr := C.office_editable_save_to_bytes(ed.handle, &outLen, &errCode)
	if ptr == nil {
		return nil, &Error{Code: int(errCode), Op: "SaveToBytes"}
	}
	defer C.office_oxide_free_bytes(ptr, outLen)
	return C.GoBytes(unsafe.Pointer(ptr), C.int(outLen)), nil
}

// ─── One-shot helpers ───────────────────────────────────────────────────────

// ExtractText opens a file, returns its plain text, and closes it.
func ExtractText(path string) (string, error) {
	cpath := C.CString(path)
	defer C.free(unsafe.Pointer(cpath))
	var errCode C.int
	out := C.office_extract_text(cpath, &errCode)
	if out == nil {
		return "", &Error{Code: int(errCode), Op: "ExtractText"}
	}
	defer C.office_oxide_free_string(out)
	return C.GoString(out), nil
}

// ToMarkdown opens a file, converts it to markdown, and closes it.
func ToMarkdown(path string) (string, error) {
	cpath := C.CString(path)
	defer C.free(unsafe.Pointer(cpath))
	var errCode C.int
	out := C.office_to_markdown(cpath, &errCode)
	if out == nil {
		return "", &Error{Code: int(errCode), Op: "ToMarkdown"}
	}
	defer C.office_oxide_free_string(out)
	return C.GoString(out), nil
}

// ToHTML opens a file, converts it to HTML, and closes it.
func ToHTML(path string) (string, error) {
	cpath := C.CString(path)
	defer C.free(unsafe.Pointer(cpath))
	var errCode C.int
	out := C.office_to_html(cpath, &errCode)
	if out == nil {
		return "", &Error{Code: int(errCode), Op: "ToHTML"}
	}
	defer C.office_oxide_free_string(out)
	return C.GoString(out), nil
}

// CreateFromMarkdown converts a Markdown string into an Office document file.
//
// format must be "docx", "xlsx", or "pptx" (case-insensitive).
func CreateFromMarkdown(markdown, format, path string) error {
	cmd := C.CString(markdown)
	defer C.free(unsafe.Pointer(cmd))
	cfmt := C.CString(format)
	defer C.free(unsafe.Pointer(cfmt))
	cpath := C.CString(path)
	defer C.free(unsafe.Pointer(cpath))
	var errCode C.int
	rc := C.office_create_from_markdown(cmd, cfmt, cpath, &errCode)
	if rc != 0 {
		return &Error{Code: int(errCode), Op: "CreateFromMarkdown"}
	}
	return nil
}
