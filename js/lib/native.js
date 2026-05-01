// Native-library loader + FFI bindings for office_oxide.
//
// Uses koffi's `disposable()` pattern so that C strings allocated by the
// library are auto-decoded to JS strings *and* freed via office_oxide_free_string.

import koffi from 'koffi';
import { createRequire } from 'node:module';
import path from 'node:path';
import process from 'node:process';

const require = createRequire(import.meta.url);

const ext =
  process.platform === 'win32' ? '.dll' :
  process.platform === 'darwin' ? '.dylib' : '.so';
const prefix = process.platform === 'win32' ? '' : 'lib';

function candidatePaths() {
  const paths = [];
  if (process.env.OFFICE_OXIDE_LIB) paths.push(process.env.OFFICE_OXIDE_LIB);
  const hereDir = path.dirname(require.resolve('../package.json'));
  paths.push(path.join(
    hereDir, 'prebuilds',
    `${process.platform}-${process.arch}`,
    `${prefix}office_oxide${ext}`,
  ));
  paths.push(`${prefix}office_oxide${ext}`);
  return paths;
}

function loadLib() {
  let lastErr;
  for (const p of candidatePaths()) {
    try { return koffi.load(p); }
    catch (e) { lastErr = e; }
  }
  throw new Error(
    `office-oxide: failed to load native library. Tried:\n  ${candidatePaths().join('\n  ')}\n` +
    `Original error: ${lastErr?.message ?? 'unknown'}`,
  );
}

const lib = loadLib();

// Declare the raw string-freeing function first so we can compose it into a
// disposable type. Heap strings are returned as `HeapStr` — koffi decodes
// them to JS strings and automatically calls office_oxide_free_string.
const freeStringRaw = lib.func('void office_oxide_free_string(void* ptr)');
const freeBytesRaw = lib.func('void office_oxide_free_bytes(void* ptr, size_t len)');

// Register a disposable `HeapStr` type. koffi keeps the type table global,
// so the type name (registered here) is usable by name in function prototypes.
koffi.disposable('HeapStr', 'char*', freeStringRaw);

export const native = {
  version: lib.func('const char* office_oxide_version()'),
  detectFormat: lib.func('const char* office_oxide_detect_format(const char* path)'),
  freeString: freeStringRaw,
  freeBytes: freeBytesRaw,

  documentOpen: lib.func('void* office_document_open(const char* path, _Out_ int* error_code)'),
  documentOpenFromBytes: lib.func('void* office_document_open_from_bytes(const uint8_t* data, size_t len, const char* format, _Out_ int* error_code)'),
  documentFree: lib.func('void office_document_free(void* handle)'),
  documentFormat: lib.func('const char* office_document_format(void* handle)'),
  documentPlainText: lib.func('HeapStr office_document_plain_text(void* handle, _Out_ int* error_code)'),
  documentToMarkdown: lib.func('HeapStr office_document_to_markdown(void* handle, _Out_ int* error_code)'),
  documentToHtml: lib.func('HeapStr office_document_to_html(void* handle, _Out_ int* error_code)'),
  documentToIrJson: lib.func('HeapStr office_document_to_ir_json(void* handle, _Out_ int* error_code)'),
  documentSaveAs: lib.func('int32_t office_document_save_as(void* handle, const char* path, _Out_ int* error_code)'),

  editableOpen: lib.func('void* office_editable_open(const char* path, _Out_ int* error_code)'),
  editableFree: lib.func('void office_editable_free(void* handle)'),
  editableReplaceText: lib.func('int64_t office_editable_replace_text(void* handle, const char* find, const char* replace, _Out_ int* error_code)'),
  editableSetCell: lib.func('int32_t office_editable_set_cell(void* handle, uint32_t sheet_index, const char* cell_ref, int32_t value_type, const char* value_str, double value_num, _Out_ int* error_code)'),
  editableSave: lib.func('int32_t office_editable_save(void* handle, const char* path, _Out_ int* error_code)'),

  extractText: lib.func('HeapStr office_extract_text(const char* path, _Out_ int* error_code)'),
  toMarkdown: lib.func('HeapStr office_to_markdown(const char* path, _Out_ int* error_code)'),
  toHtml: lib.func('HeapStr office_to_html(const char* path, _Out_ int* error_code)'),
  createFromMarkdown: lib.func('int32_t office_create_from_markdown(const char* markdown, const char* format, const char* path, _Out_ int* error_code)'),

  // XLSX writer
  xlsxWriterNew: lib.func('void* office_xlsx_writer_new()'),
  xlsxWriterFree: lib.func('void office_xlsx_writer_free(void* handle)'),
  xlsxWriterAddSheet: lib.func('uint32_t office_xlsx_writer_add_sheet(void* handle, const char* name)'),
  xlsxSheetSetCell: lib.func('void office_xlsx_sheet_set_cell(void* handle, uint32_t sheet, uint32_t row, uint32_t col, int32_t value_type, const char* value_str, double value_num)'),
  xlsxSheetSetCellStyled: lib.func('void office_xlsx_sheet_set_cell_styled(void* handle, uint32_t sheet, uint32_t row, uint32_t col, int32_t value_type, const char* value_str, double value_num, bool bold, const char* bg_color)'),
  xlsxSheetMergeCells: lib.func('void office_xlsx_sheet_merge_cells(void* handle, uint32_t sheet, uint32_t row, uint32_t col, uint32_t row_span, uint32_t col_span)'),
  xlsxSheetSetColumnWidth: lib.func('void office_xlsx_sheet_set_column_width(void* handle, uint32_t sheet, uint32_t col, double width)'),
  xlsxWriterSave: lib.func('int32_t office_xlsx_writer_save(void* handle, const char* path, _Out_ int* error_code)'),
  xlsxWriterToBytes: lib.func('uint8_t* office_xlsx_writer_to_bytes(void* handle, _Out_ size_t* out_len, _Out_ int* error_code)'),

  // PPTX writer
  pptxWriterNew: lib.func('void* office_pptx_writer_new()'),
  pptxWriterFree: lib.func('void office_pptx_writer_free(void* handle)'),
  pptxWriterSetPresentationSize: lib.func('void office_pptx_writer_set_presentation_size(void* handle, uint64_t cx, uint64_t cy)'),
  pptxWriterAddSlide: lib.func('uint32_t office_pptx_writer_add_slide(void* handle)'),
  pptxSlideSetTitle: lib.func('void office_pptx_slide_set_title(void* handle, uint32_t slide, const char* title)'),
  pptxSlideAddText: lib.func('void office_pptx_slide_add_text(void* handle, uint32_t slide, const char* text)'),
  pptxSlideAddImage: lib.func('void office_pptx_slide_add_image(void* handle, uint32_t slide, const uint8_t* data, size_t len, const char* format, int64_t x, int64_t y, uint64_t cx, uint64_t cy)'),
  pptxWriterSave: lib.func('int32_t office_pptx_writer_save(void* handle, const char* path, _Out_ int* error_code)'),
  pptxWriterToBytes: lib.func('uint8_t* office_pptx_writer_to_bytes(void* handle, _Out_ size_t* out_len, _Out_ int* error_code)'),
};
