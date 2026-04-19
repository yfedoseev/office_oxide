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
};
