// CommonJS entry point — mirrors lib/index.js (ESM) exactly.
'use strict';

const koffi = require('koffi');
const path = require('node:path');
const process = require('node:process');

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
    `Original error: ${lastErr ? lastErr.message : 'unknown'}`,
  );
}

const lib = loadLib();
const freeStringRaw = lib.func('void office_oxide_free_string(void* ptr)');
const freeBytesRaw = lib.func('void office_oxide_free_bytes(void* ptr, size_t len)');
koffi.disposable('HeapStr', 'char*', freeStringRaw);

const n = {
  version: lib.func('const char* office_oxide_version()'),
  detectFormat: lib.func('const char* office_oxide_detect_format(const char* path)'),
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

class OfficeOxideError extends Error {
  constructor(code, operation) {
    const kind = ({
      0: 'ok', 1: 'invalid argument', 2: 'io error', 3: 'parse error',
      4: 'extraction failed', 5: 'internal error', 6: 'unsupported format',
    })[code] || `code=${code}`;
    super(`office_oxide: ${operation}: ${kind}`);
    this.name = 'OfficeOxideError';
    this.code = code;
    this.operation = operation;
  }
}

const emptyToNull = (v) => (v === null || v === undefined || v === '' ? null : v);

function version() { return n.version() || ''; }
function detectFormat(p) { return emptyToNull(n.detectFormat(p)); }

class Document {
  constructor(h, src = null) { this._h = h; this._src = src; }
  static open(p) {
    const e = [0]; const h = n.documentOpen(p, e);
    if (!h) throw new OfficeOxideError(e[0], 'open');
    return new Document(h, p);
  }
  static fromBytes(data, fmt) {
    if (!(data instanceof Uint8Array)) throw new TypeError('data must be Uint8Array/Buffer');
    const e = [0]; const h = n.documentOpenFromBytes(data, data.length, fmt, e);
    if (!h) throw new OfficeOxideError(e[0], 'fromBytes');
    return new Document(h);
  }
  _ensure() { if (!this._h) throw new Error('Document is closed'); }
  get format() { this._ensure(); return emptyToNull(n.documentFormat(this._h)); }
  _call(fn, op) {
    this._ensure();
    const e = [0]; const s = fn(this._h, e);
    if (s === null || s === undefined) throw new OfficeOxideError(e[0], op);
    return s;
  }
  plainText() { return this._call(n.documentPlainText, 'plainText'); }
  toMarkdown() { return this._call(n.documentToMarkdown, 'toMarkdown'); }
  toHtml() { return this._call(n.documentToHtml, 'toHtml'); }
  toIr() { return JSON.parse(this._call(n.documentToIrJson, 'toIr')); }
  saveAs(p) {
    this._ensure(); const e = [0];
    const rc = n.documentSaveAs(this._h, p, e);
    if (rc !== 0) throw new OfficeOxideError(e[0], 'saveAs');
  }
  close() { if (this._h) { n.documentFree(this._h); this._h = null; } }
  [Symbol.dispose]() { this.close(); }
}

class EditableDocument {
  constructor(h) { this._h = h; }
  static open(p) {
    const e = [0]; const h = n.editableOpen(p, e);
    if (!h) throw new OfficeOxideError(e[0], 'open');
    return new EditableDocument(h);
  }
  _ensure() { if (!this._h) throw new Error('EditableDocument is closed'); }
  replaceText(find, repl) {
    this._ensure(); const e = [0];
    const x = n.editableReplaceText(this._h, find, repl, e);
    if (x < 0) throw new OfficeOxideError(e[0], 'replaceText');
    return Number(x);
  }
  setCell(sheetIndex, cellRef, value) {
    this._ensure();
    let t, s = '', num = 0.0;
    if (value === null || value === undefined) t = 0;
    else if (typeof value === 'string') { t = 1; s = value; }
    else if (typeof value === 'number') { t = 2; num = value; }
    else if (typeof value === 'boolean') { t = 3; num = value ? 1 : 0; }
    else throw new TypeError('value must be null, string, number, or boolean');
    const e = [0];
    const rc = n.editableSetCell(this._h, sheetIndex, cellRef, t, s, num, e);
    if (rc !== 0) throw new OfficeOxideError(e[0], 'setCell');
  }
  save(p) {
    this._ensure(); const e = [0];
    const rc = n.editableSave(this._h, p, e);
    if (rc !== 0) throw new OfficeOxideError(e[0], 'save');
  }
  close() { if (this._h) { n.editableFree(this._h); this._h = null; } }
  [Symbol.dispose]() { this.close(); }
}

function oneShot(fn, name, p) {
  const e = [0]; const s = fn(p, e);
  if (s === null || s === undefined) throw new OfficeOxideError(e[0], name);
  return s;
}

function extractText(p) { return oneShot(n.extractText, 'extractText', p); }
function toMarkdown(p) { return oneShot(n.toMarkdown, 'toMarkdown', p); }
function toHtml(p) { return oneShot(n.toHtml, 'toHtml', p); }

module.exports = {
  OfficeOxideError, Document, EditableDocument,
  version, detectFormat, extractText, toMarkdown, toHtml,
};
