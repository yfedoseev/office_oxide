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

class XlsxWriter {
  constructor() {
    this._h = n.xlsxWriterNew();
    if (!this._h) throw new OfficeOxideError(5, 'XlsxWriter.new');
  }
  _ensure() { if (!this._h) throw new Error('XlsxWriter is closed'); }
  addSheet(name) {
    this._ensure();
    return n.xlsxWriterAddSheet(this._h, name);
  }
  setCell(sheet, row, col, value) {
    this._ensure();
    let t, s = null, num = 0;
    if (value === null || value === undefined) t = 0;
    else if (typeof value === 'string') { t = 1; s = value; }
    else if (typeof value === 'number') { t = 2; num = value; }
    else if (typeof value === 'boolean') { t = 2; num = value ? 1 : 0; }
    else { t = 1; s = String(value); }
    n.xlsxSheetSetCell(this._h, sheet, row, col, t, s, num);
  }
  setCellStyled(sheet, row, col, value, bold, bgColor) {
    this._ensure();
    let t, s = null, num = 0;
    if (value === null || value === undefined) t = 0;
    else if (typeof value === 'string') { t = 1; s = value; }
    else if (typeof value === 'number') { t = 2; num = value; }
    else { t = 1; s = String(value); }
    n.xlsxSheetSetCellStyled(this._h, sheet, row, col, t, s, num, bold, bgColor || null);
  }
  mergeCells(sheet, row, col, rowSpan, colSpan) {
    this._ensure();
    n.xlsxSheetMergeCells(this._h, sheet, row, col, rowSpan, colSpan);
  }
  setColumnWidth(sheet, col, width) {
    this._ensure();
    n.xlsxSheetSetColumnWidth(this._h, sheet, col, width);
  }
  save(path) {
    this._ensure();
    const e = [0];
    const rc = n.xlsxWriterSave(this._h, path, e);
    if (rc !== 0) throw new OfficeOxideError(e[0], 'XlsxWriter.save');
  }
  toBytes() {
    this._ensure();
    const outLen = [0]; const e = [0];
    const ptr = n.xlsxWriterToBytes(this._h, outLen, e);
    if (!ptr) throw new OfficeOxideError(e[0], 'XlsxWriter.toBytes');
    try {
      return Buffer.from(koffi.decode(ptr, 'uint8_t', outLen[0]));
    } finally {
      freeBytesRaw(ptr, outLen[0]);
    }
  }
  close() { if (this._h) { n.xlsxWriterFree(this._h); this._h = null; } }
  [Symbol.dispose]() { this.close(); }
}

class PptxWriter {
  constructor() {
    this._h = n.pptxWriterNew();
    if (!this._h) throw new OfficeOxideError(5, 'PptxWriter.new');
  }
  _ensure() { if (!this._h) throw new Error('PptxWriter is closed'); }
  setPresentationSize(cx, cy) {
    this._ensure();
    n.pptxWriterSetPresentationSize(this._h, BigInt(cx), BigInt(cy));
  }
  addSlide() {
    this._ensure();
    return n.pptxWriterAddSlide(this._h);
  }
  setSlideTitle(slide, title) {
    this._ensure();
    n.pptxSlideSetTitle(this._h, slide, title);
  }
  addSlideText(slide, text) {
    this._ensure();
    n.pptxSlideAddText(this._h, slide, text);
  }
  addSlideImage(slide, data, format, x, y, cx, cy) {
    this._ensure();
    const buf = data instanceof Uint8Array ? data : new Uint8Array(data);
    n.pptxSlideAddImage(this._h, slide, buf, buf.length, format, BigInt(x), BigInt(y), BigInt(cx), BigInt(cy));
  }
  save(path) {
    this._ensure();
    const e = [0];
    const rc = n.pptxWriterSave(this._h, path, e);
    if (rc !== 0) throw new OfficeOxideError(e[0], 'PptxWriter.save');
  }
  toBytes() {
    this._ensure();
    const outLen = [0]; const e = [0];
    const ptr = n.pptxWriterToBytes(this._h, outLen, e);
    if (!ptr) throw new OfficeOxideError(e[0], 'PptxWriter.toBytes');
    try {
      return Buffer.from(koffi.decode(ptr, 'uint8_t', outLen[0]));
    } finally {
      freeBytesRaw(ptr, outLen[0]);
    }
  }
  close() { if (this._h) { n.pptxWriterFree(this._h); this._h = null; } }
  [Symbol.dispose]() { this.close(); }
}

module.exports = {
  OfficeOxideError, Document, EditableDocument,
  XlsxWriter, PptxWriter,
  version, detectFormat, extractText, toMarkdown, toHtml,
};
