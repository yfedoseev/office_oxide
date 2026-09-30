// SPDX-License-Identifier: MIT OR Apache-2.0
// ESM entry point for office-oxide.
//
// The native functions that allocate C strings already return JS strings
// (decoded and auto-freed via the `HeapStr` disposable type registered in
// native.js). Errors surface as null returns + a non-zero error code.

import koffi from 'koffi';
import { native } from './native.js';

const n = native;
const freeBytesRaw = native.freeBytes;

// ─── Public API ─────────────────────────────────────────────────────────────
// This block is identical in index.js (ESM) and index.cjs (CommonJS); the
// test suite checks that the two entry points export the same names.

const ERROR_KINDS = {
  0: 'ok', 1: 'invalid argument', 2: 'io error', 3: 'parse error',
  4: 'extraction failed', 5: 'internal error', 6: 'unsupported format',
};

class OfficeOxideError extends Error {
  constructor(code, operation, detail = null) {
    const kind = ERROR_KINDS[code] ?? `code=${code}`;
    super(`office_oxide: ${operation}: ${detail ?? kind}`);
    this.name = 'OfficeOxideError';
    this.code = code;
    this.operation = operation;
  }
}

// koffi normalises null C strings to either null, undefined or ''; treat all as absent.
const emptyToNull = (v) => (v === null || v === undefined || v === '' ? null : v);

// The writers return a status so a write that landed nowhere (a missing
// sheet or slide, a cell outside Excel's grid, an unknown value type) is
// detectable. Ignoring it made those writes vanish while reporting success.
function checkStatus(rc, op) {
  if (rc === 0) return;
  const detail = rc === 6
    ? 'the target is out of range, so nothing was written'
    : null;
  throw new OfficeOxideError(rc, op, detail);
}

// Largest integer a double (every Excel number) holds exactly.
const MAX_EXACT = 2n ** 53n;

// Encode a cell value for the FFI: [value_type, value_str, value_num].
// Booleans are value_type 3; they were sent as the number 1/0 by the
// writer and as the text "true" by setCellStyled.
function encodeCell(value, op) {
  if (value === null || value === undefined) return [0, null, 0];
  if (typeof value === 'string') return [1, value, 0];
  if (typeof value === 'number') return [2, null, value];
  if (typeof value === 'boolean') return [3, null, value ? 1 : 0];
  if (typeof value === 'bigint') {
    // A bigint beyond 2**53 cannot be stored exactly as an Excel number;
    // rounding it silently would change an ID or account number.
    if (value > MAX_EXACT || value < -MAX_EXACT) {
      throw new RangeError(`${op}: ${value} cannot be stored exactly as an Excel number; write it as a string`);
    }
    return [2, null, Number(value)];
  }
  if (value instanceof Formula) return [4, value.formula, 0];
  return [1, String(value), 0];
}

/** A formula cell value, e.g. `new Formula('SUM(A1:A3)')`. */
class Formula {
  constructor(formula) {
    if (typeof formula !== 'string') throw new TypeError('formula must be a string');
    this.formula = formula;
  }
}

function version() { return n.version() ?? ''; }
function detectFormat(p) { return emptyToNull(n.detectFormat(p)); }

function takeBytes(ptr, len) {
  try {
    return Buffer.from(koffi.decode(ptr, 'uint8_t', len));
  } finally {
    freeBytesRaw(ptr, len);
  }
}

class Document {
  constructor(h, src = null) { this._h = h; this._src = src; }

  static open(p) {
    const e = [0];
    const h = n.documentOpen(p, e);
    if (!h) throw new OfficeOxideError(e[0], 'open');
    return new Document(h, p);
  }

  static fromBytes(data, fmt) {
    if (!(data instanceof Uint8Array)) throw new TypeError('data must be a Uint8Array or Buffer');
    const e = [0];
    const h = n.documentOpenFromBytes(data, data.length, fmt, e);
    if (!h) throw new OfficeOxideError(e[0], 'fromBytes');
    return new Document(h);
  }

  _ensure() { if (!this._h) throw new Error('Document is closed'); }

  get format() { this._ensure(); return emptyToNull(n.documentFormat(this._h)); }

  _call(fn, op) {
    this._ensure();
    const e = [0];
    const s = fn(this._h, e);
    if (s === null || s === undefined) throw new OfficeOxideError(e[0], op);
    return s;
  }

  plainText() { return this._call(n.documentPlainText, 'plainText'); }
  toMarkdown() { return this._call(n.documentToMarkdown, 'toMarkdown'); }
  /** Markdown with each image embedded inline as `[image-base64:…]`. */
  toMarkdownWithImages() { return this._call(n.documentToMarkdownWithImages, 'toMarkdownWithImages'); }
  toHtml() { return this._call(n.documentToHtml, 'toHtml'); }
  toIr() { return JSON.parse(this._call(n.documentToIrJson, 'toIr')); }

  saveAs(p) {
    this._ensure();
    const e = [0];
    const rc = n.documentSaveAs(this._h, p, e);
    if (rc !== 0) throw new OfficeOxideError(e[0], 'saveAs');
  }

  close() { if (this._h) { n.documentFree(this._h); this._h = null; } }
  [Symbol.dispose]() { this.close(); }
}

class EditableDocument {
  constructor(h) { this._h = h; }

  static open(p) {
    const e = [0];
    const h = n.editableOpen(p, e);
    if (!h) throw new OfficeOxideError(e[0], 'open');
    return new EditableDocument(h);
  }

  /** Open a document for editing from bytes ("docx" | "xlsx" | "pptx"). */
  static fromBytes(data, fmt) {
    if (!(data instanceof Uint8Array)) throw new TypeError('data must be a Uint8Array or Buffer');
    const e = [0];
    const h = n.editableOpenFromBytes(data, data.length, fmt, e);
    if (!h) throw new OfficeOxideError(e[0], 'fromBytes');
    return new EditableDocument(h);
  }

  _ensure() { if (!this._h) throw new Error('EditableDocument is closed'); }

  replaceText(find, repl) {
    this._ensure();
    const e = [0];
    const x = n.editableReplaceText(this._h, find, repl, e);
    if (x < 0) throw new OfficeOxideError(e[0], 'replaceText');
    return Number(x);
  }

  setCell(sheetIndex, cellRef, value) {
    this._ensure();
    if (value instanceof Formula) throw new TypeError('setCell: formulas are only supported by XlsxWriter');
    const [t, s, num] = encodeCell(value, 'setCell');
    const e = [0];
    const rc = n.editableSetCell(this._h, sheetIndex, cellRef, t, s ?? '', num, e);
    if (rc !== 0) throw new OfficeOxideError(e[0], 'setCell');
  }

  save(p) {
    this._ensure();
    const e = [0];
    const rc = n.editableSave(this._h, p, e);
    if (rc !== 0) throw new OfficeOxideError(e[0], 'save');
  }

  /** Serialize the edited document to a Buffer. */
  toBytes() {
    this._ensure();
    const outLen = [0];
    const e = [0];
    const ptr = n.editableSaveToBytes(this._h, outLen, e);
    if (!ptr) throw new OfficeOxideError(e[0], 'toBytes');
    return takeBytes(ptr, outLen[0]);
  }

  close() { if (this._h) { n.editableFree(this._h); this._h = null; } }
  [Symbol.dispose]() { this.close(); }
}

function oneShot(fn, name, p) {
  const e = [0];
  const s = fn(p, e);
  if (s === null || s === undefined) throw new OfficeOxideError(e[0], name);
  return s;
}

function extractText(p) { return oneShot(n.extractText, 'extractText', p); }
function toMarkdown(p) { return oneShot(n.toMarkdown, 'toMarkdown', p); }
function toHtml(p) { return oneShot(n.toHtml, 'toHtml', p); }

/**
 * Convert a Markdown string to an Office document file.
 * @param {string} markdown - The Markdown content.
 * @param {string} format - One of "docx", "xlsx", or "pptx".
 * @param {string} path - Output file path.
 */
function createFromMarkdown(markdown, format, path) {
  const e = [0];
  const rc = n.createFromMarkdown(markdown, format, path, e);
  if (rc !== 0) throw new OfficeOxideError(e[0], 'createFromMarkdown');
}

class XlsxWriter {
  constructor() {
    this._h = n.xlsxWriterNew();
    if (!this._h) throw new OfficeOxideError(5, 'XlsxWriter.new');
  }

  _ensure() { if (!this._h) throw new Error('XlsxWriter is closed'); }

  /** Add a worksheet; returns its 0-based index. */
  addSheet(name) {
    this._ensure();
    return n.xlsxWriterAddSheet(this._h, name);
  }

  /**
   * Set a cell value: null | string | number | boolean | bigint | Formula.
   * A string is always text, even one starting with '='; use setFormula
   * (or a Formula value) for a formula. Throws when nothing was written.
   */
  setCell(sheet, row, col, value) {
    this._ensure();
    const [t, s, num] = encodeCell(value, 'XlsxWriter.setCell');
    checkStatus(n.xlsxSheetSetCell(this._h, sheet, row, col, t, s, num), 'XlsxWriter.setCell');
  }

  /** Set a formula cell, e.g. setFormula(0, 3, 1, 'SUM(B1:B3)'). A leading '=' is accepted. */
  setFormula(sheet, row, col, formula) {
    this.setCell(sheet, row, col, new Formula(formula));
  }

  /** Set a cell with styling. bgColor: 6-char hex string or null. */
  setCellStyled(sheet, row, col, value, bold, bgColor = null) {
    this._ensure();
    const [t, s, num] = encodeCell(value, 'XlsxWriter.setCellStyled');
    checkStatus(
      n.xlsxSheetSetCellStyled(this._h, sheet, row, col, t, s, num, bold, bgColor || null),
      'XlsxWriter.setCellStyled',
    );
  }

  /** Merge a rectangular range. rowSpan and colSpan must be >= 1. */
  mergeCells(sheet, row, col, rowSpan, colSpan) {
    this._ensure();
    n.xlsxSheetMergeCells(this._h, sheet, row, col, rowSpan, colSpan);
  }

  /** Set column width in Excel character units (e.g. 20.0). */
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
    const outLen = [0];
    const e = [0];
    const ptr = n.xlsxWriterToBytes(this._h, outLen, e);
    if (!ptr) throw new OfficeOxideError(e[0], 'XlsxWriter.toBytes');
    return takeBytes(ptr, outLen[0]);
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

  /** Override canvas size. 914400 EMU = 1 inch. */
  setPresentationSize(cx, cy) {
    this._ensure();
    n.pptxWriterSetPresentationSize(this._h, BigInt(cx), BigInt(cy));
  }

  /** Add a slide; returns its 0-based index. */
  addSlide() {
    this._ensure();
    return n.pptxWriterAddSlide(this._h);
  }

  /** Set the title of a slide. Throws when the slide does not exist. */
  setSlideTitle(slide, title) {
    this._ensure();
    checkStatus(n.pptxSlideSetTitle(this._h, slide, title), 'PptxWriter.setSlideTitle');
  }

  /** Add a plain text paragraph to the slide body. Throws when the slide does not exist. */
  addSlideText(slide, text) {
    this._ensure();
    checkStatus(n.pptxSlideAddText(this._h, slide, text), 'PptxWriter.addSlideText');
  }

  /**
   * Embed an image on a slide.
   * data: Buffer or Uint8Array; format: "png" | "jpeg" | "gif"
   * x, y, cx, cy: EMU coordinates (914400 = 1 inch)
   */
  addSlideImage(slide, data, format, x, y, cx, cy) {
    this._ensure();
    const buf = data instanceof Uint8Array ? data : new Uint8Array(data);
    checkStatus(
      n.pptxSlideAddImage(this._h, slide, buf, buf.length, format, BigInt(x), BigInt(y), BigInt(cx), BigInt(cy)),
      'PptxWriter.addSlideImage',
    );
  }

  save(path) {
    this._ensure();
    const e = [0];
    const rc = n.pptxWriterSave(this._h, path, e);
    if (rc !== 0) throw new OfficeOxideError(e[0], 'PptxWriter.save');
  }

  toBytes() {
    this._ensure();
    const outLen = [0];
    const e = [0];
    const ptr = n.pptxWriterToBytes(this._h, outLen, e);
    if (!ptr) throw new OfficeOxideError(e[0], 'PptxWriter.toBytes');
    return takeBytes(ptr, outLen[0]);
  }

  close() { if (this._h) { n.pptxWriterFree(this._h); this._h = null; } }
  [Symbol.dispose]() { this.close(); }
}

export {
  OfficeOxideError,
  Document,
  EditableDocument,
  XlsxWriter,
  PptxWriter,
  Formula,
  version,
  detectFormat,
  extractText,
  toMarkdown,
  toHtml,
  createFromMarkdown,
};
