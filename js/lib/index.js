// SPDX-License-Identifier: MIT OR Apache-2.0
// ESM entry point for office-oxide.
//
// The native functions that allocate C strings already return JS strings
// (decoded and auto-freed via the `HeapStr` disposable type registered in
// native.js). Errors surface as null returns + a non-zero error code.

import koffi from 'koffi';
import { native } from './native.js';

export class OfficeOxideError extends Error {
  constructor(code, operation) {
    const kind = ({
      0: 'ok', 1: 'invalid argument', 2: 'io error', 3: 'parse error',
      4: 'extraction failed', 5: 'internal error', 6: 'unsupported format',
    })[code] ?? `code=${code}`;
    super(`office_oxide: ${operation}: ${kind}`);
    this.name = 'OfficeOxideError';
    this.code = code;
    this.operation = operation;
  }
}

// koffi normalises null C strings to either null, undefined or ''; treat all as absent.
const emptyToNull = (v) => (v === null || v === undefined || v === '' ? null : v);

export function version() {
  return native.version() ?? '';
}

export function detectFormat(path) {
  return emptyToNull(native.detectFormat(path));
}

export class Document {
  #handle;
  #source;

  constructor(handle, source = null) {
    this.#handle = handle;
    this.#source = source;
  }

  static open(path) {
    const err = [0];
    const h = native.documentOpen(path, err);
    if (!h) throw new OfficeOxideError(err[0], 'open');
    return new Document(h, path);
  }

  static fromBytes(data, format) {
    if (!(data instanceof Uint8Array))
      throw new TypeError('data must be a Uint8Array or Buffer');
    const err = [0];
    const h = native.documentOpenFromBytes(data, data.length, format, err);
    if (!h) throw new OfficeOxideError(err[0], 'fromBytes');
    return new Document(h, null);
  }

  #ensure() {
    if (!this.#handle) throw new Error('Document is closed');
  }

  get format() {
    this.#ensure();
    return emptyToNull(native.documentFormat(this.#handle));
  }

  #callStr(fn, op) {
    this.#ensure();
    const err = [0];
    const s = fn(this.#handle, err);
    if (s === null || s === undefined) throw new OfficeOxideError(err[0], op);
    return s;
  }

  plainText() { return this.#callStr(native.documentPlainText, 'plainText'); }
  toMarkdown() { return this.#callStr(native.documentToMarkdown, 'toMarkdown'); }
  toHtml() { return this.#callStr(native.documentToHtml, 'toHtml'); }
  toIr() { return JSON.parse(this.#callStr(native.documentToIrJson, 'toIr')); }

  saveAs(path) {
    this.#ensure();
    const err = [0];
    const rc = native.documentSaveAs(this.#handle, path, err);
    if (rc !== 0) throw new OfficeOxideError(err[0], 'saveAs');
  }

  close() {
    if (this.#handle) {
      native.documentFree(this.#handle);
      this.#handle = null;
    }
  }

  [Symbol.dispose]() { this.close(); }
}

export class EditableDocument {
  #handle;

  constructor(handle) { this.#handle = handle; }

  static open(path) {
    const err = [0];
    const h = native.editableOpen(path, err);
    if (!h) throw new OfficeOxideError(err[0], 'open');
    return new EditableDocument(h);
  }

  #ensure() {
    if (!this.#handle) throw new Error('EditableDocument is closed');
  }

  replaceText(find, replace) {
    this.#ensure();
    const err = [0];
    const n = native.editableReplaceText(this.#handle, find, replace, err);
    if (n < 0) throw new OfficeOxideError(err[0], 'replaceText');
    return Number(n);
  }

  setCell(sheetIndex, cellRef, value) {
    this.#ensure();
    let t, s = '', num = 0.0;
    if (value === null || value === undefined) t = 0;
    else if (typeof value === 'string') { t = 1; s = value; }
    else if (typeof value === 'number') { t = 2; num = value; }
    else if (typeof value === 'boolean') { t = 3; num = value ? 1 : 0; }
    else throw new TypeError('value must be null, string, number, or boolean');
    const err = [0];
    const rc = native.editableSetCell(this.#handle, sheetIndex, cellRef, t, s, num, err);
    if (rc !== 0) throw new OfficeOxideError(err[0], 'setCell');
  }

  save(path) {
    this.#ensure();
    const err = [0];
    const rc = native.editableSave(this.#handle, path, err);
    if (rc !== 0) throw new OfficeOxideError(err[0], 'save');
  }

  close() {
    if (this.#handle) {
      native.editableFree(this.#handle);
      this.#handle = null;
    }
  }

  [Symbol.dispose]() { this.close(); }
}

function oneShot(fn, name, path) {
  const err = [0];
  const s = fn(path, err);
  if (s === null || s === undefined) throw new OfficeOxideError(err[0], name);
  return s;
}

export function extractText(path) { return oneShot(native.extractText, 'extractText', path); }
export function toMarkdown(path) { return oneShot(native.toMarkdown, 'toMarkdown', path); }
export function toHtml(path) { return oneShot(native.toHtml, 'toHtml', path); }

/**
 * Convert a Markdown string to an Office document file.
 * @param {string} markdown - The Markdown content.
 * @param {string} format - One of "docx", "xlsx", or "pptx".
 * @param {string} path - Output file path.
 */
export function createFromMarkdown(markdown, format, path) {
  const err = [0];
  const rc = native.createFromMarkdown(markdown, format, path, err);
  if (rc !== 0) throw new OfficeOxideError(err[0], 'createFromMarkdown');
}

export class XlsxWriter {
  #handle;

  constructor() {
    this.#handle = native.xlsxWriterNew();
    if (!this.#handle) throw new OfficeOxideError(5, 'XlsxWriter.new');
  }

  #ensure() {
    if (!this.#handle) throw new Error('XlsxWriter is closed');
  }

  /** Add a worksheet; returns its 0-based index. */
  addSheet(name) {
    this.#ensure();
    return native.xlsxWriterAddSheet(this.#handle, name);
  }

  /** Set a cell value. value: null | string | number | boolean */
  setCell(sheet, row, col, value) {
    this.#ensure();
    let t, s = null, n = 0;
    if (value === null || value === undefined) t = 0;
    else if (typeof value === 'string') { t = 1; s = value; }
    else if (typeof value === 'number') { t = 2; n = value; }
    else if (typeof value === 'boolean') { t = 2; n = value ? 1 : 0; }
    else { t = 1; s = String(value); }
    native.xlsxSheetSetCell(this.#handle, sheet, row, col, t, s, n);
  }

  /** Set a cell with styling. bgColor: 6-char hex string or null. */
  setCellStyled(sheet, row, col, value, bold, bgColor = null) {
    this.#ensure();
    let t, s = null, n = 0;
    if (value === null || value === undefined) t = 0;
    else if (typeof value === 'string') { t = 1; s = value; }
    else if (typeof value === 'number') { t = 2; n = value; }
    else { t = 1; s = String(value); }
    native.xlsxSheetSetCellStyled(this.#handle, sheet, row, col, t, s, n, bold, bgColor);
  }

  /** Merge a rectangular range. rowSpan and colSpan must be >= 1. */
  mergeCells(sheet, row, col, rowSpan, colSpan) {
    this.#ensure();
    native.xlsxSheetMergeCells(this.#handle, sheet, row, col, rowSpan, colSpan);
  }

  /** Set column width in Excel character units (e.g. 20.0). */
  setColumnWidth(sheet, col, width) {
    this.#ensure();
    native.xlsxSheetSetColumnWidth(this.#handle, sheet, col, width);
  }

  save(path) {
    this.#ensure();
    const err = [0];
    const rc = native.xlsxWriterSave(this.#handle, path, err);
    if (rc !== 0) throw new OfficeOxideError(err[0], 'XlsxWriter.save');
  }

  toBytes() {
    this.#ensure();
    const outLen = [0];
    const err = [0];
    const ptr = native.xlsxWriterToBytes(this.#handle, outLen, err);
    if (!ptr) throw new OfficeOxideError(err[0], 'XlsxWriter.toBytes');
    try {
      return Buffer.from(koffi.decode(ptr, 'uint8_t', outLen[0]));
    } finally {
      native.freeBytes(ptr, outLen[0]);
    }
  }

  close() {
    if (this.#handle) {
      native.xlsxWriterFree(this.#handle);
      this.#handle = null;
    }
  }

  [Symbol.dispose]() { this.close(); }
}

export class PptxWriter {
  #handle;

  constructor() {
    this.#handle = native.pptxWriterNew();
    if (!this.#handle) throw new OfficeOxideError(5, 'PptxWriter.new');
  }

  #ensure() {
    if (!this.#handle) throw new Error('PptxWriter is closed');
  }

  /** Override canvas size. 914400 EMU = 1 inch. */
  setPresentationSize(cx, cy) {
    this.#ensure();
    native.pptxWriterSetPresentationSize(this.#handle, BigInt(cx), BigInt(cy));
  }

  /** Add a slide; returns its 0-based index. */
  addSlide() {
    this.#ensure();
    return native.pptxWriterAddSlide(this.#handle);
  }

  /** Set the title of a slide. */
  setSlideTitle(slide, title) {
    this.#ensure();
    native.pptxSlideSetTitle(this.#handle, slide, title);
  }

  /** Add a plain text paragraph to the slide body. */
  addSlideText(slide, text) {
    this.#ensure();
    native.pptxSlideAddText(this.#handle, slide, text);
  }

  /**
   * Embed an image on a slide.
   * data: Buffer or Uint8Array; format: "png" | "jpeg" | "gif"
   * x, y, cx, cy: EMU coordinates (914400 = 1 inch)
   */
  addSlideImage(slide, data, format, x, y, cx, cy) {
    this.#ensure();
    const buf = data instanceof Uint8Array ? data : new Uint8Array(data);
    native.pptxSlideAddImage(this.#handle, slide, buf, buf.length, format, BigInt(x), BigInt(y), BigInt(cx), BigInt(cy));
  }

  save(path) {
    this.#ensure();
    const err = [0];
    const rc = native.pptxWriterSave(this.#handle, path, err);
    if (rc !== 0) throw new OfficeOxideError(err[0], 'PptxWriter.save');
  }

  toBytes() {
    this.#ensure();
    const outLen = [0];
    const err = [0];
    const ptr = native.pptxWriterToBytes(this.#handle, outLen, err);
    if (!ptr) throw new OfficeOxideError(err[0], 'PptxWriter.toBytes');
    try {
      return Buffer.from(koffi.decode(ptr, 'uint8_t', outLen[0]));
    } finally {
      native.freeBytes(ptr, outLen[0]);
    }
  }

  close() {
    if (this.#handle) {
      native.pptxWriterFree(this.#handle);
      this.#handle = null;
    }
  }

  [Symbol.dispose]() { this.close(); }
}
