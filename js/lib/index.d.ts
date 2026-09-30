// Type definitions for office-oxide (Node native bindings).

export type DocumentFormat = 'docx' | 'xlsx' | 'pptx' | 'doc' | 'xls' | 'ppt';
export type CellValue = null | string | number | boolean;

/** A formula cell value for XlsxWriter, e.g. `new Formula('SUM(A1:A3)')`. */
export class Formula {
  constructor(formula: string);
  readonly formula: string;
}

/**
 * A value XlsxWriter accepts. A `bigint` beyond ±2**53 throws a RangeError
 * (Excel stores numbers as doubles); a string starting with '=' stays text.
 */
export type WriterCellValue = CellValue | bigint | Formula;

/**
 * A failed call. `code`: 1 invalid argument, 2 I/O, 3 parse, 4 extraction,
 * 5 internal, 6 unsupported format or out-of-range target (nothing written).
 */
export class OfficeOxideError extends Error {
  readonly code: number;
  readonly operation: string;
}

export class Document implements Disposable {
  static open(path: string): Document;
  static fromBytes(data: Uint8Array, format: DocumentFormat): Document;
  readonly format: DocumentFormat | null;
  plainText(): string;
  toMarkdown(): string;
  /** Markdown with each image embedded inline as `[image-base64:…]`. */
  toMarkdownWithImages(): string;
  toHtml(): string;
  toIr(): unknown;
  saveAs(path: string): void;
  close(): void;
  [Symbol.dispose](): void;
}

export class EditableDocument implements Disposable {
  static open(path: string): EditableDocument;
  static fromBytes(data: Uint8Array, format: 'docx' | 'xlsx' | 'pptx'): EditableDocument;
  /** Throws (code 1) for an empty `find`, (code 6) for XLSX. */
  replaceText(find: string, replace: string): number;
  setCell(sheetIndex: number, cellRef: string, value: CellValue | bigint): void;
  save(path: string): void;
  toBytes(): Buffer;
  close(): void;
  [Symbol.dispose](): void;
}

export function version(): string;
export function detectFormat(path: string): DocumentFormat | null;
export function extractText(path: string): string;
export function toMarkdown(path: string): string;
export function toHtml(path: string): string;
export function createFromMarkdown(markdown: string, format: 'docx' | 'xlsx' | 'pptx', path: string): void;

export type ImageFormat = 'png' | 'jpeg' | 'jpg' | 'gif';

export class XlsxWriter implements Disposable {
  constructor();
  addSheet(name: string): number;
  /** Throws OfficeOxideError (code 6) when the sheet or cell is out of range. */
  setCell(sheet: number, row: number, col: number, value: WriterCellValue): void;
  /** A leading '=' is accepted. */
  setFormula(sheet: number, row: number, col: number, formula: string): void;
  setCellStyled(sheet: number, row: number, col: number, value: WriterCellValue, bold: boolean, bgColor?: string | null): void;
  mergeCells(sheet: number, row: number, col: number, rowSpan: number, colSpan: number): void;
  setColumnWidth(sheet: number, col: number, width: number): void;
  save(path: string): void;
  toBytes(): Buffer;
  close(): void;
  [Symbol.dispose](): void;
}

export class PptxWriter implements Disposable {
  constructor();
  setPresentationSize(cx: number | bigint, cy: number | bigint): void;
  addSlide(): number;
  /** Throws OfficeOxideError (code 6) when the slide does not exist. */
  setSlideTitle(slide: number, title: string): void;
  addSlideText(slide: number, text: string): void;
  addSlideImage(slide: number, data: Uint8Array | Buffer, format: ImageFormat, x: number | bigint, y: number | bigint, cx: number | bigint, cy: number | bigint): void;
  save(path: string): void;
  toBytes(): Buffer;
  close(): void;
  [Symbol.dispose](): void;
}
