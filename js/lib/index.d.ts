// Type definitions for office-oxide (Node native bindings).

export type DocumentFormat = 'docx' | 'xlsx' | 'pptx' | 'doc' | 'xls' | 'ppt';
export type CellValue = null | string | number | boolean;

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
  toHtml(): string;
  toIr(): unknown;
  saveAs(path: string): void;
  close(): void;
  [Symbol.dispose](): void;
}

export class EditableDocument implements Disposable {
  static open(path: string): EditableDocument;
  replaceText(find: string, replace: string): number;
  setCell(sheetIndex: number, cellRef: string, value: CellValue): void;
  save(path: string): void;
  close(): void;
  [Symbol.dispose](): void;
}

export function version(): string;
export function detectFormat(path: string): DocumentFormat | null;
export function extractText(path: string): string;
export function toMarkdown(path: string): string;
export function toHtml(path: string): string;

export type ImageFormat = 'png' | 'jpeg' | 'jpg' | 'gif';

export class XlsxWriter implements Disposable {
  constructor();
  addSheet(name: string): number;
  setCell(sheet: number, row: number, col: number, value: CellValue): void;
  setCellStyled(sheet: number, row: number, col: number, value: CellValue, bold: boolean, bgColor?: string | null): void;
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
  setSlideTitle(slide: number, title: string): void;
  addSlideText(slide: number, text: string): void;
  addSlideImage(slide: number, data: Uint8Array | Buffer, format: ImageFormat, x: number | bigint, y: number | bigint, cx: number | bigint, cy: number | bigint): void;
  save(path: string): void;
  toBytes(): Buffer;
  close(): void;
  [Symbol.dispose](): void;
}
