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
