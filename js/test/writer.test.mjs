// Behavioural tests for the writers and the in-memory editing surface.
// Self-contained: every document is built here, no fixture file needed.
import { test } from 'node:test';
import { strict as assert } from 'node:assert';
import { createRequire } from 'node:module';
import {
  Document,
  EditableDocument,
  Formula,
  OfficeOxideError,
  PptxWriter,
  XlsxWriter,
} from '../lib/index.js';
import * as esm from '../lib/index.js';

const text = (bytes) => Document.fromBytes(bytes, 'xlsx').plainText();

test('booleans are written as booleans, not numbers or text', () => {
  const w = new XlsxWriter();
  w.addSheet('S');
  w.setCell(0, 0, 0, true);
  w.setCellStyled(0, 1, 0, false, true);
  const t = text(w.toBytes());
  assert.ok(t.includes('TRUE'), t);
  assert.ok(t.includes('FALSE'), t);
  assert.ok(!t.includes('false'), t);
  w.close();
});

test('an out-of-range write throws instead of vanishing', () => {
  const w = new XlsxWriter();
  w.addSheet('S');
  assert.throws(() => w.setCell(5, 0, 0, 'x'), OfficeOxideError);
  assert.throws(() => w.setCell(0, 1_048_576, 0, 'x'), OfficeOxideError);
  assert.throws(() => w.setCellStyled(5, 0, 0, 'x', true), OfficeOxideError);
  w.close();

  const p = new PptxWriter();
  p.addSlide();
  assert.throws(() => p.setSlideTitle(99, 'x'), OfficeOxideError);
  assert.throws(() => p.addSlideText(99, 'x'), OfficeOxideError);
  assert.throws(() => p.addSlideImage(99, new Uint8Array([1, 2, 3]), 'png', 0, 0, 1, 1), OfficeOxideError);
  assert.throws(() => p.addSlideImage(0, new Uint8Array([1, 2, 3]), 'bmp', 0, 0, 1, 1), OfficeOxideError);
  p.setSlideTitle(0, 'ok');
  p.close();
});

test('formulas are reachable; a string starting with = stays text', () => {
  const w = new XlsxWriter();
  w.addSheet('S');
  w.setCell(0, 0, 0, 2);
  w.setCell(0, 1, 0, 3);
  w.setFormula(0, 2, 0, '=SUM(A1:A2)');
  w.setCell(0, 3, 0, new Formula('A1*2'));
  w.setCell(0, 4, 0, '=literal');
  const ir = JSON.stringify(Document.fromBytes(w.toBytes(), 'xlsx').toIr());
  assert.ok(ir.includes('SUM(A1:A2)'), ir);
  assert.ok(ir.includes('A1*2'), ir);
  assert.ok(ir.includes('=literal'), ir);
  w.close();
});

test('a bigint beyond 2**53 is rejected, not rounded', () => {
  const w = new XlsxWriter();
  w.addSheet('S');
  w.setCell(0, 0, 0, 2n ** 53n);
  assert.throws(() => w.setCell(0, 1, 0, 2n ** 53n + 1n), RangeError);
  w.close();
});

test('EditableDocument.fromBytes + toBytes edit in memory', () => {
  const w = new XlsxWriter();
  w.addSheet('S');
  w.setCell(0, 0, 0, 'old');
  const ed = EditableDocument.fromBytes(w.toBytes(), 'xlsx');
  ed.setCell(0, 'A1', 'new');
  ed.setCell(0, 'B1', true);
  const t = text(ed.toBytes());
  assert.ok(t.includes('new') && t.includes('TRUE'), t);
  assert.throws(() => ed.replaceText('a', 'b'), (e) => e.code === 6);
  ed.close();
  w.close();
});

test('an empty find is an invalid-argument error', () => {
  const p = new PptxWriter();
  p.addSlide();
  p.setSlideTitle(0, 'Hello');
  const ed = EditableDocument.fromBytes(p.toBytes(), 'pptx');
  assert.throws(() => ed.replaceText('', 'X'), (e) => e.code === 1);
  assert.equal(ed.replaceText('Hello', 'Bye'), 1);
  ed.close();
  p.close();
});

test('toMarkdownWithImages embeds images', () => {
  const png = Uint8Array.from([
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
    0xde, 0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, 0x08, 0xd7, 0x63, 0xf8, 0xcf, 0xc0, 0x00,
    0x00, 0x03, 0x01, 0x01, 0x00, 0x18, 0xdd, 0x8d, 0xb0, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e,
    0x44, 0xae, 0x42, 0x60, 0x82,
  ]);
  const p = new PptxWriter();
  p.addSlide();
  p.addSlideImage(0, png, 'png', 0, 0, 914400, 914400);
  const doc = Document.fromBytes(p.toBytes(), 'pptx');
  assert.ok(doc.toMarkdownWithImages().includes('[image-base64:'));
  assert.ok(!doc.toMarkdown().includes('[image-base64:'));
  doc.close();
  p.close();
});

test('the CommonJS entry point exports the same API as the ESM one', () => {
  const cjs = createRequire(import.meta.url)('../lib/index.cjs');
  assert.deepEqual(Object.keys(cjs).sort(), Object.keys(esm).sort());
  for (const cls of ['Document', 'EditableDocument', 'XlsxWriter', 'PptxWriter']) {
    const names = (c) => Object.getOwnPropertyNames(c.prototype).concat(Object.getOwnPropertyNames(c)).sort();
    assert.deepEqual(names(cjs[cls]), names(esm[cls]), cls);
  }
});
