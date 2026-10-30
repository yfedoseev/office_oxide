#!/usr/bin/env node
/**
 * 04_open_from_bytes: build a DOCX, an XLSX and a PPTX in memory and open each from its bytes.
 *
 * Run: node examples/javascript/04_open_from_bytes.mjs
 */
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Document, PptxWriter, XlsxWriter, createFromMarkdown } from '../../js/lib/index.js';

const MARKDOWN = `# Quarterly Report

Revenue grew in every region.

| Region | Q1  | Q2  |
|--------|-----|-----|
| North  | 120 | 135 |
| South  | 95  | 110 |
`;

const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==', 'base64');

// createFromMarkdown writes to a path, so the file lives in a private temp
// directory just long enough to read its bytes back.
function markdownToBytes(markdown, format) {
  const dir = mkdtempSync(join(tmpdir(), 'oo_example_'));
  try {
    const path = join(dir, `report.${format}`);
    createFromMarkdown(markdown, format, path);
    return readFileSync(path);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

function buildDocxBytes(title = 'Quarterly Report') {
  return markdownToBytes(MARKDOWN.replaceAll('Quarterly Report', title), 'docx');
}

function buildXlsxBytes() {
  const w = new XlsxWriter();
  try {
    const s = w.addSheet('Sales');
    ['Region', 'Q1', 'Q2', 'Growth'].forEach((head, col) => w.setCellStyled(s, 0, col, head, true, null)); // bold header row
    w.setCell(s, 1, 0, 'North'); w.setCell(s, 1, 1, 120); w.setCell(s, 1, 2, 135); w.setCell(s, 1, 3, true); // D2
    w.setCell(s, 2, 0, 'South'); w.setCell(s, 2, 1, 95); w.setCell(s, 2, 2, 110);
    w.setCell(s, 3, 0, 'Total');
    w.setFormula(s, 3, 1, 'SUM(B2:B3)');
    w.setFormula(s, 3, 2, 'SUM(C2:C3)');
    w.setCell(s, 5, 0, 'Notes');
    w.mergeCells(s, 5, 0, 1, 3); // A6:C6
    w.setColumnWidth(s, 0, 18);
    return Buffer.from(w.toBytes());
  } finally {
    w.close();
  }
}

function buildPptxBytes() {
  const w = new PptxWriter();
  try {
    w.setPresentationSize(12192000, 6858000); // 16:9, in EMU (914400 per inch)
    const one = w.addSlide();
    w.setSlideTitle(one, 'Welcome');
    w.addSlideText(one, 'Quarterly Report');
    const two = w.addSlide();
    w.setSlideTitle(two, 'Results');
    w.addSlideText(two, 'Revenue grew in every region.');
    w.addSlideImage(two, PNG, 'png', 914400, 2743200, 1828800, 1828800);
    return Buffer.from(w.toBytes());
  } finally {
    w.close();
  }
}

function check(condition, message) {
  if (!condition) {
    console.error(`FAIL: ${message}`);
    process.exit(1);
  }
}

function has(text, expected) {
  check(text.includes(expected), `"${expected}" is missing from the text`);
}

function main() {
  console.log('step 1: build a DOCX, an XLSX and a PPTX in memory');
  const built = { docx: buildDocxBytes(), xlsx: buildXlsxBytes(), pptx: buildPptxBytes() };

  console.log('step 2: open each from bytes');
  const texts = {};
  for (const [format, data] of Object.entries(built)) {
    const doc = Document.fromBytes(data, format);
    try {
      check(doc.format === format, 'the opened format differs from the requested one');
      texts[format] = doc.plainText();
    } finally {
      doc.close();
    }
    console.log(`  format: ${format}`);
    console.log(`  chars: ${texts[format].length}`);
  }

  console.log('step 3: check the text');
  has(texts.docx, 'Revenue grew in every region.');
  has(texts.xlsx, 'North');
  has(texts.xlsx, 'Total');
  has(texts.pptx, 'Welcome');
  has(texts.pptx, 'Results');

  console.log('step 4: check that a wrong format is an error');
  let error = null;
  try {
    Document.fromBytes(built.xlsx, 'pptx').close();
  } catch (e) {
    error = e;
  }
  check(error !== null, 'an XLSX opened as a PPTX was accepted');
  console.log(`  error: ${error.message}`);

  console.log('OK 04_open_from_bytes');
}

main();
