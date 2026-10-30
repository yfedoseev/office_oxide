#!/usr/bin/env node
/**
 * 10_detect_and_batch: detect the format of bytes, then process a mixed batch.
 *
 * Run: node examples/javascript/10_detect_and_batch.mjs
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

const PARTS = [['word/document.xml', 'docx'], ['xl/workbook.xml', 'xlsx'], ['ppt/presentation.xml', 'pptx']];

// The binding detects from a path only. A ZIP stores each part name in the
// clear, and the main part's name tells the three formats apart.
function detectFormat(data) {
  if (data.subarray(0, 2).toString() !== 'PK') return null;
  for (const [part, format] of PARTS) {
    if (data.includes(part)) return format;
  }
  return null;
}

function check(condition, message) {
  if (!condition) {
    console.error(`FAIL: ${message}`);
    process.exit(1);
  }
}

function main() {
  console.log('step 1: build a DOCX, an XLSX and a PPTX in memory');
  const batch = [
    ['report.docx', buildDocxBytes()],
    ['report.xlsx', buildXlsxBytes()],
    ['report.pptx', buildPptxBytes()],
  ];

  console.log('step 2: detect each format and open it from bytes');
  for (const [name, data] of batch) {
    const format = detectFormat(data);
    check(format === name.split('.').pop(), 'the detected format differs from the one built');
    const doc = Document.fromBytes(data, format);
    let text;
    try {
      text = doc.plainText();
    } finally {
      doc.close();
    }
    check(text.trim() !== '', 'a document has no text');
    console.log(`  ${name}: format ${format}, chars ${text.length}`);
  }

  console.log('step 3: check that bytes that are not a document are not detected');
  check(detectFormat(Buffer.from('not an office file')) === null, 'bytes that are not a document were detected as one');

  console.log('OK 10_detect_and_batch');
}

main();
