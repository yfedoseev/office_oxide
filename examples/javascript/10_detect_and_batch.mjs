#!/usr/bin/env node
/**
 * 10_detect_and_batch — detect the format of each document, then process a mixed batch.
 *
 * Run: OFFICE_OXIDE_LIB=.../liboffice_oxide.so node examples/javascript/10_detect_and_batch.mjs
 */
import { tmpdir } from 'os';
import { join } from 'path';
import { readFileSync, unlinkSync, mkdtempSync, writeFileSync, rmSync } from 'fs';
import { Document, XlsxWriter, PptxWriter, createFromMarkdown, detectFormat } from '../../js/lib/index.js';

function check(cond, message) {
  if (!cond) { console.error(`FAIL: ${message}`); process.exit(1); }
}

const MARKDOWN = (title) => `# ${title}

Revenue grew in every region.

| Region | Q1  | Q2  |
|--------|-----|-----|
| North  | 120 | 135 |
| South  | 95  | 110 |
`;

const PNG = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==',
  'base64');

// The JS binding has no DOCX writer, so the DOCX is generated into a temp
// file that is read and deleted at once; everything after that is bytes.
function buildDocxBytes(title = 'Quarterly Report') {
  const path = join(tmpdir(), `oo_${process.pid}_${Math.random().toString(36).slice(2)}.docx`);
  try {
    createFromMarkdown(MARKDOWN(title), 'docx', path);
    return readFileSync(path);
  } finally {
    try { unlinkSync(path); } catch (_) {}
  }
}

// Row and column indices are zero-based.
function buildXlsxBytes() {
  const w = new XlsxWriter();
  try {
    const s = w.addSheet('Sales');
    ['Region', 'Q1', 'Q2', 'Growth'].forEach((h, c) => w.setCellStyled(s, 0, c, h, true, null));
    w.setCell(s, 1, 0, 'North'); w.setCell(s, 1, 1, 120); w.setCell(s, 1, 2, 135); w.setCell(s, 1, 3, true);
    w.setCell(s, 2, 0, 'South'); w.setCell(s, 2, 1, 95); w.setCell(s, 2, 2, 110);
    w.setCell(s, 3, 0, 'Total');
    w.setFormula(s, 3, 1, 'SUM(B2:B3)');
    w.setFormula(s, 3, 2, 'SUM(C2:C3)');
    w.setCell(s, 5, 0, 'Notes');
    w.mergeCells(s, 5, 0, 1, 3); // A6:C6
    w.setColumnWidth(s, 0, 18);
    return w.toBytes();
  } finally { w.close(); }
}

function buildPptxBytes() {
  const w = new PptxWriter();
  try {
    w.setPresentationSize(12192000, 6858000); // 16:9, in EMU
    const s1 = w.addSlide();
    w.setSlideTitle(s1, 'Welcome');
    w.addSlideText(s1, 'Quarterly Report');
    const s2 = w.addSlide();
    w.setSlideTitle(s2, 'Results');
    w.addSlideText(s2, 'Revenue grew in every region.');
    w.addSlideImage(s2, PNG, 'png', 914400, 2743200, 1828800, 1828800);
    return w.toBytes();
  } finally { w.close(); }
}

console.log('step 1: build a DOCX, an XLSX and a PPTX in memory');
const batch = [
  ['docx', buildDocxBytes(), 'Quarterly Report'],
  ['xlsx', buildXlsxBytes(), 'Region'],
  ['pptx', buildPptxBytes(), 'Quarterly Report'],
];

// The JS binding detects from a path only (no bytes variant), so each
// document is written under the name report.<ext> in a temp dir. detectFormat
// goes by the extension, not the content (see the report in the PR).
const dir = mkdtempSync(join(tmpdir(), 'oo_10_'));
try {
  console.log('step 2: detect the format of each document');
  const paths = [];
  for (const [ext, bytes, word] of batch) {
    const path = join(dir, `report.${ext}`);
    writeFileSync(path, bytes);
    const format = detectFormat(path);
    check(format === ext, `detected format is ${ext}`);
    paths.push([path, ext, word]);
  }

  console.log('step 3: process the batch');
  for (const [path, ext, word] of paths) {
    const doc = Document.open(path);
    try {
      const text = doc.plainText();
      console.log(`  report.${ext}: format ${doc.format}, chars ${text.length}`);
      check(text.includes(word), `text contains "${word}"`);
    } finally { doc.close(); }
  }
} finally {
  rmSync(dir, { recursive: true, force: true });
}

console.log('OK 10_detect_and_batch');
