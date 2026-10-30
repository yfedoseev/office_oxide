#!/usr/bin/env node
/**
 * 05_xlsx_writer — write a workbook with typed cells, a formula, a style, a merge and a width, then read it back.
 *
 * Run: OFFICE_OXIDE_LIB=.../liboffice_oxide.so node examples/javascript/05_xlsx_writer.mjs
 */
import { tmpdir } from 'os';
import { join } from 'path';
import { readFileSync, unlinkSync } from 'fs';
import { Document, XlsxWriter, createFromMarkdown, PptxWriter } from '../../js/lib/index.js';

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

console.log('step 1: build an XLSX in memory');
const bytes = buildXlsxBytes();

console.log('step 2: open it from bytes');
const doc = Document.fromBytes(bytes, 'xlsx');
try {
  const text = doc.plainText();
  const md = doc.toMarkdown();
  console.log(`  format: ${doc.format}`);
  console.log(`  chars: ${text.length}`);
  console.log(md);

  console.log('step 3: check the cells');
  for (const e of ['Sales', 'Region', 'Q1', 'Q2', 'Growth', 'North', '120', '135', 'South', '95', '110', 'Total', 'Notes']) {
    check(text.includes(e) || md.includes(e), `text contains "${e}"`);
  }
  check(md.includes('TRUE'), 'boolean cell D2 is TRUE');
  // Formulas are written without a cached value, so readers show the formula.
  check(md.includes('SUM(B2:B3)') && md.includes('SUM(C2:C3)'), 'formulas are preserved');
} finally { doc.close(); }

console.log('step 4: an out-of-range target is an error');
const w = new XlsxWriter();
try {
  w.setCell(99, 0, 0, 'x');
  check(false, 'setCell on a missing sheet throws');
} catch (e) {
  check(e.code === 6, 'error code is 6 (unsupported or out of range)');
  console.log(`  error: ${e.message}`);
} finally { w.close(); }

console.log('OK 05_xlsx_writer');
