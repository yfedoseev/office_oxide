#!/usr/bin/env node
/**
 * 05_xlsx_writer: write a workbook with typed cells, a formula, a style, a merge and a column width, then read it back.
 *
 * Run: node examples/javascript/05_xlsx_writer.mjs
 */
import { Document, XlsxWriter } from '../../js/lib/index.js';

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
  console.log('step 1: build an XLSX in memory');
  const data = buildXlsxBytes();

  console.log('step 2: open it from bytes');
  const doc = Document.fromBytes(data, 'xlsx');
  let text, sheets;
  try {
    text = doc.plainText();
    sheets = doc.toIr().sections.length;
    console.log(`  format: ${doc.format}`);
    console.log(`  chars: ${text.length}`);
    console.log(`  sheets: ${sheets}`);
  } finally {
    doc.close();
  }

  console.log('step 3: check the cells');
  check(sheets === 1, 'expected one sheet');
  for (const expected of ['Sales', 'Region', 'Q1', 'Q2', 'Growth', 'North', '120', '135', 'South', '95', '110', 'Total', 'Notes']) {
    has(text, expected);
  }
  check(text.includes('TRUE'), 'the boolean cell is missing');
  check(text.includes('SUM(B2:B3)') && text.includes('SUM(C2:C3)'), 'a formula is missing');

  console.log('OK 05_xlsx_writer');
}

main();
