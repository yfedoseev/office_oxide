#!/usr/bin/env node
/**
 * 03_edit: replace text in a DOCX and set cells in an XLSX, in memory.
 *
 * Run: node examples/javascript/03_edit.mjs
 */
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Document, EditableDocument, XlsxWriter, createFromMarkdown } from '../../js/lib/index.js';

const MARKDOWN = `# Quarterly Report

Revenue grew in every region.

| Region | Q1  | Q2  |
|--------|-----|-----|
| North  | 120 | 135 |
| South  | 95  | 110 |
`;

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

function check(condition, message) {
  if (!condition) {
    console.error(`FAIL: ${message}`);
    process.exit(1);
  }
}

function has(text, expected) {
  check(text.includes(expected), `"${expected}" is missing from the text`);
}

// Opens bytes, reads the plain text and closes the document.
function textOf(data, format) {
  const doc = Document.fromBytes(data, format);
  try {
    return doc.plainText();
  } finally {
    doc.close();
  }
}

function main() {
  console.log('step 1: build a DOCX in memory');
  const docx = buildDocxBytes();

  console.log('step 2: replace text');
  let ed = EditableDocument.fromBytes(docx, 'docx');
  let edited;
  try {
    const replaced = ed.replaceText('Quarterly Report', 'Annual Report');
    console.log(`  replacements: ${replaced}`);
    check(replaced === 1, 'replace_text did not make exactly one replacement');

    console.log('step 3: save to bytes and check the text');
    edited = ed.toBytes();
  } finally {
    ed.close();
  }
  let text = textOf(edited, 'docx');
  has(text, 'Annual Report');
  check(!text.includes('Quarterly Report'), 'the old title is still present');

  console.log('step 4: build an XLSX in memory');
  const xlsx = buildXlsxBytes();

  console.log('step 5: set cells');
  ed = EditableDocument.fromBytes(xlsx, 'xlsx');
  try {
    ed.setCell(0, 'B2', 999);
    ed.setCell(0, 'A2', 'Central');

    console.log('step 6: save to bytes and check the cells');
    edited = ed.toBytes();
  } finally {
    ed.close();
  }
  text = textOf(edited, 'xlsx');
  has(text, 'Central');
  has(text, '999');
  check(!text.includes('North'), 'the old A2 value is still present');

  console.log('OK 03_edit');
}

main();
