#!/usr/bin/env node
/**
 * 07_all_outputs — one document, every output: plain text, Markdown, HTML and IR JSON.
 *
 * Run: OFFICE_OXIDE_LIB=.../liboffice_oxide.so node examples/javascript/07_all_outputs.mjs
 */
import { tmpdir } from 'os';
import { join } from 'path';
import { readFileSync, unlinkSync } from 'fs';
import { Document, XlsxWriter, PptxWriter, createFromMarkdown } from '../../js/lib/index.js';

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

// plainText: text only, no structure; for search and LLM input.
// toMarkdown: headings, lists and tables as Markdown; keeps structure.
// toHtml: the same structure as HTML elements, for display.
// toIr: the format-neutral document tree as JSON, for programs that need
//       every node rather than a rendering.
console.log('step 1: build a DOCX in memory');
const bytes = buildDocxBytes();

console.log('step 2: open it from bytes');
const doc = Document.fromBytes(bytes, 'docx');
try {
  const text = doc.plainText();
  const md = doc.toMarkdown();
  const html = doc.toHtml();
  const irJson = JSON.stringify(doc.toIr()); // already parsed; re-serialise to count chars
  const ir = JSON.parse(irJson);

  console.log('step 3: check every output');
  for (const [name, out] of [['text', text], ['markdown', md], ['html', html], ['ir', irJson]]) {
    console.log(`  ${name} chars: ${out.length}`);
    check(out.length > 0, `${name} is not empty`);
    check(out.includes('Quarterly Report'), `${name} contains "Quarterly Report"`);
  }
  check(Array.isArray(ir.sections), 'IR has a sections array');
  check(/^#+\s*Quarterly Report/m.test(md), 'markdown has a heading marker');
  check(/<h\d/i.test(html) && /<table/i.test(html), 'html has heading and table elements');
  check(!/<h\d/i.test(text) && !/^#/m.test(text), 'plain text has no markup');
} finally { doc.close(); }

console.log('OK 07_all_outputs');
