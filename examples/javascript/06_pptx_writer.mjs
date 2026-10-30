#!/usr/bin/env node
/**
 * 06_pptx_writer — write a deck with a size, titles, text and an image, then read it back.
 *
 * Run: OFFICE_OXIDE_LIB=.../liboffice_oxide.so node examples/javascript/06_pptx_writer.mjs
 */
import { tmpdir } from 'os';
import { join } from 'path';
import { readFileSync, unlinkSync } from 'fs';
import { Document, PptxWriter, createFromMarkdown, XlsxWriter } from '../../js/lib/index.js';

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

console.log('step 1: build a PPTX in memory');
const bytes = buildPptxBytes();

console.log('step 2: open it from bytes');
const doc = Document.fromBytes(bytes, 'pptx');
try {
  const text = doc.plainText();
  const ir = doc.toIr();
  console.log(`  format: ${doc.format}`);
  console.log(`  chars: ${text.length}`);
  console.log(`  slides: ${ir.sections.length}`);

  console.log('step 3: check the slides');
  check(ir.sections.length === 2, 'two slides');
  for (const e of ['Welcome', 'Quarterly Report', 'Results', 'Revenue grew in every region.']) {
    check(text.includes(e), `text contains "${e}"`);
  }
  // The image sits in a positioned text box: it is in the IR and in the
  // image-embedding Markdown, but toHtml() does not render it for PPTX.
  const box = ir.sections[1].elements.find((el) => el.type === 'text_box');
  check(box && box.content.some((c) => c.type === 'image' && c.format === 'png'), 'slide 2 has a png image');
  check(doc.toMarkdownWithImages().includes('[image-base64:'), 'markdown embeds the image');
  // 19200 twips = 12192000 EMU: the presentation size round-trips.
  check(ir.sections[0].page_setup.width_twips === 19200, 'presentation width round-trips');
} finally { doc.close(); }

console.log('step 4: a missing slide is an error');
const w = new PptxWriter();
try {
  w.setSlideTitle(5, 'ghost');
  check(false, 'setSlideTitle on a missing slide throws');
} catch (e) {
  check(e.code === 6, 'error code is 6 (unsupported or out of range)');
  console.log(`  error: ${e.message}`);
} finally { w.close(); }

console.log('OK 06_pptx_writer');
