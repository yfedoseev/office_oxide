#!/usr/bin/env node
/**
 * 08_limits_strict_damage — strict mode, per-document limits and the damage report on truncated bytes.
 *
 * Run: OFFICE_OXIDE_LIB=.../liboffice_oxide.so node examples/javascript/08_limits_strict_damage.mjs
 */
import { tmpdir } from 'os';
import { join } from 'path';
import { readFileSync, unlinkSync } from 'fs';
import { Document, ParseOptions, OfficeOxideError, createFromMarkdown, XlsxWriter, PptxWriter } from '../../js/lib/index.js';

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

function openOrError(bytes, options) {
  try {
    return { doc: Document.fromBytes(bytes, 'docx', options) };
  } catch (e) {
    check(e instanceof OfficeOxideError, 'failure is an OfficeOxideError');
    return { error: e };
  }
}

console.log('step 1: build a DOCX in memory and keep the first half of its bytes');
const good = buildDocxBytes();
const truncated = good.subarray(0, Math.floor(good.length / 2));
console.log(`  option names: ${ParseOptions.names().join(', ')}`);

console.log('step 2: strict mode rejects damaged bytes');
const strict = new ParseOptions({ strict: true });
try {
  const ok = Document.fromBytes(good, 'docx', strict);
  try { check(ok.damage().length === 0, 'a sound document has no damage'); } finally { ok.close(); }
  for (const [name, bytes] of [['truncated', truncated]]) {
    const r = openOrError(bytes, strict);
    if (r.doc) r.doc.close();
    check(r.error, `strict open of ${name} bytes fails`);
    console.log(`  error: ${name}: code ${r.error.code}: ${r.error.message}`);
  }
} finally { strict.close(); }

console.log('step 3: a lenient open reports damage or fails, never looks clean');
for (const [name, bytes] of [['truncated', truncated]]) {
  const r = openOrError(bytes);
  if (r.error) { console.log(`  error: ${name}: code ${r.error.code}: ${r.error.message}`); continue; }
  try {
    const damage = r.doc.damage();
    console.log(`  damage: ${name}: ${damage.map((d) => d.kind).join(', ')}`);
    check(damage.length > 0, `lenient open of ${name} bytes reports damage`);
  } finally { r.doc.close(); }
}

console.log('step 4: a limit applies to one document only');
const capped = new ParseOptions({ package_bytes: 100 });
try {
  const r = openOrError(good, capped);
  if (r.doc) r.doc.close();
  check(r.error, 'package_bytes=100 rejects the document');
  console.log(`  error: capped: code ${r.error.code}: ${r.error.message}`);
} finally { capped.close(); }
Document.fromBytes(good, 'docx').close(); // default options still open it

console.log('OK 08_limits_strict_damage');
