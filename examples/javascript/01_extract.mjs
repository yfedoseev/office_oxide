#!/usr/bin/env node
/**
 * 01_extract — Self-contained extract demo using the JS (koffi) bindings.
 *
 * Creates a DOCX from Markdown (written to a temp file), opens it with
 * Document.open, and prints the format, plain text, and Markdown output.
 * Exit 0 on success.
 *
 * Run: OFFICE_OXIDE_LIB=.../liboffice_oxide.so node examples/javascript/01_extract.mjs
 */
import { tmpdir } from 'os';
import { join } from 'path';
import { unlinkSync } from 'fs';
import { Document, createFromMarkdown } from '../../js/lib/index.js';

const markdown = `# Office Oxide Extract Demo

This document was created from Markdown and parsed back via JavaScript bindings.

## Features

- Plain text extraction
- Markdown conversion
- IR (intermediate representation) access
`;

const tmpPath = join(tmpdir(), `oo_01_extract_${Date.now()}.docx`);

try {
  createFromMarkdown(markdown, 'docx', tmpPath);

  const doc = Document.open(tmpPath);
  try {
    const format = doc.format;
    const text = doc.plainText();
    const md = doc.toMarkdown();
    const ir = doc.toIr();

    console.log('format:', format);
    console.log('--- plain text (first 200 chars) ---');
    console.log(text.slice(0, 200));
    console.log(`--- markdown length: ${md.length} chars ---`);
    console.log(`--- IR sections: ${ir.sections?.length ?? 0} ---`);

    if (format !== 'docx') throw new Error(`expected format=docx, got ${format}`);
    if (!text) throw new Error('plain text is empty');
    if (!text.includes('Office Oxide Extract Demo')) throw new Error('heading missing');

    console.log('\nAll checks passed.');
  } finally {
    doc.close();
  }
} finally {
  try { unlinkSync(tmpPath); } catch (_) {}
}
