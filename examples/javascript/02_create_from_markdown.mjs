#!/usr/bin/env node
/**
 * 02_create_from_markdown: turn one Markdown source into a DOCX, an XLSX and a PPTX.
 *
 * Run: node examples/javascript/02_create_from_markdown.mjs
 */
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Document, createFromMarkdown } from '../../js/lib/index.js';

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
  console.log('step 1: write the Markdown source');

  console.log('step 2: create a DOCX, an XLSX and a PPTX from it');
  const created = {};
  for (const format of ['docx', 'xlsx', 'pptx']) created[format] = markdownToBytes(MARKDOWN, format);

  console.log('step 3: open each and check the text');
  for (const [format, data] of Object.entries(created)) {
    const doc = Document.fromBytes(data, format);
    try {
      const text = doc.plainText();
      console.log(`  format: ${format}`);
      console.log(`  chars: ${text.length}`);
      has(text, 'North');
    } finally {
      doc.close();
    }
  }

  console.log('OK 02_create_from_markdown');
}

main();
