#!/usr/bin/env node
/**
 * 01_extract: read a document as plain text, Markdown and IR JSON.
 *
 * Run: node examples/javascript/01_extract.mjs
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

function buildDocxBytes(title = 'Quarterly Report') {
  return markdownToBytes(MARKDOWN.replaceAll('Quarterly Report', title), 'docx');
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
  console.log('step 1: build a DOCX in memory');
  const data = buildDocxBytes();

  console.log('step 2: open it from bytes');
  const doc = Document.fromBytes(data, 'docx');
  try {
    console.log(`  format: ${doc.format}`);
    check(doc.format === 'docx', 'the opened format differs from the requested one');

    console.log('step 3: extract plain text, Markdown and IR JSON');
    const text = doc.plainText();
    const markdown = doc.toMarkdown();
    const irJson = JSON.stringify(doc.toIr());
    console.log(`  chars: ${text.length}`);

    console.log('step 4: check the output');
    has(text, 'Revenue grew in every region.');
    has(markdown, 'Quarterly Report');
    check(typeof JSON.parse(irJson) === 'object', 'the IR JSON is not an object');
  } finally {
    doc.close();
  }

  console.log('OK 01_extract');
}

main();
