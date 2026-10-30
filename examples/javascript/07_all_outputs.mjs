#!/usr/bin/env node
/**
 * 07_all_outputs: render one document as plain text, Markdown, HTML and IR JSON.
 *
 * Run: node examples/javascript/07_all_outputs.mjs
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
  let outputs;
  try {
    console.log(`  format: ${doc.format}`);

    console.log('step 3: render every output');
    outputs = {
      plain: doc.plainText(), // words only: no markup, no structure
      markdown: doc.toMarkdown(), // headings and pipe tables survive as syntax
      html: doc.toHtml(), // a fragment (no <html> wrapper), tables as <table>
      ir_json: JSON.stringify(doc.toIr()), // the format-neutral tree as JSON, for programs
    };
  } finally {
    doc.close();
  }
  for (const [name, value] of Object.entries(outputs)) console.log(`  chars: ${name} ${value.length}`);

  console.log('step 4: check each output');
  for (const [name, value] of Object.entries(outputs)) check(value.trim() !== '', `the ${name} output is empty`);
  has(outputs.plain, 'Quarterly Report');
  check(outputs.markdown.includes('# Quarterly Report'), 'the Markdown is missing the heading');
  check(outputs.html.includes('<table'), 'the HTML has no table');
  const tree = JSON.parse(outputs.ir_json);
  check(typeof tree === 'object' && tree !== null && !Array.isArray(tree), 'the IR JSON does not parse to an object');

  console.log('OK 07_all_outputs');
}

main();
