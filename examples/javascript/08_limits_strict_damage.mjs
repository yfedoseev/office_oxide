#!/usr/bin/env node
/**
 * 08_limits_strict_damage: strict mode refuses a damaged archive; a lenient open reports the damage.
 *
 * Run: node examples/javascript/08_limits_strict_damage.mjs
 */
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Document, OfficeOxideError, ParseOptions, createFromMarkdown } from '../../js/lib/index.js';

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

const PART = 'word/_rels/document.xml.rels';

// Overwrites the start of one part's compressed data. A first deflate byte of
// 0xFF declares a reserved block type, so every inflater rejects the part, and
// the central directory stays intact. The part is found by name, not by
// offset, so the recipe does not depend on how big the other parts are.
function damagePart(data) {
  const at = data.indexOf(PART); // the first hit is the part's local header
  check(at > 0, 'the part to damage is missing from the archive');
  // The extra-field length sits in the two bytes just before the name.
  const start = at + PART.length + data.readUInt16LE(at - 2);
  const damaged = Buffer.from(data);
  damaged.fill(0xff, start, start + 16);
  return damaged;
}

// Opens bytes as a DOCX; returns the error, or null when the open succeeded.
function openError(data, options) {
  try {
    Document.fromBytes(data, 'docx', options).close();
    return null;
  } catch (e) {
    if (!(e instanceof OfficeOxideError)) throw e;
    return e;
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
  console.log('step 1: build a DOCX in memory');
  const good = buildDocxBytes();

  console.log(`step 2: damage one part: overwrite 16 bytes of ${PART}`);
  const damaged = damagePart(good);

  console.log('step 3: open it in strict mode');
  const strict = new ParseOptions({ strict: true });
  try {
    const error = openError(damaged, strict);
    check(error !== null, 'strict mode accepted a damaged part');
    console.log(`  error: ${error.message}`);

    console.log('step 4: open it leniently');
    const doc = Document.fromBytes(damaged, 'docx');
    let kinds, text;
    try {
      kinds = doc.damage().map((d) => d.kind);
      text = doc.plainText();
    } finally {
      doc.close();
    }
    console.log(`  damage: ${kinds.join(', ')}`);
    check(kinds.includes('unreadable_part'), 'the damage report does not name an unreadable part');
    has(text, 'Quarterly Report');

    console.log('step 5: truncate the archive to its first half and open it both ways');
    const half = good.subarray(0, Math.floor(good.length / 2));
    for (const [options, message] of [
      [strict, 'strict mode accepted a truncated archive'],
      [undefined, 'a lenient open accepted a truncated archive'],
    ]) {
      const error = openError(half, options);
      check(error !== null, message);
      console.log(`  error: ${error.message}`);
    }
  } finally {
    strict.close();
  }

  console.log('step 6: a limit applies to one document only');
  const limited = new ParseOptions({ part_bytes: 100 });
  try {
    const error = openError(good, limited);
    check(error !== null, 'the part_bytes limit was not applied');
    console.log(`  error: ${error.message}`);
  } finally {
    limited.close();
  }
  const doc = Document.fromBytes(good, 'docx'); // the default limits still accept it
  try {
    has(doc.plainText(), 'Quarterly Report');
  } finally {
    doc.close();
  }

  console.log('OK 08_limits_strict_damage');
}

main();
