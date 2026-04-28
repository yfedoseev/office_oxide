import { test } from 'node:test';
import { strict as assert } from 'node:assert';
import { existsSync, readFileSync } from 'node:fs';
import {
  Document,
  EditableDocument,
  detectFormat,
  extractText,
  version,
} from '../lib/index.js';

const fixture = '/tmp/ffi_smoke.docx';

test('version is non-empty', () => {
  const v = version();
  assert.ok(v.length > 0);
  console.log('office_oxide version:', v);
});

test('detectFormat', () => {
  assert.equal(detectFormat('x.docx'), 'docx');
  assert.equal(detectFormat('x.unknown'), null);
});

test('Document.open + plainText + toMarkdown + toIr', { skip: !existsSync(fixture) }, () => {
  const doc = Document.open(fixture);
  try {
    assert.equal(doc.format, 'docx');
    const txt = doc.plainText();
    assert.ok(txt.includes('Hello'));
    const md = doc.toMarkdown();
    assert.ok(md.includes('# '));
    const ir = doc.toIr();
    assert.ok(ir && 'sections' in ir);
  } finally { doc.close(); }
});

test('Document.fromBytes', { skip: !existsSync(fixture) }, () => {
  const bytes = readFileSync(fixture);
  const doc = Document.fromBytes(bytes, 'docx');
  try {
    assert.ok(doc.plainText().length > 0);
  } finally { doc.close(); }
});

test('EditableDocument.replaceText round-trip', { skip: !existsSync(fixture) }, () => {
  const ed = EditableDocument.open(fixture);
  try {
    const n = ed.replaceText('Hello', 'Howdy');
    assert.ok(n >= 1);
    ed.save('/tmp/ffi_smoke_node_edit.docx');
  } finally { ed.close(); }
  const out = extractText('/tmp/ffi_smoke_node_edit.docx');
  assert.ok(out.includes('Howdy'));
});
