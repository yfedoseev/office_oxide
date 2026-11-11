// Smoke test for the wasm-bindgen glue: loads the Node build and drives
// every exported class, so a glue failure surfaces in CI rather than on
// release day.
//
//   wasm-pack build --target nodejs --dev --out-dir wasm-pkg/node \
//     --no-default-features --features wasm
//   node --test wasm-pkg/test/smoke.test.cjs wasm-pkg/test/options.test.cjs
//
// OFFICE_OXIDE_WASM_PKG overrides where the built module is loaded from.
'use strict';

const { test } = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');

const pkg = process.env.OFFICE_OXIDE_WASM_PKG
  ?? path.join(__dirname, '..', 'node', 'office_oxide.js');
const { WasmDocument, WasmEditableDocument } = require(pkg);

// The smallest DOCX the reader accepts: [Content_Types].xml, the package
// relationship and one paragraph.
const { docxWithBody } = require('./zip.cjs');
const docx = docxWithBody('<w:p><w:r><w:t>Hello wasm</w:t></w:r></w:p>');

test('WasmDocument extracts text, markdown, html and IR', () => {
  const doc = new WasmDocument(docx, 'docx');
  assert.equal(doc.formatName(), 'docx');
  assert.match(doc.plainText(), /Hello wasm/);
  assert.match(doc.toMarkdown(), /Hello wasm/);
  assert.match(doc.toMarkdownWithImages(), /Hello wasm/);
  assert.match(doc.toHtml(), /Hello wasm/);
  assert.ok(Array.isArray(doc.toIr().sections));
  doc.free();
});

test('WasmEditableDocument edits in memory', () => {
  const ed = new WasmEditableDocument(docx, 'docx');
  assert.equal(ed.replaceText('wasm', 'there'), 1);
  assert.throws(() => ed.replaceText('', 'x'), /empty/);
  assert.throws(() => ed.setCell(0, 'A1', 'x'), /XLSX/);
  const out = ed.toBytes();
  ed.free();
  const doc = new WasmDocument(out, 'docx');
  assert.match(doc.plainText(), /Hello there/);
  doc.free();
});
