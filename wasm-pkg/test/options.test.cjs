// The per-document options through the wasm-bindgen glue: `WasmParseOptions`,
// `WasmDocument.openWith` and `damage()`, as JavaScript calls them.
//
//   wasm-pack build --target nodejs --dev --out-dir wasm-pkg/node \
//     --no-default-features --features wasm
//   node --test wasm-pkg/test/smoke.test.cjs wasm-pkg/test/options.test.cjs
'use strict';

const { test } = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const { docxWithBody, xlsxFanOut } = require('./zip.cjs');

const pkg = process.env.OFFICE_OXIDE_WASM_PKG
  ?? path.join(__dirname, '..', 'node', 'office_oxide.js');
const { WasmDocument, WasmParseOptions } = require(pkg);

// Every option, as every binding names it.
const NAMES = [
  'include_hidden_text', 'include_deleted_text', 'on_incomplete', 'lazy_pictures',
  'package_entries', 'package_bytes', 'part_bytes', 'inflate_ratio',
  'embedded_depth', 'embedded_packages', 'text_chars', 'memory_bytes',
];

const paragraphs = (n) => Array.from({ length: n },
  (_, i) => `<w:p><w:r><w:t>paragraph number ${i}</w:t></w:r></w:p>`).join('');
// 100,000 characters of cell text, which a 12,000-character budget cuts.
const fanOut = xlsxFanOut(20);

test('WasmParseOptions names every option, and set takes each', () => {
  assert.deepEqual([...WasmParseOptions.names()].sort(), [...NAMES].sort());
  const options = new WasmParseOptions();
  for (const name of NAMES) {
    if (name === 'on_incomplete') options.set(name, 'report');
    else if (name.startsWith('include_') || name === 'lazy_pictures') options.set(name, true);
    else options.set(name, 1_000_000);
  }
  options.set('part_bytes', Infinity);
  options.set('text_chars', 'max');
  options.free();
});

test('WasmParseOptions refuses bad values and unknown names', () => {
  const options = new WasmParseOptions();
  for (const bad of [-1, 1.5, NaN, 2 ** 64, 'lots', {}, null]) {
    assert.throws(() => options.set('text_chars', bad), /text_chars/, String(bad));
  }
  assert.throws(() => options.set('on_incomplete', 'maybe'), /report/);
  assert.throws(() => options.set('no_such_option', 1), /text_chars/);
  options.free();
});

test('one options object opens many documents, and only those', () => {
  const options = new WasmParseOptions();
  options.set('text_chars', 12000);
  const full = new WasmDocument(fanOut, 'xlsx').plainText().length;
  const first = WasmDocument.openWith(fanOut, 'xlsx', options);
  const second = WasmDocument.openWith(fanOut, 'xlsx', options);
  for (const doc of [first, second]) {
    assert.ok(doc.plainText().length < full / 4, `${doc.plainText().length} of ${full}`);
    const damage = doc.damage();
    assert.ok(damage.some((d) => d.kind === 'truncated'), JSON.stringify(damage));
    assert.ok(damage.every((d) => d.part === null || typeof d.part === 'string'));
    doc.free();
  }
  // A document opened without options reads under the process-wide limits,
  // which the options never touched.
  const plain = new WasmDocument(fanOut, 'xlsx');
  assert.equal(plain.plainText().length, full);
  assert.deepEqual(plain.damage(), []);
  plain.free();
  options.free();
});

test('strict options make an incomplete result throw', () => {
  const strict = new WasmParseOptions(true);
  strict.set('text_chars', 12000);
  assert.throws(() => WasmDocument.openWith(fanOut, 'xlsx', strict), /incomplete/i);
  const sound = WasmDocument.openWith(docxWithBody(paragraphs(3)), 'docx', strict);
  assert.match(sound.plainText(), /paragraph number 2/);
  sound.free();
  strict.free();
});

test('the options report their limits and the bounds derived from them', () => {
  const options = new WasmParseOptions();
  options.set('memory_bytes', 1000);
  options.set('part_bytes', Infinity);
  options.set('package_bytes', 2 ** 60);
  const { limits, derived } = options.limits();
  assert.equal(limits.memory_bytes, 1000);
  assert.equal(limits.part_bytes, Infinity);
  assert.equal(limits.package_bytes, 2n ** 60n);
  assert.equal(derived.write_bytes, 500);
  // A count with no bound reads as Infinity on wasm32, whose usize is 32 bits.
  options.set('text_chars', Infinity);
  options.set('memory_bytes', Infinity);
  const unbounded = options.limits();
  assert.equal(unbounded.limits.text_chars, Infinity);
  assert.equal(unbounded.derived.grid_padding_cells, Infinity);
  assert.equal(unbounded.derived.image_bytes, Infinity);
  assert.deepEqual(Object.keys(limits).sort(),
    NAMES.filter((n) => !n.startsWith('include_') && n !== 'on_incomplete' && n !== 'lazy_pictures').sort());
  options.free();
});

test('a document as UTF-8 bytes is its string form', () => {
  const doc = new WasmDocument(docxWithBody('<w:p><w:r><w:t>café \u{1D74E}</w:t></w:r></w:p>'), 'docx');
  const utf8 = (s) => new Uint8Array(Buffer.from(s, 'utf8'));
  assert.deepEqual(doc.toBuffer('text'), utf8(doc.plainText()));
  assert.deepEqual(doc.toBuffer('markdown'), utf8(doc.toMarkdown()));
  assert.deepEqual(doc.toBuffer('markdown_images'), utf8(doc.toMarkdownWithImages()));
  assert.deepEqual(doc.toBuffer('html'), utf8(doc.toHtml()));
  assert.deepEqual(JSON.parse(Buffer.from(doc.toBuffer('ir_json')).toString('utf8')), doc.toIr());
  assert.throws(() => doc.toBuffer('pdf'), /ir_json/);
  doc.free();
});
