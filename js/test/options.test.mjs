// Per-document options: limits, hidden text, and refusing an incomplete result.
// Self-contained: every document is built here.
import { test } from 'node:test';
import { strict as assert } from 'node:assert';
import { createRequire } from 'node:module';
import { Document, OfficeOxideError, ParseOptions, XlsxWriter } from '../lib/index.js';

// 200 cells of 200 characters: 40,000 characters of text from one small sheet.
function bigSheet() {
  const w = new XlsxWriter();
  w.addSheet('S');
  for (let i = 0; i < 200; i++) w.setCell(0, i, 0, 'x'.repeat(200));
  const bytes = w.toBytes();
  w.close();
  return bytes;
}

test('a limit in the options applies to that document only', () => {
  const bytes = bigSheet();
  const opts = new ParseOptions({ text_chars: 5000 });
  const small = Document.fromBytes(bytes, 'xlsx', opts);
  const plain = Document.fromBytes(bytes, 'xlsx');
  assert.ok(small.plainText().length < plain.plainText().length);
  small.close(); plain.close(); opts.close();
  // The process-wide limit, which new options start from, is untouched.
  const fresh = new ParseOptions();
  assert.equal(fresh.limits().limits.text_chars, 256 * 2 ** 20);
  fresh.close();
});

// Every option, as every binding names it.
const OPTION_NAMES = [
  'include_hidden_text', 'include_deleted_text', 'on_incomplete', 'lazy_pictures',
  'package_entries', 'package_bytes', 'part_bytes', 'inflate_ratio',
  'embedded_depth', 'embedded_packages', 'text_chars', 'memory_bytes',
];

test('the options name every option', () => {
  assert.deepEqual([...ParseOptions.names()].sort(), [...OPTION_NAMES].sort());
  for (const name of OPTION_NAMES) {
    const opts = new ParseOptions({ [name]: name === 'on_incomplete' ? 'report' : 1 });
    opts.close();
  }
});

test('damage lists a cut and strict options refuse it with code 7', () => {
  const bytes = bigSheet();
  const lenient = new ParseOptions({ text_chars: 5000 });
  const doc = Document.fromBytes(bytes, 'xlsx', lenient);
  assert.ok(doc.damage().some((d) => d.kind === 'truncated'), JSON.stringify(doc.damage()));
  doc.close(); lenient.close();

  const strict = new ParseOptions({ text_chars: 5000, on_incomplete: 'error' });
  assert.throws(
    () => Document.fromBytes(bytes, 'xlsx', strict),
    (e) => e instanceof OfficeOxideError && e.code === 7,
  );
  strict.close();
});

test('a sound document is accepted by strict options and has nothing missing', () => {
  const w = new XlsxWriter();
  w.addSheet('S');
  w.setCell(0, 0, 0, 'hello');
  const bytes = w.toBytes();
  w.close();
  const strict = new ParseOptions({ strict: true });
  const doc = Document.fromBytes(bytes, 'xlsx', strict);
  assert.ok(doc.plainText().includes('hello'));
  assert.deepEqual(doc.damage().filter((d) => d.kind !== 'recovered'), []);
  doc.close(); strict.close();
});

test('unknown names and bad values are refused', () => {
  assert.throws(() => new ParseOptions({ no_such_limit: 1 }), OfficeOxideError);
  assert.throws(() => new ParseOptions({ text_chars: 'lots' }), TypeError);
  assert.throws(() => new ParseOptions({ text_chars: -1 }), TypeError);
  assert.ok(ParseOptions.names().includes('part_bytes'));
});

test('the CommonJS entry point exports ParseOptions too', () => {
  const cjs = createRequire(import.meta.url)('../lib/index.cjs');
  assert.equal(typeof cjs.ParseOptions, 'function');
  const opts = new cjs.ParseOptions({ text_chars: 10 });
  opts.close();
});

test('the options report their limits and the bounds derived from them', () => {
  const opts = new ParseOptions({ memory_bytes: 1000, part_bytes: Infinity });
  try {
    const { limits, derived } = opts.limits();
    assert.equal(limits.memory_bytes, 1000);
    assert.equal(limits.part_bytes, Infinity);
    assert.equal(derived.write_bytes, 500);
    for (const name of ParseOptions.names().filter((n) => n in limits)) assert.ok(name in limits);
    assert.ok(Object.keys(limits).length >= 8);
  } finally {
    opts.close();
  }
  const big = new ParseOptions({ memory_bytes: 2n ** 62n });
  try {
    assert.equal(big.limits().limits.memory_bytes, 2n ** 62n, 'a large limit is exact');
  } finally {
    big.close();
  }
});

test('a document as a Buffer or a file is its string form in UTF-8', async () => {
  const { mkdtempSync, readFileSync, rmSync } = await import('node:fs');
  const { tmpdir } = await import('node:os');
  const { join } = await import('node:path');
  const w = new XlsxWriter();
  w.addSheet('S');
  w.setCell(0, 0, 0, 'omega \u{1D74E} and é');
  const bytes = w.toBytes();
  w.close();
  const doc = Document.fromBytes(bytes, 'xlsx');
  const dir = mkdtempSync(join(tmpdir(), 'oo-out-'));
  try {
    const forms = {
      text: doc.plainText(),
      markdown: doc.toMarkdown(),
      markdown_images: doc.toMarkdownWithImages(),
      html: doc.toHtml(),
    };
    for (const [kind, str] of Object.entries(forms)) {
      assert.ok(str.includes('\u{1D74E}'), kind);
      assert.deepEqual(doc.toBuffer(kind), Buffer.from(str, 'utf8'), kind);
      const p = join(dir, kind);
      doc.writeTo(p, kind);
      assert.deepEqual(readFileSync(p), Buffer.from(str, 'utf8'), kind);
    }
    assert.deepEqual(JSON.parse(doc.toBuffer('ir_json').toString('utf8')), doc.toIr());
    assert.throws(() => doc.toBuffer('pdf'), OfficeOxideError);
    assert.throws(() => doc.writeTo(join(dir, 'no', 'such', 'dir'), 'text'), OfficeOxideError);
  } finally {
    doc.close();
    rmSync(dir, { recursive: true, force: true });
  }
});
