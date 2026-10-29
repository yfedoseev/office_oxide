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
