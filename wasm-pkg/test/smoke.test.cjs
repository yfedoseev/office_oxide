// Smoke test for the wasm-bindgen glue: loads the Node build and drives
// every exported class, so a glue failure surfaces in CI rather than on
// release day.
//
//   wasm-pack build --target nodejs --dev --out-dir wasm-pkg/node \
//     --no-default-features --features wasm
//   node --test wasm-pkg/test/
//
// OFFICE_OXIDE_WASM_PKG overrides where the built module is loaded from.
'use strict';

const { test } = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');

const pkg = process.env.OFFICE_OXIDE_WASM_PKG
  ?? path.join(__dirname, '..', 'node', 'office_oxide.js');
const { WasmDocument, WasmEditableDocument } = require(pkg);

// The smallest DOCX the reader accepts, built by hand so the test needs no
// fixture: [Content_Types].xml, the package relationship and one paragraph.
function storedZip(files) {
  const crcTable = new Uint32Array(256).map((_, n) => {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    return c >>> 0;
  });
  const crc32 = (buf) => {
    let c = 0xffffffff;
    for (const b of buf) c = crcTable[(c ^ b) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  };
  const locals = [];
  const centrals = [];
  let offset = 0;
  for (const [name, text] of files) {
    const nameBuf = Buffer.from(name);
    const data = Buffer.from(text);
    const crc = crc32(data);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(data.length, 18);
    local.writeUInt32LE(data.length, 22);
    local.writeUInt16LE(nameBuf.length, 26);
    locals.push(local, nameBuf, data);
    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(data.length, 20);
    central.writeUInt32LE(data.length, 24);
    central.writeUInt16LE(nameBuf.length, 28);
    central.writeUInt32LE(offset, 42);
    centrals.push(central, nameBuf);
    offset += 30 + nameBuf.length + data.length;
  }
  const cd = Buffer.concat(centrals);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(files.length, 8);
  end.writeUInt16LE(files.length, 10);
  end.writeUInt32LE(cd.length, 12);
  end.writeUInt32LE(offset, 16);
  return new Uint8Array(Buffer.concat([...locals, cd, end]));
}

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const docx = storedZip([
  ['[Content_Types].xml',
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    + '</Types>'],
  ['_rels/.rels',
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    + '</Relationships>'],
  ['word/document.xml',
    `<?xml version="1.0"?><w:document xmlns:w="${W}"><w:body>`
    + '<w:p><w:r><w:t>Hello wasm</w:t></w:r></w:p></w:body></w:document>'],
]);

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
