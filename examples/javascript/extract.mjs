#!/usr/bin/env node
import { Document } from 'office-oxide';

if (process.argv.length !== 3) {
  console.error('usage: extract.mjs <file>');
  process.exit(1);
}

const doc = Document.open(process.argv[2]);
try {
  console.log('format:', doc.format);
  console.log('--- plain text ---');
  console.log(doc.plainText());
  console.log('--- markdown ---');
  console.log(doc.toMarkdown());
} finally {
  doc.close();
}
