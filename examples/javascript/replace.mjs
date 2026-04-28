#!/usr/bin/env node
import { EditableDocument } from 'office-oxide';

if (process.argv.length !== 4) {
  console.error('usage: replace.mjs <template> <output>');
  process.exit(1);
}

const ed = EditableDocument.open(process.argv[2]);
try {
  let n = ed.replaceText('{{NAME}}', 'Alice');
  n += ed.replaceText('{{DATE}}', '2026-04-18');
  console.log('replacements:', n);
  ed.save(process.argv[3]);
} finally {
  ed.close();
}
