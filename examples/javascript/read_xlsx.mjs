#!/usr/bin/env node
import { Document } from 'office-oxide';

if (process.argv.length !== 3) {
  console.error('usage: read_xlsx.mjs <file.xlsx>');
  process.exit(1);
}

const doc = Document.open(process.argv[2]);
try {
  const ir = doc.toIr();
  for (let i = 0; i < ir.sections.length; i++) {
    const s = ir.sections[i];
    console.log(`# sheet ${i}: ${s.title ?? ''}`);
    for (const el of s.elements) {
      if (el.type !== 'table') continue;
      for (const row of el.rows) {
        console.log(row.cells.map(c => c.text ?? '').join('\t'));
      }
    }
  }
} finally {
  doc.close();
}
