#!/usr/bin/env node
/**
 * 02_create_from_markdown — Convert Markdown to DOCX, XLSX, and PPTX.
 *
 * For each format: calls createFromMarkdown, opens with Document.open,
 * prints a summary. Exit 0 on success.
 *
 * Run: OFFICE_OXIDE_LIB=.../liboffice_oxide.so node examples/javascript/02_create_from_markdown.mjs
 */
import { tmpdir } from 'os';
import { join } from 'path';
import { unlinkSync } from 'fs';
import { Document, createFromMarkdown } from '../../js/lib/index.js';

const markdown = `# Quarterly Report

Generated automatically from Markdown using Office Oxide.

## Highlights

- Revenue grew by **32%** year-over-year
- Customer satisfaction: 4.8 / 5.0
- New products launched: Widget Pro, Widget Lite

## Financial Summary

| Category   | Q3 2025 | Q4 2025 |
|------------|---------|---------|
| Revenue    | $1.2M   | $1.6M   |
| Expenses   | $0.8M   | $0.9M   |
| Net Profit | $0.4M   | $0.7M   |
`;

const formats = ['docx', 'xlsx', 'pptx'];
const tmps = {};

try {
  for (const fmt of formats) {
    tmps[fmt] = join(tmpdir(), `oo_02_create_${Date.now()}_${fmt}.${fmt}`);
  }

  for (const fmt of formats) {
    const path = tmps[fmt];
    createFromMarkdown(markdown, fmt, path);

    const doc = Document.open(path);
    try {
      const text = doc.plainText();
      const md = doc.toMarkdown();

      if (!text) throw new Error(`${fmt}: plain text is empty`);

      console.log(`\n=== ${fmt.toUpperCase()} ===`);
      console.log(`plain text length: ${text.length} chars`);
      console.log(`markdown length:   ${md.length} chars`);
      console.log(`first 100 chars: ${JSON.stringify(text.slice(0, 100))}`);
    } finally {
      doc.close();
    }
  }

  console.log('\nAll formats created and verified.');
} finally {
  for (const p of Object.values(tmps)) {
    try { unlinkSync(p); } catch (_) {}
  }
}
