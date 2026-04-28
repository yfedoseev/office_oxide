#!/usr/bin/env node
/**
 * 03_edit — Edit a document by replacing placeholder text.
 *
 * Creates a DOCX from Markdown with {{NAME}} and {{DATE}} placeholders,
 * opens with EditableDocument, replaces both, saves to a second temp file,
 * then reads back and verifies. Exit 0 on success.
 *
 * Run: OFFICE_OXIDE_LIB=.../liboffice_oxide.so node examples/javascript/03_edit.mjs
 */
import { tmpdir } from 'os';
import { join } from 'path';
import { unlinkSync } from 'fs';
import { Document, EditableDocument, createFromMarkdown } from '../../js/lib/index.js';

const markdown = `# Invoice

Dear {{NAME}},

Please find attached your invoice for services rendered on {{DATE}}.

## Summary

- Service: Office Oxide Pro License
- Amount: $499.00
- Due date: 30 days from {{DATE}}

Thank you for your business!
`;

const ts = Date.now();
const templatePath = join(tmpdir(), `oo_03_template_${ts}.docx`);
const outputPath = join(tmpdir(), `oo_03_output_${ts}.docx`);

try {
  // Create template
  createFromMarkdown(markdown, 'docx', templatePath);

  // Edit
  const ed = EditableDocument.open(templatePath);
  let n1, n2;
  try {
    n1 = ed.replaceText('{{NAME}}', 'Alice Smith');
    n2 = ed.replaceText('{{DATE}}', '2026-04-26');
    ed.save(outputPath);
  } finally {
    ed.close();
  }
  console.log(`Replacements: {{NAME}} x${n1}, {{DATE}} x${n2}`);

  // Verify
  const doc = Document.open(outputPath);
  try {
    const text = doc.plainText();

    if (!text.includes('Alice Smith')) throw new Error("name replacement failed: 'Alice Smith' not found");
    if (!text.includes('2026-04-26')) throw new Error("date replacement failed: '2026-04-26' not found");
    if (text.includes('{{NAME}}')) throw new Error('placeholder {{NAME}} still present');
    if (text.includes('{{DATE}}')) throw new Error('placeholder {{DATE}} still present');

    console.log('Edit verified successfully.');
    console.log('--- final text ---');
    console.log(text);
  } finally {
    doc.close();
  }
} finally {
  try { unlinkSync(templatePath); } catch (_) {}
  try { unlinkSync(outputPath); } catch (_) {}
}
