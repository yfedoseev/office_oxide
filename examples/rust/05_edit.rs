//! 05_edit — Edit an existing document via EditableDocument.
//!
//! Creates a DOCX with placeholder text "Hello {{NAME}}", saves it, opens
//! with EditableDocument, replaces "{{NAME}}" with "World", saves again, and
//! reads back to verify the final text.
//!
//! Run: `cargo run --example 05_edit`

use office_oxide::Document;
use office_oxide::docx::write::DocxWriter;
use office_oxide::edit::EditableDocument;

fn main() {
    let template_path = std::env::temp_dir().join("oo_example_05_template.docx");
    let output_path = std::env::temp_dir().join("oo_example_05_output.docx");

    // ── Create template ─────────────────────────────────────────────────────
    let mut w = DocxWriter::new();
    w.add_heading("Greeting Template", 1);
    w.add_paragraph("Hello {{NAME}}, welcome to Office Oxide!");
    w.add_paragraph("Sincerely, {{SENDER}}");
    w.save(&template_path).expect("save template");

    // ── Edit ────────────────────────────────────────────────────────────────
    let mut ed = EditableDocument::open(&template_path).expect("open editable");
    let n1 = ed.replace_text("{{NAME}}", "World");
    let n2 = ed.replace_text("{{SENDER}}", "The Office Oxide Team");
    ed.save(&output_path).expect("save edited document");

    println!("Replacements: {{NAME}} x{n1}, {{SENDER}} x{n2}");

    // ── Read back and verify ────────────────────────────────────────────────
    let doc = Document::open(&output_path).expect("open output");
    let text = doc.plain_text();

    assert!(text.contains("Hello World"), "replacement failed: expected 'Hello World'");
    assert!(text.contains("The Office Oxide Team"), "replacement failed: expected team name");
    assert!(!text.contains("{{NAME}}"), "placeholder still present");
    assert!(!text.contains("{{SENDER}}"), "placeholder still present");

    println!("Edit verified successfully.");
    println!("--- final text ---");
    println!("{text}");
}
