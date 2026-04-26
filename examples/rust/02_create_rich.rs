//! 02_create_rich — Rich DOCX creation demo.
//!
//! Creates a DOCX with a heading, rich paragraph (bold + italic + colored runs),
//! a centred paragraph, a page break, and a table. Saves to a temp file and
//! reads it back to verify the content.
//!
//! Run: `cargo run --example 02_create_rich`

use office_oxide::Document;
use office_oxide::docx::write::{Alignment, DocxWriter, Run};
use office_oxide::format::DocumentFormat;

fn main() {
    let out = std::env::temp_dir().join("oo_example_02.docx");

    // ── Build the document ──────────────────────────────────────────────────
    let mut w = DocxWriter::new();

    w.add_heading("Rich Document Demo", 1);

    w.add_rich_paragraph(&[
        Run::new("Bold text").bold(),
        Run::new(", "),
        Run::new("italic text").italic(),
        Run::new(", and "),
        Run::new("red colored text").color("FF0000"),
        Run::new("."),
    ]);

    w.add_paragraph_aligned("This paragraph is centred.", Alignment::Center);
    w.add_paragraph_aligned("This one is right-aligned.", Alignment::Right);

    w.add_page_break();

    w.add_heading("After Page Break", 2);

    w.add_table(&[
        vec!["Product", "Price", "Qty"],
        vec!["Widget A", "$10.00", "100"],
        vec!["Widget B", "$25.00", "50"],
        vec!["Total", "$2750.00", "150"],
    ]);

    // ── Save ────────────────────────────────────────────────────────────────
    w.save(&out).expect("save DOCX");

    // ── Read back and verify ────────────────────────────────────────────────
    let doc = Document::open(&out).expect("open saved DOCX");
    assert_eq!(doc.format(), DocumentFormat::Docx);

    let text = doc.plain_text();
    assert!(text.contains("Rich Document Demo"), "heading missing");
    assert!(text.contains("Bold text"), "bold run missing");
    assert!(text.contains("italic text"), "italic run missing");
    assert!(text.contains("red colored text"), "colored run missing");
    assert!(text.contains("centred"), "aligned paragraph missing");
    assert!(text.contains("Widget A"), "table content missing");
    assert!(text.contains("After Page Break"), "content after page break missing");

    println!("DOCX created and verified: {}", out.display());
    println!("Plain text snippet: {}", text.lines().next().unwrap_or(""));
}
