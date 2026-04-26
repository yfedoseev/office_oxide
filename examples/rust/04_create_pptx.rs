//! 04_create_pptx — PPTX creation with slides, bullet lists, and text boxes.
//!
//! Creates a 2-slide presentation: slide 1 has a title, bullet list, and rich
//! text (bold + colored runs); slide 2 has a positioned free-floating text box.
//! Saves to a temp file and reads back to verify content.
//!
//! Run: `cargo run --example 04_create_pptx`

use office_oxide::Document;
use office_oxide::pptx::write::{PptxWriter, Run};

fn main() {
    let out = std::env::temp_dir().join("oo_example_04.pptx");

    // ── Build the presentation ──────────────────────────────────────────────
    let mut writer = PptxWriter::new();

    {
        let slide = writer.add_slide();
        slide.set_title("Office Oxide PPTX Demo");
        slide.add_text("This presentation was created with the Rust native API.");
        slide.add_bullet_list(&[
            "Fast: parse and create in microseconds",
            "Safe: pure Rust with no unsafe C++ dependencies",
            "Complete: DOCX, XLSX, PPTX, DOC, XLS, PPT",
        ]);
        slide.add_rich_text(&[
            Run::new("Key feature: ").bold(),
            Run::new("styled runs").italic().color("0070C0"),
            Run::new(" in every paragraph."),
        ]);
    }

    {
        let slide = writer.add_slide();
        slide.set_title("Free-floating Text Boxes");
        slide.add_text("The body placeholder holds regular text.");
        // Positioned text box (x, y, cx, cy in EMUs; 914400 EMU = 1 inch)
        slide.add_text_box("This is a text box at 1in × 4in", 914_400, 3_657_600, 5_486_400, 685_800);
        slide.add_rich_text_box(
            &[
                Run::new("Rich ").bold(),
                Run::new("text box").italic().color("C00000"),
            ],
            914_400, 4_500_000, 5_486_400, 685_800,
        );
    }

    // ── Save ────────────────────────────────────────────────────────────────
    writer.save(&out).expect("save PPTX");

    // ── Read back and verify ────────────────────────────────────────────────
    let doc = Document::open(&out).expect("open saved PPTX");
    let text = doc.plain_text();

    assert!(text.contains("Office Oxide PPTX Demo"), "slide 1 title missing");
    assert!(text.contains("Fast:"), "bullet list missing");
    assert!(text.contains("styled runs"), "rich text run missing");
    assert!(text.contains("Free-floating Text Boxes"), "slide 2 title missing");
    assert!(text.contains("text box"), "text box content missing");

    println!("PPTX created and verified: {}", out.display());
    println!("--- plain text ---");
    println!("{text}");
}
