//! Round-trip idempotence: reading an IR back after writing it must return the
//! same IR, so "read to IR, edit a field, write back" does not silently drift.
//!
//! The property tested is a *fixed point*: with IR1 = parse(write(ir0)) and
//! IR2 = parse(write(IR1)), IR1 must equal IR2 (and plain-text length must be
//! preserved). DOCX previously grew a duplicate title heading each cycle and
//! PPTX dropped its slide-body text; both are covered here.

use office_oxide::{Document, DocumentFormat, DocumentIR, create};
use std::io::Cursor;

fn write_parse(ir: &DocumentIR, fmt: DocumentFormat) -> (DocumentIR, usize) {
    let mut buf = Cursor::new(Vec::new());
    create::create_from_ir_to_writer(ir, fmt, &mut buf).expect("write");
    buf.set_position(0);
    let doc = Document::from_reader(buf, fmt).expect("parse");
    (doc.to_ir(), doc.plain_text().split_whitespace().count())
}

fn assert_idempotent(fmt: DocumentFormat, md: &str) {
    let ir0 = DocumentIR::from_markdown(md, fmt);
    let (ir1, w1) = write_parse(&ir0, fmt);
    let (ir2, w2) = write_parse(&ir1, fmt);
    assert_eq!(
        ir1, ir2,
        "{fmt:?}: write→parse is not idempotent — a second cycle changed the IR"
    );
    assert_eq!(
        w1, w2,
        "{fmt:?}: plain-text word count drifted across a round-trip ({w1} vs {w2})"
    );
    assert!(w1 > 0, "{fmt:?}: round-trip produced empty text");
}

const DOCX_MD: &str = "# Title\n\n## Heading\n\nA **bold** paragraph with text.\n\n- one\n- two\n\n| A | B |\n|---|---|\n| 1 | 2 |\n";
const XLSX_MD: &str = "# Sheet1\n\n| Item | Qty |\n|------|-----|\n| Apple | 10 |\n| Pear | 5 |\n";
const PPTX_MD: &str = "# Slide One\n\n- Bullet A\n- Bullet B\n\n# Slide Two\n\nBody text here.\n";

#[test]
fn test_docx_ir_roundtrip_is_idempotent() {
    assert_idempotent(DocumentFormat::Docx, DOCX_MD);
}

#[test]
fn test_xlsx_ir_roundtrip_is_idempotent() {
    assert_idempotent(DocumentFormat::Xlsx, XLSX_MD);
}

#[test]
fn test_pptx_ir_roundtrip_is_idempotent() {
    assert_idempotent(DocumentFormat::Pptx, PPTX_MD);
}

/// Focused guard for the PPTX body-text-loss defect: the slide body content
/// (bullets) must survive a round-trip, not be replaced by the title.
#[test]
fn test_pptx_roundtrip_preserves_slide_body() {
    let ir0 = DocumentIR::from_markdown(PPTX_MD, DocumentFormat::Pptx);
    let (ir1, _) = write_parse(&ir0, DocumentFormat::Pptx);
    let text = {
        let mut buf = Cursor::new(Vec::new());
        create::create_from_ir_to_writer(&ir1, DocumentFormat::Pptx, &mut buf).unwrap();
        buf.set_position(0);
        Document::from_reader(buf, DocumentFormat::Pptx)
            .unwrap()
            .plain_text()
    };
    for needle in ["Bullet A", "Bullet B", "Body text here"] {
        assert!(text.contains(needle), "PPTX round-trip dropped body text {needle:?}: {text:?}");
    }
}

/// `section.title` must not be re-emitted as a duplicate
/// heading when the heading it came from isn't the section's literal
/// first element (e.g. a byline or date line ahead of it). The old
/// check only looked at `section.elements.first()`.
#[test]
fn test_docx_heading_not_first_element_does_not_duplicate_on_roundtrip() {
    use office_oxide::ir::{
        Element, Heading, InlineContent, Metadata, Paragraph, Section, TextSpan,
    };

    let ir = DocumentIR {
        metadata: Metadata {
            format: DocumentFormat::Docx,
            ..Default::default()
        },
        sections: vec![Section {
            title: Some("My Heading".to_string()),
            elements: vec![
                Element::Paragraph(Paragraph {
                    content: vec![InlineContent::Text(TextSpan::plain(
                        "A byline before the heading",
                    ))],
                    ..Default::default()
                }),
                Element::Heading(Heading {
                    level: 1,
                    content: vec![InlineContent::Text(TextSpan::plain("My Heading"))],
                    ..Default::default()
                }),
            ],
            ..Default::default()
        }],
        defined_names: Vec::new(),
    };

    let mut buf = Cursor::new(Vec::new());
    create::create_from_ir_to_writer(&ir, DocumentFormat::Docx, &mut buf).unwrap();
    buf.set_position(0);
    let text = Document::from_reader(buf, DocumentFormat::Docx)
        .unwrap()
        .plain_text();

    let occurrences = text.matches("My Heading").count();
    assert_eq!(
        occurrences, 1,
        "heading duplicated on write, expected exactly one occurrence: {text:?}"
    );
}

/// Independent text boxes on one slide must not merge into
/// one (or spawn a spurious empty one) on a write→reread round trip;
/// each keeps its own position.
#[test]
fn test_pptx_independent_text_boxes_stay_independent_on_roundtrip() {
    use office_oxide::ir::{
        Element, InlineContent, Metadata, Paragraph, Section, TextBox, TextSpan,
    };

    fn textbox(text: &str, x: i64) -> Element {
        Element::TextBox(TextBox {
            content: vec![Element::Paragraph(Paragraph {
                content: vec![InlineContent::Text(TextSpan::plain(text))],
                ..Default::default()
            })],
            x_emu: Some(x),
            y_emu: Some(500_000),
            width_emu: Some(2_000_000),
            height_emu: Some(500_000),
            ..Default::default()
        })
    }

    let ir = DocumentIR {
        metadata: Metadata {
            format: DocumentFormat::Pptx,
            ..Default::default()
        },
        sections: vec![Section {
            elements: vec![
                textbox("Text Box", 0),
                textbox("Aspose.Slides for .NET", 2_500_000),
                textbox("Welcome", 5_000_000),
            ],
            ..Default::default()
        }],
        defined_names: Vec::new(),
    };

    let (ir1, _) = write_parse(&ir, DocumentFormat::Pptx);
    let box_count = ir1.sections[0]
        .elements
        .iter()
        .filter(|e| matches!(e, Element::TextBox(_)))
        .count();
    assert_eq!(
        box_count, 3,
        "expected 3 independent text boxes, got {box_count}: {:?}",
        ir1.sections[0]
    );

    let texts: Vec<String> = ir1.sections[0]
        .elements
        .iter()
        .filter_map(|e| match e {
            Element::TextBox(tb) => tb.content.iter().find_map(|inner| match inner {
                Element::Paragraph(p) => Some(
                    p.content
                        .iter()
                        .filter_map(|c| match c {
                            InlineContent::Text(t) => Some(t.text.clone()),
                            _ => None,
                        })
                        .collect::<String>(),
                ),
                _ => None,
            }),
            _ => None,
        })
        .collect();
    assert!(texts.contains(&"Text Box".to_string()), "{texts:?}");
    assert!(texts.contains(&"Aspose.Slides for .NET".to_string()), "{texts:?}");
    assert!(texts.contains(&"Welcome".to_string()), "{texts:?}");
}

/// PPTX analogue of the DOCX title duplication — the title heading
/// duplicating into the slide body when it isn't the section's first element.
#[test]
fn test_pptx_title_heading_not_first_element_does_not_duplicate_on_roundtrip() {
    use office_oxide::ir::{
        Element, Heading, InlineContent, Metadata, Paragraph, Section, TextSpan,
    };

    let ir = DocumentIR {
        metadata: Metadata {
            format: DocumentFormat::Pptx,
            ..Default::default()
        },
        sections: vec![Section {
            title: Some("Slide Title".to_string()),
            elements: vec![
                Element::Paragraph(Paragraph {
                    content: vec![InlineContent::Text(TextSpan::plain(
                        "A decorative shape before the title",
                    ))],
                    ..Default::default()
                }),
                Element::Heading(Heading {
                    level: 1,
                    content: vec![InlineContent::Text(TextSpan::plain("Slide Title"))],
                    ..Default::default()
                }),
            ],
            ..Default::default()
        }],
        defined_names: Vec::new(),
    };

    let mut buf = Cursor::new(Vec::new());
    create::create_from_ir_to_writer(&ir, DocumentFormat::Pptx, &mut buf).unwrap();
    buf.set_position(0);
    let text = Document::from_reader(buf, DocumentFormat::Pptx)
        .unwrap()
        .plain_text();

    let occurrences = text.matches("Slide Title").count();
    assert_eq!(
        occurrences, 1,
        "title duplicated into body on write, expected exactly one occurrence: {text:?}"
    );
}
