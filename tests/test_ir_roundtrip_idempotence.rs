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
fn docx_ir_roundtrip_is_idempotent() {
    assert_idempotent(DocumentFormat::Docx, DOCX_MD);
}

#[test]
fn xlsx_ir_roundtrip_is_idempotent() {
    assert_idempotent(DocumentFormat::Xlsx, XLSX_MD);
}

#[test]
fn pptx_ir_roundtrip_is_idempotent() {
    assert_idempotent(DocumentFormat::Pptx, PPTX_MD);
}

/// Focused guard for the PPTX body-text-loss defect: the slide body content
/// (bullets) must survive a round-trip, not be replaced by the title.
#[test]
fn pptx_roundtrip_preserves_slide_body() {
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
