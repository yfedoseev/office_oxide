//! Text and formatting fidelity of the legacy `.doc` reader, on synthetic
//! documents built in code by `tests/common`.

mod common;

use common::{FibTweaks, Para, Subdocs, build_doc_full, open_doc, prose_grpprl};
use office_oxide::ir::{Element, InlineContent};

fn para(text: &'static str) -> Para {
    Para {
        text,
        terminator: '\r',
        grpprl: prose_grpprl(),
    }
}

fn all_text(ir: &office_oxide::ir::DocumentIR) -> String {
    fn walk(e: &Element, out: &mut String) {
        match e {
            Element::Paragraph(p) => {
                for c in &p.content {
                    if let InlineContent::Text(t) = c {
                        out.push_str(&t.text);
                    }
                }
                out.push('\n');
            },
            Element::Heading(h) => {
                for c in &h.content {
                    if let InlineContent::Text(t) = c {
                        out.push_str(&t.text);
                    }
                }
                out.push('\n');
            },
            Element::Footnote(n) | Element::Endnote(n) => {
                for e in &n.content {
                    walk(e, out);
                }
            },
            _ => {},
        }
    }
    let mut out = String::new();
    for s in &ir.sections {
        for e in &s.elements {
            walk(e, &mut out);
        }
    }
    out
}

/// Main-story reference marks — `0x02` (auto-numbered footnote/endnote
/// reference) and `0x05` (comment reference) — are anchors, not text, and
/// leaked into `plain_text()` and the IR as literal U+0002/U+0005.
/// `0x1E` is a non-breaking hyphen (drawn, U+2011) and `0x1F` an optional
/// hyphen (a line-break hint, dropped as the DOCX reader drops
/// `w:softHyphen`).
#[test]
fn test_reference_marks_and_hyphen_controls_do_not_leak_into_text() {
    let paras = [para(
        "See note\u{2} and comment\u{5}; well\u{1E}known co\u{1F}operation.",
    )];
    let subdocs = Subdocs {
        footnotes: "\u{2} First note.\r\u{2} Second note.",
        ..Default::default()
    };
    let doc = open_doc(&build_doc_full(&paras, &subdocs, FibTweaks::default()));
    let text = doc.plain_text();
    for bad in ['\u{2}', '\u{5}', '\u{1E}', '\u{1F}'] {
        assert!(!text.contains(bad), "U+{:04X} leaked: {text:?}", bad as u32);
    }
    assert!(
        text.contains("See note and comment; well\u{2011}known cooperation."),
        "{text:?}"
    );

    let ir = doc.to_ir();
    let ir_text = all_text(&ir);
    for bad in ['\u{2}', '\u{5}', '\u{1E}', '\u{1F}'] {
        assert!(!ir_text.contains(bad), "U+{:04X} leaked into IR: {ir_text:?}", bad as u32);
    }
    // The footnote bodies are still split one per note.
    let notes: Vec<_> = ir.sections[0]
        .elements
        .iter()
        .filter(|e| matches!(e, Element::Footnote(_)))
        .collect();
    assert_eq!(notes.len(), 2, "{:?}", ir.sections[0].elements);
}

/// A nested table (`sprmPItap` > 1) is flattened into its outer table.
/// That used to be announced by a paragraph of text the document does not
/// have; it is a metadata warning now, and the document's own text is
/// untouched.
#[test]
fn test_flattened_nested_table_is_a_warning_not_document_text() {
    use common::{cell_grpprl, row_grpprl};
    // Depth-2 cell and row mark: sprmPItap operand = 2.
    let mut cell = cell_grpprl();
    cell[5] = 2;
    let mut row = row_grpprl(&[0, 1000], &[0]);
    row[8] = 2;
    let paras = [
        Para {
            text: "Inner",
            terminator: '\u{7}',
            grpprl: cell,
        },
        Para {
            text: "",
            terminator: '\u{7}',
            grpprl: row,
        },
        para("After."),
    ];
    let doc = open_doc(&build_doc_full(&paras, &Subdocs::default(), FibTweaks::default()));
    let ir = doc.to_ir();
    let md = ir.to_markdown();
    assert!(!md.contains("nested table"), "fabricated text in the output: {md}");
    assert!(md.contains("Inner"), "{md}");
    assert!(
        ir.metadata
            .warnings
            .iter()
            .any(|w| w.contains("nested table")),
        "{:?}",
        ir.metadata.warnings
    );
}
