//! Integration tests for legacy `.doc` list extraction (issue #115).
//!
//! The document is built **in code** by a minimal synthetic `.doc` writer
//! (`tests/common/mod.rs`) — no third-party fixture blob is committed
//! (AGENTS.md rule #4). It contains an introductory paragraph, a flat
//! three-item bulleted list, and a trailing paragraph:
//!
//! ```text
//! This is a simple word document created with office_oxide.
//!   • First item in list
//!   • Second item in list
//!   • Third item in list
//! This is the last paragraph.
//! ```
//!
//! These tests assert the IR now emits a real `Element::List` keeping every
//! item, and that the list is flanked by its neighbouring paragraphs.

mod common;

use common::{Para, build_doc, list_grpprl, open_doc, prose_grpprl};
use office_oxide::ir::{Element, InlineContent, ListItem, Paragraph};

/// Locate the first `Element::List` in the document's single section.
fn first_list(ir: &office_oxide::ir::DocumentIR) -> Option<&office_oxide::ir::List> {
    ir.sections
        .iter()
        .flat_map(|s| s.elements.iter())
        .find_map(|e| match e {
            Element::List(l) => Some(l),
            _ => None,
        })
}

/// Concatenate the `TextSpan` text inside a list item's block elements.
fn item_text(item: &ListItem) -> String {
    let mut out = String::new();
    for el in &item.content {
        if let Element::Paragraph(Paragraph { content, .. }) = el {
            for ic in content {
                if let InlineContent::Text(t) = ic {
                    out.push_str(&t.text);
                }
            }
        }
    }
    out
}

/// Build a synthetic list `.doc`.
fn synthetic_list_doc() -> Vec<u8> {
    let paras = [
        Para {
            text: "This is a simple word document created with office_oxide.",
            terminator: '\r',
            grpprl: prose_grpprl(),
        },
        Para {
            text: "First item in list",
            terminator: '\r',
            grpprl: list_grpprl(0),
        },
        Para {
            text: "Second item in list",
            terminator: '\r',
            grpprl: list_grpprl(0),
        },
        Para {
            text: "Third item in list",
            terminator: '\r',
            grpprl: list_grpprl(0),
        },
        Para {
            text: "This is the last paragraph.",
            terminator: '\r',
            grpprl: prose_grpprl(),
        },
    ];
    build_doc(&paras)
}

#[test]
fn doc_list_produces_a_list_element() {
    let bytes = synthetic_list_doc();
    let doc = open_doc(&bytes);
    let ir = doc.to_ir();

    let list = first_list(&ir).expect("expected an Element::List in the IR");
    assert_eq!(list.items.len(), 3, "list should have 3 items");
}

#[test]
fn doc_list_item_texts_match_fixture() {
    let bytes = synthetic_list_doc();
    let doc = open_doc(&bytes);
    let ir = doc.to_ir();

    let list = first_list(&ir).expect("expected a list");
    let texts: Vec<String> = list.items.iter().map(item_text).collect();
    assert_eq!(texts.len(), 3);
    assert_eq!(texts[0], "First item in list");
    assert_eq!(texts[1], "Second item in list");
    assert_eq!(texts[2], "Third item in list");
}

#[test]
fn doc_list_is_surrounded_by_paragraphs() {
    let bytes = synthetic_list_doc();
    let doc = open_doc(&bytes);
    let ir = doc.to_ir();
    let elements = &ir.sections[0].elements;

    // [Paragraph, List, Paragraph] — the list must not swallow its neighbours.
    assert!(
        matches!(elements.first(), Some(Element::Paragraph(p)) if p.content.iter().any(|c| matches!(c, InlineContent::Text(t) if t.text.contains("simple word document"))))
    );
    assert!(matches!(elements[1], Element::List(_)));
    assert!(
        matches!(elements.last(), Some(Element::Paragraph(p)) if p.content.iter().any(|c| matches!(c, InlineContent::Text(t) if t.text.contains("last paragraph"))))
    );
}
