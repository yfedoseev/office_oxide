//! The text of a Word document embedded in a `.doc`.
//!
//! A `.doc` whose body is made of embedded Word objects named each of them
//! with a placeholder and nothing else, although the objects hold the
//! article text. The text is now read from each object's own streams, within
//! the document's embedded-package account: a count, a byte allowance, and a
//! share of the text budget. An object that cannot be read keeps its
//! placeholder and is reported.

mod common;

use common::{
    CfbNode, FibTweaks, LimitScope, Para, Subdocs, build_doc_streams, cfb_tree, default_limits,
};
use office_oxide::ir::Element;
use office_oxide::limits::{max_embedded_bytes, max_embedded_packages, max_text_chars};
use office_oxide::limits::{set_max_embedded_bytes, set_max_embedded_packages, set_max_text_chars};
use office_oxide::{Document, DocumentFormat};
use std::io::Cursor;

const PLACEHOLDER: &str = "[Embedded Microsoft Word Document]";

fn para(text: &'static str) -> Para {
    Para {
        text,
        terminator: '\r',
        grpprl: Vec::new(),
    }
}

fn streams(text: &'static str) -> (Vec<u8>, Vec<u8>) {
    build_doc_streams(&[para(text)], &Subdocs::default(), FibTweaks::default())
}

/// An object storage holding a Word document with `text`.
fn word_object(name: &str, text: &'static str) -> CfbNode {
    let (wd, table) = streams(text);
    CfbNode::Storage(
        name.to_string(),
        vec![
            CfbNode::Stream("WordDocument".into(), wd),
            CfbNode::Stream("0Table".into(), table),
            CfbNode::Stream("\u{1}CompObj".into(), vec![0u8; 28]),
        ],
    )
}

fn host(objects: Vec<CfbNode>) -> Vec<u8> {
    let (wd, table) = streams("Host body text");
    cfb_tree(vec![
        CfbNode::Stream("WordDocument".into(), wd),
        CfbNode::Stream("0Table".into(), table),
        CfbNode::Storage("ObjectPool".into(), objects),
    ])
}

fn open(bytes: Vec<u8>) -> Document {
    Document::from_reader(Cursor::new(bytes), DocumentFormat::Doc).unwrap()
}

#[test]
fn test_the_text_of_an_embedded_word_document_follows_its_placeholder() {
    let _defaults = default_limits();
    let doc = open(host(vec![word_object("_1", "Festive spread recipe")]));
    let text = doc.plain_text();
    let at = |needle: &str| {
        text.find(needle)
            .unwrap_or_else(|| panic!("{needle:?} in {text:?}"))
    };
    assert!(at("Host body text") < at(PLACEHOLDER), "{text:?}");
    assert!(at(PLACEHOLDER) < at("Festive spread recipe"), "{text:?}");
}

#[test]
fn test_every_surface_carries_the_embedded_text() {
    let _defaults = default_limits();
    let doc = open(host(vec![word_object("_1", "Festive spread recipe")]));
    assert!(doc.to_markdown().contains("Festive spread recipe"));
    let ir = doc.to_ir();
    let paragraphs: Vec<String> = ir
        .sections
        .iter()
        .flat_map(|s| &s.elements)
        .filter_map(|e| match e {
            Element::Paragraph(p) => Some(format!("{p:?}")),
            _ => None,
        })
        .collect();
    assert!(
        paragraphs
            .iter()
            .any(|p| p.contains("Festive spread recipe")),
        "{paragraphs:?}"
    );
    assert!(!ir.metadata.text_truncated);
}

#[test]
fn test_objects_beyond_the_package_limit_keep_their_placeholder_and_are_reported() {
    let mut scope = LimitScope::new();
    scope.set(max_embedded_packages, set_max_embedded_packages, 2);
    let doc = open(host(vec![
        word_object("_1", "alpha article"),
        word_object("_2", "bravo article"),
        word_object("_3", "charlie article"),
        word_object("_4", "delta article"),
    ]));
    let text = doc.plain_text();
    assert_eq!(text.matches(PLACEHOLDER).count(), 4, "{text:?}");
    let read = ["alpha", "bravo", "charlie", "delta"]
        .iter()
        .filter(|w| text.contains(&format!("{w} article")))
        .count();
    assert_eq!(read, 2, "{text:?}");
    let warnings = doc.to_ir().metadata.warnings;
    assert!(
        warnings
            .iter()
            .any(|w| w.contains("embedded") && w.contains("not")),
        "{warnings:?}"
    );
    drop(scope);
}

#[test]
fn test_embedded_text_is_charged_to_the_text_budget() {
    let mut scope = LimitScope::new();
    scope.set(max_text_chars, set_max_text_chars, 60);
    let doc = open(host(vec![
        word_object("_1", "a first embedded article of about forty characters"),
        word_object("_2", "a second embedded article that cannot also fit in"),
    ]));
    let text = doc.plain_text();
    assert!(text.contains("Host body text"), "{text:?}");
    assert!(
        !(text.contains("first embedded article") && text.contains("second embedded article")),
        "both embedded texts were read past the budget: {text:?}"
    );
    let warnings = doc.to_ir().metadata.warnings;
    assert!(
        warnings
            .iter()
            .any(|w| w.contains("embedded object") && w.contains("budget")),
        "{warnings:?}"
    );
    drop(scope);
}

#[test]
fn test_an_embedded_document_that_cannot_be_read_keeps_its_placeholder() {
    let _defaults = default_limits();
    let broken = CfbNode::Storage(
        "_1".into(),
        vec![
            CfbNode::Stream("WordDocument".into(), vec![0xFFu8; 2048]),
            CfbNode::Stream("0Table".into(), vec![0u8; 512]),
        ],
    );
    let doc = open(host(vec![broken, word_object("_2", "second article readable")]));
    let text = doc.plain_text();
    assert_eq!(text.matches(PLACEHOLDER).count(), 2, "{text:?}");
    assert!(text.contains("second article readable"), "{text:?}");
    assert!(text.contains("Host body text"), "{text:?}");
    let warnings = doc.to_ir().metadata.warnings;
    assert!(warnings.iter().any(|w| w.contains("_1")), "{warnings:?}");
}

/// Only the objects of the document itself are read: an object inside an
/// embedded document is not followed.
#[test]
fn test_objects_inside_an_embedded_document_are_not_followed() {
    let _defaults = default_limits();
    let (wd, table) = streams("outer embedded text");
    let inner = word_object("_9", "inner embedded text");
    let outer = CfbNode::Storage(
        "_1".into(),
        vec![
            CfbNode::Stream("WordDocument".into(), wd),
            CfbNode::Stream("0Table".into(), table),
            CfbNode::Storage("ObjectPool".into(), vec![inner]),
        ],
    );
    let doc = open(host(vec![outer]));
    let text = doc.plain_text();
    assert!(text.contains("outer embedded text"), "{text:?}");
    assert!(!text.contains("inner embedded text"), "{text:?}");
}

#[test]
fn test_a_document_without_embedded_objects_is_unchanged() {
    let _defaults = default_limits();
    let (wd, table) = streams("Host body text");
    let bytes = cfb_tree(vec![
        CfbNode::Stream("WordDocument".into(), wd),
        CfbNode::Stream("0Table".into(), table),
    ]);
    let text = open(bytes).plain_text();
    assert!(text.contains("Host body text") && !text.contains("[Embedded"), "{text:?}");
}

/// The bytes of the objects share one allowance: when it is too small for an
/// object it is not read, and the document says so.
#[test]
fn test_the_byte_allowance_stops_reading_embedded_documents() {
    let mut scope = LimitScope::new();
    scope.set(max_embedded_bytes, set_max_embedded_bytes, 16);
    let doc = open(host(vec![word_object("_1", "never read text")]));
    let text = doc.plain_text();
    assert!(text.contains(PLACEHOLDER) && !text.contains("never read text"), "{text:?}");
    let warnings = doc.to_ir().metadata.warnings;
    assert!(warnings.iter().any(|w| w.contains("ObjectPool/_1")), "{warnings:?}");
    drop(scope);
}

/// One document embedded under several names is read from its content each
/// time it is copied, and each copy is in the text, as it is in the file.
#[test]
fn test_a_document_embedded_under_several_names_appears_once_per_copy() {
    let _defaults = default_limits();
    let doc = open(host(vec![
        word_object("_1", "shared article"),
        word_object("_2", "shared article"),
        word_object("_3", "shared article"),
    ]));
    assert_eq!(doc.plain_text().matches("shared article").count(), 3);
}
