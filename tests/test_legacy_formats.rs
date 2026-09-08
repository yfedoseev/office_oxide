//! Legacy `.doc` / `.xls` / `.ppt` behaviour.
//!
//! Fixtures are synthesised in code (see `tests/common`). The recurring
//! defect these guard against is the same one: a file that cannot be read,
//! or content that was read and then discarded, surfacing as an empty
//! string with `Ok` — which a caller cannot distinguish from a document
//! that genuinely has no text.

mod common;

use std::io::Cursor;

#[allow(unused_imports)]
use common::{FibTweaks, Para, Subdocs, build_doc, build_doc_full, prose_grpprl};
use office_oxide::ir::Element;
use office_oxide::{Document, DocumentFormat};

fn para(text: &'static str) -> Para {
    Para {
        text,
        terminator: '\r',
        grpprl: prose_grpprl(),
    }
}

fn open_doc(bytes: Vec<u8>) -> office_oxide::Result<Document> {
    Document::from_reader(Cursor::new(bytes), DocumentFormat::Doc)
}

fn expect_err(r: office_oxide::Result<Document>, why: &str) -> office_oxide::OfficeError {
    match r {
        Ok(_) => panic!("{why}"),
        Err(e) => e,
    }
}

// ---------------------------------------------------------------------------
// #167 — Word 6.0/95
// ---------------------------------------------------------------------------

#[test]
fn a_word_6_or_95_file_is_refused_rather_than_read_with_word_97_offsets() {
    // 0xA5DC has a completely different FIB layout; reading Word 97 offsets
    // out of it produced a confident empty result for a real document.
    let bytes = build_doc_full(
        &[para("hello")],
        &Subdocs::default(),
        FibTweaks {
            wident: Some(0xA5DC),
            ..Default::default()
        },
    );
    let err = expect_err(open_doc(bytes), "Word 6.0/95 must not parse to an empty Ok");
    let msg = err.to_string();
    assert!(
        msg.contains("unsupported Word version") && msg.contains("A5DC"),
        "the error must name the version, got {msg}"
    );
}

// ---------------------------------------------------------------------------
// #169 — encrypted legacy files
// ---------------------------------------------------------------------------

#[test]
fn an_encrypted_document_reports_encryption_not_emptiness() {
    let bytes = build_doc_full(
        &[para("ciphertext")],
        &Subdocs::default(),
        FibTweaks {
            encrypted: true,
            ..Default::default()
        },
    );
    let err = expect_err(open_doc(bytes), "an encrypted file must not parse to Ok");
    assert!(err.to_string().contains("encrypted"), "got {err}");
}

#[test]
fn an_unencrypted_document_still_parses() {
    let doc = open_doc(build_doc(&[para("plain text")])).expect("parse");
    assert!(doc.plain_text().contains("plain text"));
}

// ---------------------------------------------------------------------------
// #195 — subdocuments
// ---------------------------------------------------------------------------

#[test]
fn footnotes_headers_comments_and_text_boxes_reach_the_ir() {
    // The FIB's `ccp*` lengths delimit these; they were parsed and then
    // never used, so none of this content reached a consumer.
    let bytes = build_doc_full(
        &[para("BODY")],
        &Subdocs {
            footnotes: "FOOTNOTE ONE",
            headers: "PAGE HEADER",
            comments: "REVIEW NOTE",
            endnotes: "ENDNOTE ONE",
            textboxes: "SIDEBAR",
        },
        FibTweaks::default(),
    );
    let doc = open_doc(bytes).expect("parse");
    let ir = doc.to_ir();
    let text = ir.plain_text();
    for token in [
        "BODY",
        "FOOTNOTE ONE",
        "PAGE HEADER",
        "REVIEW NOTE",
        "ENDNOTE ONE",
        "SIDEBAR",
    ] {
        assert!(text.contains(token), "{token} missing from {text:?}");
    }

    // Footnotes land in the footnote slot, not as body paragraphs.
    let kinds: Vec<&str> = ir.sections[0]
        .elements
        .iter()
        .map(|e| match e {
            Element::Footnote(_) => "footnote",
            Element::Endnote(_) => "endnote",
            Element::TextBox(_) => "textbox",
            _ => "other",
        })
        .collect();
    assert!(kinds.contains(&"footnote"), "no footnote element: {kinds:?}");
    assert!(kinds.contains(&"endnote"), "no endnote element: {kinds:?}");
    assert!(kinds.contains(&"textbox"), "no textbox element: {kinds:?}");
}

#[test]
fn a_document_with_no_subdocuments_gains_no_extra_elements() {
    let doc = open_doc(build_doc(&[para("just body")])).expect("parse");
    let ir = doc.to_ir();
    assert!(
        !ir.sections[0]
            .elements
            .iter()
            .any(|e| matches!(e, Element::Footnote(_) | Element::Endnote(_))),
        "empty ccp* lengths must produce no note elements"
    );
}

// ---------------------------------------------------------------------------
// #168 — a read failure is an error, not an empty document
// ---------------------------------------------------------------------------

#[test]
fn a_document_whose_piece_table_is_out_of_bounds_reports_the_failure() {
    // Point fcClx far past the end of the table stream.
    let bytes = build_doc_full(
        &[para("hello")],
        &Subdocs::default(),
        FibTweaks {
            clx_offset: Some(0x00FF_FFFF),
            ..Default::default()
        },
    );
    let err =
        expect_err(open_doc(bytes), "an unreadable piece table must not parse to an empty Ok");
    let msg = err.to_string();
    assert!(
        msg.contains("piece table") || msg.contains("CLX"),
        "the error must name the failing structure, got {msg}"
    );
}

// ---------------------------------------------------------------------------
// #139 — the line-shape heading guess is gated on real outline data
// ---------------------------------------------------------------------------

/// `sprmPOutLvl` (0x2640), 1-byte operand: 0 = Heading 1 … 8 = Heading 9.
fn outline_grpprl(level: u8) -> Vec<u8> {
    vec![0x40, 0x26, level]
}

#[test]
fn a_document_with_real_outline_levels_uses_them_and_does_not_guess() {
    // "ALL CAPS" would be guessed as a heading by the line-shape rule. With
    // real outline data present, the guess must not run alongside it —
    // otherwise one document emits real levels and ALL-CAPS guesses that
    // disagree about the same paragraphs.
    let doc = open_doc(build_doc_full(
        &[
            Para {
                text: "Real Heading",
                terminator: '\r',
                grpprl: outline_grpprl(1),
            },
            Para {
                text: "ALL CAPS LINE",
                terminator: '\r',
                grpprl: prose_grpprl(),
            },
        ],
        &Subdocs::default(),
        FibTweaks::default(),
    ))
    .expect("parse");

    let ir = doc.to_ir();
    let headings: Vec<(u8, String)> = ir.sections[0]
        .elements
        .iter()
        .filter_map(|e| match e {
            Element::Heading(h) => Some((
                h.level,
                h.content
                    .iter()
                    .filter_map(|c| match c {
                        office_oxide::ir::InlineContent::Text(t) => Some(t.text.as_str()),
                        _ => None,
                    })
                    .collect::<String>(),
            )),
            _ => None,
        })
        .collect();
    assert_eq!(
        headings,
        vec![(2u8, "Real Heading".to_string())],
        "outlineLvl=1 is Heading 2, and the ALL-CAPS line must stay a paragraph"
    );
}

#[test]
fn a_document_with_no_outline_data_still_gets_the_line_shape_guess() {
    // The 88 documents that would otherwise lose their headings — and with
    // them `metadata.title` — under a stylesheet-based gate.
    let doc = open_doc(build_doc(&[
        Para {
            text: "MEMORANDUM",
            terminator: '\r',
            grpprl: prose_grpprl(),
        },
        Para {
            text: "Body text follows here.",
            terminator: '\r',
            grpprl: prose_grpprl(),
        },
    ]))
    .expect("parse");
    let ir = doc.to_ir();
    assert!(
        ir.sections[0]
            .elements
            .iter()
            .any(|e| matches!(e, Element::Heading(_))),
        "a document with no outline data keeps the heuristic"
    );
    assert!(ir.metadata.title.is_some(), "and keeps its derived title");
}
