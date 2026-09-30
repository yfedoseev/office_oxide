//! Legacy `.ppt` extraction on synthetic decks built in code by
//! `tests/common/ppt.rs`.

mod common;

use common::ppt::*;
use office_oxide::ir::{Element, InlineContent};
use office_oxide::{Document, DocumentFormat};
use std::io::Cursor;

fn open(bytes: Vec<u8>) -> Result<Document, office_oxide::OfficeError> {
    Document::from_reader(Cursor::new(bytes), DocumentFormat::Ppt)
}

fn slide(text: &str) -> PptSlide {
    PptSlide {
        shapes: text_shape(1, text, &[]),
        ..Default::default()
    }
}

fn texts(elements: &[Element]) -> Vec<String> {
    let mut out = Vec::new();
    for e in elements {
        let content = match e {
            Element::Paragraph(p) => &p.content,
            Element::Heading(h) => &h.content,
            _ => continue,
        };
        let mut s = String::new();
        for c in content {
            if let InlineContent::Text(t) = c {
                s.push_str(&t.text);
            }
        }
        out.push(s);
    }
    out
}

/// The builder itself: a two-slide deck parses with each slide's text.
#[test]
fn test_synthetic_ppt_round_trips_slide_text() {
    let deck = PptBuilder {
        slides: vec![slide("First slide"), slide("Second slide")],
        ..Default::default()
    };
    let doc = open(deck.build()).expect("synthetic .ppt must parse");
    let ir = doc.to_ir();
    assert_eq!(ir.sections.len(), 2);
    assert_eq!(texts(&ir.sections[0].elements), ["First slide"]);
    assert_eq!(texts(&ir.sections[1].elements), ["Second slide"]);
}

fn notes_text(ir: &office_oxide::ir::DocumentIR, slide: usize) -> Vec<String> {
    texts(ir.sections[slide].speaker_notes.as_deref().unwrap_or(&[]))
}

/// Speaker notes live in `Notes` containers listed by the
/// `NotesListWithTextContainer` and linked from `SlideAtom.notesIdRef`
/// ([MS-PPT] `SlideAtom`, `NotesAtom`). Only the slide list was ever read,
/// so notes from a real deck never reached `speaker_notes`.
#[test]
fn test_speaker_notes_are_read_from_notes_containers() {
    let deck = PptBuilder {
        slides: vec![
            PptSlide {
                notes: Some("Say hello first.\rThen the agenda.".into()),
                link_notes_from_slide: true,
                ..slide("Slide one")
            },
            slide("Slide two"),
            PptSlide {
                // Linked only by NotesAtom.slideIdRef.
                notes: Some("Closing remarks.".into()),
                ..slide("Slide three")
            },
        ],
        ..Default::default()
    };
    let doc = open(deck.build()).unwrap();
    let ir = doc.to_ir();
    assert_eq!(ir.sections.len(), 3, "notes must not become slides");
    assert_eq!(notes_text(&ir, 0), ["Say hello first.", "Then the agenda."]);
    assert!(ir.sections[1].speaker_notes.is_none());
    assert_eq!(notes_text(&ir, 2), ["Closing remarks."]);
    // Notes are presenter-only: not in the slide body or plain text.
    assert!(
        !texts(&ir.sections[0].elements)
            .iter()
            .any(|t| t.contains("hello"))
    );
    assert!(!doc.plain_text().contains("Say hello first."));
}

fn assert_encrypted(result: Result<Document, office_oxide::OfficeError>) {
    let err = result.err().expect("an encrypted deck must be an error");
    assert!(err.to_string().contains("encrypted"), "{err}");
}

/// A password-protected deck parsed as an empty or garbage deck with `Ok`:
/// `PptError::Encrypted` was declared but never raised. [MS-PPT]
/// `CurrentUserAtom.headerToken` is 0xF3D1C4DF for an encrypted document.
#[test]
fn test_encrypted_header_token_is_an_error() {
    let deck = PptBuilder {
        slides: vec![slide("Ciphertext")],
        header_token: Some(HEADER_TOKEN_ENCRYPTED),
        ..Default::default()
    };
    assert_encrypted(open(deck.build()));
}

/// The other spec signal: `UserEditAtom.encryptSessionPersistIdRef`, the
/// optional field present only when the document is encrypted — caught
/// even when the "Current User" stream says nothing.
#[test]
fn test_encrypt_session_persist_id_ref_is_an_error() {
    let deck = PptBuilder {
        slides: vec![slide("Ciphertext")],
        encrypt_session_persist_id: Some(9),
        ..Default::default()
    };
    assert_encrypted(open(deck.build()));
}
