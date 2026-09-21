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
use common::{FibTweaks, Para, Subdocs, build_doc, build_doc_full, build_word6_doc, prose_grpprl};
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
// Word 6.0/95
// ---------------------------------------------------------------------------

/// Word 6.0/95 used to be refused outright (after an earlier release read
/// it with Word 97 offsets and returned nothing with `Ok`). Its text is
/// read now: one byte per character from `fcMin`, story lengths from the
/// Word 6 FIB, no table stream — the shape every other reader handles.
#[test]
fn test_a_word_6_or_95_file_extracts_its_text() {
    let bytes = build_word6_doc(0xA5DC, b"Hello from Word 6.\rSecond paragraph.\r");
    let doc = open_doc(bytes).expect("a Word 6 file opens");
    assert_eq!(doc.plain_text(), "Hello from Word 6.\nSecond paragraph.\n");
    let ir = doc.to_ir();
    assert_eq!(ir.sections.len(), 1);
    assert!(!ir.sections[0].elements.is_empty());
}

/// The Macintosh Word 6/95 magics (0xA697–0xA699, `nFib` 0x65/0x68) are
/// the same family and used to get the generic "unknown wIdent" refusal.
#[test]
fn test_mac_word_6_or_95_magics_are_read_as_that_family() {
    for wident in [0xA697u16, 0xA698, 0xA699] {
        let bytes = build_word6_doc(wident, b"Mac Word text.\r");
        let doc = open_doc(bytes).unwrap_or_else(|e| panic!("wIdent 0x{wident:04X}: {e}"));
        assert_eq!(doc.plain_text(), "Mac Word text.\n", "wIdent 0x{wident:04X}");
    }
}

/// `fEncrypted` sits in the same flags word in both FIB layouts; an
/// encrypted Word 6/95 file is refused as encrypted, not read as
/// ciphertext.
#[test]
fn test_an_encrypted_word_6_file_reports_encryption() {
    let mut bytes = build_word6_doc(0xA5DC, b"ciphertext\r");
    // The WordDocument stream starts at sector 2 (header + dir + FAT);
    // flip fEncrypted (0x0100) in its flags word at 0x0A.
    let wd = 512 * 3;
    bytes[wd + 0x0B] |= 0x01;
    let err = expect_err(open_doc(bytes), "an encrypted Word 6 file must be refused");
    assert!(err.to_string().to_lowercase().contains("encrypt"), "{err}");
}

// ---------------------------------------------------------------------------
// Encrypted legacy files
// ---------------------------------------------------------------------------

#[test]
fn test_an_encrypted_document_reports_encryption_not_emptiness() {
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
fn test_an_unencrypted_document_still_parses() {
    let doc = open_doc(build_doc(&[para("plain text")])).expect("parse");
    assert!(doc.plain_text().contains("plain text"));
}

// ---------------------------------------------------------------------------
// Subdocuments
// ---------------------------------------------------------------------------

#[test]
fn test_footnotes_headers_comments_and_text_boxes_reach_the_ir() {
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

    // A comment and a real endnote both land as
    // `Element::Endnote` (there is no dedicated comment variant), but their
    // `marker` must distinguish them: the comment's text ("REVIEW NOTE")
    // must carry marker "comment", and the real endnote's text ("ENDNOTE
    // ONE") must carry marker "endnote" — not the other way around, which
    // is what the FIB offset bug produced (comments always read
    // zero-length, and a real endnote's content surfaced under the
    // textbox slot instead).
    let note_markers: Vec<(Option<&str>, String)> = ir.sections[0]
        .elements
        .iter()
        .filter_map(|e| match e {
            Element::Endnote(n) => {
                let text: String = n
                    .content
                    .iter()
                    .filter_map(|c| match c {
                        Element::Paragraph(p) => p.content.iter().find_map(|ic| match ic {
                            office_oxide::ir::InlineContent::Text(t) => Some(t.text.clone()),
                            _ => None,
                        }),
                        _ => None,
                    })
                    .collect();
                Some((n.marker.as_deref(), text))
            },
            _ => None,
        })
        .collect();
    assert!(
        note_markers
            .iter()
            .any(|(marker, text)| *marker == Some("comment") && text.contains("REVIEW NOTE")),
        "the comment must carry marker \"comment\", got {note_markers:?}"
    );
    assert!(
        note_markers
            .iter()
            .any(|(marker, text)| *marker == Some("endnote") && text.contains("ENDNOTE ONE")),
        "the real endnote must carry marker \"endnote\", got {note_markers:?}"
    );
}

#[test]
fn test_a_document_with_no_subdocuments_gains_no_extra_elements() {
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
// A read failure is an error, not an empty document
// ---------------------------------------------------------------------------

#[test]
fn test_a_document_whose_piece_table_is_out_of_bounds_reports_the_failure() {
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
// The line-shape heading guess is gated on real outline data
// ---------------------------------------------------------------------------

/// `sprmPOutLvl` (0x2640), 1-byte operand: 0 = Heading 1 … 8 = Heading 9.
fn outline_grpprl(level: u8) -> Vec<u8> {
    vec![0x40, 0x26, level]
}

#[test]
fn test_a_document_with_real_outline_levels_uses_them_and_does_not_guess() {
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
fn test_a_document_with_no_outline_data_still_gets_the_line_shape_guess() {
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

/// A compound file is read as the document it holds, whatever its
/// extension says: a workbook saved as `.doc` opens as a spreadsheet
/// (MS-XLS `Workbook` stream) instead of failing with "WordDocument
/// stream not found" — the way every other reader treats it.
#[test]
fn test_a_workbook_under_a_doc_extension_opens_as_a_spreadsheet() {
    use common::{biff, cfb_with_stream};
    // BIFF8: workbook globals (BOF, BOUNDSHEET, EOF) then one sheet with
    // a LABEL cell (0x0204) "hello".
    let bof = |kind: u16| {
        let mut b = 0x0600u16.to_le_bytes().to_vec();
        b.extend_from_slice(&kind.to_le_bytes());
        b.extend_from_slice(&[0u8; 12]);
        biff(0x0809, &b)
    };
    let mut s = bof(0x0005);
    let mut bs = 0u32.to_le_bytes().to_vec();
    bs.extend_from_slice(&[0, 0, 1, 0, b'S']);
    s.extend(biff(0x0085, &bs));
    s.extend(biff(0x000A, &[]));
    s.extend(bof(0x0010));
    let mut label = 0u16.to_le_bytes().to_vec();
    label.extend_from_slice(&0u16.to_le_bytes());
    label.extend_from_slice(&0u16.to_le_bytes());
    label.extend_from_slice(&5u16.to_le_bytes());
    label.push(0);
    label.extend_from_slice(b"hello");
    s.extend(biff(0x0204, &label));
    s.extend(biff(0x000A, &[]));
    let bytes = cfb_with_stream("Workbook", &s);

    let doc = Document::from_reader(Cursor::new(bytes.clone()), DocumentFormat::Doc)
        .expect("a compound file holding a Workbook stream is a spreadsheet");
    assert_eq!(doc.format(), DocumentFormat::Xls);
    assert!(doc.plain_text().contains("hello"), "{:?}", doc.plain_text());

    let dir = std::env::temp_dir().join(format!("oo-mislabeled-{}", std::process::id()));
    std::fs::create_dir_all(&dir).unwrap();
    let path = dir.join("workbook.doc");
    std::fs::write(&path, &bytes).unwrap();
    let doc = Document::open(&path).expect("same through the path API");
    assert_eq!(doc.format(), DocumentFormat::Xls);
    std::fs::remove_dir_all(&dir).unwrap();

    // A compound file holding none of the three is still the extension's
    // own error, not a silent success.
    let other = cfb_with_stream("SomeStream", b"data");
    assert!(Document::from_reader(Cursor::new(other), DocumentFormat::Doc).is_err());
}

/// Word for Windows 2.0 wrote flat files — the FIB at byte 0, no
/// compound container — with CR LF paragraph marks and code-page text.
/// They were refused as "not a compound file" while catdoc and antiword
/// read them.
#[test]
fn test_a_word_2_flat_file_extracts_its_text() {
    let text = b"Sonderaktion Winter 1990/1991\r\nWegen g\xfcnstiger Einkaufsm\xf6glichkeiten.\r\n\r\nEnde.\r\n";
    let mut bytes = vec![0u8; 0x180];
    bytes[0..2].copy_from_slice(&0xA5DBu16.to_le_bytes());
    bytes[2..4].copy_from_slice(&45u16.to_le_bytes());
    bytes[6..8].copy_from_slice(&0x0407u16.to_le_bytes());
    bytes[0x18..0x1C].copy_from_slice(&0x180u32.to_le_bytes());
    bytes[0x1C..0x20].copy_from_slice(&(0x180 + text.len() as u32).to_le_bytes());
    bytes[0x34..0x38].copy_from_slice(&(text.len() as u32).to_le_bytes());
    bytes.extend_from_slice(text);

    let doc = open_doc(bytes.clone()).expect("a Word 2.0 file opens");
    assert_eq!(doc.format(), DocumentFormat::Doc);
    assert_eq!(
        doc.plain_text(),
        "Sonderaktion Winter 1990/1991\nWegen günstiger Einkaufsmöglichkeiten.\n\nEnde.\n"
    );
    let ir = doc.to_ir();
    assert!(
        ir.sections[0].elements.len() >= 3,
        "each paragraph mark ends a paragraph: {:?}",
        ir.sections[0].elements
    );

    // Word 1.x has the same shape under its own magic.
    bytes[0..2].copy_from_slice(&0xA59Bu16.to_le_bytes());
    assert!(
        open_doc(bytes)
            .unwrap()
            .plain_text()
            .starts_with("Sonderaktion")
    );
}

/// The first PAPX FKP run may start below the text's first piece — Word
/// leaves `rgfc[0]` at the start of the text area rather than at
/// `fcMin`. That start mapped to no CP, so the document's opening
/// paragraph was missing from `to_ir()` (and `to_html()`) while
/// `plain_text()` had it.
#[test]
fn test_an_opening_paragraph_whose_fkp_run_starts_before_the_text_is_kept() {
    let bytes = build_doc_full(
        &[para("Opening paragraph."), para("Second paragraph.")],
        &Subdocs::default(),
        FibTweaks {
            first_fkp_fc_before_text: 512,
            ..Default::default()
        },
    );
    let doc = open_doc(bytes).unwrap();
    assert!(doc.plain_text().starts_with("Opening paragraph."));
    let ir = doc.to_ir();
    let first = ir.sections[0]
        .elements
        .iter()
        .find_map(|e| match e {
            Element::Paragraph(p) => Some(office_oxide::ir::inline_to_text(&p.content)),
            Element::Heading(h) => Some(office_oxide::ir::inline_to_text(&h.content)),
            _ => None,
        })
        .unwrap_or_default();
    assert_eq!(first, "Opening paragraph.", "{:?}", ir.sections[0].elements);
    assert!(ir.to_html().contains("Opening paragraph."));
}
