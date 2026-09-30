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

const HF_HAS_DATE: u16 = 0x01;
const HF_HAS_USER_DATE: u16 = 0x04;
const HF_HAS_HEADER: u16 = 0x10;
const HF_HAS_FOOTER: u16 = 0x20;

/// Header/footer text was taken from every `HeadersFootersContainer` —
/// slide and notes/handout alike — and appended to every slide, ignoring
/// the `HeadersFootersAtom` show flags and per-slide overrides ([MS-PPT]
/// `HeadersFootersAtom`, `SlideHeadersFootersContainer`,
/// `NotesHeadersFootersContainer`).
#[test]
fn test_headers_footers_follow_their_container_and_flags() {
    let mut doc_children = headers_footers(
        3,
        HF_HAS_FOOTER | HF_HAS_DATE | HF_HAS_USER_DATE,
        &[(2, "Deck footer"), (0, "1 Jan 2000")],
    );
    // Notes/handout header: never slide text.
    doc_children.extend(headers_footers(4, HF_HAS_HEADER, &[(1, "Notes header")]));
    let mut hidden_footer = slide("Slide two");
    // Slide two overrides: footer text present but not shown.
    hidden_footer
        .shapes
        .extend(headers_footers(3, 0, &[(2, "Deck footer")]));
    let deck = PptBuilder {
        slides: vec![slide("Slide one"), hidden_footer],
        doc_children,
        ..Default::default()
    };
    let ir = open(deck.build()).unwrap().to_ir();
    let one = texts(&ir.sections[0].elements);
    assert!(one.contains(&"Deck footer".to_string()), "{one:?}");
    assert!(one.contains(&"1 Jan 2000".to_string()), "{one:?}");
    assert!(!one.iter().any(|t| t.contains("Notes header")), "{one:?}");
    let two = texts(&ir.sections[1].elements);
    assert!(!two.iter().any(|t| t.contains("Deck footer")), "{two:?}");
    assert!(!two.iter().any(|t| t.contains("Notes header")), "{two:?}");
}

/// A footer whose show flag is off is not slide content.
#[test]
fn test_hidden_footer_is_not_injected() {
    let deck = PptBuilder {
        slides: vec![slide("Only slide")],
        doc_children: headers_footers(3, 0, &[(2, "Unshown footer")]),
        ..Default::default()
    };
    let ir = open(deck.build()).unwrap().to_ir();
    assert_eq!(texts(&ir.sections[0].elements), ["Only slide"]);
}

fn spans(elements: &[Element]) -> Vec<office_oxide::ir::TextSpan> {
    let mut out = Vec::new();
    for e in elements {
        let content = match e {
            Element::Paragraph(p) => &p.content,
            Element::Heading(h) => &h.content,
            _ => continue,
        };
        for c in content {
            if let InlineContent::Text(t) = c {
                out.push(t.clone());
            }
        }
    }
    out
}

/// A run's typeface is a `fontRef` index into the deck's
/// `FontCollectionContainer` of `FontEntityAtom`s ([MS-PPT]
/// `TextCFException`, `FontEntityAtom`). The collection was never parsed,
/// so no `.ppt` span ever had a font name.
#[test]
fn test_character_typeface_resolves_through_the_font_collection() {
    const CF_TYPEFACE: u32 = 1 << 16;
    let text = "Styled";
    let prop = style_text_prop(text.len() as u32, &[(7, CF_TYPEFACE, 1u16.to_le_bytes().to_vec())]);
    let deck = PptBuilder {
        slides: vec![PptSlide {
            shapes: text_shape(1, text, &prop),
            ..Default::default()
        }],
        fonts: vec!["Arial".into(), "Georgia".into()],
        ..Default::default()
    };
    let ir = open(deck.build()).unwrap().to_ir();
    let s = spans(&ir.sections[0].elements);
    assert_eq!(s.len(), 1, "{s:?}");
    assert_eq!(s[0].font_name.as_deref(), Some("Georgia"));
}

/// A movie shape: its `OfficeArtClientData` names the object through an
/// `ExObjRefAtom`, and the `ExObjListContainer` holds an
/// `ExAviMovieContainer` whose `ExMediaAtom` carries that id. It reaches
/// the IR as a data-less image naming the object, as OLE objects do.
#[test]
fn test_media_shape_leaves_a_placeholder() {
    let media_atom = atom(0x1004, 0, &[&9u32.to_le_bytes()[..], &[0u8; 4]].concat());
    let movie = container(0x1006, 0, &container(0x1005, 0, &media_atom));
    let ex_obj_list = container(0x0409, 0, &movie);
    let obj_ref = atom(0x0BC1, 0, &9u32.to_le_bytes());
    let mut shapes = text_shape(1, "Watch this", &[]);
    shapes.extend(container(RT_SHAPE, 0, &container(0xF011, 0, &obj_ref)));
    let deck = PptBuilder {
        slides: vec![PptSlide {
            shapes,
            ..Default::default()
        }],
        doc_children: ex_obj_list,
        ..Default::default()
    };
    let ir = open(deck.build()).unwrap().to_ir();
    let alts: Vec<_> = ir.sections[0]
        .elements
        .iter()
        .filter_map(|e| match e {
            Element::Image(i) => i.alt_text.clone(),
            _ => None,
        })
        .collect();
    assert_eq!(alts, ["Embedded video"]);
}

/// A `TextCharsAtom` with its `StyleTextPropAtom` and a text-range
/// hyperlink over part of it: the formatting spans attach to the text and
/// survive the run being split around the link (`apply_style_text_prop`,
/// `slice_char_formats`, `slice_para_formats`), each piece keeping only
/// its own formatting.
#[test]
fn test_style_spans_attach_and_survive_a_text_range_hyperlink_split() {
    const CF_BOLD: u32 = 1;
    let text = "Click here now";
    // 6 plain, 4 bold ("here"), 5 plain (incl. the implicit final mark).
    let prop = style_text_prop(
        text.len() as u32,
        &[
            (6, CF_BOLD, 0u16.to_le_bytes().to_vec()),
            (4, CF_BOLD, 1u16.to_le_bytes().to_vec()),
            (5, CF_BOLD, 0u16.to_le_bytes().to_vec()),
        ],
    );
    // InteractiveInfo (II_HyperlinkAction = 4) + TextInteractiveInfoAtom
    // covering "here" (6..10), after the text atoms.
    let mut info = vec![0u8; 16];
    info[4..8].copy_from_slice(&1u32.to_le_bytes());
    info[8] = 4;
    let mut extra = prop;
    extra.extend(container(0x0FF2, 0, &atom(0x0FF3, 0, &info)));
    extra.extend(atom(0x0FDF, 0, &[&6i32.to_le_bytes()[..], &10i32.to_le_bytes()[..]].concat()));
    let mut link = atom(0x0FD3, 0, &1u32.to_le_bytes());
    link.extend(atom(RT_CSTRING, 1, &utf16("http://example.com/")));
    let deck = PptBuilder {
        slides: vec![PptSlide {
            shapes: text_shape(1, text, &extra),
            ..Default::default()
        }],
        doc_children: container(0x0409, 0, &container(0x0FD7, 0, &link)),
        ..Default::default()
    };
    let ir = open(deck.build()).unwrap().to_ir();
    // One paragraph: a link inside a sentence does not break the sentence.
    assert_eq!(texts(&ir.sections[0].elements), ["Click here now"]);
    let s = spans(&ir.sections[0].elements);
    let summary: Vec<(String, bool, bool)> = s
        .iter()
        .map(|t| (t.text.clone(), t.bold, t.hyperlink.is_some()))
        .collect();
    assert_eq!(
        summary,
        [
            ("Click ".to_string(), false, false),
            ("here".to_string(), true, true),
            (" now".to_string(), false, false),
        ]
    );
}

/// A "PowerPoint Document" stream whose sector chain ends before its
/// declared size was read short with no signal; the IR now says the text
/// is incomplete and names the stream.
#[test]
fn test_truncated_container_stream_marks_ppt_text_truncated() {
    let deck = PptBuilder {
        slides: vec![slide("Short")],
        ..Default::default()
    };
    let mut bytes = deck.build();
    let ir = open(bytes.clone()).unwrap().to_ir();
    assert!(!ir.metadata.text_truncated);
    assert!(ir.metadata.warnings.is_empty(), "{:?}", ir.metadata.warnings);
    // Directory entry 1 is "PowerPoint Document".
    let size_at = 512 + 128 + 0x78;
    let size = u32::from_le_bytes(bytes[size_at..size_at + 4].try_into().unwrap());
    bytes[size_at..size_at + 4].copy_from_slice(&(size + 4096).to_le_bytes());
    let ir = open(bytes).unwrap().to_ir();
    assert!(ir.metadata.text_truncated);
    assert!(
        ir.metadata
            .warnings
            .iter()
            .any(|w| w.contains("PowerPoint Document")),
        "{:?}",
        ir.metadata.warnings
    );
}

// ---------------------------------------------------------------------------
// Slide-master text
// ---------------------------------------------------------------------------

const MASTER_STATIC: &str = "Text that I added to the master slide";
const MASTER_PROMPT: &str = "Klicken Sie, um das Titelformat zu bearbeiten";
const MASTER_BODY_PROMPT: &str = "Second level";

/// A main master holding a title placeholder prompt (an
/// `OEPlaceholderAtom` shape), a body-type prompt with no placeholder atom
/// (caught by its placeholder text type), a date field placeholder and a
/// static text box (`Tx_TYPE_OTHER`, no placeholder atom).
fn master_shapes() -> Vec<u8> {
    let mut m = placeholder_text_shape(0, 1, MASTER_PROMPT); // MasterTitle
    m.extend(text_shape(1, MASTER_BODY_PROMPT, &[])); // Tx_TYPE_BODY
    m.extend(placeholder_text_shape(4, 7, "*")); // MasterDate
    m.extend(text_shape(4, MASTER_STATIC, &[])); // Tx_TYPE_OTHER
    m
}

fn master_deck(slides: Vec<PptSlide>) -> PptBuilder {
    PptBuilder {
        slides,
        masters: vec![master_shapes()],
        ..Default::default()
    }
}

fn on_master(text: &str, follow: bool) -> PptSlide {
    PptSlide {
        master: Some(0),
        follow_master_objects: follow,
        ..slide(text)
    }
}

/// Every surface of a document, for "appears exactly once" checks.
fn surfaces(doc: &Document) -> [(&'static str, String); 5] {
    let ir = doc.to_ir();
    [
        ("plain_text", doc.plain_text()),
        ("to_markdown", doc.to_markdown()),
        ("ir.plain_text", ir.plain_text()),
        ("ir.to_markdown", ir.to_markdown()),
        ("to_html", doc.to_html()),
    ]
}

/// Static text on a slide master is drawn on every slide that shows the
/// master's objects ([MS-PPT] `SlideAtom.slideFlags.fMasterObjects`), and
/// reference extractors include it. It is surfaced once per deck, in a
/// trailing "Slide Master" section, never once per slide; the master's
/// placeholder prompts ("Click to edit…", in any language) never are.
#[test]
fn test_ppt_master_static_text_appears_once_and_prompts_never() {
    let deck = master_deck(vec![on_master("Slide one", true), on_master("Slide two", true)]);
    let doc = open(deck.build()).unwrap();
    for (name, text) in surfaces(&doc) {
        assert_eq!(text.matches(MASTER_STATIC).count(), 1, "{name}: {text}");
        assert!(!text.contains(MASTER_PROMPT), "{name}: {text}");
        assert!(!text.contains(MASTER_BODY_PROMPT), "{name}: {text}");
        assert!(text.contains("Slide one") && text.contains("Slide two"), "{name}: {text}");
    }
    let ir = doc.to_ir();
    assert_eq!(ir.sections.len(), 3, "two slides plus the master section");
    let last = ir.sections.last().unwrap();
    assert_eq!(last.title.as_deref(), Some("Slide Master"));
    assert_eq!(texts(&last.elements), [MASTER_STATIC]);
}

/// A master whose objects every slide hides contributes nothing.
#[test]
fn test_ppt_master_text_is_absent_when_every_slide_hides_master_objects() {
    let deck = master_deck(vec![on_master("Slide one", false), on_master("Slide two", false)]);
    let doc = open(deck.build()).unwrap();
    for (name, text) in surfaces(&doc) {
        assert!(!text.contains(MASTER_STATIC), "{name}: {text}");
        assert!(!text.contains("Slide Master"), "{name}: {text}");
    }
    assert_eq!(doc.to_ir().sections.len(), 2);
}

/// One slide showing the master is enough.
#[test]
fn test_ppt_master_text_is_kept_when_one_slide_shows_master_objects() {
    let deck = master_deck(vec![on_master("Slide one", false), on_master("Slide two", true)]);
    let doc = open(deck.build()).unwrap();
    for (name, text) in surfaces(&doc) {
        assert_eq!(text.matches(MASTER_STATIC).count(), 1, "{name}: {text}");
    }
}

/// A deck whose only text is on the master: the slides have no text of
/// their own. The master text must still come out, and the slides keep
/// their places.
#[test]
fn test_ppt_master_text_survives_a_deck_whose_slides_have_no_text() {
    let blank = |follow| PptSlide {
        master: Some(0),
        follow_master_objects: follow,
        ..Default::default()
    };
    let deck = master_deck(vec![blank(true), blank(true)]);
    let doc = open(deck.build()).unwrap();
    for (name, text) in surfaces(&doc) {
        assert_eq!(text.matches(MASTER_STATIC).count(), 1, "{name}: {text}");
        assert!(!text.contains(MASTER_PROMPT), "{name}: {text}");
    }
    assert_eq!(doc.to_ir().sections.len(), 3);
}

/// `TxMasterStyleAtom` ([MS-PPT]) record type.
const RT_TX_MASTER_STYLE_ATOM: u16 = 0x0FA3;

/// A `TxMasterStyleAtom` for `text_type` (< 5, so no per-level
/// `indentLevel` field) with one level setting alignment (`PFMasks` bit
/// 11) and font size (`CFMasks` bit 17).
fn master_style(text_type: u16, alignment: u16, font_size: i16) -> Vec<u8> {
    let mut b = 1u16.to_le_bytes().to_vec(); // cLevels
    b.extend_from_slice(&(1u32 << 11).to_le_bytes());
    b.extend_from_slice(&alignment.to_le_bytes());
    b.extend_from_slice(&(1u32 << 17).to_le_bytes());
    b.extend_from_slice(&font_size.to_le_bytes());
    atom(RT_TX_MASTER_STYLE_ATOM, text_type, &b)
}

/// A slide's title with no direct formatting takes its size from the
/// master's title style. `SlideAtom.masterIdRef` is the master's
/// `masterId` from the `MasterListWithTextContainer` ([MS-PPT]
/// `MasterIdRef`), which the builder keeps unrelated to its persist id;
/// resolving it as a persist id (high bit masked) found no master, so
/// real decks never inherited master formatting.
#[test]
fn test_ppt_master_style_resolves_master_id_through_the_master_list() {
    let deck = PptBuilder {
        slides: vec![PptSlide {
            shapes: text_shape(0, "Styled Title", &[]),
            master: Some(0),
            ..Default::default()
        }],
        masters: vec![master_style(0, 1, 44)], // Title: centred, 44pt
        ..Default::default()
    };
    let ir = open(deck.build()).unwrap().to_ir();
    let Element::Heading(h) = &ir.sections[0].elements[0] else {
        panic!("expected a heading: {:?}", ir.sections[0].elements);
    };
    let InlineContent::Text(span) = &h.content[0] else {
        panic!()
    };
    assert_eq!(span.text, "Styled Title");
    assert_eq!(span.font_size_half_pt, Some(88), "44pt from the master title style");
    assert_eq!(h.alignment, Some(office_oxide::ir::ParagraphAlignment::Center));
}
