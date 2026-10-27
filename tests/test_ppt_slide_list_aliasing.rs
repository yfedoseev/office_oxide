//! A legacy `.ppt` slide list whose entries carry distinct persist ids that
//! the persist directory maps to one slide must not copy that slide's text
//! once per entry. Built in code.

mod common;

use common::cfb_with_stream;
use common::ppt::{
    RT_DOCUMENT, RT_PERSIST_DIRECTORY_ATOM, RT_SHAPE, RT_SLIDE, RT_SLIDE_ATOM,
    RT_SLIDE_LIST_WITH_TEXT, RT_SLIDE_PERSIST_ATOM, RT_TEXT_CHARS, RT_USER_EDIT_ATOM, atom,
    container, utf16,
};
use office_oxide::{Document, DocumentFormat};
use std::io::Cursor;

/// `DocumentAtom` ([MS-PPT] 2.4.2).
const RT_DOCUMENT_ATOM: u16 = 0x03E9;
/// `OfficeArtDgContainer` ([MS-ODRAW] 2.2.13).
const RT_DG_CONTAINER: u16 = 0xF002;
/// `OfficeArtSpgrContainer` ([MS-ODRAW] 2.2.16).
const RT_SPGR_CONTAINER: u16 = 0xF003;
/// Instance of the `SlideListWithText` that lists the slides ([MS-PPT]
/// 2.4.14).
const SLWT_SLIDES: u16 = 0;

/// A deck with one slide of `chars` characters, named by `entries` list
/// entries whose persist ids are all different.
fn aliased_deck(entries: u32, chars: usize) -> Vec<u8> {
    let text = utf16(&"A".repeat(chars));
    let shape = container(RT_SHAPE, 0, &atom(RT_TEXT_CHARS, 0, &text));
    let slide = container(
        RT_SLIDE,
        0,
        &[
            atom(RT_SLIDE_ATOM, 0, &[0u8; 32]),
            container(RT_DG_CONTAINER, 0, &container(RT_SPGR_CONTAINER, 0, &shape)),
        ]
        .concat(),
    );
    let list: Vec<u8> = (0..entries)
        .map(|i| {
            let mut body = (2 + i).to_le_bytes().to_vec();
            body.extend_from_slice(&[0u8; 12]);
            atom(RT_SLIDE_PERSIST_ATOM, 0, &body)
        })
        .collect::<Vec<_>>()
        .concat();
    let doc = container(
        RT_DOCUMENT,
        0,
        &[
            atom(RT_DOCUMENT_ATOM, 0, &[0u8; 48]),
            container(RT_SLIDE_LIST_WITH_TEXT, SLWT_SLIDES, &list),
        ]
        .concat(),
    );
    let edit_len = 8 + 28;
    let dir_len = 8 + 4 + 4 * (1 + entries as usize);
    let dir_offset = doc.len() + edit_len;
    let slide_offset = dir_offset + dir_len;
    // `PersistDirectoryAtom`: one run starting at persist id 1 (the document)
    // followed by one offset per entry, all but the first being the slide's.
    let mut dir = (1u32 | ((1 + entries) << 20)).to_le_bytes().to_vec();
    dir.extend_from_slice(&0u32.to_le_bytes());
    for _ in 0..entries {
        dir.extend_from_slice(&(slide_offset as u32).to_le_bytes());
    }
    let mut edit = vec![0u8; 28];
    edit[12..16].copy_from_slice(&(dir_offset as u32).to_le_bytes());
    edit[16..20].copy_from_slice(&1u32.to_le_bytes());
    let mut stream = doc;
    stream.extend(atom(RT_USER_EDIT_ATOM, 0, &edit));
    stream.extend(atom(RT_PERSIST_DIRECTORY_ATOM, 0, &dir));
    stream.extend(slide);
    cfb_with_stream("PowerPoint Document", &stream)
}

#[test]
fn test_ppt_list_entries_aliasing_one_slide_yield_it_once() {
    let file = aliased_deck(300, 20_000);
    let doc = Document::from_reader(Cursor::new(file), DocumentFormat::Ppt).expect("opens");
    let ir = doc.to_ir();
    assert_eq!(ir.sections.len(), 1, "one slide, not one per list entry");
    assert!(doc.plain_text().len() < 100_000, "{}", doc.plain_text().len());
    assert!(ir.metadata.text_truncated);
    assert!(
        ir.metadata
            .warnings
            .iter()
            .any(|w| w.contains("299 slide-list entries")),
        "{:?}",
        ir.metadata.warnings
    );
}

#[test]
fn test_ppt_a_single_entry_is_untouched() {
    let doc = Document::from_reader(Cursor::new(aliased_deck(1, 50)), DocumentFormat::Ppt)
        .expect("opens");
    assert_eq!(doc.to_ir().sections.len(), 1);
    assert!(!doc.to_ir().metadata.text_truncated);
}
