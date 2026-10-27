//! A Word 97 piece table that points many pieces at the same bytes must not
//! multiply a small file into a huge text. Built in code.

mod common;

use common::{FibTweaks, Para, Subdocs, build_doc_streams, cfb_with_streams, prose_grpprl};
use office_oxide::{Document, DocumentFormat};
use std::io::Cursor;

/// A 60,000-character paragraph piled `piles` times, with `ccpText` patched
/// to cover every pile.
fn piled_doc(piles: usize) -> Vec<u8> {
    let text: &'static str = Box::leak("A".repeat(60_000).into_boxed_str());
    let paras = vec![Para {
        text,
        terminator: '\r',
        grpprl: prose_grpprl(),
    }];
    let tweaks = FibTweaks {
        piled_pieces: piles,
        ..Default::default()
    };
    let (mut word_doc, table) = build_doc_streams(&paras, &Subdocs::default(), tweaks);
    let total = (60_001u64 * piles as u64) as u32;
    word_doc[0x4C..0x50].copy_from_slice(&total.to_le_bytes());
    cfb_with_streams(&[("WordDocument", &word_doc), ("0Table", &table)])
}

#[test]
fn test_word97_pieces_reading_far_more_than_the_stream_holds_are_refused() {
    let file = piled_doc(100);
    let err = Document::from_reader(Cursor::new(file), DocumentFormat::Doc)
        .err()
        .expect("100 pieces over the same 60,000 characters must be refused");
    assert!(err.to_string().contains("overlap"), "{err}");
}

#[test]
fn test_word97_pieces_that_share_some_bytes_are_still_read() {
    let file = piled_doc(3);
    let doc = Document::from_reader(Cursor::new(file), DocumentFormat::Doc).expect("opens");
    assert!(doc.plain_text().contains('A'));
}
