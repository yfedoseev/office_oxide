//! A `PlcfandTxt` whose boundary CPs run backwards must not make the reader
//! copy the comment text once per range. Built in code.

mod common;

use common::{
    FibTweaks, Para, Subdocs, build_doc_streams, cfb_with_streams, prose_grpprl, with_comment_plcs,
};
use office_oxide::{Document, DocumentFormat};
use std::io::Cursor;

const BODY_CHARS: usize = 20_000;

/// A document whose comment story is `BODY_CHARS` characters long, with the
/// given `PlcfandTxt` boundaries.
fn doc_with_comment_cps(cps: &[u32]) -> Vec<u8> {
    let comments: &'static str = Box::leak("c".repeat(BODY_CHARS).into_boxed_str());
    let paras = vec![Para {
        text: "body",
        terminator: '\r',
        grpprl: prose_grpprl(),
    }];
    let subdocs = Subdocs {
        comments,
        ..Default::default()
    };
    let (mut word_doc, mut table) = build_doc_streams(&paras, &subdocs, FibTweaks::default());
    with_comment_plcs(&mut word_doc, &mut table, cps);
    cfb_with_streams(&[("WordDocument", &word_doc), ("0Table", &table)])
}

fn open(bytes: Vec<u8>) -> Document {
    Document::from_reader(Cursor::new(bytes), DocumentFormat::Doc).expect("opens")
}

/// Boundaries 0, B, 0, B, ... : every range after the first runs backwards or
/// covers the whole story again.
fn alternating(n: usize) -> Vec<u32> {
    let mut cps: Vec<u32> = (0..=n)
        .map(|i| if i % 2 == 0 { 0 } else { BODY_CHARS as u32 })
        .collect();
    cps.push(0);
    cps
}

#[test]
fn test_backwards_comment_boundaries_do_not_multiply_the_comment_text() {
    let doc = open(doc_with_comment_cps(&alternating(200)));
    let ir = doc.to_ir();
    let text = ir.plain_text();
    assert!(
        text.len() < 4 * BODY_CHARS,
        "the comment story was copied once per range: {} bytes",
        text.len()
    );
    assert!(
        ir.metadata.warnings.iter().any(|w| w.contains("comment")),
        "the refused split must be reported: {:?}",
        ir.metadata.warnings
    );
}

#[test]
fn test_forward_comment_boundaries_still_split_the_comments() {
    let third = (BODY_CHARS / 3) as u32;
    let cps = [0, third, 2 * third, BODY_CHARS as u32, 0];
    let doc = open(doc_with_comment_cps(&cps));
    let ir = doc.to_ir();
    assert!(
        !ir.metadata.warnings.iter().any(|w| w.contains("comment")),
        "{:?}",
        ir.metadata.warnings
    );
}

/// The same builder the structure-aware fuzz target uses, driven by a fixed
/// generator: boundaries that fall inside, on and past the story in any order
/// never panic, and the text never grows with the number of ranges.
#[test]
fn test_random_comment_boundaries_never_panic_or_multiply_the_text() {
    let mut state = 0x2545_F491_4F6C_DD1Du64;
    let mut next = move || {
        state ^= state << 13;
        state ^= state >> 7;
        state ^= state << 17;
        state
    };
    for round in 0..60 {
        let n = 3 + (next() % 40) as usize;
        let cps: Vec<u32> = (0..n)
            .map(|_| match next() % 4 {
                0 => 0,
                1 => BODY_CHARS as u32,
                2 => (next() % (2 * BODY_CHARS as u64)) as u32,
                _ => u32::MAX - (next() % 3) as u32,
            })
            .collect();
        let doc = open(doc_with_comment_cps(&cps));
        let text = doc.to_ir().plain_text();
        assert!(text.len() < 4 * BODY_CHARS, "round {round}: {cps:?} gave {} bytes", text.len());
    }
}
