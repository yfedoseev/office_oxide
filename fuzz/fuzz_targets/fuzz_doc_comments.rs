#![no_main]

// The legacy `.doc` reader splits the merged comment story into comments by
// the boundary CPs of `PlcfandTxt`. Whole-file fuzzing rarely reaches that
// behind a valid CFB container, FIB, piece table and a matching `PlcfandRef`.
// This target builds a well-formed `.doc` with a comment story in code and
// lets the fuzzer choose the boundary CPs (as little-endian u16s, so they land
// inside, on and past the story): backwards, repeated and out-of-range
// boundaries must neither panic nor make the output grow with the number of
// ranges.
#[path = "../../tests/common/mod.rs"]
mod common;

use std::io::Cursor;

use common::{
    FibTweaks, Para, Subdocs, build_doc_streams, cfb_with_streams, prose_grpprl, with_comment_plcs,
};
use libfuzzer_sys::fuzz_target;
use office_oxide::{Document, DocumentFormat};

const STORY: &str = "comment text of a few words, long enough to be cut in many places";

fuzz_target!(|data: &[u8]| {
    let cps: Vec<u32> = data
        .chunks_exact(2)
        .take(2_000)
        .map(|c| u32::from(u16::from_le_bytes([c[0], c[1]])) % 100)
        .collect();
    if cps.len() < 3 {
        return;
    }
    let paras = vec![Para { text: "body", terminator: '\r', grpprl: prose_grpprl() }];
    let subdocs = Subdocs { comments: STORY, ..Default::default() };
    let (mut word_doc, mut table) = build_doc_streams(&paras, &subdocs, FibTweaks::default());
    with_comment_plcs(&mut word_doc, &mut table, &cps);
    let bytes = cfb_with_streams(&[("WordDocument", &word_doc), ("0Table", &table)]);
    if let Ok(doc) = Document::from_reader(Cursor::new(bytes), DocumentFormat::Doc) {
        let ir = doc.to_ir();
        // The story is copied at most once however the boundaries fall.
        assert!(ir.plain_text().len() < 64 * STORY.len());
    }
});
