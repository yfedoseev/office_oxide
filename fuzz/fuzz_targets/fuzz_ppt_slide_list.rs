#![no_main]

// The legacy `.ppt` reader resolves each slide-list entry through the persist
// directory, reads a slide container once however many entries name it, and
// refuses containers that overlap an accepted one. Whole-file fuzzing rarely
// gets past the CFB container, the `UserEditAtom` chain and the directory to
// these decisions. This target builds a well-formed deck in code and lets the
// fuzzer choose how many slides there are, whether they nest inside one
// another, and which container (or none) each list entry's persist id maps
// to. None may panic, overflow or hang, and the text may not grow with the
// number of entries.
#[path = "../../tests/common/mod.rs"]
mod common;

use std::io::Cursor;

use common::ppt::aliasing_deck;
use libfuzzer_sys::fuzz_target;
use office_oxide::{Document, DocumentFormat};

fuzz_target!(|data: &[u8]| {
    // byte 0: slide count (1-8); byte 1 bit 0: nest the later slides in the
    // first; the rest: one entry target per byte (at most 500).
    let [slides, flags, targets @ ..] = data else {
        return;
    };
    let slides = usize::from(*slides % 8) + 1;
    let texts: Vec<String> = (0..slides)
        .map(|i| format!("slide {i} {}", "t".repeat(200)))
        .collect();
    let targets = &targets[..targets.len().min(500)];
    let bytes = aliasing_deck(&texts, flags & 1 == 1, targets);
    if let Ok(doc) = Document::from_reader(Cursor::new(bytes), DocumentFormat::Ppt) {
        let ir = doc.to_ir();
        // Each container is read at most once: the text is bounded by the
        // slides, not by the entries.
        assert!(ir.plain_text().len() < 64 * slides * 400 + 4_096);
    }
});
