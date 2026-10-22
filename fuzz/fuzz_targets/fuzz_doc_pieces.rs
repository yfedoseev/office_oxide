#![no_main]

// The legacy `.doc` paragraph builder maps property runs to the pieces they
// overlap, and table spans to a column grid built from every row's boundaries.
// Whole-file fuzzing rarely reaches either behind a valid CFB container, FIB and
// piece table. This target builds a well-formed table `.doc` in code with the
// fuzzer choosing how many pieces pile on the same bytes, how many rows there
// are, and each row's `rgdxaCenter` boundaries. None may panic, overflow or hang.
#[path = "../../tests/common/mod.rs"]
mod common;

use std::io::Cursor;

use common::{FibTweaks, Para, Subdocs, build_doc_full, cell_grpprl, row_grpprl};
use libfuzzer_sys::fuzz_target;
use office_oxide::{Document, DocumentFormat};

fuzz_target!(|data: &[u8]| {
    // bytes 0-1: piled pieces (0-399) and row count (1-24); the rest are
    // little-endian i16 boundaries dealt out to the rows, up to eight a row.
    let [pieces, rows, rest @ ..] = data else {
        return;
    };
    let piled_pieces = usize::from(*pieces) * 400 / 256;
    let rows = usize::from(*rows % 24) + 1;
    let mut centers = rest
        .chunks_exact(2)
        .map(|c| i16::from_le_bytes([c[0], c[1]]));
    let mut paras = Vec::new();
    for _ in 0..rows {
        let mut row: Vec<i16> = centers.by_ref().take(8).collect();
        row.sort_unstable();
        row.dedup();
        if row.len() < 2 {
            row = vec![0, 1000];
        }
        let cells = row.len() - 1;
        for _ in 0..cells {
            paras.push(Para { text: "A", terminator: '\u{7}', grpprl: cell_grpprl() });
        }
        paras.push(Para {
            text: "",
            terminator: '\u{7}',
            grpprl: row_grpprl(&row, &vec![0; cells]),
        });
    }
    let tweaks = FibTweaks { piled_pieces, ..Default::default() };
    let bytes = build_doc_full(&paras, &Subdocs::default(), tweaks);
    if let Ok(doc) = Document::from_reader(Cursor::new(bytes), DocumentFormat::Doc) {
        let _ = doc.plain_text();
        let _ = doc.to_markdown();
        let _ = doc.to_ir();
    }
});
