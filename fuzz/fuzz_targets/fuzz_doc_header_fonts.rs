#![no_main]

// The `.doc` font table (`SttbfFfn`) names the fonts whose private-use text is
// read by a symbol table, and, when the header's version is destroyed,
// declares the character sets that give the 8-bit text its language. It sits
// behind a valid compound file, FIB and piece table. This target builds a
// well-formed one-paragraph `.doc` in code, then hands the fuzzer's bytes in as
// the font table and as the header's version and language, so every run reaches
// the font-table walk, the symbol runs and the destroyed-header inference.
// None may panic, overflow or hang.
#[path = "../../tests/common/mod.rs"]
mod common;

use std::io::Cursor;

use common::{FibTweaks, Para, Subdocs, build_doc_from_streams, build_doc_streams};
use libfuzzer_sys::fuzz_target;
use office_oxide::{Document, DocumentFormat};

fuzz_target!(|data: &[u8]| {
    if data.len() < 5 {
        return;
    }
    let (head, ffn) = data.split_at(5);
    let n_fib = u16::from_le_bytes([head[0], head[1]]);
    let lid = u16::from_le_bytes([head[2], head[3]]);
    let chpx_ftc = head[4];
    let tweaks = FibTweaks {
        chpx_runs: 1,
        // `sprmCRgFtc2` naming a font by the fuzzer's index.
        chpx_grpprl: vec![0x51, 0x4A, chpx_ftc, 0],
        ..FibTweaks::default()
    };
    let (mut word, mut table) = build_doc_streams(
        &[Para {
            text: "\u{F06D}M and more text",
            terminator: '\r',
            grpprl: Vec::new(),
        }],
        &Subdocs::default(),
        tweaks,
    );
    word[0x112..0x116].copy_from_slice(&(table.len() as u32).to_le_bytes());
    word[0x116..0x11A].copy_from_slice(&(ffn.len() as u32).to_le_bytes());
    table.extend_from_slice(ffn);
    word[2..4].copy_from_slice(&n_fib.to_le_bytes());
    word[6..8].copy_from_slice(&lid.to_le_bytes());
    let bytes = build_doc_from_streams(&word, &table);
    if let Ok(doc) = Document::from_reader(Cursor::new(bytes), DocumentFormat::Doc) {
        let _ = doc.plain_text();
        let _ = doc.to_markdown();
        let _ = doc.to_ir();
    }
});
