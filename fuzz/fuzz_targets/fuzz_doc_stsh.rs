#![no_main]

// The legacy `.doc` style sheet (STSH) sits behind a valid CFB container, a
// valid FIB and a valid piece table; byte-level fuzzing of whole files almost
// never gets that far. This target builds a well-formed one-to-four paragraph
// `.doc` in code and hands the fuzzer's bytes in as the style sheet itself,
// plus the paragraphs' style indexes, so every run reaches the STSH parser, the
// `istdBase` walk and the heading resolution. None may panic, overflow or hang;
// a malformed sheet must degrade to "no style" and a warning.
#[path = "../../tests/common/mod.rs"]
mod common;

use std::io::Cursor;

use common::{Para, build_doc_with_styles};
use libfuzzer_sys::fuzz_target;
use office_oxide::{Document, DocumentFormat};

const TEXTS: [&str; 4] = ["Introduction.", "Second Heading", "Third Heading", "Fourth Heading"];

fuzz_target!(|data: &[u8]| {
    // byte 0: paragraph count (1-4); then one little-endian u16 istd each; the
    // rest is the style sheet.
    let Some((&count, rest)) = data.split_first() else {
        return;
    };
    let n = usize::from(count % 4) + 1;
    if rest.len() < 2 * n {
        return;
    }
    let (istd_bytes, stsh) = rest.split_at(2 * n);
    let istds: Vec<u16> = istd_bytes
        .chunks_exact(2)
        .map(|c| u16::from_le_bytes([c[0], c[1]]))
        .collect();
    let paras: Vec<Para> = (0..n)
        .map(|i| Para {
            text: TEXTS[i],
            terminator: '\r',
            grpprl: Vec::new(),
        })
        .collect();
    let bytes = build_doc_with_styles(&paras, &istds, stsh);
    if let Ok(doc) = Document::from_reader(Cursor::new(bytes), DocumentFormat::Doc) {
        let _ = doc.plain_text();
        let _ = doc.to_markdown();
        let _ = doc.to_ir();
    }
});
