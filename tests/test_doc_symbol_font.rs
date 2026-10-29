//! Text typed in a symbol font in a `.doc`.
//!
//! Symbol, Wingdings and Webdings store their text as private-use characters
//! (`U+F000` plus the font's byte), so `μM` was written as `U+F06D` and `M`.
//! A character outside U+0020..U+007F takes its font from `sprmCRgFtc2`
//! ([MS-DOC] 2.6.1), an index into the font table (`SttbfFfn`), and in a run
//! whose font is a symbol font the characters are the font's.

mod common;

use common::{FibTweaks, Para, Subdocs, build_doc_from_streams, build_doc_streams, default_limits};
use office_oxide::{Document, DocumentFormat};
use std::io::Cursor;

/// An `SttbfFfn` ([MS-DOC] 2.9.286) holding fonts named `names`: `cData`,
/// `cbExtra` (0), then each `FFN` behind a one-byte length. An `FFN` is 39
/// fixed bytes (`ffid`, `wWeight`, `chs`, `ixchSzAlt`, `panose`, `fs`) and the
/// UTF-16 name, null-terminated.
fn sttbf_ffn(names: &[&str]) -> Vec<u8> {
    let mut out = (names.len() as u16).to_le_bytes().to_vec();
    out.extend_from_slice(&0u16.to_le_bytes());
    for name in names {
        let mut ffn = vec![0u8; 39];
        for unit in name.encode_utf16().chain(std::iter::once(0)) {
            ffn.extend_from_slice(&unit.to_le_bytes());
        }
        out.push(ffn.len() as u8);
        out.extend_from_slice(&ffn);
    }
    out
}

/// A document of one paragraph, `text`, whose first character is a run with
/// `sprmCRgFtc2` = `ftc`, in a document whose font table holds `fonts`.
fn doc_with_first_char_in_font(text: &'static str, fonts: &[&str], ftc: u16) -> Vec<u8> {
    let tweaks = FibTweaks {
        chpx_runs: 1,
        chpx_grpprl: [&[0x51u8, 0x4A][..], &ftc.to_le_bytes()].concat(),
        ..FibTweaks::default()
    };
    let (mut word, mut table) = build_doc_streams(
        &[Para {
            text,
            terminator: '\r',
            grpprl: Vec::new(),
        }],
        &Subdocs::default(),
        tweaks,
    );
    let ffn = sttbf_ffn(fonts);
    word[0x112..0x116].copy_from_slice(&(table.len() as u32).to_le_bytes()); // fcSttbfFfn
    word[0x116..0x11A].copy_from_slice(&(ffn.len() as u32).to_le_bytes()); // lcbSttbfFfn
    table.extend_from_slice(&ffn);
    build_doc_from_streams(&word, &table)
}

fn text_of(bytes: Vec<u8>) -> String {
    Document::from_reader(Cursor::new(bytes), DocumentFormat::Doc)
        .unwrap()
        .plain_text()
}

#[test]
fn test_a_run_in_the_symbol_font_is_read_by_the_symbol_table() {
    let _defaults = default_limits();
    let bytes =
        doc_with_first_char_in_font("\u{F06D}M and more", &["Times New Roman", "Symbol"], 1);
    assert!(text_of(bytes).contains("\u{3bc}M and more"));
}

#[test]
fn test_a_wingdings_run_is_read_by_its_table_including_non_bmp_characters() {
    let _defaults = default_limits();
    let fonts = ["Arial", "Symbol", "Wingdings"];
    assert!(text_of(doc_with_first_char_in_font("\u{F0FC} ok", &fonts, 2)).contains("\u{2714} ok"));
    // Wingdings 0xB7 is a clock face outside the Basic Multilingual Plane.
    assert!(text_of(doc_with_first_char_in_font("\u{F0B7} t", &fonts, 2)).contains("\u{1F550} t"));
}

#[test]
fn test_a_run_in_another_font_keeps_its_private_use_character() {
    let _defaults = default_limits();
    let bytes = doc_with_first_char_in_font("\u{F06D}M", &["Times New Roman", "Symbol"], 0);
    assert!(text_of(bytes).contains("\u{F06D}M"));
}

#[test]
fn test_a_font_index_past_the_table_is_ignored() {
    let _defaults = default_limits();
    let bytes = doc_with_first_char_in_font("\u{F06D}M", &["Symbol"], 9);
    assert!(text_of(bytes).contains("\u{F06D}M"));
}

/// Every surface reads the run the same way.
#[test]
fn test_every_surface_reads_the_symbol_run_alike() {
    let _defaults = default_limits();
    let bytes =
        doc_with_first_char_in_font("\u{F06D}M and more", &["Times New Roman", "Symbol"], 1);
    let doc = Document::from_reader(Cursor::new(bytes), DocumentFormat::Doc).unwrap();
    for out in [doc.plain_text(), doc.to_markdown(), doc.to_html()] {
        assert!(out.contains('\u{3bc}') && !out.contains('\u{F06D}'), "{out}");
    }
}
