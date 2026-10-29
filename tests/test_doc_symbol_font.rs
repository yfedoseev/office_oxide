//! Text typed in a symbol font in a `.doc`.
//!
//! Symbol, Wingdings and Webdings store their text as private-use characters
//! (`U+F000` plus the font's byte), so `μM` was written as `U+F06D` and `M`.
//! A character outside U+0020..U+007F takes its font from `sprmCRgFtc2`
//! ([MS-DOC] 2.6.1), an index into the font table (`SttbfFfn`), and in a run
//! whose font is a symbol font the characters are the font's.

mod common;

use common::{
    FibTweaks, Para, Style, Subdocs, build_doc_from_streams, build_doc_streams,
    build_streams_with_styles, default_limits, style_sheet,
};
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
    assert!(text_of(doc_with_first_char_in_font("\u{F0FC} ok", &fonts, 2)).contains("\u{2713} ok"));
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

/// Three paragraphs of `\u{F06D}M`, in the styles `istds`, whose font table holds
/// Times New Roman, Symbol and Wingdings. Styles: 0 Normal; 1 a paragraph
/// style in Symbol; 2 a paragraph style based on 1; 3 a character style in
/// Wingdings. `chpx` is the grpprl of the one run over the first character.
fn styled_doc(istds: &[u16], chpx: Vec<u8>) -> String {
    let paras: Vec<Para> = istds
        .iter()
        .map(|_| Para {
            text: "\u{F0FC}\u{F06D}M",
            terminator: '\r',
            grpprl: Vec::new(),
        })
        .collect();
    let stsh = style_sheet(&[
        Style::new(0, "Normal"),
        Style::based_on("Symbolic", 0).font(1),
        Style::based_on("Symbolic Too", 1),
        Style::based_on("Wing", 0).kind(2).font(2),
    ]);
    let tweaks = FibTweaks {
        chpx_runs: usize::from(!chpx.is_empty()),
        chpx_grpprl: chpx,
        ..FibTweaks::default()
    };
    let (mut word, mut table) = build_streams_with_styles(&paras, istds, &stsh, tweaks);
    let ffn = sttbf_ffn(&["Times New Roman", "Symbol", "Wingdings"]);
    word[0x112..0x116].copy_from_slice(&(table.len() as u32).to_le_bytes());
    word[0x116..0x11A].copy_from_slice(&(ffn.len() as u32).to_le_bytes());
    table.extend_from_slice(&ffn);
    text_of(build_doc_from_streams(&word, &table))
}

#[test]
fn test_a_paragraph_style_in_a_symbol_font_maps_runs_that_state_no_font() {
    let _defaults = default_limits();
    let text = styled_doc(&[0, 1, 2], Vec::new());
    let lines: Vec<&str> = text.lines().collect();
    assert_eq!(lines[0], "\u{F0FC}\u{F06D}M", "Normal names no symbol font");
    // Symbol: F0FC is the Symbol table's own character, F06D is mu.
    assert!(lines[1].contains('\u{3bc}') && !lines[1].contains('\u{F06D}'), "{lines:?}");
    assert_eq!(lines[1], lines[2], "a style based on a symbol style inherits its font");
}

#[test]
fn test_a_run_font_or_character_style_beats_the_paragraph_style() {
    let _defaults = default_limits();
    // The first character's own font is Times New Roman (index 0).
    let own = styled_doc(&[1], vec![0x51, 0x4A, 0x00, 0x00]);
    assert!(own.starts_with('\u{F0FC}') && own.contains('\u{3bc}'), "{own:?}");
    // The first character applies the Wingdings character style (istd 3):
    // F0FC is a check mark there, and the rest stays Symbol.
    let character_style = styled_doc(&[1], vec![0x30, 0x4A, 0x03, 0x00]);
    assert!(character_style.starts_with('\u{2713}') && character_style.contains('\u{3bc}'));
}
