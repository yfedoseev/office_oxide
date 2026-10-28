//! `<w:sym>` names a character of a symbol font by a private-use code
//! (`F000` + the font's byte). Read as Latin-1 it printed an unrelated letter
//! or sign: the Symbol font's nabla (0xD1) came out as `Ñ`, its club (0xA7) as
//! `§`. The Symbol font's own encoding decides what the code means.

use std::io::Cursor;

use office_oxide::core::opc::{OpcWriter, PartName};
use office_oxide::core::relationships::rel_types;
use office_oxide::{Document, DocumentFormat};

fn docx_with_runs(runs: &str) -> Vec<u8> {
    let body = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>{runs}</w:p></w:body></w:document>"#
    );
    let mut writer = OpcWriter::new(Cursor::new(Vec::new())).unwrap();
    let part = PartName::new("/word/document.xml").unwrap();
    writer
        .add_part(
            &part,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            body.as_bytes(),
        )
        .unwrap();
    writer.add_package_rel(rel_types::OFFICE_DOCUMENT, "word/document.xml");
    writer.finish().unwrap().into_inner()
}

fn text_of(runs: &str) -> String {
    let doc =
        Document::from_reader(Cursor::new(docx_with_runs(runs)), DocumentFormat::Docx).unwrap();
    doc.plain_text().trim().to_string()
}

fn sym(font: &str, code: &str) -> String {
    format!(r#"<w:r><w:sym w:font="{font}" w:char="{code}"/></w:r>"#)
}

#[test]
fn test_symbol_font_codes_use_the_symbol_encoding() {
    let runs = [
        sym("Symbol", "F0D1"), // nabla
        sym("Symbol", "F0A7"), // club
        sym("Symbol", "F061"), // alpha
        sym("Symbol", "F0B4"), // multiply
        sym("Symbol", "F0B7"), // bullet
        sym("Symbol", "F044"), // Delta (Greek capital, not the increment sign)
    ]
    .concat();
    assert_eq!(text_of(&runs), "\u{2207}\u{2663}\u{3b1}\u{d7}\u{2022}\u{394}");
}

#[test]
fn test_the_font_name_is_matched_without_regard_to_case() {
    assert_eq!(text_of(&sym("SYMBOL", "F0D1")), "\u{2207}");
    assert_eq!(text_of(&sym("symbol", "F0D1")), "\u{2207}");
}

/// A code the Symbol font has no glyph for is shown as the replacement
/// character, not as an unrelated Latin letter.
#[test]
fn test_a_symbol_code_without_a_glyph_is_not_read_as_latin_1() {
    assert_eq!(text_of(&sym("Symbol", "F060")), "\u{fffd}");
    assert_eq!(text_of(&sym("Symbol", "F0FF")), "\u{fffd}");
}

/// Other fonts keep their reading: Wingdings has no standard Unicode table,
/// so its bullets are still the round bullet and an unknown font passes the
/// code through as before.
#[test]
fn test_other_fonts_keep_their_reading() {
    assert_eq!(text_of(&sym("Wingdings", "F0B7")), "\u{2022}");
    assert_eq!(text_of(&sym("Wingdings", "F0D8")), "\u{25ba}");
    assert_eq!(text_of(&sym("Arial", "0041")), "A");
}

/// Text and symbols in one paragraph stay in order.
#[test]
fn test_a_symbol_between_words_keeps_its_place() {
    let runs = format!(
        "<w:r><w:t xml:space=\"preserve\">angle </w:t></w:r>{}<w:r><w:t xml:space=\"preserve\"> = 90</w:t></w:r>",
        sym("Symbol", "F0D0")
    );
    assert_eq!(text_of(&runs), "angle \u{2220} = 90");
}
