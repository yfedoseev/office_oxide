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

/// Wingdings, Wingdings 2 and 3 and Webdings are read by their own tables
/// (the Unicode compatibility mappings), not as Latin-1: a check mark is a
/// check mark, not `ü`. The values are the characters of Unicode's mapping.
#[test]
fn test_the_other_symbol_fonts_use_their_own_tables() {
    assert_eq!(text_of(&sym("Wingdings", "F0FC")), "\u{2714}"); // heavy check mark, not u-umlaut
    assert_eq!(text_of(&sym("Wingdings", "F0A7")), "\u{25AA}"); // small square, not a bullet
    assert_eq!(text_of(&sym("Wingdings", "F06E")), "\u{25A0}"); // black square
    assert_eq!(text_of(&sym("Wingdings", "F0D8")), "\u{2B9A}"); // arrowhead
    assert_eq!(text_of(&sym("Wingdings", "F0B7")), "\u{1F550}"); // clock face, not a bullet
    assert_eq!(text_of(&sym("Wingdings 2", "F052")), "\u{2611}"); // ballot box with check
    assert_eq!(text_of(&sym("Wingdings 2", "F050")), "\u{2713}"); // check mark
    assert_eq!(text_of(&sym("Wingdings 3", "F070")), "\u{25B2}"); // up triangle
    assert_eq!(text_of(&sym("Webdings", "F061")), "\u{2714}"); // check mark
    // The font name is matched without regard to case, spacing or a trailing space.
    assert_eq!(text_of(&sym("WINGDINGS 2", "F052")), "\u{2611}");
    assert_eq!(text_of(&sym("Wingdings2", "F052")), "\u{2611}");
    assert_eq!(text_of(&sym(" Wingdings ", "F0FC")), "\u{2714}");
}

/// A code the font has no character for (its private-use placeholders) is the
/// replacement character, not a Latin letter; the space is a space.
#[test]
fn test_a_symbol_font_code_without_a_character_is_the_replacement_character() {
    assert_eq!(text_of(&sym("Wingdings", "F0FF")), "\u{fffd}");
    assert_eq!(text_of(&sym("Wingdings 3", "F0A0")), "\u{fffd}");
    assert_eq!(text_of(&sym("Wingdings", "F020")), " ".trim());
    let runs =
        format!("<w:r><w:t>a</w:t></w:r>{}<w:r><w:t>b</w:t></w:r>", sym("Wingdings", "F020"));
    assert_eq!(text_of(&runs), "a b");
}

/// A font that is not a symbol font has no table here: a code outside the
/// private-use range passes through, and without a font the two bullet codes
/// Word's galleries use keep their old reading.
#[test]
fn test_fonts_without_a_table_keep_their_reading() {
    assert_eq!(text_of(&sym("Arial", "0041")), "A");
    assert_eq!(text_of(&sym("Arial", "2713")), "\u{2713}");
    let no_font = r#"<w:r><w:sym w:char="F0B7"/></w:r>"#;
    assert_eq!(text_of(no_font), "\u{2022}");
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
