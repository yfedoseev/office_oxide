//! A `.doc` whose header version is not a Word version.
//!
//! The language ID in the FIB decides how the file's 8-bit text is read. When
//! the header is damaged (the version, language, flags and character set all
//! overwritten) there is no language to use, and the text used to be read in
//! Windows-1252 with nothing said: a Hebrew file came out as Latin letters.
//! The reader does not guess a script from byte statistics; it says that the
//! header is damaged and what that means for the text.

mod common;

use common::{
    FibTweaks, Para, Subdocs, build_doc_from_streams, build_doc_streams, build_word6_doc,
    default_limits,
};
use office_oxide::{Document, DocumentFormat};
use std::io::Cursor;

fn warnings(bytes: Vec<u8>) -> (String, Vec<String>) {
    let doc = Document::from_reader(Cursor::new(bytes), DocumentFormat::Doc).expect("opens");
    (doc.plain_text(), doc.to_ir().metadata.warnings)
}

fn warns_of_damage(warnings: &[String]) -> bool {
    warnings
        .iter()
        .any(|w| w.contains("version") && w.contains("language") && w.contains("Windows-1252"))
}

fn word97_with(n_fib: u16, lid: u16) -> Vec<u8> {
    let (mut word, table) = build_doc_streams(
        &[Para {
            text: "Body text",
            terminator: '\r',
            grpprl: Vec::new(),
        }],
        &Subdocs::default(),
        FibTweaks::default(),
    );
    word[2..4].copy_from_slice(&n_fib.to_le_bytes());
    word[6..8].copy_from_slice(&lid.to_le_bytes());
    build_doc_from_streams(&word, &table)
}

/// A Word 6/95 file with `nFib` rewritten (the builder writes 101).
fn word6_with(n_fib: u16) -> Vec<u8> {
    let mut bytes = build_word6_doc(0xA5DC, b"Body text\r");
    let at = bytes
        .windows(4)
        .position(|w| w == [0xDC, 0xA5, 101, 0])
        .expect("the WordDocument stream starts with wIdent and nFib");
    bytes[at + 2..at + 4].copy_from_slice(&n_fib.to_le_bytes());
    bytes
}

#[test]
fn test_a_word97_header_with_a_destroyed_version_says_the_language_is_unknown() {
    let _defaults = default_limits();
    let (text, warnings) = warnings(word97_with(0x2020, 0x2020));
    assert!(text.contains("Body text"), "{text:?}");
    assert!(warns_of_damage(&warnings), "{warnings:?}");
    assert!(warnings.iter().any(|w| w.contains("0x2020")), "{warnings:?}");
}

#[test]
fn test_a_word6_header_with_a_destroyed_version_says_the_language_is_unknown() {
    let _defaults = default_limits();
    let (text, warnings) = warnings(word6_with(0x20FF));
    assert!(text.contains("Body text"), "{text:?}");
    assert!(warns_of_damage(&warnings), "{warnings:?}");
    assert!(warnings.iter().any(|w| w.contains("0x20FF")), "{warnings:?}");
}

/// Every version a real Word wrote carries no warning.
#[test]
fn test_genuine_versions_carry_no_header_warning() {
    let _defaults = default_limits();
    for n_fib in [0x00C1u16, 0x00C2, 0x0101, 0x010C, 0x0112] {
        let (_, warnings) = warnings(word97_with(n_fib, 0x0409));
        assert!(!warns_of_damage(&warnings), "{n_fib:#06x}: {warnings:?}");
    }
    for n_fib in [0x0065u16, 0x0068] {
        let (_, warnings) = warnings(word6_with(n_fib));
        assert!(!warns_of_damage(&warnings), "{n_fib:#06x}: {warnings:?}");
    }
}

/// A Word 97 file whose one piece is 8-bit text (`bytes`, then the paragraph
/// mark), with a font table whose fonts all declare character set `chs`, and a
/// header with the version and language `n_fib` / `lid`.
fn word97_8bit(bytes: &[u8], chs: u8, n_fib: u16, lid: u16) -> Vec<u8> {
    let text: &'static str = Box::leak("x".repeat(bytes.len()).into_boxed_str());
    let (mut word, mut table) = build_doc_streams(
        &[Para {
            text,
            terminator: '\r',
            grpprl: Vec::new(),
        }],
        &Subdocs::default(),
        FibTweaks::default(),
    );
    // Rewrite the text as 8-bit characters and the piece as compressed
    // ([MS-DOC] 2.9.73: `fc` carries the byte offset times two, with bit 30).
    let clx = u32::from_le_bytes(word[0x1A2..0x1A6].try_into().unwrap()) as usize;
    let pcd_fc = clx + 1 + 4 + 8 + 2;
    let fc_min = u32::from_le_bytes(table[pcd_fc..pcd_fc + 4].try_into().unwrap()) as usize;
    let mut eight_bit = bytes.to_vec();
    eight_bit.push(b'\r');
    word[fc_min..fc_min + eight_bit.len()].copy_from_slice(&eight_bit);
    table[pcd_fc..pcd_fc + 4].copy_from_slice(&((fc_min as u32 * 2) | 0x4000_0000).to_le_bytes());
    // A font table of two fonts with the character set.
    let mut ffn_table = 2u16.to_le_bytes().to_vec();
    ffn_table.extend_from_slice(&0u16.to_le_bytes());
    for name in ["Arial", "Times New Roman"] {
        let mut ffn = vec![0u8; 39];
        ffn[3] = chs;
        for unit in name.encode_utf16().chain(std::iter::once(0)) {
            ffn.extend_from_slice(&unit.to_le_bytes());
        }
        ffn_table.push(ffn.len() as u8);
        ffn_table.extend_from_slice(&ffn);
    }
    word[0x112..0x116].copy_from_slice(&(table.len() as u32).to_le_bytes());
    word[0x116..0x11A].copy_from_slice(&(ffn_table.len() as u32).to_le_bytes());
    table.extend_from_slice(&ffn_table);
    word[2..4].copy_from_slice(&n_fib.to_le_bytes());
    word[6..8].copy_from_slice(&lid.to_le_bytes());
    build_doc_from_streams(&word, &table)
}

/// "Привет" in Windows-1251.
const PRIVET_1251: [u8; 6] = [0xCF, 0xF0, 0xE8, 0xE2, 0xE5, 0xF2];

#[test]
fn test_a_destroyed_header_takes_the_language_from_the_fonts_character_set() {
    let _defaults = default_limits();
    let (text, warnings) = warnings(word97_8bit(&PRIVET_1251, 0xCC, 0x2020, 0x2020));
    assert!(text.contains("Привет"), "{text:?}");
    assert!(
        warnings
            .iter()
            .any(|w| w.contains("0x2020") && w.contains("0xCC") && w.contains("fonts declare")),
        "{warnings:?}"
    );
    assert!(
        !warns_of_damage(&warnings),
        "no longer says the language is unknown: {warnings:?}"
    );
}

#[test]
fn test_a_destroyed_header_with_western_fonts_still_reads_windows_1252_and_says_so() {
    let _defaults = default_limits();
    let (text, warnings) = warnings(word97_8bit(&PRIVET_1251, 0x00, 0x2020, 0x2020));
    assert!(text.contains("Ïðèâåò"), "{text:?}");
    assert!(warns_of_damage(&warnings), "{warnings:?}");
}

#[test]
fn test_an_intact_header_is_never_second_guessed_by_the_fonts() {
    let _defaults = default_limits();
    // A genuine English header whose fonts declare Cyrillic: the header wins.
    let (text, warnings) = warnings(word97_8bit(&PRIVET_1251, 0xCC, 0x00C1, 0x0409));
    assert!(text.contains("Ïðèâåò"), "{text:?}");
    assert!(
        !warns_of_damage(&warnings) && !warnings.iter().any(|w| w.contains("fonts declare")),
        "{warnings:?}"
    );
}
