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
