//! A piece table whose character positions are not ascending.
//!
//! [MS-DOC] 2.8.35 requires the `PlcPcd` positions to ascend; 2.4.1 step 3 finds
//! the text at a position as "the largest i such that aCp[i] <= cp", which is
//! defined for any array. A fast-saved file with one wrong position used to be
//! refused as a whole. It is read by that rule and the warnings say the table
//! was out of order.

mod common;

use common::{FibTweaks, Para, Subdocs, build_doc_streams, default_limits};
use office_oxide::{Document, DocumentFormat};
use std::io::Cursor;

const TEXT: &str = "ABCDEFGH";

/// The streams of a one-paragraph document, and the `fc` its single piece
/// reads from (UTF-16, so the byte offset into the `WordDocument` stream).
fn streams() -> (Vec<u8>, Vec<u8>, u32, u32) {
    let (word, table) = build_doc_streams(
        &[Para {
            text: TEXT,
            terminator: '\r',
            grpprl: Vec::new(),
        }],
        &Subdocs::default(),
        FibTweaks::default(),
    );
    let fc_clx = u32::from_le_bytes(word[0x1A2..0x1A6].try_into().unwrap()) as usize;
    // CLX: 0x02, lcb (4), then the PlcPcd: two CPs (8 bytes), one Pcd (8 bytes).
    let pcd = fc_clx + 5 + 8;
    let fc = u32::from_le_bytes(table[pcd + 2..pcd + 6].try_into().unwrap());
    let cp_end = u32::from_le_bytes(table[fc_clx + 9..fc_clx + 13].try_into().unwrap());
    assert_eq!(fc & 0x4000_0000, 0, "the builder writes UTF-16 text");
    (word, table, fc, cp_end)
}

/// Replace the document's CLX with a `PlcPcd` over `cps` whose pieces read from
/// `fcs` (UTF-16 byte offsets), appended to the Table stream.
fn with_pieces(cps: &[u32], fcs: &[u32]) -> Vec<u8> {
    let (mut word, mut table, _, _) = streams();
    let n = fcs.len();
    assert_eq!(cps.len(), n + 1);
    let mut clx = vec![0x02];
    clx.extend_from_slice(&((cps.len() * 4 + n * 8) as u32).to_le_bytes());
    for cp in cps {
        clx.extend_from_slice(&cp.to_le_bytes());
    }
    for fc in fcs {
        clx.extend_from_slice(&0u16.to_le_bytes());
        clx.extend_from_slice(&fc.to_le_bytes());
        clx.extend_from_slice(&0u16.to_le_bytes());
    }
    let at = table.len() as u32;
    table.extend_from_slice(&clx);
    word[0x1A2..0x1A6].copy_from_slice(&at.to_le_bytes());
    word[0x1A6..0x1AA].copy_from_slice(&(clx.len() as u32).to_le_bytes());
    // The container writer used by the other tests.
    common::build_doc_from_streams(&word, &table)
}

fn open(bytes: Vec<u8>) -> Document {
    Document::from_reader(Cursor::new(bytes), DocumentFormat::Doc).expect("opens")
}

fn warned(doc: &Document) -> bool {
    doc.to_ir()
        .metadata
        .warnings
        .iter()
        .any(|w| w.contains("ascending"))
}

/// Pieces 0 and 1 are stored next to each other; piece 2's position (200) is
/// wrong. Piece 1 therefore runs to the last position, piece 2 covers nothing,
/// and the whole text is read.
#[test]
fn test_one_wrong_position_still_yields_the_whole_text_and_is_reported() {
    let _defaults = default_limits();
    let (_, _, fc, cp_end) = streams();
    let bytes = with_pieces(&[0, 3, 200, cp_end], &[fc, fc + 6, 0]);
    let doc = open(bytes);
    assert!(doc.plain_text().contains(TEXT), "{:?}", doc.plain_text());
    assert!(warned(&doc), "{:?}", doc.to_ir().metadata.warnings);
}

/// The same text through an ascending table: identical text, no warning.
#[test]
fn test_an_ascending_table_reads_the_same_text_without_a_warning() {
    let _defaults = default_limits();
    let (_, _, fc, cp_end) = streams();
    let doc = open(with_pieces(&[0, 3, cp_end], &[fc, fc + 6]));
    assert!(doc.plain_text().contains(TEXT));
    assert!(!warned(&doc), "{:?}", doc.to_ir().metadata.warnings);
}

/// Every position reversed: no piece covers any text, and the file opens with
/// no text and says the text is incomplete, instead of failing or panicking.
#[test]
fn test_a_table_with_every_position_reversed_opens_without_text() {
    let _defaults = default_limits();
    let doc = open(with_pieces(&[900, 500, 100], &[0, 6]));
    assert!(!doc.plain_text().contains(TEXT));
    assert!(warned(&doc));
}
