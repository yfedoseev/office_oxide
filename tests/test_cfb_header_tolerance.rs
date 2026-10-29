//! A compound file whose byte-order field is wrong but whose other header
//! fields are intact.
//!
//! [MS-CFB] 2.2 requires the field to be 0xFFFE when a file is written and gives
//! a reader no rule for another value. Files with only this field damaged exist
//! (a writer that swapped it, a fuzzed header), and other implementations read
//! them. The file is read as little-endian and the warnings say so.

mod common;
#[path = "common/ppt_builder.rs"]
mod ppt_builder;

use common::{Para, biff, build_doc, cfb_with_stream};
use office_oxide::{Document, DocumentFormat};
use std::io::Cursor;

const WARNING: &str = "byte-order field";

fn with_byte_order(mut bytes: Vec<u8>, order: u16) -> Vec<u8> {
    bytes[0x1C..0x1E].copy_from_slice(&order.to_le_bytes());
    bytes
}

fn open(bytes: Vec<u8>, format: DocumentFormat) -> Document {
    Document::from_reader(Cursor::new(bytes), format).expect("opens")
}

fn warned(doc: &Document) -> bool {
    doc.to_ir()
        .metadata
        .warnings
        .iter()
        .any(|w| w.contains(WARNING))
}

fn build_xls(text: &str) -> Vec<u8> {
    let bof = |kind: u16| {
        let mut b = 0x0600u16.to_le_bytes().to_vec();
        b.extend_from_slice(&kind.to_le_bytes());
        b.extend_from_slice(&[0u8; 12]);
        biff(0x0809, &b)
    };
    let mut s = bof(0x0005);
    let mut bs = 0u32.to_le_bytes().to_vec();
    bs.extend_from_slice(&[0, 0, 1, 0, b'S']);
    s.extend(biff(0x0085, &bs));
    s.extend(biff(0x000A, &[]));
    s.extend(bof(0x0010));
    let mut label = 0u16.to_le_bytes().to_vec();
    label.extend_from_slice(&0u16.to_le_bytes());
    label.extend_from_slice(&0u16.to_le_bytes());
    label.extend_from_slice(&(text.len() as u16).to_le_bytes());
    label.push(0);
    label.extend_from_slice(text.as_bytes());
    s.extend(biff(0x0204, &label));
    s.extend(biff(0x000A, &[]));
    cfb_with_stream("Workbook", &s)
}

fn doc_bytes() -> Vec<u8> {
    build_doc(&[Para {
        text: "Word body text",
        terminator: '\r',
        grpprl: Vec::new(),
    }])
}

#[test]
fn test_doc_with_a_swapped_byte_order_is_read_and_warns() {
    let doc = open(with_byte_order(doc_bytes(), 0xFEFF), DocumentFormat::Doc);
    assert!(doc.plain_text().contains("Word body text"));
    assert!(warned(&doc), "{:?}", doc.to_ir().metadata.warnings);
}

#[test]
fn test_doc_with_a_fuzzed_byte_order_and_minor_version_is_read_and_warns() {
    let mut bytes = doc_bytes();
    bytes[0x18..0x1C].copy_from_slice(&[0x20, 0x20, 0x03, 0x00]); // minor 0x2020, major 3
    let doc = open(with_byte_order(bytes, 0xFF20), DocumentFormat::Doc);
    assert!(doc.plain_text().contains("Word body text"));
    assert!(warned(&doc));
}

#[test]
fn test_xls_with_a_swapped_byte_order_is_read_and_warns() {
    let doc = open(with_byte_order(build_xls("sheet cell"), 0xFEFF), DocumentFormat::Xls);
    assert!(doc.plain_text().contains("sheet cell"));
    assert!(warned(&doc), "{:?}", doc.to_ir().metadata.warnings);
}

#[test]
fn test_ppt_with_a_swapped_byte_order_is_read_and_warns() {
    let bytes = ppt_builder::build_ppt(&["Slide one text"], &[]);
    let doc = open(with_byte_order(bytes, 0xFEFF), DocumentFormat::Ppt);
    assert!(doc.plain_text().contains("Slide one text"));
    assert!(warned(&doc), "{:?}", doc.to_ir().metadata.warnings);
}

/// A conforming file carries no such warning.
#[test]
fn test_a_conforming_file_has_no_byte_order_warning() {
    for (bytes, format) in [
        (doc_bytes(), DocumentFormat::Doc),
        (build_xls("sheet cell"), DocumentFormat::Xls),
        (ppt_builder::build_ppt(&["Slide one text"], &[]), DocumentFormat::Ppt),
    ] {
        assert!(!warned(&open(bytes, format)), "{format:?}");
    }
}

/// The tolerance covers the marker only: a file damaged elsewhere in its
/// header is still refused.
#[test]
fn test_a_wrong_marker_does_not_excuse_a_damaged_sector_shift() {
    let mut bytes = with_byte_order(doc_bytes(), 0xFEFF);
    bytes[0x1E..0x20].copy_from_slice(&14u16.to_le_bytes());
    assert!(Document::from_reader(Cursor::new(bytes), DocumentFormat::Doc).is_err());
}
