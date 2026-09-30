//! Smoke test for the in-code synthetic `.doc` writer (`tests/common/mod.rs`).
//!
//! Guards the writer itself — that it produces a CFB/OLE2 container the
//! parser accepts and that the round-trip recovers the text and basic
//! structure. These tests run as part of `cargo test` with no fixture blob.

mod common;

use common::{Para, build_doc, open_doc, prose_grpprl};

#[test]
fn test_synthetic_single_paragraph_round_trips() {
    let paras = [Para {
        text: "Hello from a synthetic doc.",
        terminator: '\r',
        grpprl: prose_grpprl(),
    }];
    let bytes = build_doc(&paras);
    let doc = open_doc(&bytes);

    // The parser must not error, and the text must survive.
    let text = doc.plain_text();
    assert!(text.contains("Hello from a synthetic doc."), "text lost: {text:?}");
}

#[test]
fn test_synthetic_doc_is_a_real_cfb() {
    let paras = [Para {
        text: "CFB container check.",
        terminator: '\r',
        grpprl: prose_grpprl(),
    }];
    let bytes = build_doc(&paras);
    // CFB magic signature (D0 CF 11 E0 A1 B1 1A E1).
    assert_eq!(&bytes[0..8], &[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
}

#[test]
fn test_synthetic_table_and_prose_coexist() {
    use common::{cell_grpprl, row_grpprl};
    use office_oxide::ir::Element;

    // One prose paragraph, then a 2×2 table, then another prose paragraph.
    let paras = vec![
        Para {
            text: "Before the table.",
            terminator: '\r',
            grpprl: prose_grpprl(),
        },
        Para {
            text: "X",
            terminator: '\u{7}',
            grpprl: cell_grpprl(),
        },
        Para {
            text: "Y",
            terminator: '\u{7}',
            grpprl: cell_grpprl(),
        },
        Para {
            text: "",
            terminator: '\u{7}',
            grpprl: row_grpprl(&[0i16, 2000, 4000], &[0, 0]),
        },
        Para {
            text: "A",
            terminator: '\u{7}',
            grpprl: cell_grpprl(),
        },
        Para {
            text: "B",
            terminator: '\u{7}',
            grpprl: cell_grpprl(),
        },
        Para {
            text: "",
            terminator: '\u{7}',
            grpprl: row_grpprl(&[0i16, 2000, 4000], &[0, 0]),
        },
        Para {
            text: "After the table.",
            terminator: '\r',
            grpprl: prose_grpprl(),
        },
    ];
    let bytes = build_doc(&paras);
    let doc = open_doc(&bytes);
    let ir = doc.to_ir();

    let tables = ir
        .sections
        .iter()
        .flat_map(|s| s.elements.iter())
        .filter(|e| matches!(e, Element::Table(_)))
        .count();
    assert_eq!(tables, 1, "exactly one table must be emitted");
}

/// A `WordDocument` stream whose sector chain ends before its declared
/// size was read short with no signal: the document converted as if
/// complete. The short read now marks the IR as truncated.
#[test]
fn test_truncated_container_stream_marks_doc_text_truncated() {
    let paras = [Para {
        text: "Short stream.",
        terminator: '\r',
        grpprl: prose_grpprl(),
    }];
    let mut bytes = build_doc(&paras);
    assert!(!open_doc(&bytes).to_ir().metadata.text_truncated);
    // Directory entry 1 is `WordDocument`; claim two more sectors than
    // its chain holds.
    let size_at = 512 + 128 + 0x78;
    let size = u32::from_le_bytes(bytes[size_at..size_at + 4].try_into().unwrap());
    bytes[size_at..size_at + 4].copy_from_slice(&(size + 1024).to_le_bytes());
    let ir = open_doc(&bytes).to_ir();
    assert!(ir.metadata.text_truncated);
}
