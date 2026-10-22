//! Integration tests for legacy `.doc` header rows and huge-PAPX paragraphs.
//!
//! Documents are built **in code** from the synthetic `.doc` writer in
//! `tests/common/mod.rs` — no third-party fixture is committed.
//!
//! Two defect classes are guarded:
//!
//! - a row's `sprmTTableHeader` ([MS-DOC] §2.6.3, `0x3404`) never reached
//!   `TableRow::is_header` / `repeat_as_header`;
//! - a paragraph whose PAPX holds only `sprmPHugePapx` ([MS-DOC] §2.6.2,
//!   `0x6646`) keeps its real property list in the `Data` stream
//!   (`PrcData`, [MS-DOC] §2.9.210). Ignoring it lost the paragraph's
//!   table flags: a row mark became a plain paragraph and the table was
//!   split in two at that row.

mod common;

use common::{
    FibTweaks, Para, Subdocs, build_doc, build_doc_streams, cell_grpprl, cfb_with_streams,
    row_grpprl,
};
use office_oxide::ir::{DocumentIR, Element, Table};
use office_oxide::{Document, DocumentFormat};

const CENTERS: [i16; 3] = [0, 1500, 3000];

/// `sprmTTableHeader` (`0x3404`) = 1.
const TABLE_HEADER_ON: [u8; 3] = [0x04, 0x34, 0x01];

fn cell(text: &'static str) -> Para {
    Para {
        text,
        terminator: '\u{7}',
        grpprl: cell_grpprl(),
    }
}

fn row_mark(grpprl: Vec<u8>) -> Para {
    Para {
        text: "",
        terminator: '\u{7}',
        grpprl,
    }
}

/// A row-end grpprl: the usual TAP, then `sprmTTableHeader` = 1.
fn header_row_grpprl() -> Vec<u8> {
    let mut g = row_grpprl(&CENTERS, &[0, 0]);
    g.extend_from_slice(&TABLE_HEADER_ON);
    g
}

/// A PAPX holding only `sprmPHugePapx` (`0x6646`) naming `fc` in `Data`.
fn huge_papx_grpprl(fc: u32) -> Vec<u8> {
    let mut g = vec![0x46, 0x66];
    g.extend_from_slice(&fc.to_le_bytes());
    g
}

/// A `Data` stream with `PrcData` (`cbGrpprl: i16` + the grpprl) at `fc`.
fn data_with_prc_data(fc: usize, grpprl: &[u8]) -> Vec<u8> {
    let mut data = vec![0xEE; fc];
    data.extend_from_slice(&(grpprl.len() as i16).to_le_bytes());
    data.extend_from_slice(grpprl);
    data
}

/// Header row `H1 | H2`, then body row `a | b`. `header_mark` is the PAPX of
/// the header row's end mark.
fn two_row_paras(header_mark: Vec<u8>) -> Vec<Para> {
    vec![
        cell("H1"),
        cell("H2"),
        row_mark(header_mark),
        cell("a"),
        cell("b"),
        row_mark(row_grpprl(&CENTERS, &[0, 0])),
    ]
}

fn open_with_data(paras: &[Para], data: &[u8]) -> Document {
    let (word_doc, table) = build_doc_streams(paras, &Subdocs::default(), FibTweaks::default());
    let bytes = cfb_with_streams(&[
        ("WordDocument", &word_doc),
        ("0Table", &table),
        ("Data", data),
    ]);
    Document::from_reader(std::io::Cursor::new(bytes), DocumentFormat::Doc).expect("open .doc")
}

fn tables(ir: &DocumentIR) -> Vec<&Table> {
    ir.sections
        .iter()
        .flat_map(|s| s.elements.iter())
        .filter_map(|e| match e {
            Element::Table(t) => Some(t),
            _ => None,
        })
        .collect()
}

#[test]
fn test_doc_table_header_sprm_sets_is_header_on_that_row_only() {
    let bytes = build_doc(&two_row_paras(header_row_grpprl()));
    let doc = common::open_doc(&bytes);
    let ir = doc.to_ir();

    let tables = tables(&ir);
    assert_eq!(tables.len(), 1, "one logical table");
    let rows = &tables[0].rows;
    assert_eq!(rows.len(), 2);
    assert!(rows[0].is_header, "the row carrying sprmTTableHeader is a header row");
    assert!(rows[0].repeat_as_header, "a header row repeats, as in .docx");
    assert!(!rows[1].is_header, "the body row is not a header row");
    assert!(!rows[1].repeat_as_header);
}

#[test]
fn test_doc_table_without_header_sprm_has_no_header_row() {
    let bytes = build_doc(&two_row_paras(row_grpprl(&CENTERS, &[0, 0])));
    let doc = common::open_doc(&bytes);
    let ir = doc.to_ir();

    let tables = tables(&ir);
    assert_eq!(tables.len(), 1);
    assert!(
        tables[0]
            .rows
            .iter()
            .all(|r| !r.is_header && !r.repeat_as_header)
    );
}

#[test]
fn test_doc_huge_papx_row_mark_keeps_the_table_whole() {
    // The header row's end mark stores its property list in the Data stream.
    let fc = 0x40usize;
    let data = data_with_prc_data(fc, &header_row_grpprl());
    let doc = open_with_data(&two_row_paras(huge_papx_grpprl(fc as u32)), &data);
    let ir = doc.to_ir();

    let tables = tables(&ir);
    assert_eq!(tables.len(), 1, "a huge-PAPX row mark must not split the table");
    let rows = &tables[0].rows;
    assert_eq!(rows.len(), 2);
    assert!(rows[0].is_header, "the flag inside the PrcData reaches the row");
    assert!(!rows[1].is_header);
    assert_eq!(rows[0].cells.len(), 2);
}

#[test]
fn test_doc_huge_papx_pointing_outside_data_is_reported_not_fatal() {
    let data = data_with_prc_data(0x10, &header_row_grpprl());
    // fc lies far past the end of the Data stream.
    let doc = open_with_data(&two_row_paras(huge_papx_grpprl(0x00FF_0000)), &data);
    let ir = doc.to_ir();

    assert!(
        ir.metadata
            .warnings
            .iter()
            .any(|w| w.contains("sprmPHugePapx")),
        "an unresolvable huge PAPX must surface as a warning, got {:?}",
        ir.metadata.warnings
    );
}

#[test]
fn test_doc_huge_papx_with_no_data_stream_is_reported_not_fatal() {
    let (word_doc, table) = build_doc_streams(
        &two_row_paras(huge_papx_grpprl(0x40)),
        &Subdocs::default(),
        FibTweaks::default(),
    );
    let bytes = cfb_with_streams(&[("WordDocument", &word_doc), ("0Table", &table)]);
    let doc =
        Document::from_reader(std::io::Cursor::new(bytes), DocumentFormat::Doc).expect("open .doc");
    let ir = doc.to_ir();

    assert!(
        ir.metadata
            .warnings
            .iter()
            .any(|w| w.contains("sprmPHugePapx"))
    );
}

#[test]
fn test_doc_huge_papx_inside_prc_data_is_not_followed_again() {
    // A PrcData that itself names a huge PAPX at its own offset must not
    // recurse: the expansion is one level deep, the inner reference is left
    // alone, and the table stays whole.
    let fc = 0x20usize;
    let data = data_with_prc_data(fc, &huge_papx_grpprl(fc as u32));
    let doc = open_with_data(&two_row_paras(huge_papx_grpprl(fc as u32)), &data);
    let ir = doc.to_ir();
    assert!(
        !ir.metadata
            .warnings
            .iter()
            .any(|w| w.contains("sprmPHugePapx")),
        "a resolvable reference is not a warning: {:?}",
        ir.metadata.warnings
    );
}

#[test]
fn test_doc_huge_papx_word97_opcode_keeps_the_table_whole() {
    // The same sprm under its Word 97 number (0x6645) must resolve too.
    let fc = 0x40usize;
    let data = data_with_prc_data(fc, &header_row_grpprl());
    let mut g = vec![0x45, 0x66];
    g.extend_from_slice(&(fc as u32).to_le_bytes());
    let doc = open_with_data(&two_row_paras(g), &data);
    let ir = doc.to_ir();
    let tables = tables(&ir);
    assert_eq!(tables.len(), 1);
    assert!(tables[0].rows[0].is_header);
}

#[test]
fn test_doc_huge_papx_expansion_is_bounded_by_the_data_stream() {
    // Overlapping PrcData: every even offset of this 40 KB stream is a valid
    // maximum-size PrcData, so distinct references can name far more bytes
    // than the stream holds. The total expansion is budgeted against the size
    // of the stream, and the excess is reported.
    let data = [0xA2u8, 0x3F].repeat(20_000); // cbGrpprl = 0x3FA2 at every even fc
    let paras: Vec<Para> = (0..200u32)
        .map(|i| row_mark(huge_papx_grpprl(i * 2)))
        .collect();
    let doc = open_with_data(&paras, &data);
    let ir = doc.to_ir();
    assert!(
        ir.metadata
            .warnings
            .iter()
            .any(|w| w.contains("sprmPHugePapx")),
        "an over-budget expansion must be reported, got {:?}",
        ir.metadata.warnings
    );
}

#[test]
fn test_doc_huge_papx_shared_by_many_rows_is_expanded_once_and_keeps_one_table() {
    // A long, wide table: every row end names the same PrcData, which holds
    // the row definition plus a few kilobytes of other properties. The Data
    // stream stores it once, so the expansion budget is spent once, not once
    // per row.
    let mut prc = row_grpprl(&CENTERS, &[0, 0]);
    prc.extend_from_slice(&[0x03, 0x24, 0x00].repeat(1_500)); // sprmPJc80 padding
    let data = data_with_prc_data(0, &prc);
    let mut paras = Vec::new();
    for _ in 0..1_000 {
        paras.push(cell("a"));
        paras.push(cell("b"));
        paras.push(row_mark(huge_papx_grpprl(0)));
    }
    let doc = open_with_data(&paras, &data);
    let ir = doc.to_ir();
    let tables = tables(&ir);
    assert_eq!(tables.len(), 1, "rows sharing one PrcData must stay one table");
    assert_eq!(tables[0].rows.len(), 1_000);
    assert!(
        ir.metadata
            .warnings
            .iter()
            .all(|w| !w.contains("sprmPHugePapx")),
        "a shared PrcData is within budget, got {:?}",
        ir.metadata.warnings
    );
}

#[test]
fn test_doc_huge_papx_negative_length_is_rejected() {
    let fc = 0x08usize;
    let mut data = vec![0u8; fc];
    data.extend_from_slice(&(-1i16).to_le_bytes());
    data.extend_from_slice(&[0u8; 32]);
    let doc = open_with_data(&two_row_paras(huge_papx_grpprl(fc as u32)), &data);
    let ir = doc.to_ir();
    assert!(
        ir.metadata
            .warnings
            .iter()
            .any(|w| w.contains("sprmPHugePapx"))
    );
}

#[test]
fn test_doc_nested_table_header_flag_does_not_mark_a_flattened_row() {
    // A nested table (itap = 2) is flattened into its outer grid, so its
    // header row is just a row in the middle of the outer table: the flag
    // must not carry over.
    let nested_cell = || Para {
        text: "n",
        terminator: '\u{7}',
        // sprmPFInTable = 1, sprmPItap = 2.
        grpprl: vec![0x16, 0x24, 0x01, 0x49, 0x66, 0x02, 0x00, 0x00, 0x00],
    };
    let mut nested_mark = row_grpprl(&CENTERS, &[0, 0]);
    // row_grpprl writes sprmPItap = 1; raise it to 2 (operand byte at index 8).
    assert_eq!(&nested_mark[6..9], &[0x49, 0x66, 0x01]);
    nested_mark[8] = 0x02;
    nested_mark.extend_from_slice(&TABLE_HEADER_ON);
    let paras = vec![
        cell("o1"),
        cell("o2"),
        row_mark(row_grpprl(&CENTERS, &[0, 0])),
        nested_cell(),
        nested_cell(),
        row_mark(nested_mark),
    ];
    let bytes = build_doc(&paras);
    let doc = common::open_doc(&bytes);
    let ir = doc.to_ir();

    let tables = tables(&ir);
    assert_eq!(tables.len(), 1, "the nested rows are flattened into the outer table");
    assert!(
        tables[0].rows.iter().all(|r| !r.is_header),
        "a nested table's header flag must not mark a row of the flattened outer table"
    );
}
