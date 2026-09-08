//! Integration tests for legacy `.doc` table extraction (issue #115).
//!
//! The document is built **in code** by a minimal synthetic `.doc` writer
//! (`tests/common/mod.rs`) — no third-party fixture blob is committed
//! (AGENTS.md rule #4). It is a 3×3 table:
//!
//! ```text
//! 1 | 2 | 4
//! 6 | 9 | 8
//! 7 | 5 | 3
//! ```
//!
//! Before PAPX/TAP parsing landed, `.doc` collapsed this table into a single
//! heading/paragraph string. These tests assert the IR now emits a real
//! `Element::Table` with the right shape.

mod common;

use common::{Para, build_doc, cell_grpprl, open_doc, row_grpprl};
use office_oxide::ir::{Element, TableCell, TableRow};

/// Locate the first `Element::Table` in the document's single section.
fn first_table(ir: &office_oxide::ir::DocumentIR) -> Option<&office_oxide::ir::Table> {
    ir.sections
        .iter()
        .flat_map(|s| s.elements.iter())
        .find_map(|e| match e {
            Element::Table(t) => Some(t),
            _ => None,
        })
}

/// Collect the plain cell text from a table row (one `String` per cell).
fn row_texts(row: &TableRow) -> Vec<String> {
    row.cells.iter().map(|c: &TableCell| cell_text(c)).collect()
}

/// Concatenate all `TextSpan` text inside a cell's block elements.
fn cell_text(cell: &TableCell) -> String {
    use office_oxide::ir::{Heading, InlineContent, Paragraph};

    let mut out = String::new();
    for el in &cell.content {
        match el {
            Element::Paragraph(Paragraph { content, .. })
            | Element::Heading(Heading { content, .. }) => {
                for ic in content {
                    if let InlineContent::Text(t) = ic {
                        out.push_str(&t.text);
                    }
                }
            },
            _ => {},
        }
    }
    out
}

/// Build a 3×3 synthetic table `.doc`.
fn synthetic_table_doc() -> Vec<u8> {
    let cells = [["1", "2", "4"], ["6", "9", "8"], ["7", "5", "3"]];
    let mut paras = Vec::new();
    // Uniform 3-column grid (edges at 0/1000/2000/3000 twips).
    let centers = [0i16, 1000, 2000, 3000];
    for row in &cells {
        for c in row {
            paras.push(Para {
                text: c,
                terminator: '\u{7}',
                grpprl: cell_grpprl(),
            });
        }
        paras.push(Para {
            text: "",
            terminator: '\u{7}',
            grpprl: row_grpprl(&centers, &[0, 0, 0]),
        });
    }
    build_doc(&paras)
}

#[test]
fn doc_table_produces_a_table_element() {
    let bytes = synthetic_table_doc();
    let doc = open_doc(&bytes);
    let ir = doc.to_ir();

    let table = first_table(&ir).expect("expected an Element::Table in the IR");
    assert_eq!(table.rows.len(), 3, "table should have 3 rows");
}

#[test]
fn doc_table_has_three_cells_per_row() {
    let bytes = synthetic_table_doc();
    let doc = open_doc(&bytes);
    let ir = doc.to_ir();

    let table = first_table(&ir).expect("expected a table");
    for (i, row) in table.rows.iter().enumerate() {
        assert_eq!(row.cells.len(), 3, "row {i} should have 3 cells, got {}", row.cells.len());
    }
}

#[test]
fn doc_table_cell_values_match_fixture() {
    let bytes = synthetic_table_doc();
    let doc = open_doc(&bytes);
    let ir = doc.to_ir();

    let table = first_table(&ir).expect("expected a table");
    let rows: Vec<Vec<String>> = table.rows.iter().map(row_texts).collect();
    assert_eq!(rows.len(), 3);
    assert_eq!(rows[0], vec!["1", "2", "4"]);
    assert_eq!(rows[1], vec!["6", "9", "8"]);
    assert_eq!(rows[2], vec!["7", "5", "3"]);
}
