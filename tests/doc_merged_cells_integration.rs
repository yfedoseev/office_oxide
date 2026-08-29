//! Integration tests for merged table cells in legacy `.doc` files.
//!
//! The document is built **in code** by a minimal synthetic `.doc` writer
//! (`tests/common/mod.rs`) — no third-party fixture blob is committed
//! (AGENTS.md rule #4). It is a 4-row table with horizontally merged cells
//! (`col_span`) and one vertically merged cell (`row_span`):
//!
//! ```text
//! +---------------+---------------+
//! | A (col=3)     | B (col=2)     |
//! +---+---+-------+-------+-------+
//! | C (row=2) | D | E (col=2) | F |
//! |          +---+-------+---+---+
//! |          | G | H (col=2) | I |
//! +---+---+---------------+---+---+
//! | K (col=5)                     |
//! +-------------------------------+
//! ```
//!
//! The column grid is the union of every row's `rgdxaCenter` boundaries
//! (here `[0, 1000, 2000, 3000, 4000, 5000]` in twips). A cell's `col_span`
//! is the number of grid edges inside its own boundary interval, which
//! yields `colspan="3"` (cell A) and `colspan="2"` (cell B). Cell C carries
//! `fVertMerge | fVertRestart` (0x0060); the first cell of row 2 carries only
//! `fVertMerge` (0x0020), so it is absorbed into C, which spans rows 1-2.

mod common;

use common::{Para, build_doc, cell_grpprl, open_doc, row_grpprl};
use office_oxide::ir::{Element, InlineContent, Table, TableCell};

/// Locate the first `Element::Table` in the document's single section.
fn first_table(ir: &office_oxide::ir::DocumentIR) -> Option<&Table> {
    ir.sections
        .iter()
        .flat_map(|s| s.elements.iter())
        .find_map(|e| match e {
            Element::Table(t) => Some(t),
            _ => None,
        })
}

/// Concatenate the text inside a cell; interior paragraphs join with `\n`.
fn cell_text(cell: &TableCell) -> String {
    let mut out = String::new();
    for (i, el) in cell.content.iter().enumerate() {
        if i > 0 {
            out.push('\n');
        }
        if let Element::Paragraph(p) = el {
            for ic in &p.content {
                if let InlineContent::Text(t) = ic {
                    out.push_str(&t.text);
                }
            }
        }
    }
    out
}

/// `(text, col_span, row_span)` for one cell.
fn cell_shape(cell: &TableCell) -> (String, u32, u32) {
    (cell_text(cell), cell.col_span, cell.row_span)
}

/// Build a synthetic merged-cell table `.doc`.
fn synthetic_merged_doc() -> Vec<u8> {
    let mut paras = Vec::new();

    // Row 0: A spans cols 0-3, B spans cols 3-5.
    let r0_centers = [0i16, 3000, 5000];
    paras.push(Para {
        text: "A",
        terminator: '\u{7}',
        grpprl: cell_grpprl(),
    });
    paras.push(Para {
        text: "B",
        terminator: '\u{7}',
        grpprl: cell_grpprl(),
    });
    paras.push(Para {
        text: "",
        terminator: '\u{7}',
        grpprl: row_grpprl(&r0_centers, &[0, 0]),
    });

    // Row 1: C (vert merge start), D, E (cols 2-4), F.
    let r1_centers = [0i16, 1000, 2000, 4000, 5000];
    paras.push(Para {
        text: "C",
        terminator: '\u{7}',
        grpprl: cell_grpprl(),
    });
    paras.push(Para {
        text: "D",
        terminator: '\u{7}',
        grpprl: cell_grpprl(),
    });
    paras.push(Para {
        text: "E",
        terminator: '\u{7}',
        grpprl: cell_grpprl(),
    });
    paras.push(Para {
        text: "F",
        terminator: '\u{7}',
        grpprl: cell_grpprl(),
    });
    paras.push(Para {
        text: "",
        terminator: '\u{7}',
        grpprl: row_grpprl(&r1_centers, &[0x0060, 0, 0, 0]),
    });

    // Row 2: first cell continues C (absorbed), then G, H (cols 2-4), I.
    let r2_centers = [0i16, 1000, 2000, 4000, 5000];
    paras.push(Para {
        text: "",
        terminator: '\u{7}',
        grpprl: cell_grpprl(),
    });
    paras.push(Para {
        text: "G",
        terminator: '\u{7}',
        grpprl: cell_grpprl(),
    });
    paras.push(Para {
        text: "H",
        terminator: '\u{7}',
        grpprl: cell_grpprl(),
    });
    paras.push(Para {
        text: "I",
        terminator: '\u{7}',
        grpprl: cell_grpprl(),
    });
    paras.push(Para {
        text: "",
        terminator: '\u{7}',
        grpprl: row_grpprl(&r2_centers, &[0x0020, 0, 0, 0]),
    });

    // Row 3: K spans the whole five-column grid.
    let r3_centers = [0i16, 5000];
    paras.push(Para {
        text: "K",
        terminator: '\u{7}',
        grpprl: cell_grpprl(),
    });
    paras.push(Para {
        text: "",
        terminator: '\u{7}',
        grpprl: row_grpprl(&r3_centers, &[0]),
    });

    build_doc(&paras)
}

#[test]
fn doc_merged_cells_have_four_rows() {
    let bytes = synthetic_merged_doc();
    let doc = open_doc(&bytes);
    let ir = doc.to_ir();

    let table = first_table(&ir).expect("expected an Element::Table in the IR");
    assert_eq!(table.rows.len(), 4, "the merged table has four rows");
}

#[test]
fn doc_horizontal_merges_set_col_span() {
    let bytes = synthetic_merged_doc();
    let doc = open_doc(&bytes);
    let ir = doc.to_ir();

    let table = first_table(&ir).expect("expected a table");

    // Row 0: A spans the first three grid columns, B the remaining two.
    let shapes: Vec<_> = table.rows[0].cells.iter().map(cell_shape).collect();
    assert_eq!(shapes, vec![("A".to_string(), 3, 1), ("B".to_string(), 2, 1)]);

    // Row 1: E spans the two middle grid columns; C/D/F are single-column.
    let shapes: Vec<_> = table.rows[1].cells.iter().map(cell_shape).collect();
    assert_eq!(
        shapes,
        vec![
            ("C".to_string(), 1, 2),
            ("D".to_string(), 1, 1),
            ("E".to_string(), 2, 1),
            ("F".to_string(), 1, 1),
        ]
    );

    // Row 2: the vertical-merge continuation cell of C is absorbed, so the
    // row has only three cells (G, H, I). H spans two grid columns.
    let shapes: Vec<_> = table.rows[2].cells.iter().map(cell_shape).collect();
    assert_eq!(shapes.len(), 3, "the vertical-merge continuation cell is skipped");
    assert_eq!(shapes[0], ("G".to_string(), 1, 1));
    assert_eq!(shapes[1], ("H".to_string(), 2, 1));
    assert_eq!(shapes[2], ("I".to_string(), 1, 1));

    // Row 3: K spans the whole five-column grid.
    let shapes: Vec<_> = table.rows[3].cells.iter().map(cell_shape).collect();
    assert_eq!(shapes, vec![("K".to_string(), 5, 1)]);
}

#[test]
fn doc_vertical_merge_sets_row_span() {
    let bytes = synthetic_merged_doc();
    let doc = open_doc(&bytes);
    let ir = doc.to_ir();

    let table = first_table(&ir).expect("expected a table");
    let c = &table.rows[1].cells[0];
    assert_eq!(cell_text(c), "C");
    assert_eq!(c.row_span, 2, "C merges with the empty cell in row 2");
    assert_eq!(c.col_span, 1);
}
