//! Resource bounds on the create path. These tests lower a process-wide
//! limit, so each holds the scope that serialises them and restores it.

use std::io::Cursor;

use office_oxide::DocumentFormat;
use office_oxide::OfficeError;
use office_oxide::core::Error as CoreError;
use office_oxide::create::create_from_ir_to_writer;
use office_oxide::ir::*;
use office_oxide::limits::{
    max_write_sheets, max_write_table_cells, set_max_write_sheets, set_max_write_table_cells,
};

mod common;
use common::LimitScope;

const CELL_LIMIT: usize = 50;

fn table(rows: usize, cols: usize) -> Table {
    let cell = || TableCell {
        content: vec![Element::Paragraph(Paragraph {
            content: vec![InlineContent::Text(TextSpan::plain("c".to_string()))],
            ..Default::default()
        })],
        col_span: 1,
        row_span: 1,
        ..Default::default()
    };
    Table {
        rows: (0..rows)
            .map(|_| TableRow {
                cells: (0..cols).map(|_| cell()).collect(),
                ..Default::default()
            })
            .collect(),
        ..Default::default()
    }
}

fn ir_with(elements: Vec<Element>, format: DocumentFormat) -> DocumentIR {
    DocumentIR {
        metadata: Metadata {
            format,
            ..Default::default()
        },
        sections: vec![Section {
            elements,
            ..Default::default()
        }],
    }
}

fn write(ir: &DocumentIR, format: DocumentFormat) -> Result<(), OfficeError> {
    create_from_ir_to_writer(ir, format, Cursor::new(Vec::new()))
}

#[test]
fn test_a_table_with_too_many_cells_is_refused_when_writing() {
    let mut limits = LimitScope::new();
    limits.set(max_write_table_cells, set_max_write_table_cells, CELL_LIMIT);
    for format in [DocumentFormat::Docx, DocumentFormat::Pptx] {
        let big = ir_with(vec![Element::Table(table(10, 10))], format);
        let err = write(&big, format).expect_err("100 cells exceed a limit of 50");
        assert!(
            matches!(err, OfficeError::Core(CoreError::PackageLimit(_))),
            "{format:?}: {err:?}"
        );
        assert!(err.to_string().contains("set_max_write_table_cells"));

        // Within the limit the document is written as before.
        let small = ir_with(vec![Element::Table(table(5, 5))], format);
        write(&small, format).unwrap_or_else(|e| panic!("{format:?}: {e}"));
    }
}

/// Rows padded to the widest row: 20 one-cell rows and one 20-cell row are
/// only 40 cell entries, but the writers build 21 x 20 cells.
#[test]
fn test_one_wide_row_among_many_narrow_rows_is_charged_at_the_padded_width() {
    let mut limits = LimitScope::new();
    limits.set(max_write_table_cells, set_max_write_table_cells, CELL_LIMIT);
    let mut t = table(20, 1);
    t.rows.extend(table(1, 20).rows);
    for format in [DocumentFormat::Docx, DocumentFormat::Pptx] {
        let ir = ir_with(vec![Element::Table(t.clone())], format);
        let err = write(&ir, format).expect_err("the padded grid is 420 cells");
        assert!(
            matches!(err, OfficeError::Core(CoreError::PackageLimit(_))),
            "{format:?}: {err:?}"
        );
    }
}

/// A column span widens the grid the `.docx` writer lays out (clamped to
/// its column limit), so three cells spanning 100 columns are 300 cells.
/// The `.pptx` writer ignores spans: its table stays three cells.
#[test]
fn test_column_spans_are_charged_as_the_grid_they_widen() {
    let mut limits = LimitScope::new();
    limits.set(max_write_table_cells, set_max_write_table_cells, CELL_LIMIT);
    let mut t = table(3, 1);
    for row in &mut t.rows {
        row.cells[0].col_span = 100;
        row.cells[0].row_span = u32::MAX;
    }
    let docx = ir_with(vec![Element::Table(t.clone())], DocumentFormat::Docx);
    assert!(write(&docx, DocumentFormat::Docx).is_err());
    let pptx = ir_with(vec![Element::Table(t)], DocumentFormat::Pptx);
    write(&pptx, DocumentFormat::Pptx).expect("the pptx table is 3 cells");

    // A span far beyond any grid is clamped, not summed without bound.
    let mut huge = table(1, 2);
    huge.rows[0].cells[0].col_span = u32::MAX;
    huge.rows[0].cells[1].col_span = u32::MAX;
    let ir = ir_with(vec![Element::Table(huge)], DocumentFormat::Docx);
    assert!(write(&ir, DocumentFormat::Docx).is_err());
}

#[test]
fn test_cells_are_counted_across_tables_and_inside_nested_content() {
    let mut limits = LimitScope::new();
    limits.set(max_write_table_cells, set_max_write_table_cells, CELL_LIMIT);
    // Three 20-cell tables: none is over the limit alone, together they are.
    let three =
        ir_with((0..3).map(|_| Element::Table(table(4, 5))).collect(), DocumentFormat::Docx);
    assert!(write(&three, DocumentFormat::Docx).is_err());

    // A table nested in a cell of another counts too.
    let mut outer = table(1, 1);
    outer.rows[0].cells[0].content = vec![Element::Table(table(8, 8))];
    let nested = ir_with(vec![Element::Table(outer)], DocumentFormat::Docx);
    assert!(write(&nested, DocumentFormat::Docx).is_err());
}

#[test]
fn test_a_workbook_with_too_many_sections_is_refused_when_writing() {
    let mut limits = LimitScope::new();
    limits.set(max_write_sheets, set_max_write_sheets, 3);
    let sections = |n: usize| DocumentIR {
        metadata: Metadata {
            format: DocumentFormat::Xlsx,
            ..Default::default()
        },
        sections: (0..n).map(|_| Section::default()).collect(),
    };
    let err =
        write(&sections(4), DocumentFormat::Xlsx).expect_err("4 sections exceed a limit of 3");
    assert!(matches!(err, OfficeError::Core(CoreError::PackageLimit(_))), "{err:?}");
    assert!(err.to_string().contains("set_max_write_sheets"));
    write(&sections(3), DocumentFormat::Xlsx).expect("3 sections are within the limit");
}

#[test]
fn test_a_direct_writer_caller_past_the_sheet_limit_gets_an_error() {
    use office_oxide::xlsx::write::{CellData, XlsxWriter};
    let mut limits = LimitScope::new();
    limits.set(max_write_sheets, set_max_write_sheets, 3);
    let mut w = XlsxWriter::new();
    for i in 0..3 {
        w.add_sheet(&format!("s{i}"));
    }
    w.write_to(Cursor::new(Vec::new()))
        .expect("3 sheets are within the limit");
    // The fourth is not kept; the handle still accepts writes.
    w.add_sheet("extra")
        .set_cell(0, 0, CellData::String("x".into()));
    assert_eq!(w.add_sheet_get_index("more"), 3, "no index is handed out past the limit");
    let err = w
        .write_to(Cursor::new(Vec::new()))
        .expect_err("two sheets were refused");
    assert!(err.to_string().contains("set_max_write_sheets"), "{err}");
}

/// The `.docx` and `.pptx` writers cost a couple of kilobytes a table cell,
/// which is why they are counted; the `.xlsx` writer's cost is a small
/// multiple of the IR it is handed, so a worksheet is bounded by the sheet
/// limit alone and the table-cell limit does not apply to it.
#[test]
fn test_the_table_cell_limit_does_not_apply_to_a_worksheet() {
    let mut limits = LimitScope::new();
    limits.set(max_write_table_cells, set_max_write_table_cells, CELL_LIMIT);
    let sheet = ir_with(vec![Element::Table(table(10, 10))], DocumentFormat::Xlsx);
    write(&sheet, DocumentFormat::Xlsx).expect("100 cells in a worksheet are written");
    let doc = ir_with(vec![Element::Table(table(10, 10))], DocumentFormat::Docx);
    assert!(write(&doc, DocumentFormat::Docx).is_err());
}
