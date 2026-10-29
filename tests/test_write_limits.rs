//! Resource bounds on the create path. These tests lower a process-wide
//! limit, so each holds the scope that serialises them and restores it.

use std::io::Cursor;

use office_oxide::DocumentFormat;
use office_oxide::OfficeError;
use office_oxide::core::Error as CoreError;
use office_oxide::create::create_from_ir_to_writer;
use office_oxide::ir::*;
use office_oxide::limits::{
    DEFAULT_MAX_WRITE_BYTES, WRITE_BYTES_PER_DOCX_CELL, WRITE_BYTES_PER_XLSX_CELL, max_write_bytes,
    max_write_sheets, set_max_write_bytes, set_max_write_sheets,
};

mod common;
use common::LimitScope;

/// The write-memory limit these tests lower to: what 50 `.docx` cells cost.
const BYTE_LIMIT: u64 = 50 * WRITE_BYTES_PER_DOCX_CELL;

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
    limits.set(max_write_bytes, set_max_write_bytes, BYTE_LIMIT);
    for format in [DocumentFormat::Docx, DocumentFormat::Pptx] {
        // A `.pptx` cell weighs about a third of a `.docx` cell to the writer,
        // so the same limit holds three times as many of them.
        let rows = if format == DocumentFormat::Docx {
            10
        } else {
            30
        };
        let big = ir_with(vec![Element::Table(table(rows, 10))], format);
        let err = write(&big, format).expect_err("past a limit of 50 cells");
        assert!(
            matches!(err, OfficeError::Core(CoreError::PackageLimit(_))),
            "{format:?}: {err:?}"
        );
        assert!(err.to_string().contains("set_max_write_bytes"));

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
    limits.set(max_write_bytes, set_max_write_bytes, BYTE_LIMIT);
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
    limits.set(max_write_bytes, set_max_write_bytes, BYTE_LIMIT);
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
    limits.set(max_write_bytes, set_max_write_bytes, BYTE_LIMIT);
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

/// What a cell costs the writer differs by format, so the one limit is in
/// bytes and a cell is charged what its format takes: 100 cells are past a
/// limit of 50 `.docx` cells in a `.docx`, and within it in a `.pptx` (about a
/// third the cost) and in a worksheet (about a tenth).
#[test]
fn test_a_cell_is_charged_what_its_formats_writer_costs() {
    let mut limits = LimitScope::new();
    limits.set(max_write_bytes, set_max_write_bytes, BYTE_LIMIT);
    let hundred = |format| ir_with(vec![Element::Table(table(10, 10))], format);
    assert!(write(&hundred(DocumentFormat::Docx), DocumentFormat::Docx).is_err());
    write(&hundred(DocumentFormat::Pptx), DocumentFormat::Pptx).expect("a pptx cell costs less");
    write(&hundred(DocumentFormat::Xlsx), DocumentFormat::Xlsx)
        .expect("a worksheet cell costs less");
}

#[test]
fn test_a_worksheet_with_too_many_cells_is_refused_when_writing() {
    let mut limits = LimitScope::new();
    limits.set(max_write_bytes, set_max_write_bytes, 50 * WRITE_BYTES_PER_XLSX_CELL);
    assert_eq!(max_write_bytes(), 50 * WRITE_BYTES_PER_XLSX_CELL);
    let big = ir_with(vec![Element::Table(table(10, 10))], DocumentFormat::Xlsx);
    let err = write(&big, DocumentFormat::Xlsx).expect_err("100 cells exceed a limit of 50");
    assert!(matches!(err, OfficeError::Core(CoreError::PackageLimit(_))), "{err:?}");
    assert!(err.to_string().contains("set_max_write_bytes"), "{err}");
    // Cells are counted across the workbook's sheets, padded to the widest row.
    let two = DocumentIR {
        metadata: Metadata {
            format: DocumentFormat::Xlsx,
            ..Default::default()
        },
        sections: (0..2)
            .map(|_| Section {
                elements: vec![Element::Table(table(4, 7))],
                ..Default::default()
            })
            .collect(),
    };
    assert!(write(&two, DocumentFormat::Xlsx).is_err(), "56 cells exceed a limit of 50");
    // Within the limit the workbook is written as before.
    let small = ir_with(vec![Element::Table(table(5, 5))], DocumentFormat::Xlsx);
    write(&small, DocumentFormat::Xlsx).expect("25 cells are within a limit of 50");
}

#[test]
fn test_the_default_write_limit_leaves_ordinary_documents_alone() {
    let _defaults = common::default_limits();
    assert_eq!(max_write_bytes(), DEFAULT_MAX_WRITE_BYTES);
    for format in [
        DocumentFormat::Docx,
        DocumentFormat::Pptx,
        DocumentFormat::Xlsx,
    ] {
        let ir = ir_with(vec![Element::Table(table(1_000, 20))], format);
        write(&ir, format).unwrap_or_else(|e| panic!("{format:?}: 20,000 cells: {e}"));
    }
}

/// The writers used directly are held to the same limit as the IR path: the
/// cells they were handed are counted when the document is written.
#[test]
fn test_a_writer_used_directly_is_held_to_the_write_limit() {
    use office_oxide::docx::write::DocxWriter;
    use office_oxide::pptx::write::{PptxWriter, Run};
    use office_oxide::xlsx::write::{CellData, XlsxWriter};
    let mut limits = LimitScope::new();
    limits.set(max_write_bytes, set_max_write_bytes, BYTE_LIMIT);
    let refused = |err: office_oxide::OfficeError| {
        assert!(err.to_string().contains("set_max_write_bytes"), "{err}");
    };

    // 10 x 10 = 100 cells: past 50 `.docx` cells.
    let row: Vec<&str> = vec!["c"; 10];
    let rows: Vec<Vec<&str>> = vec![row; 10];
    let mut docx = DocxWriter::new();
    docx.add_table(&rows);
    refused(
        docx.write_to(Cursor::new(Vec::new()))
            .expect_err("100 docx cells")
            .into(),
    );
    let mut small = DocxWriter::new();
    small.add_table(&[vec!["a", "b"], vec!["c", "d"]]);
    small
        .write_to(Cursor::new(Vec::new()))
        .expect("4 cells are within the limit");

    // A `.pptx` cell costs a third as much: 300 are past the limit, 100 are not.
    let cell = || vec![Run::new("c")];
    let pptx_rows = |n: usize| vec![(0..10).map(|_| cell()).collect::<Vec<_>>(); n];
    let mut big = PptxWriter::new();
    big.add_slide().add_table(pptx_rows(30));
    refused(
        big.write_to(Cursor::new(Vec::new()))
            .expect_err("300 pptx cells")
            .into(),
    );
    let mut ok = PptxWriter::new();
    ok.add_slide().add_table(pptx_rows(10));
    ok.write_to(Cursor::new(Vec::new()))
        .expect("100 pptx cells fit");

    // A worksheet cell costs a tenth: 600 cells are past the limit.
    let mut book = XlsxWriter::new();
    {
        let mut sheet = book.add_sheet("S");
        for r in 0..30 {
            sheet.add_row(
                (0..20)
                    .map(|c| CellData::String(format!("{r}{c}")))
                    .collect(),
            );
        }
    }
    refused(
        book.write_to(Cursor::new(Vec::new()))
            .expect_err("600 cells")
            .into(),
    );
}

/// `save` is held to the limit as `write_to` is, and tables in a header, a
/// footnote or a slide's notes are counted with the body's.
#[test]
fn test_save_and_the_other_containers_are_held_to_the_write_limit_too() {
    use office_oxide::docx::write::DocxWriter;
    use office_oxide::pptx::write::{PptxWriter, Run};
    let mut limits = LimitScope::new();
    limits.set(max_write_bytes, set_max_write_bytes, BYTE_LIMIT);
    let dir = std::env::temp_dir().join(format!("office_oxide_save_limit_{}", std::process::id()));
    std::fs::create_dir_all(&dir).unwrap();
    let docx_path = dir.join("big.docx");
    let pptx_path = dir.join("big.pptx");

    let row: Vec<&str> = vec!["c"; 10];
    let rows: Vec<Vec<&str>> = vec![row; 10];
    let mut docx = DocxWriter::new();
    docx.add_table(&rows);
    assert!(docx.save(&docx_path).is_err(), "save must refuse what write_to refuses");

    // The same table in the IR, placed in a footnote.
    let mut with_note = DocxWriter::new();
    with_note.add_footnote(1, &[Element::Table(table(10, 10))], None);
    assert!(
        with_note.write_to(Cursor::new(Vec::new())).is_err(),
        "a table in a footnote is counted"
    );

    let cell = || vec![Run::new("c")];
    let mut pptx = PptxWriter::new();
    pptx.add_slide()
        .add_table(vec![(0..10).map(|_| cell()).collect::<Vec<_>>(); 30]);
    assert!(pptx.save(&pptx_path).is_err(), "save must refuse what write_to refuses");
    std::fs::remove_dir_all(&dir).ok();
}
