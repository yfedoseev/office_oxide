//! Resource bounds on the create path. These tests lower a process-wide
//! limit, so they have a test binary to themselves.

use std::io::Cursor;

use office_oxide::DocumentFormat;
use office_oxide::OfficeError;
use office_oxide::core::Error as CoreError;
use office_oxide::create::create_from_ir_to_writer;
use office_oxide::ir::*;

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
    office_oxide::limits::set_max_write_table_cells(CELL_LIMIT);
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

#[test]
fn test_cells_are_counted_across_tables_and_inside_nested_content() {
    office_oxide::limits::set_max_write_table_cells(CELL_LIMIT);
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
