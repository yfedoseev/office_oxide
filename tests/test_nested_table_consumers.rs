//! A table nested in a cell of another table must reach every writer and
//! renderer that has no cell-in-cell, whatever formatting the outer cell has.
//!
//! The IR is built directly: the synthetic `.doc` writer cannot make formatted
//! runs, and these consumers do not care which reader produced the tree.

use office_oxide::ir::{
    DocumentIR, Element, InlineContent, Metadata, Paragraph, Section, Table, TableCell, TableRow,
    TextSpan,
};
use office_oxide::{Document, DocumentFormat};

fn span(text: &str, bold: bool) -> InlineContent {
    let mut s = TextSpan::plain(text);
    s.bold = bold;
    InlineContent::Text(s)
}

fn para(text: &str, bold: bool) -> Element {
    Element::Paragraph(Paragraph {
        content: vec![span(text, bold)],
        ..Default::default()
    })
}

fn cell(content: Vec<Element>) -> TableCell {
    TableCell {
        content,
        col_span: 1,
        row_span: 1,
        ..Default::default()
    }
}

fn row(cells: Vec<TableCell>) -> TableRow {
    TableRow {
        cells,
        ..Default::default()
    }
}

fn table(rows: Vec<TableRow>) -> Element {
    Element::Table(Table {
        rows,
        ..Default::default()
    })
}

/// `Dleft | <outer_text> + nested table | Gright`, the nested rows given as text.
fn ir_with_nested(outer_bold: bool, nested: &[&[&str]]) -> DocumentIR {
    let nested_rows = nested
        .iter()
        .map(|r| row(r.iter().map(|t| cell(vec![para(t, false)])).collect()))
        .collect();
    let middle = cell(vec![para("Emid", outer_bold), table(nested_rows)]);
    let outer = table(vec![row(vec![
        cell(vec![para("Dleft", false)]),
        middle,
        cell(vec![para("Gright", false)]),
    ])]);
    DocumentIR {
        metadata: Metadata {
            format: DocumentFormat::Docx,
            ..Default::default()
        },
        sections: vec![Section {
            elements: vec![outer],
            ..Default::default()
        }],
    }
}

fn written_text(ir: &DocumentIR, format: DocumentFormat, ext: &str) -> String {
    // One directory per call: the tests run in parallel.
    static NEXT: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);
    let n = NEXT.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
    let dir = std::env::temp_dir().join(format!("oo-nested-consumers-{}-{n}", std::process::id()));
    std::fs::create_dir_all(&dir).unwrap();
    let path = dir.join(format!("out.{ext}"));
    office_oxide::create::create_from_ir(ir, format, &path).unwrap();
    let text = Document::open(&path).unwrap().plain_text();
    let _ = std::fs::remove_dir_all(&dir);
    text
}

#[test]
fn test_nested_table_text_survives_xlsx_when_the_outer_cell_is_formatted() {
    // Any bold or coloured run makes the spreadsheet cell a rich string, which
    // is built from the cell's paragraphs only: the nested table's text was
    // dropped entirely.
    let ir = ir_with_nested(true, &[&["Nalpha", "Nbeta"], &["Ngamma", "Ndelta"]]);
    let text = written_text(&ir, DocumentFormat::Xlsx, "xlsx");
    for w in [
        "Dleft", "Emid", "Nalpha", "Nbeta", "Ngamma", "Ndelta", "Gright",
    ] {
        assert!(text.contains(w), "{w} lost: {text:?}");
    }
}

#[test]
fn test_nested_table_text_is_separated_from_the_outer_text_in_every_writer() {
    for (format, ext) in [
        (DocumentFormat::Xlsx, "xlsx"),
        (DocumentFormat::Pptx, "pptx"),
    ] {
        for bold in [false, true] {
            let ir = ir_with_nested(bold, &[&["Nalpha", "Nbeta"]]);
            let text = written_text(&ir, format, ext);
            assert!(
                !text.contains("EmidNalpha"),
                "{ext} bold={bold}: words run together: {text:?}"
            );
            assert!(text.contains("Nalpha"), "{ext} bold={bold}: {text:?}");
        }
    }
}

#[test]
fn test_markdown_keeps_a_nested_row_of_dashes() {
    // `-` is a common "nil" placeholder, and a rendered row `| - | - |` looked
    // like the separator row, so every such row of the nested table vanished.
    let ir = ir_with_nested(false, &[&["Nalpha", "Nbeta"], &["-", "-"], &["Ngamma", "Ndelta"]]);
    let md = ir.to_markdown();
    assert!(md.contains("- \\| -"), "the dash row is missing:\n{md}");
    for w in ["Nalpha", "Nbeta", "Ngamma", "Ndelta"] {
        assert!(md.contains(w), "{w} missing:\n{md}");
    }
    assert_eq!(
        md.lines()
            .filter(|l| l.starts_with('|') && l.contains("---"))
            .count(),
        1,
        "{md}"
    );
}

#[test]
fn test_markdown_keeps_the_columns_of_a_nested_row_with_an_empty_cell() {
    let ir = ir_with_nested(false, &[&["a", "", "c"]]);
    let md = ir.to_markdown();
    assert!(md.contains("a \\|  \\| c"), "the empty cell moved the columns:\n{md}");
}
