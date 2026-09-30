//! `Document::to_markdown()` (each format's direct renderer) and
//! `to_ir().to_markdown()` (the IR renderer) are two pipelines behind the
//! same promise. They agreed on words but not on layout: DOCX blocks ran
//! together with no blank line (so a table after a list became part of
//! the last list item), PPTX bullets were marked only when indented and
//! numbering vanished, and an XLSX sheet was a table on one path and
//! paragraphs on the other. These tests build representative documents in
//! code and require the two outputs to be identical.

use std::io::Cursor;

use office_oxide::xlsx::write::{CellData, XlsxWriter};
use office_oxide::{Document, DocumentFormat};

const RICH_MD: &str = "\
# Title

Intro with **bold**, *italic* and a [link](https://example.com/a).

## Second level

- one
  - one-a
- two

1. first
2. second

| Col A | Col B |
|---|---|
| a1 | b1 |

Tail paragraph.
";

fn from_markdown(format: DocumentFormat) -> Document {
    let mut buf = Cursor::new(Vec::new());
    office_oxide::create::create_from_markdown_to_writer(RICH_MD, format, &mut buf).unwrap();
    Document::from_reader(Cursor::new(buf.into_inner()), format).unwrap()
}

fn assert_same_markdown(name: &str, doc: &Document) {
    let direct = doc.to_markdown();
    let via_ir = doc.to_ir().to_markdown();
    assert_eq!(
        direct.trim_end(),
        via_ir.trim_end(),
        "{name}: direct and IR markdown differ\n--- direct ---\n{direct}\n--- via IR ---\n{via_ir}"
    );
}

#[test]
fn test_docx_direct_and_ir_markdown_are_identical() {
    let doc = from_markdown(DocumentFormat::Docx);
    assert_same_markdown("docx", &doc);
    let md = doc.to_markdown();
    // Blocks are separated by a blank line; items of one list are not.
    assert!(md.contains("- one\n  - one-a\n- two\n\n1. first\n2. second\n\n| Col A"), "{md}");
}

#[test]
fn test_pptx_direct_and_ir_markdown_are_identical() {
    let doc = from_markdown(DocumentFormat::Pptx);
    assert_same_markdown("pptx", &doc);
    let md = doc.to_markdown();
    // Plain paragraphs are not bullets; a level-0 bulleted list is one; a
    // numbered run keeps its numbers.
    assert!(
        md.contains("\n\nSecond level\n\n- one\n  - one-a\n- two\n\n1. first\n2. second"),
        "{md}"
    );
}

fn xlsx(build: impl FnOnce(&mut XlsxWriter)) -> Document {
    let mut w = XlsxWriter::new();
    build(&mut w);
    let mut buf = Cursor::new(Vec::new());
    w.write_to(&mut buf).unwrap();
    Document::from_reader(Cursor::new(buf.into_inner()), DocumentFormat::Xlsx).unwrap()
}

fn s(text: &str) -> CellData {
    CellData::String(text.to_string())
}

#[test]
fn test_xlsx_direct_and_ir_markdown_are_identical() {
    // A grid is a table on both paths.
    let grid = xlsx(|w| {
        let mut sh = w.add_sheet("Grid");
        for (r, row) in [["Name", "Qty"], ["apple", "3"], ["pear", "5"]]
            .iter()
            .enumerate()
        {
            for (c, v) in row.iter().enumerate() {
                sh.set_cell(r, c, s(v));
            }
        }
    });
    assert_same_markdown("xlsx grid", &grid);
    assert!(grid.to_markdown().contains("| Name | Qty |"), "{}", grid.to_markdown());

    // A single column of short values is data: still a table on both.
    let column = xlsx(|w| {
        let mut sh = w.add_sheet("Codes");
        for (r, v) in ["A1", "B2", "C3", "D4"].iter().enumerate() {
            sh.set_cell(r, 0, s(v));
        }
    });
    assert_same_markdown("xlsx column", &column);
    assert!(column.to_markdown().contains("| A1 |"), "{}", column.to_markdown());

    // A sheet of prose rows (one row with two cells among them) is
    // paragraphs on both, the two-cell row tab-separated.
    let prose = xlsx(|w| {
        let mut sh = w.add_sheet("Notes");
        sh.set_cell(0, 0, s("The quarter closed ahead of plan."));
        sh.set_cell(1, 0, s("Costs were flat against the prior year."));
        sh.set_cell(2, 0, s("Left"));
        sh.set_cell(2, 3, s("Right"));
        sh.set_cell(3, 0, s("Hiring resumes next quarter."));
        sh.set_cell(4, 0, s("No further changes are expected."));
    });
    assert_same_markdown("xlsx prose", &prose);
    let md = prose.to_markdown();
    assert!(md.contains("plan.\n\nCosts") && md.contains("Left\tRight"), "{md}");
}

/// The markdown-created workbook carries cell formatting the direct
/// renderer does not draw (bold header cells, a hyperlink), so only its
/// layout is compared: the same blocks in the same order.
#[test]
fn test_xlsx_from_markdown_has_the_same_layout_on_both_paths() {
    let doc = from_markdown(DocumentFormat::Xlsx);
    let strip = |md: &str| {
        let mut out = md.replace("**", "");
        // `[text](target)` → `text`.
        while let Some(open) = out.find("](") {
            let start = out[..open].rfind('[').unwrap_or(open);
            let end = out[open..].find(')').map_or(out.len(), |e| open + e + 1);
            let text = out[start + 1..open].to_string();
            out.replace_range(start..end, &text);
        }
        out.trim_end().to_string()
    };
    let direct = strip(&doc.to_markdown());
    let via_ir = strip(&doc.to_ir().to_markdown());
    assert_eq!(direct, via_ir);
}
