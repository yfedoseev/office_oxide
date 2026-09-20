//! Document text must not inject structure into rendered markdown: a
//! cell containing `|`, a line break, `<tag>`-shaped text or `*stars*`
//! stays one cell with that text. The four format-specific
//! `to_markdown()` table renderers escaped nothing, and the IR renderer
//! left a raw newline in a cell — which ends a GFM table.

mod common;

use std::io::Cursor;

use office_oxide::{Document, DocumentFormat};

const CELLS: [&str; 4] = ["a|b", "line1\nline2", "<Company Name>", "*not bold*"];

/// Every table row of `md` has `cols` cells, and the escaped forms of
/// every hostile cell are present, on both markdown paths.
fn assert_escaped(doc: &Document, cols: usize) {
    for (name, md) in [
        ("to_markdown", doc.to_markdown()),
        ("ir", doc.to_ir().to_markdown()),
    ] {
        let rows: Vec<&str> = md.lines().filter(|l| l.starts_with('|')).collect();
        assert!(rows.len() >= 3, "{name}: no table rendered:\n{md}");
        for row in &rows {
            let n = row.matches(" |").count();
            assert_eq!(n, cols, "{name}: row has {n} cells, not {cols}: {row:?}\n{md}");
        }
        for want in [
            "a\\|b",
            "line1<br>line2",
            "\\<Company Name\\>",
            "\\*not bold\\*",
        ] {
            assert!(md.contains(want), "{name}: {want} missing:\n{md}");
        }
    }
}

#[test]
fn test_xlsx_cells_cannot_break_the_markdown_table() {
    use office_oxide::xlsx::write::{CellData, XlsxWriter};
    let mut wb = XlsxWriter::new();
    let mut s = wb.add_sheet("S");
    s.set_cell(0, 0, CellData::String("Name".into()));
    s.set_cell(0, 1, CellData::String("Value".into()));
    s.set_cell(1, 0, CellData::String(CELLS[0].into()));
    s.set_cell(1, 1, CellData::String(CELLS[1].into()));
    s.set_cell(2, 0, CellData::String(CELLS[2].into()));
    s.set_cell(2, 1, CellData::String(CELLS[3].into()));
    let mut buf = Cursor::new(Vec::new());
    wb.write_to(&mut buf).unwrap();
    let doc = Document::from_reader(Cursor::new(buf.into_inner()), DocumentFormat::Xlsx).unwrap();
    assert_escaped(&doc, 2);
}

#[test]
fn test_docx_cells_cannot_break_the_markdown_table() {
    let mut w = office_oxide::docx::write::DocxWriter::new();
    w.add_table(&[
        vec!["Name", "Value"],
        vec![CELLS[0], CELLS[1]],
        vec![CELLS[2], CELLS[3]],
    ]);
    let mut buf = Cursor::new(Vec::new());
    w.write_to(&mut buf).unwrap();
    let doc = Document::from_reader(Cursor::new(buf.into_inner()), DocumentFormat::Docx).unwrap();
    assert_escaped(&doc, 2);
}

#[test]
fn test_pptx_cells_cannot_break_the_markdown_table() {
    use office_oxide::pptx::write::{PptxWriter, Run};
    let mut w = PptxWriter::new();
    let cell = |t: &str| vec![Run::new(t)];
    w.add_slide().add_table(vec![
        vec![cell("Name"), cell("Value")],
        vec![cell(CELLS[0]), cell(CELLS[1])],
        vec![cell(CELLS[2]), cell(CELLS[3])],
    ]);
    let mut buf = Cursor::new(Vec::new());
    w.write_to(&mut buf).unwrap();
    let doc = Document::from_reader(Cursor::new(buf.into_inner()), DocumentFormat::Pptx).unwrap();
    assert_escaped(&doc, 2);
}

#[test]
fn test_xls_cells_cannot_break_the_markdown_table() {
    use common::{biff, cfb_with_stream};
    let bof = |kind: u16| {
        let mut b = 0x0600u16.to_le_bytes().to_vec();
        b.extend_from_slice(&kind.to_le_bytes());
        b.extend_from_slice(&[0u8; 12]);
        biff(0x0809, &b)
    };
    // BIFF8 LABEL (0x0204): row, col, xf, cch, flags(0 = 8-bit), chars.
    let label = |row: u16, col: u16, text: &str| {
        let mut d = row.to_le_bytes().to_vec();
        d.extend_from_slice(&col.to_le_bytes());
        d.extend_from_slice(&0u16.to_le_bytes());
        d.extend_from_slice(&(text.len() as u16).to_le_bytes());
        d.push(0);
        d.extend_from_slice(text.as_bytes());
        biff(0x0204, &d)
    };
    let mut s = bof(0x0005);
    let mut bs = 0u32.to_le_bytes().to_vec();
    bs.extend_from_slice(&[0, 0, 1, 0, b'S']);
    s.extend(biff(0x0085, &bs));
    s.extend(biff(0x000A, &[]));
    s.extend(bof(0x0010));
    s.extend(label(0, 0, "Name"));
    s.extend(label(0, 1, "Value"));
    s.extend(label(1, 0, CELLS[0]));
    s.extend(label(1, 1, CELLS[1]));
    s.extend(label(2, 0, CELLS[2]));
    s.extend(label(2, 1, CELLS[3]));
    s.extend(biff(0x000A, &[]));
    let doc =
        Document::from_reader(Cursor::new(cfb_with_stream("Workbook", &s)), DocumentFormat::Xls)
            .unwrap();
    assert_escaped(&doc, 2);
}
