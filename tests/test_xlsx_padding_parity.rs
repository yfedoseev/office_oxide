//! The renderers of a spreadsheet must agree about what a sheet within the
//! documented budgets contains: a workbook that `to_markdown` renders in
//! full is not cut short in `to_ir`/`to_html` by a cost model of its own.
//!
//! Budgets are process-global, so this binary holds only tests that want
//! the same text budget and the default padding allowance.

use std::io::{Cursor, Write};

use office_oxide::limits::set_max_text_chars;
use office_oxide::{Document, DocumentFormat};

const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// A workbook package: `sheets` are the `<sheet>` elements, `workbook_rels`
/// the relationships they name, `parts` every other part.
fn package(sheets: &str, workbook_rels: &str, parts: &[(&str, Vec<u8>)]) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts: zip::write::FileOptions<'_, '_, ()> = zip::write::FileOptions::default();
    let fixed: Vec<(&str, Vec<u8>)> = vec![
        (
            "[Content_Types].xml",
            br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>"#.to_vec(),
        ),
        (
            "_rels/.rels",
            format!(r#"<Relationships xmlns="{REL_NS}"><Relationship Id="rId1" Type="{R}/officeDocument" Target="xl/workbook.xml"/></Relationships>"#).into_bytes(),
        ),
        (
            "xl/workbook.xml",
            format!(r#"<workbook xmlns="{NS}" xmlns:r="{R}"><sheets>{sheets}</sheets></workbook>"#).into_bytes(),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            format!(r#"<Relationships xmlns="{REL_NS}">{workbook_rels}</Relationships>"#).into_bytes(),
        ),
    ];
    for (name, data) in fixed.iter().chain(parts.iter()) {
        zip.start_file(*name, opts).unwrap();
        zip.write_all(data).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

fn open(bytes: Vec<u8>) -> office_oxide::Result<Document> {
    Document::from_reader(Cursor::new(bytes), DocumentFormat::Xlsx)
}

fn worksheet_rel(id: &str, target: &str) -> String {
    format!(r#"<Relationship Id="{id}" Type="{R}/worksheet" Target="{target}"/>"#)
}

/// The letters of 1-based column `n` (`1` is `A`, `27` is `AA`).
fn col_name(mut n: usize) -> String {
    let mut out = Vec::new();
    while n > 0 {
        n -= 1;
        out.push(b'A' + (n % 26) as u8);
        n /= 26;
    }
    out.reverse();
    String::from_utf8(out).unwrap()
}

/// `sheets` sheets of `rows` rows, each with a cell in column A and one in
/// column 100 (`CV`): 98 absent columns between them in every row. The last
/// row of the last sheet holds `LASTROW`.
fn wide_export(sheets: usize, rows: usize) -> Vec<u8> {
    let far = col_name(100);
    let mut names = String::new();
    let mut rels = String::new();
    let mut parts = Vec::new();
    for n in 1..=sheets {
        names.push_str(&format!(r#"<sheet name="S{n}" sheetId="{n}" r:id="rId{n}"/>"#));
        rels.push_str(&worksheet_rel(&format!("rId{n}"), &format!("worksheets/sheet{n}.xml")));
        let mut sheet = format!(r#"<worksheet xmlns="{NS}"><sheetData>"#);
        for r in 1..=rows {
            let a = if n == sheets && r == rows {
                "LASTROW"
            } else {
                "x"
            };
            sheet.push_str(&format!(
                r#"<row r="{r}"><c r="A{r}" t="inlineStr"><is><t>{a}</t></is></c><c r="{far}{r}" t="inlineStr"><is><t>y</t></is></c></row>"#
            ));
        }
        sheet.push_str("</sheetData></worksheet>");
        parts.push((format!("xl/worksheets/sheet{n}.xml"), sheet.into_bytes()));
    }
    let parts: Vec<(&str, Vec<u8>)> = parts.iter().map(|(n, d)| (n.as_str(), d.clone())).collect();
    package(&names, &rels, &parts)
}

#[test]
fn test_a_wide_multi_sheet_export_is_the_same_in_markdown_and_the_ir() {
    // 8 sheets x 20 rows x 98 absent columns: 15,680 padding cells, well
    // inside the default allowance. In Markdown each absent column costs 3
    // characters of a 100,000 budget; the IR used to charge 64, and cut
    // the workbook short.
    set_max_text_chars(100_000);
    let doc = open(wide_export(8, 20)).unwrap();
    let md = doc.to_markdown();
    assert!(md.contains("LASTROW") && !md.contains("character budget"), "Markdown was cut");
    let ir = doc.to_ir();
    assert!(!ir.metadata.text_truncated, "the IR is flagged truncated");
    assert!(ir.metadata.warnings.is_empty(), "{:?}", ir.metadata.warnings);
    assert!(ir.plain_text().contains("LASTROW"), "the IR is missing the last row");
    assert!(doc.to_html().contains("LASTROW"), "the HTML is missing the last row");
}

#[test]
fn test_a_single_far_cell_over_many_narrow_rows_stays_bounded_in_the_ir() {
    use office_oxide::ir::Element;
    set_max_text_chars(100_000);
    // 500 rows with a cell in A and one cell in XFD (16,384 columns): 8
    // million empty table cells if every row were laid out in full, against
    // a 100,000 character budget that Markdown spends 3 per column of.
    let far = col_name(16_384);
    let mut sheet = format!(r#"<worksheet xmlns="{NS}"><sheetData>"#);
    for r in 1..=500 {
        sheet.push_str(&format!(r#"<row r="{r}"><c r="A{r}" t="inlineStr"><is><t>x</t></is></c>"#));
        if r == 1 {
            sheet.push_str(&format!(r#"<c r="{far}1" t="inlineStr"><is><t>y</t></is></c>"#));
        }
        sheet.push_str("</row>");
    }
    sheet.push_str("</sheetData></worksheet>");
    let bytes = package(
        r#"<sheet name="S" sheetId="1" r:id="rId1"/>"#,
        &worksheet_rel("rId1", "worksheets/sheet1.xml"),
        &[("xl/worksheets/sheet1.xml", sheet.into_bytes())],
    );
    let ir = open(bytes).unwrap().to_ir();
    let cells: usize = ir
        .sections
        .iter()
        .flat_map(|s| &s.elements)
        .filter_map(|e| match e {
            Element::Table(t) => Some(t.rows.iter().map(|r| r.cells.len()).sum::<usize>()),
            _ => None,
        })
        .sum();
    // Three characters a cell, plus the one row the budget runs out in.
    assert!(cells <= 100_000 / 3 + 2 * 16_384, "{cells} table cells");
    assert!(ir.metadata.text_truncated);
}
