//! A workbook is converted, used and released one sheet at a time: what the
//! IR spends on a sheet (its padding cells, its pictures) is that sheet's
//! allowance, given back when the sheet is released, while the text budget
//! caps what the whole run may output.
//!
//! Limits are process-global, so every test holds the scope that serialises
//! them and restores them, and builds workbooks of a few kilobytes.

use std::io::{Cursor, Write};

use office_oxide::limits::*;
use office_oxide::{Document, DocumentFormat};

mod common;
use common::LimitScope;

const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// Sheets in the workbook every test uses, rows in each, and the column
/// (`AD`) of the cell that makes every row 30 columns wide: 28 empty cells of
/// padding in each of 40 rows, 1,120 a sheet, 3,360 in all.
const SHEETS: usize = 3;
const ROWS: usize = 40;
const FAR: usize = 30;
const PADDING_PER_SHEET: usize = ROWS * (FAR - 2);

fn cell_bytes() -> usize {
    std::mem::size_of::<office_oxide::ir::TableCell>()
        + std::mem::size_of::<office_oxide::ir::Element>()
}

/// Picture bytes in the one picture each sheet of the picture workbook draws.
const PICTURE: usize = 6_000;

/// The document-wide allowance holds one sheet's padding and a half, so the
/// eager IR is cut in the second sheet; each sheet has the same allowance to
/// itself, so each fits. A sheet's pictures have room for one picture and a
/// half, and the whole run for two and a third.
fn set_budgets() -> LimitScope {
    let padding = PADDING_PER_SHEET * 3 / 2 * cell_bytes();
    let mut scope = LimitScope::new();
    scope.set(max_text_chars, set_max_text_chars, 200_000);
    scope.set(max_ir_padding_bytes, set_max_ir_padding_bytes, padding);
    scope.set(max_sheet_ir_padding_bytes, set_max_sheet_ir_padding_bytes, padding);
    scope.set(max_image_bytes, set_max_image_bytes, PICTURE * 3 / 2);
    scope.set(max_total_image_bytes, set_max_total_image_bytes, (PICTURE * 7 / 3) as u64);
    scope
}

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

fn package(sheets: &str, rels: &str, parts: &[(String, Vec<u8>)]) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts: zip::write::FileOptions<'_, '_, ()> = zip::write::FileOptions::default();
    let fixed: Vec<(String, Vec<u8>)> = vec![
        (
            "[Content_Types].xml".into(),
            br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>"#.to_vec(),
        ),
        (
            "_rels/.rels".into(),
            format!(r#"<Relationships xmlns="{REL_NS}"><Relationship Id="rId1" Type="{R}/officeDocument" Target="xl/workbook.xml"/></Relationships>"#).into_bytes(),
        ),
        (
            "xl/workbook.xml".into(),
            format!(r#"<workbook xmlns="{NS}" xmlns:r="{R}"><sheets>{sheets}</sheets></workbook>"#).into_bytes(),
        ),
        (
            "xl/_rels/workbook.xml.rels".into(),
            format!(r#"<Relationships xmlns="{REL_NS}">{rels}</Relationships>"#).into_bytes(),
        ),
    ];
    for (name, data) in fixed.iter().chain(parts.iter()) {
        zip.start_file(name.as_str(), opts).unwrap();
        zip.write_all(data).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

/// `sheets` sheets of `rows` rows: a label in column A and a number in column
/// `far`, so every row pads out to `far` columns.
fn workbook(sheets: usize, rows: usize, far: usize) -> Vec<u8> {
    workbook_of_rows(&vec![rows; sheets], far)
}

/// Like [`workbook`], with `rows[i]` rows in sheet `i + 1`.
fn workbook_of_rows(rows: &[usize], far: usize) -> Vec<u8> {
    let far_name = col_name(far);
    let (mut names, mut rels, mut parts) = (String::new(), String::new(), Vec::new());
    for (i, &rows) in (1..).zip(rows) {
        let mut sheet = format!(r#"<worksheet xmlns="{NS}"><sheetData>"#);
        for r in 1..=rows {
            sheet.push_str(&format!(
                r#"<row r="{r}"><c r="A{r}" t="inlineStr"><is><t>s{i}r{r}</t></is></c><c r="{far_name}{r}"><v>{r}</v></c></row>"#
            ));
        }
        sheet.push_str("</sheetData></worksheet>");
        names.push_str(&format!(r#"<sheet name="Sheet{i}" sheetId="{i}" r:id="rId{i}"/>"#));
        rels.push_str(&format!(
            r#"<Relationship Id="rId{i}" Type="{R}/worksheet" Target="worksheets/sheet{i}.xml"/>"#
        ));
        parts.push((format!("xl/worksheets/sheet{i}.xml"), sheet.into_bytes()));
    }
    package(&names, &rels, &parts)
}

fn open(bytes: Vec<u8>) -> Document {
    Document::from_reader(Cursor::new(bytes), DocumentFormat::Xlsx).unwrap()
}

#[test]
fn test_every_renderer_shows_each_sheet_whole_where_the_eager_ir_is_cut() {
    let _limits = set_budgets();
    let doc = open(workbook(SHEETS, ROWS, FAR));
    let last = format!("s{SHEETS}r{ROWS}");

    // The whole-document IR holds everything at once, so it is cut by the
    // document-wide allowance, loudly.
    let ir = doc.to_ir();
    assert!(ir.metadata.text_truncated, "the eager IR is cut without a flag");
    assert!(
        ir.metadata
            .warnings
            .iter()
            .any(|w| w.contains("empty table cells")),
        "no warning: {:?}",
        ir.metadata.warnings
    );

    // The renderers hold one sheet at a time, and each sheet is within its
    // own allowance.
    for (name, out) in [
        ("plain_text", doc.plain_text()),
        ("to_markdown", doc.to_markdown()),
        ("to_markdown_with", doc.to_markdown_with(Default::default())),
        ("to_html", doc.to_html()),
        ("to_html_with", doc.to_html_with(Default::default())),
    ] {
        assert!(out.contains(&last), "{name} lost the last row of the last sheet");
        assert!(!out.contains("truncated"), "{name} carries a truncation notice");
    }
}

/// How many table cells the section's tables hold.
fn table_cells(section: &office_oxide::ir::Section) -> usize {
    section
        .elements
        .iter()
        .map(|e| match e {
            office_oxide::ir::Element::Table(t) => t.rows.iter().map(|r| r.cells.len()).sum(),
            _ => 0,
        })
        .sum()
}

fn table_rows(section: &office_oxide::ir::Section) -> usize {
    section
        .elements
        .iter()
        .map(|e| match e {
            office_oxide::ir::Element::Table(t) => t.rows.len(),
            _ => 0,
        })
        .sum()
}

fn mentions(section: &office_oxide::ir::Section, notice: &str) -> bool {
    serde_json::to_string(section).unwrap().contains(notice)
}

#[test]
fn test_the_sheet_by_sheet_ir_returns_each_sheet_whole() {
    let _limits = set_budgets();
    let doc = open(workbook(SHEETS, ROWS, FAR));
    let xlsx = doc.as_xlsx().unwrap();
    let mut it = xlsx.sheets_ir();
    let mut seen = 0;
    for item in it.by_ref() {
        seen += 1;
        assert_eq!(table_rows(&item.section), ROWS, "sheet {seen} is cut");
        assert_eq!(table_cells(&item.section), ROWS * FAR);
        assert!(!item.text_truncated, "sheet {seen} is flagged");
        assert!(item.warnings.is_empty(), "{:?}", item.warnings);
        assert!(!mentions(&item.section, "truncated"));
    }
    assert_eq!(seen, SHEETS);
    let metadata = it.metadata();
    assert!(!metadata.text_truncated);
    assert!(metadata.warnings.is_empty(), "{:?}", metadata.warnings);
    // The document-wide IR of the same workbook is cut.
    assert!(doc.to_ir().metadata.text_truncated);
    // One sheet on its own, by index.
    let second = xlsx.sheet_ir(1).unwrap();
    assert_eq!(table_rows(&second.section), ROWS);
    assert_eq!(second.section.title.as_deref(), Some("Sheet2"));
    assert!(xlsx.sheet_ir(SHEETS).is_none());
}

#[test]
fn test_a_sheet_past_its_own_allowance_is_cut_and_the_next_sheet_is_whole() {
    let _limits = set_budgets();
    // The middle sheet needs 100 rows of padding, the allowance holds 60.
    let doc = open(workbook_of_rows(&[ROWS, 100, ROWS], FAR));
    let items: Vec<_> = doc.as_xlsx().unwrap().sheets_ir().collect();
    assert_eq!(items.len(), 3);
    assert_eq!(table_rows(&items[0].section), ROWS);
    assert!(!items[0].text_truncated);
    let cut = table_rows(&items[1].section);
    assert!((60..100).contains(&cut), "{cut} rows");
    assert!(items[1].text_truncated, "the cut sheet is not flagged");
    assert!(mentions(&items[1].section, "[sheet truncated"), "no notice on the sheet");
    assert!(
        items[1]
            .warnings
            .iter()
            .any(|w| w.contains("Sheet2") && w.contains("empty table cells")),
        "no warning naming the sheet: {:?}",
        items[1].warnings
    );
    assert_eq!(table_rows(&items[2].section), ROWS, "the next sheet is not whole");
    assert!(!items[2].text_truncated);
    assert!(items[2].warnings.is_empty(), "{:?}", items[2].warnings);
    assert!(!mentions(&items[2].section, "truncated"));
    // The streaming renderers show the notice of the sheet that was cut, and
    // every row of the sheets that were not.
    let html = doc.to_html();
    assert!(html.contains("[sheet truncated"), "HTML names no cause");
    assert!(html.contains(&format!("s3r{ROWS}")));
    assert!(html.contains("set_max_sheet_ir_padding_bytes"));
}

/// Every sheet draws one picture of [`PICTURE`] bytes, all from one media part.
fn picture_sheets(sheets: usize) -> Vec<u8> {
    let drawing = r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="100" cy="100"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="P"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>"#;
    let (mut names, mut rels, mut parts) = (String::new(), String::new(), Vec::new());
    for i in 1..=sheets {
        names.push_str(&format!(r#"<sheet name="Sheet{i}" sheetId="{i}" r:id="rId{i}"/>"#));
        rels.push_str(&format!(
            r#"<Relationship Id="rId{i}" Type="{R}/worksheet" Target="worksheets/sheet{i}.xml"/>"#
        ));
        parts.push((
            format!("xl/worksheets/sheet{i}.xml"),
            format!(r#"<worksheet xmlns="{NS}" xmlns:r="{R}"><sheetData><row r="1"><c r="A1"><v>{i}</v></c></row></sheetData><drawing r:id="rId1"/></worksheet>"#).into_bytes(),
        ));
        parts.push((
            format!("xl/worksheets/_rels/sheet{i}.xml.rels"),
            format!(r#"<Relationships xmlns="{REL_NS}"><Relationship Id="rId1" Type="{R}/drawing" Target="../drawings/drawing{i}.xml"/></Relationships>"#).into_bytes(),
        ));
        parts.push((format!("xl/drawings/drawing{i}.xml"), drawing.as_bytes().to_vec()));
        parts.push((
            format!("xl/drawings/_rels/drawing{i}.xml.rels"),
            format!(r#"<Relationships xmlns="{REL_NS}"><Relationship Id="rId1" Type="{R}/image" Target="../media/image1.png"/></Relationships>"#).into_bytes(),
        ));
    }
    parts.push(("xl/media/image1.png".into(), vec![7u8; PICTURE]));
    package(&names, &rels, &parts)
}

fn images(section: &office_oxide::ir::Section) -> usize {
    fn count(elements: &[office_oxide::ir::Element]) -> usize {
        elements
            .iter()
            .map(|e| match e {
                office_oxide::ir::Element::Image(_) => 1,
                office_oxide::ir::Element::TextBox(t) => count(&t.content),
                _ => 0,
            })
            .sum()
    }
    count(&section.elements)
}

#[test]
fn test_a_sheets_pictures_are_its_own_allowance_and_the_run_has_a_total() {
    let _limits = set_budgets();
    let doc = open(picture_sheets(4));
    // Held at once, four pictures of 6,000 bytes do not fit 9,000.
    let eager = doc.to_ir();
    let eager_images: usize = eager.sections.iter().map(images).sum();
    assert_eq!(eager_images, 1, "the whole-document IR keeps one picture");
    assert!(eager.metadata.text_truncated);
    // One sheet at a time, each picture fits its sheet's allowance, until
    // the total the run may copy (14,000) is spent: two pictures.
    let items: Vec<_> = doc.as_xlsx().unwrap().sheets_ir().collect();
    let counts: Vec<usize> = items.iter().map(|i| images(&i.section)).collect();
    assert_eq!(counts, vec![1, 1, 0, 0], "{counts:?}");
    assert!(!items[1].text_truncated);
    assert!(items[2].text_truncated, "the sheet that spent the total is not flagged");
    assert!(mentions(&items[2].section, "set_max_total_image_bytes"), "no notice");
    assert!(
        items[2]
            .warnings
            .iter()
            .any(|w| w.contains("set_max_total_image_bytes")),
        "{:?}",
        items[2].warnings
    );
    assert!(!mentions(&items[3].section, "set_max_total_image_bytes"), "the notice repeats");
}
