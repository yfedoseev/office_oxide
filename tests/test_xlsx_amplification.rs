//! A spreadsheet can name one shared thing from many places: a formula
//! master shared by many cells, a picture drawn from many anchors, a
//! comment author on many comments. The thing is stored once; each place
//! gets a copy. These tests build a small workbook that fans out one
//! shared item, lower the document budgets so the product is small, and
//! assert the product is bounded (an `Err` or a documented truncation)
//! rather than materialised.
//!
//! The budgets are process-global: every test that lowers them holds the scope
//! that serialises them and restores them, and the tests that rely on the
//! defaults hold the matching guard. Documents are a few kilobytes.

use std::io::{Cursor, Write};

use office_oxide::limits::{
    max_memory_bytes, max_text_chars, set_max_memory_bytes, set_max_text_chars,
};
use office_oxide::{Document, DocumentFormat};

mod common;
use common::LimitScope;

/// The text budget every test here runs under.
const TEXT_BUDGET: usize = 60_000;

fn set_budgets() -> LimitScope {
    let mut scope = LimitScope::new();
    scope.set(max_text_chars, set_max_text_chars, TEXT_BUDGET);
    // About a thousand empty cells: five rows of the 200-column sheets below.
    let allowance = 1_000
        * (std::mem::size_of::<office_oxide::ir::TableCell>()
            + std::mem::size_of::<office_oxide::ir::Element>());
    // The document-wide padding allowance is half the memory budget (and a
    // sheet's own, which `to_html` converts against, all of it).
    scope.set(max_memory_bytes, set_max_memory_bytes, 2 * allowance as u64);
    scope
}

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

/// One sheet whose column B holds a `master_len`-character shared-formula
/// master and `followers` bare followers in the rows below it.
fn shared_formula_workbook(master_len: usize, followers: usize) -> Vec<u8> {
    let master = format!("A1+{}", "1+".repeat(master_len / 2));
    let mut sheet = format!(r#"<worksheet xmlns="{NS}"><sheetData>"#);
    sheet.push_str(&format!(
        r#"<row r="1"><c r="B1"><f t="shared" ref="B1:B{}" si="0">{master}</f><v>1</v></c></row>"#,
        followers + 1
    ));
    for r in 2..=followers + 1 {
        sheet.push_str(&format!(
            r#"<row r="{r}"><c r="B{r}"><f t="shared" si="0"/><v>1</v></c></row>"#
        ));
    }
    sheet.push_str("</sheetData></worksheet>");
    package(
        r#"<sheet name="S" sheetId="1" r:id="rId1"/>"#,
        &worksheet_rel("rId1", "worksheets/sheet1.xml"),
        &[("xl/worksheets/sheet1.xml", sheet.into_bytes())],
    )
}

/// The sheet's cells in column B that still have a formula, and how many
/// have a value.
fn formula_and_value_counts(ws: &office_oxide::xlsx::Worksheet) -> (usize, usize) {
    let cells = ws.rows.iter().flat_map(|r| &r.cells);
    (
        cells.clone().filter(|c| c.formula.is_some()).count(),
        cells
            .filter(|c| !matches!(c.value, office_oxide::xlsx::CellValue::Empty))
            .count(),
    )
}

#[test]
fn test_shared_formula_master_copied_into_many_followers_is_bounded() {
    let _limits = set_budgets();
    // 4,000 characters (under the 8,192 limit) copied into 100 followers is
    // 400,000 characters against a 60,000 budget: the followers that fit keep
    // their formula, the others keep their cell and value without one.
    let doc = open(shared_formula_workbook(4_000, 100)).expect("the sheet stays readable");
    let ws = &doc.as_xlsx().unwrap().worksheets[0];
    let (formulas, values) = formula_and_value_counts(ws);
    assert_eq!(values, 101, "every cell keeps its value");
    assert!(formulas > 1 && formulas < 20, "{formulas} cells kept a formula");
    assert_eq!(ws.formulas_dropped, 101 - formulas);
    let ir = doc.to_ir();
    assert!(ir.metadata.text_truncated, "the loss is not flagged");
    assert!(
        ir.metadata
            .warnings
            .iter()
            .any(|w| w.contains("shared formula"))
    );
}

#[test]
fn test_shared_formula_master_over_excel_limit_is_dropped_but_the_cell_is_kept() {
    let _limits = set_budgets();
    let doc = open(shared_formula_workbook(9_000, 1)).expect("the sheet stays readable");
    let ws = &doc.as_xlsx().unwrap().worksheets[0];
    assert_eq!(formula_and_value_counts(ws), (0, 2));
    assert_eq!(ws.formulas_dropped, 2, "the master and its follower");
    assert!(
        doc.to_ir()
            .metadata
            .warnings
            .iter()
            .any(|w| w.contains("8192"))
    );
}

#[test]
fn test_shared_formula_within_budget_still_translates() {
    let _limits = set_budgets();
    let doc = open(shared_formula_workbook(20, 5)).expect("a modest shared formula opens");
    assert!(doc.plain_text().contains('1'));
    let json = serde_json::to_string(&doc.to_ir()).unwrap();
    assert!(json.contains("A6+"), "the last follower's formula is translated");
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

/// `rows` rows with one cell each in column A, and one cell in column
/// `width` of the first row: every row is laid out `width` columns wide.
fn sparse_wide_workbook(rows: usize, width: usize) -> Vec<u8> {
    let far = col_name(width);
    let mut sheet = format!(r#"<worksheet xmlns="{NS}"><sheetData>"#);
    for r in 1..=rows {
        sheet.push_str(&format!(r#"<row r="{r}"><c r="A{r}" t="inlineStr"><is><t>x</t></is></c>"#));
        if r == 1 {
            sheet.push_str(&format!(r#"<c r="{far}1" t="inlineStr"><is><t>y</t></is></c>"#));
        }
        sheet.push_str("</row>");
    }
    sheet.push_str("</sheetData></worksheet>");
    package(
        r#"<sheet name="S" sheetId="1" r:id="rId1"/>"#,
        &worksheet_rel("rId1", "worksheets/sheet1.xml"),
        &[("xl/worksheets/sheet1.xml", sheet.into_bytes())],
    )
}

#[test]
fn test_column_padding_of_sparse_rows_is_charged_to_the_text_budget() {
    let _limits = set_budgets();
    // 2,000 rows, each laid out 200 columns wide: 3 characters of padding
    // per column in Markdown is 1.2 MB against a 60,000 budget.
    let doc = open(sparse_wide_workbook(2_000, 200)).unwrap();
    let md = doc.to_markdown();
    assert!(md.contains("character budget"), "Markdown carries no truncation notice");
    assert!(md.len() < TEXT_BUDGET + 2_000, "{} bytes of Markdown", md.len());
}

#[test]
fn test_ir_table_rows_padded_to_the_sheet_width_are_bounded() {
    let _limits = set_budgets();
    let doc = open(sparse_wide_workbook(2_000, 200)).unwrap();
    let ir = doc.to_ir();
    let cells: usize = ir
        .sections
        .iter()
        .flat_map(|s| &s.elements)
        .filter_map(|e| match e {
            office_oxide::ir::Element::Table(t) => {
                Some(t.rows.iter().map(|r| r.cells.len()).sum::<usize>())
            },
            _ => None,
        })
        .sum();
    assert!(cells < 2_000, "{cells} table cells were built from a 3 KB workbook");
    assert!(ir.metadata.text_truncated, "the truncation is not flagged");
    let html = doc.to_html();
    assert!(html.contains("empty table cells"), "HTML names no cause");
    assert!(!html.contains("character budget"), "the cause is not the text budget");
    assert!(
        ir.metadata
            .warnings
            .iter()
            .any(|w| w.contains("empty table cells")),
        "no warning: {:?}",
        ir.metadata.warnings
    );
}

#[test]
fn test_a_dense_wide_sheet_is_unaffected_by_padding_charges() {
    let _limits = set_budgets();
    // Every column of every row is filled, so nothing is padding.
    let width = 150;
    let mut sheet = format!(r#"<worksheet xmlns="{NS}"><sheetData>"#);
    for r in 1..=5 {
        sheet.push_str(&format!(r#"<row r="{r}">"#));
        for c in 1..=width {
            sheet.push_str(&format!(r#"<c r="{}{r}"><v>{c}</v></c>"#, col_name(c)));
        }
        sheet.push_str("</row>");
    }
    sheet.push_str("</sheetData></worksheet>");
    let bytes = package(
        r#"<sheet name="S" sheetId="1" r:id="rId1"/>"#,
        &worksheet_rel("rId1", "worksheets/sheet1.xml"),
        &[("xl/worksheets/sheet1.xml", sheet.into_bytes())],
    );
    let doc = open(bytes).unwrap();
    assert!(!doc.to_markdown().contains("character budget"));
    assert!(!doc.to_ir().metadata.text_truncated);
}

#[test]
fn test_shared_formula_budget_is_per_document_not_per_sheet() {
    let _limits = set_budgets();
    // Each of six sheets translates about 20,000 characters, under the
    // 60,000 budget on its own; together they are twice the budget.
    let master = format!("A1+{}", "1+".repeat(500));
    let mut sheet = format!(r#"<worksheet xmlns="{NS}"><sheetData>"#);
    sheet.push_str(&format!(
        r#"<row r="1"><c r="B1"><f t="shared" ref="B1:B21" si="0">{master}</f><v>1</v></c></row>"#
    ));
    for r in 2..=21 {
        sheet.push_str(&format!(
            r#"<row r="{r}"><c r="B{r}"><f t="shared" si="0"/><v>1</v></c></row>"#
        ));
    }
    sheet.push_str("</sheetData></worksheet>");
    let mut sheets = String::new();
    let mut rels = String::new();
    let mut parts = Vec::new();
    for i in 1..=6 {
        sheets.push_str(&format!(r#"<sheet name="S{i}" sheetId="{i}" r:id="rId{i}"/>"#));
        rels.push_str(&worksheet_rel(&format!("rId{i}"), &format!("worksheets/sheet{i}.xml")));
        parts.push((format!("xl/worksheets/sheet{i}.xml"), sheet.clone().into_bytes()));
    }
    let parts: Vec<(&str, Vec<u8>)> = parts.iter().map(|(n, d)| (n.as_str(), d.clone())).collect();
    // Whichever order the sheets are parsed in, the first sheets keep their
    // formulas and the last lose them.
    let bytes = package(&sheets, &rels, &parts);
    let dropped = |bytes: &[u8]| -> Vec<usize> {
        let doc = open(bytes.to_vec()).expect("no sheet is unreadable");
        doc.as_xlsx()
            .unwrap()
            .worksheets
            .iter()
            .map(|w| w.formulas_dropped)
            .collect()
    };
    let first = dropped(&bytes);
    assert_eq!(first.len(), 6, "every sheet is read");
    assert_eq!(&first[..2], &[0, 0], "{first:?}");
    assert_eq!(first[5], 20, "{first:?}");
    for _ in 0..5 {
        assert_eq!(dropped(&bytes), first, "the outcome depends on scheduling");
    }
}

#[test]
fn test_padding_is_charged_after_the_text_has_spent_its_budget() {
    let _limits = set_budgets();
    // The last row's text spends the whole text budget; the rows before it
    // are still laid out 200 columns wide and must still be paid for.
    let far = col_name(200);
    let mut sheet = format!(r#"<worksheet xmlns="{NS}"><sheetData>"#);
    for r in 1..=2_000 {
        sheet.push_str(&format!(r#"<row r="{r}"><c r="A{r}" t="inlineStr"><is><t>x</t></is></c>"#));
        if r == 1 {
            sheet.push_str(&format!(r#"<c r="{far}1" t="inlineStr"><is><t>y</t></is></c>"#));
        }
        if r == 2_000 {
            sheet.push_str(&format!(
                r#"<c r="B{r}" t="inlineStr"><is><t>{}</t></is></c>"#,
                "z".repeat(TEXT_BUDGET + 1)
            ));
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
            office_oxide::ir::Element::Table(t) => {
                Some(t.rows.iter().map(|r| r.cells.len()).sum::<usize>())
            },
            _ => None,
        })
        .sum();
    assert!(cells < 2_000, "{cells} table cells were built after the text budget was spent");
}

/// The picture budget every test here runs under.
const IMAGE_BUDGET: usize = 10_000;

/// A sheet whose one drawing anchors the same `media_len`-byte picture
/// `anchors` times.
fn picture_fan_out_workbook(anchors: usize, media_len: usize) -> Vec<u8> {
    picture_workbook(anchors, media_len, 0)
}

/// [`picture_fan_out_workbook`] followed by `plain_sheets` sheets with one
/// number each and no drawing.
fn picture_workbook(anchors: usize, media_len: usize, plain_sheets: usize) -> Vec<u8> {
    let mut drawing = String::from(
        r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
    );
    for i in 0..anchors {
        drawing.push_str(&format!(
            r#"<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{i}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="100" cy="100"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{}" name="P{i}"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>"#,
            i + 2
        ));
    }
    drawing.push_str("</xdr:wsDr>");
    let sheet = format!(
        r#"<worksheet xmlns="{NS}" xmlns:r="{R}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><drawing r:id="rId1"/></worksheet>"#
    );
    let mut sheets = String::from(r#"<sheet name="S" sheetId="1" r:id="rId1"/>"#);
    let mut rels = worksheet_rel("rId1", "worksheets/sheet1.xml");
    let mut parts = Vec::new();
    for i in 0..plain_sheets {
        let n = i + 2;
        sheets.push_str(&format!(r#"<sheet name="P{n}" sheetId="{n}" r:id="rId{n}"/>"#));
        rels.push_str(&worksheet_rel(&format!("rId{n}"), &format!("worksheets/sheet{n}.xml")));
        parts.push((
            format!("xl/worksheets/sheet{n}.xml"),
            format!(r#"<worksheet xmlns="{NS}"><sheetData><row r="1"><c r="A1"><v>{n}</v></c></row></sheetData></worksheet>"#).into_bytes(),
        ));
    }
    let mut parts: Vec<(&str, Vec<u8>)> =
        parts.iter().map(|(n, d)| (n.as_str(), d.clone())).collect();
    parts.extend([
            ("xl/worksheets/sheet1.xml", sheet.into_bytes()),
            (
                "xl/worksheets/_rels/sheet1.xml.rels",
                format!(r#"<Relationships xmlns="{REL_NS}"><Relationship Id="rId1" Type="{R}/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#).into_bytes(),
            ),
            ("xl/drawings/drawing1.xml", drawing.into_bytes()),
            (
                "xl/drawings/_rels/drawing1.xml.rels",
                format!(r#"<Relationships xmlns="{REL_NS}"><Relationship Id="rId1" Type="{R}/image" Target="../media/image1.png"/></Relationships>"#).into_bytes(),
            ),
            ("xl/media/image1.png", vec![7u8; media_len]),
    ]);
    package(&sheets, &rels, &parts)
}

fn ir_image_count(ir: &office_oxide::DocumentIR) -> usize {
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
    ir.sections.iter().map(|s| count(&s.elements)).sum()
}

#[test]
fn test_one_picture_drawn_from_many_anchors_shares_its_bytes_and_is_bounded_in_the_ir() {
    let mut limits = set_budgets();
    // A picture's allowance is an eighth of the memory budget.
    limits.set(max_memory_bytes, set_max_memory_bytes, 8 * IMAGE_BUDGET as u64);
    // 100 anchors of a 1,000-byte picture: 100,000 bytes of IR images
    // against a 10,000-byte budget.
    let doc = open(picture_fan_out_workbook(100, 1_000)).unwrap();
    let ws = &doc.as_xlsx().unwrap().worksheets[0];
    assert_eq!(ws.images.len(), 100);
    assert!(
        std::sync::Arc::ptr_eq(&ws.images[0].data, &ws.images[99].data),
        "the anchors each hold their own copy of the picture"
    );
    let ir = doc.to_ir();
    assert_eq!(ir_image_count(&ir), 10, "pictures past the budget are left out");
    assert!(ir.metadata.text_truncated);
    assert!(ir.plain_text().contains("picture budget"));
}

#[test]
fn test_the_picture_notice_is_given_once_not_on_every_later_sheet() {
    let mut limits = set_budgets();
    // A picture's allowance is an eighth of the memory budget.
    limits.set(max_memory_bytes, set_max_memory_bytes, 8 * IMAGE_BUDGET as u64);
    // The first sheet spends the picture budget; the two sheets after it have
    // no pictures and nothing to be told about.
    let doc = open(picture_workbook(100, 1_000, 2)).unwrap();
    let ir = doc.to_ir();
    let notices = ir.plain_text().matches("picture budget").count();
    assert_eq!(notices, 1, "the notice appears {notices} times");
    let warnings = ir
        .metadata
        .warnings
        .iter()
        .filter(|w| w.contains("pictures were left out"))
        .count();
    assert_eq!(warnings, 1, "{:?}", ir.metadata.warnings);
    assert!(doc.plain_text().contains("P3"), "the later sheets are kept");
}

#[test]
fn test_many_distinct_small_pictures_within_the_budget_are_all_kept() {
    let mut limits = set_budgets();
    // A picture's allowance is an eighth of the memory budget.
    limits.set(max_memory_bytes, set_max_memory_bytes, 8 * IMAGE_BUDGET as u64);
    let doc = open(picture_fan_out_workbook(5, 1_000)).unwrap();
    let ir = doc.to_ir();
    assert_eq!(ir_image_count(&ir), 5);
    assert!(!ir.metadata.text_truncated);
}

/// A sheet with `comments` comments by one author whose name is
/// `author_len` characters.
fn comment_fan_out_workbook(comments: usize, author_len: usize) -> Vec<u8> {
    let author = "a".repeat(author_len);
    let mut part = format!(
        r#"<comments xmlns="{NS}"><authors><author>{author}</author></authors><commentList>"#
    );
    for i in 1..=comments {
        part.push_str(&format!(
            r#"<comment ref="A{i}" authorId="0"><text><r><t>note</t></r></text></comment>"#
        ));
    }
    part.push_str("</commentList></comments>");
    let sheet = format!(
        r#"<worksheet xmlns="{NS}" xmlns:r="{R}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#
    );
    package(
        r#"<sheet name="S" sheetId="1" r:id="rId1"/>"#,
        &worksheet_rel("rId1", "worksheets/sheet1.xml"),
        &[
            ("xl/worksheets/sheet1.xml", sheet.into_bytes()),
            (
                "xl/worksheets/_rels/sheet1.xml.rels",
                format!(r#"<Relationships xmlns="{REL_NS}"><Relationship Id="rId1" Type="{R}/comments" Target="../comments1.xml"/></Relationships>"#).into_bytes(),
            ),
            ("xl/comments1.xml", part.into_bytes()),
        ],
    )
}

#[test]
fn test_one_comment_author_named_by_many_comments_is_charged_to_the_text_budget() {
    let _limits = set_budgets();
    // 100 comments by an author whose name is 2,000 characters: 200,000
    // characters of rendered names against a 60,000 budget.
    let doc = open(comment_fan_out_workbook(100, 2_000)).unwrap();
    let ws = &doc.as_xlsx().unwrap().worksheets[0];
    assert!(
        std::sync::Arc::ptr_eq(
            ws.comments[0].author.as_ref().unwrap(),
            ws.comments[99].author.as_ref().unwrap()
        ),
        "every comment holds its own copy of the author's name"
    );
    for (surface, out) in [
        ("plain_text", doc.plain_text()),
        ("to_markdown", doc.to_markdown()),
    ] {
        assert!(out.contains("character budget"), "{surface}: no truncation notice");
        assert!(out.len() < 2 * TEXT_BUDGET, "{surface}: {} bytes", out.len());
    }
    let ir = doc.to_ir();
    let notes: usize = ir
        .sections
        .iter()
        .flat_map(|s| &s.elements)
        .filter(|e| matches!(e, office_oxide::ir::Element::Endnote(_)))
        .count();
    assert!(notes < 100, "{notes} comment notes were built");
    assert!(ir.metadata.text_truncated);
}

#[test]
fn test_a_few_comments_by_one_author_are_unaffected() {
    let _limits = set_budgets();
    let doc = open(comment_fan_out_workbook(3, 20)).unwrap();
    assert!(!doc.plain_text().contains("character budget"));
    assert!(!doc.to_ir().metadata.text_truncated);
}

/// Two sheets, each with a cell and `threads` threaded comments (one
/// message each) by one person whose name is `author_len` characters.
fn threaded_comment_workbook(threads: usize, author_len: usize) -> Vec<u8> {
    let persons = format!(
        r#"<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"><person displayName="{}" id="{{P}}" userId="u" providerId="AD"/></personList>"#,
        "a".repeat(author_len)
    );
    let mut sheets = String::new();
    let mut rels = String::new();
    let mut parts: Vec<(String, Vec<u8>)> =
        vec![("xl/persons/person.xml".to_string(), persons.into_bytes())];
    for i in 1..=2 {
        sheets.push_str(&format!(r#"<sheet name="S{i}" sheetId="{i}" r:id="rId{i}"/>"#));
        rels.push_str(&worksheet_rel(&format!("rId{i}"), &format!("worksheets/sheet{i}.xml")));
        parts.push((
            format!("xl/worksheets/sheet{i}.xml"),
            format!(
                r#"<worksheet xmlns="{NS}"><sheetData><row r="1"><c r="A1"><v>4{i}</v></c></row></sheetData></worksheet>"#
            )
            .into_bytes(),
        ));
        parts.push((
            format!("xl/worksheets/_rels/sheet{i}.xml.rels"),
            format!(r#"<Relationships xmlns="{REL_NS}"><Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/t{i}.xml"/></Relationships>"#).into_bytes(),
        ));
        let mut tc = String::from(
            r#"<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">"#,
        );
        for t in 1..=threads {
            tc.push_str(&format!(
                r#"<threadedComment ref="A{t}" id="{{T{t}}}" personId="{{P}}"><text>note</text></threadedComment>"#
            ));
        }
        tc.push_str("</ThreadedComments>");
        parts.push((format!("xl/threadedComments/t{i}.xml"), tc.into_bytes()));
    }
    let parts: Vec<(&str, Vec<u8>)> = parts.iter().map(|(n, d)| (n.as_str(), d.clone())).collect();
    package(&sheets, &rels, &parts)
}

#[test]
fn test_threaded_comment_flood_keeps_the_cells_and_reports_the_truncation() {
    let _limits = set_budgets();
    // 20 threads of 2,000 characters on each of two sheets: 40,000 per
    // sheet, under the 60,000 budget alone, 80,000 together.
    let doc = open(threaded_comment_workbook(20, 2_000)).unwrap();
    let xlsx = doc.as_xlsx().unwrap();
    assert!(xlsx.unreadable_sheets.is_empty(), "no sheet is dropped for its comments");
    assert_eq!(xlsx.worksheets.len(), 2);
    assert!(!xlsx.worksheets[0].comments_truncated);
    assert!(
        xlsx.worksheets[1].comments_truncated,
        "the budget is the document's, not each sheet's"
    );
    assert!(xlsx.worksheets[1].comments.len() < 20);
    for ws in &xlsx.worksheets {
        assert_eq!(ws.rows.len(), 1, "{}: the cells are kept", ws.name);
    }
    for out in [doc.plain_text(), doc.to_markdown()] {
        assert!(out.contains("truncated"), "the truncation is not reported");
    }
    let ir = doc.to_ir();
    assert!(ir.metadata.text_truncated);
}

/// One sheet with `pivots` relationships to the same pivot table part, whose
/// cache (external source) holds `records` records of `value_len` characters.
fn pivot_fan_out_workbook(
    pivots: usize,
    records: usize,
    value_len: usize,
    fields: usize,
    field_len: usize,
) -> Vec<u8> {
    let mut sheet_rels = format!(r#"<Relationships xmlns="{REL_NS}">"#);
    for i in 1..=pivots {
        sheet_rels.push_str(&format!(
            r#"<Relationship Id="rId{i}" Type="{R}/pivotTable" Target="../pivotTables/pivotTable1.xml"/>"#
        ));
    }
    sheet_rels.push_str("</Relationships>");
    let mut cache_records = format!(r#"<pivotCacheRecords xmlns="{NS}">"#);
    for _ in 0..records {
        cache_records.push_str(&format!(r#"<r><s v="{}"/></r>"#, "p".repeat(value_len)));
    }
    cache_records.push_str("</pivotCacheRecords>");
    let mut cache_fields = format!(r#"<cacheFields count="{fields}">"#);
    for f in 0..fields {
        cache_fields.push_str(&format!(
            r#"<cacheField name="{}{f}"><sharedItems/></cacheField>"#,
            "f".repeat(field_len)
        ));
    }
    cache_fields.push_str("</cacheFields>");
    package(
        r#"<sheet name="S" sheetId="1" r:id="rId1"/>"#,
        &worksheet_rel("rId1", "worksheets/sheet1.xml"),
        &[
            (
                "xl/worksheets/sheet1.xml",
                format!(r#"<worksheet xmlns="{NS}"><sheetData><row r="1"><c r="A1"><v>42</v></c></row></sheetData></worksheet>"#).into_bytes(),
            ),
            ("xl/worksheets/_rels/sheet1.xml.rels", sheet_rels.into_bytes()),
            (
                "xl/pivotTables/pivotTable1.xml",
                format!(r#"<pivotTableDefinition xmlns="{NS}" name="P"><location ref="A1:B2"/></pivotTableDefinition>"#).into_bytes(),
            ),
            (
                "xl/pivotTables/_rels/pivotTable1.xml.rels",
                format!(r#"<Relationships xmlns="{REL_NS}"><Relationship Id="rId1" Type="{R}/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition1.xml"/></Relationships>"#).into_bytes(),
            ),
            (
                "xl/pivotCache/pivotCacheDefinition1.xml",
                format!(r#"<pivotCacheDefinition xmlns="{NS}" xmlns:r="{R}" r:id="rId1"><cacheSource type="external"/>{cache_fields}</pivotCacheDefinition>"#).into_bytes(),
            ),
            (
                "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels",
                format!(r#"<Relationships xmlns="{REL_NS}"><Relationship Id="rId1" Type="{R}/pivotCacheRecords" Target="pivotCacheRecords1.xml"/></Relationships>"#).into_bytes(),
            ),
            ("xl/pivotCache/pivotCacheRecords1.xml", cache_records.into_bytes()),
        ],
    )
}

#[test]
fn test_one_pivot_cache_named_by_many_pivots_is_shared_and_its_output_charged() {
    let _limits = set_budgets();
    // 40 pivots over a cache of 50 records of 100 characters: 5,000
    // characters stored once, 200,000 rendered against a 60,000 budget.
    let doc = open(pivot_fan_out_workbook(40, 50, 100, 1, 1)).unwrap();
    let pivots = &doc.as_xlsx().unwrap().worksheets[0].pivot_tables;
    assert_eq!(pivots.len(), 40);
    assert!(
        std::sync::Arc::ptr_eq(
            pivots[0].cache_data.as_ref().unwrap(),
            pivots[39].cache_data.as_ref().unwrap()
        ),
        "every pivot holds its own copy of the cache"
    );
    for (surface, out) in [
        ("plain_text", doc.plain_text()),
        ("to_markdown", doc.to_markdown()),
    ] {
        assert!(out.contains("character budget"), "{surface}: no truncation notice");
        assert!(out.len() < 2 * TEXT_BUDGET, "{surface}: {} bytes", out.len());
    }
    assert!(doc.to_ir().metadata.text_truncated);
}

#[test]
fn test_a_few_pivots_over_one_cache_are_unaffected() {
    let _limits = set_budgets();
    let doc = open(pivot_fan_out_workbook(2, 5, 20, 1, 1)).unwrap();
    assert!(!doc.to_markdown().contains("character budget"));
    assert!(!doc.to_ir().metadata.text_truncated);
}

#[test]
fn test_sheets_naming_one_part_are_read_once_and_the_rest_reported() {
    let _limits = set_budgets();
    let mut sheets = String::new();
    for i in 1..=30 {
        sheets.push_str(&format!(r#"<sheet name="S{i}" sheetId="{i}" r:id="rId1"/>"#));
    }
    let sheet = format!(
        r#"<worksheet xmlns="{NS}"><sheetData><row r="1"><c r="A1"><v>42</v></c></row></sheetData></worksheet>"#
    );
    let doc = open(package(
        &sheets,
        &worksheet_rel("rId1", "worksheets/sheet1.xml"),
        &[("xl/worksheets/sheet1.xml", sheet.into_bytes())],
    ))
    .unwrap();
    let xlsx = doc.as_xlsx().unwrap();
    assert_eq!(xlsx.worksheets.len(), 1, "the part is parsed once");
    assert_eq!(xlsx.unreadable_sheets.len(), 29, "the other entries are reported");
    assert!(doc.plain_text().contains("already the content of sheet"));
}

#[test]
fn test_pivot_headers_are_charged_when_a_cache_has_no_records() {
    let _limits = set_budgets();
    // 100 pivots over a cache of 100 fields of 100 characters and no
    // records: the field names are stored once (10,000 characters) and
    // written once per pivot, which is 1,000,000 against a 60,000 budget.
    let doc = open(pivot_fan_out_workbook(100, 0, 0, 100, 100)).unwrap();
    for (surface, out) in [
        ("plain_text", doc.plain_text()),
        ("to_markdown", doc.to_markdown()),
    ] {
        assert!(out.contains("character budget"), "{surface}: no truncation notice");
        assert!(out.len() < 3 * TEXT_BUDGET, "{surface}: {} bytes", out.len());
    }
    let ir = doc.to_ir();
    let cells: usize = ir
        .sections
        .iter()
        .flat_map(|s| &s.elements)
        .filter_map(|e| match e {
            office_oxide::ir::Element::Table(t) => {
                Some(t.rows.iter().map(|r| r.cells.len()).sum::<usize>())
            },
            _ => None,
        })
        .sum();
    assert!(cells < 100 * 100, "{cells} header cells were built");
    assert!(ir.metadata.text_truncated);
}

#[test]
fn test_one_sheet_part_under_several_spellings_is_still_read_once() {
    let _limits = set_budgets();
    // The zip layer finds one entry under all of these.
    let spellings = [
        "worksheets/sheet2.xml",
        "Worksheets\\Sheet2.xml",
        "worksheets/../worksheets/sheet2.xml",
        "/xl/worksheets/sheet2.xml",
        "./worksheets/SHEET2.XML",
        // A part the package lacks falls back to the entry its position
        // names, which is the same part.
        "worksheets/missing.xml",
    ];
    let mut sheets = String::new();
    let mut rels = String::new();
    for (i, target) in spellings.iter().enumerate() {
        let n = i + 1;
        sheets.push_str(&format!(r#"<sheet name="S{n}" sheetId="{n}" r:id="rId{n}"/>"#));
        rels.push_str(&worksheet_rel(&format!("rId{n}"), target));
    }
    let sheet = format!(
        r#"<worksheet xmlns="{NS}"><sheetData><row r="1"><c r="A1"><v>42</v></c></row></sheetData></worksheet>"#
    );
    let doc =
        open(package(&sheets, &rels, &[("xl/worksheets/sheet2.xml", sheet.into_bytes())])).unwrap();
    let xlsx = doc.as_xlsx().unwrap();
    assert_eq!(xlsx.worksheets.len(), 1, "the part is parsed once");
    assert_eq!(xlsx.unreadable_sheets.len(), spellings.len() - 1);
}

/// One sheet of `links` numeric cells (A1, A2, ...), each hyperlinked
/// through one relationship whose external target is `target_len`
/// characters long.
fn hyperlink_fan_out_workbook(links: usize, target_len: usize) -> Vec<u8> {
    let mut sheet = format!(r#"<worksheet xmlns="{NS}" xmlns:r="{R}"><sheetData>"#);
    for r in 1..=links {
        sheet.push_str(&format!(r#"<row r="{r}"><c r="A{r}"><v>{r}</v></c></row>"#));
    }
    sheet.push_str("</sheetData><hyperlinks>");
    for r in 1..=links {
        sheet.push_str(&format!(r#"<hyperlink ref="A{r}" r:id="rId1"/>"#));
    }
    sheet.push_str("</hyperlinks></worksheet>");
    let target = format!("http://example.com/{}", "a".repeat(target_len));
    package(
        r#"<sheet name="S" sheetId="1" r:id="rId1"/>"#,
        &worksheet_rel("rId1", "worksheets/sheet1.xml"),
        &[
            ("xl/worksheets/sheet1.xml", sheet.into_bytes()),
            (
                "xl/worksheets/_rels/sheet1.xml.rels",
                format!(r#"<Relationships xmlns="{REL_NS}"><Relationship Id="rId1" Type="{R}/hyperlink" Target="{target}" TargetMode="External"/></Relationships>"#).into_bytes(),
            ),
        ],
    )
}

/// Bytes of hyperlink target held by the IR's text spans, and how many spans
/// carry one.
fn ir_hyperlinks(ir: &office_oxide::DocumentIR) -> (usize, usize) {
    use office_oxide::ir::{Element, InlineContent};
    let (mut bytes, mut count) = (0, 0);
    let mut paragraph = |content: &[InlineContent]| {
        for c in content {
            if let InlineContent::Text(span) = c {
                if let Some(h) = &span.hyperlink {
                    bytes += h.len();
                    count += 1;
                }
            }
        }
    };
    for e in ir.sections.iter().flat_map(|s| &s.elements) {
        match e {
            Element::Paragraph(p) => paragraph(&p.content),
            Element::Table(t) => {
                for cell in t.rows.iter().flat_map(|r| &r.cells) {
                    for el in &cell.content {
                        if let Element::Paragraph(p) = el {
                            paragraph(&p.content);
                        }
                    }
                }
            },
            _ => {},
        }
    }
    (bytes, count)
}

#[test]
fn test_a_hyperlink_target_past_the_package_limit_is_reported_not_dropped_silently() {
    let _limits = set_budgets();
    // A relationship target of 100,000 characters is past the 64 KiB every
    // format accepts: the relationships part is refused, and the sheet's
    // links with it. That is on record, not an empty set of relationships.
    let doc = open(hyperlink_fan_out_workbook(40, 100_000)).unwrap();
    let ir = doc.to_ir();
    assert_eq!(ir_hyperlinks(&ir), (0, 0), "an over-long target was copied into the IR");
    let text = doc.plain_text();
    assert!(text.contains("40"), "the cells stay: {text}");
    assert!(ir.metadata.text_truncated, "the loss is not flagged");
    assert!(
        ir.metadata
            .warnings
            .iter()
            .any(|w| w.contains("sheet1.xml.rels")),
        "no warning: {:?}",
        ir.metadata.warnings
    );
}

#[test]
fn test_a_hyperlink_target_within_the_package_limit_is_kept_and_charged() {
    let _budgets = set_budgets();
    // 40 links to a 30,000-character target (inside 64 KiB, as docx and
    // pptx accept) is 1.2 million characters against a 60,000 budget: the
    // links that fit are kept, the rest are cut, loudly.
    let doc = open(hyperlink_fan_out_workbook(40, 30_000)).unwrap();
    let ir = doc.to_ir();
    let (bytes, count) = ir_hyperlinks(&ir);
    assert!(count >= 1, "a legitimate long target is dropped");
    assert!(bytes <= 120_000, "{bytes} bytes of target copied into the IR");
    assert!(ir.metadata.text_truncated);
}

#[test]
fn test_one_hyperlink_target_on_many_cells_is_charged_to_the_text_budget() {
    let _limits = set_budgets();
    // 2,000 characters on each of 100 cells is 200,000 against a 60,000
    // budget.
    let doc = open(hyperlink_fan_out_workbook(100, 2_000)).unwrap();
    let ir = doc.to_ir();
    let (bytes, count) = ir_hyperlinks(&ir);
    assert!(count > 0, "the links that fit are kept");
    assert!(bytes <= TEXT_BUDGET, "{bytes} bytes of hyperlink in the IR");
    assert!(ir.metadata.text_truncated, "the truncation is not flagged");
    assert!(
        ir.metadata
            .warnings
            .iter()
            .any(|w| w.contains("character budget"))
    );
}

#[test]
fn test_a_few_ordinary_hyperlinks_are_unaffected() {
    let _limits = set_budgets();
    let doc = open(hyperlink_fan_out_workbook(3, 40)).unwrap();
    let ir = doc.to_ir();
    assert_eq!(ir_hyperlinks(&ir).1, 3);
    assert!(!ir.metadata.text_truncated);
    assert!(ir.metadata.warnings.is_empty(), "{:?}", ir.metadata.warnings);
}

/// One shared string of `runs` one-character rich-text runs (each with its
/// own font), named by `cells` cells in column A.
fn rich_text_fan_out_workbook(runs: usize, cells: usize) -> Vec<u8> {
    let mut sst = format!(r#"<sst xmlns="{NS}" count="{cells}" uniqueCount="1"><si>"#);
    for _ in 0..runs {
        sst.push_str(r#"<r><rPr><b/><rFont val="Arial"/><sz val="11"/></rPr><t>x</t></r>"#);
    }
    sst.push_str("</si></sst>");
    let mut sheet = format!(r#"<worksheet xmlns="{NS}"><sheetData>"#);
    for r in 1..=cells {
        sheet.push_str(&format!(r#"<row r="{r}"><c r="A{r}" t="s"><v>0</v></c></row>"#));
    }
    sheet.push_str("</sheetData></worksheet>");
    package(
        r#"<sheet name="S" sheetId="1" r:id="rId1"/>"#,
        &format!(
            r#"{}<Relationship Id="rId9" Type="{R}/sharedStrings" Target="sharedStrings.xml"/>"#,
            worksheet_rel("rId1", "worksheets/sheet1.xml")
        ),
        &[
            ("xl/worksheets/sheet1.xml", sheet.into_bytes()),
            ("xl/sharedStrings.xml", sst.into_bytes()),
        ],
    )
}

/// How many text spans the IR holds in table cells and paragraphs.
fn ir_span_count(ir: &office_oxide::DocumentIR) -> usize {
    use office_oxide::ir::{Element, InlineContent};
    let count = |content: &[InlineContent]| {
        content
            .iter()
            .filter(|c| matches!(c, InlineContent::Text(_)))
            .count()
    };
    let mut total = 0;
    for e in ir.sections.iter().flat_map(|s| &s.elements) {
        match e {
            Element::Paragraph(p) => total += count(&p.content),
            Element::Table(t) => {
                for el in t
                    .rows
                    .iter()
                    .flat_map(|r| &r.cells)
                    .flat_map(|c| &c.content)
                {
                    if let Element::Paragraph(p) = el {
                        total += count(&p.content);
                    }
                }
            },
            _ => {},
        }
    }
    total
}

#[test]
fn test_rich_text_runs_of_one_shared_string_are_charged_for_every_cell_that_names_it() {
    let _limits = set_budgets();
    // 100 runs of one character each in 50 cells: 5,000 characters of text,
    // but 5,000 spans of a few hundred bytes each against a 60,000 budget.
    let doc = open(rich_text_fan_out_workbook(100, 50)).unwrap();
    let ir = doc.to_ir();
    let spans = ir_span_count(&ir);
    assert!(spans > 0, "the cells that fit are kept");
    assert!(spans < 1_000, "{spans} spans were built from a 3 KB workbook");
    assert!(ir.metadata.text_truncated, "the truncation is not flagged");
    assert!(
        ir.metadata
            .warnings
            .iter()
            .any(|w| w.contains("character budget"))
    );
}

#[test]
fn test_a_few_rich_text_cells_are_unaffected() {
    let _limits = set_budgets();
    let doc = open(rich_text_fan_out_workbook(3, 5)).unwrap();
    let ir = doc.to_ir();
    assert_eq!(ir_span_count(&ir), 15);
    assert!(!ir.metadata.text_truncated);
}

/// `cells` numeric cells in column A, all in a style whose number format code
/// is `code`.
fn number_format_workbook(code: &str, cells: usize) -> Vec<u8> {
    let styles = format!(
        r#"<styleSheet xmlns="{NS}"><numFmts count="1"><numFmt numFmtId="164" formatCode="{code}"/></numFmts><fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="164" applyNumberFormat="1"/></cellXfs></styleSheet>"#
    );
    let mut sheet = format!(r#"<worksheet xmlns="{NS}"><sheetData>"#);
    for r in 1..=cells {
        sheet.push_str(&format!(r#"<row r="{r}"><c r="A{r}" s="1"><v>{r}</v></c></row>"#));
    }
    sheet.push_str("</sheetData></worksheet>");
    package(
        r#"<sheet name="S" sheetId="1" r:id="rId1"/>"#,
        &worksheet_rel("rId1", "worksheets/sheet1.xml"),
        &[
            ("xl/worksheets/sheet1.xml", sheet.into_bytes()),
            ("xl/styles.xml", styles.into_bytes()),
        ],
    )
}

/// The number-format code of every table cell in the IR that has one.
fn ir_number_formats(ir: &office_oxide::DocumentIR) -> Vec<String> {
    use office_oxide::ir::Element;
    ir.sections
        .iter()
        .flat_map(|s| &s.elements)
        .filter_map(|e| match e {
            Element::Table(t) => Some(t),
            _ => None,
        })
        .flat_map(|t| t.rows.iter().flat_map(|r| &r.cells))
        .filter_map(|c| c.number_format.clone())
        .collect()
}

#[test]
fn test_a_number_format_code_longer_than_excels_limit_is_not_copied_into_every_cell() {
    let _limits = set_budgets();
    // Excel accepts at most 255 characters of format code; a 100,000
    // character one was copied into every cell that used it.
    let code = format!("0.{}", "0".repeat(100_000));
    let doc = open(number_format_workbook(&code, 40)).unwrap();
    let ir = doc.to_ir();
    let copied: usize = ir_number_formats(&ir).iter().map(String::len).sum();
    assert_eq!(copied, 0, "{copied} bytes of format code were copied into the IR");
    assert!(
        ir.metadata
            .warnings
            .iter()
            .any(|w| w.contains("number format")),
        "no warning: {:?}",
        ir.metadata.warnings
    );
    assert!(doc.plain_text().contains("40"), "the cells keep their values");
}

#[test]
fn test_an_ordinary_number_format_code_is_still_reported_on_each_cell() {
    let _limits = set_budgets();
    let doc = open(number_format_workbook("0.00", 5)).unwrap();
    let ir = doc.to_ir();
    assert_eq!(ir_number_formats(&ir), vec!["0.00"; 5]);
    assert!(ir.metadata.warnings.is_empty(), "{:?}", ir.metadata.warnings);
}

#[test]
fn test_rich_text_runs_longer_than_the_cell_text_shown_are_charged_in_full() {
    let _limits = set_budgets();
    // The cell's displayed text stops at 32 KB, but its spans are built from
    // the runs: two 40,000-character runs, named by 20 cells.
    let mut sst = format!(r#"<sst xmlns="{NS}" count="20" uniqueCount="1"><si>"#);
    for _ in 0..2 {
        sst.push_str(&format!(r#"<r><rPr><b/></rPr><t>{}</t></r>"#, "y".repeat(40_000)));
    }
    sst.push_str("</si></sst>");
    let mut sheet = format!(r#"<worksheet xmlns="{NS}"><sheetData>"#);
    for r in 1..=20 {
        sheet.push_str(&format!(r#"<row r="{r}"><c r="A{r}" t="s"><v>0</v></c></row>"#));
    }
    sheet.push_str("</sheetData></worksheet>");
    let bytes = package(
        r#"<sheet name="S" sheetId="1" r:id="rId1"/>"#,
        &format!(
            r#"{}<Relationship Id="rId9" Type="{R}/sharedStrings" Target="sharedStrings.xml"/>"#,
            worksheet_rel("rId1", "worksheets/sheet1.xml")
        ),
        &[
            ("xl/worksheets/sheet1.xml", sheet.into_bytes()),
            ("xl/sharedStrings.xml", sst.into_bytes()),
        ],
    );
    let ir = open(bytes).unwrap().to_ir();
    let held: usize = serde_json::to_string(&ir).unwrap().matches('y').count();
    assert!(held <= TEXT_BUDGET, "{held} characters of run text in the IR");
    assert!(ir.metadata.text_truncated);
}

#[test]
fn test_hyperlinks_over_one_relationship_share_one_target_allocation() {
    use office_oxide::xlsx::HyperlinkTarget;
    let _limits = set_budgets();
    let doc = open(hyperlink_fan_out_workbook(100, 2_000)).unwrap();
    let links = &doc.as_xlsx().unwrap().worksheets[0].hyperlinks;
    assert_eq!(links.len(), 100);
    let targets: Vec<_> = links
        .iter()
        .map(|l| match &l.target {
            HyperlinkTarget::External(t) => t.clone(),
            other => panic!("{other:?}"),
        })
        .collect();
    assert!(
        targets
            .iter()
            .all(|t| std::sync::Arc::ptr_eq(t, &targets[0]))
    );
}

/// Cells merged across columns and down rows, with a merge covering a cell
/// that holds a value.
fn merged_workbook() -> Vec<u8> {
    let sheet = format!(
        r#"<worksheet xmlns="{NS}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>head</t></is></c><c r="D1" t="inlineStr"><is><t>d</t></is></c></row><row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>2</v></c></row><row r="3"><c r="A3"><v>3</v></c></row><row r="4"><c r="C4"><v>9</v></c></row></sheetData><mergeCells count="2"><mergeCell ref="A1:C1"/><mergeCell ref="A2:A3"/></mergeCells></worksheet>"#
    );
    package(
        r#"<sheet name="M" sheetId="1" r:id="rId1"/>"#,
        &worksheet_rel("rId1", "worksheets/sheet1.xml"),
        &[("xl/worksheets/sheet1.xml", sheet.into_bytes())],
    )
}

/// A sheet of flowing sentences (laid out as paragraphs) and an empty sheet.
fn prose_workbook() -> Vec<u8> {
    let mut sheet = format!(r#"<worksheet xmlns="{NS}"><sheetData>"#);
    for r in 1..=12 {
        sheet.push_str(&format!(
            r#"<row r="{r}"><c r="A{r}" t="inlineStr"><is><t>This is a flowing sentence number {r} of some prose in a document style sheet.</t></is></c></row>"#
        ));
    }
    sheet.push_str("</sheetData></worksheet>");
    package(
        r#"<sheet name="Prose" sheetId="1" r:id="rId1"/><sheet name="Empty" sheetId="2" r:id="rId2"/>"#,
        &format!(
            "{}{}",
            worksheet_rel("rId1", "worksheets/sheet1.xml"),
            worksheet_rel("rId2", "worksheets/sheet2.xml")
        ),
        &[
            ("xl/worksheets/sheet1.xml", sheet.into_bytes()),
            (
                "xl/worksheets/sheet2.xml",
                format!(r#"<worksheet xmlns="{NS}"><sheetData/></worksheet>"#).into_bytes(),
            ),
        ],
    )
}

/// `sheets` sheets of `rows` rows, a label in column A and a number in
/// column `far`.
fn multi_sheet_workbook(sheets: usize, rows: usize, far: usize) -> Vec<u8> {
    let far_name = col_name(far);
    let (mut names, mut rels, mut parts) = (String::new(), String::new(), Vec::new());
    for i in 1..=sheets {
        let mut sheet = format!(r#"<worksheet xmlns="{NS}"><sheetData>"#);
        for r in 1..=rows {
            sheet.push_str(&format!(
                r#"<row r="{r}"><c r="A{r}" t="inlineStr"><is><t>s{i}r{r}</t></is></c><c r="{far_name}{r}"><v>{r}</v></c></row>"#
            ));
        }
        sheet.push_str("</sheetData></worksheet>");
        names.push_str(&format!(r#"<sheet name="Sheet{i}" sheetId="{i}" r:id="rId{i}"/>"#));
        rels.push_str(&worksheet_rel(&format!("rId{i}"), &format!("worksheets/sheet{i}.xml")));
        parts.push((format!("xl/worksheets/sheet{i}.xml"), sheet.into_bytes()));
    }
    let parts: Vec<(&str, Vec<u8>)> = parts.iter().map(|(n, d)| (n.as_str(), d.clone())).collect();
    package(&names, &rels, &parts)
}

/// Within the budgets, a workbook converted and rendered a sheet at a time is
/// the workbook converted and rendered whole: the same sections, metadata,
/// HTML and Markdown, down to the byte.
#[test]
fn test_the_sheet_by_sheet_path_matches_the_whole_document_one() {
    use office_oxide::ir_render::{HtmlOptions, ImageEmbed, MarkdownOptions};
    let _limits = set_budgets();
    let fixtures = vec![
        ("sparse", sparse_wide_workbook(30, 10)),
        ("wide", sparse_wide_workbook(3, 150)),
        ("merged", merged_workbook()),
        ("prose", prose_workbook()),
        ("pictures", picture_workbook(3, 50, 2)),
        ("comments", comment_fan_out_workbook(5, 5)),
        ("pivots", pivot_fan_out_workbook(2, 5, 4, 3, 4)),
        ("links", hyperlink_fan_out_workbook(5, 20)),
        ("rich", rich_text_fan_out_workbook(5, 3)),
        ("number formats", number_format_workbook("0.00", 5)),
        ("several sheets", multi_sheet_workbook(3, 10, 20)),
    ];
    for (name, bytes) in fixtures {
        let doc = open(bytes).unwrap();
        let ir = doc.to_ir();
        assert!(!ir.metadata.text_truncated, "{name}: not within the budgets");

        let xlsx = doc.as_xlsx().unwrap();
        let mut sheets = xlsx.sheets_ir();
        let sections: Vec<_> = sheets.by_ref().map(|s| s.section).collect();
        assert_eq!(
            serde_json::to_string(&sections).unwrap(),
            serde_json::to_string(&ir.sections).unwrap(),
            "{name}: sections differ"
        );
        assert_eq!(
            serde_json::to_string(&sheets.metadata()).unwrap(),
            serde_json::to_string(&ir.metadata).unwrap(),
            "{name}: metadata differs"
        );

        // Written to a stream, each rendering is the string it would have been.
        let streamed = |write: &dyn Fn(&mut Vec<u8>) -> std::io::Result<()>| {
            let mut out = Vec::new();
            write(&mut out).unwrap();
            String::from_utf8(out).unwrap()
        };
        assert_eq!(streamed(&|o| doc.write_plain_text(o)), doc.plain_text(), "{name}: text");
        assert_eq!(streamed(&|o| doc.write_markdown(o)), doc.to_markdown(), "{name}: Markdown");
        assert_eq!(streamed(&|o| doc.write_html(o)), doc.to_html(), "{name}: written HTML");
        assert_eq!(
            streamed(&|o| doc.write_markdown_with(MarkdownOptions::default(), o)),
            doc.to_markdown_with(MarkdownOptions::default()),
            "{name}: written Markdown with options"
        );
        assert_eq!(doc.to_html(), ir.to_html(), "{name}: HTML differs");
        let embed = MarkdownOptions {
            image_embed: ImageEmbed::Base64,
        };
        assert_eq!(
            doc.to_html_with(HtmlOptions {
                image_embed: ImageEmbed::Base64
            }),
            ir.to_html_with(HtmlOptions {
                image_embed: ImageEmbed::Base64
            }),
            "{name}: embedded-image HTML differs"
        );
        assert_eq!(
            doc.to_markdown_with(MarkdownOptions::default()),
            ir.to_markdown_with(MarkdownOptions::default()),
            "{name}: Markdown differs"
        );
        assert_eq!(
            doc.to_markdown_with(embed),
            ir.to_markdown_with(embed),
            "{name}: embedded-image Markdown differs"
        );
    }
}
