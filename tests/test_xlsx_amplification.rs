//! A spreadsheet can name one shared thing from many places: a formula
//! master shared by many cells, a picture drawn from many anchors, a
//! comment author on many comments. The thing is stored once; each place
//! gets a copy. These tests build a small workbook that fans out one
//! shared item, lower the document budgets so the product is small, and
//! assert the product is bounded (an `Err` or a documented truncation)
//! rather than materialised.
//!
//! The budgets are process-global, so every test in this binary sets the
//! same low values and builds documents of a few kilobytes.

use std::io::{Cursor, Write};

use office_oxide::core::Error as CoreError;
use office_oxide::limits::set_max_text_chars;
use office_oxide::xlsx::XlsxError;
use office_oxide::{Document, DocumentFormat, OfficeError};

/// The message of a refusal for exceeding a resource bound, or `None` when
/// `result` is not one. A worksheet that exceeds a bound is unreadable, and a
/// workbook with no readable sheet is refused with that sheet's reason.
fn limit_message<T>(result: &office_oxide::Result<T>) -> Option<String> {
    match result {
        Err(OfficeError::Core(CoreError::PackageLimit(m))) => Some(m.clone()),
        Err(OfficeError::Core(CoreError::MalformedXml(m)))
        | Err(OfficeError::Xlsx(XlsxError::Core(CoreError::MalformedXml(m))))
            if m.contains("package limit exceeded") =>
        {
            Some(m.clone())
        },
        _ => None,
    }
}

/// The text budget every test here runs under.
const TEXT_BUDGET: usize = 60_000;

fn set_budgets() {
    set_max_text_chars(TEXT_BUDGET);
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

#[test]
fn test_shared_formula_master_copied_into_many_followers_is_bounded() {
    set_budgets();
    // 4,000 characters (under the 8,192 limit) copied into 100 followers is
    // 400,000 characters against a 60,000 budget.
    let result = open(shared_formula_workbook(4_000, 100));
    let message = limit_message(&result).unwrap_or_else(|| {
        panic!("expected a package limit error, got {:?}", result.as_ref().map(|_| ()))
    });
    assert!(message.contains("shared group 0"), "{message}");
}

#[test]
fn test_shared_formula_master_over_excel_limit_is_refused() {
    set_budgets();
    let result = open(shared_formula_workbook(9_000, 1));
    let message = limit_message(&result).expect("a package limit error");
    assert!(message.contains("8192"), "{message}");
}

#[test]
fn test_shared_formula_within_budget_still_translates() {
    set_budgets();
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
    set_budgets();
    // 2,000 rows, each laid out 200 columns wide: 3 characters of padding
    // per column in Markdown is 1.2 MB against a 60,000 budget.
    let doc = open(sparse_wide_workbook(2_000, 200)).unwrap();
    let md = doc.to_markdown();
    assert!(md.contains("character budget"), "Markdown carries no truncation notice");
    assert!(md.len() < TEXT_BUDGET + 2_000, "{} bytes of Markdown", md.len());
}

#[test]
fn test_ir_table_rows_padded_to_the_sheet_width_are_bounded() {
    set_budgets();
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
    assert!(html.contains("character budget"), "HTML carries no truncation notice");
}

#[test]
fn test_a_dense_wide_sheet_is_unaffected_by_padding_charges() {
    set_budgets();
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
    set_budgets();
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
    let doc = open(package(&sheets, &rels, &parts)).unwrap();
    let text = doc.plain_text();
    assert!(
        text.contains("unreadable sheet"),
        "the sheets past the document total are reported"
    );
    assert!(text.contains("S1"), "the sheets within it are read");
}

#[test]
fn test_padding_is_charged_after_the_text_has_spent_its_budget() {
    set_budgets();
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
    package(
        r#"<sheet name="S" sheetId="1" r:id="rId1"/>"#,
        &worksheet_rel("rId1", "worksheets/sheet1.xml"),
        &[
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
        ],
    )
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
    set_budgets();
    office_oxide::limits::set_max_image_bytes(IMAGE_BUDGET);
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
fn test_many_distinct_small_pictures_within_the_budget_are_all_kept() {
    set_budgets();
    office_oxide::limits::set_max_image_bytes(IMAGE_BUDGET);
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
    set_budgets();
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
    set_budgets();
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
    set_budgets();
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
    set_budgets();
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
    set_budgets();
    let doc = open(pivot_fan_out_workbook(2, 5, 20, 1, 1)).unwrap();
    assert!(!doc.to_markdown().contains("character budget"));
    assert!(!doc.to_ir().metadata.text_truncated);
}

#[test]
fn test_sheets_naming_one_part_are_read_once_and_the_rest_reported() {
    set_budgets();
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
    set_budgets();
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
    set_budgets();
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
