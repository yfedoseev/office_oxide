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
