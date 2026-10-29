//! Options that belong to one document, not to the process.
//!
//! `Document::open_with` keeps a `ParseOptions` (limits, which text is kept,
//! what to do with an incomplete result) with the document, and every
//! conversion of it runs under them, on any thread, without touching the
//! process-wide settings other documents read.

mod common;

use std::io::Cursor;

use common::LimitScope;
use office_oxide::core::opc::{OpcWriter, PartName};
use office_oxide::core::relationships::rel_types;
use office_oxide::limits::{self, Limits};
use office_oxide::options::{DamageKind, OnIncomplete, ParseOptions};
use office_oxide::{Document, DocumentFormat, OfficeError};

const WML: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

fn docx(body: &str) -> Vec<u8> {
    let xml = format!(r#"<w:document xmlns:w="{WML}"><w:body>{body}</w:body></w:document>"#);
    let mut writer = OpcWriter::new(Cursor::new(Vec::new())).unwrap();
    let part = PartName::new("/word/document.xml").unwrap();
    writer
        .add_part(
            &part,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            xml.as_bytes(),
        )
        .unwrap();
    writer.add_package_rel(rel_types::OFFICE_DOCUMENT, "word/document.xml");
    writer.finish().unwrap().into_inner()
}

fn paragraphs(n: usize) -> String {
    (0..n)
        .map(|i| format!("<w:p><w:r><w:t>paragraph number {i}</w:t></w:r></w:p>"))
        .collect()
}

fn open_with(bytes: &[u8], options: &ParseOptions) -> office_oxide::Result<Document> {
    Document::from_reader_with(Cursor::new(bytes.to_vec()), DocumentFormat::Docx, options)
}

/// One 5,000-character shared string referenced from `cells` cells of one
/// row: 5,000 x `cells` characters of output from a few KB of zip.
fn xlsx_fan_out(cells: usize) -> Vec<u8> {
    use std::io::Write;
    let big = "x".repeat(5_000);
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts: zip::write::FileOptions<'_, '_, ()> = zip::write::FileOptions::default();
    let mut row = String::from(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1">"#,
    );
    for c in 0..cells {
        row.push_str(&format!(r#"<c r="{}1" t="s"><v>0</v></c>"#, (b'A' + c as u8) as char));
    }
    row.push_str("</row></sheetData></worksheet>");
    let parts: Vec<(&str, String)> = vec![
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>"#.into()),
        ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#.into()),
        ("xl/workbook.xml", r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>"#.into()),
        ("xl/_rels/workbook.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>"#.into()),
        ("xl/sharedStrings.xml", format!(r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>{big}</t></si></sst>"#)),
        ("xl/worksheets/sheet1.xml", row),
    ];
    for (name, data) in parts {
        zip.start_file(name, opts).unwrap();
        zip.write_all(data.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

fn open_xlsx_with(bytes: &[u8], options: &ParseOptions) -> office_oxide::Result<Document> {
    Document::from_reader_with(Cursor::new(bytes.to_vec()), DocumentFormat::Xlsx, options)
}

#[test]
fn test_limits_of_one_document_do_not_change_another() {
    let _scope = LimitScope::new();
    let bytes = xlsx_fan_out(20);
    let mut small = ParseOptions::default();
    small.limits.text_chars = 12_000;
    let capped = open_xlsx_with(&bytes, &small).unwrap();
    let plain = Document::from_reader(Cursor::new(bytes.clone()), DocumentFormat::Xlsx).unwrap();

    let capped_text = capped.plain_text();
    let plain_text = plain.plain_text();
    assert!(
        capped_text.len() < plain_text.len() / 4,
        "{} vs {}",
        capped_text.len(),
        plain_text.len()
    );
    assert!(capped.to_ir().metadata.text_truncated);
    assert!(!plain.to_ir().metadata.text_truncated);
    // The process-wide setting was never touched.
    assert_eq!(limits::max_text_chars(), limits::DEFAULT_MAX_TEXT_CHARS);
}

#[test]
fn test_two_threads_read_under_their_own_limits_at_the_same_time() {
    let _scope = LimitScope::new();
    let bytes = xlsx_fan_out(20);
    let barrier = std::sync::Arc::new(std::sync::Barrier::new(2));
    let run = |chars: usize| {
        let bytes = bytes.clone();
        let barrier = barrier.clone();
        std::thread::spawn(move || {
            let mut options = ParseOptions::default();
            options.limits.text_chars = chars;
            let doc = open_xlsx_with(&bytes, &options).unwrap();
            barrier.wait();
            (0..20).map(|_| doc.plain_text().len()).max().unwrap()
        })
    };
    let (small, large) = (run(12_000), run(usize::MAX));
    let (small, large) = (small.join().unwrap(), large.join().unwrap());
    assert!(small * 4 < large, "{small} vs {large}");
}

#[test]
fn test_a_document_converts_under_its_options_on_another_thread() {
    let _scope = LimitScope::new();
    let bytes = xlsx_fan_out(20);
    let mut options = ParseOptions::default();
    options.limits.text_chars = 12_000;
    let doc = open_xlsx_with(&bytes, &options).unwrap();
    let here = doc.plain_text();
    let there = std::thread::spawn(move || doc.plain_text()).join().unwrap();
    assert_eq!(here, there);
}

#[test]
fn test_hidden_text_is_kept_for_one_document_only() {
    let _scope = LimitScope::new();
    let bytes = docx(
        r#"<w:p><w:r><w:t>shown </w:t></w:r><w:r><w:rPr><w:vanish/></w:rPr><w:t>secret</w:t></w:r></w:p>"#,
    );
    let mut options = ParseOptions::default();
    options.include_hidden_text = true;
    let with = open_with(&bytes, &options).unwrap();
    let without = open_with(&bytes, &ParseOptions::default()).unwrap();
    assert!(with.plain_text().contains("secret"));
    assert!(!without.plain_text().contains("secret"));
    assert!(!office_oxide::read_options::include_hidden_text());
}

#[test]
fn test_a_part_limit_in_the_options_refuses_the_open() {
    let _scope = LimitScope::new();
    let bytes = docx(&paragraphs(5000));
    let mut options = ParseOptions::default();
    options.limits.part_bytes = 4096;
    assert!(open_with(&bytes, &options).is_err());
    assert!(open_with(&bytes, &ParseOptions::default()).is_ok());
    assert_eq!(limits::max_part_bytes(), limits::DEFAULT_MAX_PART_BYTES);
}

#[test]
fn test_an_incomplete_result_is_an_error_only_when_asked() {
    let _scope = LimitScope::new();
    // A body cut off part-way: the paragraphs before the cut are kept.
    let xml = format!(r#"<w:document xmlns:w="{WML}"><w:body>{}<w:p><w:r><w:t>cut"#, paragraphs(3));
    let mut writer = OpcWriter::new(Cursor::new(Vec::new())).unwrap();
    let part = PartName::new("/word/document.xml").unwrap();
    writer
        .add_part(
            &part,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            xml.as_bytes(),
        )
        .unwrap();
    writer.add_package_rel(rel_types::OFFICE_DOCUMENT, "word/document.xml");
    let bytes = writer.finish().unwrap().into_inner();

    let reported = open_with(&bytes, &ParseOptions::default()).unwrap();
    let damage = reported.damage();
    assert!(
        damage.iter().any(|d| d.kind == DamageKind::UnreadablePart
            && d.part
                .as_deref()
                .is_some_and(|p| p.ends_with("word/document.xml"))
            && d.is_loss()),
        "{damage:?}"
    );

    let mut strict = ParseOptions::default();
    strict.on_incomplete = OnIncomplete::Error;
    match open_with(&bytes, &strict) {
        Err(OfficeError::Incomplete(why)) => assert!(why.contains("word/document.xml"), "{why}"),
        Err(other) => panic!("a different error: {other}"),
        Ok(_) => panic!("an incomplete document was accepted"),
    }
}

#[test]
fn test_a_sound_document_is_accepted_in_strict_mode_and_has_no_damage() {
    let _scope = LimitScope::new();
    let bytes = docx(&paragraphs(3));
    let doc = open_with(&bytes, &ParseOptions::strict()).unwrap();
    assert!(doc.damage().iter().all(|d| !d.is_loss()), "{:?}", doc.damage());
    assert!(doc.plain_text().contains("paragraph number 2"));
}

#[test]
fn test_a_limit_that_cuts_the_text_is_an_incomplete_result() {
    let _scope = LimitScope::new();
    let bytes = xlsx_fan_out(20);
    let mut options = ParseOptions::default();
    options.limits.text_chars = 12_000;
    let doc = open_xlsx_with(&bytes, &options).unwrap();
    assert!(doc.damage().iter().any(|d| d.kind == DamageKind::Truncated));
    options.on_incomplete = OnIncomplete::Error;
    assert!(matches!(open_xlsx_with(&bytes, &options), Err(OfficeError::Incomplete(_))));
}

#[test]
fn test_limits_constructors_and_apply() {
    let mut scope = LimitScope::new();
    let defaults = Limits::default();
    assert_eq!(defaults.text_chars, limits::DEFAULT_MAX_TEXT_CHARS);
    assert_eq!(Limits::current(), defaults);

    let strict = Limits::strict();
    assert!(strict.text_chars < defaults.text_chars);
    assert!(strict.part_bytes < defaults.part_bytes);
    assert_eq!(strict.embedded_depth, 1);
    assert_eq!(Limits::unlimited().text_chars, usize::MAX);
    assert_eq!(Limits::unlimited().part_bytes, u64::MAX);

    // `apply` is the `set_max_*` functions in one call; put them back after.
    scope.set(limits::max_text_chars, limits::set_max_text_chars, defaults.text_chars);
    scope.set(limits::max_part_bytes, limits::set_max_part_bytes, defaults.part_bytes);
    strict.apply();
    assert_eq!(limits::max_text_chars(), strict.text_chars);
    assert_eq!(Limits::current(), strict);
    defaults.apply();
}

#[test]
fn test_strict_options_read_an_ordinary_document() {
    let _scope = LimitScope::new();
    let bytes = docx(&paragraphs(200));
    let doc = open_with(&bytes, &ParseOptions::strict()).unwrap();
    assert!(doc.plain_text().contains("paragraph number 199"));
}

#[test]
fn test_saving_runs_under_the_documents_own_write_limit() {
    let _scope = LimitScope::new();
    let cells: String = (0..50)
        .map(|r| {
            format!(
                "<w:tr><w:tc><w:p><w:r><w:t>a{r}</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>b{r}</w:t></w:r></w:p></w:tc></w:tr>"
            )
        })
        .collect();
    let bytes = docx(&format!("<w:tbl>{cells}</w:tbl>"));
    let dir = std::env::temp_dir().join(format!("oo-save-limit-{}", std::process::id()));
    std::fs::create_dir_all(&dir).unwrap();
    let target = dir.join("out.docx");

    let mut options = ParseOptions::default();
    options.limits.write_bytes = 1000;
    let limited = open_with(&bytes, &options).unwrap();
    let refused = limited.save_as(&target);
    assert!(refused.is_err(), "100 cells cost far more than 1,000 bytes to write");

    let plain = open_with(&bytes, &ParseOptions::default()).unwrap();
    plain.save_as(&target).unwrap();
    assert_eq!(limits::max_write_bytes(), limits::DEFAULT_MAX_WRITE_BYTES);
    let _ = std::fs::remove_dir_all(&dir);
}

#[test]
fn test_every_named_limit_can_be_set_by_name_and_unknown_names_are_refused() {
    let mut options = ParseOptions::default();
    for name in Limits::NAMES {
        options
            .set_option(name, 12_345)
            .unwrap_or_else(|e| panic!("{name}: {e}"));
    }
    assert_eq!(options.limits.text_chars, 12_345);
    assert_eq!(options.limits.part_bytes, 12_345);
    options.set_option("part_bytes", u64::MAX).unwrap();
    assert_eq!(options.limits.part_bytes, u64::MAX);

    options.set_option("include_hidden_text", 1).unwrap();
    options.set_option("on_incomplete", 1).unwrap();
    assert!(options.include_hidden_text);
    assert_eq!(options.on_incomplete, OnIncomplete::Error);
    options.set_option("on_incomplete", 0).unwrap();
    assert_eq!(options.on_incomplete, OnIncomplete::Report);

    let err = options.set_option("no_such_option", 1).unwrap_err();
    assert!(err.to_string().contains("text_chars"), "{err}");
    assert!(Limits::default().set("no_such_limit", 1).is_err());
}
