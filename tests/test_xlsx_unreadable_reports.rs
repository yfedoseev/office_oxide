//! What an `.xlsx` reader could not read is reported, in every rendering and
//! in the metadata, and the report is bounded: a workbook that lists many
//! sheets naming one part, with long names, must not be echoed in full by
//! every renderer.
//!
//! The limits are process-global, so every test holds the `LimitScope`.

use std::io::{Cursor, Write};

use office_oxide::limits::{max_text_chars, set_max_text_chars};
use office_oxide::{Document, DocumentFormat};

mod common;
use common::LimitScope;

const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn zip_of(parts: Vec<(String, String)>) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts: zip::write::FileOptions<'_, '_, ()> =
        zip::write::FileOptions::default().compression_method(zip::CompressionMethod::Deflated);
    for (name, data) in parts {
        zip.start_file(name, opts).unwrap();
        zip.write_all(data.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

fn fixed_parts(workbook: String, wb_rels: String) -> Vec<(String, String)> {
    vec![
        (
            "[Content_Types].xml".into(),
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>"#.into(),
        ),
        (
            "_rels/.rels".into(),
            format!(r#"<Relationships xmlns="{REL_NS}"><Relationship Id="rId1" Type="{R}/officeDocument" Target="xl/workbook.xml"/></Relationships>"#),
        ),
        ("xl/workbook.xml".into(), workbook),
        ("xl/_rels/workbook.xml.rels".into(), wb_rels),
    ]
}

const SHEET: &str = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData></worksheet>"#;

/// `sheets` entries of `name_len` characters that all name the one sheet
/// part: the first reads it, the others are unreadable (the part is the
/// content of the first).
fn duplicate_sheets(sheets: usize, name_len: usize) -> Vec<u8> {
    let long = "n".repeat(name_len);
    let mut names = String::new();
    for i in 0..sheets {
        names.push_str(&format!(r#"<sheet name="{i}{long}" sheetId="{}" r:id="rId1"/>"#, i + 1));
    }
    let mut parts = fixed_parts(
        format!(r#"<workbook xmlns="{NS}" xmlns:r="{R}"><sheets>{names}</sheets></workbook>"#),
        format!(
            r#"<Relationships xmlns="{REL_NS}"><Relationship Id="rId1" Type="{R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#
        ),
    );
    parts.push(("xl/worksheets/sheet1.xml".into(), SHEET.into()));
    zip_of(parts)
}

/// One sheet whose relationships part is not XML.
fn unreadable_sheet_rels() -> Vec<u8> {
    let mut parts = fixed_parts(
        format!(
            r#"<workbook xmlns="{NS}" xmlns:r="{R}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>"#
        ),
        format!(
            r#"<Relationships xmlns="{REL_NS}"><Relationship Id="rId1" Type="{R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#
        ),
    );
    parts.push(("xl/worksheets/sheet1.xml".into(), SHEET.into()));
    parts.push(("xl/worksheets/_rels/sheet1.xml.rels".into(), "<Relationships".into()));
    zip_of(parts)
}

fn open(bytes: Vec<u8>) -> Document {
    Document::from_reader(Cursor::new(bytes), DocumentFormat::Xlsx).unwrap()
}

/// Every rendering of `doc`, by name.
fn renderings(doc: &Document) -> Vec<(&'static str, String)> {
    let mut streamed = Vec::new();
    doc.write_markdown_with(Default::default(), &mut streamed)
        .unwrap();
    let mut written = Vec::new();
    doc.write_plain_text(&mut written).unwrap();
    vec![
        ("plain_text", doc.plain_text()),
        ("to_markdown", doc.to_markdown()),
        ("to_html", doc.to_html()),
        ("to_markdown_with", doc.to_markdown_with(Default::default())),
        ("write_markdown_with", String::from_utf8(streamed).unwrap()),
        ("write_plain_text", String::from_utf8(written).unwrap()),
        ("ir_json", serde_json::to_string(&doc.to_ir()).unwrap()),
        ("ir_sections_json", serde_json::to_string(&doc.to_ir().sections).unwrap()),
    ]
}

#[test]
fn test_unreadable_sheets_are_capped_cut_and_counted_in_every_rendering() {
    let _g = LimitScope::new();
    let doc = open(duplicate_sheets(300, 5_000));
    let xlsx = doc.as_xlsx().unwrap();
    // The first entry reads the part; the 299 others are unreadable.
    assert_eq!(xlsx.unreadable_sheets.len(), 100);
    assert_eq!(xlsx.unreadable_sheets_omitted, 199);
    for (name, reason) in &xlsx.unreadable_sheets {
        assert!(name.chars().count() <= 256 && reason.chars().count() <= 256);
    }
    for (what, out) in renderings(&doc) {
        assert!(out.len() < 300_000, "{what}: {} bytes for 299 unreadable sheets", out.len());
        assert!(out.contains("199 more unreadable sheets not listed"), "{what}");
    }
    let ir = doc.to_ir();
    assert!(ir.metadata.text_truncated);
    assert!(ir.metadata.warnings.len() < 250, "{}", ir.metadata.warnings.len());
}

#[test]
fn test_unreadable_sheet_notices_are_charged_to_the_text_budget() {
    let mut limits = LimitScope::new();
    limits.set(max_text_chars, set_max_text_chars, 3_000);
    let doc = open(duplicate_sheets(300, 200));
    for (what, out) in renderings(&doc) {
        // 100 listed sheets would be about 40,000 characters. (The metadata
        // warnings, which `ir_json` includes, are bounded by the cap alone.)
        if what == "ir_json" {
            continue;
        }
        assert!(
            out.len() < 12_000,
            "{what}: {} bytes against a 3,000 character budget",
            out.len()
        );
        assert!(out.contains("output truncated"), "{what}: {out}");
    }
    assert!(doc.to_ir().metadata.text_truncated);
}

#[test]
fn test_an_unreadable_part_is_reported_in_every_rendering() {
    let _g = LimitScope::new();
    let doc = open(unreadable_sheet_rels());
    for (what, out) in renderings(&doc) {
        assert!(
            out.contains("sheet1.xml.rels"),
            "{what} says nothing of the unreadable part: {out}"
        );
        assert!(out.contains('7'), "{what} lost the cell");
    }
    assert!(doc.to_ir().metadata.text_truncated);
}

#[test]
fn test_a_caller_who_drops_each_sheet_ir_still_learns_of_an_unreadable_part() {
    let _g = LimitScope::new();
    let doc = open(unreadable_sheet_rels());
    let xlsx = doc.as_xlsx().unwrap();
    let mut warnings = Vec::new();
    let mut truncated = false;
    for sheet in xlsx.sheets_ir() {
        truncated |= sheet.text_truncated;
        warnings.extend(sheet.warnings);
    }
    assert!(truncated);
    assert!(warnings.iter().any(|w| w.contains("sheet1.xml.rels")), "{warnings:?}");
}

const XDR: &str = r#"xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#;

/// One worksheet `S` whose body ends with `tail`, its relationships `rels`,
/// and the parts `extra`.
fn sheet_with(tail: &str, rels: &str, extra: Vec<(String, String)>) -> Vec<u8> {
    let mut parts = fixed_parts(
        format!(
            r#"<workbook xmlns="{NS}" xmlns:r="{R}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>"#
        ),
        format!(
            r#"<Relationships xmlns="{REL_NS}"><Relationship Id="rId1" Type="{R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#
        ),
    );
    parts.push((
        "xl/worksheets/sheet1.xml".into(),
        format!(
            r#"<worksheet xmlns="{NS}" xmlns:r="{R}"><sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData>{tail}</worksheet>"#
        ),
    ));
    parts.push((
        "xl/worksheets/_rels/sheet1.xml.rels".into(),
        format!(r#"<Relationships xmlns="{REL_NS}">{rels}</Relationships>"#),
    ));
    parts.extend(extra);
    zip_of(parts)
}

/// The reported unreadable parts of a workbook.
fn reported(doc: &Document) -> Vec<String> {
    let xlsx = doc.as_xlsx().unwrap();
    xlsx.unreadable_parts
        .iter()
        .map(|(part, why)| format!("{part}: {why}"))
        .collect()
}

#[test]
fn test_parts_that_cannot_be_read_or_parsed_are_reported_not_swallowed() {
    let _g = LimitScope::new();
    let rel = |id: &str, kind: &str, target: &str| {
        format!(r#"<Relationship Id="{id}" Type="{R}/{kind}" Target="{target}"/>"#)
    };
    let threaded = |target: &str| {
        format!(
            r#"<Relationship Id="rT" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="{target}"/>"#
        )
    };
    let cases: Vec<(&str, Vec<u8>)> = vec![
        // A comments part that is not XML, a table part, a threaded part and
        // a drawing that do not parse.
        (
            "comments1.xml",
            sheet_with(
                "",
                &rel("r1", "comments", "../comments1.xml"),
                vec![("xl/comments1.xml".into(), "<comments".into())],
            ),
        ),
        (
            "table1.xml",
            sheet_with(
                r#"<tableParts count="1"><tablePart r:id="r1"/></tableParts>"#,
                &rel("r1", "table", "../tables/table1.xml"),
                vec![("xl/tables/table1.xml".into(), "<table".into())],
            ),
        ),
        (
            "t1.xml",
            sheet_with(
                "",
                &threaded("../threadedComments/t1.xml"),
                vec![(
                    "xl/threadedComments/t1.xml".into(),
                    "<ThreadedComments><threadedComment".into(),
                )],
            ),
        ),
        (
            "drawing1.xml",
            sheet_with(
                r#"<drawing r:id="r1"/>"#,
                &rel("r1", "drawing", "../drawings/drawing1.xml"),
                vec![(
                    "xl/drawings/drawing1.xml".into(),
                    format!("<xdr:wsDr {XDR}><xdr:oneCellAnchor"),
                )],
            ),
        ),
        // A part that is named, present and not an image is a damaged part;
        // one that is named and absent is below.
        (
            "gone.png",
            sheet_with(
                r#"<drawing r:id="r1"/>"#,
                &rel("r1", "drawing", "../drawings/drawing1.xml"),
                vec![
                    (
                        "xl/drawings/drawing1.xml".into(),
                        format!(
                            r#"<xdr:wsDr {XDR}><xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="100" cy="100"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="p"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>"#
                        ),
                    ),
                    (
                        "xl/drawings/_rels/drawing1.xml.rels".into(),
                        format!(
                            r#"<Relationships xmlns="{REL_NS}">{}</Relationships>"#,
                            rel("rId1", "image", "../media/gone.png")
                        ),
                    ),
                ],
            ),
        ),
    ];
    for (part, bytes) in cases {
        let doc = open(bytes);
        let found = reported(&doc);
        assert!(found.iter().any(|f| f.contains(part)), "{part} is not reported: {found:?}");
        assert!(doc.plain_text().contains(part), "{part} is not in the text");
        assert!(doc.to_ir().metadata.text_truncated, "{part}");
    }
}

/// A chart part named twice by different spellings of its path is one part.
#[test]
fn test_a_chart_named_by_another_spelling_is_not_read_twice() {
    let _g = LimitScope::new();
    let chart = r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>UniqueTitle</a:t></a:r></a:p></c:rich></c:tx></c:title></c:chart></c:chartSpace>"#;
    let mut parts = fixed_parts(
        format!(
            r#"<workbook xmlns="{NS}" xmlns:r="{R}"><sheets><sheet name="C" sheetId="1" r:id="rId1"/></sheets></workbook>"#
        ),
        format!(
            r#"<Relationships xmlns="{REL_NS}"><Relationship Id="rId1" Type="{R}/chartsheet" Target="chartsheets/sheet1.xml"/></Relationships>"#
        ),
    );
    parts.push((
        "xl/chartsheets/sheet1.xml".into(),
        format!(r#"<chartsheet xmlns="{NS}" xmlns:r="{R}"><drawing r:id="rId1"/></chartsheet>"#),
    ));
    parts.push((
        "xl/chartsheets/_rels/sheet1.xml.rels".into(),
        format!(
            r#"<Relationships xmlns="{REL_NS}"><Relationship Id="rId1" Type="{R}/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#
        ),
    ));
    parts.push(("xl/drawings/drawing1.xml".into(), format!("<xdr:wsDr {XDR}/>")));
    // The relationship spells the chart's path in capitals and with a
    // backslash; it is the part `xl/charts/chart1.xml` all the same.
    parts.push((
        "xl/drawings/_rels/drawing1.xml.rels".into(),
        format!(
            r#"<Relationships xmlns="{REL_NS}"><Relationship Id="rId1" Type="{R}/chart" Target="..\Charts\CHART1.XML"/></Relationships>"#
        ),
    ));
    parts.push(("xl/charts/chart1.xml".into(), chart.into()));
    let doc = open(zip_of(parts));
    assert_eq!(doc.plain_text().matches("UniqueTitle").count(), 1, "{}", doc.plain_text());
    assert_eq!(doc.to_markdown().matches("UniqueTitle").count(), 1);
    assert_eq!(
        serde_json::to_string(&doc.to_ir())
            .unwrap()
            .matches("UniqueTitle")
            .count(),
        1
    );
}

/// The relationship of `kind` to `target`, internal or external.
fn rel_to(kind: &str, target: &str, external: bool) -> String {
    let mode = if external {
        r#" TargetMode="External""#
    } else {
        ""
    };
    format!(r#"<Relationship Id="r1" Type="{R}/{kind}" Target="{target}"{mode}/>"#)
}

fn assert_silent(label: &str, bytes: Vec<u8>) {
    let doc = open(bytes);
    assert_eq!(reported(&doc), Vec::<String>::new(), "{label}");
    let ir = doc.to_ir();
    assert!(!ir.metadata.text_truncated, "{label}: {:?}", ir.metadata.warnings);
    assert!(ir.metadata.warnings.is_empty(), "{label}: {:?}", ir.metadata.warnings);
    assert!(!doc.plain_text().contains("skipped"), "{label}");
}

/// An optional part a relationship names but the package lacks is a dangling
/// reference, as it always was: nothing is reported. Only a part that exists
/// and cannot be read or parsed is.
#[test]
fn test_a_relationship_to_an_absent_optional_part_is_not_reported() {
    let _g = LimitScope::new();
    let drawing = r#"<drawing r:id="r1"/>"#;
    let table = r#"<tableParts count="1"><tablePart r:id="r1"/></tableParts>"#;
    for (label, tail, rels) in [
        ("drawing", drawing, rel_to("drawing", "../drawings/gone.xml", false)),
        ("table", table, rel_to("table", "../tables/gone.xml", false)),
        ("comments", "", rel_to("comments", "../gone.xml", false)),
        (
            "threaded",
            "",
            rel_to("threadedComment", "../threadedComments/gone.xml", false).replace(
                &format!("{R}/threadedComment"),
                "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment",
            ),
        ),
        ("pivot", "", rel_to("pivotTable", "../pivotTables/gone.xml", false)),
    ] {
        assert_silent(label, sheet_with(tail, &rels, vec![]));
    }
}

/// A target outside the package is never looked up in it, whatever part of
/// the package happens to share its name.
#[test]
fn test_an_external_target_is_not_looked_up_in_the_package() {
    let _g = LimitScope::new();
    for (label, tail, rels, name) in [
        (
            "drawing",
            r#"<drawing r:id="r1"/>"#,
            rel_to("drawing", "../d.xml", true),
            "xl/d.xml",
        ),
        (
            "table",
            r#"<tableParts count="1"><tablePart r:id="r1"/></tableParts>"#,
            rel_to("table", "../d.xml", true),
            "xl/d.xml",
        ),
        ("comments", "", rel_to("comments", "../d.xml", true), "xl/d.xml"),
        ("pivot", "", rel_to("pivotTable", "../d.xml", true), "xl/d.xml"),
    ] {
        // The name is a part that would be reported were it read.
        assert_silent(label, sheet_with(tail, &rels, vec![(name.into(), "<broken".into())]));
    }
}

/// A name that would be markup, a link, an image or a heading if written as it is.
const HOSTILE: &str = "x](http:evil)<img src=x onerror=1>`\n\n# Heading";
const HOSTILE_XML: &str = "x](http:evil)&lt;img src=x onerror=1&gt;`&#10;&#10;# Heading";

/// Every string in the part of the IR that reports losses (the sections after
/// the sheet's own).
fn report_strings(doc: &Document) -> Vec<String> {
    fn walk(v: &serde_json::Value, out: &mut Vec<String>) {
        match v {
            serde_json::Value::String(s) => out.push(s.clone()),
            serde_json::Value::Array(a) => a.iter().for_each(|v| walk(v, out)),
            serde_json::Value::Object(o) => o.values().for_each(|v| walk(v, out)),
            _ => {},
        }
    }
    let mut out = Vec::new();
    for section in doc.to_ir().sections.iter().skip(1) {
        walk(&serde_json::to_value(section).unwrap(), &mut out);
    }
    out
}

/// What the file says about a sheet or part is echoed in the notices, and
/// stays one inert line in every rendering: not a heading, a link or markup.
#[test]
fn test_names_echoed_in_the_notices_are_inert_in_every_rendering() {
    let _g = LimitScope::new();
    let mut parts = fixed_parts(
        format!(
            r#"<workbook xmlns="{NS}" xmlns:r="{R}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/><sheet name="{HOSTILE_XML}" sheetId="2" r:id="rId1"/></sheets></workbook>"#
        ),
        format!(
            r#"<Relationships xmlns="{REL_NS}"><Relationship Id="rId1" Type="{R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#
        ),
    );
    parts.push((
        "xl/worksheets/sheet1.xml".into(),
        SHEET
            .replace("</worksheet>", r#"<legacyDrawing r:id="r1"/></worksheet>"#)
            .replace("<worksheet ", &format!(r#"<worksheet xmlns:r="{R}" "#)),
    ));
    parts.push((
        "xl/worksheets/_rels/sheet1.xml.rels".into(),
        format!(
            r#"<Relationships xmlns="{REL_NS}"><Relationship Id="r1" Type="{R}/comments" Target="../{HOSTILE_XML}.xml"/></Relationships>"#
        ),
    ));
    parts.push((format!("xl/{HOSTILE}.xml"), "<comments".into()));
    let doc = open(zip_of(parts));
    let xlsx = doc.as_xlsx().unwrap();
    assert!(!xlsx.unreadable_sheets.is_empty() && !xlsx.unreadable_parts.is_empty());
    for (what, out) in renderings(&doc) {
        if what.starts_with("ir_") {
            continue;
        }
        for line in out.lines() {
            assert!(!line.starts_with("# Heading"), "{what}: a heading from a name: {out}");
            assert!(!line.starts_with("## x]"), "{what}: a heading from a name: {out}");
        }
        if what.contains("markdown") {
            assert_eq!(
                out.matches("](http:evil)").count(),
                out.matches("\\](http:evil)").count(),
                "{what}: a link from a name: {out}"
            );
            let tags = out.matches("<img").count();
            assert_eq!(tags, out.matches("\\<img").count(), "{what}: markup from a name: {out}");
        }
        if what == "to_html" {
            assert!(!out.contains("<img"), "{what}: markup from a name: {out}");
        }
    }
    for s in report_strings(&doc) {
        assert!(!s.chars().any(char::is_control), "a control character in the IR: {s:?}");
    }
    // The metadata keeps the bounded text as it is.
    let ir = doc.to_ir();
    assert!(
        ir.metadata.warnings.iter().any(|w| w.contains("evil")),
        "{:?}",
        ir.metadata.warnings
    );
}

/// A notice the text budget cannot pay is replaced by the budget's notice in
/// the content, and its warning still reaches a caller who drops each sheet.
#[test]
fn test_a_notice_the_budget_cannot_pay_keeps_its_warning_in_the_streamed_sheet() {
    let mut limits = LimitScope::new();
    limits.set(max_text_chars, set_max_text_chars, 3_000);
    let doc = open(duplicate_sheets(300, 200));
    let xlsx = doc.as_xlsx().unwrap();
    let mut warnings = Vec::new();
    for sheet in xlsx.sheets_ir() {
        warnings.extend(sheet.warnings);
    }
    let sheets = warnings
        .iter()
        .filter(|w| w.starts_with("skipped unreadable sheet"))
        .count();
    assert_eq!(sheets, xlsx.unreadable_sheets.len(), "{} warnings", warnings.len());
    assert!(
        warnings
            .iter()
            .any(|w| w.contains("199 more unreadable sheets")),
        "{}",
        warnings.len()
    );
}
