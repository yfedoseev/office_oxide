//! A value stored once in the package and copied into every list, run or
//! paragraph that uses it: a numbering level's `w:lvlText`, a style's font
//! name, a style's tab stops. The reader keeps one copy; every copy that
//! lands in an IR is charged to the text budget first, and the conversion
//! says so when the budget refuses one.
//!
//! These tests lower the process-wide text budget, so they have a test binary
//! to themselves, and each sets the same value.

use std::io::{Cursor, Write};

use office_oxide::ir::{DocumentIR, Element, InlineContent};
use office_oxide::limits::set_max_text_chars;
use office_oxide::{Document, DocumentFormat};

const NS: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#;
const PKG_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// The length of the value stored once: a large but storable string.
const BIG: usize = 50_000;
/// About six copies of `BIG` fit.
const TEXT_BUDGET: usize = 300_000;
/// How many users of the shared value the document has.
const USERS: usize = 40;

fn limits() {
    set_max_text_chars(TEXT_BUDGET);
}

fn zip_of(parts: Vec<(String, String)>) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts = zip::write::SimpleFileOptions::default();
    for (name, data) in parts {
        zip.start_file(name, opts).unwrap();
        zip.write_all(data.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

/// A `.docx` with `body`, plus the optional `styles.xml` and `numbering.xml`.
fn docx(body: &str, styles: Option<&str>, numbering: Option<&str>) -> Vec<u8> {
    let mut ct = String::new();
    let mut rels = String::new();
    let mut parts = vec![];
    if let Some(s) = styles {
        ct.push_str(r#"<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>"#);
        rels.push_str(&format!(
            r#"<Relationship Id="rId3" Type="{REL}/styles" Target="styles.xml"/>"#
        ));
        parts.push(("word/styles.xml".to_string(), s.to_string()));
    }
    if let Some(n) = numbering {
        ct.push_str(r#"<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>"#);
        rels.push_str(&format!(
            r#"<Relationship Id="rId4" Type="{REL}/numbering" Target="numbering.xml"/>"#
        ));
        parts.push(("word/numbering.xml".to_string(), n.to_string()));
    }
    let mut all = vec![
        (
            "[Content_Types].xml".to_string(),
            format!(
                r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>{ct}</Types>"#
            ),
        ),
        (
            "_rels/.rels".to_string(),
            format!(
                r#"<Relationships xmlns="{PKG_REL}"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="word/document.xml"/></Relationships>"#
            ),
        ),
        (
            "word/_rels/document.xml.rels".to_string(),
            format!(r#"<Relationships xmlns="{PKG_REL}">{rels}</Relationships>"#),
        ),
        (
            "word/document.xml".to_string(),
            format!(r#"<w:document {NS}><w:body>{body}</w:body></w:document>"#),
        ),
    ];
    all.extend(parts);
    zip_of(all)
}

fn to_ir(bytes: Vec<u8>) -> DocumentIR {
    Document::from_reader(Cursor::new(bytes), DocumentFormat::Docx)
        .unwrap()
        .to_ir()
}

fn assert_reported(ir: &DocumentIR) {
    assert!(ir.metadata.text_truncated, "a refused copy must set the truncation flag");
    assert!(
        ir.metadata
            .warnings
            .iter()
            .any(|w| w.contains("set_max_text_chars")),
        "a refused copy must be reported: {:?}",
        ir.metadata.warnings
    );
}

fn paragraphs(ir: &DocumentIR) -> Vec<&office_oxide::ir::Paragraph> {
    ir.sections
        .iter()
        .flat_map(|s| s.elements.iter())
        .filter_map(|e| match e {
            Element::Paragraph(p) => Some(p),
            _ => None,
        })
        .collect()
}

fn lists(ir: &DocumentIR) -> Vec<&office_oxide::ir::List> {
    ir.sections
        .iter()
        .flat_map(|s| s.elements.iter())
        .filter_map(|e| match e {
            Element::List(l) => Some(l),
            _ => None,
        })
        .collect()
}

fn numbering(level_text: &str) -> String {
    format!(
        r#"<w:numbering {NS}><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="{level_text}"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>"#
    )
}

fn chart_part(text: &str) -> String {
    format!(
        r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>{text}</a:t></a:r></a:p></c:rich></c:tx></c:title></c:chart></c:chartSpace>"#
    )
}

/// `n` paragraphs, each holding a drawing of the chart `rId9`.
fn chart_body(n: usize) -> String {
    r#"<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="100" cy="100"/><wp:docPr id="1" name="c"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId9"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"#.repeat(n)
}

/// A `.docx` with `drawings` drawings of one chart whose title is `text`.
fn chart_docx(text: &str, drawings: usize) -> Vec<u8> {
    let body = chart_body(drawings);
    let mut all = vec![
        ("[Content_Types].xml".to_string(), r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>"#.to_string()),
        ("_rels/.rels".to_string(), format!(r#"<Relationships xmlns="{PKG_REL}"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="word/document.xml"/></Relationships>"#)),
        ("word/_rels/document.xml.rels".to_string(), format!(r#"<Relationships xmlns="{PKG_REL}"><Relationship Id="rId9" Type="{REL}/chart" Target="charts/chart1.xml"/></Relationships>"#)),
        ("word/document.xml".to_string(), format!(r#"<w:document {NS}><w:body>{body}</w:body></w:document>"#)),
        ("word/charts/chart1.xml".to_string(), chart_part(text)),
    ];
    zip_of(std::mem::take(&mut all))
}

/// `lists` one-item lists, each ended by an ordinary paragraph.
fn list_body(lists: usize) -> String {
    r#"<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>x</w:t></w:r></w:p><w:p><w:r><w:t>y</w:t></w:r></w:p>"#
        .repeat(lists)
}

#[test]
fn test_marker_pattern_copies_per_list_are_charged() {
    limits();
    let num = numbering(&format!("{}%1.", "x".repeat(BIG)));
    let ir = to_ir(docx(&list_body(USERS), None, Some(&num)));
    let copied: usize = lists(&ir)
        .iter()
        .map(|l| l.marker_pattern.as_deref().map_or(0, str::len))
        .sum();
    assert!(copied > 0, "the budget holds a few copies");
    assert!(
        copied <= TEXT_BUDGET,
        "{copied} characters of pattern were copied against a budget of {TEXT_BUDGET}"
    );
    assert_eq!(lists(&ir).len(), USERS, "the lists themselves stay");
    assert_reported(&ir);
}

#[test]
fn test_small_marker_pattern_is_kept_on_every_list() {
    limits();
    let ir = to_ir(docx(&list_body(USERS), None, Some(&numbering("(%1)"))));
    assert!(
        lists(&ir)
            .iter()
            .all(|l| l.marker_pattern.as_deref() == Some("(%1)"))
    );
    assert!(!ir.metadata.text_truncated);
}

fn font_styles(name: &str) -> String {
    format!(
        r#"<w:styles {NS}><w:style w:type="paragraph" w:styleId="s"><w:name w:val="S"/><w:rPr><w:rFonts w:ascii="{name}" w:hAnsi="{name}"/></w:rPr></w:style></w:styles>"#
    )
}

/// `n` styled paragraphs, each with one bold run (direct formatting).
fn bold_body(n: usize) -> String {
    r#"<w:p><w:pPr><w:pStyle w:val="s"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>x</w:t></w:r></w:p>"#
        .repeat(n)
}

fn span_font_names(ir: &DocumentIR) -> Vec<Option<&str>> {
    paragraphs(ir)
        .iter()
        .flat_map(|p| p.content.iter())
        .filter_map(|c| match c {
            InlineContent::Text(t) => Some(t.font_name.as_deref()),
            _ => None,
        })
        .collect()
}

#[test]
fn test_style_font_name_copies_per_run_are_charged() {
    limits();
    let styles = font_styles(&"F".repeat(BIG));
    let ir = to_ir(docx(&bold_body(USERS), Some(&styles), None));
    let names = span_font_names(&ir);
    let copied: usize = names.iter().flatten().map(|n| n.len()).sum();
    assert!(copied > 0, "the budget holds a few copies");
    assert!(
        copied <= TEXT_BUDGET,
        "{copied} characters of font name were copied against a budget of {TEXT_BUDGET}"
    );
    assert_eq!(names.len(), USERS, "the runs themselves stay");
    assert_reported(&ir);
}

#[test]
fn test_small_style_font_name_is_kept_on_every_run() {
    limits();
    let ir = to_ir(docx(&bold_body(USERS), Some(&font_styles("Arial")), None));
    let names = span_font_names(&ir);
    assert_eq!(names.len(), USERS);
    assert!(names.iter().all(|n| *n == Some("Arial")));
    assert!(!ir.metadata.text_truncated);
}

fn tab_styles(tabs: usize) -> String {
    let mut t = String::from("<w:pPr><w:tabs>");
    for i in 0..tabs {
        t.push_str(&format!(r#"<w:tab w:val="left" w:pos="{}"/>"#, i + 1));
    }
    t.push_str("</w:tabs></w:pPr>");
    format!(
        r#"<w:styles {NS}><w:style w:type="paragraph" w:styleId="s"><w:name w:val="S"/>{t}</w:style></w:styles>"#
    )
}

fn styled_body(n: usize) -> String {
    r#"<w:p><w:pPr><w:pStyle w:val="s"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>"#.repeat(n)
}

#[test]
fn test_style_tab_stop_copies_per_paragraph_are_charged() {
    limits();
    let ir = to_ir(docx(&styled_body(USERS), Some(&tab_styles(1_000)), None));
    let paras = paragraphs(&ir);
    assert_eq!(paras.len(), USERS, "the paragraphs themselves stay");
    let tabs: usize = paras.iter().map(|p| p.tabs.len()).sum();
    assert!(tabs > 0, "the budget holds a few copies");
    // Every tab stop serialises to at least 64 characters.
    assert!(
        tabs * 64 <= TEXT_BUDGET,
        "{tabs} tab stops were copied against a budget of {TEXT_BUDGET} characters"
    );
    assert_reported(&ir);
}

#[test]
fn test_few_style_tab_stops_are_kept_on_every_paragraph() {
    limits();
    let ir = to_ir(docx(&styled_body(USERS), Some(&tab_styles(3)), None));
    let paras = paragraphs(&ir);
    assert_eq!(paras.len(), USERS);
    assert!(paras.iter().all(|p| p.tabs.len() == 3));
    assert!(!ir.metadata.text_truncated);
}

#[test]
fn test_chart_text_copies_per_drawing_are_charged() {
    limits();
    let bytes = chart_docx(&"c".repeat(BIG), USERS);
    let doc = Document::from_reader(Cursor::new(bytes), DocumentFormat::Docx).unwrap();
    let ir = doc.to_ir();
    let copied: usize = paragraphs(&ir)
        .iter()
        .flat_map(|p| p.content.iter())
        .map(|c| match c {
            InlineContent::Text(t) => t.text.len(),
            _ => 0,
        })
        .sum();
    assert!(copied >= BIG, "the budget holds a few copies");
    assert!(
        copied <= TEXT_BUDGET,
        "{copied} characters of chart text were copied against a budget of {TEXT_BUDGET}"
    );
    assert_reported(&ir);
    for rendering in [doc.plain_text(), doc.to_markdown()] {
        assert!(
            rendering.len() <= TEXT_BUDGET + 2_000,
            "{} characters rendered against a budget of {TEXT_BUDGET}",
            rendering.len()
        );
        assert!(rendering.contains("set_max_text_chars"), "the cut is announced");
    }
}

#[test]
fn test_small_chart_text_is_kept_in_every_drawing() {
    limits();
    let doc =
        Document::from_reader(Cursor::new(chart_docx("Revenue", USERS)), DocumentFormat::Docx)
            .unwrap();
    let ir = doc.to_ir();
    let with_text = paragraphs(&ir)
        .iter()
        .filter(|p| {
            p.content
                .iter()
                .any(|c| matches!(c, InlineContent::Text(t) if t.text.contains("Revenue")))
        })
        .count();
    assert_eq!(with_text, USERS);
    assert!(!ir.metadata.text_truncated);
    assert_eq!(doc.plain_text().matches("Revenue").count(), USERS);
}
