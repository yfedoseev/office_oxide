//! The paragraph style a document applies, reached through the public API.
//!
//! A paragraph styled with one of Word's own styles (Caption, a legislation
//! style named `section-e`, …) used to come out as a plain paragraph, so a
//! consumer could not tell it from body text. `Paragraph::style` carries the
//! style's display name; `None` means the document's default paragraph style.
//!
//! The IR is inspected as JSON, which is also the contract the bindings
//! expose. Documents are built in code (no committed third-party fixture).

mod common;

use std::io::Cursor;

use common::{Para, Style, build_doc_with_styles, open_doc, style_sheet};
use office_oxide::core::opc::{OpcWriter, PartName};
use office_oxide::core::relationships::rel_types;
use office_oxide::{Document, DocumentFormat};
use serde_json::Value;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

fn docx(body: &str, styles: Option<&str>) -> Document {
    docx_with_numbering(body, styles, None)
}

fn docx_with_numbering(body: &str, styles: Option<&str>, numbering: Option<&str>) -> Document {
    let document_xml = format!(r#"<w:document xmlns:w="{W}"><w:body>{body}</w:body></w:document>"#);
    let mut writer = OpcWriter::new(Cursor::new(Vec::new())).unwrap();
    let doc_part = PartName::new("/word/document.xml").unwrap();
    writer
        .add_part(
            &doc_part,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            document_xml.as_bytes(),
        )
        .unwrap();
    writer.add_package_rel(rel_types::OFFICE_DOCUMENT, "word/document.xml");
    if let Some(styles) = styles {
        let styles_xml = format!(r#"<w:styles xmlns:w="{W}">{styles}</w:styles>"#);
        let styles_part = PartName::new("/word/styles.xml").unwrap();
        writer
            .add_part(
                &styles_part,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
                styles_xml.as_bytes(),
            )
            .unwrap();
        writer.add_part_rel(&doc_part, rel_types::STYLES, "styles.xml");
    }
    if let Some(numbering) = numbering {
        let part = PartName::new("/word/numbering.xml").unwrap();
        writer
            .add_part(
                &part,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
                numbering.as_bytes(),
            )
            .unwrap();
        writer.add_part_rel(&doc_part, rel_types::NUMBERING, "numbering.xml");
    }
    let bytes = writer.finish().unwrap().into_inner();
    Document::from_reader(Cursor::new(bytes), DocumentFormat::Docx).expect("open .docx")
}

/// `(text, style)` of every IR paragraph, in document order, wherever it sits
/// (body, table cell, …).
fn paragraphs(doc: &Document) -> Vec<(String, Option<String>)> {
    fn walk(v: &Value, out: &mut Vec<(String, Option<String>)>) {
        match v {
            Value::Object(m) => {
                if m.get("type").and_then(Value::as_str) == Some("paragraph") {
                    let text: String = m
                        .get("content")
                        .and_then(Value::as_array)
                        .into_iter()
                        .flatten()
                        .filter_map(|c| c.get("text").and_then(Value::as_str))
                        .collect();
                    let style = m.get("style").and_then(Value::as_str).map(str::to_owned);
                    out.push((text, style));
                }
                m.values().for_each(|c| walk(c, out));
            },
            Value::Array(a) => a.iter().for_each(|c| walk(c, out)),
            _ => {},
        }
    }
    let mut out = Vec::new();
    walk(&serde_json::to_value(doc.to_ir()).unwrap(), &mut out);
    out
}

fn style_of(paras: &[(String, Option<String>)], text: &str) -> Option<String> {
    paras
        .iter()
        .find(|(t, _)| t == text)
        .unwrap_or_else(|| panic!("paragraph {text:?} not found in {paras:?}"))
        .1
        .clone()
}

/// Word writes a built-in style's name in lower case (`caption`,
/// `heading 1`); the name consumers expect is the one Word shows (`Caption`).
const STYLES: &str = r#"
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Caption"><w:name w:val="caption"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="MyNote"><w:name w:val="My Note"/><w:basedOn w:val="Normal"/></w:style>
"#;

#[test]
fn test_docx_paragraph_style_is_exposed_by_display_name() {
    let doc = docx(
        r#"<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>
           <w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr><w:r><w:t>Table 1</w:t></w:r></w:p>
           <w:tbl><w:tr><w:tc><w:p><w:pPr><w:pStyle w:val="MyNote"/></w:pPr><w:r><w:t>in a cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
           <w:p><w:r><w:t>Plain body</w:t></w:r></w:p>"#,
        Some(STYLES),
    );
    let p = paragraphs(&doc);
    assert_eq!(style_of(&p, "Table 1").as_deref(), Some("Caption"));
    assert_eq!(style_of(&p, "in a cell").as_deref(), Some("My Note"));
    assert_eq!(style_of(&p, "Plain body"), None);
}

#[test]
fn test_docx_default_paragraph_style_is_none_however_it_is_named() {
    // An explicit reference to the default style is still "the default": it is
    // recognised by `w:default`, not by the name, which Word localizes.
    let styles = r#"<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Standard"/></w:style>"#;
    let doc = docx(
        r#"<w:p><w:pPr><w:pStyle w:val="a"/></w:pPr><w:r><w:t>Explicit</w:t></w:r></w:p>"#,
        Some(styles),
    );
    assert_eq!(style_of(&paragraphs(&doc), "Explicit"), None);
}

#[test]
fn test_docx_style_that_is_not_a_defined_paragraph_style_is_none() {
    // As python-docx's `get_by_id`: an id the sheet does not define, a missing
    // sheet, and an id naming a character or table style all fall back to the
    // default paragraph style, which is `None`. Inventing a style the document
    // does not define would be a plausible but wrong answer.
    let defs = r#"
      <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
      <w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/></w:style>
      <w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/></w:style>"#;
    for (id, sheet) in [
        ("Ghost", Some(defs)),
        ("Strong", Some(defs)),
        ("TableGrid", Some(defs)),
        ("Caption", None),
    ] {
        let doc = docx(
            &format!(
                r#"<w:p><w:pPr><w:pStyle w:val="{id}"/></w:pPr><w:r><w:t>Text</w:t></w:r></w:p>"#
            ),
            sheet,
        );
        assert_eq!(style_of(&paragraphs(&doc), "Text"), None, "pStyle={id}");
    }
}

#[test]
fn test_docx_without_a_flagged_default_style_normal_is_the_default() {
    // Some producers define `Normal` but never set `w:default`.
    let styles = r#"<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
                    <w:style w:type="paragraph" w:styleId="Caption"><w:name w:val="caption"/></w:style>"#;
    let doc = docx(
        r#"<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:t>Explicit</w:t></w:r></w:p>
           <w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr><w:r><w:t>Cap</w:t></w:r></w:p>"#,
        Some(styles),
    );
    let p = paragraphs(&doc);
    assert_eq!(style_of(&p, "Explicit"), None);
    assert_eq!(style_of(&p, "Cap").as_deref(), Some("Caption"));
}

#[test]
fn test_docx_style_names_that_are_not_names_are_not_exposed() {
    let long = "A".repeat(300);
    let exact = "B".repeat(255);
    let styles = format!(
        r#"<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
           <w:style w:type="paragraph" w:styleId="Long"><w:name w:val="{long}"/></w:style>
           <w:style w:type="paragraph" w:styleId="Exact"><w:name w:val="{exact}"/></w:style>
           <w:style w:type="paragraph" w:styleId="Blank"><w:name w:val="   "/></w:style>
           <w:style w:type="paragraph" w:styleId="Ctl"><w:name w:val="a&#10;b"/></w:style>
           <w:style w:type="paragraph" w:styleId="Padded"><w:name w:val="  Padded  "/></w:style>"#
    );
    let body = |id: &str, t: &str| {
        format!(r#"<w:p><w:pPr><w:pStyle w:val="{id}"/></w:pPr><w:r><w:t>{t}</w:t></w:r></w:p>"#)
    };
    let doc = docx(
        &[
            body("Long", "l"),
            body("Exact", "e"),
            body("Blank", "b"),
            body("Ctl", "c"),
            body("Padded", "p"),
        ]
        .concat(),
        Some(&styles),
    );
    let p = paragraphs(&doc);
    assert_eq!(style_of(&p, "l"), None, "over 255 characters");
    assert_eq!(style_of(&p, "e").as_deref(), Some(exact.as_str()), "exactly 255 is allowed");
    assert_eq!(style_of(&p, "b"), None, "whitespace only");
    assert_eq!(style_of(&p, "c"), None, "control character");
    assert_eq!(style_of(&p, "p").as_deref(), Some("Padded"), "trimmed");
    let warnings = doc.to_ir().metadata.warnings;
    assert!(
        warnings.iter().any(|w| w.contains("style name")),
        "rejected names must be reported, got {warnings:?}"
    );
}

#[test]
fn test_docx_style_name_cannot_multiply_a_small_file() {
    // One huge style name referenced from every paragraph used to be copied
    // into each of them: output = name length x paragraph count.
    let name = "A".repeat(100_000);
    let styles = format!(
        r#"<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
           <w:style w:type="paragraph" w:styleId="S"><w:name w:val="{name}"/></w:style>"#
    );
    let body =
        r#"<w:p><w:pPr><w:pStyle w:val="S"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>"#.repeat(300);
    let doc = docx(&body, Some(&styles));
    let json = serde_json::to_string(&doc.to_ir()).unwrap();
    assert!(json.len() < 200_000, "300 paragraphs produced {} bytes of JSON", json.len());
}

#[test]
fn test_docx_malformed_default_attribute_does_not_stop_the_document_opening() {
    let styles = r#"<w:style w:type="paragraph" w:default="&bogus;" w:styleId="Caption"><w:name w:val="caption"/></w:style>"#;
    let doc = docx(
        r#"<w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr><w:r><w:t>Cap</w:t></w:r></w:p>"#,
        Some(styles),
    );
    assert_eq!(style_of(&paragraphs(&doc), "Cap").as_deref(), Some("Caption"));
}

#[test]
fn test_paragraph_style_is_omitted_from_json_when_absent() {
    // Other IR types use a `style` key of their own (borders, lists), so look
    // at paragraph objects only.
    fn paragraph_style_keys(v: &Value, out: &mut Vec<Option<String>>) {
        match v {
            Value::Object(m) => {
                if m.get("type").and_then(Value::as_str) == Some("paragraph") {
                    out.push(m.get("style").map(|s| s.to_string()));
                }
                m.values().for_each(|c| paragraph_style_keys(c, out));
            },
            Value::Array(a) => a.iter().for_each(|c| paragraph_style_keys(c, out)),
            _ => {},
        }
    }
    let keys = |doc: &Document| {
        let mut out = Vec::new();
        paragraph_style_keys(&serde_json::to_value(doc.to_ir()).unwrap(), &mut out);
        out
    };
    let plain = docx(r#"<w:p><w:r><w:t>Plain</w:t></w:r></w:p>"#, Some(STYLES));
    assert_eq!(keys(&plain), vec![None], "an absent style must not appear");

    let cap = docx(
        r#"<w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr><w:r><w:t>Cap</w:t></w:r></w:p>"#,
        Some(STYLES),
    );
    assert_eq!(keys(&cap), vec![Some("\"Caption\"".to_owned())]);
}

#[test]
fn test_docx_heading_style_stays_a_heading_element() {
    let doc = docx(
        r#"<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>"#,
        Some(STYLES),
    );
    let ir = doc.to_ir();
    let kinds: Vec<_> = ir.sections.iter().flat_map(|s| &s.elements).collect();
    assert!(matches!(kinds.as_slice(), [office_oxide::ir::Element::Heading(h)] if h.level == 1));
    assert!(paragraphs(&doc).is_empty(), "a heading is not also a styled paragraph");
}

fn para(text: &'static str) -> Para {
    Para {
        text,
        terminator: '\r',
        grpprl: Vec::new(),
    }
}

/// A `.doc` style sheet whose `istd`s are: 0 Normal, 1 `section-e`,
/// 2 `Caption,Legende` (primary name and an alias), 3 an unnamed slot,
/// 4 built-in Header (sti 31) and 5 built-in Caption (sti 34), both stored
/// under localized names.
fn doc_with_styles(paras: &[Para], istds: &[u16]) -> Document {
    let stsh = style_sheet(&[
        Style::new(0, "Normal"),
        Style::new(0x0FFE, "section-e"),
        Style::new(0x0FFE, "Caption,Legende"),
        Style::new(0x0FFE, ""),
        Style::new(31, "Kopfzeile"),
        Style::new(34, "Beschriftung"),
    ]);
    open_doc(&build_doc_with_styles(paras, istds, &stsh))
}

#[test]
fn test_doc_paragraph_style_is_the_stored_name_and_normal_is_none() {
    let paras = [
        para("The default style paragraph is ordinary body text."),
        para("A legislation section paragraph states the rule."),
        para("A caption paragraph sits above its table."),
        para("A paragraph in an unnamed style slot has no name."),
    ];
    let doc = doc_with_styles(&paras, &[0, 1, 2, 3]);
    let p = paragraphs(&doc);
    assert_eq!(style_of(&p, "The default style paragraph is ordinary body text."), None);
    assert_eq!(
        style_of(&p, "A legislation section paragraph states the rule.").as_deref(),
        Some("section-e")
    );
    assert_eq!(
        style_of(&p, "A caption paragraph sits above its table.").as_deref(),
        Some("Caption"),
        "an alias after the comma is not part of the name"
    );
    assert_eq!(style_of(&p, "A paragraph in an unnamed style slot has no name."), None);
}

#[test]
fn test_doc_paragraph_style_reaches_a_paragraph_in_a_table_cell() {
    let cell = |text: &'static str| Para {
        text,
        terminator: '\u{7}',
        grpprl: common::cell_grpprl(),
    };
    let row_end = Para {
        text: "",
        terminator: '\u{7}',
        grpprl: common::row_grpprl(&[0, 1000, 2000], &[0, 0]),
    };
    let paras = [cell("styled cell"), cell("plain cell"), row_end];
    let doc = doc_with_styles(&paras, &[1, 0, 0]);
    let p = paragraphs(&doc);
    assert_eq!(style_of(&p, "styled cell").as_deref(), Some("section-e"));
    assert_eq!(style_of(&p, "plain cell"), None);
}

#[test]
fn test_doc_built_in_styles_get_their_canonical_name_whatever_the_language() {
    // `sti` identifies the built-in; the stored name is whatever language the
    // producing Word ran in. DOCX stores the English names, so a consumer
    // matching on "Caption" or "Header" must get the same answer from both.
    let paras = [
        para("A header style paragraph stored under a German name."),
        para("A caption style paragraph stored under a German name."),
    ];
    let doc = doc_with_styles(&paras, &[4, 5]);
    let p = paragraphs(&doc);
    assert_eq!(
        style_of(&p, "A header style paragraph stored under a German name.").as_deref(),
        Some("Header")
    );
    assert_eq!(
        style_of(&p, "A caption style paragraph stored under a German name.").as_deref(),
        Some("Caption")
    );
}

#[test]
fn test_doc_style_name_over_the_limit_is_not_exposed_and_is_reported() {
    let long = "N".repeat(300);
    let stsh = style_sheet(&[Style::new(0, "Normal"), Style::new(0x0FFE, &long)]);
    let paras = [para("A paragraph in a style with an absurdly long name.")];
    let doc = open_doc(&build_doc_with_styles(&paras, &[1], &stsh));
    let p = paragraphs(&doc);
    assert_eq!(style_of(&p, "A paragraph in a style with an absurdly long name."), None);
    assert!(
        doc.to_ir()
            .metadata
            .warnings
            .iter()
            .any(|w| w.contains("style name"))
    );
}

#[test]
fn test_doc_style_name_with_a_trailing_nul_terminator_is_the_name_without_it() {
    // Some producers count the string's NUL terminator in `cch`, so the stored
    // name is `body\0`. That is the name "body", not a name with a control
    // character in it.
    let stsh = style_sheet(&[Style::new(0, "Normal"), Style::new(0x0FFE, "body\u{0}")]);
    let paras = [para("A paragraph in a style whose name ends in a NUL.")];
    let doc = open_doc(&build_doc_with_styles(&paras, &[1], &stsh));
    assert_eq!(
        style_of(&paragraphs(&doc), "A paragraph in a style whose name ends in a NUL.").as_deref(),
        Some("body")
    );
    assert!(
        !doc.to_ir()
            .metadata
            .warnings
            .iter()
            .any(|w| w.contains("style name")),
        "a terminator is not an unusable name"
    );
}

const NUMBERING: &str = r#"
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="-"/></w:lvl>
    <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="-"/></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>"#;

#[test]
fn test_docx_list_item_paragraphs_keep_their_style() {
    let styles = format!(
        r#"{STYLES}
        <w:style w:type="paragraph" w:styleId="ListBullet"><w:name w:val="List Bullet"/></w:style>
        <w:style w:type="paragraph" w:styleId="ListBullet2"><w:name w:val="List Bullet 2"/></w:style>"#
    );
    let item = |style: Option<&str>, ilvl: u8, text: &str| {
        let ps = style.map_or(String::new(), |s| format!(r#"<w:pStyle w:val="{s}"/>"#));
        format!(
            r#"<w:p><w:pPr>{ps}<w:numPr><w:ilvl w:val="{ilvl}"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>"#
        )
    };
    let body = [
        item(Some("ListBullet"), 0, "first"),
        item(Some("ListBullet2"), 1, "nested"),
        item(None, 0, "third"),
        r#"<w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr><w:r><w:t>after the list</w:t></w:r></w:p>"#.to_owned(),
    ]
    .concat();
    let doc = docx_with_numbering(
        &body,
        Some(&styles),
        Some(&format!(r#"<w:numbering xmlns:w="{W}">{NUMBERING}</w:numbering>"#)),
    );
    let p = paragraphs(&doc);
    assert_eq!(style_of(&p, "first").as_deref(), Some("List Bullet"));
    assert_eq!(style_of(&p, "nested").as_deref(), Some("List Bullet 2"));
    assert_eq!(style_of(&p, "third"), None);
    assert_eq!(style_of(&p, "after the list").as_deref(), Some("Caption"));
}

#[test]
fn test_doc_list_item_paragraphs_keep_their_style() {
    let item = |text: &'static str, ilvl: u8| Para {
        text,
        terminator: '\r',
        grpprl: common::list_grpprl(ilvl),
    };
    let paras = [
        item("styled top item", 0),
        item("styled nested item", 1),
        item("plain top item", 0),
    ];
    // istd 1 = `section-e`, 2 = `Caption,Legende`, 0 = Normal.
    let doc = doc_with_styles(&paras, &[1, 2, 0]);
    let p = paragraphs(&doc);
    assert_eq!(style_of(&p, "styled top item").as_deref(), Some("section-e"));
    assert_eq!(style_of(&p, "styled nested item").as_deref(), Some("Caption"));
    assert_eq!(style_of(&p, "plain top item"), None);
}
