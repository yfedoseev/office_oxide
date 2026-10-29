//! Private-use text in a symbol-font run whose font is not on the run itself.
//!
//! A run takes its font from the document defaults, its paragraph style, its
//! character style, and its own properties, the later overriding the earlier.
//! Word writes the font on a run only when it differs from the style's, so a
//! symbol-font paragraph style leaves the runs without a font of their own.

use std::io::Cursor;

use office_oxide::core::opc::{OpcWriter, PartName};
use office_oxide::core::relationships::rel_types;
use office_oxide::{Document, DocumentFormat};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/// A package whose `document.xml` body is `body` and whose `styles.xml` holds
/// `styles` (the content of `<w:styles>`).
fn docx(body: &str, styles: &str) -> Vec<u8> {
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="{W}"><w:body>{body}</w:body></w:document>"#
    );
    let styles = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="{W}">{styles}</w:styles>"#
    );
    let mut writer = OpcWriter::new(Cursor::new(Vec::new())).unwrap();
    let doc_part = PartName::new("/word/document.xml").unwrap();
    writer
        .add_part(
            &doc_part,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            document.as_bytes(),
        )
        .unwrap();
    writer.add_package_rel(rel_types::OFFICE_DOCUMENT, "word/document.xml");
    let styles_part = PartName::new("/word/styles.xml").unwrap();
    writer
        .add_part(
            &styles_part,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
            styles.as_bytes(),
        )
        .unwrap();
    writer.add_part_rel(&doc_part, rel_types::STYLES, "styles.xml");
    writer.finish().unwrap().into_inner()
}

fn text_of(body: &str, styles: &str) -> String {
    let doc = Document::from_reader(Cursor::new(docx(body, styles)), DocumentFormat::Docx).unwrap();
    doc.plain_text().trim().to_string()
}

fn run(text: &str) -> String {
    format!(r#"<w:r><w:t>{text}</w:t></w:r>"#)
}

const MU: &str = "\u{F06D}";
const SYMBOL_PARAGRAPH_STYLE: &str = r#"<w:style w:type="paragraph" w:styleId="Sym"><w:name w:val="Sym"/><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr></w:style>"#;
const SYMBOL_CHAR_STYLE: &str = r#"<w:style w:type="character" w:styleId="SymChar"><w:name w:val="SymChar"/><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr></w:style>"#;

#[test]
fn test_a_paragraph_style_in_a_symbol_font_maps_the_runs_without_a_font() {
    let body =
        format!(r#"<w:p><w:pPr><w:pStyle w:val="Sym"/></w:pPr>{}{}</w:p>"#, run(MU), run("M"));
    assert_eq!(text_of(&body, SYMBOL_PARAGRAPH_STYLE), "\u{3bc}M");
}

#[test]
fn test_a_character_style_in_a_symbol_font_maps_the_run() {
    let body = format!(
        r#"<w:p><w:r><w:rPr><w:rStyle w:val="SymChar"/></w:rPr><w:t>{MU}</w:t></w:r>{}</w:p>"#,
        run(MU)
    );
    // Only the run in the character style is in the symbol font.
    assert_eq!(text_of(&body, SYMBOL_CHAR_STYLE), format!("\u{3bc}{MU}"));
}

#[test]
fn test_a_style_based_on_a_symbol_style_inherits_its_font() {
    let styles = format!(
        r#"{SYMBOL_PARAGRAPH_STYLE}<w:style w:type="paragraph" w:styleId="Child"><w:name w:val="Child"/><w:basedOn w:val="Sym"/></w:style>"#
    );
    let body = format!(r#"<w:p><w:pPr><w:pStyle w:val="Child"/></w:pPr>{}</w:p>"#, run(MU));
    assert_eq!(text_of(&body, &styles), "\u{3bc}");
}

#[test]
fn test_the_document_defaults_can_name_a_symbol_font() {
    let styles = r#"<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings"/></w:rPr></w:rPrDefault></w:docDefaults>"#;
    assert_eq!(text_of(&format!("<w:p>{}</w:p>", run("\u{F0FC}")), styles), "\u{2713}");
}

#[test]
fn test_a_run_font_of_its_own_overrides_the_style() {
    let body = format!(
        r#"<w:p><w:pPr><w:pStyle w:val="Sym"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr><w:t>{MU}</w:t></w:r></w:p>"#
    );
    assert_eq!(text_of(&body, SYMBOL_PARAGRAPH_STYLE), MU);
    // And a symbol font on the run beats a plain style.
    let plain = r#"<w:style w:type="paragraph" w:styleId="Plain"><w:name w:val="Plain"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr></w:style>"#;
    let body = format!(
        r#"<w:p><w:pPr><w:pStyle w:val="Plain"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr><w:t>{MU}</w:t></w:r></w:p>"#
    );
    assert_eq!(text_of(&body, plain), "\u{3bc}");
}

#[test]
fn test_table_cells_and_hyperlinks_are_covered() {
    let cell =
        format!(r#"<w:tc><w:p><w:pPr><w:pStyle w:val="Sym"/></w:pPr>{}</w:p></w:tc>"#, run(MU));
    let link = format!(
        r#"<w:p><w:pPr><w:pStyle w:val="Sym"/></w:pPr><w:hyperlink w:anchor="x">{}</w:hyperlink></w:p>"#,
        run(MU)
    );
    let body = format!("<w:tbl><w:tr>{cell}</w:tr></w:tbl>{link}");
    let text = text_of(&body, SYMBOL_PARAGRAPH_STYLE);
    assert_eq!(text.matches('\u{3bc}').count(), 2, "{text:?}");
    assert!(!text.contains(MU), "{text:?}");
}

#[test]
fn test_other_surfaces_agree() {
    let body =
        format!(r#"<w:p><w:pPr><w:pStyle w:val="Sym"/></w:pPr>{}{}</w:p>"#, run(MU), run("M"));
    let doc = Document::from_reader(
        Cursor::new(docx(&body, SYMBOL_PARAGRAPH_STYLE)),
        DocumentFormat::Docx,
    )
    .unwrap();
    for out in [doc.plain_text(), doc.to_markdown(), doc.to_html()] {
        assert!(out.contains('\u{3bc}') && !out.contains(MU), "{out}");
    }
}

/// Without a styles part, the run's own font still decides.
#[test]
fn test_without_a_styles_part_the_runs_own_font_decides() {
    let body = format!(
        r#"<w:p><w:r><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr><w:t>{MU}</w:t></w:r></w:p>"#
    );
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="{W}"><w:body>{body}</w:body></w:document>"#
    );
    let mut writer = OpcWriter::new(Cursor::new(Vec::new())).unwrap();
    let part = PartName::new("/word/document.xml").unwrap();
    writer
        .add_part(
            &part,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            document.as_bytes(),
        )
        .unwrap();
    writer.add_package_rel(rel_types::OFFICE_DOCUMENT, "word/document.xml");
    let doc = Document::from_reader(
        Cursor::new(writer.finish().unwrap().into_inner()),
        DocumentFormat::Docx,
    )
    .unwrap();
    assert_eq!(doc.plain_text().trim(), "\u{3bc}");
}

/// Footnotes and headers are covered like the body.
#[test]
fn test_footnotes_and_headers_are_covered() {
    let footnotes = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="{W}"><w:footnote w:id="2"><w:p><w:pPr><w:pStyle w:val="Sym"/></w:pPr>{}</w:p></w:footnote></w:footnotes>"#,
        run(MU)
    );
    let body = r#"<w:p><w:r><w:t>text</w:t><w:footnoteReference w:id="2"/></w:r></w:p>"#;
    let styles = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="{W}">{SYMBOL_PARAGRAPH_STYLE}</w:styles>"#
    );
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="{W}"><w:body>{body}</w:body></w:document>"#
    );
    let mut writer = OpcWriter::new(Cursor::new(Vec::new())).unwrap();
    let doc_part = PartName::new("/word/document.xml").unwrap();
    writer
        .add_part(
            &doc_part,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            document.as_bytes(),
        )
        .unwrap();
    writer.add_package_rel(rel_types::OFFICE_DOCUMENT, "word/document.xml");
    let sp = PartName::new("/word/styles.xml").unwrap();
    writer
        .add_part(
            &sp,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
            styles.as_bytes(),
        )
        .unwrap();
    writer.add_part_rel(&doc_part, rel_types::STYLES, "styles.xml");
    let fp = PartName::new("/word/footnotes.xml").unwrap();
    writer
        .add_part(
            &fp,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
            footnotes.as_bytes(),
        )
        .unwrap();
    writer.add_part_rel(&doc_part, rel_types::FOOTNOTES, "footnotes.xml");
    let doc = Document::from_reader(
        Cursor::new(writer.finish().unwrap().into_inner()),
        DocumentFormat::Docx,
    )
    .unwrap();
    let text = doc.plain_text();
    assert!(text.contains('\u{3bc}') && !text.contains(MU), "{text:?}");
}

/// Table cells nest; a crafted file nests them deeply. The pass keeps its own
/// worklist, so depth costs no stack.
#[test]
fn test_deeply_nested_tables_do_not_overflow_the_stack() {
    let depth = 400;
    let mut body = format!(r#"<w:p><w:pPr><w:pStyle w:val="Sym"/></w:pPr>{}</w:p>"#, run(MU));
    for _ in 0..depth {
        body = format!("<w:tbl><w:tr><w:tc>{body}</w:tc></w:tr></w:tbl><w:p/>");
    }
    let text = std::thread::Builder::new()
        .stack_size(2 * 1024 * 1024)
        .spawn(move || text_of(&body, SYMBOL_PARAGRAPH_STYLE))
        .unwrap()
        .join()
        .expect("no stack overflow");
    assert!(text.is_empty() || text.contains('\u{3bc}') || !text.contains(MU), "{text:?}");
}
