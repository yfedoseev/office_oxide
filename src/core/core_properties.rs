//! Shared `docProps/core.xml` generator used by DOCX, PPTX, and XLSX
//! writers. Emits the OOXML core-properties payload from the IR's
//! `Metadata` so document title / author / subject / created /
//! modified surface in Word, PowerPoint, and Excel "Properties"
//! dialogs.

use crate::ir::Metadata;
use quick_xml::Writer;
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};

/// MIME content type for `docProps/core.xml`.
pub const CONTENT_TYPE: &str = "application/vnd.openxmlformats-package.core-properties+xml";

/// Generate the XML payload for `docProps/core.xml`. Empty fields
/// in the input are omitted entirely (no `<dc:title></dc:title>`),
/// matching the convention Word / PowerPoint use.
pub fn generate_xml(meta: &Metadata) -> Vec<u8> {
    let mut w = Writer::new_with_indent(Vec::new(), b' ', 2);
    w.write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
        .expect("decl");

    let mut root = BytesStart::new("cp:coreProperties");
    root.push_attribute((
        "xmlns:cp",
        "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    ));
    root.push_attribute(("xmlns:dc", "http://purl.org/dc/elements/1.1/"));
    root.push_attribute(("xmlns:dcterms", "http://purl.org/dc/terms/"));
    root.push_attribute(("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"));
    w.write_event(Event::Start(root)).expect("root");

    write_text(&mut w, "dc:title", meta.title.as_deref());
    write_text(&mut w, "dc:subject", meta.subject.as_deref());
    write_text(&mut w, "dc:creator", meta.author.as_deref());
    write_text(&mut w, "dc:description", meta.description.as_deref());
    if !meta.keywords.is_empty() {
        write_text(&mut w, "cp:keywords", Some(meta.keywords.join(", ").as_str()));
    }
    write_dcterms(&mut w, "dcterms:created", meta.created.as_deref());
    write_dcterms(&mut w, "dcterms:modified", meta.modified.as_deref());

    w.write_event(Event::End(BytesEnd::new("cp:coreProperties")))
        .expect("close");
    w.into_inner()
}

fn write_text(w: &mut Writer<Vec<u8>>, tag: &str, value: Option<&str>) {
    if let Some(v) = value {
        if v.is_empty() {
            return;
        }
        w.write_event(Event::Start(BytesStart::new(tag.to_string())))
            .expect("open");
        w.write_event(Event::Text(BytesText::new(v))).expect("text");
        w.write_event(Event::End(BytesEnd::new(tag.to_string())))
            .expect("close");
    }
}

fn write_dcterms(w: &mut Writer<Vec<u8>>, tag: &str, value: Option<&str>) {
    if let Some(v) = value {
        if v.is_empty() {
            return;
        }
        let mut elem = BytesStart::new(tag.to_string());
        elem.push_attribute(("xsi:type", "dcterms:W3CDTF"));
        w.write_event(Event::Start(elem)).expect("open");
        w.write_event(Event::Text(BytesText::new(v))).expect("text");
        w.write_event(Event::End(BytesEnd::new(tag.to_string())))
            .expect("close");
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::DocumentFormat;

    fn meta_string(meta: &Metadata) -> String {
        String::from_utf8(generate_xml(meta)).unwrap()
    }

    #[test]
    fn empty_metadata_emits_only_root() {
        let meta = Metadata {
            format: DocumentFormat::Docx,
            ..Default::default()
        };
        let xml = meta_string(&meta);
        assert!(xml.contains("<cp:coreProperties"), "xml: {xml}");
        assert!(!xml.contains("<dc:title"), "xml: {xml}");
        assert!(!xml.contains("<dc:creator"), "xml: {xml}");
        assert!(!xml.contains("<dcterms:created"), "xml: {xml}");
    }

    #[test]
    fn title_and_author_are_emitted() {
        let meta = Metadata {
            format: DocumentFormat::Docx,
            title: Some("Hello".into()),
            author: Some("Yury".into()),
            ..Default::default()
        };
        let xml = meta_string(&meta);
        assert!(xml.contains("<dc:title>Hello</dc:title>"), "xml: {xml}");
        assert!(xml.contains("<dc:creator>Yury</dc:creator>"), "xml: {xml}");
    }

    #[test]
    fn empty_string_field_is_omitted() {
        let meta = Metadata {
            format: DocumentFormat::Docx,
            title: Some(String::new()),
            author: Some("Someone".into()),
            ..Default::default()
        };
        let xml = meta_string(&meta);
        // Empty title is dropped entirely; non-empty author is kept.
        assert!(!xml.contains("<dc:title"), "xml: {xml}");
        assert!(xml.contains("<dc:creator>Someone</dc:creator>"), "xml: {xml}");
    }

    #[test]
    fn dcterms_carry_w3cdtf_type_attribute() {
        let meta = Metadata {
            format: DocumentFormat::Docx,
            created: Some("2026-05-13T10:00:00Z".into()),
            modified: Some("2026-05-13T11:00:00Z".into()),
            ..Default::default()
        };
        let xml = meta_string(&meta);
        assert!(xml.contains("xsi:type=\"dcterms:W3CDTF\""), "xml: {xml}");
        assert!(xml.contains("2026-05-13T10:00:00Z"), "xml: {xml}");
        assert!(xml.contains("2026-05-13T11:00:00Z"), "xml: {xml}");
    }

    #[test]
    fn keywords_joined_with_comma() {
        let meta = Metadata {
            format: DocumentFormat::Docx,
            keywords: vec!["rust".into(), "office".into(), "oxide".into()],
            ..Default::default()
        };
        let xml = meta_string(&meta);
        assert!(xml.contains("<cp:keywords>rust, office, oxide</cp:keywords>"), "xml: {xml}");
    }

    #[test]
    fn no_keywords_omits_element() {
        let meta = Metadata {
            format: DocumentFormat::Docx,
            ..Default::default()
        };
        let xml = meta_string(&meta);
        assert!(!xml.contains("<cp:keywords"), "xml: {xml}");
    }

    #[test]
    fn content_type_is_core_properties() {
        assert!(CONTENT_TYPE.ends_with("core-properties+xml"));
    }
}
