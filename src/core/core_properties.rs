//! Shared `docProps/core.xml` generator used by DOCX, PPTX, and XLSX
//! writers. Emits the OOXML core-properties payload from the IR's
//! `Metadata` so document title / author / subject / created /
//! modified surface in Word, PowerPoint, and Excel "Properties"
//! dialogs.

use crate::core::properties::normalize_w3cdtf;
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
    write_text(&mut w, "dc:language", meta.language.as_deref());
    if !meta.keywords.is_empty() {
        write_text(&mut w, "cp:keywords", Some(join_keywords(&meta.keywords).as_str()));
    }
    write_text(&mut w, "cp:category", meta.category.as_deref());
    write_text(&mut w, "cp:contentStatus", meta.content_status.as_deref());
    write_text(&mut w, "cp:lastModifiedBy", meta.last_modified_by.as_deref());
    write_text(&mut w, "cp:revision", meta.revision.as_deref());
    // A malformed source date is recovered where the intent is unambiguous
    // and dropped otherwise (both elements are optional), rather than
    // authoring a core.xml the W3CDTF type rejects.
    let created = meta.created.as_deref().and_then(normalize_w3cdtf);
    let modified = meta.modified.as_deref().and_then(normalize_w3cdtf);
    write_dcterms(&mut w, "dcterms:created", created.as_deref());
    write_dcterms(&mut w, "dcterms:modified", modified.as_deref());

    w.write_event(Event::End(BytesEnd::new("cp:coreProperties")))
        .expect("close");
    w.into_inner()
}

/// Join keywords into one `cp:keywords` value that
/// [`split_keywords`] splits back into the same list.
pub(crate) fn join_keywords(keywords: &[String]) -> String {
    keywords.join(", ")
}

/// Split a `cp:keywords` (or legacy `PIDSI_KEYWORDS`) value into keywords.
/// OOXML defines no separator; Word writes `;`- or `,`-separated lists, so
/// both separate. Whitespace does not: "annual report" is one keyword —
/// splitting on it turned every multi-word keyword into several, and the
/// written file then carried different keywords from the source.
pub(crate) fn split_keywords(s: &str) -> Vec<String> {
    s.split([',', ';'])
        .map(|k| k.trim().to_string())
        .filter(|k| !k.is_empty())
        .collect()
}

/// MIME content type for `docProps/app.xml`.
pub const APP_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.extended-properties+xml";
/// MIME content type for `docProps/custom.xml`.
pub const CUSTOM_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.custom-properties+xml";

/// The producer name written to `docProps/app.xml` `<Application>`.
pub const APPLICATION_NAME: &str = "office_oxide";

/// Generate `docProps/app.xml`: the producing application, plus the
/// company and manager carried by `meta`. Page/word/slide counts are not
/// written — they describe the source layout, which a written document no
/// longer has, and the Office applications recompute them on save.
pub fn generate_app_xml(meta: Option<&Metadata>) -> Vec<u8> {
    crate::core::properties::AppProperties {
        application: Some(APPLICATION_NAME.to_string()),
        company: meta
            .and_then(|m| m.company.clone())
            .filter(|c| !c.is_empty()),
        manager: meta
            .and_then(|m| m.manager.clone())
            .filter(|c| !c.is_empty()),
        ..Default::default()
    }
    .serialize()
}

/// Add the package-property parts to a package being written:
/// `docProps/core.xml` when `meta` is given, `docProps/app.xml` always
/// (every Office package names its producer), and `docProps/custom.xml`
/// when `meta` carries custom properties — each with its package
/// relationship and content type.
pub fn add_property_parts<W: std::io::Write + std::io::Seek>(
    opc: &mut crate::core::opc::OpcWriter<W>,
    meta: Option<&Metadata>,
) -> crate::core::Result<()> {
    use crate::core::opc::PartName;
    use crate::core::relationships::rel_types;
    if let Some(meta) = meta {
        opc.add_package_rel(rel_types::CORE_PROPERTIES, "docProps/core.xml");
        opc.add_part(&PartName::new("/docProps/core.xml")?, CONTENT_TYPE, &generate_xml(meta))?;
    }
    opc.add_package_rel(rel_types::EXTENDED_PROPERTIES, "docProps/app.xml");
    opc.add_part(&PartName::new("/docProps/app.xml")?, APP_CONTENT_TYPE, &generate_app_xml(meta))?;
    if let Some(meta) = meta.filter(|m| !m.custom_properties.is_empty()) {
        opc.add_package_rel(rel_types::CUSTOM_PROPERTIES, "docProps/custom.xml");
        opc.add_part(
            &PartName::new("/docProps/custom.xml")?,
            CUSTOM_CONTENT_TYPE,
            &crate::core::properties::serialize_custom_properties(&meta.custom_properties),
        )?;
    }
    Ok(())
}

/// The `Metadata` fields an OOXML converter does not fill itself: the
/// remaining core properties, the app.xml company/manager, and the
/// package-level custom properties, signature flag and thumbnail. Meant
/// for struct-update syntax: `Metadata { title, .., ..ooxml_metadata_extras(..) }`.
pub(crate) fn ooxml_metadata_extras(
    core: Option<&crate::core::properties::CoreProperties>,
    app: Option<&crate::core::properties::AppProperties>,
    package: Option<&crate::core::properties::PackageProperties>,
) -> Metadata {
    Metadata {
        last_modified_by: core.and_then(|c| c.last_modified_by.clone()),
        revision: core.and_then(|c| c.revision.clone()),
        category: core.and_then(|c| c.category.clone()),
        content_status: core.and_then(|c| c.content_status.clone()),
        language: core.and_then(|c| c.language.clone()),
        company: app.and_then(|a| a.company.clone()),
        manager: app.and_then(|a| a.manager.clone()),
        custom_properties: package.map(|p| p.custom.clone()).unwrap_or_default(),
        has_digital_signature: package.is_some_and(|p| p.has_digital_signature),
        thumbnail: package.and_then(|p| p.thumbnail.clone()),
        warnings: package_warnings(package),
        ..Default::default()
    }
}

/// The package-level integrity warnings every OOXML converter reports: one
/// line per part whose bytes failed their CRC-32 and were read as stored.
pub(crate) fn package_warnings(
    package: Option<&crate::core::properties::PackageProperties>,
) -> Vec<String> {
    package
        .map(|p| {
            p.crc_mismatched_parts
                .iter()
                .map(|name| {
                    format!(
                        "part '{name}' failed its CRC-32 check; its content was read as \
                         stored and may be damaged"
                    )
                })
                .collect()
        })
        .unwrap_or_default()
}

/// The `Metadata` fields a legacy converter does not fill itself, from
/// the compound file's SummaryInformation / DocumentSummaryInformation.
pub(crate) fn legacy_metadata_extras(summary: Option<&crate::cfb::SummaryProperties>) -> Metadata {
    let text = |f: fn(&crate::cfb::SummaryProperties) -> &Option<String>| {
        summary.and_then(|s| f(s).clone()).filter(|v| !v.is_empty())
    };
    Metadata {
        last_modified_by: text(|s| &s.last_author),
        revision: text(|s| &s.revision),
        category: text(|s| &s.category),
        company: text(|s| &s.company),
        manager: text(|s| &s.manager),
        has_digital_signature: summary.is_some_and(|s| s.has_digital_signature),
        ..Default::default()
    }
}

fn write_text(w: &mut Writer<Vec<u8>>, tag: &str, value: Option<&str>) {
    if let Some(v) = value {
        if v.is_empty() {
            return;
        }
        w.write_event(Event::Start(BytesStart::new(tag.to_string())))
            .expect("open");
        w.write_event(Event::Text(BytesText::new(&crate::core::xml::sanitize_xml_text(v))))
            .expect("text");
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
        w.write_event(Event::Text(BytesText::new(&crate::core::xml::sanitize_xml_text(v))))
            .expect("text");
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
    fn test_empty_metadata_emits_only_root() {
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
    fn test_title_and_author_are_emitted() {
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
    fn test_empty_string_field_is_omitted() {
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
    fn test_dcterms_carry_w3cdtf_type_attribute() {
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
    fn test_keywords_joined_with_comma() {
        let meta = Metadata {
            format: DocumentFormat::Docx,
            keywords: vec!["rust".into(), "office".into(), "oxide".into()],
            ..Default::default()
        };
        let xml = meta_string(&meta);
        assert!(xml.contains("<cp:keywords>rust, office, oxide</cp:keywords>"), "xml: {xml}");
    }

    #[test]
    fn test_no_keywords_omits_element() {
        let meta = Metadata {
            format: DocumentFormat::Docx,
            ..Default::default()
        };
        let xml = meta_string(&meta);
        assert!(!xml.contains("<cp:keywords"), "xml: {xml}");
    }

    #[test]
    fn test_content_type_is_core_properties() {
        assert!(CONTENT_TYPE.ends_with("core-properties+xml"));
    }
}
