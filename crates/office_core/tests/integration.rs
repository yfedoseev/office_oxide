use std::io::Cursor;

use office_core::content_types::{ContentTypes, ContentTypesBuilder};
use office_core::opc::{OpcReader, OpcWriter, PartName};
use office_core::properties::{AppProperties, CoreProperties};
use office_core::relationships::{rel_types, Relationships, RelationshipsBuilder, TargetMode};
use office_core::theme::{ColorRef, RgbColor, Theme, ThemeColorSlot};
use office_core::units::{Angle60k, Emu, HalfPoint, Percentage1000, Twip};

// ---------------------------------------------------------------------------
// Round-trip test: create a package, read it back, verify everything
// ---------------------------------------------------------------------------

#[test]
fn full_round_trip() {
    let buf = Vec::new();
    let cursor = Cursor::new(buf);
    let mut writer = OpcWriter::new(cursor).unwrap();

    // Create main document part
    let doc_name = PartName::new("/word/document.xml").unwrap();
    let doc_content = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>Hello, office_oxide!</w:t></w:r></w:p></w:body>
</w:document>"#;
    writer
        .add_part(
            &doc_name,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            doc_content,
        )
        .unwrap();
    writer.add_package_rel(rel_types::OFFICE_DOCUMENT, "word/document.xml");

    // Create styles part
    let styles_name = PartName::new("/word/styles.xml").unwrap();
    writer
        .add_part(
            &styles_name,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
            b"<styles/>",
        )
        .unwrap();
    writer.add_part_rel(&doc_name, rel_types::STYLES, "styles.xml");

    // Create theme part
    let theme_name = PartName::new("/word/theme/theme1.xml").unwrap();
    let theme_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Test Theme">
  <a:themeElements>
    <a:clrScheme name="Test">
      <a:dk1><a:srgbClr val="000000"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="333333"/></a:dk2>
      <a:lt2><a:srgbClr val="CCCCCC"/></a:lt2>
      <a:accent1><a:srgbClr val="FF0000"/></a:accent1>
      <a:accent2><a:srgbClr val="00FF00"/></a:accent2>
      <a:accent3><a:srgbClr val="0000FF"/></a:accent3>
      <a:accent4><a:srgbClr val="FFFF00"/></a:accent4>
      <a:accent5><a:srgbClr val="FF00FF"/></a:accent5>
      <a:accent6><a:srgbClr val="00FFFF"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Test">
      <a:majorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="Times New Roman"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Test"/>
  </a:themeElements>
</a:theme>"#;
    writer
        .add_part(
            &theme_name,
            "application/vnd.openxmlformats-officedocument.theme+xml",
            theme_xml,
        )
        .unwrap();
    writer.add_part_rel(&doc_name, rel_types::THEME, "theme/theme1.xml");

    // Set core properties
    let core_props = CoreProperties {
        title: Some("Integration Test Document".to_string()),
        creator: Some("office_oxide".to_string()),
        created: Some("2024-01-01T00:00:00Z".to_string()),
        modified: Some("2024-06-15T12:00:00Z".to_string()),
        ..Default::default()
    };
    writer.set_core_properties(&core_props).unwrap();

    // Set app properties
    let app_props = AppProperties {
        application: Some("office_oxide".to_string()),
        app_version: Some("0.1.0".to_string()),
        pages: Some(1),
        words: Some(3),
        ..Default::default()
    };
    writer.set_app_properties(&app_props).unwrap();

    // Finalize
    let result = writer.finish().unwrap();
    let data = result.into_inner();
    assert!(!data.is_empty());

    // --- Read it back ---

    let cursor = Cursor::new(data);
    let mut reader = OpcReader::new(cursor).unwrap();

    // Verify main document part discovery
    let main_part = reader.main_document_part().unwrap();
    assert_eq!(main_part, doc_name);

    // Verify content types resolution
    let ct = reader.content_types();
    assert_eq!(
        ct.resolve(&doc_name),
        Some("application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")
    );
    assert_eq!(
        ct.resolve(&styles_name),
        Some("application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml")
    );

    // Verify part content
    let content = reader.read_part(&doc_name).unwrap();
    assert_eq!(content, doc_content.as_slice());

    // Verify part-level relationships
    let doc_rels = reader.read_rels_for(&doc_name).unwrap();
    let style_rels = doc_rels.get_by_type(rel_types::STYLES);
    assert_eq!(style_rels.len(), 1);
    assert_eq!(style_rels[0].target, "styles.xml");

    let theme_rels = doc_rels.get_by_type(rel_types::THEME);
    assert_eq!(theme_rels.len(), 1);

    // Verify theme parsing
    let theme_data = reader.read_part(&theme_name).unwrap();
    let theme = Theme::parse(&theme_data).unwrap();
    assert_eq!(theme.name, "Test Theme");
    assert_eq!(
        theme.resolve_color(ThemeColorSlot::Accent1),
        Some(&RgbColor([255, 0, 0]))
    );
    assert_eq!(theme.font_scheme.major_latin, "Arial");
    assert_eq!(theme.font_scheme.minor_latin, "Times New Roman");

    // Verify core properties
    let cp_rel = reader
        .package_rels()
        .first_by_type(rel_types::CORE_PROPERTIES)
        .unwrap();
    let cp_name = PartName::new(&format!("/{}", cp_rel.target)).unwrap();
    let cp_data = reader.read_part(&cp_name).unwrap();
    let parsed_core = CoreProperties::parse(&cp_data).unwrap();
    assert_eq!(
        parsed_core.title.as_deref(),
        Some("Integration Test Document")
    );
    assert_eq!(parsed_core.creator.as_deref(), Some("office_oxide"));

    // Verify app properties
    let ap_rel = reader
        .package_rels()
        .first_by_type(rel_types::EXTENDED_PROPERTIES)
        .unwrap();
    let ap_name = PartName::new(&format!("/{}", ap_rel.target)).unwrap();
    let ap_data = reader.read_part(&ap_name).unwrap();
    let parsed_app = AppProperties::parse(&ap_data).unwrap();
    assert_eq!(parsed_app.application.as_deref(), Some("office_oxide"));
    assert_eq!(parsed_app.pages, Some(1));
    assert_eq!(parsed_app.words, Some(3));

    // Verify has_part
    assert!(reader.has_part(&doc_name));
    assert!(!reader.has_part(&PartName::new("/nonexistent.xml").unwrap()));
}

// ---------------------------------------------------------------------------
// Edge cases
// ---------------------------------------------------------------------------

#[test]
fn part_name_case_insensitive_content_type_lookup() {
    let mut builder = ContentTypesBuilder::new();
    builder.add_override(
        PartName::new("/Word/Document.xml").unwrap(),
        "application/vnd.test",
    );
    let ct = builder.build();

    // Lookup with different case should still find it
    let pn = PartName::new("/word/document.xml").unwrap();
    assert_eq!(ct.resolve(&pn), Some("application/vnd.test"));
}

#[test]
fn content_types_default_extension_case_insensitive() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="PNG" ContentType="image/png"/>
</Types>"#;
    let ct = ContentTypes::parse(xml).unwrap();

    // Extension stored lowercase, so lookup should match
    let pn = PartName::new("/media/image1.png").unwrap();
    assert_eq!(ct.resolve(&pn), Some("image/png"));
}

#[test]
fn relationships_external_target_mode() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
    Target="https://example.com" TargetMode="External"/>
</Relationships>"#;
    let rels = Relationships::parse(xml).unwrap();
    let r = rels.get_by_id("rId1").unwrap();
    assert_eq!(r.target_mode, TargetMode::External);
    assert_eq!(r.target, "https://example.com");
}

#[test]
fn relationships_empty() {
    let rels = Relationships::empty();
    assert!(rels.all().is_empty());
    assert!(rels.get_by_id("rId1").is_none());
    assert!(rels.get_by_type("anything").is_empty());
}

#[test]
fn missing_part_returns_error() {
    let buf = Vec::new();
    let cursor = Cursor::new(buf);
    let mut writer = OpcWriter::new(cursor).unwrap();
    let doc_name = PartName::new("/word/document.xml").unwrap();
    writer
        .add_part(&doc_name, "application/xml", b"<doc/>")
        .unwrap();
    writer.add_package_rel(rel_types::OFFICE_DOCUMENT, "word/document.xml");
    let result = writer.finish().unwrap();
    let data = result.into_inner();

    let cursor = Cursor::new(data);
    let mut reader = OpcReader::new(cursor).unwrap();

    let missing = PartName::new("/nonexistent/file.xml").unwrap();
    assert!(reader.read_part(&missing).is_err());
}

#[test]
fn missing_rels_returns_empty() {
    let buf = Vec::new();
    let cursor = Cursor::new(buf);
    let mut writer = OpcWriter::new(cursor).unwrap();
    let doc_name = PartName::new("/word/document.xml").unwrap();
    writer
        .add_part(&doc_name, "application/xml", b"<doc/>")
        .unwrap();
    writer.add_package_rel(rel_types::OFFICE_DOCUMENT, "word/document.xml");
    let result = writer.finish().unwrap();
    let data = result.into_inner();

    let cursor = Cursor::new(data);
    let mut reader = OpcReader::new(cursor).unwrap();

    // Part has no relationships file -> should return empty
    let styles = PartName::new("/word/styles.xml").unwrap();
    let rels = reader.read_rels_for(&styles).unwrap();
    assert!(rels.all().is_empty());
}

#[test]
fn relative_uri_resolution_with_dotdot() {
    let source = PartName::new("/word/document.xml").unwrap();

    // Simple relative
    let r = source.resolve_relative("styles.xml").unwrap();
    assert_eq!(r.as_str(), "/word/styles.xml");

    // Parent directory
    let r = source.resolve_relative("../docProps/core.xml").unwrap();
    assert_eq!(r.as_str(), "/docProps/core.xml");

    // Subdirectory
    let r = source.resolve_relative("media/image1.png").unwrap();
    assert_eq!(r.as_str(), "/word/media/image1.png");

    // Absolute (ignores source)
    let r = source.resolve_relative("/xl/workbook.xml").unwrap();
    assert_eq!(r.as_str(), "/xl/workbook.xml");
}

// ---------------------------------------------------------------------------
// Theme color resolution
// ---------------------------------------------------------------------------

#[test]
fn theme_color_ref_with_shade() {
    let theme_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T">
  <a:themeElements>
    <a:clrScheme name="T">
      <a:dk1><a:srgbClr val="000000"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="333333"/></a:dk2>
      <a:lt2><a:srgbClr val="CCCCCC"/></a:lt2>
      <a:accent1><a:srgbClr val="FF0000"/></a:accent1>
      <a:accent2><a:srgbClr val="00FF00"/></a:accent2>
      <a:accent3><a:srgbClr val="0000FF"/></a:accent3>
      <a:accent4><a:srgbClr val="FFFF00"/></a:accent4>
      <a:accent5><a:srgbClr val="FF00FF"/></a:accent5>
      <a:accent6><a:srgbClr val="00FFFF"/></a:accent6>
      <a:hlink><a:srgbClr val="0000FF"/></a:hlink>
      <a:folHlink><a:srgbClr val="800080"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="T">
      <a:majorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="T"/>
  </a:themeElements>
</a:theme>"#;

    let theme = Theme::parse(theme_xml).unwrap();

    // Accent1 = pure red (255,0,0), with 50% shade = (128,0,0)
    let color = ColorRef::Theme {
        slot: ThemeColorSlot::Accent1,
        tint: None,
        shade: Some(0.5),
    };
    let resolved = color.resolve(&theme);
    assert_eq!(resolved, RgbColor([128, 0, 0]));

    // White (255,255,255) with 50% tint = still white (tint lightens)
    let color = ColorRef::Theme {
        slot: ThemeColorSlot::Lt1,
        tint: Some(0.5),
        shade: None,
    };
    let resolved = color.resolve(&theme);
    assert_eq!(resolved, RgbColor([255, 255, 255]));
}

// ---------------------------------------------------------------------------
// Unit conversion invariants
// ---------------------------------------------------------------------------

#[test]
fn unit_conversion_invariants() {
    // US Letter dimensions in twips
    let letter_width = Twip(12240);
    let letter_height = Twip(15840);
    assert!((letter_width.to_inches() - 8.5).abs() < 0.001);
    assert!((letter_height.to_inches() - 11.0).abs() < 0.001);

    // 12pt font in half-points
    let twelve_pt = HalfPoint(24);
    assert!((twelve_pt.to_points() - 12.0).abs() < f64::EPSILON);

    // 100% in Percentage1000
    let hundred = Percentage1000(100_000);
    assert!((hundred.to_fraction() - 1.0).abs() < f64::EPSILON);

    // 90 degrees in Angle60k
    let right_angle = Angle60k(5_400_000);
    assert!((right_angle.to_degrees() - 90.0).abs() < f64::EPSILON);

    // EMU -> Twip -> EMU round-trip (within precision)
    let emu = Emu::from_inches(8.5);
    let twip = emu.to_twip();
    assert_eq!(twip.0, 12240);
    let back = twip.to_emu();
    assert!((back.0 - emu.0).abs() < 635); // Within 1 twip of precision
}

// ---------------------------------------------------------------------------
// Content types builder defaults
// ---------------------------------------------------------------------------

#[test]
fn content_types_builder_has_standard_defaults() {
    let builder = ContentTypesBuilder::new();
    let ct = builder.build();

    // .rels and .xml should have default content types
    let rels_part = PartName::new("/word/_rels/document.xml.rels").unwrap();
    assert_eq!(
        ct.resolve(&rels_part),
        Some("application/vnd.openxmlformats-package.relationships+xml")
    );

    let xml_part = PartName::new("/some/random.xml").unwrap();
    assert_eq!(ct.resolve(&xml_part), Some("application/xml"));
}

// ---------------------------------------------------------------------------
// Relationships builder with multiple rels
// ---------------------------------------------------------------------------

#[test]
fn relationships_builder_generates_sequential_ids() {
    let mut builder = RelationshipsBuilder::new();
    let id1 = builder.add(rel_types::OFFICE_DOCUMENT, "word/document.xml");
    let id2 = builder.add(rel_types::CORE_PROPERTIES, "docProps/core.xml");
    let id3 = builder.add_with_mode(
        rel_types::HYPERLINK,
        "https://example.com",
        TargetMode::External,
    );
    assert_eq!(id1, "rId1");
    assert_eq!(id2, "rId2");
    assert_eq!(id3, "rId3");

    // Round-trip through serialization
    let xml = builder.serialize();
    let rels = Relationships::parse(&xml).unwrap();
    assert_eq!(rels.all().len(), 3);

    let hyperlink = rels.get_by_id("rId3").unwrap();
    assert_eq!(hyperlink.target_mode, TargetMode::External);
    assert_eq!(hyperlink.target, "https://example.com");
}

// ---------------------------------------------------------------------------
// Properties with empty/minimal content
// ---------------------------------------------------------------------------

#[test]
fn core_properties_empty() {
    let props = CoreProperties::default();
    let xml = props.serialize();
    let parsed = CoreProperties::parse(&xml).unwrap();
    assert!(parsed.title.is_none());
    assert!(parsed.creator.is_none());
    assert!(parsed.created.is_none());
}

#[test]
fn app_properties_empty() {
    let props = AppProperties::default();
    let xml = props.serialize();
    let parsed = AppProperties::parse(&xml).unwrap();
    assert!(parsed.application.is_none());
    assert!(parsed.pages.is_none());
}

// ---------------------------------------------------------------------------
// Part name edge cases
// ---------------------------------------------------------------------------

#[test]
fn part_name_root_slash_only_is_valid() {
    // Single slash "/" is a valid part name (represents root)
    assert!(PartName::new("/").is_ok());
}

#[test]
fn part_name_extension_none_for_no_dot() {
    let pn = PartName::new("/data/noext").unwrap();
    assert_eq!(pn.extension(), None);
}

#[test]
fn part_name_deep_nesting() {
    let pn = PartName::new("/a/b/c/d/e/file.xml").unwrap();
    assert_eq!(pn.directory(), "/a/b/c/d/e/");
    assert_eq!(pn.filename(), "file.xml");
    assert_eq!(pn.rels_path(), "/a/b/c/d/e/_rels/file.xml.rels");
}
