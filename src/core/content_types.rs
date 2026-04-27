use std::collections::HashMap;

use quick_xml::Writer;
use quick_xml::events::{BytesDecl, BytesStart, Event};

use super::error::Result;
use super::opc::PartName;
use super::xml;

/// Content types table parsed from `[Content_Types].xml`.
///
/// Maps parts to MIME content types via default (by extension) and override (by part name) entries.
#[derive(Debug, Clone)]
pub struct ContentTypes {
    /// Extension -> content type (both lowercased).
    defaults: HashMap<String, String>,
    /// Part name -> content type.
    overrides: HashMap<PartName, String>,
}

impl ContentTypes {
    /// Parse `[Content_Types].xml` from raw XML bytes.
    pub fn parse(xml_data: &[u8]) -> Result<Self> {
        let mut reader = xml::make_fast_reader(xml_data);
        let mut defaults = HashMap::new();
        let mut overrides = HashMap::new();

        loop {
            match reader.read_event()? {
                Event::Start(ref e) | Event::Empty(ref e) => {
                    let local = e.local_name();
                    let local_bytes = local.as_ref();

                    match local_bytes {
                        b"Default" => {
                            let ext = xml::required_attr_str(e, b"Extension")?;
                            let ct = xml::required_attr_str(e, b"ContentType")?;
                            defaults.insert(ext.to_ascii_lowercase(), ct.into_owned());
                        },
                        b"Override" => {
                            let pn = xml::required_attr_str(e, b"PartName")?;
                            let ct = xml::required_attr_str(e, b"ContentType")?;
                            let part_name = PartName::new(&pn)?;
                            overrides.insert(part_name, ct.into_owned());
                        },
                        _ => {},
                    }
                },
                Event::Eof => break,
                _ => {},
            }
        }

        Ok(Self {
            defaults,
            overrides,
        })
    }

    /// Resolve the content type for a part name.
    /// Checks overrides first, then defaults by file extension.
    pub fn resolve(&self, part_name: &PartName) -> Option<&str> {
        // Override takes precedence (case-insensitive via PartName's Eq)
        if let Some(ct) = self.overrides.get(part_name) {
            return Some(ct.as_str());
        }
        // Fall back to default by extension
        if let Some(ext) = part_name.extension() {
            if let Some(ct) = self.defaults.get(&ext.to_ascii_lowercase()) {
                return Some(ct.as_str());
            }
        }
        None
    }

    /// Return the default content-type map keyed by lowercase file extension.
    pub fn defaults(&self) -> &HashMap<String, String> {
        &self.defaults
    }

    /// Return the override content-type map keyed by part name.
    pub fn overrides(&self) -> &HashMap<PartName, String> {
        &self.overrides
    }
}

/// Builder for constructing a `[Content_Types].xml` for the write path.
#[derive(Debug, Clone)]
pub struct ContentTypesBuilder {
    defaults: HashMap<String, String>,
    overrides: Vec<(PartName, String)>,
}

impl ContentTypesBuilder {
    /// Create a new builder pre-populated with the standard OPC defaults.
    pub fn new() -> Self {
        let mut defaults = HashMap::new();
        // Standard defaults present in all OPC packages
        defaults.insert(
            "rels".to_string(),
            "application/vnd.openxmlformats-package.relationships+xml".to_string(),
        );
        defaults.insert("xml".to_string(), "application/xml".to_string());
        Self {
            defaults,
            overrides: Vec::new(),
        }
    }

    /// Register a default content type for a file extension (case-insensitive).
    pub fn add_default(&mut self, extension: &str, content_type: &str) {
        self.defaults
            .insert(extension.to_ascii_lowercase(), content_type.to_string());
    }

    /// Register an override content type for a specific part name.
    pub fn add_override(&mut self, part_name: PartName, content_type: &str) {
        self.overrides.push((part_name, content_type.to_string()));
    }

    /// Consume the builder and return a `ContentTypes` lookup table.
    pub fn build(self) -> ContentTypes {
        let overrides = self.overrides.into_iter().collect();
        ContentTypes {
            defaults: self.defaults,
            overrides,
        }
    }

    /// Serialize to XML bytes.
    pub fn serialize(&self) -> Vec<u8> {
        let mut writer = Writer::new_with_indent(Vec::new(), b' ', 2);

        writer
            .write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
            .expect("write decl");

        let mut types = BytesStart::new("Types");
        types.push_attribute((
            "xmlns",
            "http://schemas.openxmlformats.org/package/2006/content-types",
        ));
        writer
            .write_event(Event::Start(types))
            .expect("write start");

        // Write defaults (sorted for deterministic output)
        let mut sorted_defaults: Vec<_> = self.defaults.iter().collect();
        sorted_defaults.sort_by_key(|(k, _)| k.as_str());
        for (ext, ct) in sorted_defaults {
            let mut elem = BytesStart::new("Default");
            elem.push_attribute(("Extension", ext.as_str()));
            elem.push_attribute(("ContentType", ct.as_str()));
            writer
                .write_event(Event::Empty(elem))
                .expect("write default");
        }

        // Write overrides
        for (pn, ct) in &self.overrides {
            let mut elem = BytesStart::new("Override");
            elem.push_attribute(("PartName", pn.as_str()));
            elem.push_attribute(("ContentType", ct.as_str()));
            writer
                .write_event(Event::Empty(elem))
                .expect("write override");
        }

        writer
            .write_event(Event::End(quick_xml::events::BytesEnd::new("Types")))
            .expect("write end");

        writer.into_inner()
    }
}

impl Default for ContentTypesBuilder {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const SAMPLE_CT_XML: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/docProps/core.xml"
            ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>"#;

    #[test]
    fn parse_content_types() {
        let ct = ContentTypes::parse(SAMPLE_CT_XML).unwrap();
        assert_eq!(ct.defaults().len(), 3);
        assert_eq!(ct.overrides().len(), 2);
    }

    #[test]
    fn resolve_override() {
        let ct = ContentTypes::parse(SAMPLE_CT_XML).unwrap();
        let pn = PartName::new("/word/document.xml").unwrap();
        assert_eq!(
            ct.resolve(&pn),
            Some(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
            )
        );
    }

    #[test]
    fn resolve_default_by_extension() {
        let ct = ContentTypes::parse(SAMPLE_CT_XML).unwrap();
        let pn = PartName::new("/word/media/image1.png").unwrap();
        assert_eq!(ct.resolve(&pn), Some("image/png"));
    }

    #[test]
    fn resolve_unknown_returns_none() {
        let ct = ContentTypes::parse(SAMPLE_CT_XML).unwrap();
        let pn = PartName::new("/word/unknown.bin").unwrap();
        assert_eq!(ct.resolve(&pn), None);
    }

    #[test]
    fn builder_round_trip() {
        let mut builder = ContentTypesBuilder::new();
        builder.add_default("png", "image/png");
        builder.add_override(
            PartName::new("/word/document.xml").unwrap(),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        );
        let xml = builder.serialize();
        let ct = ContentTypes::parse(&xml).unwrap();
        assert_eq!(ct.defaults().get("png"), Some(&"image/png".to_string()));
        let pn = PartName::new("/word/document.xml").unwrap();
        assert!(ct.resolve(&pn).is_some());
    }
}
