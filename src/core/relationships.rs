use std::collections::HashMap;

use quick_xml::events::{BytesDecl, BytesStart, Event};
use quick_xml::Writer;

use super::error::{Error, Result};
use super::opc::PartName;
use super::xml;

/// Standard relationship type URIs (Transitional / ECMA-376).
pub mod rel_types {
    pub const OFFICE_DOCUMENT: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
    pub const CORE_PROPERTIES: &str =
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
    pub const EXTENDED_PROPERTIES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
    pub const THUMBNAIL: &str =
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail";
    pub const STYLES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
    pub const THEME: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
    pub const SETTINGS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
    pub const FONT_TABLE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable";
    pub const IMAGE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
    pub const HYPERLINK: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    pub const NUMBERING: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
    pub const COMMENTS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
    pub const HEADER: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
    pub const FOOTER: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
    pub const WORKSHEET: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
    pub const SHARED_STRINGS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
    pub const SLIDE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
    pub const SLIDE_LAYOUT: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
    pub const SLIDE_MASTER: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
    pub const NOTES_SLIDE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
    pub const CHART: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
}

/// Strict (ISO 29500) relationship type prefix.
const STRICT_REL_PREFIX: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/";
/// Transitional relationship type prefix.
const TRANSITIONAL_REL_PREFIX: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";

/// Normalize a Strict relationship type URI to its Transitional equivalent.
/// Strict uses `http://purl.oclc.org/ooxml/officeDocument/relationships/...`
/// Transitional uses `http://schemas.openxmlformats.org/officeDocument/2006/relationships/...`
/// If it's already Transitional or an unrecognized URI, return as-is.
fn normalize_rel_type(rel_type: String) -> String {
    if let Some(suffix) = rel_type.strip_prefix(STRICT_REL_PREFIX) {
        format!("{TRANSITIONAL_REL_PREFIX}{suffix}")
    } else {
        rel_type
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TargetMode {
    Internal,
    External,
}

#[derive(Debug, Clone)]
pub struct Relationship {
    pub id: String,
    pub rel_type: String,
    pub target: String,
    pub target_mode: TargetMode,
}

/// Parsed collection of relationships from a `.rels` file.
#[derive(Debug, Clone)]
pub struct Relationships {
    rels: Vec<Relationship>,
    by_id: HashMap<String, usize>,
    by_type: HashMap<String, Vec<usize>>,
}

impl Relationships {
    /// Parse a `.rels` XML file.
    pub fn parse(xml_data: &[u8]) -> Result<Self> {
        let mut reader = xml::make_fast_reader(xml_data);
        let mut rels = Vec::new();

        loop {
            match reader.read_event()? {
                Event::Start(ref e) | Event::Empty(ref e) => {
                    if e.local_name().as_ref() == b"Relationship" {
                        let id = xml::required_attr_str(e, b"Id")?.into_owned();
                        let rel_type =
                            normalize_rel_type(xml::required_attr_str(e, b"Type")?.into_owned());
                        let target = xml::required_attr_str(e, b"Target")?.into_owned();
                        let target_mode = match xml::optional_attr_str(e, b"TargetMode")? {
                            Some(ref m) if m.eq_ignore_ascii_case("External") => {
                                TargetMode::External
                            }
                            _ => TargetMode::Internal,
                        };
                        rels.push(Relationship {
                            id,
                            rel_type,
                            target,
                            target_mode,
                        });
                    }
                }
                Event::Eof => break,
                _ => {}
            }
        }

        Ok(Self::from_vec(rels))
    }

    /// Create an empty relationships collection.
    pub fn empty() -> Self {
        Self {
            rels: Vec::new(),
            by_id: HashMap::new(),
            by_type: HashMap::new(),
        }
    }

    fn from_vec(rels: Vec<Relationship>) -> Self {
        let mut by_id = HashMap::with_capacity(rels.len());
        let mut by_type: HashMap<String, Vec<usize>> = HashMap::new();
        for (i, r) in rels.iter().enumerate() {
            by_id.insert(r.id.clone(), i);
            by_type.entry(r.rel_type.clone()).or_default().push(i);
        }
        Self {
            rels,
            by_id,
            by_type,
        }
    }

    pub fn get_by_id(&self, id: &str) -> Option<&Relationship> {
        self.by_id.get(id).map(|&i| &self.rels[i])
    }

    pub fn get_by_type(&self, rel_type: &str) -> Vec<&Relationship> {
        self.by_type
            .get(rel_type)
            .map(|indices| indices.iter().map(|&i| &self.rels[i]).collect())
            .unwrap_or_default()
    }

    pub fn first_by_type(&self, rel_type: &str) -> Option<&Relationship> {
        self.by_type
            .get(rel_type)
            .and_then(|indices| indices.first().map(|&i| &self.rels[i]))
    }

    pub fn all(&self) -> &[Relationship] {
        &self.rels
    }

    /// Resolve an internal relationship target relative to a source part name.
    ///
    /// For package-level relationships (source is root), pass a source of `PartName::new("/").ok()`
    /// or use `resolve_target_from_root`.
    pub fn resolve_target(&self, id: &str, source: &PartName) -> Result<PartName> {
        let rel = self
            .get_by_id(id)
            .ok_or_else(|| Error::RelationshipNotFound(id.to_string()))?;

        if rel.target_mode == TargetMode::External {
            return Err(Error::Unsupported(format!(
                "cannot resolve external target: {}",
                rel.target
            )));
        }

        source.resolve_relative(&rel.target)
    }

    /// Resolve an internal relationship target relative to the package root.
    pub fn resolve_target_from_root(&self, id: &str) -> Result<PartName> {
        let rel = self
            .get_by_id(id)
            .ok_or_else(|| Error::RelationshipNotFound(id.to_string()))?;

        if rel.target_mode == TargetMode::External {
            return Err(Error::Unsupported(format!(
                "cannot resolve external target: {}",
                rel.target
            )));
        }

        // Package-level targets are relative to root
        let target = if rel.target.starts_with('/') {
            rel.target.clone()
        } else {
            format!("/{}", rel.target)
        };
        PartName::new(&target)
    }
}

/// Builder for constructing relationship XML for the write path.
#[derive(Debug, Clone)]
pub struct RelationshipsBuilder {
    rels: Vec<Relationship>,
    next_id: u32,
}

impl RelationshipsBuilder {
    pub fn new() -> Self {
        Self {
            rels: Vec::new(),
            next_id: 1,
        }
    }

    /// Add a relationship and return the generated rId.
    pub fn add(&mut self, rel_type: &str, target: &str) -> String {
        self.add_with_mode(rel_type, target, TargetMode::Internal)
    }

    /// Add a relationship with explicit target mode and return the generated rId.
    pub fn add_with_mode(
        &mut self,
        rel_type: &str,
        target: &str,
        target_mode: TargetMode,
    ) -> String {
        let id = format!("rId{}", self.next_id);
        self.next_id += 1;
        self.rels.push(Relationship {
            id: id.clone(),
            rel_type: rel_type.to_string(),
            target: target.to_string(),
            target_mode,
        });
        id
    }

    /// Add a relationship with an explicit ID (for round-trip preservation).
    pub fn add_with_id(
        &mut self,
        id: &str,
        rel_type: &str,
        target: &str,
        target_mode: TargetMode,
    ) {
        self.rels.push(Relationship {
            id: id.to_string(),
            rel_type: rel_type.to_string(),
            target: target.to_string(),
            target_mode,
        });
    }

    pub fn is_empty(&self) -> bool {
        self.rels.is_empty()
    }

    /// Serialize to `.rels` XML bytes.
    pub fn serialize(&self) -> Vec<u8> {
        let mut writer = Writer::new_with_indent(Vec::new(), b' ', 2);

        writer
            .write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
            .expect("write decl");

        let mut root = BytesStart::new("Relationships");
        root.push_attribute((
            "xmlns",
            "http://schemas.openxmlformats.org/package/2006/relationships",
        ));
        writer.write_event(Event::Start(root)).expect("write start");

        for rel in &self.rels {
            let mut elem = BytesStart::new("Relationship");
            elem.push_attribute(("Id", rel.id.as_str()));
            elem.push_attribute(("Type", rel.rel_type.as_str()));
            elem.push_attribute(("Target", rel.target.as_str()));
            if rel.target_mode == TargetMode::External {
                elem.push_attribute(("TargetMode", "External"));
            }
            writer.write_event(Event::Empty(elem)).expect("write rel");
        }

        writer
            .write_event(Event::End(quick_xml::events::BytesEnd::new(
                "Relationships",
            )))
            .expect("write end");

        writer.into_inner()
    }
}

impl Default for RelationshipsBuilder {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const SAMPLE_RELS: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
    Target="docProps/core.xml"/>
  <Relationship Id="rId3"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
    Target="docProps/app.xml"/>
</Relationships>"#;

    #[test]
    fn parse_relationships() {
        let rels = Relationships::parse(SAMPLE_RELS).unwrap();
        assert_eq!(rels.all().len(), 3);
    }

    #[test]
    fn get_by_id() {
        let rels = Relationships::parse(SAMPLE_RELS).unwrap();
        let r = rels.get_by_id("rId1").unwrap();
        assert_eq!(r.target, "word/document.xml");
        assert_eq!(r.target_mode, TargetMode::Internal);
    }

    #[test]
    fn get_by_type() {
        let rels = Relationships::parse(SAMPLE_RELS).unwrap();
        let docs = rels.get_by_type(rel_types::OFFICE_DOCUMENT);
        assert_eq!(docs.len(), 1);
        assert_eq!(docs[0].target, "word/document.xml");
    }

    #[test]
    fn first_by_type() {
        let rels = Relationships::parse(SAMPLE_RELS).unwrap();
        let r = rels.first_by_type(rel_types::CORE_PROPERTIES).unwrap();
        assert_eq!(r.target, "docProps/core.xml");
    }

    #[test]
    fn resolve_target_from_root() {
        let rels = Relationships::parse(SAMPLE_RELS).unwrap();
        let pn = rels.resolve_target_from_root("rId1").unwrap();
        assert_eq!(pn.as_str(), "/word/document.xml");
    }

    #[test]
    fn builder_round_trip() {
        let mut builder = RelationshipsBuilder::new();
        let id = builder.add(rel_types::OFFICE_DOCUMENT, "word/document.xml");
        assert_eq!(id, "rId1");

        let xml = builder.serialize();
        let rels = Relationships::parse(&xml).unwrap();
        assert_eq!(rels.all().len(), 1);
        assert_eq!(
            rels.get_by_id("rId1").unwrap().rel_type,
            rel_types::OFFICE_DOCUMENT
        );
    }
}
