//! Editable OPC package for read-modify-write roundtrips.
//!
//! Loads all parts and relationships into memory so unmodified parts
//! can be written back verbatim (preserving images, charts, custom XML, etc.).

use std::collections::HashMap;
use std::fs::File;
use std::io::{Read, Seek, Write};
use std::path::Path;

use zip::CompressionMethod;
use zip::write::{SimpleFileOptions, ZipWriter};

use super::content_types::{ContentTypes, ContentTypesBuilder};
use super::error::Result;
use super::opc::PartName;
use super::relationships::{Relationships, RelationshipsBuilder};

/// A mutable in-memory representation of an OPC package.
///
/// All parts are loaded into memory so individual parts can be replaced
/// while everything else is preserved on save.
pub struct EditablePackage {
    /// Raw bytes for each part.
    parts: HashMap<PartName, Vec<u8>>,
    /// Content type mapping.
    content_types: ContentTypes,
    /// Package-level relationships (_rels/.rels).
    package_rels: Relationships,
    /// Part-level relationships keyed by part name.
    part_rels: HashMap<PartName, Relationships>,
}

impl EditablePackage {
    /// Load an OPC package into an editable in-memory representation.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let file = File::open(path)?;
        Self::from_reader(file)
    }

    /// Load from any `Read + Seek` source.
    pub fn from_reader<R: Read + Seek>(reader: R) -> Result<Self> {
        let mut opc = super::opc::OpcReader::new(reader)?;
        let content_types = opc.content_types().clone();
        let package_rels = opc.package_rels().clone();

        let part_names = opc.part_names();
        let mut parts = HashMap::new();
        let mut part_rels = HashMap::new();

        for name in &part_names {
            let data = opc.read_part(name)?;
            parts.insert(name.clone(), data);

            let rels = opc.read_rels_for(name)?;
            if !rels.all().is_empty() {
                part_rels.insert(name.clone(), rels);
            }
        }

        Ok(Self {
            parts,
            content_types,
            package_rels,
            part_rels,
        })
    }

    /// Get a part's raw bytes.
    pub fn get_part(&self, name: &PartName) -> Option<&[u8]> {
        self.parts.get(name).map(|v| v.as_slice())
    }

    /// Replace or insert a part's raw bytes.
    pub fn set_part(&mut self, name: PartName, data: Vec<u8>) {
        self.parts.insert(name, data);
    }

    /// Get the content types table.
    pub fn content_types(&self) -> &ContentTypes {
        &self.content_types
    }

    /// Get the package-level relationships.
    pub fn package_rels(&self) -> &Relationships {
        &self.package_rels
    }

    /// Get part-level relationships for a part.
    pub fn part_rels(&self, name: &PartName) -> Option<&Relationships> {
        self.part_rels.get(name)
    }

    /// Save the package to a file.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        let file = File::create(path)?;
        self.write_to(file)
    }

    /// Write the package to any `Write + Seek` destination.
    pub fn write_to<W: Write + Seek>(&self, writer: W) -> Result<()> {
        let mut zip = ZipWriter::new(writer);
        let options = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);

        // Sorted: parts and part_rels are HashMaps, so iterating them directly
        // produced a different ZIP entry order on every save. Saving an
        // unchanged document then produced a different byte stream each time,
        // defeating content-hash caching and churning any VCS around the CLI.
        let mut parts: Vec<_> = self.parts.iter().collect();
        parts.sort_by(|a, b| a.0.as_str().cmp(b.0.as_str()));
        for (name, data) in parts {
            let zip_path = &name.as_str()[1..]; // strip leading /
            zip.start_file(zip_path, options)?;
            zip.write_all(data)?;
        }

        // Write part-level .rels files
        let mut part_rels: Vec<_> = self.part_rels.iter().collect();
        part_rels.sort_by(|a, b| a.0.as_str().cmp(b.0.as_str()));
        for (source, rels) in part_rels {
            if rels.all().is_empty() {
                continue;
            }
            let rels_path = source.rels_path();
            let zip_path = &rels_path[1..];
            let mut builder = RelationshipsBuilder::new();
            for rel in rels.all() {
                builder.add_with_id(&rel.id, &rel.rel_type, &rel.target, rel.target_mode);
            }
            let data = builder.serialize();
            zip.start_file(zip_path, options)?;
            zip.write_all(&data)?;
        }

        // Write _rels/.rels
        {
            let mut builder = RelationshipsBuilder::new();
            for rel in self.package_rels.all() {
                builder.add_with_id(&rel.id, &rel.rel_type, &rel.target, rel.target_mode);
            }
            let data = builder.serialize();
            zip.start_file("_rels/.rels", options)?;
            zip.write_all(&data)?;
        }

        // Write [Content_Types].xml
        {
            let mut ct_builder = ContentTypesBuilder::new();
            for (ext, ct) in self.content_types.defaults() {
                ct_builder.add_default(ext, ct);
            }
            // Sorted for the same reason as the parts above: a HashMap made
            // [Content_Types].xml differ byte-for-byte between saves.
            let mut overrides: Vec<_> = self.content_types.overrides().iter().collect();
            overrides.sort_by(|a, b| a.0.as_str().cmp(b.0.as_str()));
            for (pn, ct) in overrides {
                ct_builder.add_override(pn.clone(), ct);
            }
            let data = ct_builder.serialize();
            zip.start_file("[Content_Types].xml", options)?;
            zip.write_all(&data)?;
        }

        zip.finish()?;
        Ok(())
    }
}

/// Replace text inside every `<{tag}>…</{tag}>` element of an OOXML part.
///
/// `tag` is the fully-prefixed element name (`w:t`, `a:t`). Returns the
/// rewritten XML and the number of substitutions.
///
/// Two properties this must have, and previously did not:
///
/// * The bytes between the tags are **escaped** XML, so both the search and
///   the substitution happen on the decoded text and the result is
///   re-escaped. Matching the raw bytes meant `find` never matched text
///   containing `&`, `<` or `>` — the document holds `AT&amp;T`, not
///   `AT&T` — and a replacement containing any of them injected raw markup
///   and produced a file the Office applications refuse to open.
/// * The opening-tag search must match the element, not a prefix of it. A
///   bare `find("<w:t")` also matches `<w:tbl>`, `<w:tab/>`, `<w:tc>` and
///   `<w:trPr>`; `<a:t` likewise matches `<a:tbl>` and `<a:tc>`. Each of
///   those would then have its "text content" rewritten and its structure
///   mangled.
pub fn replace_in_text_elements(
    xml: &str,
    tag: &str,
    find: &str,
    replace: &str,
) -> (String, usize) {
    let open_prefix = format!("<{tag}");
    let close = format!("</{tag}>");
    let mut result = String::with_capacity(xml.len());
    let mut count = 0usize;
    let mut pos = 0usize;

    while pos < xml.len() {
        let Some(tag_start) = find_open_tag(xml, pos, &open_prefix) else {
            result.push_str(&xml[pos..]);
            break;
        };
        let Some(tag_end_offset) = xml[tag_start..].find('>') else {
            result.push_str(&xml[pos..]);
            break;
        };
        let tag_end = tag_start + tag_end_offset + 1;

        if xml[tag_start..tag_end].ends_with("/>") {
            result.push_str(&xml[pos..tag_end]);
            pos = tag_end;
            continue;
        }

        let Some(close_offset) = xml[tag_end..].find(&close) else {
            result.push_str(&xml[pos..]);
            break;
        };
        let close_start = tag_end + close_offset;

        let raw = &xml[tag_end..close_start];
        let decoded = quick_xml::escape::unescape(raw)
            .map(|c| c.into_owned())
            .unwrap_or_else(|_| raw.to_string());
        let hits = decoded.matches(find).count();
        result.push_str(&xml[pos..tag_end]);
        if hits == 0 {
            // Nothing changed — keep the source bytes byte-for-byte rather
            // than round-tripping them through the escaper.
            result.push_str(raw);
        } else {
            count += hits;
            result.push_str(&quick_xml::escape::escape(decoded.replace(find, replace)));
        }
        pos = close_start;
    }

    (result, count)
}

/// Find the next occurrence of `prefix` that is a complete element name —
/// i.e. followed by `>`, `/` or whitespace.
fn find_open_tag(xml: &str, from: usize, prefix: &str) -> Option<usize> {
    let mut pos = from;
    while let Some(off) = xml[pos..].find(prefix) {
        let at = pos + off;
        match xml[at + prefix.len()..].chars().next() {
            Some('>') | Some('/') | Some(' ') | Some('\t') | Some('\n') | Some('\r') => {
                return Some(at);
            },
            _ => pos = at + prefix.len(),
        }
    }
    None
}

#[cfg(test)]
mod determinism_tests {
    use super::*;

    /// Saving an unchanged package produced a different byte stream every
    /// time, because parts, part rels and content-type overrides were all
    /// iterated out of `HashMap`s.
    #[test]
    fn saving_the_same_package_twice_produces_the_same_bytes() {
        let mut wb = crate::xlsx::write::XlsxWriter::new();
        for n in ["Alpha", "Beta", "Gamma", "Delta"] {
            wb.add_sheet(n)
                .add_row(vec![crate::xlsx::write::CellData::String(n.into())]);
        }
        let mut src = std::io::Cursor::new(Vec::new());
        wb.write_to(&mut src).unwrap();

        let save = || {
            let mut r = src.clone();
            r.set_position(0);
            let pkg = EditablePackage::from_reader(r).expect("open");
            let mut out = std::io::Cursor::new(Vec::new());
            pkg.write_to(&mut out).unwrap();
            out.into_inner()
        };

        let first = save();
        for _ in 0..15 {
            assert_eq!(first, save(), "the edit path is not byte-deterministic");
        }
    }
}
