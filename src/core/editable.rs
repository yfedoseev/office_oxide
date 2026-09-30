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

    /// Save the package to a file, atomically.
    ///
    /// The package is written to a temporary file in the destination's
    /// directory, flushed to disk, and then renamed over `path`. Opening
    /// `path` with `File::create` truncated it before a single byte of the
    /// new package was written, so any failure part-way through — disk
    /// full, an I/O error, the process being killed, or a part that cannot
    /// be serialised — destroyed the original when saving in place, which is
    /// what the CLI `replace` command and the MCP `replace_text` tool do by
    /// default. On failure the destination is left exactly as it was and the
    /// temporary file is removed.
    ///
    /// Like `python-docx`'s `Document.save` and every Office application,
    /// an existing file at `path` is replaced.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        write_atomically(path.as_ref(), |file| self.write_to(file))
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

/// Reject a search string no text replacement can sensibly use.
///
/// An empty `find` matches at every character boundary, so `str::replace`
/// interleaves the replacement between every character of every run — and
/// the edit reports a large, plausible count. Only the CLI used to guard
/// against it; the MCP server (which then overwrote its input in place), the
/// FFI and every binding passed it straight through. Checking it here gives
/// every surface the same `Err`.
pub fn check_find(find: &str) -> Result<()> {
    if find.is_empty() {
        Err(super::error::Error::InvalidArgument("search string cannot be empty".into()))
    } else {
        Ok(())
    }
}

/// Write a file through a sibling temporary file, then rename it into place.
///
/// The temporary file lives in the destination's directory so the final
/// `rename` stays on one filesystem (and is therefore atomic on POSIX, and a
/// replace-existing move on Windows). It is flushed with `sync_all` before
/// the rename, so a crash cannot leave a renamed-but-empty file behind.
fn write_atomically(path: &Path, write: impl FnOnce(&mut File) -> Result<()>) -> Result<()> {
    use std::sync::atomic::{AtomicU64, Ordering};
    static COUNTER: AtomicU64 = AtomicU64::new(0);

    let dir = match path.parent() {
        Some(p) if !p.as_os_str().is_empty() => p,
        _ => Path::new("."),
    };
    let file_name = path.file_name().ok_or_else(|| {
        super::error::Error::InvalidArgument(format!("'{}' does not name a file", path.display()))
    })?;

    // `create_new` refuses to reuse a name, so a stale temporary file from a
    // killed process (or a concurrent save) is never clobbered or adopted.
    let (tmp_path, mut file) = loop {
        let mut name = std::ffi::OsString::from(".");
        name.push(file_name);
        name.push(format!(
            ".{}.{}.tmp",
            std::process::id(),
            COUNTER.fetch_add(1, Ordering::Relaxed)
        ));
        let candidate = dir.join(name);
        match std::fs::OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(&candidate)
        {
            Ok(f) => break (candidate, f),
            Err(e) if e.kind() == std::io::ErrorKind::AlreadyExists => continue,
            Err(e) => return Err(e.into()),
        }
    };

    let result = write(&mut file)
        .and_then(|()| file.sync_all().map_err(Into::into))
        .and_then(|()| {
            drop(file);
            std::fs::rename(&tmp_path, path).map_err(Into::into)
        });
    if result.is_err() {
        let _ = std::fs::remove_file(&tmp_path);
    }
    result
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
    // An empty `find` matches between every pair of characters; never
    // substitute on it. Callers that can report an error reject it first via
    // [`check_find`]; this keeps the direct per-format entry points safe too.
    if find.is_empty() {
        return (xml.to_string(), 0);
    }
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
    fn test_saving_the_same_package_twice_produces_the_same_bytes() {
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

#[cfg(test)]
mod save_tests {
    use super::*;

    fn xlsx_bytes() -> Vec<u8> {
        let mut wb = crate::xlsx::write::XlsxWriter::new();
        wb.add_sheet("S")
            .add_row(vec![crate::xlsx::write::CellData::String("keep".into())]);
        let mut buf = std::io::Cursor::new(Vec::new());
        wb.write_to(&mut buf).unwrap();
        buf.into_inner()
    }

    fn scratch_dir(tag: &str) -> std::path::PathBuf {
        let dir = std::env::temp_dir()
            .join(format!("office_oxide_editable_{tag}_{}", std::process::id()));
        std::fs::create_dir_all(&dir).unwrap();
        dir
    }

    /// Saving over the file the package was read from truncated it before a
    /// single byte of the new package was written, so any failure part-way
    /// through the write (disk full, an I/O error, a kill, or a package that
    /// cannot be serialised) destroyed the original. The CLI `replace`
    /// command and the MCP `replace_text` tool both overwrite their input by
    /// default.
    #[test]
    fn test_a_failed_in_place_save_leaves_the_original_intact() {
        let dir = scratch_dir("failed_save");
        let path = dir.join("book.xlsx");
        let original = xlsx_bytes();
        std::fs::write(&path, &original).unwrap();

        let mut pkg = EditablePackage::open(&path).unwrap();
        // A part whose zip entry collides with the package relationships the
        // writer emits near the end: serialisation fails after most parts
        // are already out.
        pkg.set_part(PartName::new("/_rels/.rels").unwrap(), b"<x/>".to_vec());
        assert!(pkg.save(&path).is_err(), "a colliding entry must fail the save");

        assert!(
            std::fs::read(&path).unwrap() == original,
            "a failed save must not touch the original file"
        );
        let leftovers: Vec<_> = std::fs::read_dir(&dir)
            .unwrap()
            .map(|e| e.unwrap().file_name())
            .collect();
        assert_eq!(leftovers.len(), 1, "no temporary file may be left behind: {leftovers:?}");
        std::fs::remove_dir_all(&dir).ok();
    }

    /// The successful path still replaces the destination, and leaves
    /// nothing else behind.
    #[test]
    fn test_an_in_place_save_replaces_the_file() {
        let dir = scratch_dir("ok_save");
        let path = dir.join("book.xlsx");
        std::fs::write(&path, xlsx_bytes()).unwrap();
        let mut pkg = EditablePackage::open(&path).unwrap();
        pkg.set_part(PartName::new("/extra.bin").unwrap(), b"payload".to_vec());
        pkg.save(&path).unwrap();
        let reopened = EditablePackage::open(&path).unwrap();
        assert_eq!(reopened.get_part(&PartName::new("/extra.bin").unwrap()), Some(&b"payload"[..]));
        let entries = std::fs::read_dir(&dir).unwrap().count();
        assert_eq!(entries, 1, "no temporary file may be left behind");
        std::fs::remove_dir_all(&dir).ok();
    }
}
