use quick_xml::Writer;
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use serde::{Deserialize, Serialize};

use super::error::Result;
use super::xml;

// ---------------------------------------------------------------------------
// Core Properties (Dublin Core metadata) — docProps/core.xml
// ---------------------------------------------------------------------------

/// Core properties (Dublin Core + OPC metadata) from `docProps/core.xml`.
#[derive(Debug, Default, Clone, Serialize, Deserialize)]
pub struct CoreProperties {
    /// Document title.
    pub title: Option<String>,
    /// Document subject.
    pub subject: Option<String>,
    /// Primary author.
    pub creator: Option<String>,
    /// Search keywords.
    pub keywords: Option<String>,
    /// Document description or abstract.
    pub description: Option<String>,
    /// User who last saved the document.
    pub last_modified_by: Option<String>,
    /// Revision number as a string.
    pub revision: Option<String>,
    /// Creation date-time (ISO 8601).
    pub created: Option<String>,
    /// Last-modified date-time (ISO 8601).
    pub modified: Option<String>,
    /// Document category.
    pub category: Option<String>,
    /// Content status (e.g., "Draft", "Final").
    pub content_status: Option<String>,
    /// Document language.
    pub language: Option<String>,
}

impl CoreProperties {
    /// Parse core properties from `docProps/core.xml` bytes.
    pub fn parse(xml_data: &[u8]) -> Result<Self> {
        let mut reader = property_reader(xml_data);
        let mut props = CoreProperties::default();

        // State: which element are we inside?
        // Since we no longer have namespace resolution, we match on local name only.
        // The element names are unique enough across namespaces to be unambiguous.
        let mut ctx = Ctx::None;

        loop {
            match reader.read_event()? {
                Event::Start(ref e) => {
                    let local = e.local_name();
                    let local_bytes = local.as_ref();

                    ctx = match local_bytes {
                        "title" => Ctx::Title,
                        "subject" => Ctx::Subject,
                        "creator" => Ctx::Creator,
                        "description" => Ctx::Description,
                        "language" => Ctx::Language,
                        "created" => Ctx::Created,
                        "modified" => Ctx::Modified,
                        "keywords" => Ctx::Keywords,
                        "lastModifiedBy" => Ctx::LastModifiedBy,
                        "revision" => Ctx::Revision,
                        "category" => Ctx::Category,
                        "contentStatus" => Ctx::ContentStatus,
                        _ => Ctx::None,
                    };
                },
                // Entity references arrive as their own event; a property
                // value like `Smith &amp; Co` would otherwise lose the `&`.
                Event::GeneralRef(ref e) => {
                    let text = crate::core::xml::resolve_general_ref(e)?;
                    append_ctx(&mut props, ctx, &text);
                },
                Event::Text(ref e) => {
                    let text = crate::core::xml::unescape_text(e)?;
                    append_ctx(&mut props, ctx, &text);
                },
                Event::End(_) => {
                    ctx = Ctx::None;
                },
                Event::Eof => break,
                _ => {},
            }
        }

        for slot in [
            &mut props.title,
            &mut props.subject,
            &mut props.creator,
            &mut props.keywords,
            &mut props.description,
            &mut props.last_modified_by,
            &mut props.revision,
            &mut props.created,
            &mut props.modified,
            &mut props.category,
            &mut props.content_status,
            &mut props.language,
        ] {
            *slot = slot
                .take()
                .map(|v| v.trim().to_string())
                .filter(|v| !v.is_empty());
        }
        Ok(props)
    }

    /// Serialize to `docProps/core.xml` bytes.
    pub fn serialize(&self) -> Vec<u8> {
        let mut w = Writer::new_with_indent(Vec::new(), b' ', 2);

        w.write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
            .expect("write decl");

        let mut root = BytesStart::new("cp:coreProperties");
        root.push_attribute((
            "xmlns:cp",
            "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
        ));
        root.push_attribute(("xmlns:dc", "http://purl.org/dc/elements/1.1/"));
        root.push_attribute(("xmlns:dcterms", "http://purl.org/dc/terms/"));
        root.push_attribute(("xmlns:dcmitype", "http://purl.org/dc/dcmitype/"));
        root.push_attribute(("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"));
        w.write_event(Event::Start(root)).expect("write root");

        write_optional_element(&mut w, "dc:title", self.title.as_deref());
        write_optional_element(&mut w, "dc:subject", self.subject.as_deref());
        write_optional_element(&mut w, "dc:creator", self.creator.as_deref());
        write_optional_element(&mut w, "dc:description", self.description.as_deref());
        write_optional_element(&mut w, "dc:language", self.language.as_deref());
        write_optional_element(&mut w, "cp:keywords", self.keywords.as_deref());
        write_optional_element(&mut w, "cp:category", self.category.as_deref());
        write_optional_element(&mut w, "cp:contentStatus", self.content_status.as_deref());
        write_optional_element(&mut w, "cp:lastModifiedBy", self.last_modified_by.as_deref());
        write_optional_element(&mut w, "cp:revision", self.revision.as_deref());

        // A malformed source `dcterms:created`/`modified` (e.g. a stray
        // "aaa" mixed into the year, or a PHP writer's space-padded
        // single-digit month/day) used to be copied through verbatim,
        // authoring an invalid docProps/core.xml out of a file someone
        // else broke. Normalize leniently where the intent is unambiguous,
        // and drop the element (both are optional) rather than emit
        // something the W3CDTF restricted union rejects.
        if let Some(created) = self.created.as_deref().and_then(normalize_w3cdtf) {
            write_datetime_element(&mut w, "dcterms:created", &created);
        }
        if let Some(modified) = self.modified.as_deref().and_then(normalize_w3cdtf) {
            write_datetime_element(&mut w, "dcterms:modified", &modified);
        }

        w.write_event(Event::End(BytesEnd::new("cp:coreProperties")))
            .expect("write end root");

        w.into_inner()
    }
}

/// A reader for the property parts. Unlike [`xml::make_fast_reader`] it
/// keeps text whitespace: a value holding an entity reference arrives as
/// several events, and trimming each one deleted the spaces around the
/// reference (`AT&amp;T Labs` read as `AT&TLabs`). Values are trimmed
/// once, whole, after parsing.
fn property_reader(xml_data: &[u8]) -> quick_xml::Reader<&[u8]> {
    let mut reader = xml::make_fast_reader(xml_data);
    reader.config_mut().trim_text(false);
    reader
}

fn write_optional_element(w: &mut Writer<Vec<u8>>, tag: &str, value: Option<&str>) {
    if let Some(text) = value {
        w.write_event(Event::Start(BytesStart::new(tag)))
            .expect("write start");
        w.write_event(Event::Text(BytesText::new(&crate::core::xml::sanitize_xml_text(text))))
            .expect("write text");
        w.write_event(Event::End(BytesEnd::new(tag)))
            .expect("write end");
    }
}

/// Leniently parse a `dcterms:W3CDTF` value and re-emit it in canonical
/// form, or return `None` when it can't be recovered as a valid one.
///
/// Handles the two shapes found in a 6,062-file real-world corpus sweep
///: a stray non-digit character mixed into a numeric
/// component (`2014aaa-10-28T11:34:00Z` — not recoverable, the intent is
/// ambiguous) and a single-digit month/day/hour written with a leading
/// space instead of zero-padding (`2021- 9- 3T20:25:22Z` — a known
/// PHP-writer quirk, unambiguously `2021-09-03T20:25:22Z`).
pub(crate) fn normalize_w3cdtf(value: &str) -> Option<String> {
    let value = value.trim();
    let (date_part, time_part) = match value.split_once('T') {
        Some((d, t)) => (d, Some(t)),
        None => (value, None),
    };

    let date_fields: Vec<&str> = date_part.split('-').collect();
    // W3CDTF allows YYYY, YYYY-MM, or YYYY-MM-DD precision.
    if date_fields.is_empty() || date_fields.len() > 3 {
        return None;
    }
    let year = date_fields[0].trim();
    if year.len() != 4 || !year.bytes().all(|b| b.is_ascii_digit()) {
        return None;
    }
    let mut normalized = year.to_string();
    for field in &date_fields[1..] {
        let f = field.trim();
        if f.is_empty() || f.len() > 2 || !f.bytes().all(|b| b.is_ascii_digit()) {
            return None;
        }
        normalized.push('-');
        normalized.push_str(&format!("{f:0>2}"));
    }

    let Some(time_part) = time_part else {
        return Some(normalized);
    };

    let (time_body, tz) = split_w3cdtf_timezone(time_part);
    let time_fields: Vec<&str> = time_body.split(':').collect();
    if time_fields.len() < 2 || time_fields.len() > 3 {
        return None;
    }
    normalized.push('T');
    for (i, field) in time_fields.iter().enumerate() {
        let f = field.trim();
        // Seconds may carry a fractional part (SS.sss) — validate the
        // integer portion strictly and keep the fraction verbatim.
        let (whole, frac) = match f.split_once('.') {
            Some((w, fr)) => (w, Some(fr)),
            None => (f, None),
        };
        if whole.is_empty() || whole.len() > 2 || !whole.bytes().all(|b| b.is_ascii_digit()) {
            return None;
        }
        if let Some(fr) = frac
            && (fr.is_empty() || !fr.bytes().all(|b| b.is_ascii_digit()))
        {
            return None;
        }
        if i > 0 {
            normalized.push(':');
        }
        normalized.push_str(&format!("{whole:0>2}"));
        if let Some(fr) = frac {
            normalized.push('.');
            normalized.push_str(fr);
        }
    }
    normalized.push_str(&tz);
    Some(normalized)
}

/// Split a W3CDTF time-of-day into its body and trailing timezone
/// designator (`Z` or `±HH:MM`), if any.
fn split_w3cdtf_timezone(time_part: &str) -> (&str, String) {
    let t = time_part.trim();
    if let Some(stripped) = t.strip_suffix('Z') {
        return (stripped.trim_end(), "Z".to_string());
    }
    if let Some(pos) = t.rfind(['+', '-'])
        && pos > 0
    {
        let (body, offset) = t.split_at(pos);
        if offset.len() >= 3 {
            return (body.trim_end(), offset.to_string());
        }
    }
    (t, String::new())
}

fn write_datetime_element(w: &mut Writer<Vec<u8>>, tag: &str, value: &str) {
    let mut elem = BytesStart::new(tag);
    elem.push_attribute(("xsi:type", "dcterms:W3CDTF"));
    w.write_event(Event::Start(elem)).expect("write start");
    w.write_event(Event::Text(BytesText::new(&crate::core::xml::sanitize_xml_text(value))))
        .expect("write text");
    w.write_event(Event::End(BytesEnd::new(tag)))
        .expect("write end");
}

/// Which `docProps/core.xml` element the reader is currently inside.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Ctx {
    None,
    Title,
    Subject,
    Creator,
    Keywords,
    Description,
    LastModifiedBy,
    Revision,
    Created,
    Modified,
    Category,
    ContentStatus,
    Language,
}

/// Append a text fragment to the core property named by `ctx`.
///
/// A single property value can arrive as several events — text split around
/// an entity reference, for example — so fragments accumulate rather than
/// overwrite.
fn append_ctx(props: &mut CoreProperties, ctx: Ctx, text: &str) {
    if text.is_empty() {
        return;
    }
    let slot = match ctx {
        Ctx::Title => &mut props.title,
        Ctx::Subject => &mut props.subject,
        Ctx::Creator => &mut props.creator,
        Ctx::Keywords => &mut props.keywords,
        Ctx::Description => &mut props.description,
        Ctx::LastModifiedBy => &mut props.last_modified_by,
        Ctx::Revision => &mut props.revision,
        Ctx::Created => &mut props.created,
        Ctx::Modified => &mut props.modified,
        Ctx::Category => &mut props.category,
        Ctx::ContentStatus => &mut props.content_status,
        Ctx::Language => &mut props.language,
        Ctx::None => return,
    };
    slot.get_or_insert_with(String::new).push_str(text);
}

/// Read and parse `docProps/core.xml` from an open OPC package.
///
/// Resolves the part through the package-level `core-properties`
/// relationship and falls back to the conventional `/docProps/core.xml`
/// path for packages that omit the relationship. Returns `None` when the
/// part is absent or unparseable — document metadata is decoration, never
/// a reason to fail opening a file.
pub fn read_core_properties<R: std::io::Read + std::io::Seek>(
    opc: &mut super::opc::OpcReader<R>,
) -> Option<CoreProperties> {
    let part = opc
        .package_rels()
        .first_by_type(super::relationships::rel_types::CORE_PROPERTIES)
        .and_then(|rel| {
            super::opc::PartName::new(&format!("/{}", rel.target.trim_start_matches('/'))).ok()
        })
        .filter(|p| opc.has_part(p))
        .or_else(|| {
            super::opc::PartName::new("/docProps/core.xml")
                .ok()
                .filter(|p| opc.has_part(p))
        })?;
    let data = opc.read_part(&part).ok()?;
    CoreProperties::parse(&data).ok()
}

/// Read and parse `docProps/app.xml` (extended/application properties —
/// company, producing application, template, editing time, and page/word/
/// character/line/paragraph/slide/notes/hidden-slide counts) from an open
/// package, the same way [`read_core_properties`] reads `docProps/core.xml`.
///
/// `AppProperties::parse` already existed, fully tested, but nothing on
/// the read side ever called it — company name and every count field were
/// unreachable through any public API.
pub fn read_app_properties<R: std::io::Read + std::io::Seek>(
    opc: &mut super::opc::OpcReader<R>,
) -> Option<AppProperties> {
    let part = opc
        .package_rels()
        .first_by_type(super::relationships::rel_types::EXTENDED_PROPERTIES)
        .and_then(|rel| {
            super::opc::PartName::new(&format!("/{}", rel.target.trim_start_matches('/'))).ok()
        })
        .filter(|p| opc.has_part(p))
        .or_else(|| {
            super::opc::PartName::new("/docProps/app.xml")
                .ok()
                .filter(|p| opc.has_part(p))
        })?;
    let data = opc.read_part(&part).ok()?;
    AppProperties::parse(&data).ok()
}

// ---------------------------------------------------------------------------
// Package-level properties: custom properties, signature, thumbnail
// ---------------------------------------------------------------------------

/// The package-level properties beyond `core.xml`/`app.xml`: the custom
/// properties part, whether the package is signed, and its thumbnail.
#[derive(Debug, Default, Clone)]
pub struct PackageProperties {
    /// User-defined properties from `docProps/custom.xml`.
    pub custom: Vec<crate::ir::CustomProperty>,
    /// The package has a digital-signature origin relationship.
    pub has_digital_signature: bool,
    /// The package thumbnail, when present.
    pub thumbnail: Option<crate::ir::Image>,
    /// Zip entries read while opening the package whose bytes failed their
    /// recorded CRC-32. They were read as stored and may be damaged; every
    /// converter reports them through `Metadata::warnings`.
    pub crc_mismatched_parts: Vec<String>,
}

/// Every package-level property part, located the same way from any
/// package reader.
#[derive(Debug, Default, Clone)]
pub(crate) struct PackageMetadata {
    pub(crate) core: Option<CoreProperties>,
    pub(crate) app: Option<AppProperties>,
    pub(crate) package: PackageProperties,
}

/// Read every package-level property part named by `rels` (the package's
/// `_rels/.rels`), with `read` returning a zip entry's bytes by its path
/// (no leading `/`). Parts are found through their relationships, never by
/// a hard-coded path — except that `core.xml` and `app.xml` fall back to
/// their conventional paths for packages that omit the relationship.
/// Malformed parts yield nothing: metadata is never a reason to fail
/// opening a document.
pub(crate) fn read_package_metadata(
    rels: &super::relationships::Relationships,
    mut read: impl FnMut(&str) -> Option<Vec<u8>>,
) -> PackageMetadata {
    use super::relationships::rel_types;
    let target_of = |rel_type: &str| -> Option<String> {
        let rel = rels.first_by_type(rel_type)?;
        if rel.target_mode == super::relationships::TargetMode::External {
            return None;
        }
        let part =
            super::opc::PartName::new(&format!("/{}", rel.target.trim_start_matches('/'))).ok()?;
        Some(part.as_str()[1..].to_string())
    };
    let mut read_part = |rel_type: &str, fallback: Option<&str>| -> Option<(String, Vec<u8>)> {
        if let Some(path) = target_of(rel_type) {
            if let Some(data) = read(&path) {
                return Some((path, data));
            }
        }
        let path = fallback?;
        read(path).map(|d| (path.to_string(), d))
    };

    let core = read_part(rel_types::CORE_PROPERTIES, Some("docProps/core.xml"))
        .and_then(|(_, d)| CoreProperties::parse(&d).ok());
    let app = read_part(rel_types::EXTENDED_PROPERTIES, Some("docProps/app.xml"))
        .and_then(|(_, d)| AppProperties::parse(&d).ok());
    let custom = read_part(rel_types::CUSTOM_PROPERTIES, None)
        .map(|(path, d)| {
            parse_custom_properties(&d).unwrap_or_else(|e| {
                log::warn!("custom properties part '{path}' is malformed and was skipped: {e}");
                Vec::new()
            })
        })
        .unwrap_or_default();
    let thumbnail = read_part(rel_types::THUMBNAIL, None).map(|(path, data)| crate::ir::Image {
        format: image_format_from_path(&path),
        data: Some(data),
        ..Default::default()
    });
    PackageMetadata {
        core,
        app,
        package: PackageProperties {
            custom,
            has_digital_signature: rels
                .first_by_type(rel_types::DIGITAL_SIGNATURE_ORIGIN)
                .is_some(),
            thumbnail,
            crc_mismatched_parts: Vec::new(),
        },
    }
}

/// Read the custom properties, signature presence and thumbnail of an open
/// OPC package. See [`read_package_metadata`].
pub fn read_package_properties<R: std::io::Read + std::io::Seek>(
    opc: &mut super::opc::OpcReader<R>,
) -> PackageProperties {
    let rels = opc.package_rels().clone();
    read_package_metadata(&rels, |path| {
        let part = super::opc::PartName::new(&format!("/{path}")).ok()?;
        opc.has_part(&part).then(|| opc.read_part(&part).ok())?
    })
    .package
}

/// The image format a package thumbnail's file extension names.
fn image_format_from_path(path: &str) -> Option<crate::ir::ImageFormat> {
    use crate::ir::ImageFormat;
    let ext = path.rsplit_once('.')?.1.to_ascii_lowercase();
    Some(match ext.as_str() {
        "png" => ImageFormat::Png,
        "jpg" | "jpeg" | "jpe" => ImageFormat::Jpeg,
        "gif" => ImageFormat::Gif,
        "tif" | "tiff" => ImageFormat::Tiff,
        "bmp" => ImageFormat::Bmp,
        "emf" => ImageFormat::Emf,
        "wmf" => ImageFormat::Wmf,
        _ => return None,
    })
}

// ---------------------------------------------------------------------------
// Custom Properties — docProps/custom.xml (ECMA-376 Part 1 §22.3)
// ---------------------------------------------------------------------------

/// Namespace of `docProps/custom.xml` (ECMA-376 Part 1 §22.3).
const CUSTOM_PROPERTIES_NS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
/// Namespace of the variant types (ECMA-376 Part 1 §22.4).
const VT_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
/// The format id every custom property carries (ECMA-376 Part 1
/// §22.3.2.2: `fmtid` is `{D5CDD505-2E9C-101B-9397-08002B2CF9AE}` for
/// user-defined properties).
const CUSTOM_PROPERTY_FMTID: &str = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
/// The scalar variant types a custom property value can be written back
/// as (ECMA-376 Part 1 §22.4.2). Anything else is written as `lpwstr`.
const SCALAR_VARIANT_TYPES: &[&str] = &[
    "lpwstr", "lpstr", "bstr", "i1", "i2", "i4", "i8", "int", "ui1", "ui2", "ui4", "ui8", "uint",
    "r4", "r8", "decimal", "bool", "date", "filetime", "cy", "error", "clsid",
];

/// Parse `docProps/custom.xml`. Properties whose value is not a scalar
/// variant (a `vt:vector`, `vt:array`, blob or stream) are skipped with a
/// warning — their text alone cannot be written back faithfully.
pub fn parse_custom_properties(xml_data: &[u8]) -> Result<Vec<crate::ir::CustomProperty>> {
    let mut reader = property_reader(xml_data);
    let mut out = Vec::new();
    // (name, value type, value) of the property being read.
    let mut current: Option<(String, Option<String>, String)> = None;
    let mut depth_in_property = 0u32;
    loop {
        match reader.read_event()? {
            Event::Start(ref e) => {
                let local = e.local_name().as_ref().to_string();
                if current.is_none() && local == "property" {
                    let name = xml::optional_attr_str(e, "name")?
                        .map(|n| n.into_owned())
                        .unwrap_or_default();
                    current = Some((name, None, String::new()));
                    depth_in_property = 0;
                } else if let Some((_, ref mut ty, _)) = current {
                    depth_in_property += 1;
                    if depth_in_property == 1 {
                        *ty = Some(local);
                    }
                }
            },
            Event::Empty(ref e) => {
                if let Some((_, ref mut ty, _)) = current {
                    if depth_in_property == 0 {
                        *ty = Some(e.local_name().as_ref().to_string());
                    }
                }
            },
            Event::Text(ref e) => {
                if let Some((_, _, ref mut value)) = current {
                    if depth_in_property == 1 {
                        value.push_str(&xml::unescape_text(e)?);
                    }
                }
            },
            Event::GeneralRef(ref e) => {
                if let Some((_, _, ref mut value)) = current {
                    if depth_in_property == 1 {
                        value.push_str(&xml::resolve_general_ref(e)?);
                    }
                }
            },
            Event::End(ref e) => {
                if current.is_some() && depth_in_property > 0 {
                    depth_in_property -= 1;
                } else if e.local_name().as_ref() == "property" {
                    if let Some((name, ty, value)) = current.take() {
                        match ty {
                            Some(ty) if SCALAR_VARIANT_TYPES.contains(&ty.as_str()) => {
                                out.push(crate::ir::CustomProperty {
                                    name,
                                    value,
                                    value_type: ty,
                                });
                            },
                            ty => log::warn!(
                                "custom property '{name}' has a non-scalar value ({}); skipped",
                                ty.as_deref().unwrap_or("none")
                            ),
                        }
                    }
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(out)
}

/// Serialize custom properties to `docProps/custom.xml` bytes. Property
/// ids (`pid`) are assigned from 2 upward, as the spec requires
/// (ECMA-376 Part 1 §22.3.2.2).
pub fn serialize_custom_properties(props: &[crate::ir::CustomProperty]) -> Vec<u8> {
    let mut w = Writer::new(Vec::new());
    w.write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
        .expect("write decl");
    let mut root = BytesStart::new("Properties");
    root.push_attribute(("xmlns", CUSTOM_PROPERTIES_NS));
    root.push_attribute(("xmlns:vt", VT_NS));
    w.write_event(Event::Start(root)).expect("write root");
    for (i, p) in props.iter().enumerate() {
        let mut prop = BytesStart::new("property");
        prop.push_attribute(("fmtid", CUSTOM_PROPERTY_FMTID));
        prop.push_attribute(("pid", (i + 2).to_string().as_str()));
        prop.push_attribute(("name", xml::sanitize_xml_text(&p.name).as_ref()));
        w.write_event(Event::Start(prop)).expect("write property");
        let ty = if SCALAR_VARIANT_TYPES.contains(&p.value_type.as_str()) {
            p.value_type.as_str()
        } else {
            "lpwstr"
        };
        write_optional_element(&mut w, &format!("vt:{ty}"), Some(&p.value));
        w.write_event(Event::End(BytesEnd::new("property")))
            .expect("write property end");
    }
    w.write_event(Event::End(BytesEnd::new("Properties")))
        .expect("write root end");
    w.into_inner()
}

// ---------------------------------------------------------------------------
// App (Extended) Properties — docProps/app.xml
// ---------------------------------------------------------------------------

/// Extended/application properties from `docProps/app.xml`.
#[derive(Debug, Default, Clone, Serialize, Deserialize)]
pub struct AppProperties {
    /// Producing application name.
    pub application: Option<String>,
    /// Application version string.
    pub app_version: Option<String>,
    /// Company or organisation name.
    pub company: Option<String>,
    /// The author's manager (`<Manager>`).
    #[serde(default)]
    pub manager: Option<String>,
    /// Template the document is based on.
    pub template: Option<String>,
    /// Total editing time in minutes.
    pub total_time: Option<u32>,
    /// Number of pages.
    pub pages: Option<u32>,
    /// Word count.
    pub words: Option<u32>,
    /// Character count (excluding spaces).
    pub characters: Option<u32>,
    /// Character count (including spaces).
    pub characters_with_spaces: Option<u32>,
    /// Line count.
    pub lines: Option<u32>,
    /// Paragraph count.
    pub paragraphs: Option<u32>,
    /// Slide count (presentations).
    pub slides: Option<u32>,
    /// Notes-page count (presentations).
    pub notes: Option<u32>,
    /// Hidden slide count (presentations).
    pub hidden_slides: Option<u32>,
}

impl AppProperties {
    /// Parse app properties from `docProps/app.xml` bytes.
    pub fn parse(xml_data: &[u8]) -> Result<Self> {
        let mut reader = property_reader(xml_data);
        let mut props = AppProperties::default();
        // The element being read and its text so far. A value arrives as
        // several events when it holds an entity reference (`AT&amp;T` is
        // Text, GeneralRef, Text), so fragments accumulate and the value
        // is stored when the element ends.
        let mut current: Option<(String, String)> = None;

        loop {
            match reader.read_event()? {
                Event::Start(ref e) => {
                    current = Some((e.local_name().as_ref().to_string(), String::new()));
                },
                Event::Text(ref e) => {
                    if let Some((_, ref mut text)) = current {
                        text.push_str(&crate::core::xml::unescape_text(e)?);
                    }
                },
                Event::GeneralRef(ref e) => {
                    if let Some((_, ref mut text)) = current {
                        text.push_str(&crate::core::xml::resolve_general_ref(e)?);
                    }
                },
                Event::End(_) => {
                    if let Some((tag, text)) = current.take() {
                        props.set(&tag, text);
                    }
                },
                Event::Eof => break,
                _ => {},
            }
        }

        Ok(props)
    }

    /// Store the text of the `docProps/app.xml` element `tag`.
    fn set(&mut self, tag: &str, text: String) {
        let text_field = |v: String| {
            let v = v.trim();
            (!v.is_empty()).then(|| v.to_string())
        };
        match tag {
            "Application" => self.application = text_field(text),
            "AppVersion" => self.app_version = text_field(text),
            "Company" => self.company = text_field(text),
            "Manager" => self.manager = text_field(text),
            "Template" => self.template = text_field(text),
            "TotalTime" => self.total_time = text.trim().parse().ok(),
            "Pages" => self.pages = text.trim().parse().ok(),
            "Words" => self.words = text.trim().parse().ok(),
            "Characters" => self.characters = text.trim().parse().ok(),
            "CharactersWithSpaces" => self.characters_with_spaces = text.trim().parse().ok(),
            "Lines" => self.lines = text.trim().parse().ok(),
            "Paragraphs" => self.paragraphs = text.trim().parse().ok(),
            "Slides" => self.slides = text.trim().parse().ok(),
            "Notes" => self.notes = text.trim().parse().ok(),
            "HiddenSlides" => self.hidden_slides = text.trim().parse().ok(),
            _ => {},
        }
    }

    /// Serialize to `docProps/app.xml` bytes.
    pub fn serialize(&self) -> Vec<u8> {
        let mut w = Writer::new_with_indent(Vec::new(), b' ', 2);

        w.write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
            .expect("write decl");

        let mut root = BytesStart::new("Properties");
        root.push_attribute((
            "xmlns",
            "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
        ));
        root.push_attribute((
            "xmlns:vt",
            "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
        ));
        w.write_event(Event::Start(root)).expect("write root");

        write_optional_element(&mut w, "Application", self.application.as_deref());
        write_optional_element(&mut w, "AppVersion", self.app_version.as_deref());
        write_optional_element(&mut w, "Company", self.company.as_deref());
        write_optional_element(&mut w, "Manager", self.manager.as_deref());
        write_optional_element(&mut w, "Template", self.template.as_deref());
        write_optional_u32(&mut w, "TotalTime", self.total_time);
        write_optional_u32(&mut w, "Pages", self.pages);
        write_optional_u32(&mut w, "Words", self.words);
        write_optional_u32(&mut w, "Characters", self.characters);
        write_optional_u32(&mut w, "CharactersWithSpaces", self.characters_with_spaces);
        write_optional_u32(&mut w, "Lines", self.lines);
        write_optional_u32(&mut w, "Paragraphs", self.paragraphs);
        write_optional_u32(&mut w, "Slides", self.slides);
        write_optional_u32(&mut w, "Notes", self.notes);
        write_optional_u32(&mut w, "HiddenSlides", self.hidden_slides);

        w.write_event(Event::End(BytesEnd::new("Properties")))
            .expect("write end root");

        w.into_inner()
    }
}

fn write_optional_u32(w: &mut Writer<Vec<u8>>, tag: &str, value: Option<u32>) {
    if let Some(v) = value {
        let s = v.to_string();
        w.write_event(Event::Start(BytesStart::new(tag)))
            .expect("write start");
        w.write_event(Event::Text(BytesText::new(&crate::core::xml::sanitize_xml_text(&s))))
            .expect("write text");
        w.write_event(Event::End(BytesEnd::new(tag)))
            .expect("write end");
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const SAMPLE_CORE: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties
    xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:dcterms="http://purl.org/dc/terms/"
    xmlns:dcmitype="http://purl.org/dc/dcmitype/"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Quarterly Report</dc:title>
  <dc:subject>Q4 2024 Financial Summary</dc:subject>
  <dc:creator>Jane Smith</dc:creator>
  <cp:keywords>finance; quarterly; report</cp:keywords>
  <cp:category>Report</cp:category>
  <cp:lastModifiedBy>John Doe</cp:lastModifiedBy>
  <cp:revision>4</cp:revision>
  <cp:contentStatus>Final</cp:contentStatus>
  <dcterms:created xsi:type="dcterms:W3CDTF">2024-10-01T09:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2024-12-22T16:45:00Z</dcterms:modified>
</cp:coreProperties>"#;

    #[test]
    fn test_parse_core_properties() {
        let props = CoreProperties::parse(SAMPLE_CORE).unwrap();
        assert_eq!(props.title.as_deref(), Some("Quarterly Report"));
        assert_eq!(props.creator.as_deref(), Some("Jane Smith"));
        assert_eq!(props.keywords.as_deref(), Some("finance; quarterly; report"));
        assert_eq!(props.revision.as_deref(), Some("4"));
        assert_eq!(props.created.as_deref(), Some("2024-10-01T09:00:00Z"));
        assert_eq!(props.content_status.as_deref(), Some("Final"));
    }

    #[test]
    fn test_core_properties_round_trip() {
        let original = CoreProperties {
            title: Some("Test Doc".to_string()),
            creator: Some("Test Author".to_string()),
            created: Some("2024-01-01T00:00:00Z".to_string()),
            ..Default::default()
        };
        let xml = original.serialize();
        let parsed = CoreProperties::parse(&xml).unwrap();
        assert_eq!(parsed.title, original.title);
        assert_eq!(parsed.creator, original.creator);
        assert_eq!(parsed.created, original.created);
    }

    #[test]
    fn test_space_padded_single_digit_date_is_normalized_not_dropped() {
        // A real-world PHP writer's quirk: single-digit
        // month/day written with a leading space instead of zero-padding.
        // The intent is unambiguous, so this must be recovered, not
        // dropped.
        let props = CoreProperties {
            modified: Some("2021- 9- 3T20:25:22Z".to_string()),
            ..Default::default()
        };
        let xml = props.serialize();
        let xml_str = String::from_utf8(xml.clone()).unwrap();
        assert!(
            xml_str.contains("2021-09-03T20:25:22Z"),
            "expected the normalized value, got: {xml_str}"
        );
        let parsed = CoreProperties::parse(&xml).unwrap();
        assert_eq!(parsed.modified.as_deref(), Some("2021-09-03T20:25:22Z"));
    }

    #[test]
    fn test_unrecoverably_malformed_date_is_dropped_not_copied_verbatim() {
        // Garbage mixed into a numeric component (a
        // deliberately corrupt OpenXML SDK test fixture) has no
        // unambiguous recovery; the invalid element must be omitted
        // entirely rather than authoring an invalid docProps/core.xml.
        let props = CoreProperties {
            modified: Some("2015sss-06-20T07:40:00Z".to_string()),
            ..Default::default()
        };
        let xml = props.serialize();
        let xml_str = String::from_utf8(xml).unwrap();
        assert!(
            !xml_str.contains("dcterms:modified"),
            "an unrecoverable date must be dropped, not written: {xml_str}"
        );
    }

    const SAMPLE_APP: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Office Word</Application>
  <AppVersion>16.0000</AppVersion>
  <Company>Acme Corp</Company>
  <Template>Normal.dotm</Template>
  <TotalTime>45</TotalTime>
  <Pages>3</Pages>
  <Words>1250</Words>
  <Characters>7125</Characters>
  <Lines>62</Lines>
  <Paragraphs>17</Paragraphs>
</Properties>"#;

    #[test]
    fn test_parse_app_properties() {
        let props = AppProperties::parse(SAMPLE_APP).unwrap();
        assert_eq!(props.application.as_deref(), Some("Microsoft Office Word"));
        assert_eq!(props.app_version.as_deref(), Some("16.0000"));
        assert_eq!(props.company.as_deref(), Some("Acme Corp"));
        assert_eq!(props.pages, Some(3));
        assert_eq!(props.words, Some(1250));
        assert_eq!(props.lines, Some(62));
    }

    /// Entity references arrive as their own event; `AT&amp;T` used to
    /// lose its `&` (and anything after the first fragment).
    #[test]
    fn test_app_properties_keep_entity_references() {
        let xml = br#"<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Company>AT&amp;T &lt;Labs&gt;</Company>
  <Manager>Smith &#38; Co</Manager>
  <HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Title</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs>
  <Template>A&amp;B.dotm</Template>
  <Pages> 3 </Pages>
</Properties>"#;
        let props = AppProperties::parse(xml).unwrap();
        assert_eq!(props.company.as_deref(), Some("AT&T <Labs>"));
        assert_eq!(props.manager.as_deref(), Some("Smith & Co"));
        assert_eq!(props.template.as_deref(), Some("A&B.dotm"));
        assert_eq!(props.pages, Some(3));
    }

    /// The spaces on either side of an entity reference are part of the
    /// value; per-fragment trimming deleted them.
    #[test]
    fn test_core_properties_keep_spaces_around_entity_references() {
        let xml = br#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>  Smith &amp; Co Report </dc:title><dc:creator>   </dc:creator></cp:coreProperties>"#;
        let props = CoreProperties::parse(xml).unwrap();
        assert_eq!(props.title.as_deref(), Some("Smith & Co Report"));
        assert_eq!(props.creator, None);
    }

    #[test]
    fn test_app_properties_round_trip() {
        let original = AppProperties {
            application: Some("office_oxide".to_string()),
            app_version: Some("0.1.0".to_string()),
            pages: Some(5),
            words: Some(2000),
            ..Default::default()
        };
        let xml = original.serialize();
        let parsed = AppProperties::parse(&xml).unwrap();
        assert_eq!(parsed.application, original.application);
        assert_eq!(parsed.pages, original.pages);
        assert_eq!(parsed.words, original.words);
    }
}
