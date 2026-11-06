//! Shape text in `metroBlob` packages inside legacy `.ppt` streams.
//!
//! A shape can carry an alternative DrawingML form of itself: the tertiary
//! property table (`0xF122`, `OfficeArtTertiaryFOPT`, [MS-ODRAW] 2.2.11) holds
//! a `metroBlob` property ([MS-ODRAW] 2.3.4.41) whose complex data
//! ([MS-ODRAW] 2.3.4.42) is an OPC package, with the shape's `<p:sp>` fragment
//! in its `drs/shapexml.xml` part. PowerPoint writes these beside full legacy
//! text. WPS, saving an OOXML-origin deck, can write them with the legacy text
//! atoms left nearly empty, so a reader of the atoms alone returns an
//! almost-empty deck with `Ok`.
//!
//! A package is therefore used only for a shape with no legacy text of its
//! own: otherwise it is the same text a second time. Everything read here is
//! bounded per file and per slide, every package attempted is charged whether
//! or not it can be read, and a limit or a damaged package leaves a notice.

use std::collections::HashSet;
use std::hash::{BuildHasher, RandomState};
use std::io::Read;

use quick_xml::events::{BytesStart, Event};
use zip::result::ZipError;

use super::records::{
    RT_OUTLINE_TEXT_REF_ATOM, RT_SHAPE, RT_TEXT_BYTES, RT_TEXT_CHARS, RecordIter,
};
use super::text::{
    TextRun, TextType, decode_text_bytes, decode_utf16le, is_field_placeholder_only,
};

/// How deep into a slide's record tree to look for carrier records.
/// Mirrors the shape-nesting guard in [`super::text`].
const MAX_CARRIER_DEPTH: usize = 64;

/// How far into a malformed carrier's payload to scan for a ZIP signature.
const MAX_SIGNATURE_SCAN: usize = 4096;

/// Cap on one shape fragment's decompressed size. The largest of 1,039 real
/// fragments was about 16 KB.
const MAX_SHAPE_XML_BYTES: u64 = 64 * 1024;

/// Cap on packages attempted per slide.
const MAX_DOWNREV_PACKAGES: usize = 256;

/// Cap on packages attempted per file. Each attempt parses a ZIP directory.
const DOWNREV_MAX_PACKAGES_PER_FILE: usize = 8192;

/// Cap on decompressed bytes per file, charged for every byte inflated whether
/// or not it turns out to be usable. Real fragments average about 2.8 KB, so
/// this is room for over 10,000 shapes.
const DOWNREV_MAX_BYTES_PER_FILE: u64 = 32 * 1024 * 1024;

/// The ZIP entry holding a shape's DrawingML fragment.
const SHAPE_XML_ENTRY: &str = "drs/shapexml.xml";

/// `OfficeArtTertiaryFOPT` record type ([MS-ODRAW] 2.2.11). Its `recInstance`
/// is the number of properties in the table.
const RT_TERTIARY_FOPT: u16 = 0xF122;

/// `OfficeArtFSP` record type ([MS-ODRAW] 2.2.40): a shape's id and flags.
const RT_FSP: u16 = 0xF00A;

/// `fDeleted`, in the flags that follow `OfficeArtFSP`'s 4-byte `spid`.
const FSP_FDELETED: u32 = 0x0000_0008;

/// `metroBlob` property id ([MS-ODRAW] 2.3.4.41): the low 14 bits of the
/// opid word.
const METROBLOB_OPID: u16 = 0x03A9;

/// `fComplex` bit of an `OfficeArtFOPTEOPID` ([MS-ODRAW] 2.2.8).
const OPID_FCOMPLEX: u16 = 0x8000;

/// Mask that keeps only the property id, dropping `fComplex` and `fBid`.
const OPID_ID_MASK: u16 = 0x3FFF;

/// Size of one `OfficeArtFOPTE` ([MS-ODRAW] 2.2.7).
const FOPTE_BYTES: usize = 6;

/// What one file may spend on packages, and what it could not read.
pub(super) struct DownrevBudget {
    /// Packages still allowed across the whole file.
    packages_remaining: usize,
    /// Decompressed shape-XML bytes still allowed across the whole file.
    bytes_remaining: u64,
    /// Packages still allowed on the current slide (reset per slide).
    slide_packages_remaining: usize,
    /// Hash of every package already paid for on the *current* slide: real
    /// files repeat a shape's package across records, and each slide that
    /// carries one keeps its own copy. The keyed hasher means a crafted file
    /// cannot aim a collision at a package to hide its text.
    seen_packages: HashSet<u64>,
    hasher: RandomState,
    /// A limit stopped text from being read.
    limit_reached: bool,
    /// A package that should have held text could not be read.
    unreadable: bool,
    /// Runs added so far, across the file.
    runs_added: usize,
}

impl DownrevBudget {
    /// A fresh, full budget for one file.
    pub(super) fn new() -> Self {
        DownrevBudget {
            packages_remaining: DOWNREV_MAX_PACKAGES_PER_FILE,
            bytes_remaining: DOWNREV_MAX_BYTES_PER_FILE,
            slide_packages_remaining: MAX_DOWNREV_PACKAGES,
            seen_packages: HashSet::new(),
            hasher: RandomState::new(),
            limit_reached: false,
            unreadable: false,
            runs_added: 0,
        }
    }

    /// Runs added so far, across the file.
    pub(super) fn runs_added(&self) -> usize {
        self.runs_added
    }

    /// What the reader should be told about text it did not get.
    pub(super) fn notices(&self) -> Vec<String> {
        let mut notices = Vec::new();
        if self.limit_reached {
            notices.push(
                "Some shape text in embedded metroBlob packages was not read: a package limit \
                 was reached."
                    .to_string(),
            );
        }
        if self.unreadable {
            notices.push(
                "Some shape text in embedded metroBlob packages was not read: a package is \
                 damaged."
                    .to_string(),
            );
        }
        notices
    }

    fn exhausted(&self) -> bool {
        self.packages_remaining == 0
            || self.slide_packages_remaining == 0
            || self.bytes_remaining == 0
    }
}

/// Add the text of the packages in `slide_children` (a `Slide` container's
/// children) to `runs`, for each shape with no legacy text of its own.
/// `outline_texts` resolves a shape's `OutlineTextRefAtom`, as it does for the
/// legacy walk; pass `&[]` where that walk has none.
///
/// A title takes its place ahead of the slide's other text. The rest follows
/// the legacy runs, in document order.
pub(super) fn append_slide_downrev(
    slide_children: &[u8],
    outline_texts: &[TextRun],
    runs: &mut Vec<TextRun>,
    budget: &mut DownrevBudget,
    text_budget: &mut crate::limits::TextBudget,
) {
    budget.slide_packages_remaining = MAX_DOWNREV_PACKAGES;
    budget.seen_packages.clear();

    let mut fragments = Vec::new();
    walk_records(slide_children, 0, outline_texts, &mut fragments, budget);
    for fragment in fragments {
        let is_title = matches!(fragment.role, TextType::Title | TextType::CenterTitle);
        let mut texts: Vec<String> = fragment
            .paragraphs
            .into_iter()
            .filter(|p| !is_placeholder_mask(p))
            .collect();
        if is_title && !texts.is_empty() {
            texts = vec![texts.join("\r")];
        }
        for text in texts {
            if !text_budget.charge(text.len()) {
                return;
            }
            let run = TextRun {
                text_type: fragment.role,
                text,
                ..Default::default()
            };
            budget.runs_added += 1;
            if is_title {
                let at = runs
                    .iter()
                    .take_while(|r| matches!(r.text_type, TextType::Title | TextType::CenterTitle))
                    .count();
                runs.insert(at, run);
            } else {
                runs.push(run);
            }
        }
    }
}

fn walk_records(
    data: &[u8],
    depth: usize,
    outline_texts: &[TextRun],
    fragments: &mut Vec<ShapeFragment>,
    budget: &mut DownrevBudget,
) {
    if depth > MAX_CARRIER_DEPTH {
        return;
    }
    for rec in RecordIter::new(data) {
        let Ok(rec) = rec else { break };
        if rec.header.rec_type == RT_SHAPE
            && (shape_is_deleted(&rec.data) || shape_has_legacy_text(&rec.data, 0, outline_texts))
        {
            continue;
        }
        if rec.header.is_container() {
            walk_records(&rec.data, depth + 1, outline_texts, fragments, budget);
        } else if rec.header.rec_type == RT_TERTIARY_FOPT {
            let properties = usize::from(rec.header.rec_instance);
            if let Some(fragment) = read_carrier(&rec.data, properties, budget) {
                fragments.push(fragment);
            }
            if budget.limit_reached && budget.exhausted() {
                return;
            }
        }
    }
}

/// True when a shape's `OfficeArtFSP` marks it deleted: it isn't on the slide.
fn shape_is_deleted(shape: &[u8]) -> bool {
    RecordIter::new(shape)
        .map_while(|r| r.ok())
        .find(|rec| rec.header.rec_type == RT_FSP)
        .and_then(|rec| {
            let flags = rec.data.get(4..8)?;
            Some(u32::from_le_bytes([flags[0], flags[1], flags[2], flags[3]]))
        })
        .is_some_and(|flags| flags & FSP_FDELETED != 0)
}

/// True when a shape's own text is in the legacy atoms: stored in the shape,
/// or by reference into the slide's outline text. A stub of whitespace or a
/// field marker does not count: that is what the packages are there to replace.
fn shape_has_legacy_text(shape: &[u8], depth: usize, outline_texts: &[TextRun]) -> bool {
    if depth > MAX_CARRIER_DEPTH {
        return false;
    }
    let real = |text: &str| !text.trim().is_empty() && !is_field_placeholder_only(text);
    RecordIter::new(shape).map_while(|r| r.ok()).any(|rec| {
        if rec.header.is_container() {
            return shape_has_legacy_text(&rec.data, depth + 1, outline_texts);
        }
        match rec.header.rec_type {
            RT_TEXT_CHARS => real(&decode_utf16le(&rec.data)),
            RT_TEXT_BYTES => real(&decode_text_bytes(&rec.data)),
            RT_OUTLINE_TEXT_REF_ATOM if rec.data.len() >= 4 => {
                let index =
                    i32::from_le_bytes([rec.data[0], rec.data[1], rec.data[2], rec.data[3]]);
                usize::try_from(index)
                    .ok()
                    .and_then(|i| outline_texts.get(i))
                    .is_some_and(|run| real(&run.text))
            },
            _ => false,
        }
    })
}

/// A carrier's property table, read per `OfficeArtRGFOPTE` ([MS-ODRAW] 2.3.1):
/// `properties` entries, then the complex data of those with `fComplex` set,
/// in entry order.
enum PropertyTable<'a> {
    /// The table holds a `metroBlob`, with this complex data.
    Package(&'a [u8]),
    /// The table is sound and accounts for the whole payload: no package.
    Nothing,
    /// The table does not fit, or has no `metroBlob` and leaves bytes it does
    /// not explain.
    Unexplained,
}

fn read_property_table(payload: &[u8], properties: usize) -> PropertyTable<'_> {
    let Some(entries) = properties
        .checked_mul(FOPTE_BYTES)
        .and_then(|len| payload.get(..len))
    else {
        return PropertyTable::Unexplained;
    };
    let mut data_at = entries.len();
    let mut package = None;
    for entry in entries.as_chunks::<FOPTE_BYTES>().0 {
        let opid = u16::from_le_bytes([entry[0], entry[1]]);
        if opid & OPID_FCOMPLEX == 0 {
            continue;
        }
        let len = u32::from_le_bytes([entry[2], entry[3], entry[4], entry[5]]) as usize;
        let Some(end) = data_at.checked_add(len).filter(|&e| e <= payload.len()) else {
            return PropertyTable::Unexplained;
        };
        if opid & OPID_ID_MASK == METROBLOB_OPID && package.is_none() {
            package = Some(&payload[data_at..end]);
        }
        data_at = end;
    }
    match package {
        Some(package) => PropertyTable::Package(package),
        None if data_at == payload.len() => PropertyTable::Nothing,
        None => PropertyTable::Unexplained,
    }
}

/// Read the package a carrier holds. A table that does not explain its payload
/// is not trusted: it is searched for a ZIP signature instead.
fn read_carrier(
    payload: &[u8],
    properties: usize,
    budget: &mut DownrevBudget,
) -> Option<ShapeFragment> {
    let (package, declared) = match read_property_table(payload, properties) {
        PropertyTable::Package(package) => (package, true),
        PropertyTable::Nothing => return None,
        PropertyTable::Unexplained => (&payload[find_zip_signature(payload)?..], false),
    };
    if package.is_empty() {
        return None;
    }
    // An identical package on this slide has been paid for already.
    if !budget.seen_packages.insert(budget.hasher.hash_one(package)) {
        return None;
    }
    if budget.exhausted() {
        budget.limit_reached = true;
        return None;
    }
    budget.packages_remaining -= 1;
    budget.slide_packages_remaining -= 1;
    read_shape_xml(package, declared, budget)
}

/// The fragment in a package. Charges what it inflates, whatever comes of it.
/// A fragment cut by a limit keeps the paragraphs it holds in full.
fn read_shape_xml(
    package: &[u8],
    declared: bool,
    budget: &mut DownrevBudget,
) -> Option<ShapeFragment> {
    let Ok(mut archive) = zip::ZipArchive::new(std::io::Cursor::new(package)) else {
        // Only a declared package is known to be one; bytes found by the
        // signature scan may be anything.
        budget.unreadable |= declared;
        return None;
    };
    let entry = match archive.by_name(SHAPE_XML_ENTRY) {
        Ok(entry) => entry,
        // Pictures, connectors and groups carry other parts.
        Err(ZipError::FileNotFound) => return None,
        Err(_) => {
            budget.unreadable = true;
            return None;
        },
    };
    let limit = MAX_SHAPE_XML_BYTES.min(budget.bytes_remaining);
    let mut bytes = Vec::new();
    let read = entry.take(limit + 1).read_to_end(&mut bytes);
    budget.bytes_remaining -= (bytes.len() as u64).min(limit);
    if read.is_err() {
        budget.unreadable = true;
        return None;
    }
    let truncated = bytes.len() as u64 > limit;
    if truncated {
        bytes.truncate(limit as usize);
        budget.limit_reached = true;
    }
    let xml = match std::str::from_utf8(&bytes) {
        Ok(xml) => xml,
        // A cut can land inside a character; anything else is damage.
        Err(e) if truncated && bytes.len() - e.valid_up_to() < 4 => {
            std::str::from_utf8(&bytes[..e.valid_up_to()]).ok()?
        },
        Err(_) => {
            budget.unreadable = true;
            return None;
        },
    };
    let fragment = parse_shape_fragment(xml, truncated);
    budget.unreadable |= fragment.malformed;
    (!fragment.skip).then_some(fragment)
}

fn find_zip_signature(payload: &[u8]) -> Option<usize> {
    let limit = payload.len().min(MAX_SIGNATURE_SCAN);
    payload[..limit].windows(4).position(|w| w == b"PK\x03\x04")
}

/// True for text that holds nothing to read: the runs of underscores real
/// files substitute for letters, and the slide-number marker `‹#›`.
fn is_placeholder_mask(s: &str) -> bool {
    s.chars()
        .all(|c| c == '_' || c == '‹' || c == '#' || c == '›' || c.is_whitespace())
}

/// One shape fragment's text.
#[derive(Debug)]
struct ShapeFragment {
    /// The placeholder's text type, from `<p:ph type>`.
    role: TextType,
    paragraphs: Vec<String>,
    /// A slide-number placeholder: its text is a cached field value.
    skip: bool,
    /// The XML is damaged, and ended before a limit cut it.
    malformed: bool,
}

/// `<a:t>` runs grouped per `<a:p>` (a paragraph's text is split across runs
/// arbitrarily), with `<a:br>` soft returns kept as line breaks. Entities and
/// CDATA resolve through the shared OOXML text reader.
///
/// The reader deliberately does not trim text events (unlike
/// [`crate::core::xml::make_fast_reader`]): `"Hello "` + `"world"` must not
/// collapse to `"Helloworld"`. Slide-number and date field values are left out:
/// they are cached results, not the slide's text. With `truncated` the XML was
/// cut by a limit, so the paragraph it ends in is left out: it isn't all there.
fn parse_shape_fragment(xml: &str, truncated: bool) -> ShapeFragment {
    use crate::core::xml::{optional_attr_str, read_text_content_fast};

    let attr = |e: &BytesStart, key: &str| {
        optional_attr_str(e, key)
            .ok()
            .flatten()
            .map(|v| v.into_owned())
    };
    let mut reader = quick_xml::Reader::from_reader(xml.as_bytes());
    reader.config_mut().check_end_names = false;
    reader.config_mut().check_comments = false;
    let mut fragment = ShapeFragment {
        role: TextType::Other,
        paragraphs: Vec::new(),
        skip: false,
        malformed: false,
    };
    let mut current = String::new();
    let mut in_volatile_field = false;
    let placeholder =
        |e: &BytesStart, fragment: &mut ShapeFragment| match attr(e, "type").as_deref() {
            Some("title") => fragment.role = TextType::Title,
            Some("ctrTitle") => fragment.role = TextType::CenterTitle,
            Some("subTitle") => fragment.role = TextType::CenterBody,
            Some("body") | None => fragment.role = TextType::Body,
            Some("sldNum") => fragment.skip = true,
            Some(_) => {},
        };
    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) => match e.local_name().as_ref() {
                "p" if !current.is_empty() => {
                    fragment.paragraphs.push(std::mem::take(&mut current));
                },
                "ph" => placeholder(&e, &mut fragment),
                "br" => current.push('\u{b}'),
                "fld" => {
                    let kind = attr(&e, "type").unwrap_or_default();
                    in_volatile_field =
                        kind.starts_with("slidenum") || kind.starts_with("datetime");
                },
                "t" => match read_text_content_fast(&mut reader) {
                    Ok(t) if !in_volatile_field => current.push_str(&t),
                    Ok(_) => {},
                    Err(_) => {
                        fragment.malformed = !truncated;
                        break;
                    },
                },
                _ => {},
            },
            Ok(Event::Empty(e)) => match e.local_name().as_ref() {
                "br" => current.push('\u{b}'),
                "ph" => placeholder(&e, &mut fragment),
                _ => {},
            },
            Ok(Event::End(e)) => match e.local_name().as_ref() {
                "p" if !current.is_empty() => {
                    fragment.paragraphs.push(std::mem::take(&mut current));
                },
                "fld" => in_volatile_field = false,
                _ => {},
            },
            Ok(Event::Eof) => break,
            Err(_) => {
                fragment.malformed = !truncated;
                break;
            },
            _ => {},
        }
    }
    if !truncated && !current.is_empty() {
        fragment.paragraphs.push(current);
    }
    fragment
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::limits::TextBudget;

    const METROBLOB_COMPLEX: u16 = OPID_FCOMPLEX | METROBLOB_OPID;
    /// The same id with `fBid` set too: most real files write this.
    const METROBLOB_COMPLEX_BID: u16 = 0xC3A9;
    /// A complex property that isn't `metroBlob`: `pibName`.
    const OTHER_COMPLEX: u16 = OPID_FCOMPLEX | 0x0105;
    /// A simple property real tables put after `metroBlob`: the Group Shape
    /// Boolean properties word.
    const TRAILING_SIMPLE: u16 = 0x03BF;
    const SP_CONTAINER: u16 = RT_SHAPE;
    const CLIENT_TEXTBOX: u16 = 0xF00D;

    const SAMPLE_XML: &str = r#"<p:sp><a:p><a:r><a:t>Real package text</a:t></a:r></a:p></p:sp>"#;

    fn sp(text: &str) -> String {
        format!("<p:sp><a:p><a:r><a:t>{text}</a:t></a:r></a:p></p:sp>")
    }

    fn zip_with(entry: &str, data: &[u8]) -> Vec<u8> {
        use std::io::Write;
        let mut buf = Vec::new();
        {
            let mut w = zip::ZipWriter::new(std::io::Cursor::new(&mut buf));
            let opts = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Deflated);
            w.start_file(entry, opts).expect("zip entry");
            w.write_all(data).expect("zip payload");
            w.finish().expect("zip finish");
        }
        buf
    }

    fn package(xml: &str) -> Vec<u8> {
        zip_with(SHAPE_XML_ENTRY, xml.as_bytes())
    }

    fn record_inst(rec_type: u16, instance: u16, ver: u8, payload: &[u8]) -> Vec<u8> {
        let ver_instance = ((instance & 0x0FFF) << 4) | (ver & 0x0F) as u16;
        let mut v = Vec::with_capacity(8 + payload.len());
        v.extend_from_slice(&ver_instance.to_le_bytes());
        v.extend_from_slice(&rec_type.to_le_bytes());
        v.extend_from_slice(&(payload.len() as u32).to_le_bytes());
        v.extend_from_slice(payload);
        v
    }

    fn atom(rec_type: u16, payload: &[u8]) -> Vec<u8> {
        record_inst(rec_type, 0, 0, payload)
    }

    fn container(rec_type: u16, children: &[Vec<u8>]) -> Vec<u8> {
        record_inst(rec_type, 0, 0x0F, &children.concat())
    }

    /// A payload per `OfficeArtRGFOPTE`: every entry, then the complex data in
    /// entry order. A `None` property is simple, with value 0.
    fn table_payload(props: &[(u16, Option<&[u8]>)]) -> Vec<u8> {
        let mut entries = Vec::new();
        let mut data = Vec::new();
        for (opid, complex) in props {
            entries.extend_from_slice(&opid.to_le_bytes());
            let value = complex.map_or(0, |c| c.len() as u32);
            entries.extend_from_slice(&value.to_le_bytes());
            data.extend_from_slice(complex.unwrap_or_default());
        }
        entries.extend(data);
        entries
    }

    fn table_record(props: &[(u16, Option<&[u8]>)]) -> Vec<u8> {
        record_inst(RT_TERTIARY_FOPT, props.len() as u16, 3, &table_payload(props))
    }

    fn carrier_record(xml: &str) -> Vec<u8> {
        table_record(&[(METROBLOB_COMPLEX, Some(&package(xml)))])
    }

    fn slide(children: &[Vec<u8>]) -> Vec<u8> {
        children.concat()
    }

    fn run_slide(children: &[Vec<u8>], budget: &mut DownrevBudget) -> Vec<TextRun> {
        let mut runs = Vec::new();
        let mut text = TextBudget::with_limit(usize::MAX);
        append_slide_downrev(&slide(children), &[], &mut runs, budget, &mut text);
        runs
    }

    fn texts(runs: &[TextRun]) -> Vec<&str> {
        runs.iter().map(|r| r.text.as_str()).collect()
    }

    fn legacy_shape(atom_rec: Vec<u8>, carrier: Vec<u8>) -> Vec<u8> {
        container(SP_CONTAINER, &[container(CLIENT_TEXTBOX, &[atom_rec]), carrier])
    }

    fn utf16(text: &str) -> Vec<u8> {
        text.encode_utf16().flat_map(u16::to_le_bytes).collect()
    }

    fn fragment_paragraphs(xml: &str) -> Vec<String> {
        parse_shape_fragment(xml, false).paragraphs
    }

    // Property table layout

    /// Every layout real tables use, including those with a property after
    /// `metroBlob`: its data follows all the entries, not its own entry.
    #[test]
    fn test_package_is_found_at_any_position_in_the_table() {
        let pkg = package(SAMPLE_XML);
        let cases: Vec<(&str, Vec<(u16, Option<&[u8]>)>)> = vec![
            ("alone", vec![(METROBLOB_COMPLEX, Some(&pkg))]),
            ("fBid set", vec![(METROBLOB_COMPLEX_BID, Some(&pkg))]),
            ("after simple", vec![(0x01FF, None), (METROBLOB_COMPLEX, Some(&pkg))]),
            (
                "before simple",
                vec![(METROBLOB_COMPLEX_BID, Some(&pkg)), (TRAILING_SIMPLE, None)],
            ),
            (
                "between simple",
                vec![
                    (0x01FF, None),
                    (METROBLOB_COMPLEX_BID, Some(&pkg)),
                    (TRAILING_SIMPLE, None),
                ],
            ),
            (
                "after complex",
                vec![
                    (OTHER_COMPLEX, Some(b"d\0")),
                    (METROBLOB_COMPLEX, Some(&pkg)),
                ],
            ),
            (
                "before complex",
                vec![
                    (METROBLOB_COMPLEX, Some(&pkg)),
                    (OTHER_COMPLEX, Some(b"d\0")),
                ],
            ),
        ];
        for (name, props) in cases {
            let mut budget = DownrevBudget::new();
            let runs = run_slide(&[table_record(&props)], &mut budget);
            assert_eq!(texts(&runs), ["Real package text"], "{name}");
            assert!(budget.notices().is_empty(), "{name}: {:?}", budget.notices());
        }
    }

    /// A ZIP in another complex property isn't this shape's package, even
    /// though a signature scan would find it first.
    #[test]
    fn test_zip_in_another_complex_property_is_not_read() {
        let decoy = package(&sp("Decoy package text"));
        let real = package(SAMPLE_XML);
        let mut budget = DownrevBudget::new();
        let runs = run_slide(
            &[table_record(&[
                (OTHER_COMPLEX, Some(&decoy)),
                (METROBLOB_COMPLEX, Some(&real)),
            ])],
            &mut budget,
        );
        assert_eq!(texts(&runs), ["Real package text"]);

        let runs = run_slide(&[table_record(&[(OTHER_COMPLEX, Some(&decoy))])], &mut budget);
        assert!(runs.is_empty(), "a sound table with no metroBlob has no package: {runs:?}");
    }

    /// The declared length bounds the read: a short one isn't carved around,
    /// and the damage is reported.
    #[test]
    fn test_short_declared_length_bounds_the_read_and_warns() {
        let pkg = package(SAMPLE_XML);
        let mut payload = Vec::new();
        payload.extend_from_slice(&METROBLOB_COMPLEX.to_le_bytes());
        payload.extend_from_slice(&((pkg.len() - 8) as u32).to_le_bytes());
        payload.extend_from_slice(&pkg);
        let mut budget = DownrevBudget::new();
        let runs = run_slide(&[record_inst(RT_TERTIARY_FOPT, 1, 3, &payload)], &mut budget);
        assert!(runs.is_empty(), "{runs:?}");
        assert_eq!(budget.notices().len(), 1, "{:?}", budget.notices());
    }

    /// A table that doesn't explain its payload falls back to the signature
    /// scan, without panicking on any of these.
    #[test]
    fn test_unexplained_table_falls_back_to_the_signature_scan() {
        let pkg = package(SAMPLE_XML);
        let cases: Vec<(&str, u16, Vec<u8>)> = vec![
            ("not a table", 1, [vec![0u8; 6], pkg.clone()].concat()),
            ("over-long length", 1, {
                let mut p = METROBLOB_COMPLEX.to_le_bytes().to_vec();
                p.extend_from_slice(&((pkg.len() + 64) as u32).to_le_bytes());
                p.extend_from_slice(&pkg);
                p
            }),
            ("huge length", 1, {
                let mut p = METROBLOB_COMPLEX.to_le_bytes().to_vec();
                p.extend_from_slice(&u32::MAX.to_le_bytes());
                p.extend_from_slice(&pkg);
                p
            }),
            ("no complex bit", 1, {
                let mut p = METROBLOB_OPID.to_le_bytes().to_vec();
                p.extend_from_slice(&0u32.to_le_bytes());
                p.extend_from_slice(&pkg);
                p
            }),
            ("count past the payload", 0xFFF, [vec![0u8; 6], pkg.clone()].concat()),
        ];
        for (name, instance, payload) in cases {
            let mut budget = DownrevBudget::new();
            let runs =
                run_slide(&[record_inst(RT_TERTIARY_FOPT, instance, 3, &payload)], &mut budget);
            assert_eq!(texts(&runs), ["Real package text"], "{name}");
        }
    }

    /// Only the tertiary property table is a carrier: a ZIP in any other
    /// record isn't read, however it is laid out.
    #[test]
    fn test_zip_in_a_record_that_is_not_a_carrier_is_not_read() {
        let payload = table_payload(&[(METROBLOB_COMPLEX, Some(&package(SAMPLE_XML)))]);
        for rec_type in [0x03EF, 0x1388, 0xF121, 0xF00B] {
            let mut budget = DownrevBudget::new();
            let runs = run_slide(&[record_inst(rec_type, 1, 3, &payload)], &mut budget);
            assert!(runs.is_empty(), "{rec_type:#06x}: {runs:?}");
        }
    }

    #[test]
    fn test_malformed_and_foreign_carriers_add_nothing() {
        let mut garbage = b"PK\x03\x04garbage".to_vec();
        garbage.extend_from_slice(&[0xFF; 32]);
        let foreign = zip_with("theme/theme1.xml", b"<a:theme/>");
        let mut budget = DownrevBudget::new();
        let runs = run_slide(
            &[
                record_inst(RT_TERTIARY_FOPT, 3, 3, &garbage),
                table_record(&[(METROBLOB_COMPLEX, Some(&foreign))]),
                table_record(&[(METROBLOB_COMPLEX, Some(&[]))]),
                record_inst(RT_TERTIARY_FOPT, 0, 3, &[]),
            ],
            &mut budget,
        );
        assert!(runs.is_empty(), "{runs:?}");
        assert!(budget.notices().is_empty(), "other kinds of package: {:?}", budget.notices());
    }

    // Package or legacy text, per shape

    /// A package is the same shape's text again when the shape has its own, in
    /// whatever form: soft returns and spacing differ, so no text comparison
    /// could tell.
    #[test]
    fn test_shape_with_legacy_text_ignores_its_package() {
        let cases = [
            (
                "Second\u{b}line",
                "<a:p><a:r><a:t>Second</a:t></a:r><a:br/><a:r><a:t>line</a:t></a:r></a:p>",
            ),
            ("Hello", "<a:p><a:r><a:t>Hello </a:t></a:r></a:p>"),
            ("A B", "<a:p><a:r><a:t>A  B</a:t></a:r></a:p>"),
            ("Same", "<a:p><a:r><a:t>Different words</a:t></a:r></a:p>"),
        ];
        for (legacy, body) in cases {
            let xml = format!("<p:sp>{body}</p:sp>");
            for legacy_atom in [
                atom(RT_TEXT_CHARS, &utf16(legacy)),
                atom(RT_TEXT_BYTES, legacy.as_bytes()),
            ] {
                let shape = legacy_shape(legacy_atom, carrier_record(&xml));
                let runs = run_slide(&[shape], &mut DownrevBudget::new());
                assert!(runs.is_empty(), "{legacy:?}: {runs:?}");
            }
        }
    }

    /// The per-shape decision: a title that exists only in a package is kept
    /// on a slide whose body has legacy text. A per-slide check would lose it.
    #[test]
    fn test_shape_without_legacy_text_keeps_its_package_beside_a_shape_with_some() {
        let body = legacy_shape(atom(RT_TEXT_CHARS, &utf16("Body")), carrier_record(&sp("Body")));
        let title = container(SP_CONTAINER, &[carrier_record(&sp("Only here"))]);
        let runs = run_slide(&[body, title], &mut DownrevBudget::new());
        assert_eq!(texts(&runs), ["Only here"]);
    }

    /// A stub is not text: whitespace and the field marker leave the shape
    /// reading from its package.
    #[test]
    fn test_legacy_stub_does_not_count_as_legacy_text() {
        for stub in ["", "  ", "\r", "*", " * "] {
            let shape = legacy_shape(atom(RT_TEXT_CHARS, &utf16(stub)), carrier_record(SAMPLE_XML));
            let runs = run_slide(&[shape], &mut DownrevBudget::new());
            assert_eq!(texts(&runs), ["Real package text"], "{stub:?}");
        }
    }

    /// Text by reference counts only when the reference resolves to text.
    #[test]
    fn test_outline_reference_counts_as_legacy_text_when_it_resolves() {
        let reference = atom(RT_OUTLINE_TEXT_REF_ATOM, &0i32.to_le_bytes());
        let shape = legacy_shape(reference, carrier_record(SAMPLE_XML));
        let run = |outline: &[TextRun]| {
            let mut runs = Vec::new();
            let mut text = TextBudget::with_limit(usize::MAX);
            let mut budget = DownrevBudget::new();
            append_slide_downrev(&shape, outline, &mut runs, &mut budget, &mut text);
            runs.len()
        };
        let outline = |text: &str| TextRun {
            text: text.to_string(),
            ..Default::default()
        };
        assert_eq!(run(&[outline("Outline text")]), 0, "resolves to text");
        assert_eq!(run(&[outline(" ")]), 1, "resolves to a stub");
        assert_eq!(run(&[]), 1, "resolves to nothing");
    }

    /// A deleted shape isn't on the slide, whatever its package says.
    #[test]
    fn test_deleted_shape_ignores_its_package() {
        let fsp = |flags: u32| {
            let mut data = 1u32.to_le_bytes().to_vec();
            data.extend_from_slice(&flags.to_le_bytes());
            record_inst(RT_FSP, 1, 2, &data)
        };
        let shape = |flags| container(SP_CONTAINER, &[fsp(flags), carrier_record(SAMPLE_XML)]);
        let run = |flags| run_slide(&[shape(flags)], &mut DownrevBudget::new()).len();
        assert_eq!(run(FSP_FDELETED), 0, "deleted");
        assert_eq!(run(FSP_FDELETED | 0x0A00), 0, "deleted, other flags set");
        assert_eq!(run(0x0A00), 1, "not deleted");
    }

    /// Two shapes can say the same thing, and a shape can repeat a line.
    #[test]
    fn test_repeated_labels_in_distinct_packages_are_kept() {
        let first =
            "<p:sp><a:p><a:r><a:t>Yes</a:t></a:r></a:p><a:p><a:r><a:t>No</a:t></a:r></a:p></p:sp>";
        let second = "<p:sp><a:p><a:r><a:t>Yes</a:t></a:r></a:p><a:p><a:r><a:t>Maybe</a:t></a:r></a:p><a:p><a:r><a:t>Maybe</a:t></a:r></a:p></p:sp>";
        let runs =
            run_slide(&[carrier_record(first), carrier_record(second)], &mut DownrevBudget::new());
        assert_eq!(texts(&runs), ["Yes", "No", "Yes", "Maybe", "Maybe"]);
    }

    /// Real files repeat a shape's package across records: the same bytes on a
    /// slide are one shape.
    #[test]
    fn test_identical_package_collapses_within_a_slide_but_not_across_slides() {
        let carrier = carrier_record(SAMPLE_XML);
        let mut budget = DownrevBudget::new();
        let before = budget.packages_remaining;
        let runs = run_slide(&[carrier.clone(), carrier.clone()], &mut budget);
        assert_eq!(texts(&runs), ["Real package text"]);
        assert_eq!(before - budget.packages_remaining, 1, "the repeat is free");
        let runs = run_slide(&[carrier], &mut budget);
        assert_eq!(texts(&runs), ["Real package text"], "the next slide keeps its copy");
    }

    // Roles and fields

    fn with_placeholder(kind: &str, text: &str) -> String {
        format!(
            r#"<p:sp><p:nvSpPr><p:nvPr><p:ph type="{kind}" idx="1"/></p:nvPr></p:nvSpPr><a:p><a:r><a:t>{text}</a:t></a:r></a:p></p:sp>"#
        )
    }

    #[test]
    fn test_placeholder_type_sets_the_text_type() {
        for (kind, expected) in [
            ("title", TextType::Title),
            ("ctrTitle", TextType::CenterTitle),
            ("subTitle", TextType::CenterBody),
            ("body", TextType::Body),
            ("ftr", TextType::Other),
        ] {
            let runs = run_slide(
                &[carrier_record(&with_placeholder(kind, "Words"))],
                &mut DownrevBudget::new(),
            );
            assert_eq!(runs[0].text_type, expected, "{kind}");
        }
    }

    /// A title goes ahead of the text already on the slide, as one run.
    #[test]
    fn test_title_goes_first_and_is_one_run() {
        let xml = r#"<p:sp><p:nvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><a:p><a:r><a:t>One</a:t></a:r></a:p><a:p><a:r><a:t>Two</a:t></a:r></a:p></p:sp>"#;
        let mut runs = vec![TextRun {
            text: "Legacy body".to_string(),
            text_type: TextType::Body,
            ..Default::default()
        }];
        let mut budget = DownrevBudget::new();
        let mut text = TextBudget::with_limit(usize::MAX);
        append_slide_downrev(&carrier_record(xml), &[], &mut runs, &mut budget, &mut text);
        assert_eq!(texts(&runs), ["One\rTwo", "Legacy body"]);
        assert_eq!(runs[0].text_type, TextType::Title);
    }

    /// Cached field results aren't slide text; the words around them are.
    #[test]
    fn test_slide_number_and_date_values_are_left_out() {
        let number = carrier_record(&with_placeholder("sldNum", "7"));
        let field = carrier_record(
            r#"<p:sp><a:p><a:r><a:t>Page </a:t></a:r><a:fld type="slidenum"><a:t>7</a:t></a:fld></a:p><a:p><a:fld type="datetime1"><a:t>1/1/2026</a:t></a:fld></a:p><a:p><a:fld type="other"><a:t>Kept</a:t></a:fld></a:p></p:sp>"#,
        );
        let runs = run_slide(&[number, field], &mut DownrevBudget::new());
        assert_eq!(texts(&runs), ["Page ", "Kept"]);
    }

    /// Runs of underscores and the slide-number marker are nothing to read;
    /// symbols are real text.
    #[test]
    fn test_placeholder_masks_are_dropped_and_symbols_kept() {
        for mask in ["___ _______", "_", "‹#›", "  "] {
            let runs = run_slide(&[carrier_record(&sp(mask))], &mut DownrevBudget::new());
            assert!(runs.is_empty(), "{mask:?}: {runs:?}");
        }
        for symbol in ["→", "✓", "%", "@#$"] {
            let runs = run_slide(&[carrier_record(&sp(symbol))], &mut DownrevBudget::new());
            assert_eq!(texts(&runs), [symbol]);
        }
    }

    // Fragment XML

    #[test]
    fn test_runs_group_per_paragraph_and_soft_returns_survive() {
        let xml = r#"<p:sp><p:txBody><a:p><a:r><a:t>Hello </a:t></a:r><a:r><a:t>world</a:t></a:r></a:p><a:p><a:r><a:t>Second</a:t></a:r><a:br/><a:r><a:t>line</a:t></a:r></a:p></p:txBody></p:sp>"#;
        assert_eq!(fragment_paragraphs(xml), ["Hello world", "Second\u{b}line"]);
    }

    /// `<a:br>` with properties is not an empty element.
    #[test]
    fn test_soft_return_with_a_child_element_still_breaks_the_line() {
        let xml = r#"<p:sp><a:p><a:r><a:t>Second</a:t></a:r><a:br><a:rPr lang="en-US"/></a:br><a:r><a:t>line</a:t></a:r></a:p></p:sp>"#;
        assert_eq!(fragment_paragraphs(xml), ["Second\u{b}line"]);
    }

    #[test]
    fn test_entities_resolve() {
        let xml = r#"<p:sp><a:p><a:r><a:t>AT&amp;T &#8212; ok</a:t></a:r></a:p></p:sp>"#;
        assert_eq!(fragment_paragraphs(xml), ["AT&T — ok"]);
    }

    #[test]
    fn test_damaged_fragment_xml_keeps_what_it_read_and_warns() {
        let xml = r#"<p:sp><a:p><a:r><a:t>Kept</a:t></a:r></a:p><a:p a="<"#;
        let mut budget = DownrevBudget::new();
        let runs = run_slide(&[carrier_record(xml)], &mut budget);
        assert_eq!(texts(&runs), ["Kept"]);
        assert_eq!(budget.notices().len(), 1, "{:?}", budget.notices());
    }

    // Limits

    /// Past the per-slide cap, reading stops and says so. Reaching the cap
    /// exactly loses nothing and says nothing.
    #[test]
    fn test_per_slide_cap_boundary() {
        let carriers = |n: usize| -> Vec<Vec<u8>> {
            (0..n)
                .map(|i| carrier_record(&sp(&format!("Text {i}"))))
                .collect()
        };
        let mut budget = DownrevBudget::new();
        let runs = run_slide(&carriers(MAX_DOWNREV_PACKAGES), &mut budget);
        assert_eq!(runs.len(), MAX_DOWNREV_PACKAGES);
        assert!(budget.notices().is_empty(), "{:?}", budget.notices());

        let mut budget = DownrevBudget::new();
        let runs = run_slide(&carriers(MAX_DOWNREV_PACKAGES + 1), &mut budget);
        assert_eq!(runs.len(), MAX_DOWNREV_PACKAGES);
        assert_eq!(budget.notices().len(), 1, "{:?}", budget.notices());
    }

    /// The per-slide allowance comes back on the next slide; the file's doesn't.
    #[test]
    fn test_per_file_package_budget_spans_slides_and_warns() {
        let mut budget = DownrevBudget::new();
        budget.packages_remaining = 3;
        let mut total = 0;
        for s in 0..3 {
            let carriers: Vec<_> = (0..2)
                .map(|i| carrier_record(&sp(&format!("S{s} T{i}"))))
                .collect();
            total += run_slide(&carriers, &mut budget).len();
        }
        assert_eq!(total, 3);
        assert_eq!(budget.packages_remaining, 0);
        assert_eq!(budget.notices().len(), 1, "{:?}", budget.notices());
    }

    /// A package that fails or has no text costs what it cost to find out.
    /// Otherwise a file of them has no limit at all.
    #[test]
    fn test_failed_packages_are_charged() {
        let invalid = vec![0xFFu8; MAX_SHAPE_XML_BYTES as usize];
        let cases: Vec<(&str, Vec<u8>, u64)> = vec![
            ("invalid UTF-8", package_bytes(&invalid), MAX_SHAPE_XML_BYTES),
            ("no text", package(&sp("")), sp("").len() as u64),
            ("no such entry", zip_with("drs/other.xml", b"x"), 0),
            ("not a ZIP", b"PK\x03\x04 nope".to_vec(), 0),
        ];
        for (name, pkg, bytes) in cases {
            let mut budget = DownrevBudget::new();
            let (packages, inflated) = (budget.packages_remaining, budget.bytes_remaining);
            run_slide(&[table_record(&[(METROBLOB_COMPLEX, Some(&pkg))])], &mut budget);
            assert_eq!(packages - budget.packages_remaining, 1, "{name}");
            assert_eq!(inflated - budget.bytes_remaining, bytes, "{name}");
        }
    }

    fn package_bytes(data: &[u8]) -> Vec<u8> {
        zip_with(SHAPE_XML_ENTRY, data)
    }

    #[test]
    fn test_invalid_utf8_stops_at_the_byte_budget_and_warns() {
        let invalid = vec![0xFFu8; 1024];
        let carriers: Vec<_> = (0..4)
            .map(|i| {
                let mut data = invalid.clone();
                data[0] = i;
                table_record(&[(METROBLOB_COMPLEX, Some(&package_bytes(&data)))])
            })
            .collect();
        let mut budget = DownrevBudget::new();
        budget.bytes_remaining = 2048;
        let runs = run_slide(&carriers, &mut budget);
        assert!(runs.is_empty());
        assert_eq!(budget.bytes_remaining, 0);
        assert_eq!(budget.packages_remaining, DOWNREV_MAX_PACKAGES_PER_FILE - 2);
        assert_eq!(budget.notices().len(), 2, "limit and damage: {:?}", budget.notices());
    }

    /// A cut keeps the paragraphs it holds whole, and says so. It doesn't end
    /// on a part of a word.
    #[test]
    fn test_fragment_cut_by_the_cap_keeps_whole_paragraphs_and_warns() {
        let paragraphs: String = (0..3000)
            .map(|i| format!("<a:p><a:r><a:t>paragraph {i}</a:t></a:r></a:p>"))
            .collect();
        let mut budget = DownrevBudget::new();
        let runs = run_slide(&[carrier_record(&format!("<p:sp>{paragraphs}</p:sp>"))], &mut budget);
        assert_eq!(budget.bytes_remaining, DOWNREV_MAX_BYTES_PER_FILE - MAX_SHAPE_XML_BYTES);
        assert!(runs.len() > 100 && runs.len() < 3000, "{}", runs.len());
        for (i, run) in runs.iter().enumerate() {
            assert_eq!(run.text, format!("paragraph {i}"));
        }
        assert_eq!(budget.notices().len(), 1, "{:?}", budget.notices());
    }

    /// Wherever a cut lands in a multi-byte character, the paragraphs before
    /// it survive.
    #[test]
    fn test_cut_inside_a_multibyte_character_keeps_earlier_paragraphs() {
        let body = "国".repeat(30_000);
        for pad in 0..3 {
            let xml = format!(
                "<p:sp>{}<a:p><a:r><a:t>Heading</a:t></a:r></a:p><a:p><a:r><a:t>{body}</a:t></a:r></a:p></p:sp>",
                " ".repeat(pad)
            );
            let mut budget = DownrevBudget::new();
            let runs = run_slide(&[carrier_record(&xml)], &mut budget);
            assert_eq!(texts(&runs), ["Heading"], "pad {pad}");
            assert_eq!(budget.notices().len(), 1, "pad {pad}: {:?}", budget.notices());
        }
    }

    /// Text from packages counts against the document's text budget.
    #[test]
    fn test_package_text_is_charged_to_the_text_budget() {
        let mut runs = Vec::new();
        let mut budget = DownrevBudget::new();
        let mut text = TextBudget::with_limit(8);
        let carriers = slide(&[carrier_record(&sp("12345")), carrier_record(&sp("67890"))]);
        append_slide_downrev(&carriers, &[], &mut runs, &mut budget, &mut text);
        assert_eq!(texts(&runs), ["12345"]);
        assert!(text.exhausted());
    }

    // Damaged input

    #[test]
    fn test_carriers_nested_past_the_depth_limit_are_not_read() {
        let nest = |mut record: Vec<u8>, depth: usize| {
            for _ in 0..depth {
                record = container(0xF003, &[record]);
            }
            record
        };
        let carrier = carrier_record(SAMPLE_XML);
        let runs =
            run_slide(&[nest(carrier.clone(), MAX_CARRIER_DEPTH)], &mut DownrevBudget::new());
        assert_eq!(texts(&runs), ["Real package text"], "at the limit");
        let runs = run_slide(&[nest(carrier, MAX_CARRIER_DEPTH + 1)], &mut DownrevBudget::new());
        assert!(runs.is_empty(), "past the limit: {runs:?}");
    }

    /// Legacy text too deep to count is too deep for the legacy walk to read,
    /// and the shape's package still answers for it. No overflow either way.
    #[test]
    fn test_legacy_text_past_the_depth_limit_does_not_count() {
        let mut textbox = atom(RT_TEXT_CHARS, &utf16("Deep"));
        for _ in 0..=MAX_CARRIER_DEPTH + 1 {
            textbox = container(CLIENT_TEXTBOX, &[textbox]);
        }
        let shape = container(SP_CONTAINER, &[textbox, carrier_record(SAMPLE_XML)]);
        let runs = run_slide(&[shape], &mut DownrevBudget::new());
        assert_eq!(texts(&runs), ["Real package text"]);
    }

    #[test]
    fn test_short_shape_record_is_not_deleted() {
        let shape = container(
            SP_CONTAINER,
            &[
                record_inst(RT_FSP, 1, 2, &[0xFF; 4]),
                carrier_record(SAMPLE_XML),
            ],
        );
        assert_eq!(run_slide(&[shape], &mut DownrevBudget::new()).len(), 1);
    }

    #[test]
    fn test_unexplained_table_with_no_zip_adds_nothing() {
        let carrier = record_inst(RT_TERTIARY_FOPT, 1, 3, &[0x11; 32]);
        let mut budget = DownrevBudget::new();
        assert!(run_slide(&[carrier], &mut budget).is_empty());
        assert_eq!(budget.packages_remaining, DOWNREV_MAX_PACKAGES_PER_FILE, "nothing to read");
        assert!(budget.notices().is_empty());
    }

    fn stored_package(xml: &str) -> Vec<u8> {
        use std::io::Write;
        let mut buf = Vec::new();
        {
            let mut w = zip::ZipWriter::new(std::io::Cursor::new(&mut buf));
            let opts = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Stored);
            w.start_file(SHAPE_XML_ENTRY, opts).expect("zip entry");
            w.write_all(xml.as_bytes()).expect("zip payload");
            w.finish().expect("zip finish");
        }
        buf
    }

    fn assert_damage_reported(pkg: &[u8]) {
        let mut budget = DownrevBudget::new();
        let runs = run_slide(&[table_record(&[(METROBLOB_COMPLEX, Some(pkg))])], &mut budget);
        assert!(runs.is_empty(), "{runs:?}");
        assert_eq!(budget.notices().len(), 1, "{:?}", budget.notices());
    }

    /// The checksum fails at the end of the read.
    #[test]
    fn test_corrupt_entry_data_is_reported() {
        let mut pkg = stored_package(SAMPLE_XML);
        let at = pkg
            .windows(4)
            .position(|w| w == b"Real")
            .expect("stored text");
        pkg[at] = b'Z';
        assert_damage_reported(&pkg);
    }

    /// An encrypted entry can be found but not opened.
    #[test]
    fn test_entry_that_cannot_be_opened_is_reported() {
        let mut pkg = stored_package(SAMPLE_XML);
        let central = pkg
            .windows(4)
            .position(|w| w == b"PK\x01\x02")
            .expect("central directory");
        pkg[central + 8] |= 1; // general purpose flag: encrypted
        assert_damage_reported(&pkg);
    }

    /// Text outside a paragraph is still text, and ends its paragraph at the
    /// next one or at the end of the fragment.
    #[test]
    fn test_text_outside_a_paragraph_is_kept() {
        let xml =
            "<p:sp><a:t>Loose</a:t><a:p><a:r><a:t>Inside</a:t></a:r></a:p><a:t>Last</a:t></p:sp>";
        assert_eq!(fragment_paragraphs(xml), ["Loose", "Inside", "Last"]);
    }

    /// A placeholder element can be written with a body.
    #[test]
    fn test_placeholder_with_a_child_element_sets_the_text_type() {
        let xml = r#"<p:sp><p:ph type="title"><p:extLst/></p:ph><a:p><a:r><a:t>T</a:t></a:r></a:p></p:sp>"#;
        assert_eq!(parse_shape_fragment(xml, false).role, TextType::Title);
    }

    /// A run with no end is damage, unless a limit is what cut it.
    #[test]
    fn test_unclosed_text_is_damage_only_when_no_limit_cut_it() {
        let xml = "<p:sp><a:p><a:r><a:t>Cut &bad";
        assert!(parse_shape_fragment(xml, false).malformed);
        assert!(!parse_shape_fragment(xml, true).malformed);
    }

    /// A malformed table never panics, whatever the count says.
    #[test]
    fn test_property_table_counts_and_lengths_cannot_overflow() {
        let huge = [0xFFu8; 12];
        for properties in [0, 1, 2, 3, 0xFFF, usize::MAX] {
            assert!(!matches!(read_property_table(&huge, properties), PropertyTable::Package(_)));
            assert!(!matches!(read_property_table(&[], properties), PropertyTable::Package(_)));
        }
    }
}
