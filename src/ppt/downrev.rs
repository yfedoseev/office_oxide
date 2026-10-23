//! WPS "downrev" packages inside legacy `.ppt` streams.
//!
//! When WPS saves an OOXML-origin presentation to the legacy binary
//! format, shapes that carry OOXML-only content are written back as
//! self-contained ZIP packages attached to records inside the standard
//! OfficeArt shape tree: each package holds a `drs/shapexml.xml` entry —
//! a standard ECMA-376 DrawingML `<p:sp>` fragment whose `<a:t>` runs
//! carry the shape's text. The legacy `TextCharsAtom`/`TextBytesAtom`
//! atoms are left nearly empty for such shapes, so a reader that walks
//! only those atoms returns an almost-empty deck with `Ok` — a deck whose
//! words all live in the packages.
//!
//! The carrier is the OfficeArt tertiary property table (`0xF122`,
//! `msofbtTertiaryOpt`, [MS-ODRAW] 2.2.11), and the package is hung off
//! its `metroBlob` complex property ([MS-ODRAW] 2.3.4.41): the 2-byte opid
//! with `fComplex` set, a 4-byte byte count, then that many bytes. Real
//! files carry it that way — verified against a real WPS file, 383 of 383
//! packages declared an exact length. Extraction therefore parses the
//! property table and reads exactly the declared bytes; a signature scan
//! remains as a fallback for carriers that are not a property table (or
//! carry a malformed one). Either way a malformed or unrelated package is
//! skipped — never a panic (robustness contract), and at most
//! [`MAX_DOWNREV_PACKAGES`] packages are read per slide.

use std::collections::HashSet;
use std::io::Read;

use quick_xml::events::Event;

use super::records::RecordIter;
use super::text::{TextRun, TextType};

/// How deep into a slide's record tree to look for carrier records.
/// Mirrors the shape-nesting guard in [`super::text`].
const MAX_CARRIER_DEPTH: usize = 64;

/// How far into a record payload to scan for a ZIP signature. A
/// spec-shaped carrier puts the package right behind its 6-byte property
/// header; the window leaves room for a longer table in front of it.
const MAX_SIGNATURE_SCAN: usize = 4096;

/// Cap on one shape fragment's decompressed size. Real fragments are a
/// couple of KB; the cap keeps a crafted package from turning extraction
/// into a decompression bomb.
const MAX_SHAPE_XML_BYTES: u64 = 1 << 20;

/// Cap on downrev packages read per slide. Real decks carry a handful per
/// slide; the cap bounds the work a crafted record tree can force.
const MAX_DOWNREV_PACKAGES: usize = 256;

/// The ZIP entry holding a shape's DrawingML fragment.
const SHAPE_XML_ENTRY: &str = "drs/shapexml.xml";

/// `metroBlob` property id ([MS-ODRAW] 2.3.4.41): the low 14 bits of the
/// opid word; the two high bits are `fBid` (undefined, ignored) and
/// `fComplex` (set means the complex data follows the property header).
const METROBLOB_OPID: u16 = 0x03A9;

/// `fComplex` bit of an `OfficeArtFOPTEOPID` ([MS-ODRAW] 2.2.8).
const OPID_FCOMPLEX: u16 = 0x8000;

/// Mask that keeps only the property id, dropping `fComplex` and `fBid`.
const OPID_ID_MASK: u16 = 0x3FFF;

/// Append the text of every WPS downrev package found in
/// `slide_children` — a `Slide` container's children region — to `runs`.
///
/// Paragraphs already present in `runs` (a legacy text atom keeps several
/// paragraphs in one run, split at `\r` / soft returns) are not repeated:
/// a shape WPS wrote both ways must not surface its text twice.
pub fn append_slide_downrev(slide_children: &[u8], runs: &mut Vec<TextRun>) {
    let mut seen: HashSet<String> = HashSet::new();
    for run in runs.iter() {
        for para in run.text.split(['\r', '\n', '\u{b}']) {
            let para = para.trim();
            if !para.is_empty() {
                seen.insert(para.to_string());
            }
        }
    }
    for text in collect_downrev_texts(slide_children) {
        if seen.insert(text.clone()) {
            runs.push(TextRun {
                text_type: TextType::Other,
                text,
                ..Default::default()
            });
        }
    }
}

/// Every distinct paragraph text carried by downrev packages anywhere in
/// the record subtree, in document order. At most
/// [`MAX_DOWNREV_PACKAGES`] packages are read; a crafted record tree cannot
/// force more work than that.
fn collect_downrev_texts(data: &[u8]) -> Vec<String> {
    let mut texts = Vec::new();
    let mut seen = HashSet::new();
    let mut budget = MAX_DOWNREV_PACKAGES;
    walk_records(data, 0, &mut texts, &mut seen, &mut budget);
    texts
}

fn walk_records(
    data: &[u8],
    depth: usize,
    texts: &mut Vec<String>,
    seen: &mut HashSet<String>,
    budget: &mut usize,
) {
    if depth > MAX_CARRIER_DEPTH {
        return;
    }
    for rec in RecordIter::new(data) {
        if *budget == 0 {
            return;
        }
        let Ok(rec) = rec else { break };
        if rec.header.is_container() {
            // A container's data is a record stream, not a payload: probing
            // it would read a nested package through the container and past
            // whatever length the carrier declared. Walk into it instead.
            walk_records(&rec.data, depth + 1, texts, seen, budget);
        } else if let Some(xml) = shape_xml_in_payload(&rec.data) {
            *budget -= 1;
            for text in shape_xml_paragraphs(&xml) {
                if !text.is_empty() && seen.insert(text.clone()) {
                    texts.push(text);
                }
            }
        }
    }
}

/// The package carried by the record's `metroBlob` complex property, if the
/// payload is a tertiary property table that carries one ([MS-ODRAW]
/// 2.3.4.41). A property entry is 6 bytes — a 2-byte opid (the high bits
/// are `fComplex` and `fBid`) and a 4-byte value; for a complex property
/// the value is the byte count of the data that follows the header, so the
/// declared length bounds the read exactly. A table whose complex data
/// runs off the payload is malformed and yields nothing.
fn metro_blob_package(payload: &[u8]) -> Option<&[u8]> {
    let mut pos = 0;
    while pos + 6 <= payload.len() {
        let opid = u16::from_le_bytes([payload[pos], payload[pos + 1]]);
        let declared = u32::from_le_bytes([
            payload[pos + 2],
            payload[pos + 3],
            payload[pos + 4],
            payload[pos + 5],
        ]) as usize;
        pos += 6;
        if opid & OPID_FCOMPLEX == 0 {
            continue; // non-complex: the 4-byte value is the whole property
        }
        let end = pos.checked_add(declared)?;
        if end > payload.len() {
            return None; // malformed table: the data runs off the payload
        }
        if opid & OPID_ID_MASK == METROBLOB_OPID {
            return Some(&payload[pos..end]);
        }
        pos = end; // some other complex property: skip its data
    }
    None
}

/// The `drs/shapexml.xml` content of the downrev package carried in
/// `payload`, if the payload carries one at all.
///
/// The spec-shaped carrier is read first: the property table's
/// `metroBlob` complex property, bounded by its declared length. Only a
/// payload that is not a property table — or carries a malformed one —
/// falls back to scanning for a ZIP local-header signature, which may sit
/// behind a small WPS header and is therefore located by scan rather than
/// assumed at offset 0.
fn shape_xml_in_payload(payload: &[u8]) -> Option<String> {
    if let Some(package) = metro_blob_package(payload) {
        return read_shape_xml(package);
    }
    let start = find_zip_signature(payload)?;
    read_shape_xml(&payload[start..])
}

/// The `drs/shapexml.xml` content of a self-contained package, with the
/// decompressed size capped at [`MAX_SHAPE_XML_BYTES`].
fn read_shape_xml(package: &[u8]) -> Option<String> {
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(package)).ok()?;
    let entry = archive.by_name(SHAPE_XML_ENTRY).ok()?;
    let mut xml = String::new();
    entry
        .take(MAX_SHAPE_XML_BYTES)
        .read_to_string(&mut xml)
        .ok()?;
    Some(xml)
}

fn find_zip_signature(payload: &[u8]) -> Option<usize> {
    let limit = payload.len().min(MAX_SIGNATURE_SCAN);
    payload[..limit].windows(4).position(|w| w == b"PK\x03\x04")
}

/// The shape fragment's paragraph texts: `<a:t>` runs grouped per `<a:p>`
/// (a paragraph's text is split across runs arbitrarily), with `<a:br/>`
/// soft returns kept as line breaks. Entities and CDATA resolve through
/// the shared OOXML text reader.
///
/// The reader deliberately does not trim text events (unlike
/// [`crate::core::xml::make_fast_reader]`): a paragraph's words are split
/// across runs at arbitrary points, so `"Hello "` + `"world"` must not
/// collapse to `"Helloworld"`.
fn shape_xml_paragraphs(xml: &str) -> Vec<String> {
    use crate::core::xml::read_text_content_fast;

    let mut reader = quick_xml::Reader::from_reader(xml.as_bytes());
    reader.config_mut().check_end_names = false;
    reader.config_mut().check_comments = false;
    let mut paragraphs = Vec::new();
    let mut current = String::new();
    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) => match e.local_name().as_ref() {
                "p" if !current.is_empty() => {
                    paragraphs.push(std::mem::take(&mut current));
                },
                "t" => match read_text_content_fast(&mut reader) {
                    Ok(t) => current.push_str(&t),
                    Err(_) => break,
                },
                _ => {},
            },
            Ok(Event::Empty(e)) => {
                if e.local_name().as_ref() == "br" {
                    current.push('\u{b}');
                }
            },
            Ok(Event::End(e)) => {
                if e.local_name().as_ref() == "p" && !current.is_empty() {
                    paragraphs.push(std::mem::take(&mut current));
                }
            },
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {},
        }
    }
    if !current.is_empty() {
        paragraphs.push(current);
    }
    paragraphs
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The opid word a real file writes for the `metroBlob` complex
    /// property: the property id with `fComplex` set ([MS-ODRAW] 2.3.4.41).
    const METROBLOB_OPID_COMPLEX: u16 = OPID_FCOMPLEX | METROBLOB_OPID;

    /// A complex (non-metroBlob) tertiary property id, used to hang a decoy
    /// package in front of the real one.
    const DECOY_COMPLEX_OPID: u16 = 0x8000 | 0x0180;

    /// The carrier record type: `msofbtTertiaryOpt` ([MS-ODRAW] 2.2.11).
    const RT_CARRIER: u16 = 0xF122;

    /// A shape fragment with one paragraph.
    const SAMPLE_XML: &str = r#"<p:sp><a:p><a:r><a:t>Real package text</a:t></a:r></a:p></p:sp>"#;

    /// A carrier whose payload is the spec-shaped `metroBlob` complex
    /// property: opid `0x83A9`, the 4-byte byte count, then the package —
    /// the shape real WPS files write ([MS-ODRAW] 2.3.4.41).
    fn carrier(xml: &str) -> Vec<u8> {
        use std::io::Write;
        let mut buf = Vec::new();
        {
            let mut w = zip::ZipWriter::new(std::io::Cursor::new(&mut buf));
            let opts = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Deflated);
            w.start_file(SHAPE_XML_ENTRY, opts).expect("zip entry");
            w.write_all(xml.as_bytes()).expect("zip payload");
            w.finish().expect("zip finish");
        }
        let mut payload = Vec::with_capacity(6 + buf.len());
        payload.extend_from_slice(&METROBLOB_OPID_COMPLEX.to_le_bytes());
        payload.extend_from_slice(&(buf.len() as u32).to_le_bytes());
        payload.extend_from_slice(&buf);
        payload
    }

    /// A carrier as a full `0xF122` record (recInstance 1, recVer 3 — the
    /// header real files write), wrapping the spec-shaped payload.
    fn carrier_record(xml: &str) -> Vec<u8> {
        record(RT_CARRIER, 3, &carrier(xml))
    }

    /// A carrier whose property table opens with a non-metroBlob **complex**
    /// property whose data is a complete, valid ZIP carrying the *wrong*
    /// text, followed by the real metroBlob package. A signature scan stops
    /// at the first ZIP and reads the wrong one; a property-table walk skips
    /// the non-metroBlob property and reads the real package.
    fn carrier_with_decoy(xml: &str) -> Vec<u8> {
        use std::io::Write;
        let mut decoy = Vec::new();
        {
            let mut w = zip::ZipWriter::new(std::io::Cursor::new(&mut decoy));
            let opts = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Deflated);
            w.start_file(SHAPE_XML_ENTRY, opts)
                .expect("decoy zip entry");
            w.write_all(b"<p:sp><a:p><a:r><a:t>Decoy package text</a:t></a:r></a:p></p:sp>")
                .expect("decoy zip payload");
            w.finish().expect("decoy zip finish");
        }
        let pkg = carrier(xml);
        let mut payload = Vec::with_capacity(12 + decoy.len() + pkg.len());
        payload.extend_from_slice(&DECOY_COMPLEX_OPID.to_le_bytes()); // complex, not metroBlob
        payload.extend_from_slice(&(decoy.len() as u32).to_le_bytes());
        payload.extend_from_slice(&decoy);
        payload.extend_from_slice(&METROBLOB_OPID_COMPLEX.to_le_bytes());
        payload.extend_from_slice(&((pkg.len() - 6) as u32).to_le_bytes());
        payload.extend_from_slice(&pkg[6..]);
        payload
    }

    /// A carrier whose metroBlob declares fewer bytes than the package
    /// holds — a corrupt length that must bound the read.
    fn carrier_short_length(xml: &str, shortfall: usize) -> Vec<u8> {
        let pkg = carrier(xml);
        let declared = pkg.len() - 6 - shortfall;
        let mut payload = Vec::with_capacity(pkg.len());
        payload.extend_from_slice(&METROBLOB_OPID_COMPLEX.to_le_bytes());
        payload.extend_from_slice(&(declared as u32).to_le_bytes());
        payload.extend_from_slice(&pkg[6..]); // the full package still follows
        payload
    }

    /// A carrier whose metroBlob declares more bytes than the payload
    /// holds — a malformed table that must not read out of bounds.
    fn carrier_over_long_length(xml: &str, excess: usize) -> Vec<u8> {
        let pkg = carrier(xml);
        let declared = (pkg.len() - 6) + excess;
        let mut payload = Vec::with_capacity(pkg.len());
        payload.extend_from_slice(&METROBLOB_OPID_COMPLEX.to_le_bytes());
        payload.extend_from_slice(&(declared as u32).to_le_bytes());
        payload.extend_from_slice(&pkg[6..]);
        payload
    }

    /// A carrier whose opid is metroBlob's id without the complex bit: the
    /// spec requires the complex data only when fComplex is set, so the
    /// walk finds no package here.
    fn carrier_without_complex_bit(xml: &str) -> Vec<u8> {
        let pkg = carrier(xml);
        let mut payload = Vec::with_capacity(pkg.len());
        payload.extend_from_slice(&METROBLOB_OPID.to_le_bytes()); // no fComplex
        payload.extend_from_slice(&0u32.to_le_bytes()); // value must be 0
        payload.extend_from_slice(&pkg[6..]);
        payload
    }

    /// A carrier that is not a property table at all: six zero bytes (a
    /// degenerate non-complex property) followed by the package — the
    /// shape an earlier revision of the fixture wrote.
    fn degenerate_carrier(xml: &str) -> Vec<u8> {
        let pkg = carrier(xml);
        let mut payload = Vec::with_capacity(pkg.len());
        payload.extend_from_slice(&[0u8; 6]);
        payload.extend_from_slice(&pkg[6..]);
        payload
    }

    fn record(rec_type: u16, ver: u8, payload: &[u8]) -> Vec<u8> {
        let ver_instance = ((1u16 & 0x0FFF) << 4) | (ver & 0x0F) as u16;
        let mut v = Vec::with_capacity(8 + payload.len());
        v.extend_from_slice(&ver_instance.to_le_bytes());
        v.extend_from_slice(&rec_type.to_le_bytes());
        v.extend_from_slice(&(payload.len() as u32).to_le_bytes());
        v.extend_from_slice(payload);
        v
    }

    fn container(rec_type: u16, children: &[Vec<u8>]) -> Vec<u8> {
        let body: Vec<u8> = children.concat();
        record(rec_type, 0x0F, &body)
    }

    #[test]
    fn test_shape_xml_paragraphs_groups_runs_per_paragraph() {
        let xml = r#"<p:sp xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:txBody><a:p><a:r><a:t>Hello </a:t></a:r><a:r><a:t>world</a:t></a:r></a:p><a:p><a:r><a:t>Second</a:t></a:r><a:br/><a:r><a:t>line</a:t></a:r></a:p></p:txBody></p:sp>"#;
        assert_eq!(
            shape_xml_paragraphs(xml),
            vec!["Hello world".to_string(), "Second\u{b}line".to_string()]
        );
    }

    #[test]
    fn test_shape_xml_paragraphs_resolves_entities() {
        let xml = r#"<p:sp><a:p><a:r><a:t>AT&amp;T &#8212; ok</a:t></a:r></a:p></p:sp>"#;
        assert_eq!(shape_xml_paragraphs(xml), vec!["AT&T — ok".to_string()]);
    }

    #[test]
    fn test_find_zip_signature_skips_wps_prefix() {
        let mut payload = vec![0u8; 6];
        payload.extend_from_slice(b"PK\x03\x04rest");
        assert_eq!(find_zip_signature(&payload), Some(6));
        assert_eq!(find_zip_signature(b"no zip here"), None);
    }

    #[test]
    fn test_append_slide_downrev_collects_and_dedups() {
        let good = carrier_record(r#"<p:sp><a:p><a:r><a:t>Shape text</a:t></a:r></a:p></p:sp>"#);
        let dup = carrier_record(r#"<p:sp><a:p><a:r><a:t>Shape text</a:t></a:r></a:p></p:sp>"#);
        let slide = container(0x03EE, &[record(0x03EF, 0, &[0u8; 8]), good, dup]);
        let mut runs = Vec::new();
        append_slide_downrev(&slide, &mut runs);
        assert_eq!(runs.len(), 1, "duplicate packages collapse: {runs:?}");
        assert_eq!(runs[0].text, "Shape text");
        assert_eq!(runs[0].text_type, TextType::Other);
    }

    #[test]
    fn test_append_slide_downrev_skips_malformed_and_foreign_packages() {
        let broken = {
            let mut p = b"PK\x03\x04garbage".to_vec();
            p.extend_from_slice(&[0xFF; 32]);
            record(RT_CARRIER, 3, &p)
        };
        let slide = container(0x03EE, &[record(0x03EF, 0, &[0u8; 8]), broken]);
        let mut runs = Vec::new();
        append_slide_downrev(&slide, &mut runs);
        assert!(runs.is_empty(), "malformed package is skipped: {runs:?}");
    }

    #[test]
    fn test_metro_blob_declared_length_bounds_the_package() {
        // A declared length shorter than the package must bound the read:
        // the declared bytes are all that may be consumed, and the truncated
        // package is not carved around by scanning past the declaration.
        let payload = carrier_short_length(SAMPLE_XML, 8);
        assert!(
            shape_xml_in_payload(&payload).is_none(),
            "a corrupt (short) declared length must not be carved around"
        );
    }

    #[test]
    fn test_decoy_zip_signature_inside_a_property_value_is_skipped() {
        // A non-metroBlob complex property carries a complete valid ZIP with
        // the wrong text, in front of the real metroBlob package. A
        // signature scan stops at the first ZIP and reads the wrong one; a
        // property-table walk skips the decoy and reads the real package.
        let payload = carrier_with_decoy(SAMPLE_XML);
        let xml = shape_xml_in_payload(&payload)
            .expect("the real package must be read past the decoy package");
        assert!(xml.contains("Real package text"), "the real package's text must surface: {xml}");
        assert!(
            !xml.contains("Decoy package text"),
            "the decoy package's text must not surface: {xml}"
        );
    }

    #[test]
    fn test_degenerate_carrier_falls_back_to_the_signature_scan() {
        // A payload that is not a property table at all (six zero bytes —
        // a degenerate non-complex property) still yields its package
        // through the signature-scan fallback.
        let payload = degenerate_carrier(SAMPLE_XML);
        assert!(
            shape_xml_in_payload(&payload).is_some(),
            "a non-property-table carrier must still be read by the scan"
        );
    }

    #[test]
    fn test_metro_blob_without_the_complex_bit_falls_back_to_the_scan() {
        // opid 0x03A9 without fComplex carries no complex data per the spec,
        // so the walk finds no package and the scan fallback applies.
        let payload = carrier_without_complex_bit(SAMPLE_XML);
        assert!(
            shape_xml_in_payload(&payload).is_some(),
            "a metroBlob opid without the complex bit must fall back to the scan"
        );
    }

    #[test]
    fn test_over_long_declared_length_falls_back_without_panicking() {
        // A declared length past the end of the payload is a malformed
        // table: no panic, no out-of-bounds read, and the scan fallback
        // still recovers the package.
        let payload = carrier_over_long_length(SAMPLE_XML, 64);
        assert!(
            shape_xml_in_payload(&payload).is_some(),
            "an over-long declared length must not panic and must fall back"
        );
    }

    #[test]
    fn test_package_count_cap_drops_packages_beyond_the_cap() {
        // 257 distinct packages on one slide: the cap bounds how many are
        // read, so the 257th package's text must not surface.
        let carriers: Vec<Vec<u8>> = (0..MAX_DOWNREV_PACKAGES + 1)
            .map(|i| {
                carrier_record(&format!("<p:sp><a:p><a:r><a:t>Text {i}</a:t></a:r></a:p></p:sp>"))
            })
            .collect();
        let slide = container(0x03EE, &carriers);
        let mut runs = Vec::new();
        append_slide_downrev(&slide, &mut runs);
        assert!(
            runs.iter().any(|r| r.text == "Text 0"),
            "the first package must be read: {runs:?}"
        );
        assert!(
            runs.iter()
                .any(|r| r.text == format!("Text {}", MAX_DOWNREV_PACKAGES - 1)),
            "the last package inside the cap must be read: {runs:?}"
        );
        assert!(
            !runs
                .iter()
                .any(|r| r.text == format!("Text {MAX_DOWNREV_PACKAGES}")),
            "the package beyond the cap must not be read: {runs:?}"
        );
    }

    #[test]
    fn test_package_count_cap_boundary_reads_every_package_at_the_cap() {
        // Exactly at the cap: every package is still read.
        let carriers: Vec<Vec<u8>> = (0..MAX_DOWNREV_PACKAGES)
            .map(|i| {
                carrier_record(&format!("<p:sp><a:p><a:r><a:t>Text {i}</a:t></a:r></a:p></p:sp>"))
            })
            .collect();
        let slide = container(0x03EE, &carriers);
        let mut runs = Vec::new();
        append_slide_downrev(&slide, &mut runs);
        for i in 0..MAX_DOWNREV_PACKAGES {
            assert!(
                runs.iter().any(|r| r.text == format!("Text {i}")),
                "package {i} must be read at the cap boundary: {runs:?}"
            );
        }
    }
}
