#![allow(dead_code)]
//! PowerPoint binary record types and parsing.
//!
//! PPT records have an 8-byte header:
//! - Bits 0-3: recVer (version)
//! - Bits 4-15: recInstance
//! - Bytes 2-3: recType (u16)
//! - Bytes 4-7: recLen (u32)
//!
//! Container records (recVer = 0xF) contain child records.
//! Atom records contain raw data.

use super::error::{PptError, Result};

// ── Record type IDs ──
pub const RT_DOCUMENT: u16 = 0x03E8;
pub const RT_SLIDE: u16 = 0x03EE;
pub const RT_SLIDE_BASE: u16 = 0x03EC;
pub const RT_NOTES: u16 = 0x03F0;
pub const RT_SLIDE_LIST_WITH_TEXT: u16 = 0x0FF0;
pub const RT_TEXT_HEADER: u16 = 0x0F9F;
/// OutlineTextRefAtom ([MS-PPT] 2.4.15.6) — a shape's text stored *by
/// reference* as a zero-based index into the TextHeaderAtom sequence that
/// follows its slide's SlidePersistAtom in SlideListWithTextContainer,
/// instead of embedded directly in the shape.
pub const RT_OUTLINE_TEXT_REF_ATOM: u16 = 0x0F9E;
pub const RT_TEXT_CHARS: u16 = 0x0FA0;
pub const RT_TEXT_BYTES: u16 = 0x0FA8;
pub const RT_SLIDE_PERSIST_ATOM: u16 = 0x03F3;
pub const RT_USER_EDIT_ATOM: u16 = 0x0FF5;
/// PersistDirectoryAtom ([MS-PPT] 2.3.4).
pub const RT_PERSIST_DIRECTORY_ATOM: u16 = 0x1772;
pub const RT_CURRENT_USER_ATOM: u16 = 0x0FF6;
pub const RT_HEADER_FOOTER: u16 = 0x0FDA;
pub const RT_TEXT_SPECIAL_INFO: u16 = 0x0FAA;
pub const RT_TEXT_RULER: u16 = 0x0FA2;
pub const RT_STYLE_TEXT_PROP: u16 = 0x0FA1;
pub const RT_CSTRING: u16 = 0x0FBA;

// ── SlideListWithText `rh.recInstance` discriminants ([MS-PPT] 2.4.14) ──
/// `rh.recInstance` value identifying a `SlideListWithTextContainer` (real slides).
pub const SLWT_SLIDES: u16 = 0;
/// `rh.recInstance` value identifying a `MasterListWithTextContainer`.
pub const SLWT_MASTERS: u16 = 1;
/// `rh.recInstance` value identifying a `NotesListWithTextContainer`.
pub const SLWT_NOTES: u16 = 2;

/// A parsed PPT record header.
#[derive(Debug, Clone, Copy)]
pub struct RecordHeader {
    pub rec_ver: u8,
    pub rec_instance: u16,
    pub rec_type: u16,
    pub rec_len: u32,
}

impl RecordHeader {
    pub fn is_container(&self) -> bool {
        self.rec_ver == 0x0F
    }

    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() < 8 {
            return Err(PptError::InvalidRecord("record header too short".into()));
        }
        let ver_instance = u16::from_le_bytes([data[0], data[1]]);
        let rec_ver = (ver_instance & 0x0F) as u8;
        let rec_instance = ver_instance >> 4;
        let rec_type = u16::from_le_bytes([data[2], data[3]]);
        let rec_len = u32::from_le_bytes([data[4], data[5], data[6], data[7]]);
        Ok(Self {
            rec_ver,
            rec_instance,
            rec_type,
            rec_len,
        })
    }
}

/// A PPT record with its header and data.
#[derive(Debug, Clone)]
pub struct PptRecord {
    pub header: RecordHeader,
    pub data: Vec<u8>,
    pub offset: usize,
}

/// Iterate over the immediate children of one record region (single level,
/// non-recursive).
///
/// Each yielded record's `data` is bounded strictly by that record's own
/// declared `rec_len`, clamped to whatever bytes actually remain in the slice
/// passed to [`RecordIter::new`] — never by descending into (or trusting the
/// declared lengths of) anything nested further down. For a container record,
/// `data` is its bounded *children* region; recurse by constructing a new
/// `RecordIter::new(&container_record.data)`.
///
/// This bounding is what keeps a corrupted or maliciously oversized `rec_len`
/// on one record from swallowing bytes that belong to its siblings, or to
/// anything outside its own enclosing container: skipping to the next sibling
/// only ever relies on the current record's own header, never on correctly
/// parsing what's inside it.
pub struct RecordIter<'a> {
    data: &'a [u8],
    pos: usize,
}

impl<'a> RecordIter<'a> {
    pub fn new(data: &'a [u8]) -> Self {
        Self { data, pos: 0 }
    }
}

impl<'a> Iterator for RecordIter<'a> {
    type Item = Result<PptRecord>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.pos + 8 > self.data.len() {
            return None;
        }

        let header = match RecordHeader::parse(&self.data[self.pos..]) {
            Ok(h) => h,
            Err(e) => return Some(Err(e)),
        };

        let record_offset = self.pos;
        let data_start = (self.pos + 8).min(self.data.len());
        // Clamp to *this* slice's own end — never trust rec_len past what's
        // actually here, whether the record is an atom or a container.
        let data_end = data_start
            .saturating_add(header.rec_len as usize)
            .min(self.data.len());

        self.pos = data_end;

        Some(Ok(PptRecord {
            header,
            data: self.data[data_start..data_end].to_vec(),
            offset: record_offset,
        }))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn make_atom(rec_type: u16, instance: u16, data: &[u8]) -> Vec<u8> {
        let ver_instance: u16 = instance << 4; // ver=0 (atom)
        let mut buf = Vec::new();
        buf.extend_from_slice(&ver_instance.to_le_bytes());
        buf.extend_from_slice(&rec_type.to_le_bytes());
        buf.extend_from_slice(&(data.len() as u32).to_le_bytes());
        buf.extend_from_slice(data);
        buf
    }

    fn make_container(rec_type: u16, instance: u16, children: &[u8]) -> Vec<u8> {
        let ver_instance: u16 = (instance << 4) | 0x0F; // ver=0xF (container)
        let mut buf = Vec::new();
        buf.extend_from_slice(&ver_instance.to_le_bytes());
        buf.extend_from_slice(&rec_type.to_le_bytes());
        buf.extend_from_slice(&(children.len() as u32).to_le_bytes());
        buf.extend_from_slice(children);
        buf
    }

    #[test]
    fn parse_record_header() {
        let data = make_atom(0x0FA0, 0, &[0x41, 0x00]);
        let header = RecordHeader::parse(&data).unwrap();
        assert_eq!(header.rec_type, RT_TEXT_CHARS);
        assert_eq!(header.rec_ver, 0);
        assert!(!header.is_container());
        assert_eq!(header.rec_len, 2);
    }

    #[test]
    fn parse_container_header() {
        let child = make_atom(RT_TEXT_CHARS, 0, &[0x41, 0x00]);
        let data = make_container(RT_SLIDE, 0, &child);
        let header = RecordHeader::parse(&data).unwrap();
        assert!(header.is_container());
        assert_eq!(header.rec_type, RT_SLIDE);
    }

    #[test]
    fn iterate_flat_atoms() {
        let mut stream = make_atom(RT_TEXT_HEADER, 0, &[0x00, 0x00, 0x00, 0x00]);
        stream.extend(make_atom(RT_TEXT_CHARS, 0, &[0x48, 0x00, 0x69, 0x00]));
        let records: Vec<_> = RecordIter::new(&stream)
            .collect::<std::result::Result<_, _>>()
            .unwrap();
        assert_eq!(records.len(), 2);
        assert_eq!(records[0].header.rec_type, RT_TEXT_HEADER);
        assert_eq!(records[1].header.rec_type, RT_TEXT_CHARS);
    }

    #[test]
    fn container_children_are_bounded_not_flattened() {
        let child1 = make_atom(RT_TEXT_HEADER, 0, &[0x00, 0x00, 0x00, 0x00]);
        let child2 = make_atom(RT_TEXT_CHARS, 0, &[0x41, 0x00]);
        let mut children = child1.clone();
        children.extend(&child2);
        let container = make_container(RT_SLIDE_LIST_WITH_TEXT, 0, &children);

        // A single-level iteration over the container yields the container
        // itself, not its descendants — it must not implicitly flatten.
        let records: Vec<_> = RecordIter::new(&container)
            .collect::<std::result::Result<_, _>>()
            .unwrap();
        assert_eq!(records.len(), 1);
        assert_eq!(records[0].header.rec_type, RT_SLIDE_LIST_WITH_TEXT);
        assert!(records[0].header.is_container());
        assert_eq!(records[0].data, children);

        // Recursing explicitly into the container's own bounded child region
        // reaches both children.
        let nested: Vec<_> = RecordIter::new(&records[0].data)
            .collect::<std::result::Result<_, _>>()
            .unwrap();
        assert_eq!(nested.len(), 2);
        assert_eq!(nested[0].header.rec_type, RT_TEXT_HEADER);
        assert_eq!(nested[1].header.rec_type, RT_TEXT_CHARS);
    }

    #[test]
    fn empty_stream() {
        let records: Vec<_> = RecordIter::new(&[])
            .collect::<std::result::Result<_, _>>()
            .unwrap();
        assert!(records.is_empty());
    }

    /// A single record with a corrupted/oversized declared length, nested
    /// inside a container, must not swallow bytes belonging to a sibling
    /// record *outside* that container. Skipping to the next sibling must
    /// rely only on the container's own header length, never on successfully
    /// parsing (or bounding) what's inside it.
    #[test]
    fn corrupt_nested_record_length_does_not_swallow_top_level_siblings() {
        // A child atom that declares a wildly oversized length with no data
        // behind it (simulating real-world corrupted/non-conformant files).
        let mut corrupt_child = Vec::new();
        corrupt_child.extend_from_slice(&0u16.to_le_bytes()); // ver=0 (atom), instance=0
        corrupt_child.extend_from_slice(&RT_TEXT_CHARS.to_le_bytes());
        corrupt_child.extend_from_slice(&1_000_000u32.to_le_bytes()); // bogus length

        let outer = make_container(RT_SLIDE, 0, &corrupt_child);

        let mut stream = outer;
        stream.extend(make_atom(RT_TEXT_BYTES, 0, b"sibling text"));

        let records: Vec<_> = RecordIter::new(&stream)
            .collect::<std::result::Result<_, _>>()
            .unwrap();
        assert_eq!(
            records.len(),
            2,
            "corrupt length nested inside the container must not swallow the top-level sibling after it"
        );
        assert_eq!(records[0].header.rec_type, RT_SLIDE);
        assert_eq!(records[1].header.rec_type, RT_TEXT_BYTES);
        assert_eq!(records[1].data, b"sibling text");

        // Descending into the corrupt container's own children must clamp to
        // the bytes actually available, not panic or run past the container.
        let inner: Vec<_> = RecordIter::new(&records[0].data)
            .collect::<std::result::Result<_, _>>()
            .unwrap();
        assert_eq!(inner.len(), 1);
        assert!(inner[0].data.len() < 1_000_000);
    }
}
