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
pub const RT_TEXT_CHARS: u16 = 0x0FA0;
pub const RT_TEXT_BYTES: u16 = 0x0FA8;
pub const RT_SLIDE_PERSIST_ATOM: u16 = 0x03F3;
pub const RT_USER_EDIT_ATOM: u16 = 0x0FF5;
pub const RT_PERSIST_DIR: u16 = 0x1772;
pub const RT_PERSIST_DIR_ENTRY: u16 = 0x1772;
pub const RT_CURRENT_USER_ATOM: u16 = 0x0FF6;
pub const RT_HEADER_FOOTER: u16 = 0x0FDA;
pub const RT_TEXT_SPECIAL_INFO: u16 = 0x0FAA;
pub const RT_TEXT_RULER: u16 = 0x0FA2;
pub const RT_STYLE_TEXT_PROP: u16 = 0x0FA1;
pub const RT_CSTRING: u16 = 0x0FBA;

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

/// Iterate over records in a byte stream (flat, non-recursive).
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

        let data_start = self.pos + 8;
        let data_end = data_start + header.rec_len as usize;
        let record_offset = self.pos;

        if header.is_container() {
            // For containers, advance past header only — children follow.
            self.pos = data_start;
        } else {
            // For atoms, skip over the data.
            if data_end > self.data.len() {
                // Truncated record — skip to end.
                self.pos = self.data.len();
            } else {
                self.pos = data_end;
            }
        }

        let data = if header.is_container() {
            Vec::new()
        } else if data_end <= self.data.len() {
            self.data[data_start..data_end].to_vec()
        } else {
            self.data[data_start..].to_vec()
        };

        Some(Ok(PptRecord {
            header,
            data,
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
    fn iterate_container_with_children() {
        let child1 = make_atom(RT_TEXT_HEADER, 0, &[0x00, 0x00, 0x00, 0x00]);
        let child2 = make_atom(RT_TEXT_CHARS, 0, &[0x41, 0x00]);
        let mut children = child1.clone();
        children.extend(&child2);
        let container = make_container(RT_SLIDE_LIST_WITH_TEXT, 0, &children);

        let records: Vec<_> = RecordIter::new(&container)
            .collect::<std::result::Result<_, _>>()
            .unwrap();
        // Container header + 2 children = 3 records.
        assert_eq!(records.len(), 3);
        assert_eq!(records[0].header.rec_type, RT_SLIDE_LIST_WITH_TEXT);
        assert!(records[0].header.is_container());
        assert_eq!(records[1].header.rec_type, RT_TEXT_HEADER);
        assert_eq!(records[2].header.rec_type, RT_TEXT_CHARS);
    }

    #[test]
    fn empty_stream() {
        let records: Vec<_> = RecordIter::new(&[])
            .collect::<std::result::Result<_, _>>()
            .unwrap();
        assert!(records.is_empty());
    }
}
