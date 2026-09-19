//! `.xls` cell comments: `NOTE` records (BIFF `0x001C`) carry the cell
//! position and author; the actual comment text lives in a `TXO` record
//! (`0x01B6`) attached to an `OBJ` record (`0x005D`) declaring the
//! comment's drawing shape. None of the three had any record handling at
//! all before this module.
//!
//! Byte layouts verified against Apache POI's `NoteRecord`,
//! `TextObjectRecord`, and `CommonObjectDataSubRecord` (a mature,
//! independent implementation of the same [MS-XLS] structures).
//!
//! Correlating a `NOTE` to its text needs one extra hop: `NOTE::shapeid`
//! doesn't point at the `TXO` directly, it points at the *object id* an
//! `OBJ` record declares in its own `ftCmo` sub-record — and a `TXO` is
//! always immediately preceded by the `OBJ` record for the same shape
//! ([MS-XLS] itself requires this ordering). Rather than parsing Escher/
//! `MSODRAWING` shape records to establish that link independently, this
//! reuses the ordering guarantee: `xls::workbook` tracks "the most
//! recently seen `OBJ`'s id" and records `(id -> TXO text)` whenever a
//! `TXO` immediately follows one. `NOTE` records are collected separately
//! and resolved against that map once the whole sheet has been scanned,
//! since a `NOTE` isn't guaranteed to appear after its `OBJ`/`TXO` pair in
//! every file (real files commonly group all of a sheet's `NOTE` records
//! together near the end of the sheet's own record stream).
//!
//! Scope: `TXO`'s own rich-text formatting runs and its rare "linked to a
//! cell formula" variant (used by some textboxes/chart elements, not
//! ordinary comments) are not decoded — a comment's plain text is what's
//! needed here, matching the same tier of extraction the PPT picture and WordArt paths already
//! use for other embedded-text formats.

/// One resolved cell comment: position, author (if present), and text.
#[derive(Debug, Clone, PartialEq)]
pub struct XlsComment {
    /// 0-based row the comment is attached to.
    pub row: u16,
    /// 0-based column the comment is attached to.
    pub col: u16,
    /// Comment author, when the `NOTE` record names one.
    pub author: Option<String>,
    /// Comment body text.
    pub text: String,
}

/// Extract an `OBJ` record's object id from its leading `ftCmo`
/// (`CommonObjectDataSubRecord`) sub-record: `ft: u16` (must be `0x0015`,
/// required to be the first sub-record by spec), `cb: u16` (sub-record
/// payload length, unused here), `objectType: i16`, `objectId: u16` —
/// the id sits at byte offset 6.
pub fn obj_id(data: &[u8]) -> Option<u16> {
    if data.len() < 8 {
        return None;
    }
    let ft = u16::from_le_bytes([data[0], data[1]]);
    if ft != 0x0015 {
        return None;
    }
    Some(u16::from_le_bytes([data[6], data[7]]))
}

/// Extract a `TXO` record's plain text. `data` is the record's bytes
/// after `CONTINUE` merging (this crate's `RecordIter` already does this
/// generically), matching how a "ContinuableRecord" reader transparently
/// spans those boundaries in a reference implementation.
///
/// Fixed 18-byte header: `options/textOrientation/reserved4/reserved5/
/// reserved6: u16` each (offsets 0-9), `textLength: u16` (offset 10),
/// `formattingDataLength: u16` (offset 12, unused — formatting runs
/// aren't decoded), `reserved7: u32` (offset 14). Then, when
/// `textLength > 0`: a 1-byte compression flag (bit 0 clear =
/// compressed/Latin-1, set = UTF-16LE) at offset 18, followed by
/// `textLength` characters. A link-formula block can appear between the
/// header and the text for some non-comment `TXO` uses (linked textboxes/
/// chart titles); not handled here — see the module doc's scope note.
pub fn txo_text(data: &[u8]) -> Option<String> {
    if data.len() < 18 {
        return None;
    }
    let text_length = u16::from_le_bytes([data[10], data[11]]) as usize;
    if text_length == 0 {
        return Some(String::new());
    }
    let flag = *data.get(18)?;
    let is_compressed = flag & 0x01 == 0;
    let char_start = 19usize;
    if is_compressed {
        let end = char_start.checked_add(text_length)?;
        let bytes = data.get(char_start..end)?;
        Some(bytes.iter().map(|&b| b as char).collect())
    } else {
        let byte_len = text_length.checked_mul(2)?;
        let end = char_start.checked_add(byte_len)?;
        let bytes = data.get(char_start..end)?;
        let units: Vec<u16> = bytes
            .as_chunks::<2>()
            .0
            .iter()
            .map(|c| u16::from_le_bytes(*c))
            .collect();
        Some(String::from_utf16_lossy(&units))
    }
}

/// One `NOTE` record's own fields, before its `shapeid` has been resolved
/// to a `TXO`'s text: `(row, col, shapeid, author)`.
///
/// Layout: `row: u16`, `col: u16`, `flags: u16` (offset 4, unused — hidden/
/// visible state isn't carried into the IR), `shapeid: u16` (offset 6),
/// `author_len: u16` (offset 8, character count), `has_multibyte: u8`
/// (offset 10), then `author_len` author characters (compressed or
/// UTF-16LE per the multibyte flag).
pub fn parse_note(data: &[u8]) -> Option<(u16, u16, u16, Option<String>)> {
    if data.len() < 11 {
        return None;
    }
    let row = u16::from_le_bytes([data[0], data[1]]);
    let col = u16::from_le_bytes([data[2], data[3]]);
    let shapeid = u16::from_le_bytes([data[6], data[7]]);
    let author_len = u16::from_le_bytes([data[8], data[9]]) as usize;
    let has_multibyte = data[10] != 0;

    let author = if author_len == 0 {
        None
    } else if has_multibyte {
        let byte_len = author_len.checked_mul(2)?;
        let end = 11usize.checked_add(byte_len)?;
        let bytes = data.get(11..end)?;
        let units: Vec<u16> = bytes
            .as_chunks::<2>()
            .0
            .iter()
            .map(|c| u16::from_le_bytes(*c))
            .collect();
        Some(String::from_utf16_lossy(&units))
    } else {
        let end = 11usize.checked_add(author_len)?;
        let bytes = data.get(11..end)?;
        Some(bytes.iter().map(|&b| b as char).collect())
    };

    Some((row, col, shapeid, author))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn obj_record(object_id: u16) -> Vec<u8> {
        let mut d = Vec::new();
        d.extend_from_slice(&0x0015u16.to_le_bytes()); // ft = ftCmo
        d.extend_from_slice(&18u16.to_le_bytes()); // cb
        d.extend_from_slice(&0x0019u16.to_le_bytes()); // objectType = Note
        d.extend_from_slice(&object_id.to_le_bytes());
        d.extend_from_slice(&[0u8; 14]); // option + reserved1..3
        d
    }

    fn txo_record(text: &str, compressed: bool) -> Vec<u8> {
        let mut d = Vec::new();
        d.extend_from_slice(&0u16.to_le_bytes()); // options
        d.extend_from_slice(&0u16.to_le_bytes()); // textOrientation
        d.extend_from_slice(&0u16.to_le_bytes()); // reserved4
        d.extend_from_slice(&0u16.to_le_bytes()); // reserved5
        d.extend_from_slice(&0u16.to_le_bytes()); // reserved6
        d.extend_from_slice(&(text.chars().count() as u16).to_le_bytes()); // textLength
        d.extend_from_slice(&0u16.to_le_bytes()); // formattingDataLength
        d.extend_from_slice(&0u32.to_le_bytes()); // reserved7
        if !text.is_empty() {
            if compressed {
                d.push(0x00);
                d.extend_from_slice(text.as_bytes());
            } else {
                d.push(0x01);
                d.extend(text.encode_utf16().flat_map(|u| u.to_le_bytes()));
            }
        }
        d
    }

    fn note_record(row: u16, col: u16, shapeid: u16, author: Option<&str>) -> Vec<u8> {
        let mut d = Vec::new();
        d.extend_from_slice(&row.to_le_bytes());
        d.extend_from_slice(&col.to_le_bytes());
        d.extend_from_slice(&0u16.to_le_bytes()); // flags
        d.extend_from_slice(&shapeid.to_le_bytes());
        let author = author.unwrap_or("");
        d.extend_from_slice(&(author.len() as u16).to_le_bytes());
        d.push(0x00); // has_multibyte = false
        d.extend_from_slice(author.as_bytes());
        d
    }

    #[test]
    fn test_obj_id_reads_the_ftcmo_object_id() {
        assert_eq!(obj_id(&obj_record(42)), Some(42));
    }

    #[test]
    fn test_obj_id_rejects_a_non_ftcmo_first_subrecord() {
        let mut d = obj_record(1);
        d[0] = 0xFF; // corrupt the ft field
        assert_eq!(obj_id(&d), None);
    }

    #[test]
    fn test_txo_text_compressed_round_trips() {
        assert_eq!(
            txo_text(&txo_record("Schroeder: a. GWA", true)).as_deref(),
            Some("Schroeder: a. GWA")
        );
    }

    #[test]
    fn test_txo_text_uncompressed_round_trips() {
        assert_eq!(txo_text(&txo_record("caf\u{e9}", false)).as_deref(), Some("caf\u{e9}"));
    }

    #[test]
    fn test_txo_text_empty_is_empty_string_not_none() {
        assert_eq!(txo_text(&txo_record("", true)).as_deref(), Some(""));
    }

    #[test]
    fn test_parse_note_extracts_row_col_shapeid_and_author() {
        let (row, col, shapeid, author) =
            parse_note(&note_record(5, 2, 42, Some("Elemar"))).unwrap();
        assert_eq!((row, col, shapeid), (5, 2, 42));
        assert_eq!(author.as_deref(), Some("Elemar"));
    }

    #[test]
    fn test_parse_note_with_no_author_is_none() {
        let (_, _, _, author) = parse_note(&note_record(0, 0, 1, None)).unwrap();
        assert!(author.is_none());
    }

    #[test]
    fn test_truncated_records_are_none_not_a_panic() {
        assert!(obj_id(&[0u8; 4]).is_none());
        assert!(txo_text(&[0u8; 10]).is_none());
        assert!(parse_note(&[0u8; 5]).is_none());
    }
}
