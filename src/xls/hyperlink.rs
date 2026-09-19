//! `.xls` cell hyperlinks: `HLINK` records (BIFF `0x01B8`). No record
//! handling existed at all before this module — the hyperlink target
//! behind a cell (its own display text, e.g. `"Stacie@ABC.com"`, always
//! survived as an ordinary string cell) was completely invisible, not
//! merely dropped in conversion.
//!
//! Byte layout verified against Apache POI's `HyperlinkRecord` (a mature,
//! independent implementation of the same structure, which itself
//! implements [MS-OSHARED] §2.3.7.9 "Hyperlink Object" embedded inside
//! the BIFF record). The moniker-resolution logic (which of `address`/
//! `short_filename`/`text_mark` wins as the effective target) is ported
//! byte-for-byte from POI's own `getAddress()`, including its documented
//! simplification: when both a URL/file moniker AND a "place in
//! document" text mark are present on the same link, the text mark wins
//! and the moniker's own address is discarded. That combination doesn't
//! occur in any of the corpus files this was verified against (plain
//! external `http:`/`mailto:` links, no in-document fragment), so the
//! simplification is inherited rather than independently re-derived.
//!
//! Scope: `STD_MONIKER` links (raw absolute local paths, not a URL or a
//! `..\` relative file path) are rare enough in practice that POI itself
//! notes uncertainty about the format; parsed the same way POI parses it
//! (length-prefixed raw UTF-8 bytes) since the shape is simple and cheap
//! to support once everything else is written, but not independently
//! verified against a real corpus example.

/// One decoded `HLINK` record: the cell range it covers plus its
/// effective target string (already reads through the
/// moniker/UNC/place-mark cases, mirroring POI's `getAddress()`).
#[derive(Debug, Clone, PartialEq)]
pub struct XlsHyperlink {
    /// First row of the covered range (0-based).
    pub row_first: u16,
    /// Last row of the covered range (0-based, inclusive).
    pub row_last: u16,
    /// First column of the covered range (0-based).
    pub col_first: u16,
    /// Last column of the covered range (0-based, inclusive).
    pub col_last: u16,
    /// The effective link target (URL, file path, or mailto: address).
    pub target: String,
}

const HLINK_URL: u32 = 0x01;
const HLINK_PLACE: u32 = 0x08;
const HLINK_LABEL: u32 = 0x14;
const HLINK_TARGET_FRAME: u32 = 0x80;
const HLINK_UNC_PATH: u32 = 0x100;

/// [MS-OSHARED] §2.3.7.9 moniker class GUIDs, in the file's own
/// little-endian-per-field byte order (not the canonical big-endian
/// string-form order) — this is the raw 16 bytes as they appear on disk.
const URL_MONIKER: [u8; 16] = [
    0xE0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11, 0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B, 0xA9, 0x0B,
];
const FILE_MONIKER: [u8; 16] = [
    0x03, 0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46,
];
const STD_MONIKER: [u8; 16] = [
    0xD0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11, 0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B, 0xA9, 0x0B,
];
/// Trailing bytes after a URL_MONIKER address when the record wasn't
/// trimmed to exactly the string's own length.
const TAIL_SIZE: usize = 24;

fn read_u16(data: &[u8], pos: usize) -> Option<u16> {
    data.get(pos..pos + 2)
        .map(|b| u16::from_le_bytes([b[0], b[1]]))
}

fn read_u32(data: &[u8], pos: usize) -> Option<u32> {
    data.get(pos..pos + 4)
        .map(|b| u32::from_le_bytes([b[0], b[1], b[2], b[3]]))
}

/// Read `n_chars` UTF-16LE code units starting at `pos`.
fn read_utf16le(data: &[u8], pos: usize, n_chars: usize) -> Option<(String, usize)> {
    let byte_len = n_chars.checked_mul(2)?;
    let end = pos.checked_add(byte_len)?;
    let bytes = data.get(pos..end)?;
    let units: Vec<u16> = bytes
        .as_chunks::<2>()
        .0
        .iter()
        .map(|c| u16::from_le_bytes(*c))
        .collect();
    Some((String::from_utf16_lossy(&units), end))
}

/// Read `n_bytes` single-byte (Latin-1/compressed) characters starting at
/// `pos`.
fn read_compressed(data: &[u8], pos: usize, n_bytes: usize) -> Option<(String, usize)> {
    let end = pos.checked_add(n_bytes)?;
    let bytes = data.get(pos..end)?;
    Some((bytes.iter().map(|&b| b as char).collect(), end))
}

/// Truncate at the first embedded NUL — these strings are commonly
/// null-terminated even though their length is also given explicitly.
fn clean(mut s: String) -> String {
    if let Some(idx) = s.find('\0') {
        s.truncate(idx);
    }
    s
}

/// Parse one `HLINK` record. Returns `None` for a truncated record, an
/// unrecognized moniker with no address to fall back on, or a
/// `streamVersion` other than 2 (a different structure shape this parser
/// doesn't understand, not a corruption worth reporting further up).
pub fn parse_hlink(data: &[u8]) -> Option<XlsHyperlink> {
    if data.len() < 8 {
        return None;
    }
    let row_first = read_u16(data, 0)?;
    let row_last = read_u16(data, 2)?;
    let col_first = read_u16(data, 4)?;
    let col_last = read_u16(data, 6)?;

    let mut pos = 8 + 16; // skip the leading 16-byte GUID
    let stream_version = read_u32(data, pos)?;
    if stream_version != 2 {
        return None;
    }
    pos += 4;
    let link_opts = read_u32(data, pos)?;
    pos += 4;

    if link_opts & HLINK_LABEL != 0 {
        let len = read_u32(data, pos)? as usize;
        pos += 4;
        let (_, new_pos) = read_utf16le(data, pos, len)?;
        pos = new_pos;
    }

    if link_opts & HLINK_TARGET_FRAME != 0 {
        let len = read_u32(data, pos)? as usize;
        pos += 4;
        let (_, new_pos) = read_utf16le(data, pos, len)?;
        pos = new_pos;
    }

    let mut moniker_is_file = false;
    let mut address: Option<String> = None;
    let mut short_filename: Option<String> = None;

    if link_opts & HLINK_URL != 0 && link_opts & HLINK_UNC_PATH != 0 {
        let n_chars = read_u32(data, pos)? as usize;
        pos += 4;
        let (s, new_pos) = read_utf16le(data, pos, n_chars)?;
        address = Some(s);
        pos = new_pos;
    } else if link_opts & HLINK_URL != 0 {
        let moniker: [u8; 16] = data.get(pos..pos + 16)?.try_into().ok()?;
        pos += 16;

        if moniker == URL_MONIKER {
            let length = read_u32(data, pos)? as usize;
            pos += 4;
            let remaining = data.len().checked_sub(pos)?;
            let (n_chars, has_tail) = if length == remaining {
                (length / 2, false)
            } else {
                (length.checked_sub(TAIL_SIZE)? / 2, true)
            };
            let (s, new_pos) = read_utf16le(data, pos, n_chars)?;
            address = Some(s);
            pos = new_pos;
            if has_tail {
                pos = pos.checked_add(TAIL_SIZE)?;
            }
        } else if moniker == FILE_MONIKER {
            moniker_is_file = true;
            pos += 2; // fileOpts: u16, unused
            let len = read_u32(data, pos)? as usize;
            pos += 4;
            let (s, new_pos) = read_compressed(data, pos, len)?;
            short_filename = Some(s);
            pos = new_pos;
            pos = pos.checked_add(TAIL_SIZE)?;
            let size = read_u32(data, pos)? as usize;
            pos += 4;
            if size > 0 {
                let char_data_size = read_u32(data, pos)? as usize;
                pos += 4;
                pos += 2; // reserved u16
                let (s, new_pos) = read_utf16le(data, pos, char_data_size / 2)?;
                address = Some(s);
                pos = new_pos;
            }
        } else if moniker == STD_MONIKER {
            pos += 2; // fileOpts: u16, unused
            let len = read_u32(data, pos)? as usize;
            pos += 4;
            let bytes = data.get(pos..pos.checked_add(len)?)?;
            address = Some(String::from_utf8_lossy(bytes).into_owned());
            pos += len;
        } else {
            // Unrecognized moniker: nothing further in the record can be
            // located reliably relative to it.
            return None;
        }
    }

    let mut text_mark: Option<String> = None;
    if link_opts & HLINK_PLACE != 0 {
        if let Some(len) = read_u32(data, pos) {
            pos += 4;
            if let Some((s, new_pos)) = read_utf16le(data, pos, len as usize) {
                text_mark = Some(s);
                pos = new_pos;
            }
        }
    }
    let _ = pos;

    let target = if link_opts & HLINK_URL != 0 && moniker_is_file {
        address.or(short_filename)
    } else if link_opts & HLINK_PLACE != 0 {
        text_mark
    } else {
        address
    }?;
    let target = clean(target);
    if target.is_empty() {
        return None;
    }

    Some(XlsHyperlink {
        row_first,
        row_last,
        col_first,
        col_last,
        target,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn range_bytes(rf: u16, rl: u16, cf: u16, cl: u16) -> Vec<u8> {
        let mut d = Vec::new();
        d.extend_from_slice(&rf.to_le_bytes());
        d.extend_from_slice(&rl.to_le_bytes());
        d.extend_from_slice(&cf.to_le_bytes());
        d.extend_from_slice(&cl.to_le_bytes());
        d
    }

    fn utf16le_bytes(s: &str) -> Vec<u8> {
        s.encode_utf16().flat_map(|u| u.to_le_bytes()).collect()
    }

    /// A UNC-path-shaped record (`link_opts` has both URL and UNC_PATH
    /// set): the plainest possible shape, no moniker at all. Matches how
    /// a plain external URL like `http://example.com` is commonly stored.
    #[test]
    fn test_unc_path_style_url_is_decoded() {
        let mut d = range_bytes(1, 1, 2, 2);
        d.extend_from_slice(&[0u8; 16]); // guid
        d.extend_from_slice(&2u32.to_le_bytes()); // streamVersion
        let link_opts: u32 = HLINK_URL | HLINK_UNC_PATH;
        d.extend_from_slice(&link_opts.to_le_bytes());
        let url = "http://example.com/page\0";
        d.extend_from_slice(&(url.encode_utf16().count() as u32).to_le_bytes());
        d.extend_from_slice(&utf16le_bytes(url));

        let hl = parse_hlink(&d).unwrap();
        assert_eq!(hl.target, "http://example.com/page");
        assert_eq!((hl.row_first, hl.row_last, hl.col_first, hl.col_last), (1, 1, 2, 2));
    }

    /// URL_MONIKER shape, no trailing tail bytes (`length == remaining`).
    #[test]
    fn test_url_moniker_without_tail_is_decoded() {
        let mut d = range_bytes(0, 0, 0, 0);
        d.extend_from_slice(&[0u8; 16]); // guid
        d.extend_from_slice(&2u32.to_le_bytes());
        let link_opts: u32 = HLINK_URL;
        d.extend_from_slice(&link_opts.to_le_bytes());
        d.extend_from_slice(&URL_MONIKER);
        let addr = "mailto:Stacie@ABC.com\0";
        let addr_bytes = utf16le_bytes(addr);
        d.extend_from_slice(&(addr_bytes.len() as u32).to_le_bytes());
        d.extend_from_slice(&addr_bytes);

        let hl = parse_hlink(&d).unwrap();
        assert_eq!(hl.target, "mailto:Stacie@ABC.com");
    }

    /// URL_MONIKER shape WITH the 24-byte tail present (`length` covers
    /// the tail too, so `length != remaining`).
    #[test]
    fn test_url_moniker_with_tail_is_decoded() {
        let mut d = range_bytes(3, 3, 0, 0);
        d.extend_from_slice(&[0u8; 16]);
        d.extend_from_slice(&2u32.to_le_bytes());
        let link_opts: u32 = HLINK_URL;
        d.extend_from_slice(&link_opts.to_le_bytes());
        d.extend_from_slice(&URL_MONIKER);
        let addr = "http://www.ozgrid.com/\0";
        let addr_bytes = utf16le_bytes(addr);
        d.extend_from_slice(&((addr_bytes.len() + TAIL_SIZE) as u32).to_le_bytes());
        d.extend_from_slice(&addr_bytes);
        d.extend_from_slice(&[0xAAu8; TAIL_SIZE]);

        let hl = parse_hlink(&d).unwrap();
        assert_eq!(hl.target, "http://www.ozgrid.com/");
    }

    /// FILE_MONIKER shape, no extended (relative) address — falls back to
    /// the 8.3 short filename.
    #[test]
    fn test_file_moniker_falls_back_to_short_filename() {
        let mut d = range_bytes(0, 0, 0, 0);
        d.extend_from_slice(&[0u8; 16]);
        d.extend_from_slice(&2u32.to_le_bytes());
        let link_opts: u32 = HLINK_URL;
        d.extend_from_slice(&link_opts.to_le_bytes());
        d.extend_from_slice(&FILE_MONIKER);
        d.extend_from_slice(&0u16.to_le_bytes()); // fileOpts
        let short_name = "REPORT~1.XLS\0";
        d.extend_from_slice(&(short_name.len() as u32).to_le_bytes());
        d.extend_from_slice(short_name.as_bytes());
        d.extend_from_slice(&[0u8; TAIL_SIZE]);
        d.extend_from_slice(&0u32.to_le_bytes()); // size == 0, no extended address

        let hl = parse_hlink(&d).unwrap();
        assert_eq!(hl.target, "REPORT~1.XLS");
    }

    /// FILE_MONIKER shape WITH an extended address — the long-path
    /// Unicode form wins over the short filename.
    #[test]
    fn test_file_moniker_extended_address_wins_over_short_filename() {
        let mut d = range_bytes(0, 0, 0, 0);
        d.extend_from_slice(&[0u8; 16]);
        d.extend_from_slice(&2u32.to_le_bytes());
        let link_opts: u32 = HLINK_URL;
        d.extend_from_slice(&link_opts.to_le_bytes());
        d.extend_from_slice(&FILE_MONIKER);
        d.extend_from_slice(&0u16.to_le_bytes());
        let short_name = "REPORT~1.XLS\0";
        d.extend_from_slice(&(short_name.len() as u32).to_le_bytes());
        d.extend_from_slice(short_name.as_bytes());
        d.extend_from_slice(&[0u8; TAIL_SIZE]);
        let ext_addr = "Annual Report.xls\0";
        let ext_bytes = utf16le_bytes(ext_addr);
        d.extend_from_slice(&((4 + ext_bytes.len()) as u32).to_le_bytes()); // size > 0
        d.extend_from_slice(&(ext_bytes.len() as u32).to_le_bytes()); // charDataSize
        d.extend_from_slice(&3u16.to_le_bytes()); // usKeyValue
        d.extend_from_slice(&ext_bytes);

        let hl = parse_hlink(&d).unwrap();
        assert_eq!(hl.target, "Annual Report.xls");
    }

    #[test]
    fn test_truncated_record_is_none_not_a_panic() {
        assert!(parse_hlink(&[0u8; 8]).is_none());
        assert!(parse_hlink(&[]).is_none());
        let mut d = range_bytes(0, 0, 0, 0);
        d.extend_from_slice(&[0u8; 16]);
        d.extend_from_slice(&2u32.to_le_bytes());
        assert!(parse_hlink(&d).is_none()); // link_opts truncated
    }

    #[test]
    fn test_wrong_stream_version_is_none() {
        let mut d = range_bytes(0, 0, 0, 0);
        d.extend_from_slice(&[0u8; 16]);
        d.extend_from_slice(&1u32.to_le_bytes()); // not 2
        assert!(parse_hlink(&d).is_none());
    }

    #[test]
    fn test_unrecognized_moniker_is_none() {
        let mut d = range_bytes(0, 0, 0, 0);
        d.extend_from_slice(&[0u8; 16]);
        d.extend_from_slice(&2u32.to_le_bytes());
        let link_opts: u32 = HLINK_URL;
        d.extend_from_slice(&link_opts.to_le_bytes());
        d.extend_from_slice(&[0x42u8; 16]); // not a known moniker
        assert!(parse_hlink(&d).is_none());
    }
}
