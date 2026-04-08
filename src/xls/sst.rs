//! Shared String Table (SST) parsing for BIFF8.
//!
//! The SST record contains all unique strings referenced by cells.
//! Strings can be compressed (Latin-1) or uncompressed (UTF-16LE).

use super::error::{Result, XlsError};

/// Parse the SST record data (already merged with CONTINUE records).
///
/// Returns a vector of strings.
pub fn parse_sst(data: &[u8]) -> Result<Vec<String>> {
    if data.len() < 8 {
        return Err(XlsError::Corrupted("SST too short".into()));
    }

    // Total string count (appearances) at offset 0 (u32).
    // Unique string count at offset 4 (u32).
    let unique_count = u32::from_le_bytes([data[4], data[5], data[6], data[7]]) as usize;

    let mut strings = Vec::with_capacity(unique_count.min(100_000));
    let mut pos = 8;

    for _ in 0..unique_count {
        if pos >= data.len() {
            break;
        }
        match read_unicode_string(data, pos) {
            Ok((s, new_pos)) => {
                strings.push(s);
                pos = new_pos;
            }
            Err(_) => break, // Tolerate truncated SST
        }
    }

    Ok(strings)
}

/// Read a BIFF8 Unicode string from the data at the given position.
///
/// BIFF8 strings: [char_count: u16][flags: u8][optional rt_count: u16][optional ext_size: u32][chars][rich text runs][ext data]
///
/// Returns (string, new_position).
pub fn read_unicode_string(data: &[u8], pos: usize) -> Result<(String, usize)> {
    if pos + 3 > data.len() {
        return Err(XlsError::Corrupted(format!(
            "unicode string header truncated at {pos}"
        )));
    }

    let char_count = u16::from_le_bytes([data[pos], data[pos + 1]]) as usize;
    let flags = data[pos + 2];
    let mut offset = pos + 3;

    // Sanity check: char_count shouldn't exceed remaining data.
    let max_possible = (data.len() - offset) * 2; // generous upper bound
    if char_count > max_possible {
        return Err(XlsError::Corrupted(format!(
            "string char_count {char_count} exceeds data at {pos}"
        )));
    }

    let is_wide = (flags & 0x01) != 0; // 16-bit characters
    let has_rich = (flags & 0x08) != 0; // rich text runs follow
    let has_ext = (flags & 0x04) != 0; // extended (Far East) data follows

    let rt_count = if has_rich {
        if offset + 2 > data.len() {
            return Err(XlsError::Corrupted("rich text count truncated".into()));
        }
        let n = u16::from_le_bytes([data[offset], data[offset + 1]]) as usize;
        offset += 2;
        n
    } else {
        0
    };

    let ext_size = if has_ext {
        if offset + 4 > data.len() {
            return Err(XlsError::Corrupted("ext size truncated".into()));
        }
        let n = u32::from_le_bytes([data[offset], data[offset + 1], data[offset + 2], data[offset + 3]]) as usize;
        offset += 4;
        n
    } else {
        0
    };

    // Read the character data.
    let s = if is_wide {
        let byte_len = char_count * 2;
        if offset + byte_len > data.len() {
            return Err(XlsError::Corrupted(format!(
                "wide string data truncated at {offset}, need {byte_len}, have {}",
                data.len() - offset
            )));
        }
        let chars: Vec<u16> = (0..char_count)
            .map(|i| {
                let o = offset + i * 2;
                u16::from_le_bytes([data[o], data[o + 1]])
            })
            .collect();
        offset += byte_len;
        String::from_utf16_lossy(&chars)
    } else {
        // Compressed: 1 byte per char, Latin-1.
        if offset + char_count > data.len() {
            return Err(XlsError::Corrupted(format!(
                "compressed string truncated at {offset}"
            )));
        }
        let s: String = data[offset..offset + char_count]
            .iter()
            .map(|&b| b as char)
            .collect();
        offset += char_count;
        s
    };

    // Skip rich text formatting runs (4 bytes each).
    offset += rt_count * 4;

    // Skip extended data.
    offset += ext_size;

    Ok((s, offset))
}

/// Read a short Unicode string (1-byte char count, used in BOUNDSHEET etc.)
pub fn read_short_unicode_string(data: &[u8], pos: usize) -> Result<(String, usize)> {
    if pos + 2 > data.len() {
        return Err(XlsError::Corrupted("short string truncated".into()));
    }

    let char_count = data[pos] as usize;
    let flags = data[pos + 1];
    let is_wide = (flags & 0x01) != 0;
    let mut offset = pos + 2;

    let s = if is_wide {
        let byte_len = char_count * 2;
        if offset + byte_len > data.len() {
            return Err(XlsError::Corrupted("short wide string truncated".into()));
        }
        let chars: Vec<u16> = (0..char_count)
            .map(|i| {
                let o = offset + i * 2;
                u16::from_le_bytes([data[o], data[o + 1]])
            })
            .collect();
        offset += byte_len;
        String::from_utf16_lossy(&chars)
    } else {
        if offset + char_count > data.len() {
            return Err(XlsError::Corrupted(
                "short compressed string truncated".into(),
            ));
        }
        let s: String = data[offset..offset + char_count]
            .iter()
            .map(|&b| b as char)
            .collect();
        offset += char_count;
        s
    };

    Ok((s, offset))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_compressed_sst() {
        // SST: total=2, unique=2, then two compressed strings "AB" and "CD".
        let mut data = Vec::new();
        data.extend_from_slice(&2u32.to_le_bytes()); // total
        data.extend_from_slice(&2u32.to_le_bytes()); // unique
        // String "AB": char_count=2, flags=0 (compressed, no rich, no ext)
        data.extend_from_slice(&2u16.to_le_bytes());
        data.push(0x00); // flags
        data.push(b'A');
        data.push(b'B');
        // String "CD"
        data.extend_from_slice(&2u16.to_le_bytes());
        data.push(0x00);
        data.push(b'C');
        data.push(b'D');

        let strings = parse_sst(&data).unwrap();
        assert_eq!(strings, vec!["AB", "CD"]);
    }

    #[test]
    fn parse_wide_sst() {
        let mut data = Vec::new();
        data.extend_from_slice(&1u32.to_le_bytes());
        data.extend_from_slice(&1u32.to_le_bytes());
        // Wide string "Hi" (UTF-16LE)
        data.extend_from_slice(&2u16.to_le_bytes()); // 2 chars
        data.push(0x01); // flags = wide
        data.extend_from_slice(&b'H'.to_le_bytes()); // 'H' as u16 LE
        data.push(0x00);
        data.extend_from_slice(&b'i'.to_le_bytes());
        data.push(0x00);

        let strings = parse_sst(&data).unwrap();
        assert_eq!(strings, vec!["Hi"]);
    }

    #[test]
    fn parse_sst_with_rich_text() {
        let mut data = Vec::new();
        data.extend_from_slice(&1u32.to_le_bytes());
        data.extend_from_slice(&1u32.to_le_bytes());
        // String with rich text: "AB"
        data.extend_from_slice(&2u16.to_le_bytes()); // 2 chars
        data.push(0x08); // flags = has_rich (compressed)
        data.extend_from_slice(&1u16.to_le_bytes()); // 1 rich text run
        data.push(b'A');
        data.push(b'B');
        // Rich text run (4 bytes, we skip it)
        data.extend_from_slice(&[0x00, 0x00, 0x01, 0x00]);

        let strings = parse_sst(&data).unwrap();
        assert_eq!(strings, vec!["AB"]);
    }

    #[test]
    fn read_short_compressed() {
        // "Test" as short string: len=4, flags=0, "Test"
        let data = [4, 0x00, b'T', b'e', b's', b't'];
        let (s, pos) = read_short_unicode_string(&data, 0).unwrap();
        assert_eq!(s, "Test");
        assert_eq!(pos, 6);
    }

    #[test]
    fn read_short_wide() {
        let mut data = vec![2u8, 0x01]; // 2 chars, wide
        data.extend_from_slice(&(b'O' as u16).to_le_bytes());
        data.extend_from_slice(&(b'K' as u16).to_le_bytes());
        let (s, pos) = read_short_unicode_string(&data, 0).unwrap();
        assert_eq!(s, "OK");
        assert_eq!(pos, 6);
    }
}
