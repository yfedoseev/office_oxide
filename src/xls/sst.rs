//! Shared String Table (SST) parsing for BIFF8.
//!
//! The SST record contains all unique strings referenced by cells.
//! Strings can be compressed (Latin-1) or uncompressed (UTF-16LE).

use super::error::{Result, XlsError};

/// Parse the SST record data (merged with its CONTINUE records) into the
/// table's strings.
///
/// `continue_at` holds the offsets into `data` where each `CONTINUE`
/// record began. A string whose *character data* is cut by one of those
/// boundaries resumes with a fresh option-flags byte ([MS-XLS] §2.5.293),
/// which may switch the rest of the string between 8-bit and 16-bit
/// characters; a cut anywhere else (between strings, inside a header,
/// inside the rich-text runs or extended data) inserts nothing. Reading
/// the merged bytes as one flat string left every such flags byte inside
/// the text, and the first string past the first boundary was misread —
/// so an SST of 20,000 strings yielded ~470, and every cell referencing
/// the rest came back empty.
pub fn parse_sst(data: &[u8], continue_at: &[usize]) -> Result<Vec<String>> {
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
        match read_unicode_string_across(data, pos, continue_at) {
            Ok((s, new_pos)) => {
                strings.push(s);
                pos = new_pos;
            },
            Err(_) => break, // Tolerate truncated SST
        }
    }

    Ok(strings)
}

/// As [`read_unicode_string`], for a string that may be cut by the
/// `CONTINUE` boundaries in `continue_at` (ascending offsets into `data`).
///
/// The header and the trailing rich-text/extended data are read as flat
/// bytes — a boundary there inserts nothing. Character data is read one
/// segment at a time: at each boundary with characters still owed, one
/// option-flags byte is consumed and its `fHighByte` bit decides the width
/// of the characters that follow.
pub fn read_unicode_string_across(
    data: &[u8],
    pos: usize,
    continue_at: &[usize],
) -> Result<(String, usize)> {
    if continue_at.is_empty() {
        return read_unicode_string(data, pos);
    }
    if pos + 3 > data.len() {
        return Err(XlsError::Corrupted(format!("unicode string header truncated at {pos}")));
    }
    let char_count = u16::from_le_bytes([data[pos], data[pos + 1]]) as usize;
    let flags = data[pos + 2];
    let mut offset = pos + 3;
    let has_rich = (flags & 0x08) != 0;
    let has_ext = (flags & 0x04) != 0;
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
        let n = u32::from_le_bytes([
            data[offset],
            data[offset + 1],
            data[offset + 2],
            data[offset + 3],
        ]) as usize;
        offset += 4;
        n
    } else {
        0
    };

    let mut is_wide = (flags & 0x01) != 0;
    let mut remaining = char_count;
    let mut out = String::with_capacity(char_count);
    while remaining > 0 {
        // The next boundary past the current offset, if any.
        let next_boundary = continue_at.iter().copied().find(|&b| b > offset);
        let segment_end = next_boundary.unwrap_or(data.len()).min(data.len());
        if offset >= data.len() {
            return Err(XlsError::Corrupted(format!("string data truncated at {offset}")));
        }
        let width = if is_wide { 2 } else { 1 };
        let chars_here = ((segment_end - offset) / width).min(remaining);
        if is_wide {
            let units: Vec<u16> = (0..chars_here)
                .map(|i| {
                    let o = offset + i * 2;
                    u16::from_le_bytes([data[o], data[o + 1]])
                })
                .collect();
            out.push_str(&String::from_utf16_lossy(&units));
        } else {
            out.extend(data[offset..offset + chars_here].iter().map(|&b| b as char));
        }
        offset += chars_here * width;
        remaining -= chars_here;
        if remaining > 0 {
            // Characters still owed, so the cut is inside this string's
            // character data: the CONTINUE opens with its own flags byte.
            // (A trailing odd byte before the boundary is padding.)
            let Some(boundary) = next_boundary else {
                return Err(XlsError::Corrupted(format!(
                    "string data truncated at {offset}, {remaining} chars owed"
                )));
            };
            if boundary >= data.len() {
                return Err(XlsError::Corrupted("string continues past the data".into()));
            }
            is_wide = (data[boundary] & 0x01) != 0;
            offset = boundary + 1;
        }
    }

    // Rich-text runs (4 bytes each) and extended data are flat bytes.
    offset += rt_count * 4;
    offset += ext_size;
    Ok((out, offset))
}

/// Read a BIFF8 Unicode string from the data at the given position.
///
/// BIFF8 strings: [char_count: u16][flags: u8][optional rt_count: u16][optional ext_size: u32][chars][rich text runs][ext data]
///
/// Returns (string, new_position).
pub fn read_unicode_string(data: &[u8], pos: usize) -> Result<(String, usize)> {
    if pos + 3 > data.len() {
        return Err(XlsError::Corrupted(format!("unicode string header truncated at {pos}")));
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
        let n = u32::from_le_bytes([
            data[offset],
            data[offset + 1],
            data[offset + 2],
            data[offset + 3],
        ]) as usize;
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
            return Err(XlsError::Corrupted(format!("compressed string truncated at {offset}")));
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
            return Err(XlsError::Corrupted("short compressed string truncated".into()));
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
    fn test_parse_compressed_sst() {
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

        let strings = parse_sst(&data, &[]).unwrap();
        assert_eq!(strings, vec!["AB", "CD"]);
    }

    #[test]
    fn test_parse_wide_sst() {
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

        let strings = parse_sst(&data, &[]).unwrap();
        assert_eq!(strings, vec!["Hi"]);
    }

    #[test]
    fn test_parse_sst_with_rich_text() {
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

        let strings = parse_sst(&data, &[]).unwrap();
        assert_eq!(strings, vec!["AB"]);
    }

    #[test]
    fn test_read_short_compressed() {
        // "Test" as short string: len=4, flags=0, "Test"
        let data = [4, 0x00, b'T', b'e', b's', b't'];
        let (s, pos) = read_short_unicode_string(&data, 0).unwrap();
        assert_eq!(s, "Test");
        assert_eq!(pos, 6);
    }

    #[test]
    fn test_read_short_wide() {
        let mut data = vec![2u8, 0x01]; // 2 chars, wide
        data.extend_from_slice(&(b'O' as u16).to_le_bytes());
        data.extend_from_slice(&(b'K' as u16).to_le_bytes());
        let (s, pos) = read_short_unicode_string(&data, 0).unwrap();
        assert_eq!(s, "OK");
        assert_eq!(pos, 6);
    }
}

#[cfg(test)]
mod continue_tests {
    use super::*;

    fn sst_header(unique: u32) -> Vec<u8> {
        let mut d = Vec::new();
        d.extend_from_slice(&unique.to_le_bytes()); // total (irrelevant here)
        d.extend_from_slice(&unique.to_le_bytes());
        d
    }
    fn compressed(text: &str) -> Vec<u8> {
        let mut d = Vec::new();
        d.extend_from_slice(&(text.len() as u16).to_le_bytes());
        d.push(0x00);
        d.extend_from_slice(text.as_bytes());
        d
    }

    /// Regression: a string cut by a `CONTINUE` boundary inside its
    /// character data resumes with a fresh flags byte ([MS-XLS] §2.5.293).
    /// Treating the merged bytes as flat text put that byte inside the
    /// string and misread every header after it — the corpus file that
    /// exposed this kept 474 of 20,812 strings.
    #[test]
    fn test_string_cut_inside_its_characters_resumes_after_the_flags_byte() {
        let mut data = sst_header(3);
        data.extend(compressed("first"));
        // "West Logan town": header + "West Log" in this record...
        data.extend_from_slice(&15u16.to_le_bytes());
        data.push(0x00);
        data.extend_from_slice(b"West Log");
        let boundary = data.len();
        // ...then the CONTINUE opens with its own flags byte, then "an town".
        data.push(0x00);
        data.extend_from_slice(b"an town");
        data.extend(compressed("third"));

        let strings = parse_sst(&data, &[boundary]).unwrap();
        assert_eq!(strings, ["first", "West Logan town", "third"]);
    }

    /// The flags byte at the boundary may switch the remainder to 16-bit
    /// characters.
    #[test]
    fn test_continuation_may_switch_to_wide_characters() {
        let mut data = sst_header(2);
        data.extend_from_slice(&6u16.to_le_bytes());
        data.push(0x00);
        data.extend_from_slice(b"caf");
        let boundary = data.len();
        data.push(0x01); // fHighByte: the rest is UTF-16LE
        data.extend("\u{e9}s!".encode_utf16().flat_map(|u| u.to_le_bytes()));
        data.extend(compressed("next"));

        let strings = parse_sst(&data, &[boundary]).unwrap();
        assert_eq!(strings, ["caf\u{e9}s!", "next"]);
    }

    /// A boundary that falls exactly between two strings inserts nothing:
    /// the next string's own header begins the `CONTINUE` record.
    #[test]
    fn test_boundary_between_strings_inserts_no_flags_byte() {
        let mut data = sst_header(2);
        data.extend(compressed("alpha"));
        let boundary = data.len();
        data.extend(compressed("beta"));
        let strings = parse_sst(&data, &[boundary]).unwrap();
        assert_eq!(strings, ["alpha", "beta"]);
    }

    /// A boundary inside the rich-text runs that trail a string inserts
    /// nothing either — only character data restarts with a flags byte.
    #[test]
    fn test_boundary_inside_rich_runs_inserts_no_flags_byte() {
        let mut data = sst_header(2);
        data.extend_from_slice(&4u16.to_le_bytes());
        data.push(0x08); // rich
        data.extend_from_slice(&2u16.to_le_bytes()); // two runs
        data.extend_from_slice(b"rich");
        data.extend_from_slice(&[0, 0, 5, 0]); // run 1
        let boundary = data.len();
        data.extend_from_slice(&[2, 0, 6, 0]); // run 2, after the cut
        data.extend(compressed("plain"));
        let strings = parse_sst(&data, &[boundary]).unwrap();
        assert_eq!(strings, ["rich", "plain"]);
    }

    /// Two boundaries inside one long string.
    #[test]
    fn test_string_spanning_two_continues() {
        let mut data = sst_header(1);
        data.extend_from_slice(&9u16.to_le_bytes());
        data.push(0x00);
        data.extend_from_slice(b"abc");
        let b1 = data.len();
        data.push(0x00);
        data.extend_from_slice(b"def");
        let b2 = data.len();
        data.push(0x00);
        data.extend_from_slice(b"ghi");
        let strings = parse_sst(&data, &[b1, b2]).unwrap();
        assert_eq!(strings, ["abcdefghi"]);
    }

    /// Owed characters with no further record: an error, not a panic or
    /// a silently short string.
    #[test]
    fn test_truncated_continuation_is_an_error() {
        let mut data = sst_header(1);
        data.extend_from_slice(&9u16.to_le_bytes());
        data.push(0x00);
        data.extend_from_slice(b"abc");
        assert!(read_unicode_string_across(&data, 8, &[2]).is_err());
    }
}
