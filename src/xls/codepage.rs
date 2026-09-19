//! `.xls` `CODEPAGE` record (BIFF `0x0042`) and BIFF5 8-bit text
//! decoding.
//!
//! BIFF5 (pre-Unicode Excel) strings are raw 8-bit bytes in whatever
//! single-byte codepage the workbook declares (default Windows-1252 when
//! no `CODEPAGE` record is present, per [MS-XLS] §2.4.53). The `RT_CODEPAGE`
//! constant existed but was applied nowhere — every BIFF5 string byte was
//! promoted straight to its own code point (`b as char`, i.e. raw
//! ISO-8859-1/Latin-1), which is wrong even for the *default* codepage:
//! Windows-1252 and Latin-1 only agree below 0x80 and above 0x9F; the
//! 0x80-0x9F range (smart quotes, em dash, bullet, ellipsis — common
//! punctuation) differs between them (issue #309).
//!
//! Scope: BIFF8 strings are unaffected — their "compressed" flag means
//! "low byte of a UTF-16 code unit," an entirely different (and already
//! correct) meaning that has nothing to do with a codepage. Only BIFF5's
//! `LABEL`/`RSTRING` (and the duplicate fast-path in `workbook.rs`) go
//! through this module.

/// Parse a `CODEPAGE` record: a single `u16` LE codepage identifier.
pub fn parse_codepage(data: &[u8]) -> Option<u16> {
    if data.len() < 2 {
        return None;
    }
    Some(u16::from_le_bytes([data[0], data[1]]))
}

/// Decode raw BIFF5 8-bit text bytes using the workbook's declared
/// codepage (`None` defaults to Windows-1252, the BIFF5 default per
/// spec). Unrecognized codepages fall back to the previous raw
/// byte-as-code-point behavior rather than guessing wrong — no worse
/// than before, for whatever handful of real files use a codepage this
/// crate doesn't have a table for yet.
pub fn decode_biff5_text(bytes: &[u8], codepage: Option<u16>) -> String {
    let cp = codepage.unwrap_or(1252);
    if cp == MAC_CENTRAL_EUROPE {
        return bytes.iter().map(|&b| mac_central_europe_char(b)).collect();
    }
    if let Some(encoding) = encoding_for_codepage(cp) {
        let (decoded, _, _) = encoding.decode(bytes);
        return decoded.into_owned();
    }
    bytes.iter().map(|&b| b as char).collect()
}

/// Map a Windows/Macintosh codepage identifier to the `encoding_rs`
/// encoding that decodes it. Only single-byte legacy codepages this
/// crate has verified against a real corpus file are included; anything
/// else falls back to the raw byte-as-code-point behavior in
/// [`decode_biff5_text`].
fn encoding_for_codepage(cp: u16) -> Option<&'static encoding_rs::Encoding> {
    use encoding_rs::*;
    Some(match cp {
        1250 => WINDOWS_1250,
        1251 => WINDOWS_1251,
        1252 => WINDOWS_1252,
        1253 => WINDOWS_1253,
        1254 => WINDOWS_1254,
        1255 => WINDOWS_1255,
        1256 => WINDOWS_1256,
        1257 => WINDOWS_1257,
        1258 => WINDOWS_1258,
        874 => WINDOWS_874,
        866 => IBM866,
        10000 => MACINTOSH,
        _ => return None,
    })
}

const MAC_CENTRAL_EUROPE: u16 = 10029;

fn mac_central_europe_char(b: u8) -> char {
    if b < 0x80 { b as char } else { MAC_CENTRAL_EUROPE_TABLE[(b - 0x80) as usize] }
}

/// Mac OS Central European (codepage 10029) high-byte (0x80-0xFF) to
/// Unicode mapping. `encoding_rs` doesn't implement this encoding (it's
/// not in the WHATWG Encoding Standard's legacy set, unlike Mac Roman/
/// `MACINTOSH`), so it's a small hand-rolled table instead — verified
/// against the official Unicode.org mapping
/// (unicode.org/Public/MAPPINGS/VENDORS/APPLE/CENTEURO.TXT).
#[rustfmt::skip]
const MAC_CENTRAL_EUROPE_TABLE: [char; 128] = [
    '\u{00C4}', '\u{0100}', '\u{0101}', '\u{00C9}', '\u{0104}', '\u{00D6}', '\u{00DC}', '\u{00E1}',
    '\u{0105}', '\u{010C}', '\u{00E4}', '\u{010D}', '\u{0106}', '\u{0107}', '\u{00E9}', '\u{0179}',
    '\u{017A}', '\u{010E}', '\u{00ED}', '\u{010F}', '\u{0112}', '\u{0113}', '\u{0116}', '\u{00F3}',
    '\u{0117}', '\u{00F4}', '\u{00F6}', '\u{00F5}', '\u{00FA}', '\u{011A}', '\u{011B}', '\u{00FC}',
    '\u{2020}', '\u{00B0}', '\u{0118}', '\u{00A3}', '\u{00A7}', '\u{2022}', '\u{00B6}', '\u{00DF}',
    '\u{00AE}', '\u{00A9}', '\u{2122}', '\u{0119}', '\u{00A8}', '\u{2260}', '\u{0123}', '\u{012E}',
    '\u{012F}', '\u{012A}', '\u{2264}', '\u{2265}', '\u{012B}', '\u{0136}', '\u{2202}', '\u{2211}',
    '\u{0142}', '\u{013B}', '\u{013C}', '\u{013D}', '\u{013E}', '\u{0139}', '\u{013A}', '\u{0145}',
    '\u{0146}', '\u{0143}', '\u{00AC}', '\u{221A}', '\u{0144}', '\u{0147}', '\u{2206}', '\u{00AB}',
    '\u{00BB}', '\u{2026}', '\u{00A0}', '\u{0148}', '\u{0150}', '\u{00D5}', '\u{0151}', '\u{014C}',
    '\u{2013}', '\u{2014}', '\u{201C}', '\u{201D}', '\u{2018}', '\u{2019}', '\u{00F7}', '\u{25CA}',
    '\u{014D}', '\u{0154}', '\u{0155}', '\u{0158}', '\u{2039}', '\u{203A}', '\u{0159}', '\u{0156}',
    '\u{0157}', '\u{0160}', '\u{201A}', '\u{201E}', '\u{0161}', '\u{015A}', '\u{015B}', '\u{00C1}',
    '\u{0164}', '\u{0165}', '\u{00CD}', '\u{017D}', '\u{017E}', '\u{016A}', '\u{00D3}', '\u{00D4}',
    '\u{016B}', '\u{016E}', '\u{00DA}', '\u{016F}', '\u{0170}', '\u{0171}', '\u{0172}', '\u{0173}',
    '\u{00DD}', '\u{00FD}', '\u{0137}', '\u{017B}', '\u{0141}', '\u{017C}', '\u{0122}', '\u{02C7}',
];

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_codepage_reads_the_u16() {
        assert_eq!(parse_codepage(&[0xE5, 0x04]), Some(1253)); // 1253 = 0x04E5
    }

    #[test]
    fn truncated_codepage_is_none() {
        assert!(parse_codepage(&[0x01]).is_none());
    }

    #[test]
    fn greek_codepage_1253_decodes_correctly() {
        // Windows-1253 bytes for "Προμηθευτής" ("Supplier").
        let (encoded, _, _) = encoding_rs::WINDOWS_1253.encode("Προμηθευτής");
        assert_eq!(decode_biff5_text(&encoded, Some(1253)), "Προμηθευτής");
    }

    #[test]
    fn mac_central_europe_codepage_10029_decodes_correctly() {
        // Spot-check a few entries against the official CENTEURO.TXT
        // mapping, then decode a whole word through the public API.
        assert_eq!(mac_central_europe_char(0xFC), '\u{0141}'); // Ł
        assert_eq!(mac_central_europe_char(0x8D), '\u{0107}'); // ć
        assert_eq!(mac_central_europe_char(0x87), '\u{00E1}'); // á

        // "Ładowność" ("load capacity"): Ł a d o w n o ś ć
        let bytes = [0xFC, b'a', b'd', b'o', b'w', b'n', b'o', 0xE6, 0x8D];
        assert_eq!(decode_biff5_text(&bytes, Some(10029)), "Ładowność");
    }

    #[test]
    fn ascii_bytes_are_unchanged_regardless_of_codepage() {
        assert_eq!(decode_biff5_text(b"Hello", Some(1253)), "Hello");
        assert_eq!(decode_biff5_text(b"Hello", Some(10029)), "Hello");
        assert_eq!(decode_biff5_text(b"Hello", None), "Hello");
    }

    #[test]
    fn windows_1252_high_bytes_differ_from_raw_latin1() {
        // 0x93/0x94 are curly quotes in windows-1252, not the C1 control
        // codes ISO-8859-1/raw-Latin1 promotion would produce.
        let decoded = decode_biff5_text(&[0x93, b'x', 0x94], Some(1252));
        assert_ne!(decoded, "\u{0093}x\u{0094}");
        assert_eq!(decoded, "\u{201C}x\u{201D}");
    }

    #[test]
    fn unmapped_codepage_falls_back_to_raw_byte_promotion() {
        // No behavior regression for a codepage this crate doesn't know:
        // still the pre-existing raw-Latin1 behavior, not a panic or loss.
        assert_eq!(decode_biff5_text(&[0x41, 0xE9], Some(65535)), "A\u{00E9}");
    }

    #[test]
    fn default_codepage_is_windows_1252_not_raw_latin1() {
        let decoded = decode_biff5_text(&[0x93], None);
        assert_eq!(decoded, "\u{201C}");
    }
}
