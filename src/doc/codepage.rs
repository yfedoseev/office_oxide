//! DOC FIB `lid` (language ID) to codepage mapping, applied to
//! compressed (8-bit) text runs. Mirrors the XLS `CODEPAGE` fix, and
//! is source-confirmed rather than corpus-demonstrated:
//! `Fib` had no field consulting the document's language at all before
//! this, so any compressed non-Western-European text was corrupted by
//! construction, whether or not a real example was on hand to prove it.

/// A pseudo language ID for a document whose FIB marks a Macintosh character
/// set (`FibBase.chs` = 0x0100): its 8-bit text is Mac Roman whatever the
/// document's language. `0xFFFF` is not a Windows LCID, so it cannot collide
/// with a real `lid`; it travels wherever the `lid` does.
pub(crate) const MAC_TEXT: u16 = 0xFFFF;

/// Map a FIB `lid` (Windows LCID) to the codepage compressed text in
/// that language is stored in. Only the primary language id (the low 10
/// bits) determines the codepage — the sublanguage/region bits (e.g.
/// distinguishing Mexican from European Spanish) don't change it.
/// Defaults to 1252 (Western European) for English and any
/// unrecognized/neutral `lid`, matching Word's own default and every
/// codepage-unaware caller's existing behavior.
pub(crate) fn codepage_for_lid(lid: u16) -> u16 {
    if lid == MAC_TEXT {
        return 10000; // Mac Roman
    }
    let primary = lid & 0x03FF;
    match primary {
        // Russian, Bulgarian, Ukrainian, Belarusian.
        0x19 | 0x02 | 0x22 | 0x23 => 1251,
        // Polish, Czech, Slovak, Hungarian, Slovenian, Croatian, Romanian.
        0x15 | 0x05 | 0x1B | 0x0E | 0x24 | 0x1A | 0x18 => 1250,
        0x08 => 1253,        // Greek
        0x1F => 1254,        // Turkish
        0x0D => 1255,        // Hebrew
        0x01 => 1256,        // Arabic
        0x25..=0x27 => 1257, // Estonian, Latvian, Lithuanian
        0x2A => 1258,        // Vietnamese
        0x1E => 874,         // Thai
        _ => 1252,
    }
}

/// A language ID whose code page is the one a font character set (`chs`,
/// [MS-DOC] §2.9.102 `FFN`, the Windows `LOGFONT` charset) names; `None` for
/// the Western, symbol, OEM and default sets, which are Windows-1252 already.
pub(crate) fn lid_for_charset(chs: u8) -> Option<u16> {
    Some(match chs {
        0x80 => 0x0411, // SHIFTJIS_CHARSET, 932
        0x81 => 0x0412, // HANGUL_CHARSET, 949
        0x86 => 0x0804, // GB2312_CHARSET, 936
        0x88 => 0x0404, // CHINESEBIG5_CHARSET, 950
        0xA1 => 0x0408, // GREEK_CHARSET, 1253
        0xA2 => 0x041F, // TURKISH_CHARSET, 1254
        0xA3 => 0x042A, // VIETNAMESE_CHARSET, 1258
        0xB1 => 0x040D, // HEBREW_CHARSET, 1255
        0xB2 => 0x0401, // ARABIC_CHARSET, 1256
        0xBA => 0x0425, // BALTIC_CHARSET, 1257
        0xCC => 0x0419, // RUSSIAN_CHARSET, 1251
        0xDE => 0x041E, // THAI_CHARSET, 874
        0xEE => 0x0415, // EASTEUROPE_CHARSET, 1250
        _ => return None,
    })
}

/// Decode one compressed-text byte using the codepage its FIB `lid`
/// implies. Byte-for-byte identical to the pre-existing `cp1252_to_char`
/// for the default/English case (no behavior change for the vast
/// majority of real files), and correctly codepage-aware for the rest.
pub(crate) fn decode_byte(b: u8, lid: u16) -> char {
    let cp = codepage_for_lid(lid);
    if cp == 1252 {
        return super::piece_table::cp1252_to_char(b);
    }
    if let Some(encoding) = crate::core::codepage::encoding_for_codepage(cp) {
        let bytes = [b];
        let (decoded, _, _) = encoding.decode(&bytes);
        return decoded.chars().next().unwrap_or(b as char);
    }
    b as char
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_english_lid_maps_to_1252() {
        assert_eq!(codepage_for_lid(0x0409), 1252); // en-US
    }

    #[test]
    fn test_russian_lid_maps_to_1251() {
        assert_eq!(codepage_for_lid(0x0419), 1251); // ru-RU
    }

    #[test]
    fn test_sublanguage_region_does_not_change_the_codepage() {
        // Mexican Spanish (0x080A) and Spanish Spanish (0x0C0A) share a
        // primary language id and must map to the same codepage.
        assert_eq!(codepage_for_lid(0x080A), codepage_for_lid(0x0C0A));
    }

    #[test]
    fn test_greek_lid_decodes_a_byte_correctly() {
        // Windows-1253 0xE1 = alpha (α).
        assert_eq!(decode_byte(0xE1, 0x0408), 'α');
    }

    #[test]
    fn test_default_lid_matches_the_old_cp1252_behavior() {
        assert_eq!(decode_byte(0x93, 0x0409), '\u{201C}');
        assert_eq!(decode_byte(0x93, 0x0409), super::super::piece_table::cp1252_to_char(0x93));
    }

    /// A Macintosh-authored document stores its 8-bit text in Mac Roman:
    /// 0xD5 is the right single quote and 0x8E is e-acute, which Windows-1252
    /// reads as `Õ` and `Ž`.
    #[test]
    fn test_macintosh_text_decodes_as_mac_roman() {
        assert_eq!(decode_byte(0xD5, MAC_TEXT), '\u{2019}');
        assert_eq!(decode_byte(0x8E, MAC_TEXT), '\u{e9}');
        assert_eq!(decode_byte(b'a', MAC_TEXT), 'a');
        assert_eq!(codepage_for_lid(MAC_TEXT), 10000);
        // A byte that is the same in both is unchanged for a Windows document.
        assert_eq!(decode_byte(0xD5, 0x0409), 'Õ');
    }
}
