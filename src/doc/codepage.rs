//! DOC FIB `lid` (language ID) to codepage mapping, applied to
//! compressed (8-bit) text runs. Mirrors the XLS `CODEPAGE` fix, and
//! is source-confirmed rather than corpus-demonstrated:
//! `Fib` had no field consulting the document's language at all before
//! this, so any compressed non-Western-European text was corrupted by
//! construction, whether or not a real example was on hand to prove it.

/// Map a FIB `lid` (Windows LCID) to the codepage compressed text in
/// that language is stored in. Only the primary language id (the low 10
/// bits) determines the codepage — the sublanguage/region bits (e.g.
/// distinguishing Mexican from European Spanish) don't change it.
/// Defaults to 1252 (Western European) for English and any
/// unrecognized/neutral `lid`, matching Word's own default and every
/// codepage-unaware caller's existing behavior.
pub(crate) fn codepage_for_lid(lid: u16) -> u16 {
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
}
