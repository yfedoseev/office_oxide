//! Shared Windows single-byte codepage support, used by both XLS's BIFF5
//! `CODEPAGE` record (issue #309) and DOC's FIB `lid`-derived codepage
//! (issue #310) to decode legacy 8-bit text.

/// Map a Windows codepage identifier to the `encoding_rs` encoding that
/// decodes it. Only codepages a real corpus file has been confirmed to
/// need (XLS) or that a FIB `lid` can plausibly select (DOC) are
/// included; an unmapped id returns `None` so the caller can fall back
/// to its own prior behavior rather than guess wrong.
pub(crate) fn encoding_for_codepage(cp: u16) -> Option<&'static encoding_rs::Encoding> {
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

/// Decode bytes with a Windows codepage, falling back to raw
/// byte-as-code-point (ISO-8859-1-shaped) promotion for a codepage this
/// module doesn't have a table for — no worse than the pre-existing
/// behavior for whatever handful of files use one.
pub(crate) fn decode_windows_codepage(bytes: &[u8], cp: u16) -> String {
    if let Some(encoding) = encoding_for_codepage(cp) {
        let (decoded, _, _) = encoding.decode(bytes);
        return decoded.into_owned();
    }
    bytes.iter().map(|&b| b as char).collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn known_codepage_decodes_via_encoding_rs() {
        let (encoded, _, _) = encoding_rs::WINDOWS_1251.encode("Привет");
        assert_eq!(decode_windows_codepage(&encoded, 1251), "Привет");
    }

    #[test]
    fn unmapped_codepage_falls_back_to_raw_byte_promotion() {
        assert_eq!(decode_windows_codepage(&[0x41, 0xE9], 65535), "A\u{00E9}");
    }
}
