//! Minimal [MS-OLEPS] (OLE Property Set) decoder for the
//! `\x05SummaryInformation` stream every legacy `.doc`/`.xls`/`.ppt` file
//! carries by default.
//!
//! Only the handful of well-known `PIDSI_*` properties `Metadata` has
//! fields for are decoded — title, subject, author, keywords, comments
//! and the two `FILETIME` dates. This is deliberately not a general
//! OLEPS/VARIANT decoder: no vectors, arrays, non-simple (storage-backed)
//! properties, or code-page-aware string decoding (`VT_LPSTR`'s bytes are
//! read as Latin-1/CP1252, the same simplification this crate's other
//! legacy-format readers already make — see #309/#310 for the tracked,
//! separate gap in real code-page support).
//!
//! Byte layouts verified against the published [MS-OLEPS] spec pages
//! (PropertySetStream, PropertySet, PropertyIdentifierAndOffset,
//! TypedPropertyValue, CodePageString, UnicodeString) rather than
//! recalled from memory — a wrong offset here produces silently wrong
//! metadata, not a caught error, and this format has a real CVE history
//! (integer overflow in offset arithmetic) in peer implementations.

/// Well-known property IDs in the `SummaryInformation` FMTID
/// ({F29F85E0-4FF9-1068-AB91-08002B27B3D9}), [MS-OLEPS] §2.16.
const PIDSI_TITLE: u32 = 2;
const PIDSI_SUBJECT: u32 = 3;
const PIDSI_AUTHOR: u32 = 4;
const PIDSI_KEYWORDS: u32 = 5;
const PIDSI_COMMENTS: u32 = 6;
const PIDSI_CREATE_DTM: u32 = 12;
const PIDSI_LASTSAVE_DTM: u32 = 13;

const VT_LPSTR: u16 = 0x001E;
const VT_LPWSTR: u16 = 0x001F;
const VT_FILETIME: u16 = 0x0040;

/// The subset of `SummaryInformation` this crate's `Metadata` has fields
/// for. `None` for a field means the property was absent, not that
/// parsing failed — a stream that declares only 3 of the 7 properties is
/// normal, not malformed.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct SummaryProperties {
    /// `PIDSI_TITLE` (2).
    pub title: Option<String>,
    /// `PIDSI_SUBJECT` (3).
    pub subject: Option<String>,
    /// `PIDSI_AUTHOR` (4).
    pub author: Option<String>,
    /// `PIDSI_KEYWORDS` (5).
    pub keywords: Option<String>,
    /// `PIDSI_COMMENTS` (6).
    pub comments: Option<String>,
    /// `PIDSI_CREATE_DTM` (12), ISO-8601 (`YYYY-MM-DDTHH:MM:SSZ`),
    /// converted from the raw `FILETIME` (100-ns ticks since 1601-01-01
    /// UTC).
    pub created: Option<String>,
    /// `PIDSI_LASTSAVE_DTM` (13). See `created`.
    pub modified: Option<String>,
}

impl SummaryProperties {
    fn is_empty(&self) -> bool {
        self.title.is_none()
            && self.subject.is_none()
            && self.author.is_none()
            && self.keywords.is_none()
            && self.comments.is_none()
            && self.created.is_none()
            && self.modified.is_none()
    }
}

/// Parse a `\x05SummaryInformation` stream's raw bytes.
///
/// Returns `None` for a stream that is missing, too short, carries the
/// wrong byte-order marker, or ends up with none of the 7 tracked
/// properties resolved — every failure mode here means "no metadata
/// recovered", never a panic or an out-of-bounds read: every offset is
/// checked against the buffer length before use.
pub fn parse_summary_information(data: &[u8]) -> Option<SummaryProperties> {
    // PropertySetStream header ([MS-OLEPS] §2.21):
    // ByteOrder(2) Version(2) SystemIdentifier(4) CLSID(16) NumPropertySets(4)
    // then, per property set: FMTID(16) Offset(4).
    if data.len() < 28 {
        return None;
    }
    let byte_order = u16::from_le_bytes([data[0], data[1]]);
    if byte_order != 0xFFFE {
        return None;
    }
    let num_property_sets = u32::from_le_bytes([data[24], data[25], data[26], data[27]]);
    if num_property_sets == 0 {
        return None;
    }
    // Only the first property set matters here — SummaryInformation is
    // always a single-property-set stream; DocumentSummaryInformation's
    // (rarely present) second set carries different, unrelated properties.
    let fmtid_offset = 28usize;
    if data.len() < fmtid_offset + 20 {
        return None;
    }
    let offset = u32::from_le_bytes([
        data[fmtid_offset + 16],
        data[fmtid_offset + 17],
        data[fmtid_offset + 18],
        data[fmtid_offset + 19],
    ]) as usize;

    let props = parse_property_set(data, offset)?;
    if props.is_empty() { None } else { Some(props) }
}

/// Parse one PropertySet packet ([MS-OLEPS] §2.17) at `base` within `data`.
fn parse_property_set(data: &[u8], base: usize) -> Option<SummaryProperties> {
    // Size(4) NumProperties(4) then NumProperties * PropertyIdentifierAndOffset(8).
    let header_end = base.checked_add(8)?;
    if data.len() < header_end {
        return None;
    }
    let num_properties =
        u32::from_le_bytes([data[base + 4], data[base + 5], data[base + 6], data[base + 7]])
            as usize;
    // A hostile stream can claim billions of properties; bound the work
    // to what a real SummaryInformation set could ever hold.
    let num_properties = num_properties.min(64);

    let mut out = SummaryProperties::default();
    for i in 0..num_properties {
        let entry = header_end.checked_add(i.checked_mul(8)?)?;
        if data.len() < entry + 8 {
            break;
        }
        let id = u32::from_le_bytes([data[entry], data[entry + 1], data[entry + 2], data[entry + 3]]);
        let rel_offset = u32::from_le_bytes([
            data[entry + 4],
            data[entry + 5],
            data[entry + 6],
            data[entry + 7],
        ]) as usize;
        // Offset is relative to the start of *this* PropertySet packet
        // (i.e. `base`, where its own Size field begins), not the stream.
        let Some(value_at) = base.checked_add(rel_offset) else {
            continue;
        };
        let field = match id {
            PIDSI_TITLE => &mut out.title,
            PIDSI_SUBJECT => &mut out.subject,
            PIDSI_AUTHOR => &mut out.author,
            PIDSI_KEYWORDS => &mut out.keywords,
            PIDSI_COMMENTS => &mut out.comments,
            PIDSI_CREATE_DTM => {
                out.created = read_filetime(data, value_at);
                continue;
            },
            PIDSI_LASTSAVE_DTM => {
                out.modified = read_filetime(data, value_at);
                continue;
            },
            _ => continue,
        };
        if let Some(s) = read_string(data, value_at) {
            if !s.is_empty() {
                *field = Some(s);
            }
        }
    }
    Some(out)
}

/// Read a `TypedPropertyValue` whose `Type` is `VT_LPSTR` or `VT_LPWSTR`
/// at `pos`. Any other `Type` (or an out-of-bounds `pos`) yields `None`
/// rather than misreading unrelated bytes as text.
fn read_string(data: &[u8], pos: usize) -> Option<String> {
    if data.len() < pos + 4 {
        return None;
    }
    let ty = u16::from_le_bytes([data[pos], data[pos + 1]]);
    let value_start = pos + 4; // past Type(2) + Padding(2)
    match ty {
        VT_LPSTR => {
            // CodePageString: cch(4, includes the NUL, excludes padding)
            // then that many single bytes.
            if data.len() < value_start + 4 {
                return None;
            }
            let cch = u32::from_le_bytes([
                data[value_start],
                data[value_start + 1],
                data[value_start + 2],
                data[value_start + 3],
            ]) as usize;
            let cch = cch.min(1 << 20); // 1M-char cap against a hostile length
            let start = value_start + 4;
            let end = start.checked_add(cch)?;
            if data.len() < end {
                return None;
            }
            let bytes = &data[start..end];
            // Latin-1/CP1252 fallback (see module docs); strip the NUL
            // terminator `cch` includes.
            let s: String = bytes.iter().map(|&b| b as char).collect();
            Some(s.trim_end_matches('\0').to_string())
        },
        VT_LPWSTR => {
            // UnicodeString: Length(4, UTF-16 code units, includes the
            // NUL, excludes padding) then that many UTF-16LE units.
            if data.len() < value_start + 4 {
                return None;
            }
            let len = u32::from_le_bytes([
                data[value_start],
                data[value_start + 1],
                data[value_start + 2],
                data[value_start + 3],
            ]) as usize;
            let len = len.min(1 << 20);
            let start = value_start + 4;
            let byte_len = len.checked_mul(2)?;
            let end = start.checked_add(byte_len)?;
            if data.len() < end {
                return None;
            }
            let units: Vec<u16> = data[start..end]
                .chunks_exact(2)
                .map(|c| u16::from_le_bytes([c[0], c[1]]))
                .collect();
            let s = String::from_utf16_lossy(&units);
            Some(s.trim_end_matches('\0').to_string())
        },
        _ => None,
    }
}

/// Read a `TypedPropertyValue` whose `Type` is `VT_FILETIME` at `pos`,
/// converted to an ISO-8601 UTC string. `None` for any other `Type`, an
/// out-of-bounds `pos`, or a zero `FILETIME` (Office writes an all-zero
/// value for "no date", not the 1601 epoch).
fn read_filetime(data: &[u8], pos: usize) -> Option<String> {
    if data.len() < pos + 12 {
        return None;
    }
    let ty = u16::from_le_bytes([data[pos], data[pos + 1]]);
    if ty != VT_FILETIME {
        return None;
    }
    let value_start = pos + 4;
    let low = u32::from_le_bytes([
        data[value_start],
        data[value_start + 1],
        data[value_start + 2],
        data[value_start + 3],
    ]);
    let high = u32::from_le_bytes([
        data[value_start + 4],
        data[value_start + 5],
        data[value_start + 6],
        data[value_start + 7],
    ]);
    let ticks = ((high as u64) << 32) | (low as u64);
    filetime_to_iso8601(ticks)
}

/// Convert a Windows `FILETIME` (100-ns ticks since 1601-01-01 00:00:00
/// UTC) to an ISO-8601 string, without pulling in a date/time crate this
/// project otherwise has no need for.
fn filetime_to_iso8601(ticks: u64) -> Option<String> {
    if ticks == 0 {
        return None;
    }
    const TICKS_PER_SEC: u64 = 10_000_000;
    // Seconds between the FILETIME epoch (1601-01-01) and the Unix epoch
    // (1970-01-01): 11,644,473,600, per [MS-DTYP] §2.3.3.
    const EPOCH_DELTA_SECS: u64 = 11_644_473_600;
    let total_secs = ticks / TICKS_PER_SEC;
    let unix_secs = total_secs.checked_sub(EPOCH_DELTA_SECS)?;

    let days = unix_secs / 86_400;
    let secs_of_day = unix_secs % 86_400;
    let (hour, minute, second) = (secs_of_day / 3600, (secs_of_day / 60) % 60, secs_of_day % 60);
    let (year, month, day) = civil_from_days(days as i64);
    Some(format!(
        "{year:04}-{month:02}-{day:02}T{hour:02}:{minute:02}:{second:02}Z"
    ))
}

/// Days-since-1970-01-01 to (year, month, day), Howard Hinnant's
/// `civil_from_days` algorithm (public domain) — proleptic Gregorian,
/// correct for the full range any real document date falls in.
fn civil_from_days(z: i64) -> (i64, u32, u32) {
    let z = z + 719_468;
    let era = if z >= 0 { z } else { z - 146_096 } / 146_097;
    let doe = (z - era * 146_097) as u64; // [0, 146096]
    let yoe = (doe - doe / 1460 + doe / 36524 - doe / 146_096) / 365; // [0, 399]
    let y = yoe as i64 + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100); // [0, 365]
    let mp = (5 * doy + 2) / 153; // [0, 11]
    let d = (doy - (153 * mp + 2) / 5 + 1) as u32; // [1, 31]
    let m = if mp < 10 { mp + 3 } else { mp - 9 } as u32; // [1, 12]
    let y = if m <= 2 { y + 1 } else { y };
    (y, m, d)
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Build a minimal, valid `SummaryInformation` stream with exactly
    /// the properties in `props`, each `(id, TypedPropertyValue bytes)`.
    fn build_stream(props: &[(u32, Vec<u8>)]) -> Vec<u8> {
        let mut out = Vec::new();
        out.extend_from_slice(&0xFFFEu16.to_le_bytes()); // ByteOrder
        out.extend_from_slice(&0u16.to_le_bytes()); // Version
        out.extend_from_slice(&0u32.to_le_bytes()); // SystemIdentifier
        out.extend_from_slice(&[0u8; 16]); // CLSID
        out.extend_from_slice(&1u32.to_le_bytes()); // NumPropertySets

        // FMTID (arbitrary — not checked by the decoder) + Offset.
        let fmtid_and_offset_pos = out.len();
        out.extend_from_slice(&[0u8; 16]); // FMTID
        out.extend_from_slice(&0u32.to_le_bytes()); // Offset placeholder
        let property_set_offset = out.len() as u32;
        out[fmtid_and_offset_pos + 16..fmtid_and_offset_pos + 20]
            .copy_from_slice(&property_set_offset.to_le_bytes());

        // PropertySet: Size, NumProperties, then the
        // PropertyIdentifierAndOffset array, then the values.
        let size_pos = out.len();
        out.extend_from_slice(&0u32.to_le_bytes()); // Size placeholder
        out.extend_from_slice(&(props.len() as u32).to_le_bytes());

        let header_len = 8 + props.len() * 8;
        let mut value_offset = header_len;
        for (id, value) in props {
            out.extend_from_slice(&id.to_le_bytes());
            out.extend_from_slice(&(value_offset as u32).to_le_bytes());
            value_offset += value.len();
        }
        for (_, value) in props {
            out.extend_from_slice(value);
        }
        let total_size = out.len() - size_pos;
        out[size_pos..size_pos + 4].copy_from_slice(&(total_size as u32).to_le_bytes());
        out
    }

    fn lpstr(s: &str) -> Vec<u8> {
        let mut v = Vec::new();
        v.extend_from_slice(&VT_LPSTR.to_le_bytes());
        v.extend_from_slice(&0u16.to_le_bytes()); // padding
        let cch = (s.len() + 1) as u32; // + NUL
        v.extend_from_slice(&cch.to_le_bytes());
        v.extend_from_slice(s.as_bytes());
        v.push(0); // NUL terminator
        while v.len() % 4 != 0 {
            v.push(0);
        }
        v
    }

    fn lpwstr(s: &str) -> Vec<u8> {
        let mut v = Vec::new();
        v.extend_from_slice(&VT_LPWSTR.to_le_bytes());
        v.extend_from_slice(&0u16.to_le_bytes());
        let units: Vec<u16> = s.encode_utf16().collect();
        let len = (units.len() + 1) as u32; // + NUL
        v.extend_from_slice(&len.to_le_bytes());
        for u in &units {
            v.extend_from_slice(&u.to_le_bytes());
        }
        v.extend_from_slice(&0u16.to_le_bytes()); // NUL terminator
        while v.len() % 4 != 0 {
            v.push(0);
        }
        v
    }

    fn filetime(ticks: u64) -> Vec<u8> {
        let mut v = Vec::new();
        v.extend_from_slice(&VT_FILETIME.to_le_bytes());
        v.extend_from_slice(&0u16.to_le_bytes());
        v.extend_from_slice(&(ticks as u32).to_le_bytes());
        v.extend_from_slice(&((ticks >> 32) as u32).to_le_bytes());
        v
    }

    #[test]
    fn decodes_lpstr_title_and_author() {
        let stream = build_stream(&[
            (PIDSI_TITLE, lpstr("Hello World")),
            (PIDSI_AUTHOR, lpstr("Jane Doe")),
        ]);
        let props = parse_summary_information(&stream).expect("must parse");
        assert_eq!(props.title.as_deref(), Some("Hello World"));
        assert_eq!(props.author.as_deref(), Some("Jane Doe"));
        assert!(props.subject.is_none());
    }

    #[test]
    fn decodes_lpwstr_unicode_title() {
        let stream = build_stream(&[(PIDSI_TITLE, lpwstr("Café Résumé"))]);
        let props = parse_summary_information(&stream).expect("must parse");
        assert_eq!(props.title.as_deref(), Some("Café Résumé"));
    }

    #[test]
    fn decodes_filetime_dates() {
        // 2006-09-05T00:00:00Z, a value already used as a reference date
        // elsewhere in this crate's own XLS date tests.
        // Unix seconds for 2006-09-05T00:00:00Z = 1157414400.
        let ticks = (1_157_414_400u64 + 11_644_473_600) * 10_000_000;
        let stream = build_stream(&[(PIDSI_CREATE_DTM, filetime(ticks))]);
        let props = parse_summary_information(&stream).expect("must parse");
        assert_eq!(props.created.as_deref(), Some("2006-09-05T00:00:00Z"));
    }

    #[test]
    fn zero_filetime_is_none_not_the_1601_epoch() {
        let stream = build_stream(&[(PIDSI_CREATE_DTM, filetime(0))]);
        // Every other property absent too, and a zero FILETIME resolves
        // to None, so this whole stream carries nothing usable.
        assert!(parse_summary_information(&stream).is_none());
    }

    #[test]
    fn wrong_byte_order_is_rejected() {
        let mut stream = build_stream(&[(PIDSI_TITLE, lpstr("x"))]);
        stream[0] = 0x00;
        stream[1] = 0x00;
        assert!(parse_summary_information(&stream).is_none());
    }

    #[test]
    fn truncated_stream_does_not_panic() {
        let stream = build_stream(&[(PIDSI_TITLE, lpstr("Hello World"))]);
        for cut in 0..stream.len() {
            let _ = parse_summary_information(&stream[..cut]);
        }
    }

    #[test]
    fn empty_stream_is_none() {
        assert!(parse_summary_information(&[]).is_none());
    }

    #[test]
    fn unknown_property_ids_are_ignored_not_misread() {
        let stream = build_stream(&[
            (999, lpstr("unrelated property")),
            (PIDSI_SUBJECT, lpstr("Real Subject")),
        ]);
        let props = parse_summary_information(&stream).expect("must parse");
        assert_eq!(props.subject.as_deref(), Some("Real Subject"));
    }

    #[test]
    fn civil_from_days_matches_known_dates() {
        assert_eq!(civil_from_days(0), (1970, 1, 1));
        assert_eq!(civil_from_days(13_396), (2006, 9, 5));
        assert_eq!(civil_from_days(-1), (1969, 12, 31));
    }
}
