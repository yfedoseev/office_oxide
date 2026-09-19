//! `.xls` data validation: `DV` records (`0x01BE`), grouped under a `DVAL`
//! header (`0x01B2`). No record handling for either existed at all before
//! this module — data validation was a total scope gap, not merely an
//! extraction bug, matching the already-fixed conditional-formatting gap
//! at the same severity tier.
//!
//! Byte layout verified against Apache POI's `DVRecord`/`DVALRecord`
//! (a mature, independent reference implementation of the same [MS-XLS]
//! structures) rather than guessed from the field names alone.
//!
//! Scope: `DV`'s `formula1`/`formula2` fields are BIFF formula byte-code
//! (`Ptg` tokens), not text — decoding that requires a full formula
//! disassembler this crate doesn't have anywhere (the same already-
//! documented limitation as `CF`'s `rgce1`/`rgce2` for conditional
//! formatting). This module skips over them by their declared byte size
//! without attempting to decode them; `DataValidation::formula1`/
//! `formula2` are always `None` for XLS.
//!
//! `DVAL` itself carries no per-rule information (just a dialog-position/
//! object-id/count header for the following `DV` records) — its own
//! fields aren't needed by the IR, so it isn't parsed at all, only used
//! as a signal that `DV` records follow.

use super::condfmt::range_ref;
use super::sst::read_unicode_string;
use crate::ir::DataValidation;

/// Parse one `DV` record into a [`DataValidation`], or `None` if the
/// record is truncated or defines no ranges.
///
/// Layout ([MS-XLS] §2.4.44, cross-checked against POI's `DVRecord`):
/// `dwDVFlags: u32`, then four `XLUnicodeString`s (prompt title, error
/// title, prompt text, error text — each `cch: u16, flags: u8, chars`,
/// skipped over since none reach the IR), `cbFormula1: u16` + 2 reserved
/// bytes + `cbFormula1` formula bytes, the same shape again for
/// `formula2`, then a `Ref8U` list: `count: u16` followed by `count`
/// `(rwFirst, rwLast, colFirst, colLast)` `u16` quadruples.
pub fn parse_dv(data: &[u8]) -> Option<DataValidation> {
    if data.len() < 4 {
        return None;
    }
    let flags = u32::from_le_bytes([data[0], data[1], data[2], data[3]]);
    let mut pos = 4usize;

    for _ in 0..4 {
        let (_, new_pos) = read_unicode_string(data, pos).ok()?;
        pos = new_pos;
    }

    for _ in 0..2 {
        if pos + 4 > data.len() {
            return None;
        }
        let formula_size = u16::from_le_bytes([data[pos], data[pos + 1]]) as usize;
        pos += 4; // cb*: u16 + 2 reserved bytes ("not used", always fixed values)
        pos = pos.checked_add(formula_size)?;
    }

    if pos + 2 > data.len() {
        return None;
    }
    let count = u16::from_le_bytes([data[pos], data[pos + 1]]) as usize;
    pos += 2;
    let mut ranges = Vec::with_capacity(count.min(64));
    for _ in 0..count {
        if pos + 8 > data.len() {
            break;
        }
        let row_first = u16::from_le_bytes([data[pos], data[pos + 1]]);
        let row_last = u16::from_le_bytes([data[pos + 2], data[pos + 3]]);
        let col_first = u16::from_le_bytes([data[pos + 4], data[pos + 5]]);
        let col_last = u16::from_le_bytes([data[pos + 6], data[pos + 7]]);
        ranges.push(range_ref(row_first, row_last, col_first, col_last));
        pos += 8;
    }
    if ranges.is_empty() {
        return None;
    }

    let data_type = (flags & 0x0000_000F) as u8;
    let allow_blank = flags & 0x0000_0100 != 0;
    let validation_type = match data_type {
        0x00 => "none",
        0x01 => "whole",
        0x02 => "decimal",
        0x03 => "list",
        0x04 => "date",
        0x05 => "time",
        0x06 => "textLength",
        0x07 => "custom",
        _ => "none",
    }
    .to_string();

    let operator = if matches!(validation_type.as_str(), "list" | "custom" | "none") {
        None
    } else {
        let operator_bits = ((flags & 0x0070_0000) >> 20) as u8;
        Some(
            match operator_bits {
                0 => "between",
                1 => "notBetween",
                2 => "equal",
                3 => "notEqual",
                4 => "greaterThan",
                5 => "lessThan",
                6 => "greaterThanOrEqual",
                7 => "lessThanOrEqual",
                _ => "between",
            }
            .to_string(),
        )
    };

    Some(DataValidation {
        range: ranges.join(" "),
        validation_type,
        operator,
        formula1: None,
        formula2: None,
        allow_blank,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn empty_unicode_string() -> Vec<u8> {
        vec![0, 0, 0] // cch=0, flags=0
    }

    fn dv_record(
        flags: u32,
        formula1: &[u8],
        formula2: &[u8],
        ranges: &[(u16, u16, u16, u16)],
    ) -> Vec<u8> {
        let mut d = Vec::new();
        d.extend_from_slice(&flags.to_le_bytes());
        for _ in 0..4 {
            d.extend_from_slice(&empty_unicode_string());
        }
        d.extend_from_slice(&(formula1.len() as u16).to_le_bytes());
        d.extend_from_slice(&0x3FE0u16.to_le_bytes()); // not_used_1
        d.extend_from_slice(formula1);
        d.extend_from_slice(&(formula2.len() as u16).to_le_bytes());
        d.extend_from_slice(&0u16.to_le_bytes()); // not_used_2
        d.extend_from_slice(formula2);
        d.extend_from_slice(&(ranges.len() as u16).to_le_bytes());
        for &(rf, rl, cf, cl) in ranges {
            d.extend_from_slice(&rf.to_le_bytes());
            d.extend_from_slice(&rl.to_le_bytes());
            d.extend_from_slice(&cf.to_le_bytes());
            d.extend_from_slice(&cl.to_le_bytes());
        }
        d
    }

    #[test]
    fn test_whole_number_between_rule() {
        // data_type=0x01 (whole), operator bits=0 (between), allow_blank set.
        let flags = 0x01 | 0x0000_0100;
        let data = dv_record(flags, &[0xAA; 4], &[0xBB; 4], &[(0, 0, 0, 0)]);
        let dv = parse_dv(&data).unwrap();
        assert_eq!(dv.range, "A1");
        assert_eq!(dv.validation_type, "whole");
        assert_eq!(dv.operator.as_deref(), Some("between"));
        assert!(dv.allow_blank);
        assert!(dv.formula1.is_none());
        assert!(dv.formula2.is_none());
    }

    #[test]
    fn test_list_type_has_no_operator() {
        let flags = 0x03; // list, allow_blank unset
        let data = dv_record(flags, &[], &[], &[(1, 1, 1, 1)]);
        let dv = parse_dv(&data).unwrap();
        assert_eq!(dv.validation_type, "list");
        assert!(dv.operator.is_none());
        assert!(!dv.allow_blank);
        assert_eq!(dv.range, "B2");
    }

    #[test]
    fn test_greater_than_operator_decoded() {
        // data_type=0x02 (decimal), operator bits=4 (greaterThan).
        let flags = 0x02 | (4u32 << 20);
        let data = dv_record(flags, &[0xCC; 2], &[], &[(0, 4, 0, 0)]);
        let dv = parse_dv(&data).unwrap();
        assert_eq!(dv.validation_type, "decimal");
        assert_eq!(dv.operator.as_deref(), Some("greaterThan"));
        assert_eq!(dv.range, "A1:A5");
    }

    #[test]
    fn test_multi_range_sqref_is_space_joined() {
        let data = dv_record(0x04, &[], &[], &[(0, 0, 0, 0), (2, 4, 1, 2)]);
        let dv = parse_dv(&data).unwrap();
        assert_eq!(dv.range, "A1 B3:C5");
    }

    #[test]
    fn test_truncated_record_is_none_not_a_panic() {
        assert!(parse_dv(&[0u8; 3]).is_none());
        assert!(parse_dv(&[0u8; 10]).is_none());
    }

    #[test]
    fn test_no_ranges_is_none() {
        let data = dv_record(0x01, &[], &[], &[]);
        assert!(parse_dv(&data).is_none());
    }
}
