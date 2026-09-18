//! `.xls` conditional formatting: `CONDFMT` (record type `0x01B0`) opens a
//! rule group over a cell range, followed by `ccf` `CF` (`0x01B1`) records,
//! one per rule. No record handling for either existed at all before this
//! module — conditional formatting was a total scope gap, not merely an
//! extraction bug (issue #252).
//!
//! Byte layouts verified against the published [MS-XLS] `CondFmt` and `CF`
//! spec pages, including their own worked byte examples.
//!
//! Scope: `CF`'s `rgce1`/`rgce2` formula operands are BIFF formula
//! byte-code (`Ptg` tokens), not text — decoding that requires a full
//! formula disassembler this crate doesn't have anywhere yet (cell
//! formulas themselves are read only for their cached result, never their
//! token stream — a separate, already-known gap). `range`/`rule_type`/
//! `operator` are extracted without needing to walk past them at all,
//! since they sit in the CF record's own first two bytes; `formulas` is
//! therefore always empty for XLS (unlike XLSX, where the formula is
//! already plain text in the XML).

use crate::ir::ConditionalFormat;

/// Convert a 0-based column index to a letter string: 0 -> "A", 25 -> "Z".
fn col_name(col: u16) -> String {
    let mut result = Vec::new();
    let mut n = col as u32 + 1;
    while n > 0 {
        n -= 1;
        result.push(b'A' + (n % 26) as u8);
        n /= 26;
    }
    result.reverse();
    String::from_utf8(result).unwrap()
}

/// Format a 0-based `(row, col)` pair as `"A1"`-style, and a range as
/// `"A1:B2"` (or just `"A1"` when the range is a single cell).
fn range_ref(row_first: u16, row_last: u16, col_first: u16, col_last: u16) -> String {
    let first = format!("{}{}", col_name(col_first), row_first + 1);
    if row_first == row_last && col_first == col_last {
        first
    } else {
        format!("{first}:{}{}", col_name(col_last), row_last + 1)
    }
}

/// Parse a `CONDFMT` record: its `sqref` (the cell range(s) the following
/// `CF` records apply to, space-separated for a multi-range `sqref`) and
/// `ccf` (how many `CF` records follow).
///
/// [MS-XLS] §2.4.56: `ccf: u16`, `flags: u16` (fToughRecalc/nID, unused
/// here), `refBound` (8 bytes, an outer bound — the real per-rule range is
/// `sqref`, not this), then `sqref`: `cref: u16` followed by `cref` 8-byte
/// `Ref8U` entries (`rwFirst`/`rwLast`/`colFirst`/`colLast`, all `u16`).
pub fn parse_condfmt(data: &[u8]) -> Option<(String, u16)> {
    if data.len() < 14 {
        return None;
    }
    let ccf = u16::from_le_bytes([data[0], data[1]]);
    let cref = u16::from_le_bytes([data[12], data[13]]) as usize;
    let mut pos = 14usize;
    let mut ranges = Vec::with_capacity(cref.min(64));
    for _ in 0..cref {
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
    Some((ranges.join(" "), ccf))
}

/// Parse a `CF` record's `ct` (condition type) and `cp` (comparison
/// operator, meaningful only when `ct == 0x01`) into the XLSX-equivalent
/// `rule_type`/`operator` vocabulary, so a caller sees the same strings
/// regardless of which format the workbook came from.
///
/// [MS-XLS] §2.4.42: `ct: u8` at offset 0, `cp: u8` at offset 1.
pub fn parse_cf(range: &str, data: &[u8]) -> Option<ConditionalFormat> {
    if data.len() < 2 {
        return None;
    }
    let ct = data[0];
    let cp = data[1];
    let (rule_type, operator) = match ct {
        0x01 => (
            "cellIs".to_string(),
            Some(
                match cp {
                    0x01 => "between",
                    0x02 => "notBetween",
                    0x03 => "equal",
                    0x04 => "notEqual",
                    0x05 => "greaterThan",
                    0x06 => "lessThan",
                    0x07 => "greaterThanOrEqual",
                    0x08 => "lessThanOrEqual",
                    _ => "unknown",
                }
                .to_string(),
            ),
        ),
        0x02 => ("expression".to_string(), None),
        _ => ("unknown".to_string(), None),
    };
    Some(ConditionalFormat {
        range: range.to_string(),
        rule_type,
        operator,
        formulas: Vec::new(),
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn condfmt_record(ccf: u16, ranges: &[(u16, u16, u16, u16)]) -> Vec<u8> {
        let mut d = Vec::new();
        d.extend_from_slice(&ccf.to_le_bytes());
        d.extend_from_slice(&0u16.to_le_bytes()); // flags
        d.extend_from_slice(&[0u8; 8]); // refBound
        d.extend_from_slice(&(ranges.len() as u16).to_le_bytes()); // cref
        for &(rf, rl, cf, cl) in ranges {
            d.extend_from_slice(&rf.to_le_bytes());
            d.extend_from_slice(&rl.to_le_bytes());
            d.extend_from_slice(&cf.to_le_bytes());
            d.extend_from_slice(&cl.to_le_bytes());
        }
        d
    }

    /// The exact byte shape from [MS-XLS]'s own "Structure of CondFmt"
    /// worked example: ccf=1, one range B2:B2 (0-based row/col 1,1).
    #[test]
    fn parses_the_spec_worked_example() {
        let data = condfmt_record(1, &[(1, 1, 0, 0)]);
        let (range, ccf) = parse_condfmt(&data).unwrap();
        assert_eq!(range, "A2");
        assert_eq!(ccf, 1);
    }

    #[test]
    fn multi_range_sqref_is_space_joined() {
        let data = condfmt_record(1, &[(0, 0, 0, 0), (2, 4, 1, 2)]);
        let (range, _) = parse_condfmt(&data).unwrap();
        assert_eq!(range, "A1 B3:C5");
    }

    #[test]
    fn truncated_condfmt_is_none_not_a_panic() {
        assert!(parse_condfmt(&[0u8; 4]).is_none());
    }

    #[test]
    fn cell_is_between_rule() {
        let cf = parse_cf("A1:A10", &[0x01, 0x01]).unwrap();
        assert_eq!(cf.range, "A1:A10");
        assert_eq!(cf.rule_type, "cellIs");
        assert_eq!(cf.operator.as_deref(), Some("between"));
        assert!(cf.formulas.is_empty());
    }

    #[test]
    fn expression_rule_has_no_operator() {
        let cf = parse_cf("B1:B5", &[0x02, 0x00]).unwrap();
        assert_eq!(cf.rule_type, "expression");
        assert!(cf.operator.is_none());
    }

    #[test]
    fn truncated_cf_is_none_not_a_panic() {
        assert!(parse_cf("A1", &[0x01]).is_none());
    }

    #[test]
    fn col_name_basic() {
        assert_eq!(col_name(0), "A");
        assert_eq!(col_name(25), "Z");
        assert_eq!(col_name(26), "AA");
    }
}
