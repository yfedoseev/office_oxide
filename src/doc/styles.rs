//! Style sheet (STSH) parsing for legacy binary `.doc` (MS-DOC §2.7.1).
//!
//! The style sheet maps a paragraph's style index (`istd`) to a built-in
//! style id (`sti`) and a name. Built-in heading styles carry `sti` 1–9
//! (Heading 1–9); user-defined heading styles are named `Heading N`. Both let
//! us derive a paragraph's real heading level instead of the line heuristic in
//! `convert_doc.rs`.
//!
//! Every parse step is bounds-checked: a malformed or truncated style sheet
//! yields an empty `Vec`, so callers degrade to "no style" (and the heuristic)
//! rather than panicking (AGENTS.md rule 6).

use super::{MAX_OUTLINE_LEVEL, fib::Fib};

/// One style definition, indexed by `istd`.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct StyleDef {
    /// Built-in style id (`StdfBase.sti`). `0x0FFE` means user-defined.
    pub sti: u16,
    /// Style name (from the style-name STTB).
    pub name: String,
}

/// Parse the document style sheet (STSH) from the Table stream.
///
/// Returns an empty vector when the style sheet is absent (`fcStshf == 0`),
/// out of bounds, or malformed.
pub fn parse_style_sheet(table_stream: &[u8], fib: &Fib) -> Vec<StyleDef> {
    if fib.fc_stshf == 0 || fib.lcb_stshf == 0 {
        return Vec::new();
    }
    let start = fib.fc_stshf as usize;
    // `saturating_add` so a malformed `fc_stshf`/`lcb_stshf` (e.g. near
    // `u32::MAX`) cannot overflow the pointer arithmetic on 32-bit targets;
    // on 64-bit the sum already fits, but the contract is "no overflow, ever"
    // (AGENTS.md rule 6).
    let end = start
        .saturating_add(fib.lcb_stshf as usize)
        .min(table_stream.len());
    if start >= table_stream.len() || end <= start {
        return Vec::new();
    }
    parse_stsh(&table_stream[start..end])
}

/// Resolve a paragraph's `istd` (optionally overridden by `sprmPStyle`) to a
/// heading level (1–9), or `None` when the style is not a heading.
pub fn heading_level_for_istd(styles: &[StyleDef], istd: u16) -> Option<u8> {
    let s = styles.get(istd as usize)?;
    // Built-in heading styles: `sti` 1..9 == Heading 1..9.
    if (1..=u16::from(MAX_OUTLINE_LEVEL)).contains(&s.sti) {
        return Some(s.sti as u8);
    }
    // User-defined heading styles are named "Heading N".
    if let Some(level) = heading_level_from_name(&s.name) {
        return Some(level);
    }
    None
}

fn heading_level_from_name(name: &str) -> Option<u8> {
    // Real Word style names are "Heading N" (capital H); match case-insensitively
    // so "heading 2" / "HEADING 2" also resolve.
    //
    // `xstzName` may carry aliases as "primary,alias,alias" (MS-DOC §2.9.258);
    // a heading level is always carried by the primary name, so look at the
    // segment before the first comma.
    let lowered: String = name.trim().to_ascii_lowercase();
    let primary = lowered.split(',').next().unwrap_or("");
    let rest = primary.trim().strip_prefix("heading ")?;
    let level: u8 = rest.trim().parse().ok()?;
    if (1..=MAX_OUTLINE_LEVEL).contains(&level) {
        Some(level)
    } else {
        None
    }
}

/// `data` is the `stshf` slice (the STSH). Layout (MS-DOC §2.9.271):
/// `LPStshi` = `cbStshi(u16)` + `Stshif(cbStshi bytes)`, immediately followed by
/// `rglpstd`, an array of `cstd` `LPStd` entries (`cbStd(u16)` + `STD`).
///
/// There is **no** style-name table between `Stshif` and `rglpstd`: `Stshif` is
/// exactly 18 bytes and carries no names (§2.9.274). Each style's name lives
/// inside its own `STD`, which is `stdf` + `xstzName` + `grLPUpxSw` (§2.9.258),
/// and `stdf` occupies `cbSTDBaseInFile` bytes.
fn parse_stsh(data: &[u8]) -> Vec<StyleDef> {
    if data.len() < 2 {
        return Vec::new();
    }
    let cb_stshi = u16::from_le_bytes([data[0], data[1]]) as usize;
    let mut pos = 2;
    if cb_stshi < 18 || pos + cb_stshi > data.len() {
        return Vec::new();
    }
    // Stshif (§2.9.274): `cstd` at 0 (u16), `cbSTDBaseInFile` at 2 (u16).
    let cstd = u16::from_le_bytes([data[pos], data[pos + 1]]) as usize;
    let cb_std_base = u16::from_le_bytes([data[pos + 2], data[pos + 3]]) as usize;
    pos += cb_stshi;

    // `cbSTDBaseInFile` MUST be 0x000A (`StdfBase` alone) or 0x0012 (`StdfBase`
    // + `StdfPost2000OrNone`) per §2.9.274. Anything else is malformed: slicing
    // `xstzName` at a bogus offset would yield plausible-but-wrong names, and
    // hence plausible-but-wrong heading levels, so reject and let the caller
    // degrade to the line heuristic rather than guess (AGENTS.md rule: fail
    // loudly, never fall back to a silent plausible-but-wrong result).
    if cb_std_base != 0x000A && cb_std_base != 0x0012 {
        return Vec::new();
    }

    // `rglpstd`: `cstd` `LPStd` entries (§2.9.135).
    let cap = cstd.min(4096);
    let mut styles = Vec::with_capacity(cap);
    for _ in 0..cap {
        if pos + 2 > data.len() {
            break;
        }
        let cb_std = u16::from_le_bytes([data[pos], data[pos + 1]]) as usize;
        pos += 2;
        if cb_std == 0 {
            // Empty style (fixed-index slots MAY be empty, §2.9.271).
            styles.push(StyleDef::default());
            continue;
        }
        if pos + cb_std > data.len() {
            break;
        }
        let std = &data[pos..pos + cb_std];
        // `StdfBase.sti` is the low 12 bits of the first u16 (§2.9.260).
        let sti = if std.len() >= 2 {
            u16::from_le_bytes([std[0], std[1]]) & 0x0FFF
        } else {
            0
        };
        // The name follows `stdf`, which occupies `cbSTDBaseInFile` bytes.
        let name = if std.len() > cb_std_base {
            parse_xstz(&std[cb_std_base..]).unwrap_or_default()
        } else {
            String::new()
        };
        styles.push(StyleDef { sti, name });
        pos += cb_std;
        // LPStd entries are stored on even-byte boundaries and `cbStd` does NOT
        // include that padding byte (§2.9.135).
        if !cb_std.is_multiple_of(2) {
            pos += 1;
        }
    }
    styles
}

/// Parse an `Xstz` (§2.9.354) — a style name: an `Xst` (a `cch: u16` followed
/// by `cch` UTF-16 code units) plus a 2-byte null terminator.
///
/// Per §2.9.258 the name may be `"primary,alias,alias"`; callers that want only
/// the primary name split on `','` themselves (`heading_level_from_name` does).
fn parse_xstz(data: &[u8]) -> Option<String> {
    if data.len() < 2 {
        return None;
    }
    let cch = u16::from_le_bytes([data[0], data[1]]) as usize;
    let end = 2usize.saturating_add(cch.saturating_mul(2));
    if end > data.len() {
        return None;
    }
    let units: Vec<u16> = (0..cch)
        .map(|i| u16::from_le_bytes([data[2 + 2 * i], data[3 + 2 * i]]))
        .collect();
    Some(String::from_utf16_lossy(&units))
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Build one `LPStd` for a style sheet: `cbStd(u16)` + `STD`.
    ///
    /// `STD` = `stdf` (`cbSTDBaseInFile` bytes, here `StdfBase` alone) +
    /// `xstzName`, per MS-DOC §2.9.258. `xstzName` is an `Xstz`: `cch(u16)` +
    /// `cch` UTF-16 code units + a 2-byte null terminator (§2.9.354).
    fn lpstd(sti: u16, name: &str) -> Vec<u8> {
        const CB_STD_BASE: usize = 0x000A;
        // `StdfBase` (10 bytes): `sti` is the low 12 bits of the first u16.
        let mut std = vec![0u8; CB_STD_BASE];
        std[0..2].copy_from_slice(&sti.to_le_bytes());
        // `xstzName`.
        let units: Vec<u16> = name.encode_utf16().collect();
        std.extend_from_slice(&(units.len() as u16).to_le_bytes()); // cch
        for u in &units {
            std.extend_from_slice(&u.to_le_bytes());
        }
        std.extend_from_slice(&0u16.to_le_bytes()); // chTerm

        let mut out = Vec::new();
        out.extend_from_slice(&(std.len() as u16).to_le_bytes()); // cbStd
        out.extend_from_slice(&std);
        out
    }

    /// An empty `LPStd` (`cbStd` = 0). Fixed-index slots 13–14 MUST be empty
    /// (§2.9.271) and 0–12 MAY be.
    fn empty_lpstd() -> Vec<u8> {
        vec![0, 0]
    }

    /// Build a spec-conformant STSH (§2.9.271 / §2.9.274): `LPStshi`
    /// (`cbStshi` = 18 + an 18-byte `Stshif`) followed directly by `rglpstd`.
    ///
    /// The 16 entries follow the fixed-index table: istd 0 = Normal (sti 0),
    /// istd 1–9 = Heading 1–9 (sti 1–9), istd 10–12 = sti 65/105/107, istd
    /// 13–14 empty, and istd 15 a user-defined style (sti 0x0FFE) whose name
    /// carries its level. Deliberately built from the spec, not from our
    /// parser, so it cannot merely re-encode our own bugs.
    fn synthetic_stsh() -> Vec<u8> {
        let mut d = Vec::new();
        // ── LPStshi: cbStshi, then Stshif (18 bytes total, no names) ──
        d.extend_from_slice(&18u16.to_le_bytes()); // cbStshi = 18
        d.extend_from_slice(&16u16.to_le_bytes()); // cstd
        d.extend_from_slice(&0x000Au16.to_le_bytes()); // cbSTDBaseInFile = 10
        d.extend_from_slice(&[0u8; 2]); // fStdStylenamesWritten + fReserved
        d.extend_from_slice(&108u16.to_le_bytes()); // stiMaxWhenSaved
        d.extend_from_slice(&0x000Fu16.to_le_bytes()); // istdMaxFixedWhenSaved
        d.extend_from_slice(&0u16.to_le_bytes()); // nVerBuiltInNamesWhenSaved
        d.extend_from_slice(&[0u8; 6]); // ftcAsci, ftcFE, ftcOther

        // ── rglpstd ──
        d.extend_from_slice(&lpstd(0, "Normal")); // istd 0
        for lvl in 1..=9u16 {
            d.extend_from_slice(&lpstd(lvl, &format!("Heading {lvl}"))); // 1–9
        }
        d.extend_from_slice(&lpstd(65, "Fixed Ten")); // istd 10
        d.extend_from_slice(&lpstd(105, "Fixed Eleven")); // istd 11
        d.extend_from_slice(&lpstd(107, "Fixed Twelve")); // istd 12
        d.extend_from_slice(&empty_lpstd()); // istd 13 — MUST be empty
        d.extend_from_slice(&empty_lpstd()); // istd 14 — MUST be empty
        d.extend_from_slice(&lpstd(0x0FFE, "heading 7")); // istd 15: user style
        d
    }

    #[test]
    fn parse_stsh_reads_sti_and_names() {
        let styles = parse_stsh(&synthetic_stsh());
        assert_eq!(styles.len(), 16, "all 16 LPStd entries must be decoded");
        // istd 0: Normal.
        assert_eq!(styles[0].sti, 0);
        assert_eq!(styles[0].name, "Normal");
        // istd 1..9: built-in Heading 1..9 — names come from each STD, not
        // from any shared table.
        assert_eq!(styles[1].sti, 1);
        assert_eq!(styles[1].name, "Heading 1");
        assert_eq!(styles[3].sti, 3);
        assert_eq!(styles[3].name, "Heading 3");
        assert_eq!(styles[9].sti, 9);
        assert_eq!(styles[9].name, "Heading 9");
        // istd 13/14 are empty.
        assert_eq!(styles[13], StyleDef::default());
        assert_eq!(styles[14], StyleDef::default());
        // istd 15: user-defined (sti 0x0FFE), name carries the level.
        assert_eq!(styles[15].sti, 0x0FFE);
        assert_eq!(styles[15].name, "heading 7");
    }

    /// Regression guard: an earlier revision decoded the style sheet as
    /// `cbStshi + Stshif + a style-name STTB + rglpstd`. That is wrong —
    /// `Stshif` is exactly 18 bytes and carries no names, and `rglpstd`
    /// follows immediately (§2.9.271); names live in each `STD.xstzName`.
    /// A style sheet laid out the old way must not be silently decoded into
    /// plausible-but-wrong style names: it must yield nothing, so the caller
    /// degrades to the line heuristic instead of inventing heading levels.
    #[test]
    fn parse_stsh_rejects_spurious_name_table_layout() {
        let mut d = Vec::new();
        d.extend_from_slice(&18u16.to_le_bytes()); // cbStshi
        d.extend_from_slice(&4u16.to_le_bytes()); // cstd
        d.extend_from_slice(&0x000Au16.to_le_bytes()); // cbSTDBaseInFile
        d.extend_from_slice(&[0u8; 14]); // rest of Stshif
        // The spurious style-name "STTB" that the old revision expected.
        d.push(1); // f2
        d.extend_from_slice(&4u16.to_le_bytes()); // cData
        d.extend_from_slice(&0u16.to_le_bytes()); // cbData
        for name in ["Normal", "", "", "Heading 3"] {
            let units: Vec<u16> = name.encode_utf16().collect();
            d.extend_from_slice(&(units.len() as u16).to_le_bytes());
            for u in units {
                d.extend_from_slice(&u.to_le_bytes());
            }
        }
        d.extend_from_slice(&0u16.to_le_bytes()); // trailing null cch
        for &sti in &[0u16, 0u16, 0u16, 3u16] {
            d.extend_from_slice(&10u16.to_le_bytes()); // cbStd
            d.extend_from_slice(&sti.to_le_bytes());
            d.extend_from_slice(&[0u8; 8]);
        }

        let styles = parse_stsh(&d);
        assert!(
            styles.iter().all(|s| s.name.is_empty()),
            "a spurious name table must not yield plausible style names; got {styles:?}"
        );
    }

    /// The name must be read from `STD.xstzName`, i.e. at `cbSTDBaseInFile`
    /// bytes into each `STD`. A fixture that shifts the name by even one byte
    /// must not silently produce a plausible-but-wrong name.
    #[test]
    fn parse_stsh_rejects_unexpected_cb_std_base_in_file() {
        let mut stsh = synthetic_stsh();
        // `cbSTDBaseInFile` is the u16 at offset 4 (cbStshi + cstd). Spec allows
        // only 0x000A / 0x0012; anything else is malformed.
        stsh[4..6].copy_from_slice(&0x0020u16.to_le_bytes());
        assert!(
            parse_stsh(&stsh).is_empty(),
            "a bogus cbSTDBaseInFile must be rejected, not guess a name offset"
        );
    }

    /// `LPStd` entries sit on even-byte boundaries and `cbStd` excludes the
    /// padding byte (§2.9.135), so an odd-sized `STD` must not desynchronise the
    /// following entries.
    #[test]
    fn parse_stsh_skips_odd_sized_lpstd_padding() {
        let mut stsh = synthetic_stsh();
        // Rebuild rglpstd: one odd-sized entry, then a normal one.
        let mut rglpstd = Vec::new();
        let mut odd = lpstd(2, "Heading 2");
        // Trim one byte off `STD` (and fix up cbStd) to make cbStd odd.
        let cb = u16::from_le_bytes([odd[0], odd[1]]) as usize;
        odd.truncate(cb + 1); // 2 (cbStd) + cb-1 bytes of STD
        let new_cb = (cb - 1) as u16;
        odd[0..2].copy_from_slice(&new_cb.to_le_bytes());
        rglpstd.extend_from_slice(&odd);
        rglpstd.push(0x00); // the padding byte `cbStd` does not count
        rglpstd.extend_from_slice(&lpstd(3, "Heading 3"));
        stsh.truncate(20); // keep cbStshi + Stshif
        stsh.extend_from_slice(&rglpstd);

        let styles = parse_stsh(&stsh);
        assert_eq!(styles.len(), 2, "both entries must be read");
        assert_eq!(styles[0].sti, 2, "odd-sized entry decodes normally");
        assert_eq!(
            styles[1].name, "Heading 3",
            "the padding byte must not desynchronise the next entry"
        );
    }

    #[test]
    fn heading_level_resolves_builtin_and_user() {
        let styles = parse_stsh(&synthetic_stsh());
        // Built-in: sti 1..9 resolve directly (istd 1..9).
        assert_eq!(heading_level_for_istd(&styles, 1), Some(1));
        assert_eq!(heading_level_for_istd(&styles, 3), Some(3));
        assert_eq!(heading_level_for_istd(&styles, 9), Some(9));
        // User-defined: sti 0x0FFE named "heading 7" resolves via name.
        assert_eq!(heading_level_for_istd(&styles, 15), Some(7));
        // sti 0 (Normal) is not a heading.
        assert_eq!(heading_level_for_istd(&styles, 0), None);
        // Empty fixed-index slots are not headings.
        assert_eq!(heading_level_for_istd(&styles, 13), None);
        // out-of-range istd.
        assert_eq!(heading_level_for_istd(&styles, 99), None);
    }

    /// `xstzName` may carry aliases as "primary,alias" (§2.9.258); the level is
    /// carried by the primary name, so a comma-suffixed name must still resolve.
    #[test]
    fn heading_level_uses_primary_name_before_aliases() {
        let styles = vec![
            StyleDef::default(),
            StyleDef {
                sti: 0x0FFE,
                name: "Heading 4,Title 4".into(),
            },
        ];
        assert_eq!(heading_level_for_istd(&styles, 1), Some(4));
    }

    #[test]
    fn truncated_style_sheet_is_empty() {
        assert!(parse_stsh(&[0u8; 4]).is_empty());
        assert!(parse_stsh(&[18, 0, 0, 0]).is_empty());
    }

    /// Build a `Fib` whose only non-zero fields are `fc_stshf` / `lcb_stshf`,
    /// pointing at a style sheet in the Table stream.
    fn fib_with_stsh(start: u32, len: u32) -> Fib {
        Fib {
            version: 0,
            use_table1: false,
            clx_offset: 0,
            clx_size: 0,
            text_len: 0,
            footnote_len: 0,
            header_len: 0,
            comment_len: 0,
            endnote_len: 0,
            textbox_len: 0,
            header_textbox_len: 0,
            fc_plcf_bte_papx: 0,
            lcb_plcf_bte_papx: 0,
            fc_plcf_lst: 0,
            lcb_plcf_lst: 0,
            fc_stshf: start,
            lcb_stshf: len,
        }
    }

    /// `parse_style_sheet` is the FIB-aware wrapper that `document.rs` calls: it
    /// must slice the STSH out of the Table stream at `fc_stshf` (bounds-checked)
    /// and hand the bytes to `parse_stsh`. This exercises the wrapper end-to-end
    /// with real STSH bytes placed at a non-zero offset — the path the POI corpus
    /// never reaches (every file reports `fc_stshf == 0`).
    #[test]
    fn parse_style_sheet_reads_stsh_at_fib_offset() {
        let stsh = synthetic_stsh();
        let start = 64usize;
        let mut table_stream = vec![0u8; start + stsh.len()];
        table_stream[start..start + stsh.len()].copy_from_slice(&stsh);
        let fib = fib_with_stsh(start as u32, stsh.len() as u32);
        let styles = parse_style_sheet(&table_stream, &fib);
        assert_eq!(styles.len(), 16);
        assert_eq!(styles[0].sti, 0);
        assert_eq!(styles[0].name, "Normal");
        assert_eq!(styles[3].sti, 3);
        assert_eq!(styles[3].name, "Heading 3");
        assert_eq!(styles[15].sti, 0x0FFE);
        assert_eq!(styles[15].name, "heading 7");
    }

    #[test]
    fn parse_style_sheet_absent_is_empty() {
        let table_stream = vec![0u8; 64];
        assert!(parse_style_sheet(&table_stream, &fib_with_stsh(0, 0)).is_empty());
    }

    #[test]
    fn parse_style_sheet_out_of_bounds_is_empty() {
        let stsh = synthetic_stsh();
        // Point `fc_stshf` past the end of the table stream.
        assert!(parse_style_sheet(&stsh, &fib_with_stsh(1000, stsh.len() as u32)).is_empty());
    }
}
