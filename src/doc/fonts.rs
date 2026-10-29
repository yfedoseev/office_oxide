//! The runs of a `.doc` whose font is a symbol font.
//!
//! Symbol, Wingdings, Wingdings 2 and 3 and Webdings store their text as
//! private-use characters (`U+F000` plus the font's byte). A character outside
//! U+0020..U+007F takes its font from `sprmCRgFtc2` ([MS-DOC] §2.6.1), an
//! index into the font table (`SttbfFfn`, §2.9.286). [`parse_symbol_fonts`]
//! reads that table, and [`SymbolRuns`] says which characters of the text are
//! in a symbol font, so the decoder can replace them with the font's own
//! characters ([`crate::core::symbol_font`]).

use super::sprm::ChpProps;
use super::styles::{StyleDef, style_ftc_other};
use crate::core::symbol_font::{SymbolFont, char_for, symbol_font};

/// `cData` MUST NOT exceed this ([MS-DOC] §2.9.286).
const MAX_FONTS: usize = 0x7FF0;
/// The fixed part of an `FFN`: `ffid` (1), `wWeight` (2), `chs` (1),
/// `ixchSzAlt` (1), `panose` (10) and `fs` (24); the UTF-16 name follows.
const FFN_FIXED_LEN: usize = 39;

/// For each entry of the `SttbfFfn` at `fc`/`lcb` in the table stream, the
/// symbol font it names, if any. Bounded by the structure and the stream: a
/// damaged table gives a short list, never an error.
pub(crate) fn parse_symbol_fonts(table: &[u8], fc: u32, lcb: u32) -> Vec<Option<SymbolFont>> {
    let mut fonts = Vec::new();
    for_each_font(table, fc, lcb, |ffn| {
        fonts.push(symbol_font(&ffn_name(ffn)));
    });
    fonts
}

/// The name of the font an `FFN` describes.
fn ffn_name(ffn: &[u8]) -> String {
    char::decode_utf16(
        ffn.get(FFN_FIXED_LEN..)
            .unwrap_or_default()
            .as_chunks::<2>()
            .0
            .iter()
            .map(|c| u16::from_le_bytes(*c))
            .take_while(|&u| u != 0)
            .take(64),
    )
    .map(|r| r.unwrap_or('\u{FFFD}'))
    .collect()
}

/// Call `f` with each `FFN` of the `SttbfFfn` at `fc`/`lcb` in the table
/// stream, in order. Bounded by the structure and the stream: a damaged
/// table gives fewer fonts, never an error.
fn for_each_font(table: &[u8], fc: u32, lcb: u32, mut f: impl FnMut(&[u8])) {
    let start = fc as usize;
    if lcb < 4 || start >= table.len() {
        return;
    }
    let end = start.saturating_add(lcb as usize).min(table.len());
    let sttb = &table[start..end];
    // The size is clamped to the stream, so fewer than the four header bytes
    // can remain.
    let [c0, c1, e0, e1, ..] = *sttb else {
        return;
    };
    let c_data = usize::from(u16::from_le_bytes([c0, c1])).min(MAX_FONTS);
    let cb_extra = usize::from(u16::from_le_bytes([e0, e1]));
    let mut pos = 4usize;
    for _ in 0..c_data {
        let Some(&cch) = sttb.get(pos) else { break };
        let cch = usize::from(cch);
        let Some(ffn) = sttb.get(pos + 1..pos + 1 + cch) else {
            break;
        };
        f(ffn);
        pos += 1 + cch + cb_extra;
    }
}

/// The language a damaged header leaves open, taken from what the font table
/// declares: each `FFN` states the character set (`chs`, [MS-DOC] §2.9.102)
/// its font was chosen for, and the 8-bit text of a Cyrillic, Greek, Hebrew,
/// Arabic, Baltic, Central European, Turkish, Vietnamese, Thai or East Asian
/// document is in the code page of those fonts. The character set most of the
/// fonts that state a non-Western one name decides, the first named winning a
/// tie; `None` when no font states one. Returns that character set and a
/// language ID with its code page.
pub(crate) fn language_from_font_charsets(table: &[u8], fc: u32, lcb: u32) -> Option<(u8, u16)> {
    let mut counts: Vec<(u8, usize)> = Vec::new();
    for_each_font(table, fc, lcb, |ffn| {
        let Some(&chs) = ffn.get(3) else { return };
        if super::codepage::lid_for_charset(chs).is_none() {
            return;
        }
        match counts.iter_mut().find(|(c, _)| *c == chs) {
            Some((_, n)) => *n += 1,
            None => counts.push((chs, 1)),
        }
    });
    let best = counts.iter().map(|&(_, n)| n).max()?;
    let (chs, _) = *counts.iter().find(|&&(_, n)| n == best)?;
    Some((chs, super::codepage::lid_for_charset(chs)?))
}

/// The character positions whose run is in a symbol font.
#[derive(Debug, Clone)]
pub(crate) struct SymbolRuns {
    /// Runs that state their font (directly or through a character style):
    /// `(start, end, font)`, sorted by `start`. `None` is a run set in some
    /// other font, which must not fall through to its paragraph's.
    runs: Vec<(u32, u32, Option<SymbolFont>)>,
    /// Paragraphs whose style gives them a symbol font, sorted by `start`.
    paragraphs: Vec<(u32, u32, SymbolFont)>,
    /// No range puts any character in a symbol font.
    empty: bool,
}

impl SymbolRuns {
    /// No run is in a symbol font.
    pub(crate) fn none() -> Self {
        Self {
            runs: Vec::new(),
            paragraphs: Vec::new(),
            empty: true,
        }
    }

    /// The symbol-font ranges among `cp_runs` (every character run of the
    /// document with its properties, as character positions), given the
    /// document's `fonts` ([`parse_symbol_fonts`]).
    ///
    /// A run's font is its own `sprmCRgFtc2`, else the one its character
    /// style gives; a run with neither is in its paragraph's style font, from
    /// `paragraphs` (`(start, end, istd)`), so a symbol-font paragraph style
    /// reaches the runs that carry no font of their own ([MS-DOC] §2.4.6.1
    /// applies the paragraph style, then the character style, then the run).
    pub(crate) fn from_runs(
        cp_runs: &[(u32, u32, ChpProps)],
        fonts: &[Option<SymbolFont>],
        styles: &[StyleDef],
        paragraphs: &[(u32, u32, u16)],
    ) -> Self {
        if fonts.iter().all(Option::is_none) {
            return Self::none();
        }
        let font_of = |index: u16| fonts.get(usize::from(index)).copied().flatten();
        let mut runs: Vec<_> = cp_runs
            .iter()
            .filter(|(start, end, _)| end > start)
            .filter_map(|(start, end, props)| {
                let index = props
                    .ftc_other
                    .or_else(|| style_ftc_other(styles, props.istd_char?))?;
                Some((*start, *end, font_of(index)))
            })
            .collect();
        runs.sort_by_key(|&(start, ..)| start);
        let mut paragraphs: Vec<_> = paragraphs
            .iter()
            .filter(|(start, end, _)| end > start)
            .filter_map(|&(start, end, istd)| {
                Some((start, end, font_of(style_ftc_other(styles, istd)?)?))
            })
            .collect();
        paragraphs.sort_by_key(|&(start, ..)| start);
        let empty = paragraphs.is_empty() && runs.iter().all(|r| r.2.is_none());
        Self {
            runs,
            paragraphs,
            empty,
        }
    }

    pub(crate) fn is_empty(&self) -> bool {
        self.empty
    }

    /// The symbol font of the run that covers `cp`, if any.
    fn font_at(&self, cp: u32) -> Option<SymbolFont> {
        let after = self.runs.partition_point(|&(start, ..)| start <= cp);
        if let Some(&(_, end, font)) = after.checked_sub(1).and_then(|i| self.runs.get(i)) {
            if cp < end {
                return font;
            }
        }
        let after = self.paragraphs.partition_point(|&(start, ..)| start <= cp);
        let &(_, end, font) = self.paragraphs.get(after.checked_sub(1)?)?;
        (cp < end).then_some(font)
    }

    /// The font's character for the UTF-16 unit `unit` at `cp`: only a unit in
    /// the private-use range `U+F020..=U+F0FF`, in a symbol-font run, that the
    /// font has a character for.
    pub(crate) fn map_unit(&self, cp: u32, unit: u16) -> Option<char> {
        if !(0xF020..=0xF0FF).contains(&unit) {
            return None;
        }
        char_for(self.font_at(cp)?, (unit - 0xF000) as u8)
    }
}

/// Append the text of UTF-16 `units`, the first at character position
/// `first_cp`, with the symbol-font characters replaced.
pub(crate) fn push_utf16(out: &mut String, units: &[u16], first_cp: u32, symbols: &SymbolRuns) {
    if symbols.is_empty() {
        out.push_str(&String::from_utf16_lossy(units));
        return;
    }
    let mut from = 0;
    for (i, &unit) in units.iter().enumerate() {
        if let Some(c) = symbols.map_unit(first_cp.wrapping_add(i as u32), unit) {
            out.push_str(&String::from_utf16_lossy(&units[from..i]));
            out.push(c);
            from = i + 1;
        }
    }
    out.push_str(&String::from_utf16_lossy(&units[from..]));
}

#[cfg(test)]
mod tests {
    use super::*;
    use SymbolFont::*;

    /// An `SttbfFfn` holding fonts named `names`.
    fn sttbf(names: &[&str]) -> Vec<u8> {
        let mut out = (names.len() as u16).to_le_bytes().to_vec();
        out.extend_from_slice(&0u16.to_le_bytes());
        for name in names {
            let mut ffn = vec![0u8; FFN_FIXED_LEN];
            for unit in name.encode_utf16().chain(std::iter::once(0)) {
                ffn.extend_from_slice(&unit.to_le_bytes());
            }
            out.push(ffn.len() as u8);
            out.extend_from_slice(&ffn);
        }
        out
    }

    #[test]
    fn test_font_names_are_read_from_the_table() {
        let table = sttbf(&[
            "Times New Roman",
            "Symbol",
            "Wingdings 2",
            "Webdings",
            "SYMBOL",
        ]);
        let fonts = parse_symbol_fonts(&table, 0, table.len() as u32);
        assert_eq!(
            fonts,
            vec![
                None,
                Some(Symbol),
                Some(Wingdings2),
                Some(Webdings),
                Some(Symbol)
            ]
        );
    }

    #[test]
    fn test_the_table_is_read_at_its_offset_and_bounded_by_its_size() {
        let mut table = vec![0xAAu8; 10];
        let at = table.len() as u32;
        table.extend(sttbf(&["Symbol", "Wingdings"]));
        let size = (table.len() - 10) as u32;
        assert_eq!(parse_symbol_fonts(&table, at, size), vec![Some(Symbol), Some(Wingdings)]);
        // A size that ends inside the second entry gives the first only.
        assert_eq!(parse_symbol_fonts(&table, at, size - 6), vec![Some(Symbol)]);
    }

    #[test]
    fn test_a_damaged_table_gives_a_short_list_not_an_error() {
        assert!(parse_symbol_fonts(&[], 0, 100).is_empty());
        assert!(parse_symbol_fonts(&[1, 2, 3], 0, 3).is_empty());
        assert!(parse_symbol_fonts(&sttbf(&["Symbol"]), 500, 20).is_empty());
        // cData claims 30,000 fonts over a table that holds one.
        let mut table = sttbf(&["Symbol"]);
        table[0..2].copy_from_slice(&30_000u16.to_le_bytes());
        assert_eq!(parse_symbol_fonts(&table, 0, table.len() as u32), vec![Some(Symbol)]);
        // An entry too short to hold the fixed part has no name.
        let short = [1u8, 0, 0, 0, 3, 1, 2, 3];
        assert_eq!(parse_symbol_fonts(&short, 0, short.len() as u32), vec![None]);
    }

    /// A font table whose offset lies within the last bytes of the stream leaves
    /// fewer than the four header bytes after clamping: no panic, no fonts.
    #[test]
    fn test_a_font_table_at_the_very_end_of_the_stream_is_ignored() {
        let table = vec![0x55u8; 10];
        for fc in 0..=12u32 {
            for lcb in [0u32, 1, 3, 4, 5, 100, u32::MAX] {
                let _ = parse_symbol_fonts(&table, fc, lcb);
            }
        }
        assert!(parse_symbol_fonts(&table, 8, 100).is_empty());
        assert!(parse_symbol_fonts(&table, 9, 4).is_empty());
        assert!(parse_symbol_fonts(&table, u32::MAX, u32::MAX).is_empty());
    }

    fn props(ftc: Option<u16>) -> ChpProps {
        ChpProps {
            ftc_other: ftc,
            ..ChpProps::default()
        }
    }

    #[test]
    fn test_runs_in_symbol_fonts_are_found_by_character_position() {
        let fonts = [None, Some(Symbol), Some(Wingdings)];
        let runs = [
            (0, 5, props(Some(1))),
            (5, 8, props(Some(0))),
            (8, 12, props(Some(2))),
            (12, 20, props(None)),
            (20, 25, props(Some(7))), // past the font table
        ];
        let symbols = SymbolRuns::from_runs(&runs, &fonts, &[], &[]);
        assert_eq!(symbols.font_at(0), Some(Symbol));
        assert_eq!(symbols.font_at(4), Some(Symbol));
        assert_eq!(symbols.font_at(5), None);
        assert_eq!(symbols.font_at(8), Some(Wingdings));
        assert_eq!(symbols.font_at(11), Some(Wingdings));
        assert_eq!(symbols.font_at(12), None);
        assert_eq!(symbols.font_at(22), None);
        assert!(SymbolRuns::from_runs(&runs, &[None, None], &[], &[]).is_empty());
    }

    #[test]
    fn test_only_private_use_units_in_a_symbol_run_are_mapped() {
        let symbols = SymbolRuns::from_runs(&[(0, 4, props(Some(0)))], &[Some(Symbol)], &[], &[]);
        assert_eq!(symbols.map_unit(1, 0xF06D), Some('\u{3bc}'));
        assert_eq!(symbols.map_unit(1, 0x006D), None); // not private-use
        assert_eq!(symbols.map_unit(1, 0xF0FF), None); // no character for it
        assert_eq!(symbols.map_unit(9, 0xF06D), None); // outside the run
    }

    #[test]
    fn test_utf16_text_is_rewritten_around_the_mapped_units() {
        let symbols =
            SymbolRuns::from_runs(&[(10, 12, props(Some(0)))], &[Some(Wingdings)], &[], &[]);
        let mut out = String::new();
        // "a", Wingdings clock (non-BMP), "b" -- the clock is at cp 11.
        push_utf16(&mut out, &[0x61, 0x61, 0xF0B7, 0x62], 9, &symbols);
        assert_eq!(out, "aa\u{1F550}b");
        let mut plain = String::new();
        push_utf16(&mut plain, &[0xD83D, 0xDE00, 0xF0B7], 0, &SymbolRuns::none());
        assert_eq!(plain, "\u{1F600}\u{F0B7}"); // a surrogate pair and an untouched unit
    }

    fn style(base: Option<u16>, ftc: Option<u16>) -> StyleDef {
        StyleDef {
            base,
            own_ftc_other: ftc,
            ..StyleDef::default()
        }
    }

    #[test]
    fn test_a_paragraph_style_font_reaches_runs_that_state_none() {
        let fonts = [None, Some(Symbol)];
        // istd 1 sets font 1 (Symbol); istd 2 is based on it; istd 0 sets none.
        let styles = [
            style(None, None),
            style(None, Some(1)),
            style(Some(1), None),
        ];
        let paragraphs = [(0, 10, 2), (10, 20, 0)];
        let runs = [
            (0, 4, props(None)),
            (4, 6, props(Some(0))), // a run of its own, in no symbol font
        ];
        let symbols = SymbolRuns::from_runs(&runs, &fonts, &styles, &paragraphs);
        assert_eq!(symbols.font_at(2), Some(Symbol), "inherited through the base");
        assert_eq!(symbols.font_at(5), None, "the run's own font wins");
        assert_eq!(symbols.font_at(8), Some(Symbol));
        assert_eq!(symbols.font_at(12), None, "a paragraph whose style names no font");
        assert!(!symbols.is_empty());
    }

    #[test]
    fn test_a_character_style_font_wins_over_the_paragraph_style() {
        let fonts = [None, Some(Symbol), Some(Wingdings)];
        let styles = [
            style(None, None),
            style(None, Some(1)),
            style(None, Some(2)),
        ];
        let runs = [(
            0,
            4,
            ChpProps {
                istd_char: Some(2),
                ..ChpProps::default()
            },
        )];
        let symbols = SymbolRuns::from_runs(&runs, &fonts, &styles, &[(0, 10, 1)]);
        assert_eq!(symbols.font_at(1), Some(Wingdings));
        assert_eq!(symbols.font_at(6), Some(Symbol));
    }

    #[test]
    fn test_a_cyclic_style_chain_gives_no_font() {
        let fonts = [Some(Symbol)];
        let styles = [style(Some(1), None), style(Some(0), None)];
        let symbols = SymbolRuns::from_runs(&[], &fonts, &styles, &[(0, 10, 0)]);
        assert!(symbols.is_empty());
    }

    /// A font table of fonts with the given character sets.
    fn table_with_charsets(sets: &[u8]) -> Vec<u8> {
        let mut out = (sets.len() as u16).to_le_bytes().to_vec();
        out.extend_from_slice(&0u16.to_le_bytes());
        for &chs in sets {
            let mut ffn = vec![0u8; FFN_FIXED_LEN];
            ffn[3] = chs;
            ffn.extend_from_slice(&[b'A', 0, 0, 0]);
            out.push(ffn.len() as u8);
            out.extend_from_slice(&ffn);
        }
        out
    }

    #[test]
    fn test_the_language_follows_the_character_set_most_fonts_declare() {
        let lang = |sets: &[u8]| {
            let t = table_with_charsets(sets);
            language_from_font_charsets(&t, 0, t.len() as u32)
        };
        assert_eq!(lang(&[0xCC, 0xCC, 0x00]), Some((0xCC, 0x0419)));
        // Western, symbol, OEM and default sets name no other language.
        assert_eq!(lang(&[0x00, 0x01, 0x02, 0xFF]), None);
        // The most named wins; the first named wins a tie.
        assert_eq!(lang(&[0xA1, 0xB1, 0xB1]), Some((0xB1, 0x040D)));
        assert_eq!(lang(&[0xA1, 0xB1]), Some((0xA1, 0x0408)));
        assert_eq!(lang(&[]), None);
    }

    #[test]
    fn test_a_damaged_font_table_gives_no_language_and_no_panic() {
        let t = table_with_charsets(&[0xCC, 0xCC]);
        for cut in 0..t.len() {
            let _ = language_from_font_charsets(&t[..cut], 0, cut as u32);
        }
        assert_eq!(language_from_font_charsets(&t, u32::MAX, u32::MAX), None);
    }
}
