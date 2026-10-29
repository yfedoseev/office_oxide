//! Style sheet (STSH) parsing for legacy binary `.doc` (MS-DOC §2.7.1).
//!
//! The style sheet maps a paragraph's style index (`istd`) to a built-in
//! style id (`sti`), a name, the style it is based on (`istdBase`) and the
//! outline level its own paragraph properties state. Built-in heading styles
//! carry `sti` 1–9 (Heading 1–9); user-defined heading styles are named
//! `Heading N`; any other style may state an outline level itself or inherit
//! one from its base. All of these let us derive a paragraph's real heading
//! level instead of the line heuristic in `convert_doc.rs`.
//!
//! Every parse step is bounds-checked: a malformed style sheet yields no styles
//! and a truncated one yields the styles read before the cut, so callers
//! degrade to "no style" (and the heuristic) rather than panicking (AGENTS.md
//! rule 6). Either way [`StyleSheet::warning`] says why, so the loss reaches
//! `Metadata::warnings` instead of passing unnoticed.

use super::{
    MAX_OUTLINE_LEVEL, OutlineLevel,
    fib::Fib,
    sprm::{extract_chp_props, extract_pap_props},
};

/// One style definition, indexed by `istd`.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct StyleDef {
    /// Built-in style id (`StdfBase.sti`). `0x0FFE` means user-defined.
    pub sti: u16,
    /// Style name (from the style's `xstzName` inside `STD`, [MS-DOC] §2.9.271).
    pub name: String,
    /// `istdBase` ([MS-DOC] §2.9.260 `StdfBase`): the style this one inherits
    /// from. `None` when the field is `0x0FFF` ("no base").
    pub base: Option<u16>,
    /// The outline level the style's *own* paragraph properties state: the
    /// `sprmPOutLvl` (0x2640) in its `UpxPapx`. Zero-based like the SPRM
    /// operand: `0..=8` are Heading 1–9 and `9` is explicit body text. `None`
    /// when the style states none (it may still inherit one from `base`).
    pub own_outline: Option<u8>,
    /// Style kind (`StdfBase.stk`, [MS-DOC] §2.9.253): 1 paragraph, 2
    /// character, 3 table, 4 numbering. 0 for an empty slot.
    pub stk: u16,
    /// The font-table index this style's own character properties give to
    /// characters outside U+0020..U+007F (`sprmCRgFtc2`, [MS-DOC] §2.6.1), from
    /// its `UpxChpx`. `None` when it states none (it may inherit one).
    pub own_ftc_other: Option<u16>,
    /// The name exposed on `Paragraph::style`, worked out once when the style
    /// sheet is parsed (see [`assign_display_names`]) so that a lookup per
    /// paragraph does not re-examine a name that can be tens of kilobytes.
    /// `None` for a style that is not a paragraph style, one with a blank or
    /// unusable name, and the built-ins reported under a canonical name
    /// (those do not need one stored).
    pub(crate) display_name: Option<String>,
}

/// `StdfBase.stk` of a paragraph style ([MS-DOC] §2.9.253).
const STK_PARAGRAPH: u16 = 1;

/// The canonical English name of a built-in paragraph style that `.docx` also
/// normalizes, from its `sti` ([MS-DOC] §2.9.260): 1–9 Heading 1–9, 31
/// Header, 32 Footer, 34 Caption.
fn canonical_name(sti: u16) -> Option<String> {
    match sti {
        n @ 1..=9 => Some(format!("Heading {n}")),
        31 => Some("Header".to_owned()),
        32 => Some("Footer".to_owned()),
        34 => Some("Caption".to_owned()),
        _ => None,
    }
}

/// Work out every paragraph style's exposed name, once. Returns how many
/// stored names cannot be exposed (longer than 255 characters or holding a
/// control character).
///
/// A name may be stored as `"primary,alias,alias"` ([MS-DOC] §2.9.258); only
/// the primary is the style's name. `istd` 0 is the default style however it is
/// named, so it never gets one.
fn assign_display_names(styles: &mut [StyleDef]) -> usize {
    use crate::limits::{StyleName, classify_style_name};
    let mut unusable = 0;
    for style in styles.iter_mut().skip(1) {
        style.display_name = None;
        if style.stk != STK_PARAGRAPH {
            continue;
        }
        if let Some(name) = canonical_name(style.sti) {
            style.display_name = Some(name);
            continue;
        }
        match classify_style_name(style.name.split(',').next().unwrap_or("")) {
            StyleName::Usable(name) => style.display_name = Some(name),
            StyleName::Unusable => unusable += 1,
            StyleName::Absent => {},
        }
    }
    unusable
}

/// The display name of style `istd`, or `None` for the default style, a style
/// that is not a paragraph style, a style with no usable name, or an index
/// outside the sheet.
///
/// `istd` 0 is the default paragraph style (`Normal`, [MS-DOC] §2.9.260
/// `StdfBase.sti` 0) however the producing application named it: the name is
/// localized, so it is the position and not the name that identifies it. A
/// paragraph's `istd` can name a character, table or numbering style
/// (`StdfBase.stk`, §2.9.253); only a paragraph style is the paragraph's style,
/// as in `.docx`. A built-in style is identified by its `sti` and its stored
/// name is the producing Word's language, so the styles `.docx` also
/// normalizes (Heading 1–9, Header, Footer, Caption) are reported under their
/// canonical English name; the others keep the stored name.
pub fn style_name_for_istd(styles: &[StyleDef], istd: u16) -> Option<String> {
    styles.get(usize::from(istd))?.display_name.clone()
}

/// The longest `istdBase` chain followed when resolving a style's heading
/// level. Real style sheets nest a handful of levels; the bound is what ends a
/// cyclic chain (or a hostile one thousands deep) without a visited set, and
/// it caps the work per paragraph.
const MAX_STYLE_CHAIN: usize = 32;

/// The styles read from an STSH, and why the read fell short, if it did.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct StyleSheet {
    /// Styles indexed by `istd`. Empty when the sheet was rejected; the styles
    /// read before the cut when it was truncated.
    pub styles: Vec<StyleDef>,
    /// Why the sheet could not be used in full: absent, rejected, cut short, or
    /// carrying base chains that cannot be followed. Surfaced through
    /// `Metadata::warnings`: style-derived headings are unavailable (or only
    /// partly available) for this document.
    pub warnings: Vec<String>,
}

impl StyleSheet {
    fn rejected(reason: String) -> Self {
        Self {
            styles: Vec::new(),
            warnings: vec![format!(
                "style sheet (STSH) {reason}; heading levels from paragraph styles are unavailable"
            )],
        }
    }
}

/// Parse the document style sheet (STSH) from the Table stream.
///
/// Yields no styles, with a warning, when the style sheet is absent
/// (`lcbStshf == 0`), out of bounds, or malformed.
///
/// Only the *length* is a valid absence test. `fcStshf` is an offset into
/// the Table stream and `0` is the usual value, because the style sheet is
/// normally the first thing written there; [MS-DOC] `FibRgFcLcb97` requires
/// `lcbStshf` to be non-zero and `STSH` states every FIB contains a style
/// sheet. Testing the offset silently discarded the style sheet of nearly
/// every real document.
pub fn parse_style_sheet(table_stream: &[u8], fib: &Fib) -> StyleSheet {
    // Guard on the length only: `fc_stshf` is an offset and 0 is its normal
    // value, so treating 0 as "absent" drops real style sheets.
    if fib.lcb_stshf == 0 {
        return StyleSheet::rejected("is absent (lcbStshf is 0)".into());
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
        return StyleSheet::rejected(format!(
            "lies outside the Table stream (fcStshf {start}, lcbStshf {}, stream {} bytes)",
            fib.lcb_stshf,
            table_stream.len()
        ));
    }
    parse_stsh(&table_stream[start..end])
}

/// Resolve a paragraph's `istd` (optionally overridden by `sprmPIstd`) to a
/// heading level (1–9), or `None` when the style is not a heading.
///
/// A style's level is, in order: the outline level its own paragraph
/// properties state (`9` meaning explicit body text, which also stops the
/// walk), else Heading 1–9 by built-in `sti`, else by a `Heading N` name, else
/// whatever the style it is based on resolves to. The base chain is followed at
/// most [`MAX_STYLE_CHAIN`] steps, so a cycle resolves to `None`.
pub fn heading_level_for_istd(styles: &[StyleDef], istd: u16) -> Option<u8> {
    resolve_heading_level(styles, istd).0
}

/// The level of `istd` and whether the walk gave up at [`MAX_STYLE_CHAIN`]
/// steps (a cycle, or a chain too deep to be real) before it could decide.
fn resolve_heading_level(styles: &[StyleDef], istd: u16) -> (Option<u8>, bool) {
    let mut current = istd;
    for _ in 0..MAX_STYLE_CHAIN {
        let Some(s) = styles.get(current as usize) else {
            return (None, false);
        };
        match s.own_outline {
            Some(zero_based) if zero_based < MAX_OUTLINE_LEVEL => {
                return (Some(zero_based + 1), false);
            },
            Some(MAX_OUTLINE_LEVEL) => return (None, false),
            _ => {},
        }
        // Built-in heading styles: `sti` 1..9 == Heading 1..9.
        if (1..=u16::from(MAX_OUTLINE_LEVEL)).contains(&s.sti) {
            return (Some(s.sti as u8), false);
        }
        // User-defined heading styles are named "Heading N".
        if let Some(level) = heading_level_from_name(&s.name) {
            return (Some(level), false);
        }
        match s.base {
            Some(base) => current = base,
            None => return (None, false),
        }
    }
    (None, true)
}

fn heading_level_from_name(name: &str) -> Option<u8> {
    // Real Word style names are "Heading N" (capital H); match case-insensitively
    // so "heading 2" / "HEADING 2" also resolve.
    //
    // `xstzName` may carry aliases as "primary,alias,alias" (MS-DOC §2.9.258);
    // a heading level is always carried by the primary name, so look at the
    // segment before the first comma.
    //
    // Matched **without allocating**: this is the fallback for every paragraph
    // whose style is not a built-in `sti` 1–9 heading, so a
    // `to_ascii_lowercase()` here would cost one `String` per paragraph in the
    // document. The empty name is the common case (fixed-index slots), and
    // fails the length check immediately.
    const PREFIX: &[u8] = b"heading ";
    let primary = name.split(',').next().unwrap_or("").trim();
    if primary.len() < PREFIX.len()
        || !primary.as_bytes()[..PREFIX.len()].eq_ignore_ascii_case(PREFIX)
    {
        return None;
    }
    let level: u8 = primary[PREFIX.len()..].trim().parse().ok()?;
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
fn parse_stsh(data: &[u8]) -> StyleSheet {
    if data.len() < 2 {
        return StyleSheet::rejected("is too short to hold its header".into());
    }
    let cb_stshi = u16::from_le_bytes([data[0], data[1]]) as usize;
    let mut pos = 2;
    if cb_stshi < 18 || pos + cb_stshi > data.len() {
        return StyleSheet::rejected(format!(
            "header rejected (cbStshi {cb_stshi} is below 18 or runs past the sheet)"
        ));
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
        return StyleSheet::rejected(format!(
            "rejected (cbSTDBaseInFile {cb_std_base:#06x} is neither 0x000A nor 0x0012)"
        ));
    }

    // `rglpstd`: `cstd` `LPStd` entries (§2.9.135).
    // Every entry takes at least its 2-byte `cbStd`, so no more than half the
    // remaining bytes can be entries: a lying `cstd` reserves nothing like 4096.
    let cap = cstd.min(4096).min(data.len() / 2);
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
        let (base, own_outline) = style_inheritance(std, cb_std_base);
        // `StdfBase.stk` is the low 4 bits of the second u16 (§2.9.253).
        let stk = if std.len() >= 4 {
            u16::from_le_bytes([std[2], std[3]]) & 0x000F
        } else {
            0
        };
        styles.push(StyleDef {
            sti,
            name,
            base,
            own_outline,
            stk,
            own_ftc_other: style_own_ftc_other(std, cb_std_base),
            display_name: None,
        });
        pos += cb_std;
        // LPStd entries are stored on even-byte boundaries and `cbStd` does NOT
        // include that padding byte (§2.9.135).
        if !cb_std.is_multiple_of(2) {
            pos += 1;
        }
    }
    // The loop stops early on a truncated entry and is capped at 4096 entries,
    // so one comparison covers a cut sheet, a lying `cstd`, and the cap. The
    // styles already read are kept: they are correct, only incomplete.
    let mut warnings = Vec::new();
    if styles.len() < cstd {
        warnings.push(format!(
            "style sheet (STSH) declares {cstd} styles but only {} could be read; \
             heading levels from the remaining paragraph styles are unavailable",
            styles.len()
        ));
    }
    let unfollowable = (0..styles.len())
        .filter(|&i| resolve_heading_level(&styles, i as u16).1)
        .count();
    if unfollowable > 0 {
        warnings.push(format!(
            "style sheet (STSH) has {unfollowable} styles whose istdBase chain is cyclic or \
             deeper than {MAX_STYLE_CHAIN}, so their heading levels are unavailable"
        ));
    }
    let unusable = assign_display_names(&mut styles);
    if unusable > 0 {
        warnings.push(format!(
            "{unusable} style name(s) were not exposed because they exceed 255 characters or \
             contain control characters"
        ));
    }
    StyleSheet { styles, warnings }
}

/// A style's `istdBase` and the outline level its own paragraph properties
/// state, read from the `STD` ([MS-DOC] §2.9.258).
///
/// `StdfBase` holds `stk` in the low 4 bits of its second word and `istdBase` in
/// the high 12; the third word's low 4 bits are `cupx`, the number of `UPX`
/// entries that follow the name. Only a paragraph style (`stk` 1) carries a
/// `UpxPapx`, which is the first `LPUpx` after the name: `cbUpx(u16)`, then
/// `istd(u16)`, then the paragraph `grpprl`. Any bound that does not hold
/// yields no outline level rather than a guess.
fn style_inheritance(std: &[u8], cb_std_base: usize) -> (Option<u16>, Option<u8>) {
    if std.len() < 6 {
        return (None, None);
    }
    let word1 = u16::from_le_bytes([std[2], std[3]]);
    let stk = word1 & 0x000F;
    let base = (word1 >> 4 != 0x0FFF).then_some(word1 >> 4);
    let cupx = u16::from_le_bytes([std[4], std[5]]) & 0x000F;
    if stk != STK_PARAGRAPH || cupx == 0 {
        return (base, None);
    }
    // `xstzName` is `cch(u16)` + `cch` UTF-16 units + a 2-byte terminator; the
    // `LPUpx` array starts on the next even offset within the `STD`.
    let Some(cch) = std
        .get(cb_std_base..cb_std_base.saturating_add(2))
        .map(|b| usize::from(u16::from_le_bytes([b[0], b[1]])))
    else {
        return (base, None);
    };
    let mut at = cb_std_base + 2 + 2 * cch + 2;
    if !at.is_multiple_of(2) {
        at += 1;
    }
    let Some(cb_upx) = std
        .get(at..at.saturating_add(2))
        .map(|b| usize::from(u16::from_le_bytes([b[0], b[1]])))
    else {
        return (base, None);
    };
    // `UpxPapx`: the style's own `istd` (2 bytes), then its `grpprl`.
    let Some(grpprl) = std
        .get(at + 2..at + 2 + cb_upx)
        .filter(|upx| upx.len() >= 2)
        .map(|upx| &upx[2..])
    else {
        return (base, None);
    };
    let own_outline = match extract_pap_props(grpprl).outline_level {
        Some(OutlineLevel::Heading { level, .. }) => Some(level),
        Some(OutlineLevel::BodyText) => Some(MAX_OUTLINE_LEVEL),
        None => None,
    };
    (base, own_outline)
}

/// `STK_CHARACTER`: `StdfBase.stk` of a character style ([MS-DOC] §2.9.253).
const STK_CHARACTER: u16 = 2;

/// The font a style's own `UpxChpx` gives characters outside U+0020..U+007F.
///
/// After the name come `cupx` `LPUpx` entries (`cbUpx(u16)`, then `cbUpx`
/// bytes, each starting on an even offset within the `STD`, [MS-DOC] §2.9.258).
/// A paragraph style holds a `UpxPapx` then a `UpxChpx`; a character style
/// holds a `UpxChpx` only. The `UpxChpx` is a bare character `grpprl`
/// ([MS-DOC] §2.9.328). Any bound that does not hold yields `None`.
fn style_own_ftc_other(std: &[u8], cb_std_base: usize) -> Option<u16> {
    if std.len() < 6 {
        return None;
    }
    let stk = u16::from_le_bytes([std[2], std[3]]) & 0x000F;
    let cupx = usize::from(u16::from_le_bytes([std[4], std[5]]) & 0x000F);
    let chpx_index = match stk {
        STK_PARAGRAPH => 1,
        STK_CHARACTER => 0,
        _ => return None,
    };
    if cupx <= chpx_index {
        return None;
    }
    let cch = usize::from(u16::from_le_bytes(
        std.get(cb_std_base..cb_std_base.checked_add(2)?)?
            .try_into()
            .ok()?,
    ));
    let mut at = cb_std_base.checked_add(2 + 2 * cch + 2)?;
    for index in 0..=chpx_index {
        at += at % 2;
        let cb_upx =
            usize::from(u16::from_le_bytes(std.get(at..at.checked_add(2)?)?.try_into().ok()?));
        let data = std.get(at + 2..(at + 2).checked_add(cb_upx)?)?;
        if index == chpx_index {
            return extract_chp_props(data).ftc_other;
        }
        at += 2 + cb_upx;
    }
    None
}

/// The font-table index that `istd`'s character properties give characters
/// outside U+0020..U+007F, following `istdBase` for at most
/// [`MAX_STYLE_CHAIN`] steps (a cycle resolves to `None`).
pub fn style_ftc_other(styles: &[StyleDef], istd: u16) -> Option<u16> {
    let mut current = istd;
    for _ in 0..MAX_STYLE_CHAIN {
        let style = styles.get(usize::from(current))?;
        if style.own_ftc_other.is_some() {
            return style.own_ftc_other;
        }
        current = style.base?;
    }
    None
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
    /// A paragraph style (`sgc` 1) based on nothing (`istdBase` 0x0FFF) and
    /// stating no outline level of its own; see [`lpstd_styled`].
    fn lpstd(sti: u16, name: &str) -> Vec<u8> {
        lpstd_styled(sti, name, None, None)
    }

    /// Build one `LPStd` for a paragraph style: `cbStd(u16)` + `STD`.
    ///
    /// `STD` = `stdf` (`cbSTDBaseInFile` bytes, here `StdfBase` alone) +
    /// `xstzName` + `grLPUpxSw`, per MS-DOC §2.9.258. `xstzName` is an `Xstz`:
    /// `cch(u16)` + `cch` UTF-16 code units + a 2-byte null terminator
    /// (§2.9.354). `StdfBase` carries `sti` (low 12 bits of word 0), `sgc` (low 4
    /// bits of word 1) with `istdBase` above it, and `cupx` (low 4 bits of word
    /// 2). When `outline` is given the style states it in its own `UpxPapx`: an
    /// `LPUpx` (`cbUpx` + the style's `istd` + a `sprmPOutLvl` grpprl), starting
    /// on an even offset within the `STD`.
    fn lpstd_styled(sti: u16, name: &str, base: Option<u16>, outline: Option<u8>) -> Vec<u8> {
        const CB_STD_BASE: usize = 0x000A;
        const SGC_PARAGRAPH: u16 = 1;
        let mut std = vec![0u8; CB_STD_BASE];
        std[0..2].copy_from_slice(&sti.to_le_bytes());
        let istd_base = base.unwrap_or(0x0FFF);
        std[2..4].copy_from_slice(&(SGC_PARAGRAPH | (istd_base << 4)).to_le_bytes());
        let cupx: u16 = u16::from(outline.is_some());
        std[4..6].copy_from_slice(&cupx.to_le_bytes());
        // `xstzName`.
        let units: Vec<u16> = name.encode_utf16().collect();
        std.extend_from_slice(&(units.len() as u16).to_le_bytes()); // cch
        for u in &units {
            std.extend_from_slice(&u.to_le_bytes());
        }
        std.extend_from_slice(&0u16.to_le_bytes()); // chTerm
        if let Some(op) = outline {
            if std.len() % 2 == 1 {
                std.push(0); // LPUpx starts on an even offset
            }
            let grpprl = [0x40, 0x26, op]; // sprmPOutLvl, 1-byte operand
            let cb_upx = (2 + grpprl.len()) as u16; // istd + grpprl
            std.extend_from_slice(&cb_upx.to_le_bytes());
            std.extend_from_slice(&0u16.to_le_bytes()); // UpxPapx.istd
            std.extend_from_slice(&grpprl);
        }

        let mut out = Vec::new();
        out.extend_from_slice(&(std.len() as u16).to_le_bytes()); // cbStd
        out.extend_from_slice(&std);
        if std.len() % 2 == 1 {
            out.push(0); // the next LPStd starts on an even offset; cbStd excludes this
        }
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
    fn test_parse_stsh_reads_sti_and_names() {
        let styles = parse_stsh(&synthetic_stsh()).styles;
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
    fn test_parse_stsh_rejects_spurious_name_table_layout() {
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

        // An earlier version of this assertion was vacuous: it ran `.all()` over
        // an empty vector, which holds for any implementation. What actually
        // has to be pinned is the *behaviour* — these bytes must not keep
        // producing a Heading 3, which they did when names were attached by
        // index from that table.
        let styles = parse_stsh(&d).styles;
        assert!(
            heading_level_for_istd(&styles, 3).is_none(),
            "a spurious name table must not keep yielding Heading 3; got {styles:?}"
        );
        // Contrast, so the line above cannot pass by doing nothing: the same
        // intent expressed in the spec layout does resolve.
        assert_eq!(
            heading_level_for_istd(&parse_stsh(&synthetic_stsh()).styles, 3),
            Some(3),
            "the spec layout must still resolve istd 3 to Heading 3"
        );
    }

    /// The name must be read from `STD.xstzName`, i.e. at `cbSTDBaseInFile`
    /// bytes into each `STD`. A fixture that shifts the name by even one byte
    /// must not silently produce a plausible-but-wrong name.
    #[test]
    fn test_parse_stsh_rejects_unexpected_cb_std_base_in_file() {
        let mut stsh = synthetic_stsh();
        // `cbSTDBaseInFile` is the u16 at offset 4 (cbStshi + cstd). Spec allows
        // only 0x000A / 0x0012; anything else is malformed.
        stsh[4..6].copy_from_slice(&0x0020u16.to_le_bytes());
        assert!(
            parse_stsh(&stsh).styles.is_empty(),
            "a bogus cbSTDBaseInFile must be rejected, not guess a name offset"
        );
    }

    /// `LPStd` entries sit on even-byte boundaries and `cbStd` excludes the
    /// padding byte (§2.9.135), so an odd-sized `STD` must not desynchronise the
    /// following entries.
    #[test]
    fn test_parse_stsh_skips_odd_sized_lpstd_padding() {
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

        let styles = parse_stsh(&stsh).styles;
        assert_eq!(styles.len(), 2, "both entries must be read");
        assert_eq!(styles[0].sti, 2, "odd-sized entry decodes normally");
        assert_eq!(
            styles[1].name, "Heading 3",
            "the padding byte must not desynchronise the next entry"
        );
    }

    #[test]
    fn test_heading_level_resolves_builtin_and_user() {
        let styles = parse_stsh(&synthetic_stsh()).styles;
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

    /// The built-in `sti` must resolve a heading on its own, without help from
    /// the name. Real Word documents localise the built-in style names — German
    /// "Überschrift 3", French "Titre 3" — and those carry `sti` 1-9 while not
    /// matching `heading N` at all. Without the `sti` branch such a document
    /// would lose every heading.
    ///
    /// This also pins the branch against being masked: every other fixture names
    /// its styles "Heading N", so the name-based path alone would satisfy them.
    #[test]
    fn test_heading_level_resolves_localized_builtin_style_via_sti() {
        let styles = vec![
            StyleDef::default(), // istd 0
            StyleDef {
                sti: 3,
                name: "\u{00dc}berschrift 3".into(),
                ..Default::default()
            }, // istd 1 (de)
            StyleDef {
                sti: 5,
                name: "Titre 5".into(),

                ..Default::default()
            }, // istd 2 (fr)
            StyleDef {
                sti: 0x0FFE,
                name: "Custom".into(),

                ..Default::default()
            }, // istd 3
        ];
        assert_eq!(heading_level_for_istd(&styles, 1), Some(3));
        assert_eq!(heading_level_for_istd(&styles, 2), Some(5));
        // Not a heading: user-defined (sti 0x0FFE) with a non-heading name.
        assert_eq!(heading_level_for_istd(&styles, 3), None);
    }

    /// `xstzName` may carry aliases as "primary,alias" (§2.9.258); the level is
    /// carried by the primary name, so a comma-suffixed name must still resolve.
    #[test]
    fn test_heading_level_uses_primary_name_before_aliases() {
        let styles = vec![
            StyleDef::default(),
            StyleDef {
                sti: 0x0FFE,
                name: "Heading 4,Title 4".into(),

                ..Default::default()
            },
        ];
        assert_eq!(heading_level_for_istd(&styles, 1), Some(4));
    }

    /// Pin the name parser's edges directly, so the allocation-free rewrite
    /// cannot change what resolves. `heading_level_from_name` is the fallback
    /// for every style that is not a built-in `sti` 1–9 heading.
    #[test]
    fn test_heading_level_from_name_resolves_only_primary_heading_names() {
        assert_eq!(heading_level_from_name("Heading 3"), Some(3));
        assert_eq!(heading_level_from_name("heading 3"), Some(3), "case-insensitive");
        assert_eq!(heading_level_from_name("HEADING 3"), Some(3));
        assert_eq!(heading_level_from_name("  Heading 3  "), Some(3), "trims");
        assert_eq!(
            heading_level_from_name("Heading 3,Heading 3 Char"),
            Some(3),
            "only the primary name carries the level"
        );
        assert_eq!(heading_level_from_name(""), None, "empty name is not a heading");
        assert_eq!(heading_level_from_name("   "), None);
        assert_eq!(heading_level_from_name("Not A Heading"), None);
        assert_eq!(heading_level_from_name("Heading"), None, "no level");
        assert_eq!(heading_level_from_name("Heading 0"), None, "0 is out of range");
        assert_eq!(heading_level_from_name("Heading 10"), None, "10 is above MAX_OUTLINE_LEVEL");
        assert_eq!(heading_level_from_name("Headin 3"), None, "prefix must match fully");
    }

    /// A style that is neither a built-in heading (`sti` 1-9) nor named
    /// `Heading N`, and an `istd` past the end of the sheet, resolve to no level.
    #[test]
    fn test_heading_level_for_istd_does_not_resolve_unstyled_paragraphs() {
        let styles = vec![
            StyleDef::default(),
            StyleDef {
                sti: 0x0FFE,
                name: "Body Text".into(),

                ..Default::default()
            },
        ];
        assert_eq!(heading_level_for_istd(&styles, 0), None);
        assert_eq!(heading_level_for_istd(&styles, 1), None);
        assert_eq!(heading_level_for_istd(&styles, 99), None, "out of range istd");
    }

    #[test]
    fn test_truncated_style_sheet_is_empty() {
        assert!(parse_stsh(&[0u8; 4]).styles.is_empty());
        assert!(parse_stsh(&[18, 0, 0, 0]).styles.is_empty());
    }

    /// Build a `Fib` whose only non-zero fields are `fc_stshf` / `lcb_stshf`,
    /// pointing at a style sheet in the Table stream.
    fn fib_with_stsh(start: u32, len: u32) -> Fib {
        Fib {
            fc_stshf: start,
            lcb_stshf: len,
            ..Default::default()
        }
    }

    /// `parse_style_sheet` is the FIB-aware wrapper that `document.rs` calls: it
    /// must slice the STSH out of the Table stream at `fc_stshf` (bounds-checked)
    /// and hand the bytes to `parse_stsh`. This exercises the wrapper end-to-end
    /// with real STSH bytes placed at a non-zero offset — the path the POI corpus
    /// never reaches (every file reports `fc_stshf == 0`).
    #[test]
    fn test_parse_style_sheet_reads_stsh_at_fib_offset() {
        let stsh = synthetic_stsh();
        let start = 64usize;
        let mut table_stream = vec![0u8; start + stsh.len()];
        table_stream[start..start + stsh.len()].copy_from_slice(&stsh);
        let fib = fib_with_stsh(start as u32, stsh.len() as u32);
        let styles = parse_style_sheet(&table_stream, &fib).styles;
        assert_eq!(styles.len(), 16);
        assert_eq!(styles[0].sti, 0);
        assert_eq!(styles[0].name, "Normal");
        assert_eq!(styles[3].sti, 3);
        assert_eq!(styles[3].name, "Heading 3");
        assert_eq!(styles[15].sti, 0x0FFE);
        assert_eq!(styles[15].name, "heading 7");
    }

    /// `fcStshf` is an *offset* into the Table stream, and the style sheet is
    /// normally the first thing written there, so `0` is a **valid** offset, not
    /// a marker for "absent". [MS-DOC] `FibRgFcLcb97` puts the requirement on
    /// the length: `lcbStshf` "MUST be a nonzero value" (and `STSH` adds "Each
    /// FIB MUST contain a stylesheet"). Guarding on the offset therefore
    /// discards the style sheet of essentially every real document.
    ///
    /// This is the case a real `.doc` hits: surveyed over 160 POI corpus files,
    /// 152 report `lcbStshf != 0` and 147 of those report `fcStshf == 0`.
    #[test]
    fn test_parse_style_sheet_at_offset_zero_is_parsed() {
        let stsh = synthetic_stsh();
        // STSH at Table-stream offset 0 — where Word normally writes it.
        let mut table_stream = stsh.clone();
        table_stream.resize(stsh.len() + 64, 0);
        let fib = fib_with_stsh(0, stsh.len() as u32);
        let styles = parse_style_sheet(&table_stream, &fib).styles;
        assert_eq!(styles.len(), 16, "an offset of 0 must not be read as 'no style sheet'");
        assert_eq!(styles[3].name, "Heading 3");
    }

    #[test]
    fn test_parse_style_sheet_absent_is_empty() {
        let table_stream = vec![0u8; 64];
        assert!(
            parse_style_sheet(&table_stream, &fib_with_stsh(0, 0))
                .styles
                .is_empty()
        );
    }

    /// Malformed `fc_stshf` / `lcb_stshf` near `u32::MAX` must not overflow the
    /// pointer arithmetic: on a 32-bit target `start + len` wraps, and a wrapped
    /// slice would read arbitrary memory. AGENTS.md rule 6 — no input may panic
    /// or overflow.
    ///
    /// This pins the *contract* (never panics, yields no styles) rather than the
    /// `saturating_add` itself: on a 64-bit host the sum cannot wrap, so no
    /// assertion here can tell `saturating_add` from `wrapping_add`.
    #[test]
    fn test_parse_style_sheet_extreme_fib_values_do_not_panic() {
        let table_stream = vec![0u8; 64];
        for (start, len) in [(u32::MAX, 100u32), (60u32, u32::MAX), (u32::MAX, u32::MAX)] {
            assert!(
                parse_style_sheet(&table_stream, &fib_with_stsh(start, len))
                    .styles
                    .is_empty(),
                "fc_stshf={start:#x} lcb_stshf={len:#x} must yield no styles"
            );
        }
    }

    /// A style sheet that advertises far more styles than it carries must still
    /// terminate and return what is actually there — the entry loop and the
    /// allocation cap are what stop a lying `cstd` from running away.
    #[test]
    fn test_parse_stsh_huge_cstd_terminates_without_inventing_styles() {
        let mut d = 18u16.to_le_bytes().to_vec(); // cbStshi
        d.extend_from_slice(&0xFFFFu16.to_le_bytes()); // cstd — lies
        d.extend_from_slice(&0x000Au16.to_le_bytes()); // cbSTDBaseInFile
        d.extend_from_slice(&[0u8; 14]);
        d.extend_from_slice(&lpstd(3, "Heading 3"));
        let styles = parse_stsh(&d).styles;
        assert_eq!(styles.len(), 1, "only the one entry actually present");
        assert_eq!(styles[0].name, "Heading 3");
    }

    #[test]
    fn test_parse_style_sheet_out_of_bounds_is_empty() {
        let stsh = synthetic_stsh();
        // Point `fc_stshf` past the end of the table stream.
        assert!(
            parse_style_sheet(&stsh, &fib_with_stsh(1000, stsh.len() as u32))
                .styles
                .is_empty()
        );
    }

    #[test]
    fn test_a_well_formed_style_sheet_carries_no_warning() {
        let sheet = parse_stsh(&synthetic_stsh());
        assert_eq!(sheet.styles.len(), 16);
        assert!(sheet.warnings.is_empty(), "{:?}", sheet.warnings);
    }

    #[test]
    fn test_a_rejected_style_sheet_says_why() {
        let mut stsh = synthetic_stsh();
        stsh[4..6].copy_from_slice(&0x0020u16.to_le_bytes()); // cbSTDBaseInFile
        let sheet = parse_stsh(&stsh);
        assert!(sheet.styles.is_empty());
        let w = sheet
            .warnings
            .first()
            .cloned()
            .expect("a rejected sheet must warn");
        assert!(w.contains("cbSTDBaseInFile 0x0020"), "{w}");

        let w = parse_stsh(&[0u8; 1])
            .warnings
            .first()
            .cloned()
            .expect("too short");
        assert!(w.contains("too short"), "{w}");

        let mut stsh = synthetic_stsh();
        stsh[0..2].copy_from_slice(&4u16.to_le_bytes()); // cbStshi < 18
        let w = parse_stsh(&stsh)
            .warnings
            .first()
            .cloned()
            .expect("bad cbStshi");
        assert!(w.contains("cbStshi 4"), "{w}");
    }

    /// A sheet cut mid-entry keeps the styles read before the cut and says how
    /// many are missing.
    #[test]
    fn test_a_truncated_style_sheet_keeps_what_it_read_and_warns() {
        let mut stsh = synthetic_stsh();
        stsh.truncate(stsh.len() - 20); // lose the tail entries
        let sheet = parse_stsh(&stsh);
        assert!(!sheet.styles.is_empty() && sheet.styles.len() < 16);
        assert_eq!(sheet.styles[1].name, "Heading 1");
        let w = sheet
            .warnings
            .first()
            .cloned()
            .expect("a cut sheet must warn");
        assert!(w.contains("declares 16 styles"), "{w}");
    }

    #[test]
    fn test_an_absent_or_out_of_bounds_style_sheet_warns() {
        let table_stream = vec![0u8; 64];
        let w = parse_style_sheet(&table_stream, &fib_with_stsh(0, 0))
            .warnings
            .first()
            .cloned()
            .expect("lcbStshf 0");
        assert!(w.contains("absent"), "{w}");
        let w = parse_style_sheet(&table_stream, &fib_with_stsh(1000, 50))
            .warnings
            .first()
            .cloned()
            .expect("out of bounds");
        assert!(w.contains("outside the Table stream"), "{w}");
    }

    /// Build an STSH from `LPStd` entries with `cbSTDBaseInFile` = 0x0A.
    fn stsh_of(entries: &[Vec<u8>]) -> Vec<u8> {
        let mut d = 18u16.to_le_bytes().to_vec(); // cbStshi
        d.extend_from_slice(&(entries.len() as u16).to_le_bytes()); // cstd
        d.extend_from_slice(&0x000Au16.to_le_bytes()); // cbSTDBaseInFile
        d.extend_from_slice(&[0u8; 14]);
        for e in entries {
            d.extend_from_slice(e);
        }
        d
    }

    #[test]
    fn test_parse_stsh_reads_istd_base_and_own_outline_level() {
        let sheet = parse_stsh(&stsh_of(&[
            lpstd(0, "Normal"),
            lpstd_styled(1, "Heading 1", Some(0), Some(0)),
            lpstd_styled(0x0FFE, "Custom", Some(1), Some(3)),
            lpstd_styled(0x0FFE, "Plain Body", Some(0), Some(9)),
            lpstd_styled(0x0FFE, "No Outline", Some(1), None),
        ]));
        assert!(sheet.warnings.is_empty(), "{:?}", sheet.warnings);
        let st = &sheet.styles;
        assert_eq!((st[0].base, st[0].own_outline), (None, None), "0x0FFF means no base");
        assert_eq!((st[1].base, st[1].own_outline), (Some(0), Some(0)));
        assert_eq!((st[2].base, st[2].own_outline), (Some(1), Some(3)));
        assert_eq!((st[3].base, st[3].own_outline), (Some(0), Some(9)), "9 = body text");
        assert_eq!((st[4].base, st[4].own_outline), (Some(1), None));
        assert_eq!(st[2].name, "Custom", "the name is still read at the right offset");
    }

    /// A style's own outline level wins over its name and over what it would
    /// inherit; level 9 ends the walk as explicit body text.
    #[test]
    fn test_heading_level_prefers_the_styles_own_outline_then_name_then_base() {
        let sheet = parse_stsh(&stsh_of(&[
            lpstd(0, "Normal"),                                    // 0
            lpstd_styled(1, "Heading 1", Some(0), None),           // 1: by sti
            lpstd_styled(0x0FFE, "H4", Some(0), Some(3)),          // 2: own outline 3 -> level 4
            lpstd_styled(0x0FFE, "Child", Some(2), None),          // 3: inherits H4
            lpstd_styled(0x0FFE, "Grandchild", Some(3), None),     // 4: inherits through two
            lpstd_styled(0x0FFE, "Heading 5", Some(1), Some(1)),   // 5: own outline beats its name
            lpstd_styled(0x0FFE, "Not Heading", Some(1), Some(9)), // 6: body text beats base Heading 1
            lpstd_styled(0x0FFE, "Under Body", Some(6), None),     // 7: inherits explicit body text
            lpstd_styled(0x0FFE, "Plain", Some(0), None),          // 8: nothing anywhere
        ]));
        let lvl = |i| heading_level_for_istd(&sheet.styles, i);
        assert_eq!(lvl(1), Some(1));
        assert_eq!(lvl(2), Some(4));
        assert_eq!(lvl(3), Some(4), "one step of inheritance");
        assert_eq!(lvl(4), Some(4), "two steps of inheritance");
        assert_eq!(lvl(5), Some(2), "the style's own outline level beats its name");
        assert_eq!(lvl(6), None, "explicit body text beats an inherited heading");
        assert_eq!(lvl(7), None, "and is itself inherited");
        assert_eq!(lvl(8), None);
        assert_eq!(lvl(99), None, "out of range istd");
    }

    /// A cycle, a self-base and a chain deeper than the bound must all end, give
    /// no level, and be reported once as a warning.
    #[test]
    fn test_cyclic_and_over_deep_base_chains_end_and_are_reported() {
        let mut entries = vec![
            lpstd(0, "Normal"),                        // 0
            lpstd_styled(0x0FFE, "A", Some(2), None),  // 1 -> 2
            lpstd_styled(0x0FFE, "B", Some(1), None),  // 2 -> 1 (cycle)
            lpstd_styled(0x0FFE, "Me", Some(3), None), // 3 -> 3 (self)
        ];
        // 40 styles deep, the last one a real Heading 1: the bound is 32.
        let first_deep = entries.len() as u16;
        for i in 0..40u16 {
            let base = if i == 0 { 0 } else { first_deep + i - 1 };
            entries.push(lpstd_styled(0x0FFE, "Deep", Some(base), None));
        }
        entries.push(lpstd_styled(1, "Heading 1", Some(first_deep + 39), None));
        let sheet = parse_stsh(&stsh_of(&entries));
        let lvl = |i| heading_level_for_istd(&sheet.styles, i);
        assert_eq!(lvl(1), None);
        assert_eq!(lvl(2), None);
        assert_eq!(lvl(3), None);
        assert_eq!(lvl(first_deep + 40), Some(1), "a Heading 1 by sti needs no base");
        assert_eq!(lvl(first_deep + 39), None, "40 steps from any heading is past the bound");
        let w = sheet.warnings.join(" | ");
        assert!(w.contains("istdBase chain is cyclic or deeper than 32"), "{w}");
        assert_eq!(sheet.warnings.len(), 1, "reported once, not per style: {w}");
    }

    /// Frame `std` as an `LPStd` (`cbStd` + `STD` + the even-offset pad).
    fn frame_std(std: &[u8]) -> Vec<u8> {
        let mut e = (std.len() as u16).to_le_bytes().to_vec();
        e.extend_from_slice(std);
        if std.len() % 2 == 1 {
            e.push(0);
        }
        e
    }

    /// The `STD` bytes inside an `LPStd` built by [`lpstd_styled`].
    fn std_of(lpstd: &[u8]) -> Vec<u8> {
        let cb = usize::from(u16::from_le_bytes([lpstd[0], lpstd[1]]));
        lpstd[2..2 + cb].to_vec()
    }

    /// A style whose `UpxPapx` is cut short or lies about its size yields no
    /// outline level and no panic; the name is still read.
    #[test]
    fn test_malformed_style_upx_degrades_to_no_outline_level() {
        let std = std_of(&lpstd_styled(0x0FFE, "Cut Short", Some(0), Some(2)));
        // The STD ends with `cbUpx`(2) + `istd`(2) + grpprl(3). Cut into those 7
        // bytes one at a time, re-framing so only the UPX is short.
        for cut in 1..=7usize {
            let e = frame_std(&std[..std.len() - cut]);
            let sheet = parse_stsh(&stsh_of(&[lpstd(0, "Normal"), e]));
            assert_eq!(sheet.styles[1].name, "Cut Short", "cut {cut}");
            assert_eq!(sheet.styles[1].own_outline, None, "cut {cut}");
            assert!(
                sheet.warnings.is_empty(),
                "a short UPX is not a sheet error: {:?}",
                sheet.warnings
            );
        }
        // `cbUpx` claiming far more than the STD holds.
        let mut lying = std.clone();
        let at = lying.len() - 7;
        lying[at..at + 2].copy_from_slice(&0xFFFFu16.to_le_bytes());
        let sheet = parse_stsh(&stsh_of(&[lpstd(0, "Normal"), frame_std(&lying)]));
        assert_eq!(sheet.styles[1].own_outline, None);
        // A non-paragraph style (`sgc` 2 = character) never reads a PAPX.
        let mut character = std.clone();
        character[2] = (character[2] & 0xF0) | 0x02;
        let sheet = parse_stsh(&stsh_of(&[lpstd(0, "Normal"), frame_std(&character)]));
        assert_eq!(sheet.styles[1].own_outline, None);
        // `cupx` = 0: the style has no UPX at all, whatever bytes follow.
        let mut no_upx = std.clone();
        no_upx[4] &= 0xF0;
        let sheet = parse_stsh(&stsh_of(&[lpstd(0, "Normal"), frame_std(&no_upx)]));
        assert_eq!(sheet.styles[1].own_outline, None);
    }

    /// An operand above 9 is not an outline level: ignored, so the style falls
    /// through to its name and base instead of inventing a level.
    #[test]
    fn test_an_out_of_range_style_outline_operand_is_ignored() {
        let sheet = parse_stsh(&stsh_of(&[
            lpstd(0, "Normal"),
            lpstd_styled(0x0FFE, "Heading 2", Some(0), Some(200)),
        ]));
        assert_eq!(sheet.styles[1].own_outline, None);
        assert_eq!(heading_level_for_istd(&sheet.styles, 1), Some(2));
    }

    /// The base-chain bound is part of the behaviour: a heading reached in 31
    /// hops resolves and is not warned about; one 32 hops away does neither.
    #[test]
    fn test_the_base_chain_bound_is_pinned() {
        let chain = |hops: u16| {
            // 0 Normal, 1 Heading 1, then `hops` styles, each based on the last;
            // the paragraph's style is the final one, `hops` hops from Heading 1.
            let mut entries = vec![
                lpstd(0, "Normal"),
                lpstd_styled(1, "Heading 1", Some(0), None),
            ];
            for i in 0..hops {
                entries.push(lpstd_styled(0x0FFE, "Link", Some(1 + i), None));
            }
            let sheet = parse_stsh(&stsh_of(&entries));
            let level = heading_level_for_istd(&sheet.styles, hops + 1);
            (level, sheet.warnings.len())
        };
        assert_eq!(chain(1), (Some(1), 0));
        assert_eq!(chain(30), (Some(1), 0));
        assert_eq!(chain(31), (Some(1), 0), "the longest chain that still resolves");
        assert_eq!(
            chain(32),
            (None, 1),
            "one hop past the bound neither resolves nor goes unreported"
        );
    }

    /// Deterministic mutation fuzzing of the whole STSH path: thousands of
    /// byte flips, truncations, extensions and size-field overwrites of a rich
    /// sheet (bases, own outline levels, an odd-length entry) must never panic
    /// or hang, never yield more styles than the 4096 cap, and every level that
    /// does resolve must be a real one (1-9). The same property the
    /// `fuzz_doc_stsh` target explores open-endedly.
    #[test]
    fn test_mutated_style_sheets_never_panic_and_resolve_only_real_levels() {
        let seed_sheet = stsh_of(&[
            lpstd(0, "Normal"),
            lpstd_styled(1, "Heading 1", Some(0), None),
            lpstd_styled(0x0FFE, "H4", Some(0), Some(3)),
            lpstd_styled(0x0FFE, "Child", Some(2), None),
            lpstd_styled(0x0FFE, "Body", Some(1), Some(9)),
            lpstd_styled(0x0FFE, "Odd Len", Some(0), Some(1)),
            empty_lpstd(),
            lpstd_styled(0x0FFE, "Loop", Some(8), None),
            lpstd_styled(0x0FFE, "Loop 2", Some(7), None),
        ]);
        // xorshift64: deterministic, no dependency.
        let mut state = 0x9E37_79B9_7F4A_7C15u64;
        let mut next = move || {
            state ^= state << 13;
            state ^= state >> 7;
            state ^= state << 17;
            state
        };
        for round in 0..6000u32 {
            let mut d = seed_sheet.clone();
            match round % 4 {
                0 => {
                    for _ in 0..=(next() % 8) {
                        let at = (next() as usize) % d.len();
                        d[at] = next() as u8;
                    }
                },
                1 => d.truncate((next() as usize) % (d.len() + 1)),
                2 => {
                    d.extend((0..(next() % 64)).map(|_| next() as u8));
                    let at =
                        (next() as usize) % d.len().max(2).min(d.len().saturating_sub(1)).max(1);
                    d[at] = next() as u8;
                },
                _ => {
                    // Overwrite a u16 anywhere: sizes, counts and bases alike.
                    let at = (next() as usize) % d.len().saturating_sub(1).max(1);
                    d[at..at + 2].copy_from_slice(&(next() as u16).to_le_bytes());
                },
            }
            let sheet = parse_stsh(&d);
            assert!(sheet.styles.len() <= 4096, "round {round}: {} styles", sheet.styles.len());
            for istd in 0..sheet.styles.len().min(64) as u16 {
                if let Some(level) = heading_level_for_istd(&sheet.styles, istd) {
                    assert!(
                        (1..=MAX_OUTLINE_LEVEL).contains(&level),
                        "round {round}: level {level}"
                    );
                }
            }
        }
    }

    #[test]
    fn test_style_name_for_istd_default_alias_empty_and_out_of_range() {
        let sheet = parse_stsh(&stsh_bytes(&[
            (0, 1, "Normal"),
            (0x0FFE, 1, "Caption,Legende"),
            (0x0FFE, 1, ""),
            (0x0FFE, 1, "  spaced  ,x"),
        ]));
        let styles = &sheet.styles;
        assert_eq!(style_name_for_istd(styles, 0), None, "istd 0 is the default style");
        assert_eq!(style_name_for_istd(styles, 1).as_deref(), Some("Caption"));
        assert_eq!(style_name_for_istd(styles, 2), None, "an unnamed slot");
        assert_eq!(style_name_for_istd(styles, 3).as_deref(), Some("spaced"));
        assert_eq!(style_name_for_istd(styles, 4), None, "past the end of the sheet");
        assert_eq!(style_name_for_istd(styles, u16::MAX), None);
        assert_eq!(style_name_for_istd(&[], 5), None);
    }

    #[test]
    fn test_style_name_for_istd_built_ins_are_canonical_and_the_rest_keep_the_stored_name() {
        let sheet = parse_stsh(&stsh_bytes(&[
            (0, 1, "Standard"),
            (1, 1, "Überschrift 1"),
            (9, 1, "Überschrift 9"),
            (31, 1, "Kopfzeile"),
            (32, 1, "Fußzeile"),
            (34, 1, "Beschriftung"),
            (30, 1, "Commentaire"),
            (0x0FFE, 1, "section-e"),
        ]));
        let n = |i| style_name_for_istd(&sheet.styles, i);
        assert_eq!(n(0), None);
        assert_eq!(n(1).as_deref(), Some("Heading 1"));
        assert_eq!(n(2).as_deref(), Some("Heading 9"));
        assert_eq!(n(3).as_deref(), Some("Header"));
        assert_eq!(n(4).as_deref(), Some("Footer"));
        assert_eq!(n(5).as_deref(), Some("Caption"));
        assert_eq!(n(6).as_deref(), Some("Commentaire"), "not in the canonical set");
        assert_eq!(n(7).as_deref(), Some("section-e"));
    }

    #[test]
    fn test_style_name_for_istd_rejects_unusable_names() {
        let long = "N".repeat(256);
        let exact = "M".repeat(255);
        let sheet = parse_stsh(&stsh_bytes(&[
            (0, 1, "Normal"),
            (0x0FFE, 1, &long),
            (0x0FFE, 1, "a\u{0}b"),
            (0x0FFE, 1, &exact),
        ]));
        let n = |i| style_name_for_istd(&sheet.styles, i);
        assert_eq!(n(1), None, "over 255 characters");
        assert_eq!(n(2), None, "control character");
        assert_eq!(n(3).map(|n| n.len()), Some(255));
    }

    /// A style sheet of `entries` (`sti`, `stk`, name) in `istd` order, written
    /// as the bytes `parse_stsh` reads ([MS-DOC] §2.9.274 `Stshif`, §2.9.135
    /// `LPStd`, §2.9.253 `StdfBase`).
    fn stsh_bytes(entries: &[(u16, u16, &str)]) -> Vec<u8> {
        let mut d = 18u16.to_le_bytes().to_vec(); // cbStshi
        d.extend_from_slice(&(entries.len() as u16).to_le_bytes()); // cstd
        d.extend_from_slice(&0x000Au16.to_le_bytes()); // cbSTDBaseInFile
        d.extend_from_slice(&[0u8; 14]);
        for &(sti, stk, name) in entries {
            let mut std = vec![0u8; 10];
            std[0..2].copy_from_slice(&sti.to_le_bytes());
            std[2..4].copy_from_slice(&(stk | (0x0FFF << 4)).to_le_bytes());
            let units: Vec<u16> = name.encode_utf16().collect();
            std.extend_from_slice(&(units.len() as u16).to_le_bytes());
            for u in &units {
                std.extend_from_slice(&u.to_le_bytes());
            }
            std.extend_from_slice(&0u16.to_le_bytes());
            d.extend_from_slice(&(std.len() as u16).to_le_bytes());
            d.extend_from_slice(&std);
        }
        d
    }

    #[test]
    fn test_style_name_is_examined_once_per_style_not_once_per_paragraph() {
        use crate::limits::STYLE_NAME_BYTES_SCANNED as SCANNED;
        // A name of 60 KB (the 16-bit `cbStd` allows little more): looked up for every
        // paragraph that uses the style.
        let long = "N".repeat(30_000);
        let sheet = parse_stsh(&stsh_bytes(&[
            (0, 1, "Normal"),
            (0x0FFE, 1, &long),
            (0x0FFE, 1, "section-e"),
        ]));
        let after_parse = SCANNED.with(std::cell::Cell::get);
        for _ in 0..2_000 {
            assert_eq!(style_name_for_istd(&sheet.styles, 1), None, "over 255 characters");
            assert_eq!(style_name_for_istd(&sheet.styles, 2).as_deref(), Some("section-e"));
        }
        assert_eq!(
            SCANNED.with(std::cell::Cell::get),
            after_parse,
            "looking a style name up must not re-examine the stored name"
        );
    }

    #[test]
    fn test_unusable_style_name_is_reported_once_per_style_sheet() {
        let long = "N".repeat(300);
        let sheet = parse_stsh(&stsh_bytes(&[
            (0, 1, "Normal"),
            (0x0FFE, 1, &long),
            (0x0FFE, 1, "a\u{1}b"),
        ]));
        let hits = sheet
            .warnings
            .iter()
            .filter(|w| w.contains("style name"))
            .count();
        assert_eq!(hits, 1, "{:?}", sheet.warnings);
        assert!(sheet.warnings.iter().any(|w| w.starts_with('2')), "{:?}", sheet.warnings);
    }

    #[test]
    fn test_style_name_is_none_unless_the_style_is_a_paragraph_style() {
        // [MS-DOC] §2.9.253 `StdfBase.stk`: 1 paragraph, 2 character, 3 table,
        // 4 numbering. A paragraph's `istd` can name any kind of style; only a
        // paragraph style is the paragraph's style.
        let sheet = parse_stsh(&stsh_bytes(&[
            (0, 1, "Normal"),
            (0x0FFE, 1, "section-e"),
            (0x0FFE, 2, "Emphasis"),
            (0x0FFE, 3, "Table Grid"),
            (0x0FFE, 4, "List Bullet Style"),
            (34, 2, "Beschriftung"),
        ]));
        let n = |i| style_name_for_istd(&sheet.styles, i);
        assert_eq!(n(1).as_deref(), Some("section-e"));
        assert_eq!(n(2), None, "character style");
        assert_eq!(n(3), None, "table style");
        assert_eq!(n(4), None, "numbering style");
        assert_eq!(n(5), None, "a character style is not Caption whatever its sti");
    }

    /// A character style (`stk` 2) with one `UpxChpx` holding `sprmCRgFtc2` = 5.
    fn char_style_std() -> Vec<u8> {
        let mut std = vec![0u8; 10];
        std[2..4].copy_from_slice(&(2u16 | (0x0FFFu16 << 4)).to_le_bytes());
        std[4..6].copy_from_slice(&1u16.to_le_bytes()); // cupx
        std.extend_from_slice(&1u16.to_le_bytes()); // cch
        std.extend_from_slice(&u16::from(b'X').to_le_bytes());
        std.extend_from_slice(&0u16.to_le_bytes()); // chTerm
        std.extend_from_slice(&4u16.to_le_bytes()); // cbUpx
        std.extend_from_slice(&[0x51, 0x4A, 5, 0]); // sprmCRgFtc2
        std
    }

    #[test]
    fn test_a_character_style_font_is_read_and_a_cut_style_never_panics() {
        let std = char_style_std();
        assert_eq!(style_own_ftc_other(&std, 10), Some(5));
        for cut in 0..std.len() {
            assert_eq!(style_own_ftc_other(&std[..cut], 10), None, "cut at {cut}");
        }
        // A `cbUpx` that runs past the style, and a name length that does.
        let mut long = std.clone();
        long[16..18].copy_from_slice(&0xFFF0u16.to_le_bytes());
        assert_eq!(style_own_ftc_other(&long, 10), None);
        let mut wide = std;
        wide[10..12].copy_from_slice(&0xFFFFu16.to_le_bytes());
        assert_eq!(style_own_ftc_other(&wide, 10), None);
    }
}
