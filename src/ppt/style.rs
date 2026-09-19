//! `.ppt` direct character- and paragraph-level text formatting:
//! `StyleTextPropAtom` (`0x0FA1`) — declared as a constant since the
//! crate's earliest days but never matched anywhere until now. Every
//! "bold title" or "aligned paragraph" previously visible in the IR was
//! synthetic, keyed off `TextType` classification alone, never actually
//! read from the file.
//!
//! Byte layouts verified against the published [MS-PPT] `StyleTextPropAtom`,
//! `TextPFRun`/`TextCFRun`, `TextPFException`/`TextCFException`, and
//! `PFMasks`/`CFMasks` spec pages, including the "Paragraph Formatting"
//! section's own worked byte example: a run's `count` field there is
//! `0x2A` (42) against 41 bytes of actual text, "because of the
//! terminating line break" — every text body has an implicit trailing
//! paragraph mark that the run counts here include but that is *not*
//! present in the decoded text. Callers pass `text_char_len + 1` in and
//! this module clamps output ranges back down to `text_char_len`.
//!
//! Scope, deliberately not covered here:
//! - `ColorIndexStruct.index` values other than `0xFE` (explicit RGB) are
//!   color-scheme indices (background/text/title/accent-N) that require
//!   resolving the slide's `SlideSchemeColorSchemeAtom`; that resolution
//!   doesn't exist anywhere in the crate yet, so scheme-indexed colors are
//!   left unset here rather than guessed at.
//! - `tabStops` and `wrapFlags` are parsed only far enough to skip their
//!   bytes correctly (required to keep every later field in the record
//!   aligned) — their values aren't surfaced; the IR has no field for
//!   per-tab-stop layout.
//! - Master-slide style inheritance (formatting a placeholder doesn't
//!   directly override, but should visually inherit from
//!   `TextMasterStyleAtom`) is a separate, larger gap, filed as its own
//!   issue rather than folded into this one.

/// Character-level formatting resolved from a single `TextCFRun`. Each
/// field is `None` when the corresponding `CFMasks` bit was unset, i.e.
/// this run doesn't specify (and doesn't override) that property.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CharFormat {
    /// `None` = not specified by this run's `CFMasks.bold` bit.
    pub bold: Option<bool>,
    /// `None` = not specified by this run's `CFMasks.italic` bit.
    pub italic: Option<bool>,
    /// `None` = not specified by this run's `CFMasks.underline` bit.
    pub underline: Option<bool>,
    /// Font size in points ([MS-PPT]: 1..=4000).
    pub font_size: Option<i16>,
    /// Explicit sRGB color (`ColorIndexStruct.index == 0xFE` only).
    pub color: Option<[u8; 3]>,
    /// Baseline position, percent of line height (-100..=100); positive
    /// is superscript-like, negative subscript-like.
    pub position: Option<i16>,
}

/// Paragraph-level formatting resolved from a single `TextPFRun`.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ParaFormat {
    /// Raw `TextAlignmentEnum` value (0=left, 1=center, 2=right,
    /// 3=justify, 4=distributed, 5=Thai distributed, 6=justify-low).
    pub alignment: Option<u16>,
}

/// A formatting run resolved to a clamped `[start, end)` character range
/// over the corresponding text's `chars()`.
#[derive(Debug, Clone, PartialEq)]
pub struct CharFormatSpan {
    /// Start character offset (inclusive), into the corresponding text.
    pub start: usize,
    /// End character offset (exclusive), into the corresponding text.
    pub end: usize,
    /// The formatting that applies to `[start, end)`.
    pub format: CharFormat,
}

/// A paragraph-formatting run resolved to a clamped `[start, end)`
/// character range over the corresponding text's `chars()`.
#[derive(Debug, Clone, PartialEq)]
pub struct ParaFormatSpan {
    /// Start character offset (inclusive), into the corresponding text.
    pub start: usize,
    /// End character offset (exclusive), into the corresponding text.
    pub end: usize,
    /// The formatting that applies to `[start, end)`.
    pub format: ParaFormat,
}

struct Cursor<'a> {
    data: &'a [u8],
    pos: usize,
}

impl<'a> Cursor<'a> {
    fn new(data: &'a [u8]) -> Self {
        Self { data, pos: 0 }
    }

    fn remaining(&self) -> usize {
        self.data.len().saturating_sub(self.pos)
    }

    fn u16(&mut self) -> Option<u16> {
        if self.remaining() < 2 {
            return None;
        }
        let v = u16::from_le_bytes([self.data[self.pos], self.data[self.pos + 1]]);
        self.pos += 2;
        Some(v)
    }

    fn i16(&mut self) -> Option<i16> {
        self.u16().map(|v| v as i16)
    }

    fn u32(&mut self) -> Option<u32> {
        if self.remaining() < 4 {
            return None;
        }
        let v = u32::from_le_bytes([
            self.data[self.pos],
            self.data[self.pos + 1],
            self.data[self.pos + 2],
            self.data[self.pos + 3],
        ]);
        self.pos += 4;
        Some(v)
    }

    fn skip(&mut self, n: usize) -> Option<()> {
        if self.remaining() < n {
            return None;
        }
        self.pos += n;
        Some(())
    }

    /// `ColorIndexStruct` (4 bytes): red, green, blue, index. Only
    /// `index == 0xFE` (explicit sRGB) resolves to a color here.
    fn color(&mut self) -> Option<Option<[u8; 3]>> {
        if self.remaining() < 4 {
            return None;
        }
        let r = self.data[self.pos];
        let g = self.data[self.pos + 1];
        let b = self.data[self.pos + 2];
        let idx = self.data[self.pos + 3];
        self.pos += 4;
        Some(if idx == 0xFE { Some([r, g, b]) } else { None })
    }
}

// PFMasks bit positions ([MS-PPT] 2.9.32).
const PF_HAS_BULLET: u32 = 1 << 0;
const PF_BULLET_HAS_FONT: u32 = 1 << 1;
const PF_BULLET_HAS_COLOR: u32 = 1 << 2;
const PF_BULLET_HAS_SIZE: u32 = 1 << 3;
const PF_BULLET_FONT: u32 = 1 << 4;
const PF_BULLET_COLOR: u32 = 1 << 5;
const PF_BULLET_SIZE: u32 = 1 << 6;
const PF_BULLET_CHAR: u32 = 1 << 7;
const PF_LEFT_MARGIN: u32 = 1 << 8;
const PF_INDENT: u32 = 1 << 10;
const PF_ALIGN: u32 = 1 << 11;
const PF_LINE_SPACING: u32 = 1 << 12;
const PF_SPACE_BEFORE: u32 = 1 << 13;
const PF_SPACE_AFTER: u32 = 1 << 14;
const PF_DEFAULT_TAB_SIZE: u32 = 1 << 15;
const PF_FONT_ALIGN: u32 = 1 << 16;
const PF_CHAR_WRAP: u32 = 1 << 17;
const PF_WORD_WRAP: u32 = 1 << 18;
const PF_OVERFLOW: u32 = 1 << 19;
const PF_TAB_STOPS: u32 = 1 << 20;
const PF_TEXT_DIRECTION: u32 = 1 << 21;

/// Parse one `TextPFRun`: `count: u32`, `indentLevel: u16`, then a
/// `TextPFException` (`masks: u32` followed by whichever optional fields
/// the mask bits select, in the exact order [MS-PPT] lays them out).
fn parse_pf_run(c: &mut Cursor) -> Option<(usize, ParaFormat)> {
    let count = c.u32()? as usize;
    c.skip(2)?; // indentLevel
    let masks = c.u32()?;
    Some((count, parse_pf_body(c, masks)?))
}

/// The `TextPFException` body shared by `TextPFRun` and each
/// level of a `TxMasterStyleAtom` — same `masks: u32` +
/// mask-selected fields, just with a different prefix before this point
/// (`count: u32` + `indentLevel: u16` for a run; a conditional 2-byte
/// `indentLevel` for a master style level, see
/// [`parse_master_style_level`]).
fn parse_pf_body(c: &mut Cursor, masks: u32) -> Option<ParaFormat> {
    if masks & (PF_HAS_BULLET | PF_BULLET_HAS_FONT | PF_BULLET_HAS_COLOR | PF_BULLET_HAS_SIZE) != 0 {
        c.skip(2)?; // bulletFlags
    }
    if masks & PF_BULLET_CHAR != 0 {
        c.skip(2)?; // bulletChar
    }
    if masks & PF_BULLET_FONT != 0 {
        c.skip(2)?; // bulletFontRef
    }
    if masks & PF_BULLET_SIZE != 0 {
        c.skip(2)?; // bulletSize
    }
    if masks & PF_BULLET_COLOR != 0 {
        c.skip(4)?; // bulletColor (ColorIndexStruct)
    }
    let alignment = if masks & PF_ALIGN != 0 { c.u16() } else { None };
    if masks & PF_LINE_SPACING != 0 {
        c.skip(2)?;
    }
    if masks & PF_SPACE_BEFORE != 0 {
        c.skip(2)?;
    }
    if masks & PF_SPACE_AFTER != 0 {
        c.skip(2)?;
    }
    if masks & PF_LEFT_MARGIN != 0 {
        c.skip(2)?;
    }
    if masks & PF_INDENT != 0 {
        c.skip(2)?;
    }
    if masks & PF_DEFAULT_TAB_SIZE != 0 {
        c.skip(2)?;
    }
    if masks & PF_TAB_STOPS != 0 {
        let n = c.u16()? as usize;
        c.skip(n * 4)?; // rgTabStop: count * 4 bytes
    }
    if masks & PF_FONT_ALIGN != 0 {
        c.skip(2)?;
    }
    if masks & (PF_CHAR_WRAP | PF_WORD_WRAP | PF_OVERFLOW) != 0 {
        c.skip(2)?; // wrapFlags
    }
    if masks & PF_TEXT_DIRECTION != 0 {
        c.skip(2)?;
    }

    Some(ParaFormat { alignment })
}

// CFMasks bit positions ([MS-PPT] 2.9.13).
const CF_BOLD: u32 = 1 << 0;
const CF_ITALIC: u32 = 1 << 1;
const CF_UNDERLINE: u32 = 1 << 2;
const CF_FONT_STYLE_ANY: u32 = 0x3FFF; // bits 0-13: bold..emboss plus fHasStyle
const CF_TYPEFACE: u32 = 1 << 16;
const CF_SIZE: u32 = 1 << 17;
const CF_COLOR: u32 = 1 << 18;
const CF_POSITION: u32 = 1 << 19;
const CF_OLD_EA_TYPEFACE: u32 = 1 << 21;
const CF_ANSI_TYPEFACE: u32 = 1 << 22;
const CF_SYMBOL_TYPEFACE: u32 = 1 << 23;

/// Parse one `TextCFRun`: `count: u32`, then a `TextCFException`
/// (`masks: u32` followed by whichever optional fields the mask bits
/// select, in the exact order [MS-PPT] lays them out).
fn parse_cf_run(c: &mut Cursor) -> Option<(usize, CharFormat)> {
    let count = c.u32()? as usize;
    let masks = c.u32()?;
    Some((count, parse_cf_body(c, masks)?))
}

/// The `TextCFException` body shared by `TextCFRun` and each
/// level of a `TxMasterStyleAtom` — same `masks: u32` +
/// mask-selected fields, just with a different prefix (a `count: u32` for
/// a run; nothing for a master style level, see
/// [`parse_master_style_level`]).
fn parse_cf_body(c: &mut Cursor, masks: u32) -> Option<CharFormat> {
    let mut fmt = CharFormat::default();

    if masks & CF_FONT_STYLE_ANY != 0 {
        let style = c.u16()?;
        if masks & CF_BOLD != 0 {
            fmt.bold = Some(style & 0x0001 != 0);
        }
        if masks & CF_ITALIC != 0 {
            fmt.italic = Some(style & 0x0002 != 0);
        }
        if masks & CF_UNDERLINE != 0 {
            fmt.underline = Some(style & 0x0004 != 0);
        }
    }
    if masks & CF_TYPEFACE != 0 {
        c.skip(2)?; // fontRef
    }
    if masks & CF_OLD_EA_TYPEFACE != 0 {
        c.skip(2)?; // oldEAFontRef
    }
    if masks & CF_ANSI_TYPEFACE != 0 {
        c.skip(2)?; // ansiFontRef
    }
    if masks & CF_SYMBOL_TYPEFACE != 0 {
        c.skip(2)?; // symbolFontRef
    }
    if masks & CF_SIZE != 0 {
        fmt.font_size = c.i16();
    }
    if masks & CF_COLOR != 0 {
        fmt.color = c.color()?;
    }
    if masks & CF_POSITION != 0 {
        fmt.position = c.i16();
    }

    Some(fmt)
}

/// Parse one indent level's `(ParaFormat, CharFormat)` pair from a
/// `TxMasterStyleAtom` body cursor, positioned right after the level
/// count (or the previous level's character style).
///
/// `has_indent_level_field`: per [MS-PPT] and cross-checked against
/// Apache POI's `TxMasterStyleAtom#init()`, only text types whose
/// numeric `TextTypeEnum` value is `>= 5` (`CenterBody`/`CenterTitle`/
/// `HalfBody`/`QuarterBody` — placeholder-layout variants, not the
/// ordinary `Title`(0)/`Body`(1)/`Notes`(2)/`Other`(4) this crate
/// resolves master inheritance for) carry an explicit 2-byte
/// `indentLevel` before each level's paragraph mask; for the rest, a
/// level's position in the array *is* its indent level, with no
/// separate field to skip.
fn parse_master_style_level(c: &mut Cursor, has_indent_level_field: bool) -> Option<(ParaFormat, CharFormat)> {
    if has_indent_level_field {
        c.skip(2)?; // indentLevel
    }
    let pf_masks = c.u32()?;
    let pf = parse_pf_body(c, pf_masks)?;
    let cf_masks = c.u32()?;
    let cf = parse_cf_body(c, cf_masks)?;
    Some((pf, cf))
}

/// Parse a `TxMasterStyleAtom` body (`rec.data`, header already stripped)
/// into its per-level `(ParaFormat, CharFormat)` pairs, in indent-level
/// order (index 0 = no indentation). Up to 5 levels per [MS-PPT]/POI's
/// own `TxMasterStyleAtom.MAX_INDENT`. `text_type_native_id` is the
/// record's own `recInstance` — "the atom instance value is the text
/// type" (POI's own doc comment on this record), encoded exactly like
/// `TextHeaderAtom`'s `txType`.
///
/// Stops (returning whatever levels parsed cleanly so far) on any
/// malformed/truncated level rather than propagating an error — master
/// style inheritance is a best-effort enhancement to direct formatting,
/// never a hard requirement for reading the rest of the document.
pub fn parse_tx_master_style_atom(data: &[u8], text_type_native_id: u16) -> Vec<(ParaFormat, CharFormat)> {
    let mut c = Cursor::new(data);
    let Some(levels) = c.u16() else { return Vec::new() };
    let has_indent_level_field = text_type_native_id >= 5;
    let mut out = Vec::with_capacity((levels as usize).min(5));
    for _ in 0..levels.min(5) {
        match parse_master_style_level(&mut c, has_indent_level_field) {
            Some(pair) => out.push(pair),
            None => break,
        }
    }
    out
}

impl CharFormat {
    /// Fill any field this format left unset (`None`) from `master`,
    /// keeping every field this format *did* specify untouched — the
    /// "only fill in what's missing" inheritance [MS-PPT] describes for
    /// a placeholder shape falling back to its master's
    /// `TextMasterStyleAtom`.
    pub fn inherit_from(&self, master: &CharFormat) -> CharFormat {
        CharFormat {
            bold: self.bold.or(master.bold),
            italic: self.italic.or(master.italic),
            underline: self.underline.or(master.underline),
            font_size: self.font_size.or(master.font_size),
            color: self.color.or(master.color),
            position: self.position.or(master.position),
        }
    }
}

impl ParaFormat {
    /// As [`CharFormat::inherit_from`], for paragraph-level formatting.
    pub fn inherit_from(&self, master: &ParaFormat) -> ParaFormat {
        ParaFormat { alignment: self.alignment.or(master.alignment) }
    }
}

/// Parse a `StyleTextPropAtom` body (`rec.data`, i.e. with the 8-byte
/// record header already stripped) into clamped character- and
/// paragraph-formatting spans over a text of `text_char_len` characters.
///
/// `rgTextPFRun` entries are read first, until their running `count` total
/// reaches `text_char_len + 1` (the implicit trailing paragraph mark — see
/// the module doc), then `rgTextCFRun` entries the same way. A run whose
/// `count` would overshoot the real text is clamped to `text_char_len`
/// rather than dropped, so direct formatting on the last real character of
/// a run is never lost to the phantom trailing mark.
pub fn parse_style_text_prop(data: &[u8], text_char_len: usize) -> (Vec<ParaFormatSpan>, Vec<CharFormatSpan>) {
    let target = text_char_len + 1;
    let mut c = Cursor::new(data);

    let mut para_spans = Vec::new();
    let mut covered = 0usize;
    while covered < target {
        let Some((count, fmt)) = parse_pf_run(&mut c) else { break };
        let start = covered.min(text_char_len);
        let end = covered.saturating_add(count).min(text_char_len);
        if end > start {
            para_spans.push(ParaFormatSpan { start, end, format: fmt });
        }
        covered = covered.saturating_add(count);
        if count == 0 {
            break; // guard against a zero-length run stalling the loop forever
        }
    }

    let mut char_spans = Vec::new();
    covered = 0;
    while covered < target {
        let Some((count, fmt)) = parse_cf_run(&mut c) else { break };
        let start = covered.min(text_char_len);
        let end = covered.saturating_add(count).min(text_char_len);
        if end > start {
            char_spans.push(CharFormatSpan { start, end, format: fmt });
        }
        covered = covered.saturating_add(count);
        if count == 0 {
            break;
        }
    }

    (para_spans, char_spans)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn le32(v: u32) -> [u8; 4] {
        v.to_le_bytes()
    }
    fn le16(v: u16) -> [u8; 2] {
        v.to_le_bytes()
    }

    #[test]
    fn test_spec_worked_example_pf_run_with_bullet_size_and_color() {
        // [MS-PPT] "Paragraph Formatting" 3.9.1 worked example: count=42,
        // indentLevel=0, masks has bulletHasSize+bulletColor+bulletSize
        // set, bulletFlags=0x..., bulletSize=0x0032, bulletColor=4 bytes.
        // No CF runs follow in this fixture.
        let masks = PF_BULLET_HAS_SIZE | PF_BULLET_COLOR | PF_BULLET_SIZE;
        let mut data = Vec::new();
        data.extend(le32(42)); // count
        data.extend(le16(0)); // indentLevel
        data.extend(le32(masks));
        data.extend(le16(0x0004)); // bulletFlags (fBulletHasSize bit set)
        data.extend(le16(0x0032)); // bulletSize = 50
        data.extend([0, 0, 0xFF, 0xFE]); // bulletColor: rgb=(0,0,255), index=0xFE -> explicit
        // One CF run covering all 42 (41 + implicit trailing mark), no formatting.
        data.extend(le32(42));
        data.extend(le32(0));

        let (para, chars) = parse_style_text_prop(&data, 41);
        assert_eq!(para.len(), 1);
        assert_eq!(para[0].start, 0);
        assert_eq!(para[0].end, 41);
        assert_eq!(chars.len(), 1);
        assert_eq!(chars[0].end, 41);
    }

    #[test]
    fn test_bold_italic_underline_decoded_from_cf_run() {
        let mut data = Vec::new();
        // One PF run covering everything, no optional fields.
        data.extend(le32(6));
        data.extend(le16(0));
        data.extend(le32(0));
        // One CF run: bold+italic+underline all TRUE.
        data.extend(le32(5));
        data.extend(le32(CF_BOLD | CF_ITALIC | CF_UNDERLINE));
        data.extend(le16(0x0001 | 0x0002 | 0x0004)); // fontStyle

        let (_, chars) = parse_style_text_prop(&data, 5);
        assert_eq!(chars.len(), 1);
        assert_eq!(chars[0].format.bold, Some(true));
        assert_eq!(chars[0].format.italic, Some(true));
        assert_eq!(chars[0].format.underline, Some(true));
    }

    #[test]
    fn test_two_char_runs_with_different_formatting_split_correctly() {
        let mut data = Vec::new();
        data.extend(le32(10));
        data.extend(le16(0));
        data.extend(le32(0)); // PF run, no fields
        // CF run 1: chars 0..4, bold.
        data.extend(le32(4));
        data.extend(le32(CF_BOLD));
        data.extend(le16(0x0001));
        // CF run 2: chars 4..10, italic + size 18pt.
        data.extend(le32(6));
        data.extend(le32(CF_ITALIC | CF_SIZE));
        data.extend(le16(0x0002));
        data.extend(le16(18u16)); // fontSize as i16 bit pattern

        let (_, chars) = parse_style_text_prop(&data, 9);
        assert_eq!(chars.len(), 2);
        assert_eq!(chars[0].start, 0);
        assert_eq!(chars[0].end, 4);
        assert_eq!(chars[0].format.bold, Some(true));
        assert_eq!(chars[1].start, 4);
        assert_eq!(chars[1].end, 9); // clamped from 10 (count) down to text_char_len
        assert_eq!(chars[1].format.italic, Some(true));
        assert_eq!(chars[1].format.font_size, Some(18));
    }

    #[test]
    fn test_explicit_rgb_color_decoded_scheme_index_left_unset() {
        let mut data = Vec::new();
        data.extend(le32(2));
        data.extend(le16(0));
        data.extend(le32(0));
        data.extend(le32(2));
        data.extend(le32(CF_COLOR));
        data.extend([0x10, 0x20, 0x30, 0xFE]); // explicit RGB

        let (_, chars) = parse_style_text_prop(&data, 1);
        assert_eq!(chars[0].format.color, Some([0x10, 0x20, 0x30]));

        let mut data2 = Vec::new();
        data2.extend(le32(2));
        data2.extend(le16(0));
        data2.extend(le32(0));
        data2.extend(le32(2));
        data2.extend(le32(CF_COLOR));
        data2.extend([0x10, 0x20, 0x30, 0x03]); // scheme index (title text color)

        let (_, chars2) = parse_style_text_prop(&data2, 1);
        assert_eq!(chars2[0].format.color, None);
    }

    #[test]
    fn test_alignment_decoded_from_pf_run() {
        let mut data = Vec::new();
        data.extend(le32(5));
        data.extend(le16(0));
        data.extend(le32(PF_ALIGN));
        data.extend(le16(1)); // Tx_ALIGNCenter
        data.extend(le32(5));
        data.extend(le32(0));

        let (para, _) = parse_style_text_prop(&data, 4);
        assert_eq!(para[0].format.alignment, Some(1));
    }

    #[test]
    fn test_tab_stops_are_skipped_without_desyncing_later_fields() {
        let mut data = Vec::new();
        // PF run with tabStops (2 entries = 8 bytes) followed by align,
        // to prove the fontAlign/align fields after tabStops still parse
        // at the right offset.
        data.extend(le32(3));
        data.extend(le16(0));
        data.extend(le32(PF_TAB_STOPS | PF_ALIGN));
        // NOTE: byte order in the wire format is mask-bit order, but this
        // module reads fields in the TextPFException struct's declared
        // order (align before tabStops in the field table, tabStops is
        // actually declared after align/leftMargin/indent/defaultTabSize).
        data.extend(le16(2)); // textAlignment (align bit, read before tabStops)
        data.extend(le16(2)); // tabStops.count = 2
        data.extend(le32(0)); // tab stop 1 (4 bytes)
        data.extend(le32(0)); // tab stop 2 (4 bytes)
        data.extend(le32(3));
        data.extend(le32(0));

        let (para, chars) = parse_style_text_prop(&data, 2);
        assert_eq!(para[0].format.alignment, Some(2));
        assert_eq!(chars.len(), 1); // proves the CF run after it parsed cleanly
    }

    #[test]
    fn test_truncated_atom_does_not_panic() {
        let (para, chars) = parse_style_text_prop(&[0xFF, 0x00], 10);
        assert!(para.is_empty());
        assert!(chars.is_empty());
    }

    #[test]
    fn test_zero_length_run_does_not_infinite_loop() {
        let mut data = Vec::new();
        data.extend(le32(0));
        data.extend(le16(0));
        data.extend(le32(0));
        let (para, _) = parse_style_text_prop(&data, 5);
        assert_eq!(para.len(), 0); // zero-length span isn't pushed, but parsing terminates
    }

    /// A `TxMasterStyleAtom` body with one indent level
    /// (Title/Body text types carry no per-level `indentLevel` field,
    /// per Apache POI's own `TxMasterStyleAtom#init()`), setting
    /// alignment + font size.
    #[test]
    fn test_tx_master_style_atom_single_level_no_indent_field() {
        let mut data = Vec::new();
        data.extend(le16(1)); // levels = 1
        data.extend(le32(PF_ALIGN));
        data.extend(le16(2)); // alignment = right
        data.extend(le32(CF_SIZE));
        data.extend((32i16).to_le_bytes()); // font_size = 32

        let levels = parse_tx_master_style_atom(&data, 0); // Title (native id 0, no indent field)
        assert_eq!(levels.len(), 1);
        assert_eq!(levels[0].0.alignment, Some(2));
        assert_eq!(levels[0].1.font_size, Some(32));
    }

    /// `CenterBody`(5)/`CenterTitle`(6)/`HalfBody`(7)/`QuarterBody`(8) DO
    /// carry an explicit 2-byte `indentLevel` before each level's
    /// paragraph mask — getting this wrong would desync every field
    /// after it.
    #[test]
    fn test_tx_master_style_atom_center_body_has_indent_level_field() {
        let mut data = Vec::new();
        data.extend(le16(1)); // levels = 1
        data.extend(le16(0)); // indentLevel (present for type >= 5)
        data.extend(le32(PF_ALIGN));
        data.extend(le16(1)); // alignment = center
        data.extend(le32(0)); // no CF fields set
        let levels = parse_tx_master_style_atom(&data, 5); // CenterBody
        assert_eq!(levels.len(), 1);
        assert_eq!(levels[0].0.alignment, Some(1));
    }

    #[test]
    fn test_tx_master_style_atom_truncated_does_not_panic() {
        let levels = parse_tx_master_style_atom(&[0x01, 0x00], 0);
        assert!(levels.is_empty());
    }

    #[test]
    fn test_char_format_inherit_from_fills_only_unset_fields() {
        let direct = CharFormat { bold: Some(true), ..Default::default() };
        let master = CharFormat { bold: Some(false), font_size: Some(44), ..Default::default() };
        let merged = direct.inherit_from(&master);
        assert_eq!(merged.bold, Some(true), "a direct value must never be overridden by the master");
        assert_eq!(merged.font_size, Some(44), "an unset field must be filled from the master");
    }

    #[test]
    fn test_para_format_inherit_from_fills_only_unset_fields() {
        let direct = ParaFormat { alignment: Some(0) };
        let master = ParaFormat { alignment: Some(2) };
        assert_eq!(direct.inherit_from(&master).alignment, Some(0));
        let unset = ParaFormat { alignment: None };
        assert_eq!(unset.inherit_from(&master).alignment, Some(2));
    }
}
