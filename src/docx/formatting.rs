use crate::core::theme::ColorRef;
use crate::core::units::{HalfPoint, Twip};

/// Run-level formatting properties (`w:rPr`).
#[derive(Debug, Clone, Default)]
pub struct RunProperties {
    /// Bold toggle.
    pub bold: Option<bool>,
    /// Italic toggle.
    pub italic: Option<bool>,
    /// Underline style.
    pub underline: Option<UnderlineType>,
    /// Single strikethrough toggle.
    pub strike: Option<bool>,
    /// Double strikethrough toggle.
    pub dstrike: Option<bool>,
    /// Font size in half-points.
    pub font_size: Option<HalfPoint>,
    /// Primary font family name.
    pub font_name: Option<String>,
    /// Text color reference.
    pub color: Option<ColorRef>,
    /// Highlight color name (e.g., `"yellow"`).
    pub highlight: Option<String>,
    /// Vertical alignment (superscript/subscript).
    pub vertical_align: Option<VerticalAlign>,
    /// Character style ID.
    pub style_id: Option<String>,
    /// All-caps toggle (`<w:caps/>`).
    pub caps: Option<bool>,
    /// Small-caps toggle (`<w:smallCaps/>`).
    pub small_caps: Option<bool>,
    /// Character spacing from `<w:spacing w:val="N"/>` inside `w:rPr`.
    /// Signed, in the same units the writer emits (twentieths of a point).
    pub char_spacing: Option<i32>,
    /// Run shading fill from `<w:shd w:fill="RRGGBB"/>` inside `w:rPr`.
    /// The writer uses this (not `<w:highlight>`) to encode
    /// `TextSpan::highlight`, so reading it back is what closes the
    /// write→read loop for highlighted text.
    pub shading_fill: Option<String>,
}

/// Paragraph-level formatting properties (`w:pPr`).
#[derive(Debug, Clone, Default)]
pub struct ParagraphProperties {
    /// Paragraph style ID.
    pub style_id: Option<String>,
    /// Text justification.
    pub justification: Option<Justification>,
    /// Paragraph indentation.
    pub indent: Option<ParagraphIndent>,
    /// Paragraph spacing.
    pub spacing: Option<ParagraphSpacing>,
    /// Numbering reference for list paragraphs.
    pub numbering_ref: Option<NumberingRef>,
    /// Outline level (0 = Heading 1, 1 = Heading 2, …).
    pub outline_level: Option<u8>,
    /// Paragraph-mark run properties (`w:rPr` inside `w:pPr`).
    pub run_properties: Option<RunProperties>,
    /// Frame position from `<w:framePr>`. When present this paragraph is
    /// absolutely positioned on the page (used by layout-preserving
    /// PDF-derived DOCX, e.g. pdf_oxide's `to_docx_bytes_layout`).
    pub frame_position: Option<FrameProps>,
    /// Section properties from `<w:sectPr>` inside this paragraph's `<w:pPr>`.
    /// When present this paragraph terminates a section — the properties
    /// describe the section that ends here.
    pub section_properties: Option<super::SectionProperties>,
    /// True when the paragraph has a `<w:pBdr><w:bottom .../></w:pBdr>`.
    /// Used to recover horizontal rules: pdf_to_ir emits
    /// `Element::ThematicBreak` which round-trips through DOCX as an
    /// empty paragraph with a single bottom border. Without
    /// preserving this flag the rule would be silently dropped on
    /// re-parse and turned into a plain empty paragraph.
    #[allow(dead_code)]
    pub has_bottom_border: bool,
    /// Full `<w:pBdr>` edge styling. `has_bottom_border` stays as the
    /// cheap horizontal-rule probe; this carries the actual widths,
    /// colours and styles so they survive a read.
    pub borders: Option<ParagraphBorders>,
    /// `<w:keepNext/>` — keep with the following paragraph. `None` when the
    /// element is absent, so a style-inherited value is distinguishable from
    /// an explicit `w:val="0"` that turns it off.
    pub keep_next: Option<bool>,
    /// `<w:keepLines/>` — keep all lines of this paragraph together.
    pub keep_lines: Option<bool>,
    /// `<w:pageBreakBefore/>` — force a page break before this paragraph.
    pub page_break_before: Option<bool>,
    /// Paragraph shading (`<w:shd>`) — background fill.
    pub shading: Option<super::table::Shading>,
    /// Custom tab stops from `<w:tabs>`.
    pub tabs: Vec<TabStopDef>,
}

/// One `<w:tab>` entry inside `<w:tabs>`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TabStopDef {
    /// Position in twips (`w:pos`).
    pub position_twips: i32,
    /// Alignment (`w:val`): left/center/right/decimal/bar.
    pub alignment: String,
    /// Leader character style (`w:leader`).
    pub leader: Option<String>,
}

/// A single border edge (`<w:top>`, `<w:left>`, …) as it appears
/// inside `<w:pBdr>`, `<w:tblBorders>` and `<w:tcBorders>`.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct BorderEdge {
    /// `w:val` — the line style name (`single`, `double`, `dotted`, …).
    pub style: Option<String>,
    /// `w:color` — `RRGGBB` or `auto`.
    pub color: Option<String>,
    /// `w:sz` — line width in eighths of a point.
    pub size: Option<u32>,
    /// `w:space` — padding between border and text, in points.
    pub space: Option<u32>,
}

/// The five paragraph border edges (`<w:pBdr>`).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ParagraphBorders {
    /// Top edge.
    pub top: Option<BorderEdge>,
    /// Bottom edge.
    pub bottom: Option<BorderEdge>,
    /// Left edge.
    pub left: Option<BorderEdge>,
    /// Right edge.
    pub right: Option<BorderEdge>,
    /// Edge drawn between consecutive paragraphs sharing this border set.
    pub between: Option<BorderEdge>,
}

/// Table / cell border edges (`<w:tblBorders>`, `<w:tcBorders>`).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TableBorders {
    /// Top edge.
    pub top: Option<BorderEdge>,
    /// Bottom edge.
    pub bottom: Option<BorderEdge>,
    /// Left (`w:left` / `w:start`) edge.
    pub left: Option<BorderEdge>,
    /// Right (`w:right` / `w:end`) edge.
    pub right: Option<BorderEdge>,
    /// Interior horizontal edges.
    pub inside_h: Option<BorderEdge>,
    /// Interior vertical edges.
    pub inside_v: Option<BorderEdge>,
}

/// Parse one border edge element's attributes.
pub(crate) fn parse_border_edge(e: &BytesStart) -> BorderEdge {
    BorderEdge {
        style: xml::optional_attr_str(e, b"w:val")
            .ok()
            .flatten()
            .map(|v| v.into_owned()),
        color: xml::optional_attr_str(e, b"w:color")
            .ok()
            .flatten()
            .map(|v| v.into_owned()),
        size: xml::optional_attr_str(e, b"w:sz")
            .ok()
            .flatten()
            .and_then(|v| v.parse().ok()),
        space: xml::optional_attr_str(e, b"w:space")
            .ok()
            .flatten()
            .and_then(|v| v.parse().ok()),
    }
}

impl RunProperties {
    /// Overlay `src` on top of `self`: every field `src` states explicitly
    /// wins, everything else is left alone. Used to fold a style chain
    /// (document defaults → style → parent style → direct `w:rPr`) into one
    /// effective property set.
    pub fn overlay(&mut self, src: &RunProperties) {
        macro_rules! take {
            ($($f:ident),* $(,)?) => {$(
                if src.$f.is_some() { self.$f = src.$f.clone(); }
            )*};
        }
        take!(
            bold,
            italic,
            underline,
            strike,
            dstrike,
            font_size,
            font_name,
            color,
            highlight,
            vertical_align,
            style_id,
            caps,
            small_caps,
            char_spacing,
            shading_fill,
        );
    }
}

impl ParagraphProperties {
    /// Overlay `src` on top of `self`. See [`RunProperties::overlay`].
    ///
    /// `run_properties` is merged recursively rather than replaced, so a
    /// paragraph style that sets only the font does not wipe out the bold
    /// flag its parent style set.
    pub fn overlay(&mut self, src: &ParagraphProperties) {
        macro_rules! take {
            ($($f:ident),* $(,)?) => {$(
                if src.$f.is_some() { self.$f = src.$f.clone(); }
            )*};
        }
        take!(
            style_id,
            justification,
            indent,
            spacing,
            numbering_ref,
            outline_level,
            frame_position,
            section_properties,
            borders,
            keep_next,
            keep_lines,
            page_break_before,
            shading,
        );
        if !src.tabs.is_empty() {
            self.tabs = src.tabs.clone();
        }
        if src.has_bottom_border {
            self.has_bottom_border = true;
        }
        match (self.run_properties.as_mut(), src.run_properties.as_ref()) {
            (Some(dst), Some(s)) => dst.overlay(s),
            (None, Some(s)) => self.run_properties = Some(s.clone()),
            _ => {},
        }
    }
}

/// `<w:framePr>` attributes — page-anchored frame coordinates in twips.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct FrameProps {
    /// X position in twips, anchored to the page (top-left).
    pub x_twips: i32,
    /// Y position in twips, anchored to the page (top-left).
    pub y_twips: i32,
    /// Frame width in twips.
    pub width_twips: i32,
    /// Frame height in twips.
    pub height_twips: i32,
}

/// Underline style.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum UnderlineType {
    /// Single underline.
    Single,
    /// Double underline.
    Double,
    /// Thick (heavy) underline.
    Thick,
    /// Dotted underline.
    Dotted,
    /// Dashed underline.
    Dash,
    /// Dot-dash underline.
    DotDash,
    /// Dot-dot-dash underline.
    DotDotDash,
    /// Wave underline.
    Wave,
    /// Underline under words only.
    Words,
    /// No underline (explicit removal).
    None,
    /// Any other underline value.
    Other(String),
}

/// Vertical alignment (superscript/subscript).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VerticalAlign {
    /// Superscript text.
    Superscript,
    /// Subscript text.
    Subscript,
    /// Normal (baseline) alignment.
    Baseline,
}

/// Paragraph justification / alignment.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Justification {
    /// Left-aligned (default).
    Left,
    /// Centered.
    Center,
    /// Right-aligned.
    Right,
    /// Justified (both edges).
    Both,
    /// Distributed (East Asian spacing).
    Distribute,
}

/// Paragraph indentation values.
#[derive(Debug, Clone, Default)]
pub struct ParagraphIndent {
    /// Left indent in twips.
    pub left: Option<Twip>,
    /// Right indent in twips.
    pub right: Option<Twip>,
    /// First-line indent in twips (positive = indent).
    pub first_line: Option<Twip>,
    /// Hanging indent in twips (first-line de-indent).
    pub hanging: Option<Twip>,
}

/// Paragraph spacing.
#[derive(Debug, Clone, Default)]
pub struct ParagraphSpacing {
    /// Space before the paragraph in twips.
    pub before: Option<Twip>,
    /// Space after the paragraph in twips.
    pub after: Option<Twip>,
    /// Line spacing value and rule.
    pub line: Option<SpacingLine>,
}

/// Line spacing rule and value.
#[derive(Debug, Clone)]
pub struct SpacingLine {
    /// Line spacing value (interpretation depends on `rule`).
    pub value: i32,
    /// The rule governing how `value` is applied.
    pub rule: Option<LineSpacingRule>,
}

/// Line spacing rule.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LineSpacingRule {
    /// Automatic (proportional to font size).
    Auto,
    /// Exact line height in twips.
    Exact,
    /// Minimum line height in twips.
    AtLeast,
}

/// Reference to a numbering definition from a paragraph.
#[derive(Debug, Clone)]
pub struct NumberingRef {
    /// Numbering instance ID.
    pub num_id: u32,
    /// Indent level (0-based).
    pub ilvl: u8,
}

// ---------------------------------------------------------------------------
// Parsing helpers
// ---------------------------------------------------------------------------

use quick_xml::events::{BytesStart, Event};

use crate::core::xml;

/// Parse a justification value string.
pub(crate) fn parse_justification_value(val: &str) -> Justification {
    match val {
        "left" | "start" => Justification::Left,
        "center" => Justification::Center,
        "right" | "end" => Justification::Right,
        "both" => Justification::Both,
        "distribute" => Justification::Distribute,
        _ => Justification::Left,
    }
}

/// Parse `w:rPr` element children using NsReader. Caller has already consumed the `Start(w:rPr)` event.
/// NOTE: Only used by test code via NsReader. Production code uses `parse_run_properties_fast`.
#[cfg(test)]
pub(crate) fn parse_run_properties(
    reader: &mut quick_xml::NsReader<&[u8]>,
) -> crate::core::Result<RunProperties> {
    let wml = xml::ns::WML;
    let mut props = RunProperties::default();

    loop {
        match reader.read_resolved_event()? {
            (ref resolve, Event::Start(ref e)) => {
                if xml::matches_ns(resolve, wml) {
                    let local = e.local_name();
                    match local.as_ref() {
                        b"b" => {
                            props.bold = Some(parse_toggle(e));
                            xml::skip_element(reader)?;
                        },
                        b"i" => {
                            props.italic = Some(parse_toggle(e));
                            xml::skip_element(reader)?;
                        },
                        b"strike" => {
                            props.strike = Some(parse_toggle(e));
                            xml::skip_element(reader)?;
                        },
                        b"dstrike" => {
                            props.dstrike = Some(parse_toggle(e));
                            xml::skip_element(reader)?;
                        },
                        b"u" => {
                            props.underline = Some(parse_underline(e));
                            xml::skip_element(reader)?;
                        },
                        b"sz" => {
                            if let Some(val) = parse_half_point_val(e)? {
                                props.font_size = Some(val);
                            }
                            xml::skip_element(reader)?;
                        },
                        b"rFonts" => {
                            if let Ok(Some(ascii)) = xml::optional_attr_str(e, b"w:ascii") {
                                props.font_name = Some(ascii.into_owned());
                            }
                            xml::skip_element(reader)?;
                        },
                        b"color" => {
                            props.color = parse_color_ref(e)?;
                            xml::skip_element(reader)?;
                        },
                        b"highlight" => {
                            if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                                props.highlight = Some(val.into_owned());
                            }
                            xml::skip_element(reader)?;
                        },
                        b"vertAlign" => {
                            if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                                props.vertical_align = Some(match val.as_ref() {
                                    "superscript" => VerticalAlign::Superscript,
                                    "subscript" => VerticalAlign::Subscript,
                                    _ => VerticalAlign::Baseline,
                                });
                            }
                            xml::skip_element(reader)?;
                        },
                        b"rStyle" => {
                            if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                                props.style_id = Some(val.into_owned());
                            }
                            xml::skip_element(reader)?;
                        },
                        _ => {
                            xml::skip_element(reader)?;
                        },
                    }
                } else {
                    xml::skip_element(reader)?;
                }
            },
            (ref resolve, Event::Empty(ref e)) if xml::matches_ns(resolve, wml) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"b" => props.bold = Some(parse_toggle(e)),
                    b"i" => props.italic = Some(parse_toggle(e)),
                    b"strike" => props.strike = Some(parse_toggle(e)),
                    b"dstrike" => props.dstrike = Some(parse_toggle(e)),
                    b"u" => props.underline = Some(parse_underline(e)),
                    b"sz" => {
                        if let Some(val) = parse_half_point_val(e)? {
                            props.font_size = Some(val);
                        }
                    },
                    b"rFonts" => {
                        if let Ok(Some(ascii)) = xml::optional_attr_str(e, b"w:ascii") {
                            props.font_name = Some(ascii.into_owned());
                        }
                    },
                    b"color" => {
                        props.color = parse_color_ref(e)?;
                    },
                    b"highlight" => {
                        if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                            props.highlight = Some(val.into_owned());
                        }
                    },
                    b"vertAlign" => {
                        if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                            props.vertical_align = Some(match val.as_ref() {
                                "superscript" => VerticalAlign::Superscript,
                                "subscript" => VerticalAlign::Subscript,
                                _ => VerticalAlign::Baseline,
                            });
                        }
                    },
                    b"rStyle" => {
                        if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                            props.style_id = Some(val.into_owned());
                        }
                    },
                    _ => {},
                }
            },
            (ref resolve, Event::End(ref e))
                if xml::matches_ns(resolve, wml) && e.local_name().as_ref() == b"rPr" =>
            {
                break;
            },
            (_, Event::Eof) => break,
            _ => {},
        }
    }
    Ok(props)
}

/// Parse `w:pPr` element children using NsReader. Caller has already consumed the `Start(w:pPr)` event.
/// NOTE: Only used by test code via NsReader. Production code uses `parse_paragraph_properties_fast`.
#[cfg(test)]
pub(crate) fn parse_paragraph_properties(
    reader: &mut quick_xml::NsReader<&[u8]>,
) -> crate::core::Result<ParagraphProperties> {
    let wml = xml::ns::WML;
    let mut props = ParagraphProperties::default();

    loop {
        match reader.read_resolved_event()? {
            (ref resolve, Event::Start(ref e)) => {
                if xml::matches_ns(resolve, wml) {
                    let local = e.local_name();
                    match local.as_ref() {
                        b"pStyle" => {
                            if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                                props.style_id = Some(val.into_owned());
                            }
                            xml::skip_element(reader)?;
                        },
                        b"jc" => {
                            if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                                props.justification = Some(parse_justification_value(&val));
                            }
                            xml::skip_element(reader)?;
                        },
                        b"ind" => {
                            props.indent = Some(parse_indent(e)?);
                            xml::skip_element(reader)?;
                        },
                        b"spacing" => {
                            props.spacing = Some(parse_spacing(e)?);
                            xml::skip_element(reader)?;
                        },
                        b"numPr" => {
                            props.numbering_ref = Some(parse_num_pr(reader)?);
                        },
                        b"outlineLvl" => {
                            if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                                if let Ok(lvl) = val.parse::<u8>() {
                                    props.outline_level = Some(lvl);
                                }
                            }
                            xml::skip_element(reader)?;
                        },
                        b"rPr" => {
                            props.run_properties = Some(parse_run_properties(reader)?);
                        },
                        _ => {
                            xml::skip_element(reader)?;
                        },
                    }
                } else {
                    xml::skip_element(reader)?;
                }
            },
            (ref resolve, Event::Empty(ref e)) if xml::matches_ns(resolve, wml) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"pStyle" => {
                        if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                            props.style_id = Some(val.into_owned());
                        }
                    },
                    b"jc" => {
                        if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                            props.justification = Some(parse_justification_value(&val));
                        }
                    },
                    b"ind" => {
                        props.indent = Some(parse_indent(e)?);
                    },
                    b"spacing" => {
                        props.spacing = Some(parse_spacing(e)?);
                    },
                    b"outlineLvl" => {
                        if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                            if let Ok(lvl) = val.parse::<u8>() {
                                props.outline_level = Some(lvl);
                            }
                        }
                    },
                    _ => {},
                }
            },
            (ref resolve, Event::End(ref e))
                if xml::matches_ns(resolve, wml) && e.local_name().as_ref() == b"pPr" =>
            {
                break;
            },
            (_, Event::Eof) => break,
            _ => {},
        }
    }
    Ok(props)
}

// ---------------------------------------------------------------------------
// Fast (plain Reader) variants — no namespace resolution
// ---------------------------------------------------------------------------

/// Parse `w:rPr` using plain `Reader` (no namespace resolution).
pub(crate) fn parse_run_properties_fast(
    reader: &mut quick_xml::Reader<&[u8]>,
) -> crate::core::Result<RunProperties> {
    let mut props = RunProperties::default();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"b" => {
                        props.bold = Some(parse_toggle(e));
                        xml::skip_element_fast(reader)?;
                    },
                    b"i" => {
                        props.italic = Some(parse_toggle(e));
                        xml::skip_element_fast(reader)?;
                    },
                    b"strike" => {
                        props.strike = Some(parse_toggle(e));
                        xml::skip_element_fast(reader)?;
                    },
                    b"dstrike" => {
                        props.dstrike = Some(parse_toggle(e));
                        xml::skip_element_fast(reader)?;
                    },
                    b"u" => {
                        props.underline = Some(parse_underline(e));
                        xml::skip_element_fast(reader)?;
                    },
                    b"sz" => {
                        if let Some(val) = parse_half_point_val(e)? {
                            props.font_size = Some(val);
                        }
                        xml::skip_element_fast(reader)?;
                    },
                    b"rFonts" => {
                        if let Ok(Some(ascii)) = xml::optional_attr_str(e, b"w:ascii") {
                            props.font_name = Some(ascii.into_owned());
                        }
                        xml::skip_element_fast(reader)?;
                    },
                    b"color" => {
                        props.color = parse_color_ref(e)?;
                        xml::skip_element_fast(reader)?;
                    },
                    b"highlight" => {
                        if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                            props.highlight = Some(val.into_owned());
                        }
                        xml::skip_element_fast(reader)?;
                    },
                    b"vertAlign" => {
                        if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                            props.vertical_align = Some(match val.as_ref() {
                                "superscript" => VerticalAlign::Superscript,
                                "subscript" => VerticalAlign::Subscript,
                                _ => VerticalAlign::Baseline,
                            });
                        }
                        xml::skip_element_fast(reader)?;
                    },
                    b"rStyle" => {
                        if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                            props.style_id = Some(val.into_owned());
                        }
                        xml::skip_element_fast(reader)?;
                    },
                    b"caps" => {
                        props.caps = Some(parse_toggle(e));
                        xml::skip_element_fast(reader)?;
                    },
                    b"smallCaps" => {
                        props.small_caps = Some(parse_toggle(e));
                        xml::skip_element_fast(reader)?;
                    },
                    b"spacing" => {
                        props.char_spacing = parse_signed_val(e);
                        xml::skip_element_fast(reader)?;
                    },
                    b"shd" => {
                        props.shading_fill = xml::optional_attr_str(e, b"w:fill")
                            .ok()
                            .flatten()
                            .map(|v| v.into_owned());
                        xml::skip_element_fast(reader)?;
                    },
                    _ => {
                        xml::skip_element_fast(reader)?;
                    },
                }
            },
            Event::Empty(ref e) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"caps" => props.caps = Some(parse_toggle(e)),
                    b"smallCaps" => props.small_caps = Some(parse_toggle(e)),
                    b"spacing" => props.char_spacing = parse_signed_val(e),
                    b"shd" => {
                        props.shading_fill = xml::optional_attr_str(e, b"w:fill")
                            .ok()
                            .flatten()
                            .map(|v| v.into_owned());
                    },
                    b"b" => props.bold = Some(parse_toggle(e)),
                    b"i" => props.italic = Some(parse_toggle(e)),
                    b"strike" => props.strike = Some(parse_toggle(e)),
                    b"dstrike" => props.dstrike = Some(parse_toggle(e)),
                    b"u" => props.underline = Some(parse_underline(e)),
                    b"sz" => {
                        if let Some(val) = parse_half_point_val(e)? {
                            props.font_size = Some(val);
                        }
                    },
                    b"rFonts" => {
                        if let Ok(Some(ascii)) = xml::optional_attr_str(e, b"w:ascii") {
                            props.font_name = Some(ascii.into_owned());
                        }
                    },
                    b"color" => {
                        props.color = parse_color_ref(e)?;
                    },
                    b"highlight" => {
                        if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                            props.highlight = Some(val.into_owned());
                        }
                    },
                    b"vertAlign" => {
                        if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                            props.vertical_align = Some(match val.as_ref() {
                                "superscript" => VerticalAlign::Superscript,
                                "subscript" => VerticalAlign::Subscript,
                                _ => VerticalAlign::Baseline,
                            });
                        }
                    },
                    b"rStyle" => {
                        if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                            props.style_id = Some(val.into_owned());
                        }
                    },
                    _ => {},
                }
            },
            Event::End(ref e) if e.local_name().as_ref() == b"rPr" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(props)
}

/// Parse `w:pPr` using plain `Reader` (no namespace resolution).
pub(crate) fn parse_paragraph_properties_fast(
    reader: &mut quick_xml::Reader<&[u8]>,
) -> crate::core::Result<ParagraphProperties> {
    let mut props = ParagraphProperties::default();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"pStyle" => {
                        if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                            props.style_id = Some(val.into_owned());
                        }
                        xml::skip_element_fast(reader)?;
                    },
                    b"jc" => {
                        if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                            props.justification = Some(parse_justification_value(&val));
                        }
                        xml::skip_element_fast(reader)?;
                    },
                    b"ind" => {
                        props.indent = Some(parse_indent(e)?);
                        xml::skip_element_fast(reader)?;
                    },
                    b"spacing" => {
                        props.spacing = Some(parse_spacing(e)?);
                        xml::skip_element_fast(reader)?;
                    },
                    b"numPr" => {
                        props.numbering_ref = Some(parse_num_pr_fast(reader)?);
                    },
                    b"outlineLvl" => {
                        if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                            if let Ok(lvl) = val.parse::<u8>() {
                                props.outline_level = Some(lvl);
                            }
                        }
                        xml::skip_element_fast(reader)?;
                    },
                    b"rPr" => {
                        props.run_properties = Some(parse_run_properties_fast(reader)?);
                    },
                    b"framePr" => {
                        props.frame_position = parse_frame_pr(e);
                        xml::skip_element_fast(reader)?;
                    },
                    b"sectPr" => {
                        props.section_properties =
                            Some(super::parse_section_properties(reader, e)?);
                    },
                    b"pBdr" => {
                        // Capture every `<w:pBdr>` edge with its full
                        // styling. `has_bottom_border` stays the cheap
                        // horizontal-rule probe (empty paragraph + bottom
                        // border = the conventional DOCX `<hr/>`), but the
                        // widths, colours and styles now survive the read
                        // instead of being narrowed to that one boolean.
                        let borders = parse_paragraph_borders_fast(reader)?;
                        props.has_bottom_border = borders.bottom.is_some();
                        props.borders = Some(borders);
                    },
                    b"keepNext" => {
                        props.keep_next = Some(parse_toggle(e));
                        xml::skip_element_fast(reader)?;
                    },
                    b"keepLines" => {
                        props.keep_lines = Some(parse_toggle(e));
                        xml::skip_element_fast(reader)?;
                    },
                    b"pageBreakBefore" => {
                        props.page_break_before = Some(parse_toggle(e));
                        xml::skip_element_fast(reader)?;
                    },
                    b"shd" => {
                        props.shading = Some(parse_shading(e));
                        xml::skip_element_fast(reader)?;
                    },
                    b"tabs" => {
                        props.tabs = parse_tabs_fast(reader)?;
                    },
                    _ => {
                        xml::skip_element_fast(reader)?;
                    },
                }
            },
            Event::Empty(ref e) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"keepNext" => props.keep_next = Some(parse_toggle(e)),
                    b"keepLines" => props.keep_lines = Some(parse_toggle(e)),
                    b"pageBreakBefore" => props.page_break_before = Some(parse_toggle(e)),
                    b"shd" => props.shading = Some(parse_shading(e)),
                    b"pStyle" => {
                        if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                            props.style_id = Some(val.into_owned());
                        }
                    },
                    b"jc" => {
                        if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                            props.justification = Some(parse_justification_value(&val));
                        }
                    },
                    b"ind" => {
                        props.indent = Some(parse_indent(e)?);
                    },
                    b"spacing" => {
                        props.spacing = Some(parse_spacing(e)?);
                    },
                    b"framePr" => {
                        props.frame_position = parse_frame_pr(e);
                    },
                    b"outlineLvl" => {
                        if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                            if let Ok(lvl) = val.parse::<u8>() {
                                props.outline_level = Some(lvl);
                            }
                        }
                    },
                    _ => {},
                }
            },
            Event::End(ref e) if e.local_name().as_ref() == b"pPr" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(props)
}

/// Parse a signed `w:val` integer attribute (used by `<w:spacing>` in `w:rPr`).
fn parse_signed_val(e: &BytesStart) -> Option<i32> {
    xml::optional_attr_str(e, b"w:val")
        .ok()
        .flatten()
        .and_then(|v| v.parse().ok())
}

/// Parse a `<w:shd>` element's attributes.
pub(crate) fn parse_shading(e: &BytesStart) -> super::table::Shading {
    super::table::Shading {
        fill: xml::optional_attr_str(e, b"w:fill")
            .ok()
            .flatten()
            .map(|v| v.into_owned()),
        color: xml::optional_attr_str(e, b"w:color")
            .ok()
            .flatten()
            .map(|v| v.into_owned()),
        pattern: xml::optional_attr_str(e, b"w:val")
            .ok()
            .flatten()
            .map(|v| v.into_owned()),
    }
}

/// Parse the children of `<w:pBdr>`. The caller has consumed the start tag.
fn parse_paragraph_borders_fast(
    reader: &mut quick_xml::Reader<&[u8]>,
) -> crate::core::Result<ParagraphBorders> {
    let mut b = ParagraphBorders::default();
    loop {
        match reader.read_event()? {
            Event::Start(ref e) | Event::Empty(ref e) => {
                let edge = parse_border_edge(e);
                match e.local_name().as_ref() {
                    b"top" => b.top = Some(edge),
                    b"bottom" => b.bottom = Some(edge),
                    b"left" | b"start" => b.left = Some(edge),
                    b"right" | b"end" => b.right = Some(edge),
                    b"between" => b.between = Some(edge),
                    _ => {},
                }
            },
            Event::End(ref e) if e.local_name().as_ref() == b"pBdr" => break,
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(b)
}

/// Parse the children of `<w:tblBorders>` / `<w:tcBorders>`. The caller has
/// consumed the start tag; `end` names the closing element to stop at.
pub(crate) fn parse_table_borders_fast(
    reader: &mut quick_xml::Reader<&[u8]>,
    end: &[u8],
) -> crate::core::Result<TableBorders> {
    let mut b = TableBorders::default();
    loop {
        match reader.read_event()? {
            Event::Start(ref e) | Event::Empty(ref e) => {
                let edge = parse_border_edge(e);
                match e.local_name().as_ref() {
                    b"top" => b.top = Some(edge),
                    b"bottom" => b.bottom = Some(edge),
                    b"left" | b"start" => b.left = Some(edge),
                    b"right" | b"end" => b.right = Some(edge),
                    b"insideH" => b.inside_h = Some(edge),
                    b"insideV" => b.inside_v = Some(edge),
                    _ => {},
                }
            },
            Event::End(ref e) if e.local_name().as_ref() == end => break,
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(b)
}

/// Parse the children of `<w:tabs>`. The caller has consumed the start tag.
fn parse_tabs_fast(reader: &mut quick_xml::Reader<&[u8]>) -> crate::core::Result<Vec<TabStopDef>> {
    let mut tabs = Vec::new();
    loop {
        match reader.read_event()? {
            Event::Start(ref e) | Event::Empty(ref e) if e.local_name().as_ref() == b"tab" => {
                let pos = xml::optional_attr_str(e, b"w:pos")
                    .ok()
                    .flatten()
                    .and_then(|v| v.parse::<i32>().ok());
                if let Some(position_twips) = pos {
                    tabs.push(TabStopDef {
                        position_twips,
                        alignment: xml::optional_attr_str(e, b"w:val")
                            .ok()
                            .flatten()
                            .map(|v| v.into_owned())
                            .unwrap_or_else(|| "left".to_string()),
                        leader: xml::optional_attr_str(e, b"w:leader")
                            .ok()
                            .flatten()
                            .map(|v| v.into_owned()),
                    });
                }
            },
            Event::End(ref e) if e.local_name().as_ref() == b"tabs" => break,
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(tabs)
}

fn parse_num_pr_fast(reader: &mut quick_xml::Reader<&[u8]>) -> crate::core::Result<NumberingRef> {
    let mut num_id: u32 = 0;
    let mut ilvl: u8 = 0;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) | Event::Empty(ref e) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"numId" => {
                        if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                            num_id = val.parse().unwrap_or(0);
                        }
                    },
                    b"ilvl" => {
                        if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                            ilvl = val.parse().unwrap_or(0);
                        }
                    },
                    _ => {},
                }
            },
            Event::End(ref e) if e.local_name().as_ref() == b"numPr" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(NumberingRef { num_id, ilvl })
}

/// Parse a boolean toggle attribute. `<w:b/>` = true, `<w:b w:val="0"/>` = false.
fn parse_toggle(e: &BytesStart) -> bool {
    xml::parse_toggle(e, b"w:val")
}

fn parse_underline(e: &BytesStart) -> UnderlineType {
    match xml::optional_attr_str(e, b"w:val") {
        Ok(Some(ref val)) => match val.as_ref() {
            "single" => UnderlineType::Single,
            "double" => UnderlineType::Double,
            "thick" => UnderlineType::Thick,
            "dotted" => UnderlineType::Dotted,
            "dash" => UnderlineType::Dash,
            "dotDash" => UnderlineType::DotDash,
            "dotDotDash" => UnderlineType::DotDotDash,
            "wave" => UnderlineType::Wave,
            "words" => UnderlineType::Words,
            "none" => UnderlineType::None,
            other => UnderlineType::Other(other.to_string()),
        },
        _ => UnderlineType::Single,
    }
}

/// Parse a numeric value, stripping any trailing unit suffix (e.g., "20pt" → 20).
/// OOXML Strict format uses unit suffixes and decimal values (e.g., "12.95pt");
/// Transitional uses bare integers. We truncate decimals for integer types.
fn parse_numeric<T: std::str::FromStr>(s: &str) -> std::result::Result<T, T::Err> {
    let numeric = s.trim_end_matches(|c: char| c.is_ascii_alphabetic() || c == '%');
    // Try direct parse first (fast path for integers)
    if let Ok(v) = numeric.parse() {
        return Ok(v);
    }
    // If that fails and there's a decimal point, try parsing as f64 and truncating
    if numeric.contains('.') {
        if let Ok(f) = numeric.parse::<f64>() {
            // Round to nearest integer and try to parse the string representation
            let rounded = format!("{}", f.round() as i64);
            if let Ok(v) = rounded.parse() {
                return Ok(v);
            }
        }
    }
    // Final fallback — return the original parse error
    numeric.parse()
}

fn parse_half_point_val(e: &BytesStart) -> crate::core::Result<Option<HalfPoint>> {
    match xml::optional_attr_str(e, b"w:val")? {
        Some(ref val) => {
            let v: u32 = parse_numeric(val)?;
            Ok(Some(HalfPoint(v)))
        },
        None => Ok(None),
    }
}

fn parse_color_ref(e: &BytesStart) -> crate::core::Result<Option<ColorRef>> {
    use crate::core::theme::{RgbColor, ThemeColorSlot};

    let val = xml::optional_attr_str(e, b"w:val")?;
    let theme_color = xml::optional_attr_str(e, b"w:themeColor")?;

    // The literal `w:val` doubles as the fallback for consumers that
    // cannot resolve the theme, so parse it before branching on
    // `w:themeColor` and carry it into the theme reference.
    let literal = val
        .as_deref()
        .filter(|v| v.len() == 6)
        .and_then(|v| RgbColor::from_hex(v).ok());

    if let Some(ref tc) = theme_color {
        if let Some(slot) = ThemeColorSlot::from_scheme_val(tc) {
            let tint = xml::optional_attr_str(e, b"w:themeTint")?
                .and_then(|v| u8::from_str_radix(&v, 16).ok())
                .map(|v| v as f64 / 255.0);
            let shade = xml::optional_attr_str(e, b"w:themeShade")?
                .and_then(|v| u8::from_str_radix(&v, 16).ok())
                .map(|v| v as f64 / 255.0);
            return Ok(Some(ColorRef::Theme {
                slot,
                tint,
                shade,
                fallback: literal,
            }));
        }
    }

    if let Some(ref v) = val {
        if v.as_ref() == "auto" {
            return Ok(Some(ColorRef::Auto));
        }
        if v.len() == 6 {
            if let Ok(rgb) = RgbColor::from_hex(v) {
                return Ok(Some(ColorRef::Rgb(rgb)));
            }
        }
    }
    Ok(None)
}

pub(crate) fn parse_indent(e: &BytesStart) -> crate::core::Result<ParagraphIndent> {
    let mut indent = ParagraphIndent::default();
    if let Some(val) = xml::optional_attr_str(e, b"w:left")? {
        indent.left = Some(Twip(parse_numeric(&val)?));
    }
    if indent.left.is_none() {
        if let Some(val) = xml::optional_attr_str(e, b"w:start")? {
            indent.left = Some(Twip(parse_numeric(&val)?));
        }
    }
    if let Some(val) = xml::optional_attr_str(e, b"w:right")? {
        indent.right = Some(Twip(parse_numeric(&val)?));
    }
    if indent.right.is_none() {
        if let Some(val) = xml::optional_attr_str(e, b"w:end")? {
            indent.right = Some(Twip(parse_numeric(&val)?));
        }
    }
    if let Some(val) = xml::optional_attr_str(e, b"w:firstLine")? {
        indent.first_line = Some(Twip(parse_numeric(&val)?));
    }
    if let Some(val) = xml::optional_attr_str(e, b"w:hanging")? {
        indent.hanging = Some(Twip(parse_numeric(&val)?));
    }
    Ok(indent)
}

/// Parse `<w:framePr>` attributes (`w:x`, `w:y`, `w:w`, `w:h`).
/// Returns `None` if the element doesn't carry usable absolute coords —
/// e.g. when only `wrap`/`anchor` modifiers are set without explicit
/// position/size, which we can't reproduce as positional.
fn parse_frame_pr(e: &BytesStart) -> Option<FrameProps> {
    let read_int = |attr: &[u8]| -> Option<i32> {
        xml::optional_attr_str(e, attr)
            .ok()
            .flatten()
            .and_then(|v| v.parse::<i32>().ok())
    };
    let x = read_int(b"w:x");
    let y = read_int(b"w:y");
    let w = read_int(b"w:w");
    let h = read_int(b"w:h");
    match (x, y, w, h) {
        (Some(x), Some(y), Some(w), Some(h)) => Some(FrameProps {
            x_twips: x,
            y_twips: y,
            width_twips: w,
            height_twips: h,
        }),
        _ => None,
    }
}

fn parse_spacing(e: &BytesStart) -> crate::core::Result<ParagraphSpacing> {
    let mut spacing = ParagraphSpacing::default();
    if let Some(val) = xml::optional_attr_str(e, b"w:before")? {
        spacing.before = Some(Twip(parse_numeric(&val)?));
    }
    if let Some(val) = xml::optional_attr_str(e, b"w:after")? {
        spacing.after = Some(Twip(parse_numeric(&val)?));
    }
    if let Some(val) = xml::optional_attr_str(e, b"w:line")? {
        let line_val: i32 = parse_numeric(&val)?;
        let rule = xml::optional_attr_str(e, b"w:lineRule")?.map(|r| match r.as_ref() {
            "auto" => LineSpacingRule::Auto,
            "exact" => LineSpacingRule::Exact,
            "atLeast" => LineSpacingRule::AtLeast,
            _ => LineSpacingRule::Auto,
        });
        spacing.line = Some(SpacingLine {
            value: line_val,
            rule,
        });
    }
    Ok(spacing)
}

#[cfg(test)]
fn parse_num_pr(reader: &mut quick_xml::NsReader<&[u8]>) -> crate::core::Result<NumberingRef> {
    let wml = xml::ns::WML;
    let mut num_id: u32 = 0;
    let mut ilvl: u8 = 0;

    loop {
        match reader.read_resolved_event()? {
            (ref resolve, Event::Start(ref e)) | (ref resolve, Event::Empty(ref e))
                if xml::matches_ns(resolve, wml) =>
            {
                let local = e.local_name();
                match local.as_ref() {
                    b"numId" => {
                        if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                            num_id = val.parse().unwrap_or(0);
                        }
                    },
                    b"ilvl" => {
                        if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                            ilvl = val.parse().unwrap_or(0);
                        }
                    },
                    _ => {},
                }
            },
            (ref resolve, Event::End(ref e))
                if xml::matches_ns(resolve, wml) && e.local_name().as_ref() == b"numPr" =>
            {
                break;
            },
            (_, Event::Eof) => break,
            _ => {},
        }
    }
    Ok(NumberingRef { num_id, ilvl })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_toggle_bare() {
        let xml =
            br#"<w:b xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#;
        let mut reader = xml::make_reader(xml);
        loop {
            match reader.read_resolved_event().unwrap() {
                (_, Event::Empty(ref e)) => {
                    assert!(parse_toggle(e));
                    break;
                },
                (_, Event::Eof) => panic!("unexpected eof"),
                _ => {},
            }
        }
    }

    #[test]
    fn parse_toggle_false() {
        let xml = br#"<w:b w:val="0" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#;
        let mut reader = xml::make_reader(xml);
        loop {
            match reader.read_resolved_event().unwrap() {
                (_, Event::Empty(ref e)) => {
                    assert!(!parse_toggle(e));
                    break;
                },
                (_, Event::Eof) => panic!("unexpected eof"),
                _ => {},
            }
        }
    }

    #[test]
    fn parse_run_props_bold_italic() {
        let xml =
            br#"<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:b/>
            <w:i/>
            <w:sz w:val="24"/>
        </w:rPr>"#;
        let mut reader = xml::make_reader(xml);
        // advance past <w:rPr>
        loop {
            if let (_, Event::Start(_)) = reader.read_resolved_event().unwrap() {
                break;
            }
        }
        let rp = parse_run_properties(&mut reader).unwrap();
        assert_eq!(rp.bold, Some(true));
        assert_eq!(rp.italic, Some(true));
        assert_eq!(rp.font_size, Some(HalfPoint(24)));
    }

    #[test]
    fn parse_paragraph_props_style_justification() {
        let xml =
            br#"<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:pStyle w:val="Heading1"/>
            <w:jc w:val="center"/>
            <w:outlineLvl w:val="0"/>
        </w:pPr>"#;
        let mut reader = xml::make_reader(xml);
        loop {
            if let (_, Event::Start(_)) = reader.read_resolved_event().unwrap() {
                break;
            }
        }
        let pp = parse_paragraph_properties(&mut reader).unwrap();
        assert_eq!(pp.style_id.as_deref(), Some("Heading1"));
        assert_eq!(pp.justification, Some(Justification::Center));
        assert_eq!(pp.outline_level, Some(0));
    }

    #[test]
    fn parse_indent_values() {
        let xml = br#"<w:ind w:left="720" w:hanging="360" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#;
        let mut reader = xml::make_reader(xml);
        loop {
            match reader.read_resolved_event().unwrap() {
                (_, Event::Empty(ref e)) => {
                    let indent = parse_indent(e).unwrap();
                    assert_eq!(indent.left, Some(Twip(720)));
                    assert_eq!(indent.hanging, Some(Twip(360)));
                    assert!(indent.right.is_none());
                    assert!(indent.first_line.is_none());
                    break;
                },
                (_, Event::Eof) => panic!("unexpected eof"),
                _ => {},
            }
        }
    }

    // Advance a fast reader past the opening <w:pPr> wrapper so the
    // caller can drive parse_paragraph_properties_fast directly.
    fn open_ppr_fast(xml: &[u8]) -> quick_xml::Reader<&[u8]> {
        let mut reader = xml::make_fast_reader(xml);
        loop {
            match reader.read_event().unwrap() {
                Event::Start(ref e) if e.local_name().as_ref() == b"pPr" => return reader,
                Event::Eof => panic!("no <w:pPr> in test xml"),
                _ => {},
            }
        }
    }

    // ── framePr ─────────────────────────────────────────────────────────

    #[test]
    fn parse_frame_pr_empty_element() {
        let xml =
            br#"<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:framePr w:x="720" w:y="1080" w:w="3000" w:h="500"/>
        </w:pPr>"#;
        let mut reader = open_ppr_fast(xml);
        let pp = parse_paragraph_properties_fast(&mut reader).unwrap();
        let fp = pp.frame_position.expect("framePr parsed");
        assert_eq!(fp.x_twips, 720);
        assert_eq!(fp.y_twips, 1080);
        assert_eq!(fp.width_twips, 3000);
        assert_eq!(fp.height_twips, 500);
    }

    #[test]
    fn parse_frame_pr_missing_attrs_returns_none() {
        // Missing w:h → frame_position must be None.
        let xml =
            br#"<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:framePr w:x="100" w:y="200" w:w="300"/>
        </w:pPr>"#;
        let mut reader = open_ppr_fast(xml);
        let pp = parse_paragraph_properties_fast(&mut reader).unwrap();
        assert!(pp.frame_position.is_none());
    }

    #[test]
    fn parse_frame_pr_inside_start_form() {
        // Start/End form (rather than Empty) — should still parse.
        let xml =
            br#"<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:framePr w:x="10" w:y="20" w:w="30" w:h="40"></w:framePr>
        </w:pPr>"#;
        let mut reader = open_ppr_fast(xml);
        let pp = parse_paragraph_properties_fast(&mut reader).unwrap();
        let fp = pp.frame_position.expect("framePr parsed");
        assert_eq!(fp.x_twips, 10);
        assert_eq!(fp.width_twips, 30);
    }

    // ── pBdr / has_bottom_border ────────────────────────────────────────

    #[test]
    fn parse_p_bdr_with_bottom() {
        let xml =
            br#"<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:pBdr>
            <w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/>
          </w:pBdr>
        </w:pPr>"#;
        let mut reader = open_ppr_fast(xml);
        let pp = parse_paragraph_properties_fast(&mut reader).unwrap();
        assert!(pp.has_bottom_border);
    }

    #[test]
    fn parse_p_bdr_without_bottom() {
        let xml =
            br#"<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:pBdr>
            <w:top w:val="single" w:sz="6"/>
            <w:left w:val="single" w:sz="6"/>
          </w:pBdr>
        </w:pPr>"#;
        let mut reader = open_ppr_fast(xml);
        let pp = parse_paragraph_properties_fast(&mut reader).unwrap();
        assert!(!pp.has_bottom_border);
    }

    #[test]
    fn paragraph_properties_default_has_no_frame_or_border() {
        let xml =
            br#"<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:pStyle w:val="Normal"/>
        </w:pPr>"#;
        let mut reader = open_ppr_fast(xml);
        let pp = parse_paragraph_properties_fast(&mut reader).unwrap();
        assert!(pp.frame_position.is_none());
        assert!(!pp.has_bottom_border);
    }
}
