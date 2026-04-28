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
                    _ => {
                        xml::skip_element_fast(reader)?;
                    },
                }
            },
            Event::Empty(ref e) => {
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
                    _ => {
                        xml::skip_element_fast(reader)?;
                    },
                }
            },
            Event::Empty(ref e) => {
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
            Event::End(ref e) if e.local_name().as_ref() == b"pPr" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(props)
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

    if let Some(ref tc) = theme_color {
        if let Some(slot) = ThemeColorSlot::from_scheme_val(tc) {
            let tint = xml::optional_attr_str(e, b"w:themeTint")?
                .and_then(|v| u8::from_str_radix(&v, 16).ok())
                .map(|v| v as f64 / 255.0);
            let shade = xml::optional_attr_str(e, b"w:themeShade")?
                .and_then(|v| u8::from_str_radix(&v, 16).ok())
                .map(|v| v as f64 / 255.0);
            return Ok(Some(ColorRef::Theme { slot, tint, shade }));
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
}
