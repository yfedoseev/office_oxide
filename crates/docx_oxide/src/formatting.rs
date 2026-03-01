use office_core::theme::ColorRef;
use office_core::units::{HalfPoint, Twip};

/// Run-level formatting properties (`w:rPr`).
#[derive(Debug, Clone, Default)]
pub struct RunProperties {
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<UnderlineType>,
    pub strike: Option<bool>,
    pub dstrike: Option<bool>,
    pub font_size: Option<HalfPoint>,
    pub font_name: Option<String>,
    pub color: Option<ColorRef>,
    pub highlight: Option<String>,
    pub vertical_align: Option<VerticalAlign>,
    pub style_id: Option<String>,
}

/// Paragraph-level formatting properties (`w:pPr`).
#[derive(Debug, Clone, Default)]
pub struct ParagraphProperties {
    pub style_id: Option<String>,
    pub justification: Option<Justification>,
    pub indent: Option<ParagraphIndent>,
    pub spacing: Option<ParagraphSpacing>,
    pub numbering_ref: Option<NumberingRef>,
    pub outline_level: Option<u8>,
    pub run_properties: Option<RunProperties>,
}

/// Underline style.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum UnderlineType {
    Single,
    Double,
    Thick,
    Dotted,
    Dash,
    DotDash,
    DotDotDash,
    Wave,
    Words,
    None,
    Other(String),
}

/// Vertical alignment (superscript/subscript).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VerticalAlign {
    Superscript,
    Subscript,
    Baseline,
}

/// Paragraph justification / alignment.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Justification {
    Left,
    Center,
    Right,
    Both,
    Distribute,
}

/// Paragraph indentation values.
#[derive(Debug, Clone, Default)]
pub struct ParagraphIndent {
    pub left: Option<Twip>,
    pub right: Option<Twip>,
    pub first_line: Option<Twip>,
    pub hanging: Option<Twip>,
}

/// Paragraph spacing.
#[derive(Debug, Clone, Default)]
pub struct ParagraphSpacing {
    pub before: Option<Twip>,
    pub after: Option<Twip>,
    pub line: Option<SpacingLine>,
}

/// Line spacing rule and value.
#[derive(Debug, Clone)]
pub struct SpacingLine {
    pub value: i32,
    pub rule: Option<LineSpacingRule>,
}

/// Line spacing rule.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LineSpacingRule {
    Auto,
    Exact,
    AtLeast,
}

/// Reference to a numbering definition from a paragraph.
#[derive(Debug, Clone)]
pub struct NumberingRef {
    pub num_id: u32,
    pub ilvl: u8,
}

// ---------------------------------------------------------------------------
// Parsing helpers
// ---------------------------------------------------------------------------

use quick_xml::events::{BytesStart, Event};
use quick_xml::NsReader;

use office_core::xml;

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

/// Parse `w:rPr` element children. Caller has already consumed the `Start(w:rPr)` event.
pub(crate) fn parse_run_properties(reader: &mut NsReader<&[u8]>) -> office_core::Result<RunProperties> {
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
                        }
                        b"i" => {
                            props.italic = Some(parse_toggle(e));
                            xml::skip_element(reader)?;
                        }
                        b"strike" => {
                            props.strike = Some(parse_toggle(e));
                            xml::skip_element(reader)?;
                        }
                        b"dstrike" => {
                            props.dstrike = Some(parse_toggle(e));
                            xml::skip_element(reader)?;
                        }
                        b"u" => {
                            props.underline = Some(parse_underline(e));
                            xml::skip_element(reader)?;
                        }
                        b"sz" => {
                            if let Some(val) = parse_half_point_val(e)? {
                                props.font_size = Some(val);
                            }
                            xml::skip_element(reader)?;
                        }
                        b"rFonts" => {
                            if let Ok(Some(ascii)) = xml::optional_attr_str(e, b"w:ascii") {
                                props.font_name = Some(ascii.into_owned());
                            }
                            xml::skip_element(reader)?;
                        }
                        b"color" => {
                            props.color = parse_color_ref(e)?;
                            xml::skip_element(reader)?;
                        }
                        b"highlight" => {
                            if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                                props.highlight = Some(val.into_owned());
                            }
                            xml::skip_element(reader)?;
                        }
                        b"vertAlign" => {
                            if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                                props.vertical_align = Some(match val.as_ref() {
                                    "superscript" => VerticalAlign::Superscript,
                                    "subscript" => VerticalAlign::Subscript,
                                    _ => VerticalAlign::Baseline,
                                });
                            }
                            xml::skip_element(reader)?;
                        }
                        b"rStyle" => {
                            if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                                props.style_id = Some(val.into_owned());
                            }
                            xml::skip_element(reader)?;
                        }
                        _ => {
                            xml::skip_element(reader)?;
                        }
                    }
                } else {
                    xml::skip_element(reader)?;
                }
            }
            (ref resolve, Event::Empty(ref e)) => {
                if xml::matches_ns(resolve, wml) {
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
                        }
                        b"rFonts" => {
                            if let Ok(Some(ascii)) = xml::optional_attr_str(e, b"w:ascii") {
                                props.font_name = Some(ascii.into_owned());
                            }
                        }
                        b"color" => {
                            props.color = parse_color_ref(e)?;
                        }
                        b"highlight" => {
                            if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                                props.highlight = Some(val.into_owned());
                            }
                        }
                        b"vertAlign" => {
                            if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                                props.vertical_align = Some(match val.as_ref() {
                                    "superscript" => VerticalAlign::Superscript,
                                    "subscript" => VerticalAlign::Subscript,
                                    _ => VerticalAlign::Baseline,
                                });
                            }
                        }
                        b"rStyle" => {
                            if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                                props.style_id = Some(val.into_owned());
                            }
                        }
                        _ => {}
                    }
                }
            }
            (ref resolve, Event::End(ref e)) => {
                if xml::matches_ns(resolve, wml) && e.local_name().as_ref() == b"rPr" {
                    break;
                }
            }
            (_, Event::Eof) => break,
            _ => {}
        }
    }
    Ok(props)
}

/// Parse `w:pPr` element children. Caller has already consumed the `Start(w:pPr)` event.
pub(crate) fn parse_paragraph_properties(
    reader: &mut NsReader<&[u8]>,
) -> office_core::Result<ParagraphProperties> {
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
                        }
                        b"jc" => {
                            if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                                props.justification = Some(parse_justification_value(&val));
                            }
                            xml::skip_element(reader)?;
                        }
                        b"ind" => {
                            props.indent = Some(parse_indent(e)?);
                            xml::skip_element(reader)?;
                        }
                        b"spacing" => {
                            props.spacing = Some(parse_spacing(e)?);
                            xml::skip_element(reader)?;
                        }
                        b"numPr" => {
                            props.numbering_ref = Some(parse_num_pr(reader)?);
                        }
                        b"outlineLvl" => {
                            if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                                if let Ok(lvl) = val.parse::<u8>() {
                                    props.outline_level = Some(lvl);
                                }
                            }
                            xml::skip_element(reader)?;
                        }
                        b"rPr" => {
                            props.run_properties = Some(parse_run_properties(reader)?);
                        }
                        _ => {
                            xml::skip_element(reader)?;
                        }
                    }
                } else {
                    xml::skip_element(reader)?;
                }
            }
            (ref resolve, Event::Empty(ref e)) => {
                if xml::matches_ns(resolve, wml) {
                    let local = e.local_name();
                    match local.as_ref() {
                        b"pStyle" => {
                            if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                                props.style_id = Some(val.into_owned());
                            }
                        }
                        b"jc" => {
                            if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                                props.justification = Some(parse_justification_value(&val));
                            }
                        }
                        b"ind" => {
                            props.indent = Some(parse_indent(e)?);
                        }
                        b"spacing" => {
                            props.spacing = Some(parse_spacing(e)?);
                        }
                        b"outlineLvl" => {
                            if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                                if let Ok(lvl) = val.parse::<u8>() {
                                    props.outline_level = Some(lvl);
                                }
                            }
                        }
                        _ => {}
                    }
                }
            }
            (ref resolve, Event::End(ref e)) => {
                if xml::matches_ns(resolve, wml) && e.local_name().as_ref() == b"pPr" {
                    break;
                }
            }
            (_, Event::Eof) => break,
            _ => {}
        }
    }
    Ok(props)
}

/// Parse a boolean toggle attribute. `<w:b/>` = true, `<w:b w:val="0"/>` = false.
fn parse_toggle(e: &BytesStart) -> bool {
    match xml::optional_attr_str(e, b"w:val") {
        Ok(Some(ref val)) => !matches!(val.as_ref(), "0" | "false" | "off"),
        _ => true, // bare element = true
    }
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

fn parse_half_point_val(e: &BytesStart) -> office_core::Result<Option<HalfPoint>> {
    match xml::optional_attr_str(e, b"w:val")? {
        Some(ref val) => {
            let v: u32 = parse_numeric(val)?;
            Ok(Some(HalfPoint(v)))
        }
        None => Ok(None),
    }
}

fn parse_color_ref(e: &BytesStart) -> office_core::Result<Option<ColorRef>> {
    use office_core::theme::{RgbColor, ThemeColorSlot};

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

pub(crate) fn parse_indent(e: &BytesStart) -> office_core::Result<ParagraphIndent> {
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

fn parse_spacing(e: &BytesStart) -> office_core::Result<ParagraphSpacing> {
    let mut spacing = ParagraphSpacing::default();
    if let Some(val) = xml::optional_attr_str(e, b"w:before")? {
        spacing.before = Some(Twip(parse_numeric(&val)?));
    }
    if let Some(val) = xml::optional_attr_str(e, b"w:after")? {
        spacing.after = Some(Twip(parse_numeric(&val)?));
    }
    if let Some(val) = xml::optional_attr_str(e, b"w:line")? {
        let line_val: i32 = parse_numeric(&val)?;
        let rule = xml::optional_attr_str(e, b"w:lineRule")?
            .map(|r| match r.as_ref() {
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

fn parse_num_pr(reader: &mut NsReader<&[u8]>) -> office_core::Result<NumberingRef> {
    let wml = xml::ns::WML;
    let mut num_id: u32 = 0;
    let mut ilvl: u8 = 0;

    loop {
        match reader.read_resolved_event()? {
            (ref resolve, Event::Start(ref e)) | (ref resolve, Event::Empty(ref e)) => {
                if xml::matches_ns(resolve, wml) {
                    let local = e.local_name();
                    match local.as_ref() {
                        b"numId" => {
                            if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                                num_id = val.parse().unwrap_or(0);
                            }
                        }
                        b"ilvl" => {
                            if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                                ilvl = val.parse().unwrap_or(0);
                            }
                        }
                        _ => {}
                    }
                }
            }
            (ref resolve, Event::End(ref e)) => {
                if xml::matches_ns(resolve, wml) && e.local_name().as_ref() == b"numPr" {
                    break;
                }
            }
            (_, Event::Eof) => break,
            _ => {}
        }
    }
    Ok(NumberingRef { num_id, ilvl })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_toggle_bare() {
        let xml = br#"<w:b xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#;
        let mut reader = xml::make_reader(xml);
        loop {
            match reader.read_resolved_event().unwrap() {
                (_, Event::Empty(ref e)) => {
                    assert!(parse_toggle(e));
                    break;
                }
                (_, Event::Eof) => panic!("unexpected eof"),
                _ => {}
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
                }
                (_, Event::Eof) => panic!("unexpected eof"),
                _ => {}
            }
        }
    }

    #[test]
    fn parse_run_props_bold_italic() {
        let xml = br#"<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:b/>
            <w:i/>
            <w:sz w:val="24"/>
        </w:rPr>"#;
        let mut reader = xml::make_reader(xml);
        // advance past <w:rPr>
        loop {
            match reader.read_resolved_event().unwrap() {
                (_, Event::Start(_)) => break,
                _ => {}
            }
        }
        let rp = parse_run_properties(&mut reader).unwrap();
        assert_eq!(rp.bold, Some(true));
        assert_eq!(rp.italic, Some(true));
        assert_eq!(rp.font_size, Some(HalfPoint(24)));
    }

    #[test]
    fn parse_paragraph_props_style_justification() {
        let xml = br#"<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:pStyle w:val="Heading1"/>
            <w:jc w:val="center"/>
            <w:outlineLvl w:val="0"/>
        </w:pPr>"#;
        let mut reader = xml::make_reader(xml);
        loop {
            match reader.read_resolved_event().unwrap() {
                (_, Event::Start(_)) => break,
                _ => {}
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
                }
                (_, Event::Eof) => panic!("unexpected eof"),
                _ => {}
            }
        }
    }
}
