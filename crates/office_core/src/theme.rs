use std::collections::HashMap;

use quick_xml::events::Event;

use crate::error::{Error, Result};
use crate::xml;

/// The 12 named theme color slots in DrawingML.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ThemeColorSlot {
    Dk1,
    Lt1,
    Dk2,
    Lt2,
    Accent1,
    Accent2,
    Accent3,
    Accent4,
    Accent5,
    Accent6,
    Hlink,
    FolHlink,
}

impl ThemeColorSlot {
    /// Parse from a scheme color value string (e.g. "accent1", "dk1", "tx1").
    pub fn from_scheme_val(s: &str) -> Option<Self> {
        match s {
            "dk1" | "tx1" => Some(Self::Dk1),
            "lt1" | "bg1" => Some(Self::Lt1),
            "dk2" | "tx2" => Some(Self::Dk2),
            "lt2" | "bg2" => Some(Self::Lt2),
            "accent1" => Some(Self::Accent1),
            "accent2" => Some(Self::Accent2),
            "accent3" => Some(Self::Accent3),
            "accent4" => Some(Self::Accent4),
            "accent5" => Some(Self::Accent5),
            "accent6" => Some(Self::Accent6),
            "hlink" => Some(Self::Hlink),
            "folHlink" => Some(Self::FolHlink),
            _ => None,
        }
    }

    fn from_element_name(name: &[u8]) -> Option<Self> {
        match name {
            b"dk1" => Some(Self::Dk1),
            b"lt1" => Some(Self::Lt1),
            b"dk2" => Some(Self::Dk2),
            b"lt2" => Some(Self::Lt2),
            b"accent1" => Some(Self::Accent1),
            b"accent2" => Some(Self::Accent2),
            b"accent3" => Some(Self::Accent3),
            b"accent4" => Some(Self::Accent4),
            b"accent5" => Some(Self::Accent5),
            b"accent6" => Some(Self::Accent6),
            b"hlink" => Some(Self::Hlink),
            b"folHlink" => Some(Self::FolHlink),
            _ => None,
        }
    }
}

/// An sRGB color as 3 bytes (red, green, blue).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RgbColor(pub [u8; 3]);

impl RgbColor {
    pub fn from_hex(hex: &str) -> Result<Self> {
        if hex.len() != 6 {
            return Err(Error::MalformedXml(format!(
                "invalid hex color (expected 6 chars): {hex}"
            )));
        }
        let r = u8::from_str_radix(&hex[0..2], 16)?;
        let g = u8::from_str_radix(&hex[2..4], 16)?;
        let b = u8::from_str_radix(&hex[4..6], 16)?;
        Ok(Self([r, g, b]))
    }

    pub fn to_hex(&self) -> String {
        format!("{:02X}{:02X}{:02X}", self.0[0], self.0[1], self.0[2])
    }

    pub fn red(&self) -> u8 {
        self.0[0]
    }

    pub fn green(&self) -> u8 {
        self.0[1]
    }

    pub fn blue(&self) -> u8 {
        self.0[2]
    }
}

/// Color scheme: the 12 named colors in a theme.
#[derive(Debug, Clone)]
pub struct ColorScheme {
    pub name: String,
    pub colors: HashMap<ThemeColorSlot, RgbColor>,
}

/// Font scheme: major (heading) and minor (body) fonts.
#[derive(Debug, Clone)]
pub struct FontScheme {
    pub name: String,
    pub major_latin: String,
    pub minor_latin: String,
    pub major_ea: Option<String>,
    pub minor_ea: Option<String>,
    pub major_cs: Option<String>,
    pub minor_cs: Option<String>,
}

/// A parsed DrawingML theme (`a:theme`).
#[derive(Debug, Clone)]
pub struct Theme {
    pub name: String,
    pub color_scheme: ColorScheme,
    pub font_scheme: FontScheme,
}

impl Theme {
    /// Parse a DrawingML theme from XML bytes.
    pub fn parse(xml_data: &[u8]) -> Result<Self> {
        let mut reader = xml::make_reader(xml_data);
        let dml = xml::ns::DRAWING_ML;

        let mut theme_name = String::new();
        let mut color_scheme = None;
        let mut font_scheme = None;

        loop {
            match reader.read_resolved_event()? {
                (ref resolve, Event::Start(ref e)) => {
                    let local = e.local_name();
                    let local_bytes = local.as_ref();

                    if xml::matches_ns(resolve, dml) {
                        match local_bytes {
                            b"theme" => {
                                if let Ok(Some(name)) = xml::optional_attr_str(e, b"name") {
                                    theme_name = name.into_owned();
                                }
                            }
                            b"clrScheme" => {
                                color_scheme =
                                    Some(parse_color_scheme(&mut reader, e)?);
                            }
                            b"fontScheme" => {
                                font_scheme =
                                    Some(parse_font_scheme(&mut reader, e)?);
                            }
                            _ => {}
                        }
                    }
                }
                (_, Event::Eof) => break,
                _ => {}
            }
        }

        let color_scheme = color_scheme.ok_or_else(|| {
            Error::MalformedXml("theme missing clrScheme".to_string())
        })?;
        let font_scheme = font_scheme.ok_or_else(|| {
            Error::MalformedXml("theme missing fontScheme".to_string())
        })?;

        Ok(Theme {
            name: theme_name,
            color_scheme,
            font_scheme,
        })
    }

    /// Resolve a theme color slot to its RGB value.
    pub fn resolve_color(&self, slot: ThemeColorSlot) -> Option<&RgbColor> {
        self.color_scheme.colors.get(&slot)
    }
}

/// Parse `<a:clrScheme>` element and its children.
fn parse_color_scheme(
    reader: &mut quick_xml::NsReader<&[u8]>,
    start: &quick_xml::events::BytesStart,
) -> Result<ColorScheme> {
    let name = xml::optional_attr_str(start, b"name")?
        .map(|c| c.into_owned())
        .unwrap_or_default();
    let dml = xml::ns::DRAWING_ML;
    let mut colors = HashMap::new();
    let mut current_slot: Option<ThemeColorSlot> = None;

    loop {
        match reader.read_resolved_event()? {
            (ref resolve, Event::Start(ref e)) | (ref resolve, Event::Empty(ref e)) => {
                let local = e.local_name();
                let local_bytes = local.as_ref();

                if xml::matches_ns(resolve, dml) {
                    // Check if this is a color slot element (dk1, lt1, accent1, etc.)
                    if let Some(slot) = ThemeColorSlot::from_element_name(local_bytes) {
                        current_slot = Some(slot);
                    } else if local_bytes == b"srgbClr" {
                        // <a:srgbClr val="4472C4"/>
                        if let Some(slot) = current_slot {
                            if let Ok(val) = xml::required_attr_str(e, b"val") {
                                if let Ok(rgb) = RgbColor::from_hex(&val) {
                                    colors.insert(slot, rgb);
                                }
                            }
                        }
                    } else if local_bytes == b"sysClr" {
                        // <a:sysClr val="windowText" lastClr="000000"/>
                        if let Some(slot) = current_slot {
                            // Use lastClr for the actual color value
                            if let Ok(Some(last_clr)) = xml::optional_attr_str(e, b"lastClr")
                            {
                                if let Ok(rgb) = RgbColor::from_hex(&last_clr) {
                                    colors.insert(slot, rgb);
                                }
                            }
                        }
                    }
                }
            }
            (ref resolve, Event::End(ref e)) => {
                let local = e.local_name();
                let local_bytes = local.as_ref();

                if xml::matches_ns(resolve, dml) && local_bytes == b"clrScheme" {
                    break;
                }
                // Reset current_slot when we leave a color slot element
                if xml::matches_ns(resolve, dml)
                    && ThemeColorSlot::from_element_name(local_bytes).is_some()
                {
                    current_slot = None;
                }
            }
            (_, Event::Eof) => break,
            _ => {}
        }
    }

    Ok(ColorScheme { name, colors })
}

/// Parse `<a:fontScheme>` element and its children.
fn parse_font_scheme(
    reader: &mut quick_xml::NsReader<&[u8]>,
    start: &quick_xml::events::BytesStart,
) -> Result<FontScheme> {
    let name = xml::optional_attr_str(start, b"name")?
        .map(|c| c.into_owned())
        .unwrap_or_default();
    let dml = xml::ns::DRAWING_ML;

    let mut major_latin = String::new();
    let mut minor_latin = String::new();
    let mut major_ea: Option<String> = None;
    let mut minor_ea: Option<String> = None;
    let mut major_cs: Option<String> = None;
    let mut minor_cs: Option<String> = None;

    enum FontCtx {
        None,
        Major,
        Minor,
    }
    let mut ctx = FontCtx::None;

    loop {
        match reader.read_resolved_event()? {
            (ref resolve, Event::Start(ref e)) | (ref resolve, Event::Empty(ref e)) => {
                let local = e.local_name();
                let local_bytes = local.as_ref();

                if xml::matches_ns(resolve, dml) {
                    match local_bytes {
                        b"majorFont" => ctx = FontCtx::Major,
                        b"minorFont" => ctx = FontCtx::Minor,
                        b"latin" => {
                            if let Ok(Some(tf)) = xml::optional_attr_str(e, b"typeface") {
                                let typeface = tf.into_owned();
                                match ctx {
                                    FontCtx::Major => major_latin = typeface,
                                    FontCtx::Minor => minor_latin = typeface,
                                    FontCtx::None => {}
                                }
                            }
                        }
                        b"ea" => {
                            if let Ok(Some(tf)) = xml::optional_attr_str(e, b"typeface") {
                                let typeface = tf.into_owned();
                                if !typeface.is_empty() {
                                    match ctx {
                                        FontCtx::Major => major_ea = Some(typeface),
                                        FontCtx::Minor => minor_ea = Some(typeface),
                                        FontCtx::None => {}
                                    }
                                }
                            }
                        }
                        b"cs" => {
                            if let Ok(Some(tf)) = xml::optional_attr_str(e, b"typeface") {
                                let typeface = tf.into_owned();
                                if !typeface.is_empty() {
                                    match ctx {
                                        FontCtx::Major => major_cs = Some(typeface),
                                        FontCtx::Minor => minor_cs = Some(typeface),
                                        FontCtx::None => {}
                                    }
                                }
                            }
                        }
                        _ => {}
                    }
                }
            }
            (ref resolve, Event::End(ref e)) => {
                let local = e.local_name();
                let local_bytes = local.as_ref();

                if xml::matches_ns(resolve, dml) {
                    match local_bytes {
                        b"fontScheme" => break,
                        b"majorFont" | b"minorFont" => ctx = FontCtx::None,
                        _ => {}
                    }
                }
            }
            (_, Event::Eof) => break,
            _ => {}
        }
    }

    Ok(FontScheme {
        name,
        major_latin,
        minor_latin,
        major_ea,
        minor_ea,
        major_cs,
        minor_cs,
    })
}

// ---------------------------------------------------------------------------
// ColorRef — used by format crates for color resolution
// ---------------------------------------------------------------------------

/// A color reference that may need theme resolution.
#[derive(Debug, Clone)]
pub enum ColorRef {
    /// Direct sRGB color.
    Rgb(RgbColor),
    /// Theme color with optional tint/shade transforms.
    Theme {
        slot: ThemeColorSlot,
        tint: Option<f64>,
        shade: Option<f64>,
    },
    /// System color (e.g., "windowText", "window").
    System(String),
    /// Automatic color.
    Auto,
}

impl ColorRef {
    /// Resolve this color reference to an RGB value using the given theme.
    pub fn resolve(&self, theme: &Theme) -> RgbColor {
        match self {
            Self::Rgb(rgb) => rgb.clone(),
            Self::Theme { slot, tint, shade } => {
                let base = theme
                    .resolve_color(*slot)
                    .cloned()
                    .unwrap_or(RgbColor([0, 0, 0]));
                apply_tint_shade(&base, *tint, *shade)
            }
            Self::System(name) => {
                // Map common system colors to reasonable defaults
                match name.as_str() {
                    "windowText" => RgbColor([0, 0, 0]),
                    "window" => RgbColor([255, 255, 255]),
                    "highlightText" => RgbColor([255, 255, 255]),
                    "highlight" => RgbColor([0, 120, 215]),
                    _ => RgbColor([0, 0, 0]),
                }
            }
            Self::Auto => RgbColor([0, 0, 0]),
        }
    }
}

/// Apply tint and shade transformations to a base color.
fn apply_tint_shade(base: &RgbColor, tint: Option<f64>, shade: Option<f64>) -> RgbColor {
    let mut r = base.0[0] as f64;
    let mut g = base.0[1] as f64;
    let mut b = base.0[2] as f64;

    // Apply shade (darken): component = component * shade
    if let Some(shade_val) = shade {
        r *= shade_val;
        g *= shade_val;
        b *= shade_val;
    }

    // Apply tint (lighten): component = component + (255 - component) * tint
    if let Some(tint_val) = tint {
        r = r + (255.0 - r) * tint_val;
        g = g + (255.0 - g) * tint_val;
        b = b + (255.0 - b) * tint_val;
    }

    RgbColor([
        r.round().clamp(0.0, 255.0) as u8,
        g.round().clamp(0.0, 255.0) as u8,
        b.round().clamp(0.0, 255.0) as u8,
    ])
}

#[cfg(test)]
mod tests {
    use super::*;

    const SAMPLE_THEME: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Calibri Light"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>"#;

    #[test]
    fn parse_theme() {
        let theme = Theme::parse(SAMPLE_THEME).unwrap();
        assert_eq!(theme.name, "Office Theme");
        assert_eq!(theme.color_scheme.name, "Office");
        assert_eq!(theme.font_scheme.name, "Office");
    }

    #[test]
    fn parse_theme_colors() {
        let theme = Theme::parse(SAMPLE_THEME).unwrap();
        let cs = &theme.color_scheme;

        // System colors (via lastClr)
        assert_eq!(cs.colors.get(&ThemeColorSlot::Dk1), Some(&RgbColor([0, 0, 0])));
        assert_eq!(
            cs.colors.get(&ThemeColorSlot::Lt1),
            Some(&RgbColor([255, 255, 255]))
        );

        // sRGB colors
        assert_eq!(
            cs.colors.get(&ThemeColorSlot::Accent1),
            Some(&RgbColor([0x44, 0x72, 0xC4]))
        );
        assert_eq!(
            cs.colors.get(&ThemeColorSlot::Hlink),
            Some(&RgbColor([0x05, 0x63, 0xC1]))
        );

        // All 12 slots present
        assert_eq!(cs.colors.len(), 12);
    }

    #[test]
    fn parse_theme_fonts() {
        let theme = Theme::parse(SAMPLE_THEME).unwrap();
        assert_eq!(theme.font_scheme.major_latin, "Calibri Light");
        assert_eq!(theme.font_scheme.minor_latin, "Calibri");
        assert!(theme.font_scheme.major_ea.is_none()); // empty string = None
        assert!(theme.font_scheme.minor_ea.is_none());
    }

    #[test]
    fn color_ref_resolve_rgb() {
        let theme = Theme::parse(SAMPLE_THEME).unwrap();
        let color = ColorRef::Rgb(RgbColor([128, 64, 32]));
        assert_eq!(color.resolve(&theme), RgbColor([128, 64, 32]));
    }

    #[test]
    fn color_ref_resolve_theme() {
        let theme = Theme::parse(SAMPLE_THEME).unwrap();
        let color = ColorRef::Theme {
            slot: ThemeColorSlot::Accent1,
            tint: None,
            shade: None,
        };
        assert_eq!(color.resolve(&theme), RgbColor([0x44, 0x72, 0xC4]));
    }

    #[test]
    fn color_ref_resolve_theme_with_tint() {
        let theme = Theme::parse(SAMPLE_THEME).unwrap();
        let color = ColorRef::Theme {
            slot: ThemeColorSlot::Dk1,
            tint: Some(0.5),
            shade: None,
        };
        let resolved = color.resolve(&theme);
        // Black (0,0,0) with 50% tint = (128, 128, 128) approximately
        assert_eq!(resolved, RgbColor([128, 128, 128]));
    }

    #[test]
    fn rgb_hex_round_trip() {
        let rgb = RgbColor::from_hex("4472C4").unwrap();
        assert_eq!(rgb.to_hex(), "4472C4");
        assert_eq!(rgb.red(), 0x44);
        assert_eq!(rgb.green(), 0x72);
        assert_eq!(rgb.blue(), 0xC4);
    }

    #[test]
    fn theme_color_slot_from_scheme_val() {
        assert_eq!(
            ThemeColorSlot::from_scheme_val("accent1"),
            Some(ThemeColorSlot::Accent1)
        );
        assert_eq!(
            ThemeColorSlot::from_scheme_val("dk1"),
            Some(ThemeColorSlot::Dk1)
        );
        // tx1 maps to dk1 (text1 = dark1)
        assert_eq!(
            ThemeColorSlot::from_scheme_val("tx1"),
            Some(ThemeColorSlot::Dk1)
        );
        // bg1 maps to lt1 (background1 = light1)
        assert_eq!(
            ThemeColorSlot::from_scheme_val("bg1"),
            Some(ThemeColorSlot::Lt1)
        );
    }
}
