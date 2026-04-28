use quick_xml::events::Event;

use crate::core::theme::ColorRef;
use crate::core::xml;

use super::shared_strings::parse_color_ref;

/// Parsed stylesheet from `xl/styles.xml`.
#[derive(Debug, Clone)]
pub struct StyleSheet {
    /// Custom number formats (IDs ≥ 164).
    pub number_formats: Vec<NumberFormat>,
    /// Font definitions.
    pub fonts: Vec<Font>,
    /// Fill definitions.
    pub fills: Vec<Fill>,
    /// Border definitions.
    pub borders: Vec<Border>,
    /// Cell formats (the `cellXfs` array — cells reference by index via `s` attribute).
    pub cell_formats: Vec<CellFormat>,
    /// Cell style formats (`cellStyleXfs` array).
    pub cell_style_formats: Vec<CellFormat>,
}

/// A custom number format (ID >= 164).
#[derive(Debug, Clone)]
pub struct NumberFormat {
    /// Format ID (used by `CellFormat.number_format_id`).
    pub id: u32,
    /// Excel format code string (e.g., `"#,##0.00"`).
    pub format_code: String,
}

/// A font definition.
#[derive(Debug, Clone)]
pub struct Font {
    /// Bold toggle.
    pub bold: bool,
    /// Italic toggle.
    pub italic: bool,
    /// Underline style value.
    pub underline: Option<String>,
    /// Strikethrough toggle.
    pub strike: bool,
    /// Font size in points.
    pub size: Option<f64>,
    /// Font family name.
    pub name: Option<String>,
    /// Font color reference.
    pub color: Option<ColorRef>,
}

/// A fill definition.
#[derive(Debug, Clone)]
pub struct Fill {
    /// Fill pattern type (e.g., `"solid"`).
    pub pattern_type: Option<String>,
    /// Foreground color.
    pub fg_color: Option<ColorRef>,
    /// Background color.
    pub bg_color: Option<ColorRef>,
}

/// A border definition.
#[derive(Debug, Clone)]
pub struct Border {
    /// Left border.
    pub left: Option<BorderSide>,
    /// Right border.
    pub right: Option<BorderSide>,
    /// Top border.
    pub top: Option<BorderSide>,
    /// Bottom border.
    pub bottom: Option<BorderSide>,
}

/// A single border side.
#[derive(Debug, Clone)]
pub struct BorderSide {
    /// Border style (e.g., `"thin"`, `"medium"`).
    pub style: String,
    /// Border color.
    pub color: Option<ColorRef>,
}

/// A cell format entry from `cellXfs` or `cellStyleXfs`.
#[derive(Debug, Clone)]
pub struct CellFormat {
    /// Index into the number format array (or built-in format ID).
    pub number_format_id: u32,
    /// Index into the fonts array.
    pub font_index: Option<u32>,
    /// Index into the fills array.
    pub fill_index: Option<u32>,
    /// Index into the borders array.
    pub border_index: Option<u32>,
    /// Whether the number format is explicitly applied.
    pub apply_number_format: bool,
    /// Reference to a `cellStyleXfs` entry.
    pub xf_id: Option<u32>,
}

impl StyleSheet {
    /// Parse `xl/styles.xml` from raw XML bytes.
    pub fn parse(xml_data: &[u8]) -> crate::core::Result<Self> {
        let mut reader = xml::make_fast_reader(xml_data);

        let mut number_formats = Vec::new();
        let mut fonts = Vec::new();
        let mut fills = Vec::new();
        let mut borders = Vec::new();
        let mut cell_formats = Vec::new();
        let mut cell_style_formats = Vec::new();

        loop {
            match reader.read_event()? {
                Event::Start(ref e) => match e.local_name().as_ref() {
                    b"numFmts" => {
                        number_formats = parse_num_fmts(&mut reader)?;
                    },
                    b"fonts" => {
                        fonts = parse_fonts(&mut reader)?;
                    },
                    b"fills" => {
                        fills = parse_fills(&mut reader)?;
                    },
                    b"borders" => {
                        borders = parse_borders(&mut reader)?;
                    },
                    b"cellXfs" => {
                        cell_formats = parse_xfs(&mut reader)?;
                    },
                    b"cellStyleXfs" => {
                        cell_style_formats = parse_xfs(&mut reader)?;
                    },
                    _ => {},
                },
                Event::Eof => break,
                _ => {},
            }
        }

        Ok(StyleSheet {
            number_formats,
            fonts,
            fills,
            borders,
            cell_formats,
            cell_style_formats,
        })
    }

    /// Get the number format string for a cell format index.
    pub fn number_format_for(&self, style_index: u32) -> Option<&str> {
        let xf = self.cell_formats.get(style_index as usize)?;
        let fmt_id = xf.number_format_id;
        self.number_formats
            .iter()
            .find(|nf| nf.id == fmt_id)
            .map(|nf| nf.format_code.as_str())
    }

    /// Get the font for a cell format index.
    pub fn font_for(&self, style_index: u32) -> Option<&Font> {
        let xf = self.cell_formats.get(style_index as usize)?;
        let font_idx = xf.font_index?;
        self.fonts.get(font_idx as usize)
    }

    /// Get the number format ID for a cell format index.
    pub fn number_format_id_for(&self, style_index: u32) -> Option<u32> {
        self.cell_formats
            .get(style_index as usize)
            .map(|xf| xf.number_format_id)
    }
}

/// Parse `<numFmts>` — custom number formats.
fn parse_num_fmts(reader: &mut quick_xml::Reader<&[u8]>) -> crate::core::Result<Vec<NumberFormat>> {
    let mut formats = Vec::new();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) | Event::Empty(ref e) if e.local_name().as_ref() == b"numFmt" => {
                let id: u32 = xml::required_attr_str(e, b"numFmtId")?.parse()?;
                let format_code = xml::required_attr_str(e, b"formatCode")?.into_owned();
                formats.push(NumberFormat { id, format_code });
            },
            Event::End(ref e) if e.local_name().as_ref() == b"numFmts" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(formats)
}

/// Parse `<fonts>` collection.
fn parse_fonts(reader: &mut quick_xml::Reader<&[u8]>) -> crate::core::Result<Vec<Font>> {
    let mut fonts = Vec::new();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => {
                if e.local_name().as_ref() == b"font" {
                    fonts.push(parse_font(reader)?);
                } else {
                    xml::skip_element_fast(reader)?;
                }
            },
            Event::End(ref e) if e.local_name().as_ref() == b"fonts" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(fonts)
}

/// Parse a single `<font>` element.
fn parse_font(reader: &mut quick_xml::Reader<&[u8]>) -> crate::core::Result<Font> {
    let mut bold = false;
    let mut italic = false;
    let mut underline = None;
    let mut strike = false;
    let mut size = None;
    let mut name = None;
    let mut color = None;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) | Event::Empty(ref e) => match e.local_name().as_ref() {
                b"b" => bold = parse_toggle(e),
                b"i" => italic = parse_toggle(e),
                b"u" => {
                    underline = Some(
                        xml::optional_attr_str(e, b"val")?
                            .map(|v| v.into_owned())
                            .unwrap_or_else(|| "single".to_string()),
                    );
                },
                b"strike" => strike = parse_toggle(e),
                b"sz" => {
                    size = xml::optional_attr_str(e, b"val")?.and_then(|v| v.parse().ok());
                },
                b"name" => {
                    name = xml::optional_attr_str(e, b"val")?.map(|v| v.into_owned());
                },
                b"color" => {
                    color = parse_color_ref(e)?;
                },
                _ => {},
            },
            Event::End(ref e) if e.local_name().as_ref() == b"font" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(Font {
        bold,
        italic,
        underline,
        strike,
        size,
        name,
        color,
    })
}

/// Parse `<fills>` collection.
fn parse_fills(reader: &mut quick_xml::Reader<&[u8]>) -> crate::core::Result<Vec<Fill>> {
    let mut fills = Vec::new();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => {
                if e.local_name().as_ref() == b"fill" {
                    fills.push(parse_fill(reader)?);
                } else {
                    xml::skip_element_fast(reader)?;
                }
            },
            Event::End(ref e) if e.local_name().as_ref() == b"fills" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(fills)
}

/// Parse a single `<fill>` element.
fn parse_fill(reader: &mut quick_xml::Reader<&[u8]>) -> crate::core::Result<Fill> {
    let mut pattern_type = None;
    let mut fg_color = None;
    let mut bg_color = None;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) | Event::Empty(ref e) => match e.local_name().as_ref() {
                b"patternFill" => {
                    pattern_type =
                        xml::optional_attr_str(e, b"patternType")?.map(|v| v.into_owned());
                },
                b"fgColor" => {
                    fg_color = parse_color_ref(e)?;
                },
                b"bgColor" => {
                    bg_color = parse_color_ref(e)?;
                },
                _ => {},
            },
            Event::End(ref e) if e.local_name().as_ref() == b"fill" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(Fill {
        pattern_type,
        fg_color,
        bg_color,
    })
}

/// Parse `<borders>` collection.
fn parse_borders(reader: &mut quick_xml::Reader<&[u8]>) -> crate::core::Result<Vec<Border>> {
    let mut borders = Vec::new();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => {
                if e.local_name().as_ref() == b"border" {
                    borders.push(parse_border(reader)?);
                } else {
                    xml::skip_element_fast(reader)?;
                }
            },
            Event::End(ref e) if e.local_name().as_ref() == b"borders" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(borders)
}

/// Parse a single `<border>` element.
fn parse_border(reader: &mut quick_xml::Reader<&[u8]>) -> crate::core::Result<Border> {
    let mut left = None;
    let mut right = None;
    let mut top = None;
    let mut bottom = None;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => match e.local_name().as_ref() {
                b"left" | b"start" => left = parse_border_side(reader, e)?,
                b"right" | b"end" => right = parse_border_side(reader, e)?,
                b"top" => top = parse_border_side(reader, e)?,
                b"bottom" => bottom = parse_border_side(reader, e)?,
                _ => {
                    xml::skip_element_fast(reader)?;
                },
            },
            Event::Empty(ref e) => {
                match e.local_name().as_ref() {
                    b"left" | b"start" | b"right" | b"end" | b"top" | b"bottom" => {
                        // Empty border side — check for style attribute
                        if let Some(style) = xml::optional_attr_str(e, b"style")? {
                            let side = BorderSide {
                                style: style.into_owned(),
                                color: None,
                            };
                            match e.local_name().as_ref() {
                                b"left" | b"start" => left = Some(side),
                                b"right" | b"end" => right = Some(side),
                                b"top" => top = Some(side),
                                b"bottom" => bottom = Some(side),
                                _ => {},
                            }
                        }
                    },
                    _ => {},
                }
            },
            Event::End(ref e) if e.local_name().as_ref() == b"border" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(Border {
        left,
        right,
        top,
        bottom,
    })
}

/// Parse a border side element (e.g., `<left style="thin"><color rgb="FF000000"/></left>`).
fn parse_border_side(
    reader: &mut quick_xml::Reader<&[u8]>,
    start: &quick_xml::events::BytesStart,
) -> crate::core::Result<Option<BorderSide>> {
    let style = xml::optional_attr_str(start, b"style")?.map(|v| v.into_owned());
    let mut color = None;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) | Event::Empty(ref e) if e.local_name().as_ref() == b"color" => {
                color = parse_color_ref(e)?;
            },
            Event::End(ref e) => {
                let local = e.local_name();
                if matches!(
                    local.as_ref(),
                    b"left" | b"right" | b"top" | b"bottom" | b"start" | b"end"
                ) {
                    break;
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }

    match style {
        Some(s) => Ok(Some(BorderSide { style: s, color })),
        None => Ok(None),
    }
}

/// Parse `<cellXfs>` or `<cellStyleXfs>` collection.
fn parse_xfs(reader: &mut quick_xml::Reader<&[u8]>) -> crate::core::Result<Vec<CellFormat>> {
    let mut formats = Vec::new();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) | Event::Empty(ref e) if e.local_name().as_ref() == b"xf" => {
                let number_format_id: u32 = xml::optional_attr_str(e, b"numFmtId")?
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(0);
                let font_index = xml::optional_attr_str(e, b"fontId")?.and_then(|v| v.parse().ok());
                let fill_index = xml::optional_attr_str(e, b"fillId")?.and_then(|v| v.parse().ok());
                let border_index =
                    xml::optional_attr_str(e, b"borderId")?.and_then(|v| v.parse().ok());
                let apply_number_format = xml::optional_attr_str(e, b"applyNumberFormat")?
                    .is_some_and(|v| matches!(v.as_ref(), "1" | "true"));
                let xf_id = xml::optional_attr_str(e, b"xfId")?.and_then(|v| v.parse().ok());

                formats.push(CellFormat {
                    number_format_id,
                    font_index,
                    fill_index,
                    border_index,
                    apply_number_format,
                    xf_id,
                });
            },
            Event::End(ref e) => {
                let local = e.local_name();
                if matches!(local.as_ref(), b"cellXfs" | b"cellStyleXfs") {
                    break;
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(formats)
}

/// Parse a toggle element.
fn parse_toggle(e: &quick_xml::events::BytesStart) -> bool {
    xml::parse_toggle(e, b"val")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_styles_basic() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>
  </numFmts>
  <fonts count="2">
    <font>
      <sz val="11"/>
      <name val="Calibri"/>
    </font>
    <font>
      <b/>
      <sz val="14"/>
      <name val="Arial"/>
    </font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="164" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
  </cellXfs>
</styleSheet>"#;
        let ss = StyleSheet::parse(xml).unwrap();

        // Number formats
        assert_eq!(ss.number_formats.len(), 1);
        assert_eq!(ss.number_formats[0].id, 164);
        assert_eq!(ss.number_formats[0].format_code, "yyyy-mm-dd");

        // Fonts
        assert_eq!(ss.fonts.len(), 2);
        assert!(!ss.fonts[0].bold);
        assert_eq!(ss.fonts[0].name.as_deref(), Some("Calibri"));
        assert!(ss.fonts[1].bold);
        assert_eq!(ss.fonts[1].size, Some(14.0));

        // Fills
        assert_eq!(ss.fills.len(), 2);
        assert_eq!(ss.fills[0].pattern_type.as_deref(), Some("none"));

        // Cell formats
        assert_eq!(ss.cell_formats.len(), 2);
        assert_eq!(ss.cell_formats[1].number_format_id, 164);
        assert!(ss.cell_formats[1].apply_number_format);
    }

    #[test]
    fn number_format_lookup() {
        let ss = StyleSheet {
            number_formats: vec![NumberFormat {
                id: 164,
                format_code: "yyyy-mm-dd".to_string(),
            }],
            fonts: vec![],
            fills: vec![],
            borders: vec![],
            cell_formats: vec![
                CellFormat {
                    number_format_id: 0,
                    font_index: None,
                    fill_index: None,
                    border_index: None,
                    apply_number_format: false,
                    xf_id: None,
                },
                CellFormat {
                    number_format_id: 164,
                    font_index: None,
                    fill_index: None,
                    border_index: None,
                    apply_number_format: true,
                    xf_id: None,
                },
            ],
            cell_style_formats: vec![],
        };

        assert_eq!(ss.number_format_for(0), None); // Built-in format 0 not in custom list
        assert_eq!(ss.number_format_for(1), Some("yyyy-mm-dd"));
        assert_eq!(ss.number_format_id_for(0), Some(0));
        assert_eq!(ss.number_format_id_for(1), Some(164));
    }

    #[test]
    fn font_lookup() {
        let ss = StyleSheet {
            number_formats: vec![],
            fonts: vec![
                Font {
                    bold: false,
                    italic: false,
                    underline: None,
                    strike: false,
                    size: Some(11.0),
                    name: Some("Calibri".to_string()),
                    color: None,
                },
                Font {
                    bold: true,
                    italic: false,
                    underline: None,
                    strike: false,
                    size: Some(14.0),
                    name: Some("Arial".to_string()),
                    color: None,
                },
            ],
            fills: vec![],
            borders: vec![],
            cell_formats: vec![CellFormat {
                number_format_id: 0,
                font_index: Some(1),
                fill_index: None,
                border_index: None,
                apply_number_format: false,
                xf_id: None,
            }],
            cell_style_formats: vec![],
        };

        let font = ss.font_for(0).unwrap();
        assert!(font.bold);
        assert_eq!(font.name.as_deref(), Some("Arial"));
    }
}
