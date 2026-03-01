use quick_xml::events::Event;

use office_core::theme::{ColorRef, RgbColor, ThemeColorSlot};
use office_core::xml;

/// Parsed shared string table from `xl/sharedStrings.xml`.
#[derive(Debug, Clone)]
pub struct SharedStringTable {
    pub strings: Vec<SharedString>,
}

/// A single shared string entry.
#[derive(Debug, Clone)]
pub struct SharedString {
    pub text: String,
    pub rich_text: Option<Vec<RichTextRun>>,
}

/// A run within a rich text shared string.
#[derive(Debug, Clone)]
pub struct RichTextRun {
    pub text: String,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub font_size: Option<f64>,
    pub font_name: Option<String>,
    pub color: Option<ColorRef>,
}

impl SharedStringTable {
    /// Create an empty shared string table.
    pub fn empty() -> Self {
        Self {
            strings: Vec::new(),
        }
    }

    /// Parse `xl/sharedStrings.xml` from raw XML bytes.
    pub fn parse(xml_data: &[u8]) -> office_core::Result<Self> {
        // Use non-trimming reader to preserve whitespace in text content
        let mut reader = quick_xml::NsReader::from_reader(xml_data);
        let sml = xml::ns::SML;
        let mut strings = Vec::new();

        loop {
            match reader.read_resolved_event()? {
                (ref resolve, Event::Start(ref e)) => {
                    if xml::matches_ns(resolve, sml) && e.local_name().as_ref() == b"si" {
                        strings.push(parse_si(&mut reader)?);
                    }
                }
                (_, Event::Eof) => break,
                _ => {}
            }
        }

        Ok(Self { strings })
    }

    /// Get the plain text at the given index.
    pub fn get(&self, index: u32) -> Option<&str> {
        self.strings.get(index as usize).map(|s| s.text.as_str())
    }

    /// Get the full shared string at the given index.
    pub fn get_shared(&self, index: u32) -> Option<&SharedString> {
        self.strings.get(index as usize)
    }
}

/// Parse a single `<si>` element.
///
/// An `<si>` can contain either:
/// - Plain text: `<t>text</t>`
/// - Rich text: `<r><rPr>...</rPr><t>text</t></r>` (one or more runs)
fn parse_si(reader: &mut quick_xml::NsReader<&[u8]>) -> office_core::Result<SharedString> {
    let sml = xml::ns::SML;
    let mut plain_text: Option<String> = None;
    let mut runs: Vec<RichTextRun> = Vec::new();

    loop {
        match reader.read_resolved_event()? {
            (ref resolve, Event::Start(ref e)) => {
                if xml::matches_ns(resolve, sml) {
                    match e.local_name().as_ref() {
                        b"t" => {
                            plain_text = Some(xml::read_text_content(reader)?);
                        }
                        b"r" => {
                            runs.push(parse_rich_text_run(reader)?);
                        }
                        _ => {
                            xml::skip_element(reader)?;
                        }
                    }
                } else {
                    xml::skip_element(reader)?;
                }
            }
            (ref resolve, Event::End(ref e)) => {
                if xml::matches_ns(resolve, sml) && e.local_name().as_ref() == b"si" {
                    break;
                }
            }
            (_, Event::Eof) => break,
            _ => {}
        }
    }

    if !runs.is_empty() {
        // Rich text: concatenate all run texts
        let full_text = runs.iter().map(|r| r.text.as_str()).collect::<String>();
        Ok(SharedString {
            text: full_text,
            rich_text: Some(runs),
        })
    } else {
        Ok(SharedString {
            text: plain_text.unwrap_or_default(),
            rich_text: None,
        })
    }
}

/// Parse a single `<r>` rich text run element.
fn parse_rich_text_run(reader: &mut quick_xml::NsReader<&[u8]>) -> office_core::Result<RichTextRun> {
    let sml = xml::ns::SML;
    let mut text = String::new();
    let mut bold = None;
    let mut italic = None;
    let mut font_size = None;
    let mut font_name = None;
    let mut color = None;

    loop {
        match reader.read_resolved_event()? {
            (ref resolve, Event::Start(ref e)) => {
                if xml::matches_ns(resolve, sml) {
                    match e.local_name().as_ref() {
                        b"t" => {
                            text = xml::read_text_content(reader)?;
                        }
                        b"rPr" => {
                            parse_run_properties(
                                reader,
                                &mut bold,
                                &mut italic,
                                &mut font_size,
                                &mut font_name,
                                &mut color,
                            )?;
                        }
                        _ => {
                            xml::skip_element(reader)?;
                        }
                    }
                } else {
                    xml::skip_element(reader)?;
                }
            }
            (ref resolve, Event::End(ref e)) => {
                if xml::matches_ns(resolve, sml) && e.local_name().as_ref() == b"r" {
                    break;
                }
            }
            (_, Event::Eof) => break,
            _ => {}
        }
    }

    Ok(RichTextRun {
        text,
        bold,
        italic,
        font_size,
        font_name,
        color,
    })
}

/// Parse `<rPr>` (run properties) within a shared string run.
fn parse_run_properties(
    reader: &mut quick_xml::NsReader<&[u8]>,
    bold: &mut Option<bool>,
    italic: &mut Option<bool>,
    font_size: &mut Option<f64>,
    font_name: &mut Option<String>,
    color: &mut Option<ColorRef>,
) -> office_core::Result<()> {
    let sml = xml::ns::SML;

    loop {
        match reader.read_resolved_event()? {
            (ref resolve, Event::Start(ref e)) | (ref resolve, Event::Empty(ref e)) => {
                if xml::matches_ns(resolve, sml) {
                    match e.local_name().as_ref() {
                        b"b" => {
                            *bold = Some(parse_toggle(e));
                        }
                        b"i" => {
                            *italic = Some(parse_toggle(e));
                        }
                        b"sz" => {
                            if let Some(val) = xml::optional_attr_str(e, b"val")? {
                                *font_size = val.parse().ok();
                            }
                        }
                        b"rFont" => {
                            if let Some(val) = xml::optional_attr_str(e, b"val")? {
                                *font_name = Some(val.into_owned());
                            }
                        }
                        b"color" => {
                            *color = parse_color_ref(e)?;
                        }
                        _ => {}
                    }
                }
            }
            (ref resolve, Event::End(ref e)) => {
                if xml::matches_ns(resolve, sml) && e.local_name().as_ref() == b"rPr" {
                    break;
                }
            }
            (_, Event::Eof) => break,
            _ => {}
        }
    }
    Ok(())
}

/// Parse a toggle element (e.g., `<b/>` = true, `<b val="0"/>` = false).
fn parse_toggle(e: &quick_xml::events::BytesStart) -> bool {
    match xml::optional_attr_str(e, b"val") {
        Ok(Some(ref val)) => !matches!(val.as_ref(), "0" | "false" | "off"),
        _ => true,
    }
}

/// Parse a color reference from an element's attributes.
pub(crate) fn parse_color_ref(
    e: &quick_xml::events::BytesStart,
) -> office_core::Result<Option<ColorRef>> {
    // Check for direct RGB color
    if let Some(rgb_val) = xml::optional_attr_str(e, b"rgb")? {
        let hex = rgb_val.as_ref();
        // ARGB format: "FF4472C4" — strip alpha prefix if 8 chars
        let hex = if hex.len() == 8 { &hex[2..] } else { hex };
        if hex.len() == 6 {
            return Ok(Some(ColorRef::Rgb(RgbColor::from_hex(hex)?)));
        }
    }

    // Check for theme color
    if let Some(theme_val) = xml::optional_attr_str(e, b"theme")? {
        if let Ok(theme_idx) = theme_val.parse::<u32>() {
            let slot = match theme_idx {
                0 => Some(ThemeColorSlot::Lt1),
                1 => Some(ThemeColorSlot::Dk1),
                2 => Some(ThemeColorSlot::Lt2),
                3 => Some(ThemeColorSlot::Dk2),
                4 => Some(ThemeColorSlot::Accent1),
                5 => Some(ThemeColorSlot::Accent2),
                6 => Some(ThemeColorSlot::Accent3),
                7 => Some(ThemeColorSlot::Accent4),
                8 => Some(ThemeColorSlot::Accent5),
                9 => Some(ThemeColorSlot::Accent6),
                10 => Some(ThemeColorSlot::Hlink),
                11 => Some(ThemeColorSlot::FolHlink),
                _ => None,
            };
            if let Some(slot) = slot {
                let tint = xml::optional_attr_str(e, b"tint")?
                    .and_then(|v| v.parse().ok());
                return Ok(Some(ColorRef::Theme {
                    slot,
                    tint,
                    shade: None,
                }));
            }
        }
    }

    // Check for indexed color (simplified — just return None for now)
    if xml::optional_attr_str(e, b"auto")?.is_some() {
        return Ok(Some(ColorRef::Auto));
    }

    Ok(None)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_plain_strings() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
  <si><t>Hello</t></si>
  <si><t>World</t></si>
  <si><t>Foo Bar</t></si>
</sst>"#;
        let sst = SharedStringTable::parse(xml).unwrap();
        assert_eq!(sst.strings.len(), 3);
        assert_eq!(sst.get(0), Some("Hello"));
        assert_eq!(sst.get(1), Some("World"));
        assert_eq!(sst.get(2), Some("Foo Bar"));
        assert!(sst.strings[0].rich_text.is_none());
    }

    #[test]
    fn parse_rich_text_strings() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r>
      <rPr><b/><sz val="11"/></rPr>
      <t>bold</t>
    </r>
    <r>
      <t> normal</t>
    </r>
  </si>
</sst>"#;
        let sst = SharedStringTable::parse(xml).unwrap();
        assert_eq!(sst.strings.len(), 1);
        assert_eq!(sst.get(0), Some("bold normal"));

        let rich = sst.strings[0].rich_text.as_ref().unwrap();
        assert_eq!(rich.len(), 2);
        assert_eq!(rich[0].text, "bold");
        assert_eq!(rich[0].bold, Some(true));
        assert_eq!(rich[0].font_size, Some(11.0));
        assert_eq!(rich[1].text, " normal");
        assert!(rich[1].bold.is_none());
    }

    #[test]
    fn parse_empty_sst() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0">
</sst>"#;
        let sst = SharedStringTable::parse(xml).unwrap();
        assert_eq!(sst.strings.len(), 0);
        assert_eq!(sst.get(0), None);
    }

    #[test]
    fn index_lookup() {
        let sst = SharedStringTable {
            strings: vec![
                SharedString {
                    text: "first".to_string(),
                    rich_text: None,
                },
                SharedString {
                    text: "second".to_string(),
                    rich_text: None,
                },
            ],
        };
        assert_eq!(sst.get(0), Some("first"));
        assert_eq!(sst.get(1), Some("second"));
        assert_eq!(sst.get(2), None);
    }
}
