use std::collections::HashMap;

use quick_xml::events::Event;

use crate::core::xml;

use super::formatting::{
    parse_paragraph_properties_fast, parse_run_properties_fast, ParagraphProperties, RunProperties,
};
use super::table::TableProperties;

/// Parsed stylesheet from `word/styles.xml`.
#[derive(Debug, Clone, Default)]
pub struct StyleSheet {
    pub doc_defaults: Option<DocDefaults>,
    pub styles: HashMap<String, Style>,
}

/// Document-wide default properties.
#[derive(Debug, Clone, Default)]
pub struct DocDefaults {
    pub run_properties: Option<RunProperties>,
    pub paragraph_properties: Option<ParagraphProperties>,
}

/// A single style definition.
#[derive(Debug, Clone)]
pub struct Style {
    pub style_id: String,
    pub style_type: StyleType,
    pub name: Option<String>,
    pub based_on: Option<String>,
    pub run_properties: Option<RunProperties>,
    pub paragraph_properties: Option<ParagraphProperties>,
    pub table_properties: Option<TableProperties>,
}

/// The kind of style.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StyleType {
    Paragraph,
    Character,
    Table,
    Numbering,
}

impl StyleSheet {
    /// Parse `word/styles.xml` content.
    pub fn parse(xml_data: &[u8]) -> crate::core::Result<Self> {
        let mut reader = xml::make_fast_reader(xml_data);
        let mut sheet = StyleSheet::default();

        loop {
            match reader.read_event()? {
                Event::Start(ref e) => {
                    match e.local_name().as_ref() {
                        b"docDefaults" => {
                            sheet.doc_defaults = Some(parse_doc_defaults(&mut reader)?);
                        }
                        b"style" => {
                            if let Some(style) = parse_style(&mut reader, e)? {
                                sheet.styles.insert(style.style_id.clone(), style);
                            }
                        }
                        _ => {}
                    }
                }
                Event::Eof => break,
                _ => {}
            }
        }
        Ok(sheet)
    }

    /// Resolve the effective outline level for a given style ID, walking the inheritance chain.
    pub fn resolve_outline_level(&self, style_id: &str) -> Option<u8> {
        let mut current = self.styles.get(style_id);
        let mut depth = 0;
        while let Some(style) = current {
            if depth > 20 {
                break; // prevent infinite loops
            }
            if let Some(ref pp) = style.paragraph_properties {
                if let Some(lvl) = pp.outline_level {
                    return Some(lvl);
                }
            }
            current = style.based_on.as_ref().and_then(|id| self.styles.get(id));
            depth += 1;
        }
        None
    }
}

fn parse_doc_defaults(reader: &mut quick_xml::Reader<&[u8]>) -> crate::core::Result<DocDefaults> {
    let mut defaults = DocDefaults::default();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => {
                match e.local_name().as_ref() {
                    b"rPrDefault" => {
                        // contains w:rPr
                        defaults.run_properties = parse_nested_rpr(reader)?;
                    }
                    b"pPrDefault" => {
                        defaults.paragraph_properties = parse_nested_ppr(reader)?;
                    }
                    _ => {
                        xml::skip_element_fast(reader)?;
                    }
                }
            }
            Event::End(ref e) => {
                if e.local_name().as_ref() == b"docDefaults" {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(defaults)
}

/// Parse the `w:rPr` nested inside `w:rPrDefault` (or similar wrapper).
fn parse_nested_rpr(reader: &mut quick_xml::Reader<&[u8]>) -> crate::core::Result<Option<RunProperties>> {
    let mut result = None;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => {
                if e.local_name().as_ref() == b"rPr" {
                    result = Some(parse_run_properties_fast(reader)?);
                } else {
                    xml::skip_element_fast(reader)?;
                }
            }
            Event::End(ref e) => {
                if e.local_name().as_ref() == b"rPrDefault" {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(result)
}

/// Parse the `w:pPr` nested inside `w:pPrDefault`.
fn parse_nested_ppr(
    reader: &mut quick_xml::Reader<&[u8]>,
) -> crate::core::Result<Option<ParagraphProperties>> {
    let mut result = None;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => {
                if e.local_name().as_ref() == b"pPr" {
                    result = Some(parse_paragraph_properties_fast(reader)?);
                } else {
                    xml::skip_element_fast(reader)?;
                }
            }
            Event::End(ref e) => {
                if e.local_name().as_ref() == b"pPrDefault" {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(result)
}

fn parse_style(
    reader: &mut quick_xml::Reader<&[u8]>,
    start: &quick_xml::events::BytesStart,
) -> crate::core::Result<Option<Style>> {
    let style_id = match xml::optional_attr_str(start, b"w:styleId")? {
        Some(id) => id.into_owned(),
        None => return Ok(None),
    };
    let style_type = match xml::optional_attr_str(start, b"w:type")? {
        Some(ref t) => match t.as_ref() {
            "paragraph" => StyleType::Paragraph,
            "character" => StyleType::Character,
            "table" => StyleType::Table,
            "numbering" => StyleType::Numbering,
            _ => StyleType::Paragraph,
        },
        None => StyleType::Paragraph,
    };

    let mut name = None;
    let mut based_on = None;
    let mut run_properties = None;
    let mut paragraph_properties = None;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => {
                match e.local_name().as_ref() {
                    b"name" => {
                        if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                            name = Some(val.into_owned());
                        }
                        xml::skip_element_fast(reader)?;
                    }
                    b"basedOn" => {
                        if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                            based_on = Some(val.into_owned());
                        }
                        xml::skip_element_fast(reader)?;
                    }
                    b"rPr" => {
                        run_properties = Some(parse_run_properties_fast(reader)?);
                    }
                    b"pPr" => {
                        paragraph_properties = Some(parse_paragraph_properties_fast(reader)?);
                    }
                    _ => {
                        xml::skip_element_fast(reader)?;
                    }
                }
            }
            Event::Empty(ref e) => {
                match e.local_name().as_ref() {
                    b"name" => {
                        if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                            name = Some(val.into_owned());
                        }
                    }
                    b"basedOn" => {
                        if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                            based_on = Some(val.into_owned());
                        }
                    }
                    _ => {}
                }
            }
            Event::End(ref e) => {
                if e.local_name().as_ref() == b"style" {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(Some(Style {
        style_id,
        style_type,
        name,
        based_on,
        run_properties,
        paragraph_properties,
        table_properties: None,
    }))
}

#[cfg(test)]
mod tests {
    use super::*;

    const SAMPLE_STYLES: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:sz w:val="22"/>
      </w:rPr>
    </w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:outlineLvl w:val="0"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:sz w:val="32"/>
    </w:rPr>
  </w:style>
  <w:style w:type="character" w:styleId="Strong">
    <w:name w:val="Strong"/>
    <w:rPr>
      <w:b/>
    </w:rPr>
  </w:style>
</w:styles>"#;

    #[test]
    fn parse_stylesheet() {
        let sheet = StyleSheet::parse(SAMPLE_STYLES).unwrap();
        assert_eq!(sheet.styles.len(), 3);
        assert!(sheet.doc_defaults.is_some());
    }

    #[test]
    fn parse_doc_defaults_font_size() {
        let sheet = StyleSheet::parse(SAMPLE_STYLES).unwrap();
        let defaults = sheet.doc_defaults.as_ref().unwrap();
        let rp = defaults.run_properties.as_ref().unwrap();
        assert_eq!(rp.font_size, Some(crate::core::units::HalfPoint(22)));
    }

    #[test]
    fn parse_heading1_style() {
        let sheet = StyleSheet::parse(SAMPLE_STYLES).unwrap();
        let h1 = sheet.styles.get("Heading1").unwrap();
        assert_eq!(h1.name.as_deref(), Some("heading 1"));
        assert_eq!(h1.based_on.as_deref(), Some("Normal"));
        assert_eq!(h1.style_type, StyleType::Paragraph);

        let pp = h1.paragraph_properties.as_ref().unwrap();
        assert_eq!(pp.outline_level, Some(0));

        let rp = h1.run_properties.as_ref().unwrap();
        assert_eq!(rp.bold, Some(true));
        assert_eq!(rp.font_size, Some(crate::core::units::HalfPoint(32)));
    }

    #[test]
    fn resolve_outline_level() {
        let sheet = StyleSheet::parse(SAMPLE_STYLES).unwrap();
        assert_eq!(sheet.resolve_outline_level("Heading1"), Some(0));
        assert_eq!(sheet.resolve_outline_level("Normal"), None);
    }
}
