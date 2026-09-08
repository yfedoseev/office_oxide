use std::collections::HashMap;

use quick_xml::events::Event;

use crate::core::xml;

use super::formatting::{
    ParagraphProperties, RunProperties, parse_paragraph_properties_fast, parse_run_properties_fast,
};
use super::table::TableProperties;

/// Parsed stylesheet from `word/styles.xml`.
#[derive(Debug, Clone, Default)]
pub struct StyleSheet {
    /// Document-wide default formatting.
    pub doc_defaults: Option<DocDefaults>,
    /// Map from style ID to style definition.
    pub styles: HashMap<String, Style>,
}

/// Document-wide default properties.
#[derive(Debug, Clone, Default)]
pub struct DocDefaults {
    /// Default run properties applied to all text.
    pub run_properties: Option<RunProperties>,
    /// Default paragraph properties.
    pub paragraph_properties: Option<ParagraphProperties>,
}

/// A single style definition.
#[derive(Debug, Clone)]
pub struct Style {
    /// Unique style identifier (e.g., `"Heading1"`).
    pub style_id: String,
    /// Kind of style (paragraph, character, table, or numbering).
    pub style_type: StyleType,
    /// Human-readable style name.
    pub name: Option<String>,
    /// ID of the parent style this style inherits from.
    pub based_on: Option<String>,
    /// Run-level overrides for this style.
    pub run_properties: Option<RunProperties>,
    /// Paragraph-level overrides for this style.
    pub paragraph_properties: Option<ParagraphProperties>,
    /// Table-level overrides for this style.
    pub table_properties: Option<TableProperties>,
}

/// The kind of style.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StyleType {
    /// Paragraph style.
    Paragraph,
    /// Character (inline) style.
    Character,
    /// Table style.
    Table,
    /// Numbering style.
    Numbering,
}

impl StyleSheet {
    /// Parse `word/styles.xml` content.
    pub fn parse(xml_data: &[u8]) -> crate::core::Result<Self> {
        let mut reader = xml::make_fast_reader(xml_data);
        let mut sheet = StyleSheet::default();

        loop {
            match reader.read_event()? {
                Event::Start(ref e) => match e.local_name().as_ref() {
                    b"docDefaults" => {
                        sheet.doc_defaults = Some(parse_doc_defaults(&mut reader)?);
                    },
                    b"style" => {
                        if let Some(style) = parse_style(&mut reader, e)? {
                            sheet.styles.insert(style.style_id.clone(), style);
                        }
                    },
                    _ => {},
                },
                Event::Eof => break,
                _ => {},
            }
        }
        Ok(sheet)
    }

    /// Return the inheritance chain for `style_id`, **root ancestor first**,
    /// so callers can overlay each style in turn and have the most specific
    /// one win. Cycles and pathological `w:basedOn` chains are cut off at 20
    /// links.
    fn chain(&self, style_id: &str) -> Vec<&Style> {
        let mut out = Vec::new();
        let mut current = self.styles.get(style_id);
        while let Some(style) = current {
            if out.len() >= 20 {
                break;
            }
            out.push(style);
            current = style.based_on.as_deref().and_then(|id| self.styles.get(id));
        }
        out.reverse();
        out
    }

    /// Fold the effective run formatting for a run: document defaults, then
    /// the paragraph style chain's run properties, then the character style
    /// chain, then the run's own `w:rPr`. Later stages win field by field.
    ///
    /// Without this, only direct `w:rPr` reached the IR — so a document that
    /// puts all of its formatting in styles (which is what Word's built-in
    /// styles and every template do) read back as unformatted text.
    pub fn effective_run_properties(
        &self,
        paragraph_style_id: Option<&str>,
        direct: Option<&RunProperties>,
    ) -> RunProperties {
        let mut out = self
            .doc_defaults
            .as_ref()
            .and_then(|d| d.run_properties.clone())
            .unwrap_or_default();
        if let Some(pid) = paragraph_style_id {
            for style in self.chain(pid) {
                if let Some(rp) = style.run_properties.as_ref() {
                    out.overlay(rp);
                }
            }
        }
        // `w:rStyle` on the run names a character style, which sits above
        // the paragraph style but below the run's own direct formatting.
        if let Some(cid) = direct.and_then(|d| d.style_id.as_deref()) {
            for style in self.chain(cid) {
                if let Some(rp) = style.run_properties.as_ref() {
                    out.overlay(rp);
                }
            }
        }
        if let Some(d) = direct {
            out.overlay(d);
        }
        out
    }

    /// Fold the effective paragraph formatting: document defaults, then the
    /// style chain, then the paragraph's own `w:pPr`.
    pub fn effective_paragraph_properties(
        &self,
        direct: Option<&ParagraphProperties>,
    ) -> ParagraphProperties {
        let mut out = self
            .doc_defaults
            .as_ref()
            .and_then(|d| d.paragraph_properties.clone())
            .unwrap_or_default();
        if let Some(pid) = direct.and_then(|d| d.style_id.as_deref()) {
            for style in self.chain(pid) {
                if let Some(pp) = style.paragraph_properties.as_ref() {
                    out.overlay(pp);
                }
            }
        }
        if let Some(d) = direct {
            out.overlay(d);
        }
        out
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
                    },
                    b"pPrDefault" => {
                        defaults.paragraph_properties = parse_nested_ppr(reader)?;
                    },
                    _ => {
                        xml::skip_element_fast(reader)?;
                    },
                }
            },
            Event::End(ref e) if e.local_name().as_ref() == b"docDefaults" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(defaults)
}

/// Parse the `w:rPr` nested inside `w:rPrDefault` (or similar wrapper).
fn parse_nested_rpr(
    reader: &mut quick_xml::Reader<&[u8]>,
) -> crate::core::Result<Option<RunProperties>> {
    let mut result = None;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => {
                if e.local_name().as_ref() == b"rPr" {
                    result = Some(parse_run_properties_fast(reader)?);
                } else {
                    xml::skip_element_fast(reader)?;
                }
            },
            Event::End(ref e) if e.local_name().as_ref() == b"rPrDefault" => {
                break;
            },
            Event::Eof => break,
            _ => {},
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
            },
            Event::End(ref e) if e.local_name().as_ref() == b"pPrDefault" => {
                break;
            },
            Event::Eof => break,
            _ => {},
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
            Event::Start(ref e) => match e.local_name().as_ref() {
                b"name" => {
                    if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                        name = Some(val.into_owned());
                    }
                    xml::skip_element_fast(reader)?;
                },
                b"basedOn" => {
                    if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                        based_on = Some(val.into_owned());
                    }
                    xml::skip_element_fast(reader)?;
                },
                b"rPr" => {
                    run_properties = Some(parse_run_properties_fast(reader)?);
                },
                b"pPr" => {
                    paragraph_properties = Some(parse_paragraph_properties_fast(reader)?);
                },
                _ => {
                    xml::skip_element_fast(reader)?;
                },
            },
            Event::Empty(ref e) => match e.local_name().as_ref() {
                b"name" => {
                    if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                        name = Some(val.into_owned());
                    }
                },
                b"basedOn" => {
                    if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                        based_on = Some(val.into_owned());
                    }
                },
                _ => {},
            },
            Event::End(ref e) if e.local_name().as_ref() == b"style" => {
                break;
            },
            Event::Eof => break,
            _ => {},
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
