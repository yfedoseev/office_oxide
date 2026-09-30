use std::collections::HashMap;

use quick_xml::events::Event;

use crate::core::xml;

use super::formatting::TableBorders;
use super::formatting::{
    ParagraphProperties, RunProperties, parse_paragraph_properties_fast, parse_run_properties_fast,
};
use super::table::{TableCellProperties, TableProperties};

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
    /// Table-level overrides for this style (`w:tblPr`).
    pub table_properties: Option<TableProperties>,
    /// Conditional formats of a table style (`w:tblStylePr`), applied to
    /// the parts of a table its `w:tblLook` switches on.
    pub table_conditionals: Vec<TableStyleConditional>,
}

/// One `w:tblStylePr` (ECMA-376 §17.7.6): formatting a table style applies
/// to one region of the table.
#[derive(Debug, Clone)]
pub struct TableStyleConditional {
    /// The region (`w:type`): `wholeTable`, `firstRow`, `lastRow`,
    /// `firstCol`, `lastCol`, `band1Horz`, `band2Horz`, `band1Vert`,
    /// `band2Vert`, `nwCell`, `neCell`, `swCell` or `seCell`.
    pub kind: String,
    /// Cell formatting for the region (`w:tcPr`). The region's run and
    /// paragraph formatting (`w:rPr`/`w:pPr`, e.g. a bold header row) is
    /// not read: applying it needs the table context threaded through
    /// every run renderer, not just `to_ir()`.
    pub cell_properties: Option<TableCellProperties>,
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
                    "docDefaults" => {
                        sheet.doc_defaults = Some(parse_doc_defaults(&mut reader)?);
                    },
                    "style" => {
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
    ///
    /// Collected on the stack: this runs for every run and paragraph of a
    /// document, and a `Vec` per call was a heap allocation each time.
    fn chain(&self, style_id: &str) -> impl DoubleEndedIterator<Item = &Style> {
        const MAX_CHAIN: usize = 20;
        let mut buf: [Option<&Style>; MAX_CHAIN] = [None; MAX_CHAIN];
        let mut len = 0;
        let mut current = self.styles.get(style_id);
        while let Some(style) = current {
            if len >= MAX_CHAIN {
                break;
            }
            buf[len] = Some(style);
            len += 1;
            current = style.based_on.as_deref().and_then(|id| self.styles.get(id));
        }
        buf.into_iter().take(len).rev().flatten()
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

    /// Whether a run is hidden (`<w:vanish/>`) once document defaults, the
    /// paragraph style chain, the run's character style and its own
    /// `w:rPr` are folded together — the same precedence as
    /// [`Self::effective_run_properties`], without materialising the whole
    /// property set for renderers that only need this one bit.
    pub fn effective_hidden(
        &self,
        paragraph_style_id: Option<&str>,
        direct: Option<&RunProperties>,
    ) -> bool {
        let mut hidden = self
            .doc_defaults
            .as_ref()
            .and_then(|d| d.run_properties.as_ref())
            .and_then(|rp| rp.hidden);
        let chain_hidden = |sid: &str| {
            self.chain(sid)
                .into_iter()
                .rev()
                .find_map(|style| style.run_properties.as_ref().and_then(|rp| rp.hidden))
        };
        if let Some(h) = paragraph_style_id.and_then(chain_hidden) {
            hidden = Some(h);
        }
        if let Some(h) = direct
            .and_then(|d| d.style_id.as_deref())
            .and_then(chain_hidden)
        {
            hidden = Some(h);
        }
        if let Some(h) = direct.and_then(|d| d.hidden) {
            hidden = Some(h);
        }
        hidden.unwrap_or(false)
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
                    "rPrDefault" => {
                        // contains w:rPr
                        defaults.run_properties = parse_nested_rpr(reader)?;
                    },
                    "pPrDefault" => {
                        defaults.paragraph_properties = parse_nested_ppr(reader)?;
                    },
                    _ => {
                        xml::skip_element_fast(reader)?;
                    },
                }
            },
            Event::End(ref e) if e.local_name().as_ref() == "docDefaults" => {
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
                if e.local_name().as_ref() == "rPr" {
                    result = Some(parse_run_properties_fast(reader)?);
                } else {
                    xml::skip_element_fast(reader)?;
                }
            },
            Event::End(ref e) if e.local_name().as_ref() == "rPrDefault" => {
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
                if e.local_name().as_ref() == "pPr" {
                    result = Some(parse_paragraph_properties_fast(reader)?);
                } else {
                    xml::skip_element_fast(reader)?;
                }
            },
            Event::End(ref e) if e.local_name().as_ref() == "pPrDefault" => {
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
    let style_id = match xml::optional_attr_str(start, "w:styleId")? {
        Some(id) => id.into_owned(),
        None => return Ok(None),
    };
    let style_type = match xml::optional_attr_str(start, "w:type")? {
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
    let mut table_properties = None;
    let mut table_conditionals = Vec::new();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => match e.local_name().as_ref() {
                "name" => {
                    if let Ok(Some(val)) = xml::optional_attr_str(e, "w:val") {
                        name = Some(val.into_owned());
                    }
                    xml::skip_element_fast(reader)?;
                },
                "basedOn" => {
                    if let Ok(Some(val)) = xml::optional_attr_str(e, "w:val") {
                        based_on = Some(val.into_owned());
                    }
                    xml::skip_element_fast(reader)?;
                },
                "rPr" => {
                    run_properties = Some(parse_run_properties_fast(reader)?);
                },
                "pPr" => {
                    paragraph_properties = Some(parse_paragraph_properties_fast(reader)?);
                },
                "tblPr" => {
                    table_properties = Some(super::parse_table_properties(reader)?);
                },
                "tblStylePr" => {
                    let kind = xml::optional_attr_str(e, "w:type")?
                        .map(|v| v.into_owned())
                        .unwrap_or_default();
                    table_conditionals.push(parse_table_style_conditional(reader, kind)?);
                },
                _ => {
                    xml::skip_element_fast(reader)?;
                },
            },
            Event::Empty(ref e) => match e.local_name().as_ref() {
                "name" => {
                    if let Ok(Some(val)) = xml::optional_attr_str(e, "w:val") {
                        name = Some(val.into_owned());
                    }
                },
                "basedOn" => {
                    if let Ok(Some(val)) = xml::optional_attr_str(e, "w:val") {
                        based_on = Some(val.into_owned());
                    }
                },
                _ => {},
            },
            Event::End(ref e) if e.local_name().as_ref() == "style" => {
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
        table_properties,
        table_conditionals,
    }))
}

/// The body of one `w:tblStylePr` (its start already consumed).
fn parse_table_style_conditional(
    reader: &mut quick_xml::Reader<&[u8]>,
    kind: String,
) -> crate::core::Result<TableStyleConditional> {
    let mut cond = TableStyleConditional {
        kind,
        cell_properties: None,
    };
    loop {
        match reader.read_event()? {
            Event::Start(ref e) => match e.local_name().as_ref() {
                "tcPr" => cond.cell_properties = Some(super::parse_table_cell_properties(reader)?),
                _ => xml::skip_element_fast(reader)?,
            },
            Event::End(ref e) if e.local_name().as_ref() == "tblStylePr" => break,
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(cond)
}

/// A table style resolved through its `w:basedOn` chain.
#[derive(Debug, Default)]
pub(crate) struct ResolvedTableStyle<'a> {
    /// Borders from the nearest style in the chain that sets them.
    pub borders: Option<&'a TableBorders>,
    /// Rows / columns per band.
    pub row_band_size: Option<u32>,
    pub col_band_size: Option<u32>,
    /// Conditional formats by region; a derived style's entry replaces
    /// its base style's for the same region.
    pub conditionals: Vec<&'a TableStyleConditional>,
}

impl StyleSheet {
    /// Resolve table style `style_id` through its inheritance chain.
    pub(crate) fn resolve_table_style(&self, style_id: &str) -> ResolvedTableStyle<'_> {
        let mut out = ResolvedTableStyle::default();
        for style in self.chain(style_id) {
            if let Some(tp) = style.table_properties.as_ref() {
                if tp.borders.is_some() {
                    out.borders = tp.borders.as_ref();
                }
                out.row_band_size = tp.row_band_size.or(out.row_band_size);
                out.col_band_size = tp.col_band_size.or(out.col_band_size);
            }
            for c in &style.table_conditionals {
                out.conditionals.retain(|prev| prev.kind != c.kind);
                out.conditionals.push(c);
            }
        }
        out
    }
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
    fn test_parse_stylesheet() {
        let sheet = StyleSheet::parse(SAMPLE_STYLES).unwrap();
        assert_eq!(sheet.styles.len(), 3);
        assert!(sheet.doc_defaults.is_some());
    }

    #[test]
    fn test_parse_doc_defaults_font_size() {
        let sheet = StyleSheet::parse(SAMPLE_STYLES).unwrap();
        let defaults = sheet.doc_defaults.as_ref().unwrap();
        let rp = defaults.run_properties.as_ref().unwrap();
        assert_eq!(rp.font_size, Some(crate::core::units::HalfPoint(22)));
    }

    #[test]
    fn test_parse_heading1_style() {
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
    fn test_resolve_outline_level() {
        let sheet = StyleSheet::parse(SAMPLE_STYLES).unwrap();
        assert_eq!(sheet.resolve_outline_level("Heading1"), Some(0));
        assert_eq!(sheet.resolve_outline_level("Normal"), None);
    }
}
