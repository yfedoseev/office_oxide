//! Slide-master `<p:txStyles>` default character/paragraph formatting.
//!
//! Per ECMA-376, a placeholder shape's own direct run/paragraph
//! properties are layered on top of its slide layout's placeholder
//! defaults, which are in turn layered on top of the slide master's
//! `<p:txStyles>` (`titleStyle`/`bodyStyle`/`otherStyle`, one
//! `<a:lvlNpPr>` per outline level). This module resolves only the
//! master's level-0 (`<a:lvl1pPr>`) title/body defaults — the same
//! "outline level 0 only" scope the legacy `.ppt` analogue (
//! `TxMasterStyleAtom`) settled on — and only the character/paragraph
//! properties this crate's IR already has a field for (bold, italic,
//! underline, size, color, alignment). Layout-level overrides and
//! levels 1-8 are a follow-up.

use quick_xml::events::Event;

use crate::core::Result as CoreResult;
use crate::core::xml;
use crate::ir::ParagraphAlignment;

/// The master's level-0 default formatting for one placeholder category
/// (title or body). Every field is `None` when the master's own
/// `<a:defRPr>`/`<a:lvl1pPr>` doesn't specify it — the same "unset,
/// don't guess" contract every other formatting field in this crate's
/// PPTX reader already follows.
#[derive(Debug, Clone, Default, PartialEq)]
pub(crate) struct MasterRunDefaults {
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<String>,
    pub font_size_hundredths_pt: Option<u32>,
    pub color_rgb: Option<[u8; 3]>,
    pub alignment: Option<ParagraphAlignment>,
}

/// A slide master's resolved title/body text-style defaults.
#[derive(Debug, Clone, Default, PartialEq)]
pub(crate) struct MasterTextStyles {
    pub title: Option<MasterRunDefaults>,
    pub body: Option<MasterRunDefaults>,
}

impl MasterTextStyles {
    pub(crate) fn is_empty(&self) -> bool {
        self.title.is_none() && self.body.is_none()
    }
}

/// Parse a `ppt/slideMasters/slideMasterN.xml` document's `<p:txStyles>`
/// into its title/body level-0 defaults. Returns an empty
/// `MasterTextStyles` (not an error) on any malformed/missing input —
/// master-style inheritance is a best-effort enhancement, never a hard
/// requirement for reading the rest of the file.
pub(crate) fn parse_master_text_styles(xml_data: &[u8]) -> MasterTextStyles {
    let mut reader = quick_xml::Reader::from_reader(xml_data);
    reader.config_mut().check_end_names = false;
    reader.config_mut().check_comments = false;

    let mut styles = MasterTextStyles::default();
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let target = match e.local_name().as_ref() {
                    b"titleStyle" => Some(&mut styles.title),
                    b"bodyStyle" => Some(&mut styles.body),
                    _ => None,
                };
                if let Some(target) = target {
                    if target.is_none() {
                        *target = parse_level0_defaults(&mut reader, e.name()).ok().flatten();
                    }
                }
            },
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {},
        }
    }
    styles
}

/// Parse one `<p:titleStyle>`/`<p:bodyStyle>` container: find its first
/// `<a:lvl1pPr>` (or `<a:lvl1pPr/>`) and pull out `algn` plus its
/// `<a:defRPr>`'s formatting, then skip to the container's own end tag
/// regardless of what else it contains.
fn parse_level0_defaults(
    reader: &mut quick_xml::Reader<&[u8]>,
    end_name: quick_xml::name::QName,
) -> CoreResult<Option<MasterRunDefaults>> {
    let mut result = None;
    loop {
        match reader.read_event()? {
            Event::Start(ref e) if e.local_name().as_ref() == b"lvl1pPr" => {
                result = Some(parse_lvl_pr(reader, e)?);
            },
            Event::Empty(ref e) if e.local_name().as_ref() == b"lvl1pPr" => {
                result = Some(MasterRunDefaults {
                    alignment: parse_algn(e)?,
                    ..Default::default()
                });
            },
            Event::End(ref e) if e.name() == end_name => break,
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(result)
}

fn parse_algn(e: &quick_xml::events::BytesStart) -> CoreResult<Option<ParagraphAlignment>> {
    Ok(xml::optional_attr_str(e, b"algn")?.and_then(|v| match v.as_ref() {
        "l" => Some(ParagraphAlignment::Left),
        "ctr" => Some(ParagraphAlignment::Center),
        "r" => Some(ParagraphAlignment::Right),
        "just" | "justLow" => Some(ParagraphAlignment::Justify),
        "dist" | "thaiDist" => Some(ParagraphAlignment::Distribute),
        _ => None,
    }))
}

/// Parse a non-self-closing `<a:lvl1pPr algn="…">…<a:defRPr .../>…</a:lvl1pPr>`:
/// the paragraph's own `algn` attribute plus its child `<a:defRPr>`'s
/// character formatting.
fn parse_lvl_pr(
    reader: &mut quick_xml::Reader<&[u8]>,
    start: &quick_xml::events::BytesStart,
) -> CoreResult<MasterRunDefaults> {
    let alignment = parse_algn(start)?;
    let mut defaults = MasterRunDefaults {
        alignment,
        ..Default::default()
    };
    loop {
        match reader.read_event()? {
            Event::Start(ref e) if e.local_name().as_ref() == b"defRPr" => {
                parse_def_rpr(reader, e, &mut defaults)?;
            },
            Event::Empty(ref e) if e.local_name().as_ref() == b"defRPr" => {
                apply_rpr_attrs(e, &mut defaults)?;
            },
            Event::End(ref e) if e.local_name().as_ref() == b"lvl1pPr" => break,
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(defaults)
}

/// `<a:defRPr b="1" i="0" u="sng" sz="4400">…<a:solidFill><a:srgbClr
/// val="…"/></a:solidFill>…</a:defRPr>` — same attribute/child shape as
/// an ordinary run's `<a:rPr>`, deliberately parsed fresh here (rather
/// than reused from `slide.rs`'s `parse_run_properties`) to avoid
/// disturbing that function's more complex hyperlink/nested-field
/// handling, which `defRPr` never carries.
fn parse_def_rpr(
    reader: &mut quick_xml::Reader<&[u8]>,
    start: &quick_xml::events::BytesStart,
    defaults: &mut MasterRunDefaults,
) -> CoreResult<()> {
    apply_rpr_attrs(start, defaults)?;
    let mut in_solid_fill = false;
    loop {
        match reader.read_event()? {
            Event::Start(ref e) if e.local_name().as_ref() == b"solidFill" => {
                in_solid_fill = true;
            },
            Event::End(ref e) if e.local_name().as_ref() == b"solidFill" => {
                in_solid_fill = false;
            },
            Event::Empty(ref e) if in_solid_fill && e.local_name().as_ref() == b"srgbClr" => {
                if defaults.color_rgb.is_none() {
                    defaults.color_rgb = parse_srgb_clr(e);
                }
            },
            Event::End(ref e) if e.local_name().as_ref() == b"defRPr" => break,
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(())
}

fn apply_rpr_attrs(
    e: &quick_xml::events::BytesStart,
    defaults: &mut MasterRunDefaults,
) -> CoreResult<()> {
    if let Some(v) = xml::optional_attr_str(e, b"b")? {
        defaults.bold = Some(v.as_ref() != "0");
    }
    if let Some(v) = xml::optional_attr_str(e, b"i")? {
        defaults.italic = Some(v.as_ref() != "0");
    }
    if let Some(v) = xml::optional_attr_str(e, b"u")? {
        defaults.underline = Some(v.into_owned());
    }
    if let Some(v) = xml::optional_attr_str(e, b"sz")? {
        defaults.font_size_hundredths_pt = v.parse::<u32>().ok();
    }
    Ok(())
}

fn parse_srgb_clr(e: &quick_xml::events::BytesStart) -> Option<[u8; 3]> {
    let val = xml::optional_attr_str(e, b"val").ok().flatten()?;
    let s = val.as_ref();
    if s.len() != 6 {
        return None;
    }
    let r = u8::from_str_radix(&s[0..2], 16).ok()?;
    let g = u8::from_str_radix(&s[2..4], 16).ok()?;
    let b = u8::from_str_radix(&s[4..6], 16).ok()?;
    Some([r, g, b])
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_title_and_body_level0_defaults_parsed() {
        let xml = br#"<?xml version="1.0"?>
<p:txStyles xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:titleStyle>
    <a:lvl1pPr algn="ctr">
      <a:defRPr sz="4400" b="1">
        <a:solidFill><a:srgbClr val="112233"/></a:solidFill>
      </a:defRPr>
    </a:lvl1pPr>
  </p:titleStyle>
  <p:bodyStyle>
    <a:lvl1pPr algn="l">
      <a:defRPr sz="3200"/>
    </a:lvl1pPr>
    <a:lvl2pPr>
      <a:defRPr sz="2800"/>
    </a:lvl2pPr>
  </p:bodyStyle>
</p:txStyles>"#;

        let styles = parse_master_text_styles(xml);
        let title = styles.title.expect("title style must be present");
        assert_eq!(title.alignment, Some(ParagraphAlignment::Center));
        assert_eq!(title.bold, Some(true));
        assert_eq!(title.font_size_hundredths_pt, Some(4400));
        assert_eq!(title.color_rgb, Some([0x11, 0x22, 0x33]));

        let body = styles.body.expect("body style must be present");
        assert_eq!(body.alignment, Some(ParagraphAlignment::Left));
        assert_eq!(body.font_size_hundredths_pt, Some(3200), "must take level 1, not level 2");
        assert_eq!(body.bold, None);
    }

    #[test]
    fn test_self_closing_lvl1_ppr_with_only_algn_is_not_an_error() {
        let xml = br#"<?xml version="1.0"?>
<p:txStyles xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:titleStyle><a:lvl1pPr algn="r"/></p:titleStyle>
</p:txStyles>"#;
        let styles = parse_master_text_styles(xml);
        assert_eq!(styles.title.unwrap().alignment, Some(ParagraphAlignment::Right));
        assert!(styles.body.is_none());
    }

    #[test]
    fn test_missing_tx_styles_yields_empty() {
        let xml = br#"<?xml version="1.0"?><p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#;
        let styles = parse_master_text_styles(xml);
        assert!(styles.is_empty());
    }

    #[test]
    fn test_truncated_xml_does_not_panic() {
        let styles = parse_master_text_styles(b"<p:txStyles><p:titleStyle><a:lvl1pPr");
        assert!(styles.is_empty());
    }
}
