use quick_xml::events::Event;

use crate::core::xml;

type CoreResult<T> = crate::core::Result<T>;

/// Metadata from `ppt/presentation.xml`.
#[derive(Debug, Clone)]
pub struct PresentationInfo {
    /// Ordered list of slide identifiers.
    pub slides: Vec<SlideId>,
    /// Physical slide dimensions, if present.
    pub slide_size: Option<SlideSize>,
}

/// An entry in the slide list (`p:sldIdLst`).
#[derive(Debug, Clone)]
pub struct SlideId {
    /// Numeric slide identifier from `id` attribute.
    pub id: u32,
    /// Relationship ID used to resolve the slide part path.
    pub rel_id: String,
}

/// Slide dimensions from `p:sldSz`.
#[derive(Debug, Clone)]
pub struct SlideSize {
    /// Width in EMU (English Metric Units).
    pub cx: i64,
    /// Height in EMU.
    pub cy: i64,
}

impl PresentationInfo {
    pub(crate) fn parse(xml_data: &[u8]) -> CoreResult<Self> {
        let mut reader = xml::make_fast_reader(xml_data);
        let mut slides = Vec::new();
        let mut slide_size = None;
        // PowerPoint's Sections feature stores its own `p14:sldIdLst` (per-
        // section slide membership, a different element with different
        // semantics) nested inside `p:extLst/p:ext/p14:sectionLst`.
        // Matching `sldIdLst` by local name alone (unavoidable without full
        // namespace-aware parsing) let it match that element too, and since
        // `p:extLst` comes after the real `p:sldIdLst` in document order,
        // each `p14:sldIdLst` overwrote the correct slide list in turn,
        // leaving only the last section's slides. `p:sldSz`
        // has no `extLst` analogue to collide with, but is guarded the same
        // way for consistency and against a future extension reusing it.
        let mut ext_lst_depth: u32 = 0;

        loop {
            match reader.read_event()? {
                Event::Start(ref e) if e.local_name().as_ref() == b"extLst" => {
                    ext_lst_depth += 1;
                },
                Event::End(ref e) if e.local_name().as_ref() == b"extLst" => {
                    ext_lst_depth = ext_lst_depth.saturating_sub(1);
                },
                Event::Start(ref e)
                    if ext_lst_depth == 0 && e.local_name().as_ref() == b"sldIdLst" =>
                {
                    slides = parse_slide_id_list(&mut reader)?;
                },
                Event::Empty(ref e)
                    if ext_lst_depth == 0 && e.local_name().as_ref() == b"sldSz" =>
                {
                    slide_size = Some(parse_slide_size(e)?);
                },
                Event::Eof => break,
                _ => {},
            }
        }

        Ok(PresentationInfo { slides, slide_size })
    }
}

fn parse_slide_id_list(reader: &mut quick_xml::Reader<&[u8]>) -> CoreResult<Vec<SlideId>> {
    let mut slides = Vec::new();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) | Event::Empty(ref e) if e.local_name().as_ref() == b"sldId" => {
                let id: u32 = xml::optional_attr_str(e, b"id")?
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(0);
                // r:id may be missing in some files (LibreOffice test fixtures)
                // or use a different prefix like d3p1:id instead of r:id
                let rel_id = xml::optional_attr_str(e, b"r:id")?
                    .map(|v| v.into_owned())
                    .unwrap_or_default();
                // Always add the slide — if r:id is missing, we'll try to
                // resolve by position (convention: rId2 = slide1, rId3 = slide2, etc.)
                slides.push(SlideId { id, rel_id });
            },
            Event::End(ref e) if e.local_name().as_ref() == b"sldIdLst" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(slides)
}

fn parse_slide_size(e: &quick_xml::events::BytesStart) -> CoreResult<SlideSize> {
    let cx: i64 = xml::optional_attr_str(e, b"cx")?
        .and_then(|v| v.parse().ok())
        .unwrap_or(0);
    let cy: i64 = xml::optional_attr_str(e, b"cy")?
        .and_then(|v| v.parse().ok())
        .unwrap_or(0);
    Ok(SlideSize { cx, cy })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_slide_list() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId2"/>
    <p:sldId id="257" r:id="rId3"/>
    <p:sldId id="258" r:id="rId4"/>
  </p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
</p:presentation>"#;
        let info = PresentationInfo::parse(xml).unwrap();
        assert_eq!(info.slides.len(), 3);
        assert_eq!(info.slides[0].id, 256);
        assert_eq!(info.slides[0].rel_id, "rId2");
        assert_eq!(info.slides[1].id, 257);
        assert_eq!(info.slides[2].rel_id, "rId4");
        let size = info.slide_size.unwrap();
        assert_eq!(size.cx, 9144000);
        assert_eq!(size.cy, 6858000);
    }

    #[test]
    fn test_parse_no_slides() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldIdLst/>
</p:presentation>"#;
        let info = PresentationInfo::parse(xml).unwrap();
        assert!(info.slides.is_empty());
        assert!(info.slide_size.is_none());
    }

    #[test]
    fn test_parse_with_slide_size_only() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldSz cx="12192000" cy="6858000"/>
</p:presentation>"#;
        let info = PresentationInfo::parse(xml).unwrap();
        assert!(info.slides.is_empty());
        let size = info.slide_size.unwrap();
        assert_eq!(size.cx, 12192000);
        assert_eq!(size.cy, 6858000);
    }

    /// PowerPoint Sections store per-section slide membership as
    /// `p14:sldIdLst` inside `p:extLst/p:ext/p14:sectionLst` — a different
    /// element that merely shares a local name with the real `p:sldIdLst`.
    /// Matching by local name alone let each section's list overwrite the
    /// real one in turn, so a presentation with N sections lost everything
    /// but the last section's slides. This is the issue's own
    /// minimal reproducer XML shape.
    #[test]
    fn test_sections_extension_does_not_overwrite_the_real_slide_list() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst>
    <p:sldId id="1" r:id="rId2"/>
    <p:sldId id="2" r:id="rId3"/>
  </p:sldIdLst>
  <p:extLst>
    <p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}">
      <p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">
        <p14:section name="A"><p14:sldIdLst><p14:sldId id="1"/></p14:sldIdLst></p14:section>
        <p14:section name="B"><p14:sldIdLst><p14:sldId id="2"/></p14:sldIdLst></p14:section>
      </p14:sectionLst>
    </p:ext>
  </p:extLst>
</p:presentation>"#;
        let info = PresentationInfo::parse(xml).unwrap();
        assert_eq!(
            info.slides.len(),
            2,
            "the real 2-slide p:sldIdLst must survive the extLst Sections block, got {:?}",
            info.slides
        );
        assert_eq!(info.slides[0].id, 1);
        assert_eq!(info.slides[0].rel_id, "rId2");
        assert_eq!(info.slides[1].id, 2);
        assert_eq!(info.slides[1].rel_id, "rId3");
    }
}
