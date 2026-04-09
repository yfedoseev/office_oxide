use quick_xml::events::Event;

use crate::core::xml;

type CoreResult<T> = crate::core::Result<T>;

/// Metadata from `ppt/presentation.xml`.
#[derive(Debug, Clone)]
pub struct PresentationInfo {
    pub slides: Vec<SlideId>,
    pub slide_size: Option<SlideSize>,
}

/// An entry in the slide list (`p:sldIdLst`).
#[derive(Debug, Clone)]
pub struct SlideId {
    pub id: u32,
    pub rel_id: String,
}

/// Slide dimensions from `p:sldSz`.
#[derive(Debug, Clone)]
pub struct SlideSize {
    pub cx: i64,
    pub cy: i64,
}

impl PresentationInfo {
    pub(crate) fn parse(xml_data: &[u8]) -> CoreResult<Self> {
        let mut reader = xml::make_fast_reader(xml_data);
        let mut slides = Vec::new();
        let mut slide_size = None;

        loop {
            match reader.read_event()? {
                Event::Start(ref e) => {
                    if e.local_name().as_ref() == b"sldIdLst" {
                        slides = parse_slide_id_list(&mut reader)?;
                    }
                },
                Event::Empty(ref e) => {
                    if e.local_name().as_ref() == b"sldSz" {
                        slide_size = Some(parse_slide_size(e)?);
                    }
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
            Event::Start(ref e) | Event::Empty(ref e) => {
                if e.local_name().as_ref() == b"sldId" {
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
                }
            },
            Event::End(ref e) => {
                if e.local_name().as_ref() == b"sldIdLst" {
                    break;
                }
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
    fn parse_slide_list() {
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
    fn parse_no_slides() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldIdLst/>
</p:presentation>"#;
        let info = PresentationInfo::parse(xml).unwrap();
        assert!(info.slides.is_empty());
        assert!(info.slide_size.is_none());
    }

    #[test]
    fn parse_with_slide_size_only() {
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
}
