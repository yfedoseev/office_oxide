use quick_xml::Writer;
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use serde::{Deserialize, Serialize};

use super::error::Result;
use super::xml;

// ---------------------------------------------------------------------------
// Core Properties (Dublin Core metadata) — docProps/core.xml
// ---------------------------------------------------------------------------

/// Core properties (Dublin Core + OPC metadata) from `docProps/core.xml`.
#[derive(Debug, Default, Clone, Serialize, Deserialize)]
pub struct CoreProperties {
    /// Document title.
    pub title: Option<String>,
    /// Document subject.
    pub subject: Option<String>,
    /// Primary author.
    pub creator: Option<String>,
    /// Search keywords.
    pub keywords: Option<String>,
    /// Document description or abstract.
    pub description: Option<String>,
    /// User who last saved the document.
    pub last_modified_by: Option<String>,
    /// Revision number as a string.
    pub revision: Option<String>,
    /// Creation date-time (ISO 8601).
    pub created: Option<String>,
    /// Last-modified date-time (ISO 8601).
    pub modified: Option<String>,
    /// Document category.
    pub category: Option<String>,
    /// Content status (e.g., "Draft", "Final").
    pub content_status: Option<String>,
    /// Document language.
    pub language: Option<String>,
}

impl CoreProperties {
    /// Parse core properties from `docProps/core.xml` bytes.
    pub fn parse(xml_data: &[u8]) -> Result<Self> {
        let mut reader = xml::make_fast_reader(xml_data);
        let mut props = CoreProperties::default();

        // State: which element are we inside?
        // Since we no longer have namespace resolution, we match on local name only.
        // The element names are unique enough across namespaces to be unambiguous.
        enum Ctx {
            None,
            Title,
            Subject,
            Creator,
            Keywords,
            Description,
            LastModifiedBy,
            Revision,
            Created,
            Modified,
            Category,
            ContentStatus,
            Language,
        }
        let mut ctx = Ctx::None;

        loop {
            match reader.read_event()? {
                Event::Start(ref e) => {
                    let local = e.local_name();
                    let local_bytes = local.as_ref();

                    ctx = match local_bytes {
                        b"title" => Ctx::Title,
                        b"subject" => Ctx::Subject,
                        b"creator" => Ctx::Creator,
                        b"description" => Ctx::Description,
                        b"language" => Ctx::Language,
                        b"created" => Ctx::Created,
                        b"modified" => Ctx::Modified,
                        b"keywords" => Ctx::Keywords,
                        b"lastModifiedBy" => Ctx::LastModifiedBy,
                        b"revision" => Ctx::Revision,
                        b"category" => Ctx::Category,
                        b"contentStatus" => Ctx::ContentStatus,
                        _ => Ctx::None,
                    };
                },
                Event::Text(ref e) => {
                    let text = e.unescape()?.into_owned();
                    if text.is_empty() {
                        continue;
                    }
                    match ctx {
                        Ctx::Title => props.title = Some(text),
                        Ctx::Subject => props.subject = Some(text),
                        Ctx::Creator => props.creator = Some(text),
                        Ctx::Keywords => props.keywords = Some(text),
                        Ctx::Description => props.description = Some(text),
                        Ctx::LastModifiedBy => props.last_modified_by = Some(text),
                        Ctx::Revision => props.revision = Some(text),
                        Ctx::Created => props.created = Some(text),
                        Ctx::Modified => props.modified = Some(text),
                        Ctx::Category => props.category = Some(text),
                        Ctx::ContentStatus => props.content_status = Some(text),
                        Ctx::Language => props.language = Some(text),
                        Ctx::None => {},
                    }
                },
                Event::End(_) => {
                    ctx = Ctx::None;
                },
                Event::Eof => break,
                _ => {},
            }
        }

        Ok(props)
    }

    /// Serialize to `docProps/core.xml` bytes.
    pub fn serialize(&self) -> Vec<u8> {
        let mut w = Writer::new_with_indent(Vec::new(), b' ', 2);

        w.write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
            .expect("write decl");

        let mut root = BytesStart::new("cp:coreProperties");
        root.push_attribute((
            "xmlns:cp",
            "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
        ));
        root.push_attribute(("xmlns:dc", "http://purl.org/dc/elements/1.1/"));
        root.push_attribute(("xmlns:dcterms", "http://purl.org/dc/terms/"));
        root.push_attribute(("xmlns:dcmitype", "http://purl.org/dc/dcmitype/"));
        root.push_attribute(("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"));
        w.write_event(Event::Start(root)).expect("write root");

        write_optional_element(&mut w, "dc:title", self.title.as_deref());
        write_optional_element(&mut w, "dc:subject", self.subject.as_deref());
        write_optional_element(&mut w, "dc:creator", self.creator.as_deref());
        write_optional_element(&mut w, "dc:description", self.description.as_deref());
        write_optional_element(&mut w, "dc:language", self.language.as_deref());
        write_optional_element(&mut w, "cp:keywords", self.keywords.as_deref());
        write_optional_element(&mut w, "cp:category", self.category.as_deref());
        write_optional_element(&mut w, "cp:contentStatus", self.content_status.as_deref());
        write_optional_element(&mut w, "cp:lastModifiedBy", self.last_modified_by.as_deref());
        write_optional_element(&mut w, "cp:revision", self.revision.as_deref());

        if let Some(ref created) = self.created {
            write_datetime_element(&mut w, "dcterms:created", created);
        }
        if let Some(ref modified) = self.modified {
            write_datetime_element(&mut w, "dcterms:modified", modified);
        }

        w.write_event(Event::End(BytesEnd::new("cp:coreProperties")))
            .expect("write end root");

        w.into_inner()
    }
}

fn write_optional_element(w: &mut Writer<Vec<u8>>, tag: &str, value: Option<&str>) {
    if let Some(text) = value {
        w.write_event(Event::Start(BytesStart::new(tag)))
            .expect("write start");
        w.write_event(Event::Text(BytesText::new(text)))
            .expect("write text");
        w.write_event(Event::End(BytesEnd::new(tag)))
            .expect("write end");
    }
}

fn write_datetime_element(w: &mut Writer<Vec<u8>>, tag: &str, value: &str) {
    let mut elem = BytesStart::new(tag);
    elem.push_attribute(("xsi:type", "dcterms:W3CDTF"));
    w.write_event(Event::Start(elem)).expect("write start");
    w.write_event(Event::Text(BytesText::new(value)))
        .expect("write text");
    w.write_event(Event::End(BytesEnd::new(tag)))
        .expect("write end");
}

// ---------------------------------------------------------------------------
// App (Extended) Properties — docProps/app.xml
// ---------------------------------------------------------------------------

/// Extended/application properties from `docProps/app.xml`.
#[derive(Debug, Default, Clone, Serialize, Deserialize)]
pub struct AppProperties {
    /// Producing application name.
    pub application: Option<String>,
    /// Application version string.
    pub app_version: Option<String>,
    /// Company or organisation name.
    pub company: Option<String>,
    /// Template the document is based on.
    pub template: Option<String>,
    /// Total editing time in minutes.
    pub total_time: Option<u32>,
    /// Number of pages.
    pub pages: Option<u32>,
    /// Word count.
    pub words: Option<u32>,
    /// Character count (excluding spaces).
    pub characters: Option<u32>,
    /// Character count (including spaces).
    pub characters_with_spaces: Option<u32>,
    /// Line count.
    pub lines: Option<u32>,
    /// Paragraph count.
    pub paragraphs: Option<u32>,
    /// Slide count (presentations).
    pub slides: Option<u32>,
    /// Notes-page count (presentations).
    pub notes: Option<u32>,
    /// Hidden slide count (presentations).
    pub hidden_slides: Option<u32>,
}

impl AppProperties {
    /// Parse app properties from `docProps/app.xml` bytes.
    pub fn parse(xml_data: &[u8]) -> Result<Self> {
        let mut reader = xml::make_fast_reader(xml_data);
        let mut props = AppProperties::default();
        let mut current_tag: Option<String> = None;

        loop {
            match reader.read_event()? {
                Event::Start(ref e) => {
                    let local = e.local_name();
                    let local_bytes = local.as_ref();
                    current_tag = Some(String::from_utf8_lossy(local_bytes).into_owned());
                },
                Event::Text(ref e) => {
                    let text = e.unescape()?.into_owned();
                    if let Some(ref tag) = current_tag {
                        match tag.as_str() {
                            "Application" => props.application = Some(text),
                            "AppVersion" => props.app_version = Some(text),
                            "Company" => props.company = Some(text),
                            "Template" => props.template = Some(text),
                            "TotalTime" => props.total_time = text.parse().ok(),
                            "Pages" => props.pages = text.parse().ok(),
                            "Words" => props.words = text.parse().ok(),
                            "Characters" => props.characters = text.parse().ok(),
                            "CharactersWithSpaces" => {
                                props.characters_with_spaces = text.parse().ok();
                            },
                            "Lines" => props.lines = text.parse().ok(),
                            "Paragraphs" => props.paragraphs = text.parse().ok(),
                            "Slides" => props.slides = text.parse().ok(),
                            "Notes" => props.notes = text.parse().ok(),
                            "HiddenSlides" => props.hidden_slides = text.parse().ok(),
                            _ => {},
                        }
                    }
                },
                Event::End(_) => {
                    current_tag = None;
                },
                Event::Eof => break,
                _ => {},
            }
        }

        Ok(props)
    }

    /// Serialize to `docProps/app.xml` bytes.
    pub fn serialize(&self) -> Vec<u8> {
        let mut w = Writer::new_with_indent(Vec::new(), b' ', 2);

        w.write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
            .expect("write decl");

        let mut root = BytesStart::new("Properties");
        root.push_attribute((
            "xmlns",
            "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
        ));
        root.push_attribute((
            "xmlns:vt",
            "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
        ));
        w.write_event(Event::Start(root)).expect("write root");

        write_optional_element(&mut w, "Application", self.application.as_deref());
        write_optional_element(&mut w, "AppVersion", self.app_version.as_deref());
        write_optional_element(&mut w, "Company", self.company.as_deref());
        write_optional_element(&mut w, "Template", self.template.as_deref());
        write_optional_u32(&mut w, "TotalTime", self.total_time);
        write_optional_u32(&mut w, "Pages", self.pages);
        write_optional_u32(&mut w, "Words", self.words);
        write_optional_u32(&mut w, "Characters", self.characters);
        write_optional_u32(&mut w, "CharactersWithSpaces", self.characters_with_spaces);
        write_optional_u32(&mut w, "Lines", self.lines);
        write_optional_u32(&mut w, "Paragraphs", self.paragraphs);
        write_optional_u32(&mut w, "Slides", self.slides);
        write_optional_u32(&mut w, "Notes", self.notes);
        write_optional_u32(&mut w, "HiddenSlides", self.hidden_slides);

        w.write_event(Event::End(BytesEnd::new("Properties")))
            .expect("write end root");

        w.into_inner()
    }
}

fn write_optional_u32(w: &mut Writer<Vec<u8>>, tag: &str, value: Option<u32>) {
    if let Some(v) = value {
        let s = v.to_string();
        w.write_event(Event::Start(BytesStart::new(tag)))
            .expect("write start");
        w.write_event(Event::Text(BytesText::new(&s)))
            .expect("write text");
        w.write_event(Event::End(BytesEnd::new(tag)))
            .expect("write end");
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const SAMPLE_CORE: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties
    xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:dcterms="http://purl.org/dc/terms/"
    xmlns:dcmitype="http://purl.org/dc/dcmitype/"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Quarterly Report</dc:title>
  <dc:subject>Q4 2024 Financial Summary</dc:subject>
  <dc:creator>Jane Smith</dc:creator>
  <cp:keywords>finance; quarterly; report</cp:keywords>
  <cp:category>Report</cp:category>
  <cp:lastModifiedBy>John Doe</cp:lastModifiedBy>
  <cp:revision>4</cp:revision>
  <cp:contentStatus>Final</cp:contentStatus>
  <dcterms:created xsi:type="dcterms:W3CDTF">2024-10-01T09:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2024-12-22T16:45:00Z</dcterms:modified>
</cp:coreProperties>"#;

    #[test]
    fn parse_core_properties() {
        let props = CoreProperties::parse(SAMPLE_CORE).unwrap();
        assert_eq!(props.title.as_deref(), Some("Quarterly Report"));
        assert_eq!(props.creator.as_deref(), Some("Jane Smith"));
        assert_eq!(props.keywords.as_deref(), Some("finance; quarterly; report"));
        assert_eq!(props.revision.as_deref(), Some("4"));
        assert_eq!(props.created.as_deref(), Some("2024-10-01T09:00:00Z"));
        assert_eq!(props.content_status.as_deref(), Some("Final"));
    }

    #[test]
    fn core_properties_round_trip() {
        let original = CoreProperties {
            title: Some("Test Doc".to_string()),
            creator: Some("Test Author".to_string()),
            created: Some("2024-01-01T00:00:00Z".to_string()),
            ..Default::default()
        };
        let xml = original.serialize();
        let parsed = CoreProperties::parse(&xml).unwrap();
        assert_eq!(parsed.title, original.title);
        assert_eq!(parsed.creator, original.creator);
        assert_eq!(parsed.created, original.created);
    }

    const SAMPLE_APP: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Office Word</Application>
  <AppVersion>16.0000</AppVersion>
  <Company>Acme Corp</Company>
  <Template>Normal.dotm</Template>
  <TotalTime>45</TotalTime>
  <Pages>3</Pages>
  <Words>1250</Words>
  <Characters>7125</Characters>
  <Lines>62</Lines>
  <Paragraphs>17</Paragraphs>
</Properties>"#;

    #[test]
    fn parse_app_properties() {
        let props = AppProperties::parse(SAMPLE_APP).unwrap();
        assert_eq!(props.application.as_deref(), Some("Microsoft Office Word"));
        assert_eq!(props.app_version.as_deref(), Some("16.0000"));
        assert_eq!(props.company.as_deref(), Some("Acme Corp"));
        assert_eq!(props.pages, Some(3));
        assert_eq!(props.words, Some(1250));
        assert_eq!(props.lines, Some(62));
    }

    #[test]
    fn app_properties_round_trip() {
        let original = AppProperties {
            application: Some("office_oxide".to_string()),
            app_version: Some("0.1.0".to_string()),
            pages: Some(5),
            words: Some(2000),
            ..Default::default()
        };
        let xml = original.serialize();
        let parsed = AppProperties::parse(&xml).unwrap();
        assert_eq!(parsed.application, original.application);
        assert_eq!(parsed.pages, original.pages);
        assert_eq!(parsed.words, original.words);
    }
}
