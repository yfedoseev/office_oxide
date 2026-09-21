//! High-level PPT document API.

use std::io::{Read, Seek};

use crate::cfb::{CfbReader, SummaryProperties, parse_summary_information};

use super::error::{PptError, Result};
use super::images::{PptImage, extract_images};
use super::text::{SlideText, TextType, extract_slides_text};

/// A parsed legacy PowerPoint document.
#[derive(Debug)]
pub struct PptDocument {
    /// Text content extracted from each slide.
    pub slides: Vec<SlideText>,
    /// The `Pictures` stream, decoded into `images` on first request —
    /// see `DocDocument::data_stream` for why.
    pictures_stream: Vec<u8>,
    images: std::sync::OnceLock<Vec<PptImage>>,
    has_macros: bool,
    /// Title/author/subject/keywords/comments/dates from the
    /// `\x05SummaryInformation` OLE property-set stream every real `.ppt`
    /// carries by default — parsed and then never read anywhere in the
    /// crate before.
    summary_properties: Option<SummaryProperties>,
}

impl PptDocument {
    /// Open a PPT file from a reader.
    pub fn from_reader<R: Read + Seek>(reader: R) -> Result<Self> {
        let mut cfb = CfbReader::new(reader)?;
        let has_macros = cfb.has_root_entry("_VBA_PROJECT");
        let summary_properties = cfb
            .open_stream("\u{5}SummaryInformation")
            .ok()
            .and_then(|data| parse_summary_information(&data));

        // Without the main stream there is no presentation to read. This
        // used to return an empty document with `Ok`, indistinguishable
        // from a deck that simply has no text — the same silent-empty shape
        // `.doc` and `.xls` already refuse for their own missing streams.
        // A PowerPoint 95 "dual storage" file keeps its PowerPoint 95
        // stream at the root and the PowerPoint 97 rendition of the same
        // deck under the `PP97_DUALSTORAGE` *storage*. That one is the
        // stream this parser reads, so it is looked up first — asking for
        // `PP97_DUALSTORAGE` as a root-level stream name never matched,
        // and the root PPT 95 stream then parsed as an empty deck.
        let dual = cfb
            .open_stream_by_path("PP97_DUALSTORAGE/PowerPoint Document")
            .ok();
        let (stream, current_user) = match dual {
            Some(stream) => (
                stream,
                cfb.open_stream_by_path("PP97_DUALSTORAGE/Current User")
                    .ok(),
            ),
            None => (
                cfb.open_stream("PowerPoint Document").map_err(|_| {
                    PptError::MissingStream(
                        "neither PowerPoint Document nor PP97_DUALSTORAGE/PowerPoint Document stream found".into(),
                    )
                })?,
                cfb.open_stream("Current User").ok(),
            ),
        };
        let slides = extract_slides_text(&stream, current_user.as_deref());

        // The Pictures stream (if present) holds the images; decoded lazily.
        let pictures_stream = cfb.open_stream("Pictures").unwrap_or_default();

        Ok(Self {
            slides,
            pictures_stream,
            images: std::sync::OnceLock::new(),
            has_macros,
            summary_properties,
        })
    }

    /// `true` when the file carries a `_VBA_PROJECT` storage — a cheap
    /// macro-presence signal, no VBA interpretation.
    pub fn has_macros(&self) -> bool {
        self.has_macros
    }

    /// Title/author/subject/keywords/comments/dates from the file's
    /// `\x05SummaryInformation` OLE property set, when present and
    /// well-formed.
    pub fn summary_properties(&self) -> Option<&SummaryProperties> {
        self.summary_properties.as_ref()
    }

    /// Open a PPT file from a path.
    pub fn open<P: AsRef<std::path::Path>>(path: P) -> Result<Self> {
        let file = std::fs::File::open(path)?;
        Self::from_reader(file)
    }

    /// Get all extracted images.
    pub fn images(&self) -> &[PptImage] {
        self.images
            .get_or_init(|| extract_images(&self.pictures_stream))
    }

    /// Extract plain text.
    pub fn plain_text(&self) -> String {
        let mut out = String::new();
        for (i, slide) in self.slides.iter().enumerate() {
            if i > 0 {
                out.push('\n');
            }
            for run in &slide.text_runs {
                if run.text_type != TextType::Notes {
                    // A text atom keeps `\r` between its paragraphs and
                    // `\x0B` for a soft return; both are line ends here,
                    // not characters to hand to the caller.
                    out.push_str(&run.text.replace(['\r', '\u{b}'], "\n"));
                    out.push('\n');
                }
            }
        }
        out
    }

    /// Convert to markdown.
    pub fn to_markdown(&self) -> String {
        let mut out = String::new();
        for (i, slide) in self.slides.iter().enumerate() {
            if i > 0 {
                out.push('\n');
            }
            out.push_str(&format!("## Slide {}\n\n", i + 1));

            for run in &slide.text_runs {
                match run.text_type {
                    TextType::Title | TextType::CenterTitle => {
                        out.push_str("### ");
                        out.push_str(&markdown_run_text(&run.text));
                        out.push_str("\n\n");
                    },
                    TextType::Notes => {
                        // Skip notes in main content.
                    },
                    _ => {
                        out.push_str(&markdown_run_text(&run.text));
                        out.push_str("\n\n");
                    },
                }
            }
        }
        out
    }
}

impl crate::core::OfficeDocument for PptDocument {
    fn plain_text(&self) -> String {
        self.plain_text()
    }

    fn to_markdown(&self) -> String {
        self.to_markdown()
    }
}

/// One text atom's markdown: paragraph separators (`\r`) become
/// paragraph breaks, soft returns (`\x0B`) hard line breaks, and the text
/// itself is escaped so it cannot read as markdown.
fn markdown_run_text(text: &str) -> String {
    text.split('\r')
        .map(|para| {
            para.split('\u{b}')
                .map(crate::core::markdown::escape_text)
                .collect::<Vec<_>>()
                .join("  \n")
        })
        .collect::<Vec<_>>()
        .join("\n\n")
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ppt::text::TextRun;

    /// One directory entry of a hand-built CFB: name, `1` storage / `2`
    /// stream, right-sibling index, child index, and stream bytes (a
    /// storage's are ignored). Entry 0 is the root; stream data is laid
    /// out one sector per stream in entry order.
    struct Entry {
        name: &'static str,
        kind: u8,
        right: u32,
        child: u32,
        data: Vec<u8>,
    }

    /// Assemble a minimal 512-byte-sector CFB (v3, no mini stream, no
    /// DIFAT beyond the header) from `entries`.
    fn build_cfb(entries: &[Entry]) -> Vec<u8> {
        const END_OF_CHAIN: u32 = 0xFFFF_FFFE;
        const FAT_SECT: u32 = 0xFFFF_FFFD;
        const FREE_SECT: u32 = 0xFFFF_FFFF;
        const NO_ENTRY: u32 = 0xFFFF_FFFF;
        let streams: Vec<usize> = entries
            .iter()
            .enumerate()
            .filter(|(_, e)| e.kind == 2)
            .map(|(i, _)| i)
            .collect();
        // sector 0: directory, sector 1: FAT, then one sector per stream.
        let sectors = 2 + streams.len();
        let mut file = vec![0u8; 512 * (1 + sectors)];
        file[0..8].copy_from_slice(&[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
        file[0x18..0x1A].copy_from_slice(&0x003Eu16.to_le_bytes());
        file[0x1A..0x1C].copy_from_slice(&3u16.to_le_bytes());
        file[0x1C..0x1E].copy_from_slice(&0xFFFEu16.to_le_bytes());
        file[0x1E..0x20].copy_from_slice(&9u16.to_le_bytes());
        file[0x20..0x22].copy_from_slice(&6u16.to_le_bytes());
        file[0x2C..0x30].copy_from_slice(&1u32.to_le_bytes());
        file[0x30..0x34].copy_from_slice(&0u32.to_le_bytes());
        file[0x38..0x3C].copy_from_slice(&4096u32.to_le_bytes());
        file[0x3C..0x40].copy_from_slice(&END_OF_CHAIN.to_le_bytes());
        file[0x44..0x48].copy_from_slice(&END_OF_CHAIN.to_le_bytes());
        file[0x4C..0x50].copy_from_slice(&1u32.to_le_bytes());
        for i in 1..109 {
            file[0x4C + i * 4..0x50 + i * 4].copy_from_slice(&FREE_SECT.to_le_bytes());
        }
        let fat = 1024;
        let mut fat_entries = vec![END_OF_CHAIN, FAT_SECT];
        for (n, e) in entries.iter().enumerate() {
            let buf = &mut file[512 + n * 128..512 + (n + 1) * 128];
            for (i, unit) in e.name.encode_utf16().enumerate() {
                buf[i * 2..i * 2 + 2].copy_from_slice(&unit.to_le_bytes());
            }
            let name_size = ((e.name.encode_utf16().count() + 1) * 2) as u16;
            buf[0x40..0x42].copy_from_slice(&name_size.to_le_bytes());
            buf[0x42] = if n == 0 { 5 } else { e.kind };
            buf[0x43] = 1;
            buf[0x44..0x48].copy_from_slice(&NO_ENTRY.to_le_bytes());
            buf[0x48..0x4C].copy_from_slice(&e.right.to_le_bytes());
            buf[0x4C..0x50].copy_from_slice(&e.child.to_le_bytes());
            let (start, size) = if e.kind == 2 {
                let sector = 2 + streams.iter().position(|&s| s == n).unwrap();
                let off = 512 + sector * 512;
                file[off..off + e.data.len()].copy_from_slice(&e.data);
                fat_entries.push(END_OF_CHAIN);
                (sector as u32, e.data.len() as u32)
            } else {
                (END_OF_CHAIN, 0)
            };
            let buf = &mut file[512 + n * 128..512 + (n + 1) * 128];
            buf[0x74..0x78].copy_from_slice(&start.to_le_bytes());
            buf[0x78..0x7C].copy_from_slice(&size.to_le_bytes());
        }
        for i in 0..128 {
            let v = fat_entries.get(i).copied().unwrap_or(FREE_SECT);
            file[fat + i * 4..fat + i * 4 + 4].copy_from_slice(&v.to_le_bytes());
        }
        file
    }

    const NO_ENTRY: u32 = 0xFFFF_FFFF;

    /// A CFB whose root holds only a `Current User` stream — the shape a
    /// container with a broken directory tree presents once the reader
    /// walks the tree properly instead of scanning the flat array.
    fn cfb_without_main_stream() -> Vec<u8> {
        build_cfb(&[
            Entry {
                name: "Root Entry",
                kind: 5,
                right: NO_ENTRY,
                child: 1,
                data: vec![],
            },
            Entry {
                name: "Current User",
                kind: 2,
                right: NO_ENTRY,
                child: NO_ENTRY,
                data: vec![0; 8],
            },
        ])
    }

    /// The smallest record stream `extract_slides_text` reads text from: a
    /// `Slide` container holding a `ClientTextbox` with a `TextBytesAtom`.
    fn slide_stream(text: &str) -> Vec<u8> {
        fn rec(rec_type: u16, ver: u16, data: &[u8]) -> Vec<u8> {
            let mut b = ver.to_le_bytes().to_vec();
            b.extend_from_slice(&rec_type.to_le_bytes());
            b.extend_from_slice(&(data.len() as u32).to_le_bytes());
            b.extend_from_slice(data);
            b
        }
        let mut textbox = rec(crate::ppt::records::RT_TEXT_HEADER, 0, &0u32.to_le_bytes());
        textbox.extend(rec(crate::ppt::records::RT_TEXT_BYTES, 0, text.as_bytes()));
        let textbox = rec(0xF00D, 0x000F, &textbox);
        rec(crate::ppt::records::RT_SLIDE, 0x000F, &textbox)
    }

    /// A legacy deck's macros live in a root-level `_VBA_PROJECT`
    /// storage — a storage, not a stream, which is what
    /// `has_root_entry` must see.
    #[test]
    fn test_vba_project_storage_sets_has_macros() {
        let bytes = build_cfb(&[
            Entry {
                name: "Root Entry",
                kind: 5,
                right: NO_ENTRY,
                child: 1,
                data: vec![],
            },
            Entry {
                name: "PowerPoint Document",
                kind: 2,
                right: 2,
                child: NO_ENTRY,
                data: slide_stream("Slide text"),
            },
            Entry {
                name: "_VBA_PROJECT",
                kind: 1,
                right: NO_ENTRY,
                child: NO_ENTRY,
                data: vec![],
            },
        ]);
        let doc = PptDocument::from_reader(std::io::Cursor::new(bytes)).expect("opens");
        assert!(doc.has_macros());
        assert!(crate::convert_ppt::ppt_to_ir(&doc).metadata.has_macros);
    }

    /// The `Pictures` stream was decoded into images at `open()`, so
    /// `plain_text()` on a picture-heavy deck paid for pictures it never
    /// emits (12 % of a 17 MB corpus file). It is decoded on first
    /// request now, and still yields the same pictures.
    #[test]
    fn test_pictures_decode_lazily_on_first_request() {
        let png_body = b"\x89PNG\r\n\x1a\nIHDRfakebody";
        let mut blip = Vec::new();
        blip.extend_from_slice(&0u16.to_le_bytes());
        blip.extend_from_slice(&0xF01Eu16.to_le_bytes());
        blip.extend_from_slice(&((17 + png_body.len()) as u32).to_le_bytes());
        blip.extend_from_slice(&[0u8; 17]);
        blip.extend_from_slice(png_body);
        let bytes = build_cfb(&[
            Entry {
                name: "Root Entry",
                kind: 5,
                right: NO_ENTRY,
                child: 1,
                data: vec![],
            },
            Entry {
                name: "PowerPoint Document",
                kind: 2,
                right: 2,
                child: NO_ENTRY,
                data: slide_stream("Slide text"),
            },
            Entry {
                name: "Pictures",
                kind: 2,
                right: NO_ENTRY,
                child: NO_ENTRY,
                data: blip,
            },
        ]);
        let doc = PptDocument::from_reader(std::io::Cursor::new(bytes)).expect("opens");
        assert!(doc.images.get().is_none(), "nothing decoded at open()");
        assert!(doc.plain_text().contains("Slide text"));
        assert!(doc.images.get().is_none(), "plain_text() does not decode pictures");
        let images = doc.images();
        assert_eq!(images.len(), 1);
        assert_eq!(images[0].data, png_body);
    }

    /// Regression: a PowerPoint 95 "dual storage" file keeps a PowerPoint
    /// 97 rendition of the deck under the `PP97_DUALSTORAGE` storage, and
    /// a PowerPoint 95 stream (which this parser cannot read) at the root.
    /// Looking for `PP97_DUALSTORAGE` as a root-level *stream* never
    /// matched, so the root stream was parsed instead and a 105-slide
    /// corpus deck came back with zero sections and `Ok`.
    #[test]
    fn test_dual_storage_file_reads_the_pp97_rendition() {
        let bytes = build_cfb(&[
            Entry {
                name: "Root Entry",
                kind: 5,
                right: NO_ENTRY,
                child: 1,
                data: vec![],
            },
            // The storage is the root's child; the root-level PPT 95
            // stream is its right sibling.
            Entry {
                name: "PP97_DUALSTORAGE",
                kind: 1,
                right: 2,
                child: 3,
                data: vec![],
            },
            Entry {
                name: "PowerPoint Document",
                kind: 2,
                right: NO_ENTRY,
                child: NO_ENTRY,
                data: b"PowerPoint 95 records this parser does not read".to_vec(),
            },
            Entry {
                name: "PowerPoint Document",
                kind: 2,
                right: NO_ENTRY,
                child: NO_ENTRY,
                data: slide_stream("Text from the PP97 rendition"),
            },
        ]);
        let doc = PptDocument::from_reader(std::io::Cursor::new(bytes)).expect("opens");
        let text = doc.plain_text();
        assert!(
            text.contains("Text from the PP97 rendition"),
            "the PP97_DUALSTORAGE rendition must be the one read: {text:?}"
        );
    }

    /// Regression: a container with no `PowerPoint Document` stream (a
    /// fuzzer-corrupted directory tree left it unreachable) opened as an
    /// empty deck with `Ok`, so a file that could not be read looked like
    /// one that simply had no slides. It must fail like `.doc`/`.xls` do.
    #[test]
    fn test_missing_main_stream_is_an_error_not_an_empty_deck() {
        let bytes = cfb_without_main_stream();
        let err = PptDocument::from_reader(std::io::Cursor::new(bytes))
            .expect_err("a .ppt without its main stream must not open");
        assert!(matches!(err, PptError::MissingStream(_)), "expected MissingStream, got {err:?}");
    }

    #[test]
    fn test_plain_text_basic() {
        let doc = PptDocument {
            pictures_stream: Vec::new(),
            images: std::sync::OnceLock::new(),
            has_macros: false,
            summary_properties: None,
            slides: vec![
                SlideText {
                    text_runs: vec![
                        TextRun {
                            text_type: TextType::Title,
                            text: "Welcome".into(),
                            hyperlink: None,
                            ..Default::default()
                        },
                        TextRun {
                            text_type: TextType::Body,
                            text: "Hello world".into(),
                            hyperlink: None,
                            ..Default::default()
                        },
                    ],
                    ..Default::default()
                },
                SlideText {
                    text_runs: vec![TextRun {
                        text_type: TextType::Title,
                        text: "Slide 2".into(),
                        hyperlink: None,
                        ..Default::default()
                    }],
                    ..Default::default()
                },
            ],
        };
        let text = doc.plain_text();
        assert!(text.contains("Welcome"));
        assert!(text.contains("Hello world"));
        assert!(text.contains("Slide 2"));
    }

    #[test]
    fn test_markdown_basic() {
        let doc = PptDocument {
            pictures_stream: Vec::new(),
            images: std::sync::OnceLock::new(),
            has_macros: false,
            summary_properties: None,
            slides: vec![SlideText {
                text_runs: vec![
                    TextRun {
                        text_type: TextType::Title,
                        text: "My Title".into(),
                        hyperlink: None,
                        ..Default::default()
                    },
                    TextRun {
                        text_type: TextType::Body,
                        text: "Content here".into(),
                        hyperlink: None,
                        ..Default::default()
                    },
                ],
                ..Default::default()
            }],
        };
        let md = doc.to_markdown();
        assert!(md.contains("## Slide 1"));
        assert!(md.contains("### My Title"));
        assert!(md.contains("Content here"));
    }

    #[test]
    fn test_notes_excluded_from_plain_text() {
        let doc = PptDocument {
            pictures_stream: Vec::new(),
            images: std::sync::OnceLock::new(),
            has_macros: false,
            summary_properties: None,
            slides: vec![SlideText {
                text_runs: vec![
                    TextRun {
                        text_type: TextType::Title,
                        text: "Title".into(),
                        hyperlink: None,
                        ..Default::default()
                    },
                    TextRun {
                        text_type: TextType::Notes,
                        text: "Speaker notes".into(),
                        hyperlink: None,
                        ..Default::default()
                    },
                ],
                ..Default::default()
            }],
        };
        let text = doc.plain_text();
        assert!(text.contains("Title"));
        assert!(!text.contains("Speaker notes"));
    }

    fn make_slide(runs: Vec<(TextType, &str)>) -> SlideText {
        SlideText {
            text_runs: runs
                .into_iter()
                .map(|(t, s)| TextRun {
                    text_type: t,
                    text: s.to_string(),
                    hyperlink: None,
                    ..Default::default()
                })
                .collect(),
            ..Default::default()
        }
    }

    #[test]
    fn test_ir_empty_doc_has_no_sections() {
        let doc = PptDocument {
            pictures_stream: Vec::new(),
            images: std::sync::OnceLock::new(),
            slides: Vec::new(),
            has_macros: false,
            summary_properties: None,
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        assert!(ir.sections.is_empty());
        assert!(ir.metadata.title.is_none());
    }

    /// A vertical tab is a line break inside a paragraph ([MS-PPT] soft
    /// return); the converter left it in the text, so the two lines fused
    /// into one word on every IR surface while `plain_text()` broke them.
    #[test]
    fn test_ir_soft_return_is_a_line_break_not_a_glued_word() {
        use crate::ir::{Element, InlineContent};
        let doc = PptDocument {
            pictures_stream: Vec::new(),
            images: std::sync::OnceLock::new(),
            has_macros: false,
            summary_properties: None,
            slides: vec![make_slide(vec![(
                TextType::Body,
                "Four Upload\u{b}Stations",
            )])],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        let para = ir.sections[0]
            .elements
            .iter()
            .find_map(|e| match e {
                Element::Paragraph(p) => Some(p),
                _ => None,
            })
            .expect("a body paragraph");
        assert!(
            para.content
                .iter()
                .any(|c| matches!(c, InlineContent::LineBreak)),
            "{:?}",
            para.content
        );
        let text = ir.plain_text();
        assert!(text.contains("Four Upload\nStations"), "{text:?}");
        assert!(!ir.to_html().contains("UploadStations"));
    }

    /// The direct surfaces hand no raw `\r` or `\x0B` to the caller: a
    /// text atom's paragraph separator and soft return are line ends in
    /// `plain_text()` and paragraph/line breaks in `to_markdown()`.
    #[test]
    fn test_direct_surfaces_normalise_paragraph_separators() {
        let doc = PptDocument {
            pictures_stream: Vec::new(),
            images: std::sync::OnceLock::new(),
            has_macros: false,
            summary_properties: None,
            slides: vec![make_slide(vec![(
                TextType::Body,
                "First paragraph\rFour Upload\u{b}Stations",
            )])],
        };
        let text = doc.plain_text();
        assert_eq!(text, "First paragraph\nFour Upload\nStations\n");
        let md = doc.to_markdown();
        assert!(md.contains("First paragraph\n\nFour Upload  \nStations"), "{md:?}");
        assert!(!md.contains('\r'));
    }

    /// A title placeholder holding several paragraphs: the first is the
    /// heading, the rest are body paragraphs, and none are fused.
    #[test]
    fn test_ir_title_run_with_several_paragraphs_is_a_heading_and_paragraphs() {
        use crate::ir::Element;
        let doc = PptDocument {
            pictures_stream: Vec::new(),
            images: std::sync::OnceLock::new(),
            has_macros: false,
            summary_properties: None,
            slides: vec![make_slide(vec![(
                TextType::CenterTitle,
                "Methods & Tools\rAnalyses – national studies\rEvaluation – criteria",
            )])],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        let kinds: Vec<&str> = ir.sections[0]
            .elements
            .iter()
            .map(|e| match e {
                Element::Heading(_) => "heading",
                Element::Paragraph(_) => "paragraph",
                _ => "other",
            })
            .collect();
        assert_eq!(kinds, ["heading", "paragraph", "paragraph"]);
        assert_eq!(ir.sections[0].title.as_deref(), Some("Methods & Tools"));
        let md = ir.to_markdown();
        assert!(md.contains("# Methods & Tools"), "{md}");
        assert!(!md.contains("**"), "a heading is not also bold: {md}");
        assert!(!ir.plain_text().contains("ToolsAnalyses"));
    }

    /// A free text box (`TextType::Other`) with several paragraphs is
    /// several paragraphs, not one block with its lines fused.
    #[test]
    fn test_ir_other_text_paragraphs_are_not_fused() {
        let doc = PptDocument {
            pictures_stream: Vec::new(),
            images: std::sync::OnceLock::new(),
            has_macros: false,
            summary_properties: None,
            slides: vec![make_slide(vec![(
                TextType::Other,
                "LOASP\rChap. 14: information",
            )])],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        assert_eq!(ir.sections[0].elements.len(), 2, "{:?}", ir.sections[0].elements);
        assert!(!ir.to_html().contains("LOASPChap"));
    }

    #[test]
    fn test_ir_title_becomes_heading_and_section_title() {
        use crate::ir::Element;
        let doc = PptDocument {
            pictures_stream: Vec::new(),
            images: std::sync::OnceLock::new(),
            has_macros: false,
            summary_properties: None,
            slides: vec![make_slide(vec![(TextType::Title, "My Slide")])],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        assert_eq!(ir.metadata.title.as_deref(), Some("My Slide"));
        assert!(matches!(ir.sections[0].elements[0], Element::Heading(_)));
    }

    /// A title-slide layout's real title (`textType=6`) and
    /// subtitle (`textType=5`) must not be swapped: the document title
    /// must come from the real title text, and the subtitle must not
    /// become a bold `Heading`.
    #[test]
    fn test_ir_title_slide_title_and_subtitle_are_not_swapped() {
        use crate::ir::Element;
        let doc = PptDocument {
            pictures_stream: Vec::new(),
            images: std::sync::OnceLock::new(),
            has_macros: false,
            summary_properties: None,
            slides: vec![make_slide(vec![
                (TextType::from_u32(6), "BSE in the US"), // real title
                (TextType::from_u32(5), "Lisa A. Ferguson, DVM"), // real subtitle
            ])],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        assert_eq!(
            ir.metadata.title.as_deref(),
            Some("BSE in the US"),
            "the document title must come from the real title (textType=6), not the subtitle"
        );
        assert_eq!(ir.sections[0].title.as_deref(), Some("BSE in the US"));
        // The subtitle (CenterBody) must land as an ordinary Paragraph via
        // convert_ppt.rs's catch-all arm, never as a bold Heading.
        assert!(
            ir.sections[0]
                .elements
                .iter()
                .any(|e| matches!(e, Element::Paragraph(p)
                    if p.content.iter().any(|c| matches!(c, crate::ir::InlineContent::Text(t) if t.text == "Lisa A. Ferguson, DVM")))),
            "the subtitle must reach the IR as an ordinary paragraph"
        );
        assert!(
            !ir.sections[0]
                .elements
                .iter()
                .any(|e| matches!(e, Element::Heading(h)
                    if h.content.iter().any(|c| matches!(c, crate::ir::InlineContent::Text(t) if t.text == "Lisa A. Ferguson, DVM")))),
            "the subtitle must never become the document heading"
        );
    }

    /// A `TextRun::hyperlink` resolved from `InteractiveInfo`
    /// must reach `TextSpan::hyperlink` in the IR, for every text type
    /// that hyperlink can attach to (not just plain body paragraphs).
    #[test]
    fn test_ir_hyperlink_reaches_textspan_hyperlink() {
        use crate::ir::{Element, InlineContent};
        let doc = PptDocument {
            pictures_stream: Vec::new(),
            images: std::sync::OnceLock::new(),
            has_macros: false,
            summary_properties: None,
            slides: vec![SlideText {
                text_runs: vec![TextRun {
                    text_type: TextType::Body,
                    text: "Click here".to_string(),
                    hyperlink: Some("http://testuri.org/".to_string()),
                    ..Default::default()
                }],
                ..Default::default()
            }],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        let Element::Paragraph(p) = &ir.sections[0].elements[0] else {
            panic!("expected a paragraph, got {:?}", ir.sections[0].elements[0]);
        };
        let InlineContent::Text(span) = &p.content[0] else {
            panic!("expected text content");
        };
        assert_eq!(span.hyperlink.as_deref(), Some("http://testuri.org/"));
    }

    #[test]
    fn test_ir_center_title_treated_like_title() {
        let doc = PptDocument {
            pictures_stream: Vec::new(),
            images: std::sync::OnceLock::new(),
            has_macros: false,
            summary_properties: None,
            slides: vec![make_slide(vec![(TextType::CenterTitle, "Centered")])],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        assert_eq!(ir.sections[0].title.as_deref(), Some("Centered"));
    }

    /// The PPTX side of this was already fixed; this is the same
    /// defect on the legacy binary .ppt path — `TextType::Notes` runs must
    /// land on `Section.speaker_notes`, not leak into `elements` as
    /// ordinary (if italicized) paragraphs indistinguishable from body
    /// text to most IR consumers.
    #[test]
    fn test_ir_notes_go_to_speaker_notes_not_elements() {
        use crate::ir::{Element, InlineContent};
        let doc = PptDocument {
            pictures_stream: Vec::new(),
            images: std::sync::OnceLock::new(),
            has_macros: false,
            summary_properties: None,
            slides: vec![make_slide(vec![
                (TextType::Title, "Title"),
                (TextType::Body, "Visible body text"),
                (TextType::Notes, "Presenter-only notes"),
            ])],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        assert_eq!(
            notes_text(ir.sections[0].speaker_notes.as_deref().unwrap_or_default()),
            "Presenter-only notes",
            "notes text must reach Section::speaker_notes"
        );
        let elements_text = format!("{:?}", ir.sections[0].elements);
        assert!(
            !elements_text.contains("Presenter-only notes"),
            "notes text must not also leak into elements: {elements_text}"
        );
        // Visible content is unaffected.
        assert!(
            ir.sections[0]
                .elements
                .iter()
                .any(|e| matches!(e, Element::Paragraph(p) if matches!(&p.content[0], InlineContent::Text(t) if t.text == "Visible body text"))),
            "body text must still reach elements"
        );
    }

    /// A slide with no notes at all gets `None`, not an empty string.
    #[test]
    fn test_ir_no_notes_is_none_not_empty_string() {
        let doc = PptDocument {
            pictures_stream: Vec::new(),
            images: std::sync::OnceLock::new(),
            has_macros: false,
            summary_properties: None,
            slides: vec![make_slide(vec![(TextType::Body, "Just body text")])],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        assert!(ir.sections[0].speaker_notes.is_none());
    }

    /// Multiple `TextType::Notes` runs on one slide are joined, not just
    /// the last one kept.
    #[test]
    fn test_ir_multiple_notes_runs_are_joined() {
        let doc = PptDocument {
            pictures_stream: Vec::new(),
            images: std::sync::OnceLock::new(),
            has_macros: false,
            summary_properties: None,
            slides: vec![make_slide(vec![
                (TextType::Notes, "First note"),
                (TextType::Notes, "Second note"),
            ])],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        let notes = notes_text(ir.sections[0].speaker_notes.as_deref().unwrap_or_default());
        assert!(notes.contains("First note"), "notes: {notes:?}");
        assert!(notes.contains("Second note"), "notes: {notes:?}");
    }

    /// Flatten `Section::speaker_notes`' `Element::Paragraph`s into plain
    /// text for assertions — mirrors how `ir_render.rs` does it for real.
    fn notes_text(elements: &[crate::ir::Element]) -> String {
        use crate::ir::{Element, InlineContent};
        let mut out = String::new();
        for e in elements {
            if let Element::Paragraph(p) = e {
                for c in &p.content {
                    if let InlineContent::Text(t) = c {
                        out.push_str(&t.text);
                    }
                }
                out.push('\n');
            }
        }
        out.trim_end().to_string()
    }

    #[test]
    fn test_ir_body_half_quarter_produce_paragraphs() {
        use crate::ir::Element;
        let doc = PptDocument {
            pictures_stream: Vec::new(),
            images: std::sync::OnceLock::new(),
            has_macros: false,
            summary_properties: None,
            slides: vec![make_slide(vec![
                (TextType::Body, "Body text"),
                (TextType::HalfBody, "Half body"),
                (TextType::QuarterBody, "Quarter"),
            ])],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        assert_eq!(ir.sections[0].elements.len(), 3);
        assert!(matches!(ir.sections[0].elements[0], Element::Paragraph(_)));
    }

    #[test]
    fn test_ir_notes_produce_no_visible_elements() {
        // Superseded: notes used to become an italic Element::
        // Paragraph in `elements`; they now route to Section::speaker_notes
        // instead (see ir_notes_go_to_speaker_notes_not_elements above) and
        // a notes-only slide has no visible elements at all.
        let doc = PptDocument {
            pictures_stream: Vec::new(),
            images: std::sync::OnceLock::new(),
            has_macros: false,
            summary_properties: None,
            slides: vec![make_slide(vec![(TextType::Notes, "Speaker note")])],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        assert!(ir.sections[0].elements.is_empty());
        assert_eq!(
            notes_text(ir.sections[0].speaker_notes.as_deref().unwrap_or_default()),
            "Speaker note"
        );
    }

    #[test]
    fn test_ir_other_text_type_produces_paragraph() {
        use crate::ir::Element;
        let doc = PptDocument {
            pictures_stream: Vec::new(),
            images: std::sync::OnceLock::new(),
            has_macros: false,
            summary_properties: None,
            slides: vec![make_slide(vec![(TextType::Other, "misc text")])],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        assert!(matches!(ir.sections[0].elements[0], Element::Paragraph(_)));
    }

    #[test]
    fn test_ir_slide_without_title_gets_fallback_name() {
        let doc = PptDocument {
            pictures_stream: Vec::new(),
            images: std::sync::OnceLock::new(),
            has_macros: false,
            summary_properties: None,
            slides: vec![make_slide(vec![(TextType::Body, "content")])],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        assert_eq!(ir.sections[0].title.as_deref(), Some("Slide 1"));
    }

    #[test]
    fn test_ir_format_is_ppt() {
        let doc = PptDocument {
            pictures_stream: Vec::new(),
            images: std::sync::OnceLock::new(),
            slides: Vec::new(),
            has_macros: false,
            summary_properties: None,
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        assert_eq!(ir.metadata.format, crate::format::DocumentFormat::Ppt);
    }

    /// `SummaryInformation` fields must reach `Metadata`, and
    /// the declared title must beat the first-slide-title fallback.
    #[test]
    fn test_ir_summary_properties_reach_metadata() {
        let doc = PptDocument {
            pictures_stream: Vec::new(),
            images: std::sync::OnceLock::new(),
            has_macros: false,
            summary_properties: Some(crate::cfb::SummaryProperties {
                title: Some("Declared Title".to_string()),
                subject: Some("Declared Subject".to_string()),
                author: Some("Declared Author".to_string()),
                keywords: Some("alpha, beta".to_string()),
                comments: Some("Declared Comment".to_string()),
                created: Some("2020-01-02T03:04:05Z".to_string()),
                modified: Some("2021-06-07T08:09:10Z".to_string()),
            }),
            slides: vec![make_slide(vec![(TextType::Title, "Slide Title")])],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        assert_eq!(ir.metadata.title.as_deref(), Some("Declared Title"));
        assert_eq!(ir.metadata.author.as_deref(), Some("Declared Author"));
        assert_eq!(ir.metadata.subject.as_deref(), Some("Declared Subject"));
        assert_eq!(ir.metadata.keywords, vec!["alpha".to_string(), "beta".to_string()]);
        assert_eq!(ir.metadata.description.as_deref(), Some("Declared Comment"));
        assert_eq!(ir.metadata.created.as_deref(), Some("2020-01-02T03:04:05Z"));
        assert_eq!(ir.metadata.modified.as_deref(), Some("2021-06-07T08:09:10Z"));
    }

    /// A missing/empty title in `SummaryInformation` must not shadow the
    /// first-slide-title fallback.
    #[test]
    fn test_ir_empty_summary_title_falls_back_to_slide_title() {
        let doc = PptDocument {
            pictures_stream: Vec::new(),
            images: std::sync::OnceLock::new(),
            has_macros: false,
            summary_properties: Some(crate::cfb::SummaryProperties {
                title: Some(String::new()),
                ..Default::default()
            }),
            slides: vec![make_slide(vec![(TextType::Title, "Slide Title")])],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        assert_eq!(ir.metadata.title.as_deref(), Some("Slide Title"));
    }

    // ── Direct character/paragraph formatting ──

    /// Real `bold: Some(false)` from a `StyleTextPropAtom`
    /// must override the old synthetic "titles are always bold" default,
    /// not just be ignored in its favor.
    #[test]
    fn test_ir_real_char_formatting_overrides_synthetic_title_bold() {
        use crate::ir::{Element, InlineContent};
        use crate::ppt::{CharFormat, CharFormatSpan};

        let doc = PptDocument {
            pictures_stream: Vec::new(),
            images: std::sync::OnceLock::new(),
            has_macros: false,
            summary_properties: None,
            slides: vec![SlideText {
                text_runs: vec![TextRun {
                    text_type: TextType::Title,
                    text: "Not Bold".to_string(),
                    hyperlink: None,
                    char_formats: vec![CharFormatSpan {
                        start: 0,
                        end: 8,
                        format: CharFormat {
                            bold: Some(false),
                            ..Default::default()
                        },
                    }],
                    para_formats: Vec::new(),
                    placeholder_role: None,
                }],
                ..Default::default()
            }],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        let Element::Heading(h) = &ir.sections[0].elements[0] else {
            panic!("expected a heading, got {:?}", ir.sections[0].elements[0]);
        };
        let InlineContent::Text(span) = &h.content[0] else {
            panic!("expected text content");
        };
        assert!(
            !span.bold,
            "explicit bold:false from the file must win over the old synthetic default"
        );
    }

    /// Two `TextCFRun`s with different formatting over the
    /// same run of text must produce two separately-formatted `TextSpan`s,
    /// not one span with the first (or last) run's formatting applied to
    /// everything.
    #[test]
    fn test_ir_char_formatting_produces_multiple_spans_within_one_paragraph() {
        use crate::ir::{Element, InlineContent};
        use crate::ppt::{CharFormat, CharFormatSpan};

        let doc = PptDocument {
            pictures_stream: Vec::new(),
            images: std::sync::OnceLock::new(),
            has_macros: false,
            summary_properties: None,
            slides: vec![SlideText {
                text_runs: vec![TextRun {
                    text_type: TextType::Body,
                    text: "ABCDEF".to_string(),
                    hyperlink: None,
                    char_formats: vec![
                        CharFormatSpan {
                            start: 0,
                            end: 3,
                            format: CharFormat {
                                bold: Some(true),
                                ..Default::default()
                            },
                        },
                        CharFormatSpan {
                            start: 3,
                            end: 6,
                            format: CharFormat {
                                italic: Some(true),
                                ..Default::default()
                            },
                        },
                    ],
                    para_formats: Vec::new(),
                    placeholder_role: None,
                }],
                ..Default::default()
            }],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        let Element::Paragraph(p) = &ir.sections[0].elements[0] else {
            panic!("expected a paragraph, got {:?}", ir.sections[0].elements[0]);
        };
        assert_eq!(p.content.len(), 2, "one span per formatting run: {:?}", p.content);
        let InlineContent::Text(first) = &p.content[0] else {
            panic!()
        };
        let InlineContent::Text(second) = &p.content[1] else {
            panic!()
        };
        assert_eq!(first.text, "ABC");
        assert!(first.bold);
        assert!(!first.italic);
        assert_eq!(second.text, "DEF");
        assert!(second.italic);
        assert!(!second.bold);
    }

    /// A `TextPFRun`'s alignment must reach `Paragraph::alignment`.
    #[test]
    fn test_ir_paragraph_alignment_reaches_ir() {
        use crate::ir::{Element, ParagraphAlignment};
        use crate::ppt::{ParaFormat, ParaFormatSpan};

        let doc = PptDocument {
            pictures_stream: Vec::new(),
            images: std::sync::OnceLock::new(),
            has_macros: false,
            summary_properties: None,
            slides: vec![SlideText {
                text_runs: vec![TextRun {
                    text_type: TextType::Body,
                    text: "Centered".to_string(),
                    hyperlink: None,
                    char_formats: Vec::new(),
                    para_formats: vec![ParaFormatSpan {
                        start: 0,
                        end: 8,
                        format: ParaFormat { alignment: Some(1) }, // Tx_ALIGNCenter
                    }],
                    placeholder_role: None,
                }],
                ..Default::default()
            }],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        let Element::Paragraph(p) = &ir.sections[0].elements[0] else {
            panic!("expected a paragraph, got {:?}", ir.sections[0].elements[0]);
        };
        assert_eq!(p.alignment, Some(ParagraphAlignment::Center));
    }

    /// A run's resolved placeholder role must reach
    /// `Paragraph::placeholder_role` in the IR.
    #[test]
    fn test_ir_placeholder_role_reaches_the_paragraph() {
        use crate::ir::Element;

        let doc = PptDocument {
            pictures_stream: Vec::new(),
            images: std::sync::OnceLock::new(),
            has_macros: false,
            summary_properties: None,
            slides: vec![SlideText {
                text_runs: vec![TextRun {
                    text_type: TextType::Other,
                    text: "Footer text".to_string(),
                    hyperlink: None,
                    char_formats: Vec::new(),
                    para_formats: Vec::new(),
                    placeholder_role: Some("ftr".to_string()),
                }],
                ..Default::default()
            }],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        let Element::Paragraph(p) = &ir.sections[0].elements[0] else {
            panic!("expected a paragraph, got {:?}", ir.sections[0].elements[0]);
        };
        assert_eq!(p.placeholder_role.as_deref(), Some("ftr"));
    }

    /// A lone `\r` inside one text atom (the standard PPT97
    /// multi-bullet layout, per [MS-PPT]'s own worked example) must split
    /// into separate `Paragraph` elements, not survive as a literal `\r`
    /// embedded in one giant paragraph (`str::lines()` alone doesn't split
    /// on a bare `\r`).
    #[test]
    fn test_ir_bare_cr_splits_into_multiple_paragraphs() {
        use crate::ir::{Element, InlineContent};

        let doc = PptDocument {
            pictures_stream: Vec::new(),
            images: std::sync::OnceLock::new(),
            has_macros: false,
            summary_properties: None,
            slides: vec![SlideText {
                text_runs: vec![TextRun {
                    text_type: TextType::Body,
                    text: "a sunny day\rthe blue sky\rsome green grass".to_string(),
                    hyperlink: None,
                    ..Default::default()
                }],
                ..Default::default()
            }],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        assert_eq!(ir.sections[0].elements.len(), 3, "{:?}", ir.sections[0].elements);
        let texts: Vec<&str> = ir.sections[0]
            .elements
            .iter()
            .map(|e| {
                let Element::Paragraph(p) = e else {
                    panic!("expected paragraphs")
                };
                let InlineContent::Text(t) = &p.content[0] else {
                    panic!()
                };
                t.text.as_str()
            })
            .collect();
        assert_eq!(texts, ["a sunny day", "the blue sky", "some green grass"]);
        assert!(!texts.iter().any(|t| t.contains('\r')), "no leftover literal \\r: {texts:?}");
    }

    /// A reconstructed grid-of-shapes table must reach the
    /// IR as `Element::Table`, with each cell's text in the right slot.
    #[test]
    fn test_ir_reconstructed_table_becomes_element_table() {
        use crate::ir::Element;
        use crate::ppt::TableBlock;

        fn cell_run(text: &str) -> Vec<TextRun> {
            vec![TextRun {
                text_type: TextType::Other,
                text: text.to_string(),
                ..Default::default()
            }]
        }

        let doc = PptDocument {
            pictures_stream: Vec::new(),
            images: std::sync::OnceLock::new(),
            has_macros: false,
            summary_properties: None,
            slides: vec![SlideText {
                tables: vec![TableBlock {
                    rows: vec![
                        vec![cell_run("A1"), cell_run("B1")],
                        vec![cell_run("A2"), cell_run("B2")],
                    ],
                }],
                ..Default::default()
            }],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        let Element::Table(t) = &ir.sections[0].elements[0] else {
            panic!("expected a table, got {:?}", ir.sections[0].elements[0]);
        };
        assert_eq!(t.rows.len(), 2);
        assert_eq!(t.rows[0].cells.len(), 2);

        let cell_text = |r: usize, c: usize| {
            let crate::ir::Element::Paragraph(p) = &t.rows[r].cells[c].content[0] else {
                panic!()
            };
            let crate::ir::InlineContent::Text(span) = &p.content[0] else {
                panic!()
            };
            span.text.clone()
        };
        assert_eq!(cell_text(0, 0), "A1");
        assert_eq!(cell_text(0, 1), "B1");
        assert_eq!(cell_text(1, 0), "A2");
        assert_eq!(cell_text(1, 1), "B2");
    }

    /// A picture shape's own `pib`-resolved image must
    /// attach to the slide that actually contains it, not every slide's
    /// images landing on whichever slide happens to be last.
    #[test]
    fn test_ir_image_attaches_to_its_own_slide_not_the_last_one() {
        use crate::cfb::blip::{BlipFormat, BlipImage};
        use crate::ir::Element;

        let doc = PptDocument {
            pictures_stream: Vec::new(),
            images: std::sync::OnceLock::from(vec![
                BlipImage {
                    format: BlipFormat::Png,
                    data: b"PNG0".to_vec(),
                    index: 0,
                },
                BlipImage {
                    format: BlipFormat::Jpeg,
                    data: b"JPEG1".to_vec(),
                    index: 1,
                },
            ]),
            has_macros: false,
            summary_properties: None,
            slides: vec![
                SlideText {
                    image_refs: vec![0],
                    ..Default::default()
                }, // slide 1: image 0
                SlideText {
                    ..Default::default()
                }, // slide 2: no images
                SlideText {
                    image_refs: vec![1],
                    ..Default::default()
                }, // slide 3: image 1
            ],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        assert_eq!(ir.sections.len(), 3);

        let image_data = |section: &crate::ir::Section| -> Vec<Vec<u8>> {
            section
                .elements
                .iter()
                .filter_map(|e| match e {
                    Element::Image(img) => img.data.clone(),
                    _ => None,
                })
                .collect()
        };
        assert_eq!(image_data(&ir.sections[0]), vec![b"PNG0".to_vec()]);
        assert!(image_data(&ir.sections[1]).is_empty(), "slide 2 has no images of its own");
        assert_eq!(image_data(&ir.sections[2]), vec![b"JPEG1".to_vec()]);
    }

    /// An image no shape resolved via `pib` must still
    /// reach the IR (the old fallback behavior), not vanish entirely.
    #[test]
    fn test_ir_unresolved_image_still_reaches_the_ir_as_a_leftover() {
        use crate::cfb::blip::{BlipFormat, BlipImage};
        use crate::ir::Element;

        let doc = PptDocument {
            pictures_stream: Vec::new(),
            images: std::sync::OnceLock::from(vec![BlipImage {
                format: BlipFormat::Png,
                data: b"ORPHAN".to_vec(),
                index: 0,
            }]),
            has_macros: false,
            summary_properties: None,
            slides: vec![SlideText {
                ..Default::default()
            }],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        let found = ir.sections[0].elements.iter().any(|e| {
            matches!(e, Element::Image(img) if img.data.as_deref() == Some(b"ORPHAN".as_slice()))
        });
        assert!(found, "an unresolved image must not be silently dropped");
    }
}
