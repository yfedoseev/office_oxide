//! Text extraction from PPT binary records.

use super::records::*;

/// Text type from TextHeaderAtom.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextType {
    Title,
    Body,
    Notes,
    Other,
    CenterBody,
    CenterTitle,
    HalfBody,
    QuarterBody,
}

impl TextType {
    pub fn from_u32(val: u32) -> Self {
        match val {
            0 => Self::Title,
            1 => Self::Body,
            2 => Self::Notes,
            3 => Self::Other,
            4 => Self::CenterBody,
            5 => Self::CenterTitle,
            6 => Self::HalfBody,
            7 => Self::QuarterBody,
            _ => Self::Other,
        }
    }
}

/// A text run extracted from a slide.
#[derive(Debug, Clone)]
pub struct TextRun {
    pub text_type: TextType,
    pub text: String,
}

/// Extract text runs from a "PowerPoint Document" stream.
///
/// This walks the record tree looking for TextHeaderAtom + TextCharsAtom/TextBytesAtom pairs.
pub fn extract_text_runs(data: &[u8]) -> Vec<TextRun> {
    let mut runs = Vec::new();
    let mut current_type = TextType::Other;

    for rec in RecordIter::new(data) {
        let rec = match rec {
            Ok(r) => r,
            Err(_) => break,
        };

        match rec.header.rec_type {
            RT_TEXT_HEADER => {
                if rec.data.len() >= 4 {
                    let t =
                        u32::from_le_bytes([rec.data[0], rec.data[1], rec.data[2], rec.data[3]]);
                    current_type = TextType::from_u32(t);
                }
            },
            RT_TEXT_CHARS => {
                // UTF-16LE text.
                let text = decode_utf16le(&rec.data);
                if !text.is_empty() {
                    runs.push(TextRun {
                        text_type: current_type,
                        text,
                    });
                }
            },
            RT_TEXT_BYTES => {
                // Compressed Latin-1 text.
                let text: String = rec.data.iter().map(|&b| b as char).collect();
                if !text.is_empty() {
                    runs.push(TextRun {
                        text_type: current_type,
                        text,
                    });
                }
            },
            RT_CSTRING => {
                // UTF-16LE string (used for some metadata).
                let text = decode_utf16le(&rec.data);
                if !text.is_empty() {
                    runs.push(TextRun {
                        text_type: current_type,
                        text,
                    });
                }
            },
            _ => {},
        }
    }

    runs
}

/// Extract text grouped by slide from the SlideListWithText containers.
///
/// Each SlideListWithText container holds text for multiple slides,
/// separated by SlidePersistAtom records.
pub fn extract_slides_text(data: &[u8]) -> Vec<SlideText> {
    let mut slides: Vec<SlideText> = Vec::new();
    let mut current_slide: Option<SlideText> = None;
    let mut in_slide_list = false;
    let mut current_type = TextType::Other;

    for rec in RecordIter::new(data) {
        let rec = match rec {
            Ok(r) => r,
            Err(_) => break,
        };

        match rec.header.rec_type {
            RT_SLIDE_LIST_WITH_TEXT => {
                in_slide_list = true;
            },
            RT_SLIDE_PERSIST_ATOM if in_slide_list => {
                // New slide starts.
                if let Some(slide) = current_slide.take() {
                    if !slide.text_runs.is_empty() {
                        slides.push(slide);
                    }
                }
                current_slide = Some(SlideText {
                    text_runs: Vec::new(),
                });
            },
            RT_TEXT_HEADER if in_slide_list => {
                if rec.data.len() >= 4 {
                    let t =
                        u32::from_le_bytes([rec.data[0], rec.data[1], rec.data[2], rec.data[3]]);
                    current_type = TextType::from_u32(t);
                }
            },
            RT_TEXT_CHARS if in_slide_list => {
                let text = decode_utf16le(&rec.data);
                if !text.is_empty() {
                    if let Some(slide) = &mut current_slide {
                        slide.text_runs.push(TextRun {
                            text_type: current_type,
                            text,
                        });
                    }
                }
            },
            RT_TEXT_BYTES if in_slide_list => {
                let text: String = rec.data.iter().map(|&b| b as char).collect();
                if !text.is_empty() {
                    if let Some(slide) = &mut current_slide {
                        slide.text_runs.push(TextRun {
                            text_type: current_type,
                            text,
                        });
                    }
                }
            },
            _ => {},
        }
    }

    // Push last slide.
    if let Some(slide) = current_slide {
        if !slide.text_runs.is_empty() {
            slides.push(slide);
        }
    }

    // If no SlideListWithText structure found, fall back to all text runs.
    if slides.is_empty() {
        let runs = extract_text_runs(data);
        if !runs.is_empty() {
            slides.push(SlideText { text_runs: runs });
        }
    }

    slides
}

/// Text content of a single slide.
#[derive(Debug, Clone)]
pub struct SlideText {
    pub text_runs: Vec<TextRun>,
}

fn decode_utf16le(data: &[u8]) -> String {
    let chars: Vec<u16> = data
        .chunks_exact(2)
        .map(|c| u16::from_le_bytes([c[0], c[1]]))
        .collect();
    String::from_utf16_lossy(&chars)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn make_atom(rec_type: u16, instance: u16, data: &[u8]) -> Vec<u8> {
        let ver_instance: u16 = instance << 4;
        let mut buf = Vec::new();
        buf.extend_from_slice(&ver_instance.to_le_bytes());
        buf.extend_from_slice(&rec_type.to_le_bytes());
        buf.extend_from_slice(&(data.len() as u32).to_le_bytes());
        buf.extend_from_slice(data);
        buf
    }

    fn make_container(rec_type: u16, instance: u16, children: &[u8]) -> Vec<u8> {
        let ver_instance: u16 = (instance << 4) | 0x0F;
        let mut buf = Vec::new();
        buf.extend_from_slice(&ver_instance.to_le_bytes());
        buf.extend_from_slice(&rec_type.to_le_bytes());
        buf.extend_from_slice(&(children.len() as u32).to_le_bytes());
        buf.extend_from_slice(children);
        buf
    }

    #[test]
    fn extract_text_chars() {
        // TextHeaderAtom(type=0=Title) + TextCharsAtom("Hi")
        let mut stream = make_atom(RT_TEXT_HEADER, 0, &0u32.to_le_bytes());
        // "Hi" in UTF-16LE
        stream.extend(make_atom(RT_TEXT_CHARS, 0, &[0x48, 0x00, 0x69, 0x00]));
        let runs = extract_text_runs(&stream);
        assert_eq!(runs.len(), 1);
        assert_eq!(runs[0].text, "Hi");
        assert_eq!(runs[0].text_type, TextType::Title);
    }

    #[test]
    fn extract_text_bytes() {
        let mut stream = make_atom(RT_TEXT_HEADER, 0, &1u32.to_le_bytes()); // Body
        stream.extend(make_atom(RT_TEXT_BYTES, 0, b"Hello World"));
        let runs = extract_text_runs(&stream);
        assert_eq!(runs.len(), 1);
        assert_eq!(runs[0].text, "Hello World");
        assert_eq!(runs[0].text_type, TextType::Body);
    }

    #[test]
    fn extract_multiple_runs() {
        let mut stream = make_atom(RT_TEXT_HEADER, 0, &0u32.to_le_bytes());
        stream.extend(make_atom(RT_TEXT_BYTES, 0, b"Title"));
        stream.extend(make_atom(RT_TEXT_HEADER, 0, &1u32.to_le_bytes()));
        stream.extend(make_atom(RT_TEXT_BYTES, 0, b"Body text"));
        let runs = extract_text_runs(&stream);
        assert_eq!(runs.len(), 2);
        assert_eq!(runs[0].text, "Title");
        assert_eq!(runs[1].text, "Body text");
    }

    #[test]
    fn extract_slides_from_slide_list() {
        // Build a SlideListWithText container with 2 slides.
        let mut children = Vec::new();
        // Slide 1
        children.extend(make_atom(RT_SLIDE_PERSIST_ATOM, 0, &[0u8; 20]));
        children.extend(make_atom(RT_TEXT_HEADER, 0, &0u32.to_le_bytes()));
        children.extend(make_atom(RT_TEXT_BYTES, 0, b"Slide 1 Title"));
        // Slide 2
        children.extend(make_atom(RT_SLIDE_PERSIST_ATOM, 0, &[0u8; 20]));
        children.extend(make_atom(RT_TEXT_HEADER, 0, &0u32.to_le_bytes()));
        children.extend(make_atom(RT_TEXT_BYTES, 0, b"Slide 2 Title"));

        let stream = make_container(RT_SLIDE_LIST_WITH_TEXT, 0, &children);
        let slides = extract_slides_text(&stream);
        assert_eq!(slides.len(), 2);
        assert_eq!(slides[0].text_runs[0].text, "Slide 1 Title");
        assert_eq!(slides[1].text_runs[0].text, "Slide 2 Title");
    }

    #[test]
    fn text_type_variants() {
        assert_eq!(TextType::from_u32(0), TextType::Title);
        assert_eq!(TextType::from_u32(1), TextType::Body);
        assert_eq!(TextType::from_u32(2), TextType::Notes);
        assert_eq!(TextType::from_u32(99), TextType::Other);
    }

    #[test]
    fn decode_utf16le_basic() {
        let data = [0x41, 0x00, 0x42, 0x00, 0x43, 0x00]; // "ABC"
        assert_eq!(decode_utf16le(&data), "ABC");
    }

    #[test]
    fn fallback_when_no_slide_list() {
        // Just raw text atoms without SlideListWithText.
        let mut stream = make_atom(RT_TEXT_HEADER, 0, &0u32.to_le_bytes());
        stream.extend(make_atom(RT_TEXT_BYTES, 0, b"Fallback text"));
        let slides = extract_slides_text(&stream);
        assert_eq!(slides.len(), 1);
        assert_eq!(slides[0].text_runs[0].text, "Fallback text");
    }
}
