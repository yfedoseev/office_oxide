//! Text extraction from PPT binary records.

use super::persist::{self, PersistDirectory};
use super::records::*;

/// Guards against pathologically deep (or maliciously crafted) shape nesting.
const MAX_SHAPE_DEPTH: usize = 64;

/// Text type from TextHeaderAtom.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextType {
    /// Title placeholder text.
    Title,
    /// Body / content placeholder text.
    Body,
    /// Speaker notes text.
    Notes,
    /// Other or unclassified text.
    Other,
    /// Centered body placeholder.
    CenterBody,
    /// Centered title placeholder.
    CenterTitle,
    /// Half-size body placeholder.
    HalfBody,
    /// Quarter-size body placeholder.
    QuarterBody,
}

impl TextType {
    /// Convert a `TextHeaderAtom` type integer to a `TextType`.
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
    /// The role of this text within its slide.
    pub text_type: TextType,
    /// The decoded text content.
    pub text: String,
}

/// Extract per-slide text from a "PowerPoint Document" stream.
///
/// `current_user` is the raw "Current User" stream, if present.
///
/// PPT97 files are saved incrementally: a slide's *current* content lives in
/// a `Slide` container located via the persist object directory, not
/// necessarily wherever a naive front-to-back scan first stumbles on
/// slide-shaped data (stale, superseded copies of records are routinely left
/// behind in the stream by earlier saves). This resolves each slide through
/// that directory — see the `persist` module — falling back to weaker
/// heuristics only when a usable directory can't be built at all (e.g. a
/// minimal hand-built stream that never went through a real save cycle).
pub fn extract_slides_text(stream: &[u8], current_user: Option<&[u8]>) -> Vec<SlideText> {
    if let Some(dir) = persist::build(stream, current_user) {
        if let Some(slides) = extract_slides_via_persist(stream, &dir) {
            // A resolved-but-entirely-textless result is ambiguous: it's the
            // correct answer for a genuinely text-free deck (image-only
            // slides), but it's also what a corrupted persist chain that
            // resolved to the wrong offsets looks like. Fall through to the
            // weaker heuristics below rather than committing to it — if they
            // also come up empty, this was the right answer all along.
            if !slides.is_empty() && slides.iter().any(|s| !s.text_runs.is_empty()) {
                return slides;
            }
        }
    }

    if let Some(slide_list) = find_descendant(stream, RT_SLIDE_LIST_WITH_TEXT, SLWT_SLIDES, 0) {
        let slides = extract_slides_from_slide_list_cache(&slide_list);
        if !slides.is_empty() {
            return slides;
        }
    }

    // Last resort: no resolvable structure at all — dump whatever text atoms
    // exist anywhere in the stream as a single slide.
    let mut runs = Vec::new();
    extract_shape_text(stream, 0, &[], &mut runs);
    if runs.is_empty() {
        Vec::new()
    } else {
        vec![SlideText { text_runs: runs }]
    }
}

/// Resolve the current "Slides" list through the persist directory and
/// extract each slide's shape text from its resolved `Slide` container.
///
/// Also collects each slide's own outline-text sequence — the
/// `TextHeaderAtom`/`TextCharsAtom`/`TextBytesAtom` runs that directly follow
/// its `SlidePersistAtom` in `SlideListWithTextContainer` — because many
/// placeholder shapes don't embed their text directly at all: they hold only
/// an `OutlineTextRefAtom`, an index into that same per-slide sequence
/// ([MS-PPT] 2.4.15.6). Resolving text purely from the `Slide` container's own
/// records, without this table, silently drops that text.
///
/// Returns `None` if the `DocumentContainer` or its slide list can't be
/// resolved at all (directory present but unusable); returns `Some(vec![])`
/// if the slide list resolves but is empty.
fn extract_slides_via_persist(stream: &[u8], dir: &PersistDirectory) -> Option<Vec<SlideText>> {
    let doc_offset = dir.resolve(dir.doc_persist_id)?;
    let doc_children = bounded_container_children(stream, doc_offset, RT_DOCUMENT)?;
    let slide_list = find_child(&doc_children, RT_SLIDE_LIST_WITH_TEXT, SLWT_SLIDES)?;

    let mut slides = Vec::new();
    let mut current_persist_id: Option<u32> = None;
    let mut outline_texts: Vec<TextRun> = Vec::new();
    let mut current_type = TextType::Other;

    for rec in RecordIter::new(&slide_list) {
        let Ok(rec) = rec else { break };
        match rec.header.rec_type {
            RT_SLIDE_PERSIST_ATOM if rec.data.len() >= 4 => {
                if let Some(persist_id_ref) = current_persist_id.take() {
                    slides.push(resolve_slide(stream, dir, persist_id_ref, &outline_texts));
                }
                current_persist_id =
                    Some(u32::from_le_bytes([rec.data[0], rec.data[1], rec.data[2], rec.data[3]]));
                outline_texts.clear();
                current_type = TextType::Other;
            },
            RT_TEXT_HEADER if rec.data.len() >= 4 => {
                let t = u32::from_le_bytes([rec.data[0], rec.data[1], rec.data[2], rec.data[3]]);
                current_type = TextType::from_u32(t);
            },
            RT_TEXT_CHARS => {
                // Positional index into this list is meaningful (it's what
                // OutlineTextRefAtom references) — an empty run still
                // occupies a slot and must not be skipped here.
                outline_texts.push(TextRun {
                    text_type: current_type,
                    text: decode_utf16le(&rec.data),
                });
            },
            RT_TEXT_BYTES => {
                outline_texts.push(TextRun {
                    text_type: current_type,
                    text: rec.data.iter().map(|&b| b as char).collect(),
                });
            },
            _ => {},
        }
    }
    if let Some(persist_id_ref) = current_persist_id.take() {
        slides.push(resolve_slide(stream, dir, persist_id_ref, &outline_texts));
    }

    Some(slides)
}

/// Resolve one slide's shape text: locate its `Slide` container via the
/// persist directory and walk its shape tree, resolving any
/// `OutlineTextRefAtom` references against `outline_texts`.
///
/// Every persist-directory-resolved slide is kept regardless of whether text
/// was found — an image-only slide is still a slide, and the presentation's
/// true slide count matters for numbering.
fn resolve_slide(
    stream: &[u8],
    dir: &PersistDirectory,
    persist_id_ref: u32,
    outline_texts: &[TextRun],
) -> SlideText {
    let mut text_runs = Vec::new();
    if let Some(offset) = dir.resolve(persist_id_ref) {
        if let Some(children) = bounded_container_children(stream, offset, RT_SLIDE) {
            extract_shape_text(&children, 0, outline_texts, &mut text_runs);
        }
    }
    SlideText { text_runs }
}

/// Recursively collect a shape tree's text, in document order, from a bounded
/// record region (a resolved `Slide`/`Notes`/`MainMaster` container's own
/// children).
///
/// Text comes from two places: `TextHeaderAtom` + `TextCharsAtom`/
/// `TextBytesAtom` pairs embedded directly in a shape, or an
/// `OutlineTextRefAtom` resolved against `outline_texts` (see
/// [`extract_slides_via_persist`]) for shapes that store their text there
/// instead. Pass `&[]` when no outline-text table applies (e.g. the
/// no-persist-directory fallback).
///
/// Bounded per [`RecordIter`]'s container semantics: a corrupted/oversized
/// length on one shape can, at worst, truncate the remaining shapes within
/// that *same* container — it can never affect anything outside the region
/// this function was called with.
fn extract_shape_text(
    data: &[u8],
    depth: usize,
    outline_texts: &[TextRun],
    out: &mut Vec<TextRun>,
) {
    if depth > MAX_SHAPE_DEPTH {
        return;
    }
    let mut current_type = TextType::Other;

    for rec in RecordIter::new(data) {
        let Ok(rec) = rec else { break };
        match rec.header.rec_type {
            RT_TEXT_HEADER if rec.data.len() >= 4 => {
                let t = u32::from_le_bytes([rec.data[0], rec.data[1], rec.data[2], rec.data[3]]);
                current_type = TextType::from_u32(t);
            },
            RT_TEXT_CHARS => {
                let text = decode_utf16le(&rec.data);
                if !text.is_empty() {
                    out.push(TextRun {
                        text_type: current_type,
                        text,
                    });
                }
            },
            RT_TEXT_BYTES => {
                let text: String = rec.data.iter().map(|&b| b as char).collect();
                if !text.is_empty() {
                    out.push(TextRun {
                        text_type: current_type,
                        text,
                    });
                }
            },
            RT_OUTLINE_TEXT_REF_ATOM if rec.data.len() >= 4 => {
                let index =
                    i32::from_le_bytes([rec.data[0], rec.data[1], rec.data[2], rec.data[3]]);
                if index >= 0 {
                    if let Some(run) = outline_texts.get(index as usize) {
                        if !run.text.is_empty() {
                            out.push(run.clone());
                        }
                    }
                }
            },
            _ if rec.header.is_container() => {
                extract_shape_text(&rec.data, depth + 1, outline_texts, out);
            },
            _ => {},
        }
    }
}

/// Fallback used only when the persist directory can't be resolved at all:
/// derive slide boundaries from a `SlideListWithText`'s own inline text cache
/// — a `SlidePersistAtom` followed directly by `TextHeaderAtom` +
/// `TextCharsAtom`/`TextBytesAtom` pairs ([MS-PPT] 2.4.14.3). The resolved
/// `Slide` container (primary path above) additionally resolves
/// `OutlineTextRefAtom` indirection against this same sequence; this fallback
/// is only reached when persist resolution isn't available at all.
fn extract_slides_from_slide_list_cache(slide_list: &[u8]) -> Vec<SlideText> {
    let mut slides: Vec<SlideText> = Vec::new();
    let mut current: Option<SlideText> = None;
    let mut current_type = TextType::Other;

    for rec in RecordIter::new(slide_list) {
        let Ok(rec) = rec else { break };
        match rec.header.rec_type {
            RT_SLIDE_PERSIST_ATOM => {
                if let Some(slide) = current.take() {
                    if !slide.text_runs.is_empty() {
                        slides.push(slide);
                    }
                }
                current = Some(SlideText {
                    text_runs: Vec::new(),
                });
                current_type = TextType::Other;
            },
            RT_TEXT_HEADER if rec.data.len() >= 4 => {
                let t = u32::from_le_bytes([rec.data[0], rec.data[1], rec.data[2], rec.data[3]]);
                current_type = TextType::from_u32(t);
            },
            RT_TEXT_CHARS => {
                if let Some(slide) = current.as_mut() {
                    let text = decode_utf16le(&rec.data);
                    if !text.is_empty() {
                        slide.text_runs.push(TextRun {
                            text_type: current_type,
                            text,
                        });
                    }
                }
            },
            RT_TEXT_BYTES => {
                if let Some(slide) = current.as_mut() {
                    let text: String = rec.data.iter().map(|&b| b as char).collect();
                    if !text.is_empty() {
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

    if let Some(slide) = current {
        if !slide.text_runs.is_empty() {
            slides.push(slide);
        }
    }

    slides
}

/// Bounded children of the container at `offset`, if it is one of type `rec_type`.
fn bounded_container_children(stream: &[u8], offset: usize, rec_type: u16) -> Option<Vec<u8>> {
    let header = RecordHeader::parse(stream.get(offset..offset + 8)?).ok()?;
    if header.rec_type != rec_type || !header.is_container() {
        return None;
    }
    let start = offset + 8;
    let end = start
        .saturating_add(header.rec_len as usize)
        .min(stream.len());
    Some(stream.get(start..end)?.to_vec())
}

/// Single-level search for a direct child record matching `rec_type` and
/// `instance`, returning its bounded children.
fn find_child(data: &[u8], rec_type: u16, instance: u16) -> Option<Vec<u8>> {
    for rec in RecordIter::new(data) {
        let Ok(rec) = rec else { break };
        if rec.header.rec_type == rec_type && rec.header.rec_instance == instance {
            return Some(rec.data);
        }
    }
    None
}

/// Bounded recursive search for a descendant record matching `rec_type` and
/// `instance`, returning its bounded children. Only used by the legacy
/// fallback path, where the structure isn't guaranteed to place the target at
/// any particular depth.
fn find_descendant(data: &[u8], rec_type: u16, instance: u16, depth: usize) -> Option<Vec<u8>> {
    if depth > MAX_SHAPE_DEPTH {
        return None;
    }
    for rec in RecordIter::new(data) {
        let Ok(rec) = rec else { break };
        if rec.header.rec_type == rec_type && rec.header.rec_instance == instance {
            return Some(rec.data);
        }
        if rec.header.is_container() {
            if let Some(found) = find_descendant(&rec.data, rec_type, instance, depth + 1) {
                return Some(found);
            }
        }
    }
    None
}

/// Text content of a single slide.
#[derive(Debug, Clone)]
pub struct SlideText {
    /// All text runs belonging to this slide.
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
        let mut runs = Vec::new();
        extract_shape_text(&stream, 0, &[], &mut runs);
        assert_eq!(runs.len(), 1);
        assert_eq!(runs[0].text, "Hi");
        assert_eq!(runs[0].text_type, TextType::Title);
    }

    #[test]
    fn extract_text_bytes() {
        let mut stream = make_atom(RT_TEXT_HEADER, 0, &1u32.to_le_bytes()); // Body
        stream.extend(make_atom(RT_TEXT_BYTES, 0, b"Hello World"));
        let mut runs = Vec::new();
        extract_shape_text(&stream, 0, &[], &mut runs);
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
        let mut runs = Vec::new();
        extract_shape_text(&stream, 0, &[], &mut runs);
        assert_eq!(runs.len(), 2);
        assert_eq!(runs[0].text, "Title");
        assert_eq!(runs[1].text, "Body text");
    }

    #[test]
    fn extract_text_from_nested_shape_containers() {
        // Text nested a few containers deep (shape -> group -> textbox),
        // as it actually appears inside a real Slide record.
        let header = make_atom(RT_TEXT_HEADER, 0, &0u32.to_le_bytes());
        let mut textbox_children = header;
        textbox_children.extend(make_atom(RT_TEXT_BYTES, 0, b"Nested"));
        let textbox = make_container(0xF00D, 0, &textbox_children); // ClientTextbox
        let shape = make_container(0xF004, 0, &textbox); // shape container

        let mut runs = Vec::new();
        extract_shape_text(&shape, 0, &[], &mut runs);
        assert_eq!(runs.len(), 1);
        assert_eq!(runs[0].text, "Nested");
    }

    /// Fallback path only: when there's no resolvable persist directory at
    /// all, slide boundaries are derived from a `SlideListWithText`'s own
    /// inline text cache.
    #[test]
    fn slide_list_cache_fallback_without_persist_directory() {
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

        let stream = make_container(RT_SLIDE_LIST_WITH_TEXT, SLWT_SLIDES, &children);
        let slides = extract_slides_text(&stream, None);
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
        // Just raw text atoms without SlideListWithText or a persist directory.
        let mut stream = make_atom(RT_TEXT_HEADER, 0, &0u32.to_le_bytes());
        stream.extend(make_atom(RT_TEXT_BYTES, 0, b"Fallback text"));
        let slides = extract_slides_text(&stream, None);
        assert_eq!(slides.len(), 1);
        assert_eq!(slides[0].text_runs[0].text, "Fallback text");
    }

    // ── Persist-directory resolution correctness ──
    // A minimal synthetic fixture (per CONTRIBUTING.md — no third-party
    // documents committed as fixtures) covering two defect classes an
    // incrementally-resaved PPT97 stream can trigger:
    //   1. A stale/orphaned Slide-shaped block left behind by an earlier save
    //      (unreferenced by the persist directory) must be ignored in favor
    //      of the *current*, persist-directory-resolved Slide.
    //   2. A single record with a corrupted/oversized declared length must
    //      not derail extraction of any later content in the stream.

    fn user_edit_atom_bytes(
        offset_last_edit: u32,
        offset_persist_directory: u32,
        doc_persist_id_ref: u32,
    ) -> Vec<u8> {
        let mut body = Vec::new();
        body.extend_from_slice(&0u32.to_le_bytes()); // lastSlideIdRef
        body.extend_from_slice(&0u32.to_le_bytes()); // version/minor/major
        body.extend_from_slice(&offset_last_edit.to_le_bytes());
        body.extend_from_slice(&offset_persist_directory.to_le_bytes());
        body.extend_from_slice(&doc_persist_id_ref.to_le_bytes());
        body.extend_from_slice(&0u32.to_le_bytes()); // persistIdSeed
        body.extend_from_slice(&0u32.to_le_bytes()); // lastView + unused
        make_atom(RT_USER_EDIT_ATOM, 0, &body)
    }

    fn persist_directory_bytes(entries: &[(u32, u32)]) -> Vec<u8> {
        let persist_id = entries[0].0;
        let c_persist = entries.len() as u32;
        let header = persist_id | (c_persist << 20);
        let mut body = header.to_le_bytes().to_vec();
        for (_, off) in entries {
            body.extend_from_slice(&off.to_le_bytes());
        }
        make_atom(RT_PERSIST_DIRECTORY_ATOM, 0, &body)
    }

    fn current_user_bytes(offset_to_current_edit: u32) -> Vec<u8> {
        let mut body = Vec::new();
        body.extend_from_slice(&0u32.to_le_bytes()); // size
        body.extend_from_slice(&0u32.to_le_bytes()); // headerToken
        body.extend_from_slice(&offset_to_current_edit.to_le_bytes());
        body.extend_from_slice(&0u16.to_le_bytes()); // lenUserName
        body.extend_from_slice(&0u16.to_le_bytes()); // docFileVersion
        body.push(0); // majorVersion
        body.push(0); // minorVersion
        body.extend_from_slice(&0u16.to_le_bytes()); // unused
        make_atom(RT_CURRENT_USER_ATOM, 0, &body)
    }

    fn slide_persist_atom_bytes(persist_id_ref: u32, slide_id: u32) -> Vec<u8> {
        let mut body = Vec::new();
        body.extend_from_slice(&persist_id_ref.to_le_bytes());
        body.extend_from_slice(&0u32.to_le_bytes()); // flags/reserved
        body.extend_from_slice(&0u32.to_le_bytes()); // cTexts
        body.extend_from_slice(&slide_id.to_le_bytes());
        body.extend_from_slice(&0u32.to_le_bytes()); // reserved3
        make_atom(RT_SLIDE_PERSIST_ATOM, 0, &body)
    }

    fn slide_container_bytes(title: &str) -> Vec<u8> {
        let header = make_atom(RT_TEXT_HEADER, 0, &0u32.to_le_bytes());
        let mut textbox_children = header;
        textbox_children.extend(make_atom(RT_TEXT_BYTES, 0, title.as_bytes()));
        let textbox = make_container(0xF00D, 0, &textbox_children); // ClientTextbox
        make_container(RT_SLIDE, 0, &textbox)
    }

    /// Builds a synthetic "PowerPoint Document" stream containing a
    /// stale/orphaned slide-shaped block the persist directory does *not*
    /// reference, followed by the real (persist-resolved)
    /// Document/SlideListWithText/Slide structure, with an empty inline text
    /// cache in the SlideListWithText — the shape incrementally-resaved
    /// real-world files routinely take. Returns `(stream, current_user_stream)`.
    fn build_persist_regression_fixture() -> (Vec<u8>, Vec<u8>) {
        let mut stream = Vec::new();

        // Stale, orphaned copy of slide-shaped content — not in the persist
        // directory, so it must never appear in extracted output.
        stream.extend(slide_container_bytes("DECOY STALE TEXT"));

        // DocumentContainer (persist id 1), with an empty inline text cache.
        let doc_offset = stream.len() as u32;
        let slide_persist = slide_persist_atom_bytes(2, 256);
        let slide_list = make_container(RT_SLIDE_LIST_WITH_TEXT, SLWT_SLIDES, &slide_persist);
        stream.extend(make_container(RT_DOCUMENT, 0, &slide_list));

        // The real, current Slide container (persist id 2).
        let real_slide_offset = stream.len() as u32;
        stream.extend(slide_container_bytes("REAL SLIDE TEXT"));

        // Persist directory + user edit atom.
        let pd_offset = stream.len() as u32;
        stream.extend(persist_directory_bytes(&[(1, doc_offset), (2, real_slide_offset)]));
        let edit_offset = stream.len() as u32;
        stream.extend(user_edit_atom_bytes(0, pd_offset, 1));

        let current_user = current_user_bytes(edit_offset);
        (stream, current_user)
    }

    #[test]
    fn persist_resolution_ignores_stale_orphaned_slide_copy() {
        let (stream, current_user) = build_persist_regression_fixture();
        let slides = extract_slides_text(&stream, Some(&current_user));

        assert_eq!(slides.len(), 1);
        assert_eq!(slides[0].text_runs[0].text, "REAL SLIDE TEXT");
        assert!(
            !slides
                .iter()
                .flat_map(|s| &s.text_runs)
                .any(|r| r.text.contains("DECOY")),
            "stale orphaned copy must not appear in extracted text"
        );
    }

    #[test]
    fn persist_resolution_works_without_current_user_stream() {
        let (stream, _current_user) = build_persist_regression_fixture();
        let slides = extract_slides_text(&stream, None);

        assert_eq!(slides.len(), 1);
        assert_eq!(slides[0].text_runs[0].text, "REAL SLIDE TEXT");
    }

    fn corrupt_record() -> Vec<u8> {
        // A top-level record declaring a wildly oversized length with no
        // data behind it, as produced by non-conformant real-world PPT97
        // writers.
        let mut corrupt = Vec::new();
        corrupt.extend_from_slice(&0u16.to_le_bytes()); // ver=0 (atom)
        corrupt.extend_from_slice(&RT_TEXT_CHARS.to_le_bytes());
        corrupt.extend_from_slice(&5_000_000u32.to_le_bytes()); // bogus length
        corrupt
    }

    #[test]
    fn corrupted_record_length_before_real_content_does_not_lose_it() {
        // Same shape as build_persist_regression_fixture, with a corrupt
        // top-level record spliced in right after the stale block.
        let mut stream = Vec::new();
        stream.extend(slide_container_bytes("DECOY STALE TEXT"));
        stream.extend(corrupt_record());

        let doc_offset = stream.len() as u32;
        let slide_persist = slide_persist_atom_bytes(2, 256);
        let slide_list = make_container(RT_SLIDE_LIST_WITH_TEXT, SLWT_SLIDES, &slide_persist);
        stream.extend(make_container(RT_DOCUMENT, 0, &slide_list));

        let real_slide_offset = stream.len() as u32;
        stream.extend(slide_container_bytes("REAL SLIDE TEXT"));

        let pd_offset = stream.len() as u32;
        stream.extend(persist_directory_bytes(&[(1, doc_offset), (2, real_slide_offset)]));
        let edit_offset = stream.len() as u32;
        stream.extend(user_edit_atom_bytes(0, pd_offset, 1));
        let current_user = current_user_bytes(edit_offset);

        let slides = extract_slides_text(&stream, Some(&current_user));
        assert_eq!(slides.len(), 1);
        assert_eq!(
            slides[0].text_runs[0].text, "REAL SLIDE TEXT",
            "a corrupted record's length must not prevent later real content from being found"
        );
    }

    #[test]
    fn outline_text_ref_atom_resolves_indexed_placeholder_text() {
        // A shape whose ClientTextbox holds only an OutlineTextRefAtom
        // (index 1) instead of embedding its own TextHeaderAtom/TextChars —
        // the placeholder-text-by-reference pattern real PPT97 title/body
        // placeholders commonly use ([MS-PPT] 2.4.15.6).
        let outline_texts = vec![
            TextRun {
                text_type: TextType::Title,
                text: "first".into(),
            },
            TextRun {
                text_type: TextType::Body,
                text: "second".into(),
            },
        ];

        let mut index_bytes = Vec::new();
        index_bytes.extend_from_slice(&1i32.to_le_bytes());
        let outline_ref = make_atom(RT_OUTLINE_TEXT_REF_ATOM, 0, &index_bytes);
        let textbox = make_container(0xF00D, 0, &outline_ref);
        let shape = make_container(0xF004, 0, &textbox);

        let mut runs = Vec::new();
        extract_shape_text(&shape, 0, &outline_texts, &mut runs);
        assert_eq!(runs.len(), 1);
        assert_eq!(runs[0].text, "second");
        assert_eq!(runs[0].text_type, TextType::Body);
    }

    #[test]
    fn persist_resolution_follows_outline_text_ref_atom() {
        // A Slide container whose only shape holds an OutlineTextRefAtom, with
        // the actual text living in the SlideListWithText's per-slide outline
        // cache rather than embedded in the shape itself.
        let mut stream = Vec::new();

        let doc_offset = stream.len() as u32;
        let mut slide_list_children = slide_persist_atom_bytes(2, 256);
        slide_list_children.extend(make_atom(RT_TEXT_HEADER, 0, &0u32.to_le_bytes()));
        slide_list_children.extend(make_atom(RT_TEXT_BYTES, 0, b"OUTLINE-REFERENCED TEXT"));
        let slide_list = make_container(RT_SLIDE_LIST_WITH_TEXT, SLWT_SLIDES, &slide_list_children);
        stream.extend(make_container(RT_DOCUMENT, 0, &slide_list));

        let real_slide_offset = stream.len() as u32;
        let mut index_bytes = Vec::new();
        index_bytes.extend_from_slice(&0i32.to_le_bytes());
        let outline_ref = make_atom(RT_OUTLINE_TEXT_REF_ATOM, 0, &index_bytes);
        let textbox = make_container(0xF00D, 0, &outline_ref);
        let shape = make_container(0xF004, 0, &textbox);
        stream.extend(make_container(RT_SLIDE, 0, &shape));

        let pd_offset = stream.len() as u32;
        stream.extend(persist_directory_bytes(&[(1, doc_offset), (2, real_slide_offset)]));
        let edit_offset = stream.len() as u32;
        stream.extend(user_edit_atom_bytes(0, pd_offset, 1));
        let current_user = current_user_bytes(edit_offset);

        let slides = extract_slides_text(&stream, Some(&current_user));
        assert_eq!(slides.len(), 1);
        assert_eq!(slides[0].text_runs[0].text, "OUTLINE-REFERENCED TEXT");
    }

    #[test]
    fn falls_back_to_inline_cache_when_persist_resolved_slides_are_all_textless() {
        // Persist resolution structurally succeeds (valid directory, valid
        // Slide offset) but the resolved Slide container has no extractable
        // text at all — as happens when directory corruption points at the
        // wrong offset. The SlideListWithText's own inline cache does have
        // real text, so it must be used instead of returning nothing.
        let mut stream = Vec::new();

        let doc_offset = stream.len() as u32;
        let mut slide_list_children = slide_persist_atom_bytes(2, 256);
        slide_list_children.extend(make_atom(RT_TEXT_HEADER, 0, &0u32.to_le_bytes()));
        slide_list_children.extend(make_atom(RT_TEXT_BYTES, 0, b"FALLBACK CACHE TEXT"));
        let slide_list = make_container(RT_SLIDE_LIST_WITH_TEXT, SLWT_SLIDES, &slide_list_children);
        stream.extend(make_container(RT_DOCUMENT, 0, &slide_list));

        // A resolvable but genuinely empty Slide container (no shapes at all).
        let real_slide_offset = stream.len() as u32;
        stream.extend(make_container(RT_SLIDE, 0, &[]));

        let pd_offset = stream.len() as u32;
        stream.extend(persist_directory_bytes(&[(1, doc_offset), (2, real_slide_offset)]));
        let edit_offset = stream.len() as u32;
        stream.extend(user_edit_atom_bytes(0, pd_offset, 1));
        let current_user = current_user_bytes(edit_offset);

        let slides = extract_slides_text(&stream, Some(&current_user));
        assert_eq!(slides.len(), 1);
        assert_eq!(slides[0].text_runs[0].text, "FALLBACK CACHE TEXT");
    }
}
