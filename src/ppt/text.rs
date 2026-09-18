//! Text extraction from PPT binary records.

use std::collections::HashMap;

use super::persist::{self, PersistDirectory};
use super::records::*;
use super::style::{self, CharFormatSpan, ParaFormatSpan};

/// Guards against pathologically deep (or maliciously crafted) shape nesting.
const MAX_SHAPE_DEPTH: usize = 64;

/// Text type from TextHeaderAtom.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum TextType {
    /// Title placeholder text.
    Title,
    /// Body / content placeholder text.
    Body,
    /// Speaker notes text.
    Notes,
    /// Other or unclassified text.
    #[default]
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
    ///
    /// [MS-PPT] `TextTypeEnum`: `Tx_TYPE_TITLE`=0, `_BODY`=1, `_NOTES`=2 (3
    /// is undefined — the spec's enumeration jumps straight from 2 to 4),
    /// `_OTHER`=4, `_CENTERBODY`=5, `_CENTERTITLE`=6, `_HALFBODY`=7,
    /// `_QUARTERBODY`=8. Every value from 3 upward used to be shifted one
    /// slot low (5 read as CenterTitle, 6 as HalfBody, …), which swapped a
    /// title-slide layout's real title (6, CenterTitle) and subtitle (5,
    /// CenterBody) — the single most common slide layout in real
    /// presentations (issue #253). Verified against the published
    /// [MS-PPT] TextTypeEnum spec page and its own worked byte example.
    pub fn from_u32(val: u32) -> Self {
        match val {
            0 => Self::Title,
            1 => Self::Body,
            2 => Self::Notes,
            5 => Self::CenterBody,
            6 => Self::CenterTitle,
            7 => Self::HalfBody,
            8 => Self::QuarterBody,
            _ => Self::Other, // 3 (undefined), 4 (Tx_TYPE_OTHER), and anything else
        }
    }
}

/// A text run extracted from a slide.
#[derive(Debug, Clone, Default)]
pub struct TextRun {
    /// The role of this text within its slide.
    pub text_type: TextType,
    /// The decoded text content.
    pub text: String,
    /// The URL target of the shape's own `InteractiveInfo` click action
    /// (`II_HyperlinkAction`/`II_JumpAction`/`II_CustomShowAction`),
    /// resolved through the document's `ExObjList`. `None` when the shape
    /// has no interactive info, or its action has no hyperlink to resolve
    /// (issue #257).
    pub hyperlink: Option<String>,
    /// Direct character-level formatting (bold/italic/underline/size/
    /// color/position), resolved from this run's `StyleTextPropAtom` if
    /// one was present. Empty when no such atom was found — callers must
    /// treat that as "no formatting info", not "definitely unformatted"
    /// (issue #254).
    pub char_formats: Vec<CharFormatSpan>,
    /// Direct paragraph-level formatting (currently: alignment only),
    /// resolved the same way. Empty when no `StyleTextPropAtom` was found.
    pub para_formats: Vec<ParaFormatSpan>,
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
    // A best-effort fallback scope for the weaker paths below, which have
    // no persist-resolved DocumentContainer to search within at all.
    let stream_wide_hyperlinks = parse_ex_hyperlinks(stream);
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

    // Weaker fallback: walk only `Slide` containers found anywhere in the
    // stream. A whole-stream scan also picks up `MainMaster` containers,
    // whose placeholder prompts ("Click to edit Master title style", the
    // `*` bullet placeholders) are PowerPoint's own UI strings and never
    // render on a slide — in two POI corpus files they were 116 of the 137
    // and 121 extracted characters respectively.
    let mut slides = Vec::new();
    collect_slide_containers(stream, 0, &stream_wide_hyperlinks, &mut slides);
    if slides.iter().any(|s| !s.text_runs.is_empty()) {
        return slides;
    }

    // Last resort: no resolvable structure at all — dump whatever text atoms
    // exist anywhere in the stream, minus the master boilerplate.
    let mut runs = Vec::new();
    let mut tables = Vec::new();
    let mut image_refs = Vec::new();
    extract_shape_text(
        stream,
        0,
        &[],
        &stream_wide_hyperlinks,
        None,
        &mut runs,
        &mut tables,
        &mut image_refs,
    );
    runs.retain(|r| !is_master_placeholder_prompt(&r.text));
    if runs.is_empty() && tables.is_empty() && image_refs.is_empty() {
        Vec::new()
    } else {
        vec![SlideText { text_runs: runs, tables, image_refs }]
    }
}

/// Walk the record tree collecting one `SlideText` per `Slide` container.
fn collect_slide_containers(
    data: &[u8],
    depth: usize,
    hyperlinks: &HashMap<u32, String>,
    out: &mut Vec<SlideText>,
) {
    if depth > MAX_SHAPE_DEPTH {
        return;
    }
    for rec in RecordIter::new(data) {
        let Ok(rec) = rec else { break };
        if rec.header.rec_type == RT_SLIDE {
            let mut runs = Vec::new();
            let mut tables = Vec::new();
            let mut image_refs = Vec::new();
            extract_shape_text(
                &rec.data,
                0,
                &[],
                hyperlinks,
                None,
                &mut runs,
                &mut tables,
                &mut image_refs,
            );
            runs.retain(|r| !is_master_placeholder_prompt(&r.text));
            out.push(SlideText { text_runs: runs, tables, image_refs });
            continue;
        }
        if rec.header.is_container() {
            collect_slide_containers(&rec.data, depth + 1, hyperlinks, out);
        }
    }
}

/// Whether a text run is a slide-master placeholder prompt rather than
/// document content.
///
/// PowerPoint stores the master's prompt strings as ordinary text atoms.
/// They are shown in master view and never rendered on a slide, so a
/// consumer that receives them gets a document whose "content" is the
/// application's own UI strings.
fn is_master_placeholder_prompt(text: &str) -> bool {
    let t = text.trim();
    if t.is_empty() {
        return false;
    }
    // The English prompts PowerPoint 97–2003 writes, plus the bare bullet
    // placeholders that accompany them.
    t.starts_with("Click to edit Master")
        || t.starts_with("Click to edit the outline text format")
        || t.starts_with("Click to add title")
        || t.starts_with("Click to add text")
        || t.starts_with("Click to add notes")
        || t.chars().all(|c| c == '*' || c.is_whitespace())
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
    // The current DocumentContainer's own ExObjListContainer — not a raw
    // whole-stream scan, which could resolve a stale, superseded copy left
    // behind by an earlier incremental save (the same hazard the persist
    // directory itself exists to route around for slides) (issue #257).
    let hyperlinks = parse_ex_hyperlinks(&doc_children);

    let mut slides = Vec::new();
    let mut current_persist_id: Option<u32> = None;
    let mut outline_texts: Vec<TextRun> = Vec::new();
    let mut current_type = TextType::Other;
    let mut last_outline_idx: Option<usize> = None;

    for rec in RecordIter::new(&slide_list) {
        let Ok(rec) = rec else { break };
        match rec.header.rec_type {
            RT_SLIDE_PERSIST_ATOM if rec.data.len() >= 4 => {
                if let Some(persist_id_ref) = current_persist_id.take() {
                    slides.push(resolve_slide(stream, dir, persist_id_ref, &outline_texts, &hyperlinks));
                }
                current_persist_id =
                    Some(u32::from_le_bytes([rec.data[0], rec.data[1], rec.data[2], rec.data[3]]));
                outline_texts.clear();
                current_type = TextType::Other;
                last_outline_idx = None;
            },
            RT_TEXT_HEADER if rec.data.len() >= 4 => {
                let t = u32::from_le_bytes([rec.data[0], rec.data[1], rec.data[2], rec.data[3]]);
                current_type = TextType::from_u32(t);
                last_outline_idx = None;
            },
            RT_TEXT_CHARS => {
                // Positional index into this list is meaningful (it's what
                // OutlineTextRefAtom references) — an empty run still
                // occupies a slot and must not be skipped here.
                outline_texts.push(TextRun {
                    text_type: current_type,
                    text: decode_utf16le(&rec.data),
                    hyperlink: None,
                    ..Default::default()
                });
                last_outline_idx = Some(outline_texts.len() - 1);
            },
            RT_TEXT_BYTES => {
                outline_texts.push(TextRun {
                    text_type: current_type,
                    text: rec.data.iter().map(|&b| b as char).collect(),
                    hyperlink: None,
                    ..Default::default()
                });
                last_outline_idx = Some(outline_texts.len() - 1);
            },
            RT_STYLE_TEXT_PROP => {
                if let Some(idx) = last_outline_idx {
                    apply_style_text_prop(&mut outline_texts[idx], &rec.data);
                }
            },
            _ => {},
        }
    }
    if let Some(persist_id_ref) = current_persist_id.take() {
        slides.push(resolve_slide(stream, dir, persist_id_ref, &outline_texts, &hyperlinks));
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
    hyperlinks: &HashMap<u32, String>,
) -> SlideText {
    let mut text_runs = Vec::new();
    let mut tables = Vec::new();
    let mut image_refs = Vec::new();
    if let Some(offset) = dir.resolve(persist_id_ref) {
        if let Some(children) = bounded_container_children(stream, offset, RT_SLIDE) {
            extract_shape_text(
                &children,
                0,
                outline_texts,
                hyperlinks,
                None,
                &mut text_runs,
                &mut tables,
                &mut image_refs,
            );
        }
    }
    SlideText { text_runs, tables, image_refs }
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
    hyperlinks: &HashMap<u32, String>,
    current_hyperlink: Option<&str>,
    out: &mut Vec<TextRun>,
    tables: &mut Vec<super::table::TableBlock>,
    image_refs: &mut Vec<usize>,
) {
    if depth > MAX_SHAPE_DEPTH {
        return;
    }
    let mut current_type = TextType::Other;
    // Text-run-level hyperlink state (issue #257): a `MouseClick/
    // MouseOverInteractiveInfoContainer` appearing directly as a *sibling*
    // of the text atoms (not nested in `RT_CLIENT_DATA`, which is the
    // separate whole-shape mechanism `RT_SHAPE` below already handles) is
    // immediately followed by a `MouseClick/MouseOverTextInteractiveInfoAtom`
    // giving the character range, within the *most recently pushed* text
    // run, that the hyperlink actually covers.
    let mut last_text_run_idx: Option<usize> = None;
    let mut pending_interactive: Option<(u32, u8)> = None;

    for rec in RecordIter::new(data) {
        let Ok(rec) = rec else { break };
        match rec.header.rec_type {
            RT_TEXT_HEADER if rec.data.len() >= 4 => {
                let t = u32::from_le_bytes([rec.data[0], rec.data[1], rec.data[2], rec.data[3]]);
                current_type = TextType::from_u32(t);
                last_text_run_idx = None;
                pending_interactive = None;
            },
            RT_TEXT_CHARS => {
                let text = decode_utf16le(&rec.data);
                if !text.is_empty() {
                    out.push(TextRun {
                        text_type: current_type,
                        text,
                        hyperlink: current_hyperlink.map(str::to_string),
                        ..Default::default()
                    });
                    last_text_run_idx = Some(out.len() - 1);
                }
            },
            RT_TEXT_BYTES => {
                let text: String = rec.data.iter().map(|&b| b as char).collect();
                if !text.is_empty() {
                    out.push(TextRun {
                        text_type: current_type,
                        text,
                        hyperlink: current_hyperlink.map(str::to_string),
                        ..Default::default()
                    });
                    last_text_run_idx = Some(out.len() - 1);
                }
            },
            RT_STYLE_TEXT_PROP => {
                if let Some(idx) = last_text_run_idx {
                    apply_style_text_prop(&mut out[idx], &rec.data);
                }
            },
            RT_OUTLINE_TEXT_REF_ATOM if rec.data.len() >= 4 => {
                let index =
                    i32::from_le_bytes([rec.data[0], rec.data[1], rec.data[2], rec.data[3]]);
                if index >= 0 {
                    if let Some(run) = outline_texts.get(index as usize) {
                        if !run.text.is_empty() {
                            let mut run = run.clone();
                            run.hyperlink = current_hyperlink.map(str::to_string);
                            out.push(run);
                            last_text_run_idx = Some(out.len() - 1);
                        }
                    }
                }
            },
            RT_INTERACTIVE_INFO => {
                // A sibling-level InteractiveInfo (text-run hyperlink);
                // when this same container instead sits inside
                // `RT_CLIENT_DATA` (the whole-shape case), it's reached and
                // handled separately by `resolve_shape_hyperlink` below —
                // capturing it here too is harmless since no
                // `RT_TEXT_INTERACTIVE_INFO_ATOM` ever immediately follows
                // it in that context, so `pending_interactive` just gets
                // reset at the next `RT_TEXT_HEADER` unused.
                if let Some(atom) = find_descendant(&rec.data, RT_INTERACTIVE_INFO_ATOM, 0, 0) {
                    if atom.len() >= 9 {
                        let ex_hyperlink_id_ref =
                            u32::from_le_bytes([atom[4], atom[5], atom[6], atom[7]]);
                        pending_interactive = Some((ex_hyperlink_id_ref, atom[8]));
                    }
                }
            },
            RT_TEXT_INTERACTIVE_INFO_ATOM if rec.data.len() >= 8 => {
                if let (Some(idx), Some((ex_hyperlink_id_ref, action))) =
                    (last_text_run_idx, pending_interactive.take())
                {
                    if matches!(action, 0x03 | 0x04 | 0x07) {
                        if let Some(url) = hyperlinks.get(&ex_hyperlink_id_ref) {
                            let begin = i32::from_le_bytes([
                                rec.data[0],
                                rec.data[1],
                                rec.data[2],
                                rec.data[3],
                            ])
                            .max(0) as usize;
                            let end = i32::from_le_bytes([
                                rec.data[4],
                                rec.data[5],
                                rec.data[6],
                                rec.data[7],
                            ])
                            .max(0) as usize;
                            split_run_with_hyperlink(out, idx, begin, end, url);
                        }
                    }
                }
            },
            RT_SHAPE => {
                // Resolve this shape's own hyperlink (if any) before
                // walking its subtree, so every TextRun produced from it
                // — including nested containers like a group's own child
                // shapes, which are siblings under a group, not children
                // of THIS shape's own text — carries it (issue #257).
                let shape_hyperlink = resolve_shape_hyperlink(&rec.data, hyperlinks);
                let hyperlink_ref = shape_hyperlink.as_deref().or(current_hyperlink);
                // This shape's own picture reference, if it has one
                // (issue #256) — resolved here, at the shape actually
                // carrying it, not inferred from whichever slide
                // happens to be processed last.
                if let Some(idx) = resolve_shape_pib(&rec.data) {
                    image_refs.push(idx);
                }
                extract_shape_text(
                    &rec.data,
                    depth + 1,
                    outline_texts,
                    hyperlinks,
                    hyperlink_ref,
                    out,
                    tables,
                    image_refs,
                );
            },
            RT_SPGR_CONTAINER => {
                // A group whose members form a clean rectangular grid is
                // a reconstructed table (issue #255); anything less
                // certain falls through to the ordinary flat-paragraph
                // group walk below, unchanged from before this existed.
                if let Some(table) = try_extract_table_from_spgr(
                    &rec.data,
                    depth + 1,
                    outline_texts,
                    hyperlinks,
                    current_hyperlink,
                ) {
                    tables.push(table);
                } else {
                    extract_shape_text(
                        &rec.data,
                        depth + 1,
                        outline_texts,
                        hyperlinks,
                        current_hyperlink,
                        out,
                        tables,
                        image_refs,
                    );
                }
            },
            _ if rec.header.is_container() => {
                extract_shape_text(
                    &rec.data,
                    depth + 1,
                    outline_texts,
                    hyperlinks,
                    current_hyperlink,
                    out,
                    tables,
                    image_refs,
                );
            },
            _ => {},
        }
    }
}

/// Try to recognize `spgr_data` (an `OfficeArtSpgrContainer`'s own
/// children) as a table: every `RT_SHAPE` after the group's own leading
/// placeholder shape must carry an `RT_CHILD_ANCHOR`, and the resulting
/// positions must form a clean rectangular grid (issue #255). Returns
/// `None` on the first sign this isn't a simple table (a member with no
/// anchor, or a grid [`table::build_table`] can't make sense of) — the
/// caller falls back to the ordinary flat-paragraph group walk.
fn try_extract_table_from_spgr(
    spgr_data: &[u8],
    depth: usize,
    outline_texts: &[TextRun],
    hyperlinks: &HashMap<u32, String>,
    current_hyperlink: Option<&str>,
) -> Option<super::table::TableBlock> {
    if depth > MAX_SHAPE_DEPTH {
        return None;
    }
    let mut cells = Vec::new();
    let mut seen_group_placeholder = false;
    for rec in RecordIter::new(spgr_data) {
        let Ok(rec) = rec else { break };
        if rec.header.rec_type != RT_SHAPE {
            continue;
        }
        if !seen_group_placeholder {
            // The group's own placeholder shape — its bounding box, not
            // a cell. It has no `RT_CHILD_ANCHOR` of its own.
            seen_group_placeholder = true;
            continue;
        }
        let anchor = RecordIter::new(&rec.data)
            .filter_map(Result::ok)
            .find(|c| c.header.rec_type == RT_CHILD_ANCHOR)?;
        if anchor.data.len() < 16 {
            return None;
        }
        let left =
            i32::from_le_bytes([anchor.data[0], anchor.data[1], anchor.data[2], anchor.data[3]]);
        let top =
            i32::from_le_bytes([anchor.data[4], anchor.data[5], anchor.data[6], anchor.data[7]]);

        let shape_hyperlink = resolve_shape_hyperlink(&rec.data, hyperlinks);
        let hyperlink_ref = shape_hyperlink.as_deref().or(current_hyperlink);
        let mut runs = Vec::new();
        extract_shape_text(
            &rec.data,
            depth + 1,
            outline_texts,
            hyperlinks,
            hyperlink_ref,
            &mut runs,
            &mut Vec::new(),
            &mut Vec::new(),
        );
        cells.push(super::table::TableCellData { left, top, runs });
    }
    super::table::build_table(&cells)
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
                current = Some(SlideText { text_runs: Vec::new(), ..Default::default() });
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
                            // This inline-cache walk has no shape tree to
                            // resolve a hyperlink from — it's a weaker
                            // fallback than the persist-directory path,
                            // only reached when that path isn't available.
                            hyperlink: None,
                            ..Default::default()
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
                            hyperlink: None,
                            ..Default::default()
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

/// Build the document-wide `exHyperlinkId -> target URL` table from the
/// `ExObjListContainer` ([MS-PPT] 2.10.1), a direct child of the top-level
/// `DocumentContainer` — i.e. of `stream` itself, the same "PowerPoint
/// Document" stream passed to [`extract_slides_text`]. Each entry comes
/// from one `ExHyperlinkContainer`'s `ExHyperlinkAtom.exHyperlinkId` and
/// its sibling `TargetAtom` (a `RT_CSTRING` at
/// [`CSTRING_INSTANCE_TARGET`]) (issue #257).
fn parse_ex_hyperlinks(stream: &[u8]) -> HashMap<u32, String> {
    let mut out = HashMap::new();
    let Some(ex_obj_list) = find_descendant(stream, RT_EXTERNAL_OBJECT_LIST, 0, 0) else {
        return out;
    };
    collect_ex_hyperlinks(&ex_obj_list, 0, &mut out);
    out
}

fn collect_ex_hyperlinks(data: &[u8], depth: usize, out: &mut HashMap<u32, String>) {
    if depth > MAX_SHAPE_DEPTH {
        return;
    }
    for rec in RecordIter::new(data) {
        let Ok(rec) = rec else { break };
        if rec.header.rec_type == RT_EXTERNAL_HYPERLINK {
            if let Some((id, url)) = parse_one_ex_hyperlink(&rec.data) {
                out.insert(id, url);
            }
            continue; // ExHyperlinkContainer's own children are never other ExHyperlinks
        }
        if rec.header.is_container() {
            collect_ex_hyperlinks(&rec.data, depth + 1, out);
        }
    }
}

/// Parse one `ExHyperlinkContainer`'s own children: its `ExHyperlinkAtom`
/// (`exHyperlinkId`) and its `TargetAtom` (the URL/path).
fn parse_one_ex_hyperlink(data: &[u8]) -> Option<(u32, String)> {
    let mut id = None;
    let mut target = None;
    for rec in RecordIter::new(data) {
        let Ok(rec) = rec else { break };
        match rec.header.rec_type {
            RT_EXTERNAL_HYPERLINK_ATOM if rec.data.len() >= 4 => {
                id = Some(u32::from_le_bytes([
                    rec.data[0],
                    rec.data[1],
                    rec.data[2],
                    rec.data[3],
                ]));
            },
            RT_CSTRING if rec.header.rec_instance == CSTRING_INSTANCE_TARGET => {
                let s = decode_utf16le(&rec.data);
                if !s.is_empty() {
                    target = Some(s);
                }
            },
            _ => {},
        }
    }
    Some((id?, target?))
}

/// Resolve a shape's own hyperlink target (if any), by scanning its direct
/// children for `OfficeArtClientData` (`RT_CLIENT_DATA`), then its
/// `InteractiveInfo` child, then that record's own `InteractiveInfoAtom`.
///
/// Only `II_JumpAction` (0x03), `II_HyperlinkAction` (0x04), and
/// `II_CustomShowAction` (0x07) carry a meaningful `exHyperlinkIdRef` per
/// [MS-PPT] 2.6.10; any other action (or one whose id doesn't resolve in
/// `hyperlinks`) yields `None` rather than a wrong-but-confident guess.
fn resolve_shape_hyperlink(shape_data: &[u8], hyperlinks: &HashMap<u32, String>) -> Option<String> {
    let client_data = find_descendant(shape_data, RT_CLIENT_DATA, 0, 0)?;
    // `rh.recInstance` distinguishes MouseClickInteractiveInfoContainer (0)
    // from MouseOverInteractiveInfoContainer (1); prefer the click action
    // (the real, navigable hyperlink) when a shape happens to have both.
    let interactive_info = find_descendant(&client_data, RT_INTERACTIVE_INFO, 0, 0)
        .or_else(|| find_descendant(&client_data, RT_INTERACTIVE_INFO, 1, 0))?;
    let atom = find_descendant(&interactive_info, RT_INTERACTIVE_INFO_ATOM, 0, 0)?;
    if atom.len() < 9 {
        return None;
    }
    let ex_hyperlink_id_ref = u32::from_le_bytes([atom[4], atom[5], atom[6], atom[7]]);
    let action = atom[8];
    if !matches!(action, 0x03 | 0x04 | 0x07) {
        return None;
    }
    hyperlinks.get(&ex_hyperlink_id_ref).cloned()
}

/// Resolve a shape's `pib` ("Blip to display") property from its
/// `OfficeArtFOPT` property table (a direct child of the shape, not
/// nested inside `RT_CLIENT_DATA`), if it has one.
///
/// Returns the *0-based* image index (into the document's own
/// `Pictures`-stream-derived image list, [`BlipImage::index`]) — `pib`
/// itself is a documented ONE-based index into that same array, and
/// `0x00000000` means "no picture" per [MS-ODRAW] (issue #256).
fn resolve_shape_pib(shape_data: &[u8]) -> Option<usize> {
    let fopt = RecordIter::new(shape_data)
        .filter_map(Result::ok)
        .find(|c| c.header.rec_type == RT_FOPT)?;
    let count = fopt.header.rec_instance as usize;
    for i in 0..count {
        let pos = i * 6;
        if pos + 6 > fopt.data.len() {
            break;
        }
        let opid = u16::from_le_bytes([fopt.data[pos], fopt.data[pos + 1]]);
        let pid = opid & 0x3FFF;
        let f_complex = (opid >> 15) & 1;
        // Blip:pib, [MS-ODRAW] property ID 0x0104 — only meaningful (a
        // plain 4-byte index, not a length) when fComplex is unset.
        if pid == 0x0104 && f_complex == 0 {
            let op = u32::from_le_bytes([
                fopt.data[pos + 2],
                fopt.data[pos + 3],
                fopt.data[pos + 4],
                fopt.data[pos + 5],
            ]);
            if op == 0 {
                return None; // explicitly "no picture"
            }
            return Some((op - 1) as usize);
        }
    }
    None
}

/// Split `out[idx]` into up to 3 runs at the character offsets `begin..end`
/// (`TextRange`, [MS-PPT] 2.6.12): the unlinked prefix (if any), the
/// hyperlinked `[begin, end)` slice, and the unlinked suffix (if any) — the
/// text-run-level hyperlink mechanism, where a hyperlink covers only part
/// of a run's text (e.g. a URL appearing mid-sentence) rather than the
/// whole shape (issue #257).
///
/// `begin`/`end` are [MS-PPT]'s `TextPosition` character offsets, which
/// this slices via `char` count rather than UTF-16 code units — an exact
/// match for the common case, off by one per astral-plane character
/// (surrogate pair) in the rare case one appears before the hyperlinked
/// range, which is judged not worth the extra bookkeeping here.
fn split_run_with_hyperlink(out: &mut Vec<TextRun>, idx: usize, begin: usize, end: usize, url: &str) {
    let Some(run) = out.get(idx) else { return };
    let chars: Vec<char> = run.text.chars().collect();
    let begin = begin.min(chars.len());
    let end = end.clamp(begin, chars.len());
    if begin >= end {
        return; // empty or invalid range — leave the run untouched
    }

    let text_type = run.text_type;
    let surrounding_hyperlink = run.hyperlink.clone();
    let prefix: String = chars[..begin].iter().collect();
    let linked: String = chars[begin..end].iter().collect();
    let suffix: String = chars[end..].iter().collect();
    // Splitting a run must not silently drop its own direct formatting —
    // slice each formatting span onto whichever piece(s) it overlaps,
    // re-based to that piece's own character indices.
    let prefix_char_fmt = slice_char_formats(&run.char_formats, 0..begin);
    let linked_char_fmt = slice_char_formats(&run.char_formats, begin..end);
    let suffix_char_fmt = slice_char_formats(&run.char_formats, end..chars.len());
    let prefix_para_fmt = slice_para_formats(&run.para_formats, 0..begin);
    let linked_para_fmt = slice_para_formats(&run.para_formats, begin..end);
    let suffix_para_fmt = slice_para_formats(&run.para_formats, end..chars.len());

    let mut replacement = Vec::with_capacity(3);
    if !prefix.is_empty() {
        replacement.push(TextRun {
            text_type,
            text: prefix,
            hyperlink: surrounding_hyperlink.clone(),
            char_formats: prefix_char_fmt,
            para_formats: prefix_para_fmt,
        });
    }
    replacement.push(TextRun {
        text_type,
        text: linked,
        hyperlink: Some(url.to_string()),
        char_formats: linked_char_fmt,
        para_formats: linked_para_fmt,
    });
    if !suffix.is_empty() {
        replacement.push(TextRun {
            text_type,
            text: suffix,
            hyperlink: surrounding_hyperlink,
            char_formats: suffix_char_fmt,
            para_formats: suffix_para_fmt,
        });
    }

    out.splice(idx..=idx, replacement);
}

/// Slice/clip a set of character-formatting spans onto `range`, re-based
/// so the returned spans are relative to `range.start` (i.e. valid over
/// the substring `text[range]` on its own).
fn slice_char_formats(spans: &[CharFormatSpan], range: std::ops::Range<usize>) -> Vec<CharFormatSpan> {
    spans
        .iter()
        .filter_map(|s| {
            let start = s.start.max(range.start);
            let end = s.end.min(range.end);
            (start < end).then(|| CharFormatSpan {
                start: start - range.start,
                end: end - range.start,
                format: s.format.clone(),
            })
        })
        .collect()
}

/// Same as [`slice_char_formats`] for paragraph-formatting spans.
fn slice_para_formats(spans: &[ParaFormatSpan], range: std::ops::Range<usize>) -> Vec<ParaFormatSpan> {
    spans
        .iter()
        .filter_map(|s| {
            let start = s.start.max(range.start);
            let end = s.end.min(range.end);
            (start < end).then(|| ParaFormatSpan {
                start: start - range.start,
                end: end - range.start,
                format: s.format.clone(),
            })
        })
        .collect()
}

/// Parse `data` as a `StyleTextPropAtom` body and attach the resulting
/// character-/paragraph-formatting spans to `run`, clamped against
/// `run.text`'s own character count (issue #254).
fn apply_style_text_prop(run: &mut TextRun, data: &[u8]) {
    let text_char_len = run.text.chars().count();
    let (para_spans, char_spans) = style::parse_style_text_prop(data, text_char_len);
    run.para_formats = para_spans;
    run.char_formats = char_spans;
}

/// Text content of a single slide.
#[derive(Debug, Clone, Default)]
pub struct SlideText {
    /// All text runs belonging to this slide.
    pub text_runs: Vec<TextRun>,
    /// Shape groups recognized as tables (issue #255). Rendered after
    /// `text_runs` in the IR — the binary format has no single unified
    /// reading-order concept to interleave them with, so this is a
    /// deliberate simplification, not a claim of true document order.
    pub tables: Vec<super::table::TableBlock>,
    /// 0-based indices into the document's `Pictures`-stream-derived
    /// image list ([`super::images::PptImage::index`]) for every
    /// picture shape resolved on this slide, in shape-tree encounter
    /// order (issue #256 — these used to be silently dumped onto
    /// whichever slide happened to be last, regardless of which slide
    /// actually contains the shape referencing them).
    pub image_refs: Vec<usize>,
}

fn decode_utf16le(data: &[u8]) -> String {
    let (pairs, _rest) = data.as_chunks::<2>();
    let chars: Vec<u16> = pairs.iter().copied().map(u16::from_le_bytes).collect();
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
        extract_shape_text(
            &stream,
            0,
            &[],
            &HashMap::new(),
            None,
            &mut runs,
            &mut Vec::new(),
            &mut Vec::new(),
        );
        assert_eq!(runs.len(), 1);
        assert_eq!(runs[0].text, "Hi");
        assert_eq!(runs[0].text_type, TextType::Title);
    }

    #[test]
    fn extract_text_bytes() {
        let mut stream = make_atom(RT_TEXT_HEADER, 0, &1u32.to_le_bytes()); // Body
        stream.extend(make_atom(RT_TEXT_BYTES, 0, b"Hello World"));
        let mut runs = Vec::new();
        extract_shape_text(
            &stream,
            0,
            &[],
            &HashMap::new(),
            None,
            &mut runs,
            &mut Vec::new(),
            &mut Vec::new(),
        );
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
        extract_shape_text(
            &stream,
            0,
            &[],
            &HashMap::new(),
            None,
            &mut runs,
            &mut Vec::new(),
            &mut Vec::new(),
        );
        assert_eq!(runs.len(), 2);
        assert_eq!(runs[0].text, "Title");
        assert_eq!(runs[1].text, "Body text");
    }

    /// Build one `ExHyperlinkContainer`: `ExHyperlinkAtom` (id) +
    /// `TargetAtom` (a `RT_CSTRING` at `CSTRING_INSTANCE_TARGET`, UTF-16LE).
    fn make_ex_hyperlink(id: u32, url: &str) -> Vec<u8> {
        let mut children = make_atom(RT_EXTERNAL_HYPERLINK_ATOM, 0, &id.to_le_bytes());
        let utf16: Vec<u8> = url.encode_utf16().flat_map(|u| u.to_le_bytes()).collect();
        children.extend(make_atom(RT_CSTRING, CSTRING_INSTANCE_TARGET, &utf16));
        make_container(RT_EXTERNAL_HYPERLINK, 0, &children)
    }

    /// Build one shape's `InteractiveInfoAtom` body: `soundIdRef`(4)=0,
    /// `exHyperlinkIdRef`(4), `action`(1), `oleVerb`(1)=0, `jump`(1)=0,
    /// `flags`(1)=0, `hyperlinkType`(1)=0, `unused`(3)=0 — 16 bytes total
    /// ([MS-PPT] 2.6.10, `rh.recLen` MUST be 0x10).
    fn interactive_info_atom_body(ex_hyperlink_id_ref: u32, action: u8) -> Vec<u8> {
        let mut d = vec![0u8; 16];
        d[4..8].copy_from_slice(&ex_hyperlink_id_ref.to_le_bytes());
        d[8] = action;
        d
    }

    /// Build a shape (`RT_SHAPE`) with a `ClientTextbox` carrying the given
    /// text and a `ClientData`/`InteractiveInfo` referencing
    /// `ex_hyperlink_id_ref` via a click action (`II_HyperlinkAction`).
    fn make_hyperlinked_shape(text_type: u32, text: &[u8], ex_hyperlink_id_ref: u32) -> Vec<u8> {
        let mut textbox_children = make_atom(RT_TEXT_HEADER, 0, &text_type.to_le_bytes());
        textbox_children.extend(make_atom(RT_TEXT_BYTES, 0, text));
        let textbox = make_container(0xF00D, 0, &textbox_children);

        let atom = interactive_info_atom_body(ex_hyperlink_id_ref, 0x04); // II_HyperlinkAction
        let interactive_info =
            make_container(RT_INTERACTIVE_INFO, 0, &make_atom(RT_INTERACTIVE_INFO_ATOM, 0, &atom));
        let client_data = make_container(RT_CLIENT_DATA, 0, &interactive_info);

        let mut shape_children = client_data;
        shape_children.extend(&textbox);
        make_container(RT_SHAPE, 0, &shape_children)
    }

    // ── #256: picture-shape `pib` resolution ──

    /// Build an `OfficeArtFOPT` record (issue #256) with a single
    /// `Blip:pib` property entry (`opid.opid = 0x0104`, `fBid = 1`,
    /// `fComplex = 0`), `op = pib_value` (the documented ONE-based
    /// index).
    fn make_fopt_with_pib(pib_value: u32) -> Vec<u8> {
        let opid: u16 = 0x0104 | (1 << 14); // pid=0x0104, fBid=1, fComplex=0
        let mut entry = Vec::new();
        entry.extend_from_slice(&opid.to_le_bytes());
        entry.extend_from_slice(&pib_value.to_le_bytes());

        // rh.recVer=0x3, rh.recInstance=1 (one property), rh.recType=0xF00B.
        let ver_inst: u16 = 0x3 | (1 << 4);
        let mut buf = Vec::new();
        buf.extend_from_slice(&ver_inst.to_le_bytes());
        buf.extend_from_slice(&RT_FOPT.to_le_bytes());
        buf.extend_from_slice(&(entry.len() as u32).to_le_bytes());
        buf.extend(entry);
        buf
    }

    #[test]
    fn resolve_shape_pib_finds_the_pib_property() {
        let shape_data = make_fopt_with_pib(3); // one-based
        assert_eq!(resolve_shape_pib(&shape_data), Some(2)); // zero-based
    }

    #[test]
    fn resolve_shape_pib_zero_means_no_picture() {
        let shape_data = make_fopt_with_pib(0);
        assert_eq!(resolve_shape_pib(&shape_data), None);
    }

    #[test]
    fn resolve_shape_pib_none_without_fopt() {
        assert_eq!(resolve_shape_pib(&[]), None);
    }

    #[test]
    fn resolve_shape_pib_skips_unrelated_properties() {
        // Two properties: an unrelated one (pid=0x0080, some Shape
        // Boolean property) first, then pib — the scan must not stop at
        // the first entry.
        let unrelated_opid: u16 = 0x0080;
        let mut entries = Vec::new();
        entries.extend_from_slice(&unrelated_opid.to_le_bytes());
        entries.extend_from_slice(&0xFFFFFFFFu32.to_le_bytes());
        let pib_opid: u16 = 0x0104 | (1 << 14);
        entries.extend_from_slice(&pib_opid.to_le_bytes());
        entries.extend_from_slice(&5u32.to_le_bytes());

        let ver_inst: u16 = 0x3 | (2 << 4); // 2 properties
        let mut shape_data = Vec::new();
        shape_data.extend_from_slice(&ver_inst.to_le_bytes());
        shape_data.extend_from_slice(&RT_FOPT.to_le_bytes());
        shape_data.extend_from_slice(&(entries.len() as u32).to_le_bytes());
        shape_data.extend(entries);

        assert_eq!(resolve_shape_pib(&shape_data), Some(4));
    }

    #[test]
    fn shape_with_pib_reaches_image_refs_end_to_end() {
        let mut shape_children = make_fopt_with_pib(1); // zero-based index 0
        let mut textbox_children = make_atom(RT_TEXT_HEADER, 0, &4u32.to_le_bytes());
        textbox_children.extend(make_atom(RT_TEXT_BYTES, 0, b"caption"));
        shape_children.extend(make_container(0xF00D, 0, &textbox_children));
        let shape = make_container(RT_SHAPE, 0, &shape_children);

        let mut runs = Vec::new();
        let mut tables = Vec::new();
        let mut image_refs = Vec::new();
        extract_shape_text(
            &shape,
            0,
            &[],
            &HashMap::new(),
            None,
            &mut runs,
            &mut tables,
            &mut image_refs,
        );

        assert_eq!(image_refs, vec![0]);
        assert_eq!(runs.len(), 1, "the shape's own text must still be extracted alongside its pib");
    }

    #[test]
    fn parse_ex_hyperlinks_resolves_id_to_target_url() {
        let ex_obj_list =
            make_container(RT_EXTERNAL_OBJECT_LIST, 0, &make_ex_hyperlink(1, "http://example.com"));
        let map = parse_ex_hyperlinks(&ex_obj_list);
        assert_eq!(map.get(&1).map(String::as_str), Some("http://example.com"));
    }

    #[test]
    fn parse_ex_hyperlinks_multiple_entries() {
        let mut ex_obj_list_children = make_ex_hyperlink(1, "http://a.example/");
        ex_obj_list_children.extend(make_ex_hyperlink(2, "http://b.example/"));
        let ex_obj_list = make_container(RT_EXTERNAL_OBJECT_LIST, 0, &ex_obj_list_children);
        let map = parse_ex_hyperlinks(&ex_obj_list);
        assert_eq!(map.len(), 2);
        assert_eq!(map.get(&1).map(String::as_str), Some("http://a.example/"));
        assert_eq!(map.get(&2).map(String::as_str), Some("http://b.example/"));
    }

    #[test]
    fn empty_stream_yields_no_hyperlinks() {
        assert!(parse_ex_hyperlinks(&[]).is_empty());
    }

    /// issue #257 — the real end-to-end chain: a shape's own
    /// `InteractiveInfoAtom` (`exHyperlinkIdRef` + `II_HyperlinkAction`)
    /// resolves through the document's `ExObjList` to a real URL, which
    /// ends up on the shape's own `TextRun::hyperlink`.
    #[test]
    fn shape_with_interactive_info_resolves_its_hyperlink() {
        let mut hyperlinks = HashMap::new();
        hyperlinks.insert(1u32, "http://testuri.org/".to_string());

        let shape = make_hyperlinked_shape(1, b"Click here", 1); // Body, exHyperlinkIdRef=1
        let mut runs = Vec::new();
        extract_shape_text(
            &shape,
            0,
            &[],
            &hyperlinks,
            None,
            &mut runs,
            &mut Vec::new(),
            &mut Vec::new(),
        );

        assert_eq!(runs.len(), 1);
        assert_eq!(runs[0].text, "Click here");
        assert_eq!(runs[0].hyperlink.as_deref(), Some("http://testuri.org/"));
    }

    /// issue #257 — the far more common, real-world shape: a hyperlink
    /// covering only PART of a text run's characters, via a sibling
    /// `MouseClickInteractiveInfoContainer` + `MouseClickTextInteractiveInfoAtom`
    /// pair in the `ClientTextbox` (not nested in `RT_CLIENT_DATA` at all —
    /// confirmed against real corpus bytes, not just the spec). The run
    /// must split into unlinked-prefix / linked / unlinked-suffix pieces.
    #[test]
    fn text_range_hyperlink_splits_the_run() {
        let mut hyperlinks = HashMap::new();
        hyperlinks.insert(7u32, "http://example.com/".to_string());

        let mut textbox_children = make_atom(RT_TEXT_HEADER, 0, &1u32.to_le_bytes()); // Body
        textbox_children.extend(make_atom(RT_TEXT_BYTES, 0, b"See http://example.com/ here"));
        let atom = interactive_info_atom_body(7, 0x04); // II_HyperlinkAction
        textbox_children.extend(make_container(
            RT_INTERACTIVE_INFO,
            0,
            &make_atom(RT_INTERACTIVE_INFO_ATOM, 0, &atom),
        ));
        // TextRange: begin=4, end=23 -> "http://example.com/" (chars 4..23).
        let mut range = Vec::new();
        range.extend_from_slice(&4i32.to_le_bytes());
        range.extend_from_slice(&23i32.to_le_bytes());
        textbox_children.extend(make_atom(RT_TEXT_INTERACTIVE_INFO_ATOM, 0, &range));
        let textbox = make_container(0xF00D, 0, &textbox_children);
        let shape = make_container(RT_SHAPE, 0, &textbox);

        let mut runs = Vec::new();
        extract_shape_text(
            &shape,
            0,
            &[],
            &hyperlinks,
            None,
            &mut runs,
            &mut Vec::new(),
            &mut Vec::new(),
        );

        assert_eq!(runs.len(), 3, "must split into prefix/linked/suffix: {runs:?}");
        assert_eq!(runs[0].text, "See ");
        assert_eq!(runs[0].hyperlink, None);
        assert_eq!(runs[1].text, "http://example.com/");
        assert_eq!(runs[1].hyperlink.as_deref(), Some("http://example.com/"));
        assert_eq!(runs[2].text, " here");
        assert_eq!(runs[2].hyperlink, None);
    }

    /// The hyperlinked range can cover the WHOLE run (no unlinked prefix
    /// or suffix) — must produce exactly one run, not empty placeholders.
    #[test]
    fn text_range_hyperlink_covering_the_whole_run_produces_one_run() {
        let mut hyperlinks = HashMap::new();
        hyperlinks.insert(1u32, "http://example.com/".to_string());

        let mut textbox_children = make_atom(RT_TEXT_HEADER, 0, &1u32.to_le_bytes());
        textbox_children.extend(make_atom(RT_TEXT_BYTES, 0, b"clickme"));
        let atom = interactive_info_atom_body(1, 0x04);
        textbox_children.extend(make_container(
            RT_INTERACTIVE_INFO,
            0,
            &make_atom(RT_INTERACTIVE_INFO_ATOM, 0, &atom),
        ));
        let mut range = Vec::new();
        range.extend_from_slice(&0i32.to_le_bytes());
        range.extend_from_slice(&7i32.to_le_bytes());
        textbox_children.extend(make_atom(RT_TEXT_INTERACTIVE_INFO_ATOM, 0, &range));
        let textbox = make_container(0xF00D, 0, &textbox_children);
        let shape = make_container(RT_SHAPE, 0, &textbox);

        let mut runs = Vec::new();
        extract_shape_text(
            &shape,
            0,
            &[],
            &hyperlinks,
            None,
            &mut runs,
            &mut Vec::new(),
            &mut Vec::new(),
        );

        assert_eq!(runs.len(), 1);
        assert_eq!(runs[0].text, "clickme");
        assert_eq!(runs[0].hyperlink.as_deref(), Some("http://example.com/"));
    }

    // ── #255: grid-of-shapes table reconstruction ──

    fn make_child_anchor(left: i32, top: i32, right: i32, bottom: i32) -> Vec<u8> {
        let mut body = Vec::new();
        body.extend_from_slice(&left.to_le_bytes());
        body.extend_from_slice(&top.to_le_bytes());
        body.extend_from_slice(&right.to_le_bytes());
        body.extend_from_slice(&bottom.to_le_bytes());
        make_atom(RT_CHILD_ANCHOR, 0, &body)
    }

    fn make_table_cell_shape(left: i32, top: i32, text: &[u8]) -> Vec<u8> {
        let mut children = make_child_anchor(left, top, left + 100, top + 50);
        // Tx_TYPE_OTHER
        let mut textbox_children = make_atom(RT_TEXT_HEADER, 0, &4u32.to_le_bytes());
        textbox_children.extend(make_atom(RT_TEXT_BYTES, 0, text));
        children.extend(make_container(0xF00D, 0, &textbox_children));
        make_container(RT_SHAPE, 0, &children)
    }

    /// issue #255 — a group whose members form a clean 2x2 grid must
    /// become one `TableBlock`, and the cell text must NOT also appear as
    /// flat paragraphs (that would duplicate it in the IR).
    #[test]
    fn spgr_container_with_clean_grid_becomes_a_table() {
        let mut spgr_children = make_container(RT_SHAPE, 0, &[]); // group's own placeholder shape
        spgr_children.extend(make_table_cell_shape(0, 0, b"A1"));
        spgr_children.extend(make_table_cell_shape(100, 0, b"B1"));
        spgr_children.extend(make_table_cell_shape(0, 50, b"A2"));
        spgr_children.extend(make_table_cell_shape(100, 50, b"B2"));
        let spgr = make_container(RT_SPGR_CONTAINER, 0, &spgr_children);

        let mut runs = Vec::new();
        let mut tables = Vec::new();
        extract_shape_text(
            &spgr,
            0,
            &[],
            &HashMap::new(),
            None,
            &mut runs,
            &mut tables,
            &mut Vec::new(),
        );

        assert!(runs.is_empty(), "cell text must not also appear as flat paragraphs: {runs:?}");
        assert_eq!(tables.len(), 1);
        let t = &tables[0];
        assert_eq!(t.rows.len(), 2);
        assert_eq!(t.rows[0].len(), 2);
        assert_eq!(t.rows[0][0][0].text, "A1");
        assert_eq!(t.rows[0][1][0].text, "B1");
        assert_eq!(t.rows[1][0][0].text, "A2");
        assert_eq!(t.rows[1][1][0].text, "B2");
    }

    /// issue #255 — a group that ISN'T a clean grid (here: only 3
    /// members, one short of a 2x2) must fall back to the ordinary
    /// flat-paragraph group walk, unchanged from before this feature
    /// existed — no text lost, just no table structure.
    #[test]
    fn spgr_container_that_is_not_a_grid_falls_back_to_flat_paragraphs() {
        let mut spgr_children = make_container(RT_SHAPE, 0, &[]);
        spgr_children.extend(make_table_cell_shape(0, 0, b"One"));
        spgr_children.extend(make_table_cell_shape(100, 0, b"Two"));
        spgr_children.extend(make_table_cell_shape(0, 50, b"Three"));
        let spgr = make_container(RT_SPGR_CONTAINER, 0, &spgr_children);

        let mut runs = Vec::new();
        let mut tables = Vec::new();
        extract_shape_text(
            &spgr,
            0,
            &[],
            &HashMap::new(),
            None,
            &mut runs,
            &mut tables,
            &mut Vec::new(),
        );

        assert!(tables.is_empty());
        assert_eq!(runs.len(), 3);
    }

    /// issue #255 — a group member with no `RT_CHILD_ANCHOR` at all (an
    /// odd/unexpected shape) must also bail to the flat-paragraph
    /// fallback rather than guessing a position for it.
    #[test]
    fn spgr_container_member_without_child_anchor_falls_back() {
        let mut spgr_children = make_container(RT_SHAPE, 0, &[]);
        spgr_children.extend(make_table_cell_shape(0, 0, b"A1"));
        spgr_children.extend(make_table_cell_shape(100, 0, b"B1"));
        spgr_children.extend(make_table_cell_shape(0, 50, b"A2"));
        // Fourth shape has text but no ChildAnchor at all.
        let mut textbox_children = make_atom(RT_TEXT_HEADER, 0, &4u32.to_le_bytes());
        textbox_children.extend(make_atom(RT_TEXT_BYTES, 0, b"NoAnchor"));
        let textbox = make_container(0xF00D, 0, &textbox_children);
        spgr_children.extend(make_container(RT_SHAPE, 0, &textbox));
        let spgr = make_container(RT_SPGR_CONTAINER, 0, &spgr_children);

        let mut runs = Vec::new();
        let mut tables = Vec::new();
        extract_shape_text(
            &spgr,
            0,
            &[],
            &HashMap::new(),
            None,
            &mut runs,
            &mut tables,
            &mut Vec::new(),
        );

        assert!(tables.is_empty());
        assert_eq!(runs.len(), 4);
    }

    /// A shape with no `InteractiveInfo` at all must never get a
    /// hyperlink, even when the document has some hyperlinks elsewhere.
    #[test]
    fn shape_without_interactive_info_has_no_hyperlink() {
        let mut hyperlinks = HashMap::new();
        hyperlinks.insert(1u32, "http://testuri.org/".to_string());

        let header = make_atom(RT_TEXT_HEADER, 0, &1u32.to_le_bytes());
        let mut textbox_children = header;
        textbox_children.extend(make_atom(RT_TEXT_BYTES, 0, b"Plain text"));
        let textbox = make_container(0xF00D, 0, &textbox_children);
        let shape = make_container(RT_SHAPE, 0, &textbox);

        let mut runs = Vec::new();
        extract_shape_text(
            &shape,
            0,
            &[],
            &hyperlinks,
            None,
            &mut runs,
            &mut Vec::new(),
            &mut Vec::new(),
        );
        assert_eq!(runs[0].hyperlink, None);
    }

    /// `II_NoAction` (0x00) must never resolve a hyperlink, even when
    /// `exHyperlinkIdRef` happens to name a real entry — the action type
    /// gates whether the id is meaningful at all ([MS-PPT] 2.6.10).
    #[test]
    fn non_hyperlink_action_does_not_resolve_a_hyperlink() {
        let mut hyperlinks = HashMap::new();
        hyperlinks.insert(1u32, "http://testuri.org/".to_string());

        let textbox_children = {
            let mut c = make_atom(RT_TEXT_HEADER, 0, &1u32.to_le_bytes());
            c.extend(make_atom(RT_TEXT_BYTES, 0, b"Not a link"));
            c
        };
        let textbox = make_container(0xF00D, 0, &textbox_children);
        let atom = interactive_info_atom_body(1, 0x00); // II_NoAction
        let interactive_info =
            make_container(RT_INTERACTIVE_INFO, 0, &make_atom(RT_INTERACTIVE_INFO_ATOM, 0, &atom));
        let client_data = make_container(RT_CLIENT_DATA, 0, &interactive_info);
        let mut shape_children = client_data;
        shape_children.extend(&textbox);
        let shape = make_container(RT_SHAPE, 0, &shape_children);

        let mut runs = Vec::new();
        extract_shape_text(
            &shape,
            0,
            &[],
            &hyperlinks,
            None,
            &mut runs,
            &mut Vec::new(),
            &mut Vec::new(),
        );
        assert_eq!(runs[0].hyperlink, None);
    }

    /// issue #257 — full public pipeline: `extract_slides_text` on a
    /// synthetic "PowerPoint Document" stream carrying both an
    /// `ExObjListContainer` and a hyperlinked shape (with no persist
    /// directory or `SlideListWithText` — the "last resort" fallback,
    /// which still builds `hyperlinks` from the *whole* stream first)
    /// resolves the shape's `TextRun::hyperlink` end to end.
    #[test]
    fn extract_slides_text_resolves_hyperlinks_end_to_end() {
        let shape = make_hyperlinked_shape(1, b"Hyperlink text", 1);
        let mut stream = make_container(
            RT_EXTERNAL_OBJECT_LIST,
            0,
            &make_ex_hyperlink(1, "http://testuri.org/"),
        );
        stream.extend(&shape);

        let slides = extract_slides_text(&stream, None);
        assert_eq!(slides.len(), 1);
        assert_eq!(slides[0].text_runs[0].text, "Hyperlink text");
        assert_eq!(slides[0].text_runs[0].hyperlink.as_deref(), Some("http://testuri.org/"));
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
        extract_shape_text(
            &shape,
            0,
            &[],
            &HashMap::new(),
            None,
            &mut runs,
            &mut Vec::new(),
            &mut Vec::new(),
        );
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

    /// issue #253 — every `TextTypeEnum` value from 3 upward was shifted
    /// one slot low (5 misread as `CenterTitle`, 6 as `HalfBody`, …),
    /// swapping a title slide's real title and subtitle. Values verified
    /// against the published [MS-PPT] `TextTypeEnum` spec page directly:
    /// 3 is genuinely undefined (the enum jumps from 2 to 4).
    #[test]
    fn text_type_values_3_and_up_match_the_spec_not_the_old_shifted_mapping() {
        assert_eq!(TextType::from_u32(3), TextType::Other, "3 is undefined in the spec");
        assert_eq!(TextType::from_u32(4), TextType::Other, "4 = Tx_TYPE_OTHER");
        assert_eq!(TextType::from_u32(5), TextType::CenterBody, "5 = Tx_TYPE_CENTERBODY");
        assert_eq!(TextType::from_u32(6), TextType::CenterTitle, "6 = Tx_TYPE_CENTERTITLE");
        assert_eq!(TextType::from_u32(7), TextType::HalfBody, "7 = Tx_TYPE_HALFBODY");
        assert_eq!(TextType::from_u32(8), TextType::QuarterBody, "8 = Tx_TYPE_QUARTERBODY");
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

    /// issue #257 — `ExObjListContainer` resolution must go through the
    /// same persist-directory mechanism slides do, not a raw whole-stream
    /// scan: a stale/orphaned `ExObjListContainer` left behind by an
    /// earlier incremental save (mapping the same `exHyperlinkId` to a
    /// *different*, superseded URL) must never win over the current,
    /// persist-resolved one.
    #[test]
    fn persist_resolution_uses_the_current_exobjlist_not_a_stale_one() {
        let mut stream = Vec::new();

        // Stale, orphaned ExObjListContainer — not reachable from the
        // current, persist-resolved DocumentContainer.
        let stale_ex_obj_list = make_container(
            RT_EXTERNAL_OBJECT_LIST,
            0,
            &make_ex_hyperlink(1, "http://stale.example/"),
        );
        stream.extend(&stale_ex_obj_list);

        // The real slide: a hyperlinked shape referencing exHyperlinkId=1.
        let shape = make_hyperlinked_shape(1, b"REAL SLIDE TEXT", 1);
        let real_slide = make_container(RT_SLIDE, 0, &shape);

        // The real, current DocumentContainer: SlideListWithText +
        // its own ExObjListContainer (same id, real URL).
        let doc_offset = stream.len() as u32;
        let slide_persist = slide_persist_atom_bytes(2, 256);
        let slide_list = make_container(RT_SLIDE_LIST_WITH_TEXT, SLWT_SLIDES, &slide_persist);
        let real_ex_obj_list = make_container(
            RT_EXTERNAL_OBJECT_LIST,
            0,
            &make_ex_hyperlink(1, "http://real.example/"),
        );
        let mut doc_children = slide_list;
        doc_children.extend(&real_ex_obj_list);
        stream.extend(make_container(RT_DOCUMENT, 0, &doc_children));

        let real_slide_offset = stream.len() as u32;
        stream.extend(&real_slide);

        let pd_offset = stream.len() as u32;
        stream.extend(persist_directory_bytes(&[(1, doc_offset), (2, real_slide_offset)]));
        let edit_offset = stream.len() as u32;
        stream.extend(user_edit_atom_bytes(0, pd_offset, 1));
        let current_user = current_user_bytes(edit_offset);

        let slides = extract_slides_text(&stream, Some(&current_user));
        assert_eq!(slides.len(), 1);
        assert_eq!(slides[0].text_runs[0].text, "REAL SLIDE TEXT");
        assert_eq!(
            slides[0].text_runs[0].hyperlink.as_deref(),
            Some("http://real.example/"),
            "the current ExObjList's URL must win over the stale one"
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
                hyperlink: None,
                ..Default::default()
            },
            TextRun {
                text_type: TextType::Body,
                text: "second".into(),
                hyperlink: None,
                ..Default::default()
            },
        ];

        let mut index_bytes = Vec::new();
        index_bytes.extend_from_slice(&1i32.to_le_bytes());
        let outline_ref = make_atom(RT_OUTLINE_TEXT_REF_ATOM, 0, &index_bytes);
        let textbox = make_container(0xF00D, 0, &outline_ref);
        let shape = make_container(0xF004, 0, &textbox);

        let mut runs = Vec::new();
        extract_shape_text(
            &shape,
            0,
            &outline_texts,
            &HashMap::new(),
            None,
            &mut runs,
            &mut Vec::new(),
            &mut Vec::new(),
        );
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
