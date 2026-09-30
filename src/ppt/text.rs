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
    /// presentations. Verified against the published
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
    /// has no interactive info, or its action has no hyperlink to resolve.
    pub hyperlink: Option<String>,
    /// Direct character-level formatting (bold/italic/underline/size/
    /// color/position), resolved from this run's `StyleTextPropAtom` if
    /// one was present. Empty when no such atom was found — callers must
    /// treat that as "no formatting info", not "definitely unformatted".
    pub char_formats: Vec<CharFormatSpan>,
    /// Direct paragraph-level formatting (currently: alignment only),
    /// resolved the same way. Empty when no `StyleTextPropAtom` was found.
    pub para_formats: Vec<ParaFormatSpan>,
    /// This run's shape's placeholder role (`OEPlaceholderAtom.placeholderId`,
    /// mapped to the same `ST_PlaceholderType` string vocabulary PPTX's own
    /// `<p:ph type="...">` uses), resolved from the shape actually carrying
    /// it — `None` when the shape has no `OEPlaceholderAtom`, or its
    /// `placeholderId` has no OOXML-equivalent role.
    pub placeholder_role: Option<String>,
    /// Hyperlinks covering part of the text (a `TextInteractiveInfoAtom`
    /// range), in `char` indices. The run stays one piece of text — a link
    /// inside a sentence does not break the sentence.
    pub link_ranges: Vec<LinkRange>,
}

/// A hyperlink over `[start, end)` (`char` indices) of a run's text.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LinkRange {
    /// Start character offset (inclusive).
    pub start: usize,
    /// End character offset (exclusive).
    pub end: usize,
    /// The link target.
    pub url: String,
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
#[cfg(test)]
pub fn extract_slides_text(stream: &[u8], current_user: Option<&[u8]>) -> Vec<SlideText> {
    extract_deck_text(stream, current_user).slides
}

/// What a "PowerPoint Document" stream yields: per-slide text, and the
/// static text of the masters those slides show.
#[derive(Debug, Clone, Default)]
pub struct DeckText {
    /// One entry per slide, in presentation order.
    pub slides: Vec<SlideText>,
    /// The static (non-placeholder) text of every master at least one
    /// slide shows, each distinct text once; see [`master_static_text`].
    pub master_text: Vec<TextRun>,
}

/// Extract the deck's slide text (resolved as described on
/// `extract_slides_text`) and its masters' static text.
pub fn extract_deck_text(stream: &[u8], current_user: Option<&[u8]>) -> DeckText {
    let mut resolved = DeckText::default();
    if let Some(dir) = persist::build(stream, current_user) {
        if let Some(deck) = extract_slides_via_persist(stream, &dir) {
            // A resolved-but-entirely-textless result is ambiguous: it's the
            // correct answer for a genuinely text-free deck (image-only
            // slides, or slides whose only text is on their master), but
            // it's also what a corrupted persist chain that resolved to the
            // wrong offsets looks like. Fall through to the weaker
            // heuristics below rather than committing to it — if they also
            // come up empty, this was the right answer all along, and the
            // resolved slides are kept.
            if !deck.slides.is_empty() && deck.slides.iter().any(|s| !s.text_runs.is_empty()) {
                return deck;
            }
            resolved = deck;
        }
    }
    let slides = extract_slides_without_persist(stream);
    DeckText {
        slides: if slides.is_empty() {
            resolved.slides
        } else {
            slides
        },
        // Which masters the slides show is known only through the persist
        // directory; the weaker paths cannot tell, so they add none.
        master_text: resolved.master_text,
    }
}

/// The slide-text heuristics used when the persist directory yields no
/// text: the slide list's inline text cache, then every `Slide` container
/// in the stream, then every text atom outside the masters and notes.
fn extract_slides_without_persist(stream: &[u8]) -> Vec<SlideText> {
    // A best-effort fallback scope for the weaker paths below, which have
    // no persist-resolved DocumentContainer to search within at all.
    let stream_wide_hyperlinks = parse_ex_hyperlinks(stream);
    let stream_wide_ole_objects = parse_ex_ole_objects(stream);

    if let Some(slide_list) = find_descendant(stream, RT_SLIDE_LIST_WITH_TEXT, SLWT_SLIDES, 0) {
        let slides = extract_slides_from_slide_list_cache(&slide_list);
        if !slides.is_empty() {
            return slides;
        }
    }

    // Weaker fallback: walk only `Slide` containers found anywhere in the
    // stream. A whole-stream scan also picks up `MainMaster` containers,
    // whose placeholder prompts ("Click to edit Master title style", in
    // whatever language PowerPoint was running) are its own UI strings and
    // never render on a slide — in two POI corpus files they were 116 of
    // the 137 and 121 extracted characters respectively.
    let mut slides = Vec::new();
    collect_slide_containers(
        stream,
        0,
        &stream_wide_hyperlinks,
        &stream_wide_ole_objects,
        &mut slides,
    );
    if slides.iter().any(|s| !s.text_runs.is_empty()) {
        return slides;
    }

    // Last resort: no resolvable structure at all — dump whatever text atoms
    // exist anywhere in the stream, except inside the masters (slide, title,
    // notes and handout masters), whose text is prompt boilerplate. That is
    // decided by structure, not by matching the English prompt strings: a
    // localized master's prompts leaked, and a slide's own text that
    // happened to start with "Click to add title" was deleted. Notes pages
    // are skipped too — speaker notes are not slide text.
    let mut content = Vec::new();
    for rec in RecordIter::new(stream) {
        let Ok(rec) = rec else { break };
        if matches!(rec.header.rec_type, RT_MAIN_MASTER | RT_NOTES | RT_HANDOUT) {
            continue;
        }
        let end = rec.offset + 8 + rec.data.len();
        content.extend_from_slice(&stream[rec.offset..end]);
    }
    let mut runs = Vec::new();
    let mut tables = Vec::new();
    let mut image_refs = Vec::new();
    let mut ole_object_refs = Vec::new();
    extract_shape_text(
        &content,
        0,
        &[],
        &stream_wide_hyperlinks,
        &stream_wide_ole_objects,
        None,
        None,
        &mut runs,
        &mut tables,
        &mut image_refs,
        &mut ole_object_refs,
    );
    runs.retain(|r| !is_field_placeholder_only(&r.text));
    if runs.is_empty() && tables.is_empty() && image_refs.is_empty() {
        Vec::new()
    } else {
        vec![SlideText {
            text_runs: runs,
            tables,
            image_refs,
            ole_object_refs,
            ..Default::default()
        }]
    }
}

/// Walk the record tree collecting one `SlideText` per `Slide` container.
fn collect_slide_containers(
    data: &[u8],
    depth: usize,
    hyperlinks: &HashMap<u32, String>,
    ole_objects: &HashMap<u32, OleObjectInfo>,
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
            let mut ole_object_refs = Vec::new();
            extract_shape_text(
                &rec.data,
                0,
                &[],
                hyperlinks,
                ole_objects,
                None,
                None,
                &mut runs,
                &mut tables,
                &mut image_refs,
                &mut ole_object_refs,
            );
            runs.retain(|r| !is_field_placeholder_only(&r.text));
            let hidden = slide_is_hidden(&rec.data);
            out.push(SlideText {
                text_runs: runs,
                tables,
                image_refs,
                ole_object_refs,
                hidden,
            });
            continue;
        }
        if rec.header.is_container() {
            collect_slide_containers(&rec.data, depth + 1, hyperlinks, ole_objects, out);
        }
    }
}

/// Whether a text run is nothing but `*` — the stand-in character
/// PowerPoint stores for a date, slide-number or footer field placeholder,
/// not text anyone typed. Language-independent, unlike a prompt string.
fn is_field_placeholder_only(text: &str) -> bool {
    let t = text.trim();
    !t.is_empty() && t.chars().all(|c| c == '*' || c.is_whitespace())
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
fn extract_slides_via_persist(stream: &[u8], dir: &PersistDirectory) -> Option<DeckText> {
    let doc_offset = dir.resolve(dir.doc_persist_id)?;
    let doc_children = bounded_container_children(stream, doc_offset, RT_DOCUMENT)?;
    let slide_list = find_child(&doc_children, RT_SLIDE_LIST_WITH_TEXT, SLWT_SLIDES)?;
    // The current DocumentContainer's own ExObjListContainer — not a raw
    // whole-stream scan, which could resolve a stale, superseded copy left
    // behind by an earlier incremental save (the same hazard the persist
    // directory itself exists to route around for slides).
    let hyperlinks = parse_ex_hyperlinks(&doc_children);
    let ole_objects = parse_ex_ole_objects(&doc_children);
    // The deck's slide header/footer settings ("Apply to All" in
    // PowerPoint's Header and Footer dialog): the DocumentContainer's
    // `SlideHeadersFootersContainer`, which a slide's own container
    // overrides.
    let deck_headers_footers = find_child(&doc_children, RT_HEADER_FOOTER, HF_INSTANCE_SLIDES)
        .map(|hf| shown_slide_header_footer_texts(&hf))
        .unwrap_or_default();

    let fonts = font_collection(&doc_children);

    let mut slides = Vec::new();
    // Per slide: its `SlideId` and `SlideAtom.notesIdRef`, to attach notes.
    let mut slide_links: Vec<(u32, Option<u32>)> = Vec::new();
    // Per slide: its `(masterIdRef, fMasterObjects)`.
    let mut master_refs: Vec<(u32, bool)> = Vec::new();
    let master_persist_ids = master_persist_ids(&doc_children);
    for entry in slide_list_entries(&slide_list) {
        let (mut slide, links) = resolve_slide(
            stream,
            dir,
            &master_persist_ids,
            entry.persist_id_ref,
            &entry.outline_texts,
            &hyperlinks,
            &ole_objects,
        );
        let hf_texts = links
            .headers_footers
            .as_deref()
            .unwrap_or(&deck_headers_footers);
        for text in hf_texts {
            slide.text_runs.push(TextRun {
                text_type: TextType::Other,
                text: text.clone(),
                hyperlink: None,
                ..Default::default()
            });
        }
        slides.push(slide);
        slide_links.push((entry.slide_id, links.notes_id_ref));
        master_refs.extend(links.master);
    }

    // Speaker notes: the `NotesListWithTextContainer` lists every notes
    // page; a slide names its own through `SlideAtom.notesIdRef` (matching
    // the notes entry's `SlidePersistAtom.slideId`, [MS-PPT]
    // `NotesIdRef`), and each notes page names its slide through
    // `NotesAtom.slideIdRef` — used when the slide's reference is absent.
    if let Some(notes_list) = find_child(&doc_children, RT_SLIDE_LIST_WITH_TEXT, SLWT_NOTES) {
        let mut by_notes_id: HashMap<u32, Vec<TextRun>> = HashMap::new();
        let mut by_slide_id: HashMap<u32, Vec<TextRun>> = HashMap::new();
        for entry in slide_list_entries(&notes_list) {
            let Some((runs, slide_id_ref)) = resolve_notes(
                stream,
                dir,
                entry.persist_id_ref,
                &entry.outline_texts,
                &hyperlinks,
                &ole_objects,
            ) else {
                continue;
            };
            if let Some(sid) = slide_id_ref {
                by_slide_id.entry(sid).or_insert_with(|| runs.clone());
            }
            by_notes_id.entry(entry.slide_id).or_insert(runs);
        }
        for (slide, &(slide_id, notes_id_ref)) in slides.iter_mut().zip(&slide_links) {
            let notes = notes_id_ref
                .filter(|&id| id != 0)
                .and_then(|id| by_notes_id.get(&id))
                .or_else(|| by_slide_id.get(&slide_id));
            if let Some(runs) = notes {
                slide.text_runs.extend(runs.iter().cloned());
            }
        }
    }

    let mut master_text = master_static_text(
        stream,
        dir,
        &master_persist_ids,
        master_refs.iter().copied(),
        &hyperlinks,
        &ole_objects,
    );

    // Every run's `fontRef` names a font in the deck's collection.
    if !fonts.is_empty() {
        for run in slides
            .iter_mut()
            .flat_map(|s| s.text_runs.iter_mut())
            .chain(master_text.iter_mut())
        {
            for span in &mut run.char_formats {
                span.format.typeface = span
                    .format
                    .font_ref
                    .and_then(|r| fonts.get(r as usize))
                    .cloned();
            }
        }
    }

    Some(DeckText {
        slides,
        master_text,
    })
}

/// `masterId` → `persistIdRef` for every entry of the
/// `MasterListWithTextContainer` ([MS-PPT] `MasterListWithTextContainer`,
/// a `SlideListWithText` record with `recInstance` 1; each child is a
/// `MasterPersistAtom`, whose `masterId` sits at body offset 12). A
/// `SlideAtom.masterIdRef` names a master by this `masterId` ([MS-PPT]
/// `MasterIdRef`), not by its persist id.
fn master_persist_ids(doc_children: &[u8]) -> HashMap<u32, u32> {
    let Some(list) = find_child(doc_children, RT_SLIDE_LIST_WITH_TEXT, SLWT_MASTERS) else {
        return HashMap::new();
    };
    slide_list_entries(&list)
        .into_iter()
        .map(|e| (e.slide_id, e.persist_id_ref))
        .collect()
}

/// A master resolved from its `masterId`.
enum MasterContainer {
    /// A `MainMasterContainer`'s children.
    Main(Vec<u8>),
    /// A title master — a `SlideContainer` listed in the master list —
    /// whose own `SlideAtom` names its main master.
    Title(Vec<u8>),
}

/// Resolve a `SlideAtom.masterIdRef` to its master: the `masterId` maps to
/// a persist id through the `MasterListWithTextContainer`
/// ([`master_persist_ids`]), and that persist id to a `MainMaster` or a
/// title-master `Slide` container. The one resolver for both master
/// formatting inheritance and master static text. `None` when any link is
/// missing — never a guess from the id's low bits.
fn resolve_master(
    stream: &[u8],
    dir: &PersistDirectory,
    master_persist_ids: &HashMap<u32, u32>,
    master_id: u32,
) -> Option<MasterContainer> {
    let offset = dir.resolve(*master_persist_ids.get(&master_id)?)?;
    if let Some(c) = bounded_container_children(stream, offset, RT_MAIN_MASTER) {
        return Some(MasterContainer::Main(c));
    }
    bounded_container_children(stream, offset, RT_SLIDE).map(MasterContainer::Title)
}

/// Bound on master → master references followed (a title master names
/// its main master); real decks need one hop.
const MAX_MASTER_CHAIN: usize = 8;

/// The static text of every master a slide shows, in first-use order,
/// each distinct text once.
///
/// `slide_masters` yields each slide's `(SlideAtom.masterIdRef,
/// fMasterObjects)`. A slide shows its master's shapes only when
/// `fMasterObjects` is set ([MS-PPT] `SlideAtom.slideFlags`, bit 0: "the
/// slide follows the master objects"); a master none of the slides shows
/// contributes nothing. A title master is itself a `SlideContainer` with a
/// `SlideAtom` naming its main master, followed the same way.
///
/// Only shapes that are not placeholders count — see
/// [`collect_master_static_runs`]. Placeholder text on a master is
/// PowerPoint's prompt ("Click to edit Master title style", in the
/// language PowerPoint ran in), which never appears on a slide.
fn master_static_text(
    stream: &[u8],
    dir: &PersistDirectory,
    master_persist_ids: &HashMap<u32, u32>,
    slide_masters: impl Iterator<Item = (u32, bool)>,
    hyperlinks: &HashMap<u32, String>,
    ole_objects: &HashMap<u32, OleObjectInfo>,
) -> Vec<TextRun> {
    let mut visited = std::collections::HashSet::new();
    let mut runs = Vec::new();
    for (first, follows) in slide_masters {
        let mut next = follows.then_some(first);
        for _ in 0..MAX_MASTER_CHAIN {
            let Some(master_id) = next.take() else { break };
            if master_id == 0 || !visited.insert(master_id) {
                break;
            }
            let children = match resolve_master(stream, dir, master_persist_ids, master_id) {
                Some(MasterContainer::Main(c)) => c,
                Some(MasterContainer::Title(c)) => {
                    // A title master: its own shapes, then its main master's
                    // when it follows that master's objects.
                    next = slide_master_ref(&c).and_then(|(id, follows)| follows.then_some(id));
                    c
                },
                None => break,
            };
            collect_master_static_runs(&children, 0, hyperlinks, ole_objects, &mut runs);
        }
    }
    let mut seen = std::collections::HashSet::new();
    runs.retain(|r| seen.insert(r.text.trim().to_string()));
    runs
}

/// A `Slide`/title-master container's `SlideAtom` `(masterIdRef,
/// fMasterObjects)` ([MS-PPT] `SlideAtom`: `masterIdRef` at body offset
/// 12, `slideFlags` at 20, `fMasterObjects` its bit 0).
fn slide_master_ref(children: &[u8]) -> Option<(u32, bool)> {
    let atom = RecordIter::new(children)
        .filter_map(Result::ok)
        .find(|r| r.header.rec_type == RT_SLIDE_ATOM)?;
    let id = atom.data.get(12..16)?;
    let flags = atom.data.get(20..22)?;
    Some((
        u32::from_le_bytes([id[0], id[1], id[2], id[3]]),
        u16::from_le_bytes([flags[0], flags[1]]) & SLIDE_FLAG_MASTER_OBJECTS != 0,
    ))
}

/// Collect the text of a master's non-placeholder shapes. A shape is a
/// placeholder when its `OfficeArtClientData` carries an
/// `OEPlaceholderAtom` or a `RoundTripHFPlaceholder12Atom`; as a second
/// line of defence only `Tx_TYPE_OTHER` text is kept, since every other
/// text type (title, body, notes and their variants) is placeholder text
/// by definition ([MS-PPT] `TextTypeEnum`). Field stand-ins (`*`) and
/// blank text are dropped too.
fn collect_master_static_runs(
    data: &[u8],
    depth: usize,
    hyperlinks: &HashMap<u32, String>,
    ole_objects: &HashMap<u32, OleObjectInfo>,
    out: &mut Vec<TextRun>,
) {
    if depth > MAX_SHAPE_DEPTH {
        return;
    }
    for rec in RecordIter::new(data) {
        let Ok(rec) = rec else { break };
        if rec.header.rec_type == RT_SHAPE {
            if shape_is_placeholder(&rec.data) {
                continue;
            }
            let hyperlink = resolve_shape_hyperlink(&rec.data, hyperlinks);
            let mut runs = Vec::new();
            let mut tables = Vec::new();
            extract_shape_text(
                &rec.data,
                depth + 1,
                &[],
                hyperlinks,
                ole_objects,
                hyperlink.as_deref(),
                None,
                &mut runs,
                &mut tables,
                &mut Vec::new(),
                &mut Vec::new(),
            );
            let table_runs = tables.into_iter().flat_map(|t| t.rows).flatten().flatten();
            out.extend(runs.into_iter().chain(table_runs).filter(|r| {
                r.text_type == TextType::Other
                    && !r.text.trim().is_empty()
                    && !is_field_placeholder_only(&r.text)
            }));
        } else if rec.header.is_container() {
            collect_master_static_runs(&rec.data, depth + 1, hyperlinks, ole_objects, out);
        }
    }
}

/// Whether a shape is a placeholder: its `OfficeArtClientData` holds an
/// `OEPlaceholderAtom` or a `RoundTripHFPlaceholder12Atom`, whatever the
/// role ([MS-PPT] `OfficeArtClientData`).
fn shape_is_placeholder(shape_data: &[u8]) -> bool {
    let Some(client_data) = find_descendant(shape_data, RT_CLIENT_DATA, 0, 0) else {
        return false;
    };
    find_descendant(&client_data, RT_OE_PLACEHOLDER_ATOM, 0, 0).is_some()
        || find_descendant(&client_data, RT_ROUND_TRIP_HF_PLACEHOLDER12_ATOM, 0, 0).is_some()
}

/// The deck's typeface names, in `FontCollectionContainer` order — the
/// index a `TextCFException.fontRef` uses ([MS-PPT] `FontEntityAtom`:
/// `lfFaceName` is 32 UTF-16 units, NUL-terminated or padded).
fn font_collection(doc_children: &[u8]) -> Vec<String> {
    let Some(env) = find_child(doc_children, RT_ENVIRONMENT, 0) else {
        return Vec::new();
    };
    let Some(collection) = find_child(&env, RT_FONT_COLLECTION, 0) else {
        return Vec::new();
    };
    RecordIter::new(&collection)
        .filter_map(Result::ok)
        .filter(|r| r.header.rec_type == RT_FONT_ENTITY_ATOM)
        .map(|r| {
            let name = r.data.get(..64).unwrap_or(&r.data);
            let name = decode_utf16le(name);
            name.split('\0').next().unwrap_or_default().to_string()
        })
        .collect()
}

/// Whether the deck carries a VBA project: a `VBAInfoAtom` with
/// `fHasMacros` set and a non-zero `persistIdRef` in the current
/// `DocumentContainer` ([MS-PPT] `VBAInfoAtom`). A `.ppt` keeps its project
/// as a compressed storage inside the "PowerPoint Document" stream, not as
/// a CFB root storage the way `.xls` (`_VBA_PROJECT_CUR`) and `.doc`
/// (`Macros`) do.
pub fn has_vba_project(stream: &[u8], current_user: Option<&[u8]>) -> bool {
    let doc_children = persist::build(stream, current_user)
        .and_then(|dir| dir.resolve(dir.doc_persist_id))
        .and_then(|offset| bounded_container_children(stream, offset, RT_DOCUMENT));
    let scope = doc_children.as_deref().unwrap_or(stream);
    find_record_any_instance(scope, RT_VBA_INFO_ATOM, 0).is_some_and(|atom| {
        atom.len() >= 8
            && u32::from_le_bytes([atom[0], atom[1], atom[2], atom[3]]) != 0
            && u32::from_le_bytes([atom[4], atom[5], atom[6], atom[7]]) == 1
    })
}

/// Bounded recursive search for the first record of `rec_type`, whatever
/// its instance, returning its body.
fn find_record_any_instance(data: &[u8], rec_type: u16, depth: usize) -> Option<Vec<u8>> {
    if depth > MAX_SHAPE_DEPTH {
        return None;
    }
    for rec in RecordIter::new(data) {
        let Ok(rec) = rec else { break };
        if rec.header.rec_type == rec_type {
            return Some(rec.data);
        }
        if rec.header.is_container() {
            if let Some(found) = find_record_any_instance(&rec.data, rec_type, depth + 1) {
                return Some(found);
            }
        }
    }
    None
}

/// The header/footer text a slides' `HeadersFootersContainer` actually
/// shows: the footer when `fHasFooter` is set, the user date when both
/// `fHasDate` and `fHasUserDate` are ([MS-PPT] `HeadersFootersAtom`).
/// The header string is the notes/handout page's and is never shown on a
/// slide; an automatic date has no stored text.
fn shown_slide_header_footer_texts(hf_children: &[u8]) -> Vec<String> {
    let mut flags = 0u16;
    let mut user_date = None;
    let mut footer = None;
    for rec in RecordIter::new(hf_children) {
        let Ok(rec) = rec else { break };
        match rec.header.rec_type {
            RT_HEADER_FOOTER_ATOM if rec.data.len() >= 4 => {
                flags = u16::from_le_bytes([rec.data[2], rec.data[3]]);
            },
            RT_CSTRING => {
                let text = decode_utf16le(&rec.data).trim().to_string();
                if text.is_empty() {
                    continue;
                }
                match rec.header.rec_instance {
                    HF_CSTRING_USER_DATE => user_date = Some(text),
                    HF_CSTRING_FOOTER => footer = Some(text),
                    _ => {},
                }
            },
            _ => {},
        }
    }
    let mut texts = Vec::new();
    if flags & HF_HAS_FOOTER != 0 {
        texts.extend(footer);
    }
    if flags & HF_HAS_DATE != 0 && flags & HF_HAS_USER_DATE != 0 {
        texts.extend(user_date);
    }
    texts
}

/// What resolving a slide found besides its text: its
/// `SlideAtom.notesIdRef`, and its own slide header/footer settings when it
/// overrides the deck's.
struct SlideLinks {
    notes_id_ref: Option<u32>,
    headers_footers: Option<Vec<String>>,
    /// `SlideAtom` `(masterIdRef, fMasterObjects)`.
    master: Option<(u32, bool)>,
}

/// One `SlidePersistAtom` entry of a `SlideListWithText`-family container
/// and the outline-text sequence that follows it ([MS-PPT] 2.4.14.3).
struct SlideListEntry {
    persist_id_ref: u32,
    /// `SlidePersistAtom.slideId` (body offset 12): the slide's `SlideId`,
    /// or for a notes list the notes page's `NotesId`.
    slide_id: u32,
    outline_texts: Vec<TextRun>,
}

/// Walk a `SlideListWithText`/`NotesListWithText` container's records into
/// one entry per `SlidePersistAtom`, each with its outline texts — the
/// `TextHeaderAtom`/`TextCharsAtom`/`TextBytesAtom` runs that directly
/// follow it, which an `OutlineTextRefAtom` indexes into.
fn slide_list_entries(list: &[u8]) -> Vec<SlideListEntry> {
    let mut entries: Vec<SlideListEntry> = Vec::new();
    let mut current_type = TextType::Other;
    let mut last_outline_idx: Option<usize> = None;
    for rec in RecordIter::new(list) {
        let Ok(rec) = rec else { break };
        match rec.header.rec_type {
            RT_SLIDE_PERSIST_ATOM if rec.data.len() >= 4 => {
                let persist_id_ref =
                    u32::from_le_bytes([rec.data[0], rec.data[1], rec.data[2], rec.data[3]]);
                let slide_id = rec
                    .data
                    .get(12..16)
                    .map(|b| u32::from_le_bytes([b[0], b[1], b[2], b[3]]))
                    .unwrap_or(0);
                entries.push(SlideListEntry {
                    persist_id_ref,
                    slide_id,
                    outline_texts: Vec::new(),
                });
                current_type = TextType::Other;
                last_outline_idx = None;
            },
            RT_TEXT_HEADER if rec.data.len() >= 4 => {
                let t = u32::from_le_bytes([rec.data[0], rec.data[1], rec.data[2], rec.data[3]]);
                current_type = TextType::from_u32(t);
                last_outline_idx = None;
            },
            RT_TEXT_CHARS | RT_TEXT_BYTES => {
                let Some(entry) = entries.last_mut() else {
                    continue;
                };
                let text = if rec.header.rec_type == RT_TEXT_CHARS {
                    decode_utf16le(&rec.data)
                } else {
                    decode_text_bytes(&rec.data)
                };
                // Positional index into this list is meaningful (it's what
                // OutlineTextRefAtom references) — an empty run still
                // occupies a slot and must not be skipped here.
                entry.outline_texts.push(TextRun {
                    text_type: current_type,
                    text,
                    hyperlink: None,
                    ..Default::default()
                });
                last_outline_idx = Some(entry.outline_texts.len() - 1);
            },
            RT_STYLE_TEXT_PROP => {
                if let (Some(entry), Some(idx)) = (entries.last_mut(), last_outline_idx) {
                    apply_style_text_prop(&mut entry.outline_texts[idx], &rec.data);
                }
            },
            _ => {},
        }
    }
    entries
}

/// Resolve one notes page: its `Notes` container's `Tx_TYPE_NOTES` text
/// (the notes body — not the page's slide image or date/number
/// placeholders), and the `NotesAtom.slideIdRef` of the slide it belongs
/// to. `None` when the container does not resolve.
fn resolve_notes(
    stream: &[u8],
    dir: &PersistDirectory,
    persist_id_ref: u32,
    outline_texts: &[TextRun],
    hyperlinks: &HashMap<u32, String>,
    ole_objects: &HashMap<u32, OleObjectInfo>,
) -> Option<(Vec<TextRun>, Option<u32>)> {
    let offset = dir.resolve(persist_id_ref)?;
    let children = bounded_container_children(stream, offset, RT_NOTES)?;
    let slide_id_ref = RecordIter::new(&children)
        .filter_map(Result::ok)
        .find(|r| r.header.rec_type == RT_NOTES_ATOM)
        .and_then(|r| {
            r.data
                .get(0..4)
                .map(|b| u32::from_le_bytes([b[0], b[1], b[2], b[3]]))
        });
    let mut runs = Vec::new();
    extract_shape_text(
        &children,
        0,
        outline_texts,
        hyperlinks,
        ole_objects,
        None,
        None,
        &mut runs,
        &mut Vec::new(),
        &mut Vec::new(),
        &mut Vec::new(),
    );
    runs.retain(|r| r.text_type == TextType::Notes && !r.text.trim().is_empty());
    Some((runs, slide_id_ref))
}

/// Resolve one slide's shape text: locate its `Slide` container via the
/// persist directory and walk its shape tree, resolving any
/// `OutlineTextRefAtom` references against `outline_texts`. Also returns
/// the slide's `SlideAtom.notesIdRef` (body offset 16) and its own
/// header/footer override, when present.
///
/// Every persist-directory-resolved slide is kept regardless of whether text
/// was found — an image-only slide is still a slide, and the presentation's
/// true slide count matters for numbering.
fn resolve_slide(
    stream: &[u8],
    dir: &PersistDirectory,
    master_persist_ids: &HashMap<u32, u32>,
    persist_id_ref: u32,
    outline_texts: &[TextRun],
    hyperlinks: &HashMap<u32, String>,
    ole_objects: &HashMap<u32, OleObjectInfo>,
) -> (SlideText, SlideLinks) {
    let mut text_runs = Vec::new();
    let mut tables = Vec::new();
    let mut image_refs = Vec::new();
    let mut ole_object_refs = Vec::new();
    let mut hidden = false;
    let mut notes_id_ref = None;
    let mut headers_footers = None;
    let mut master = None;
    if let Some(offset) = dir.resolve(persist_id_ref) {
        if let Some(children) = bounded_container_children(stream, offset, RT_SLIDE) {
            master = slide_master_ref(&children);
            headers_footers = find_child(&children, RT_HEADER_FOOTER, HF_INSTANCE_SLIDES)
                .map(|hf| shown_slide_header_footer_texts(&hf));
            notes_id_ref = RecordIter::new(&children)
                .filter_map(Result::ok)
                .find(|r| r.header.rec_type == RT_SLIDE_ATOM)
                .and_then(|r| {
                    r.data
                        .get(16..20)
                        .map(|b| u32::from_le_bytes([b[0], b[1], b[2], b[3]]))
                });
            extract_shape_text(
                &children,
                0,
                outline_texts,
                hyperlinks,
                ole_objects,
                None,
                None,
                &mut text_runs,
                &mut tables,
                &mut image_refs,
                &mut ole_object_refs,
            );
            hidden = slide_is_hidden(&children);

            // Placeholder character/paragraph formatting a direct
            // StyleTextPropAtom left unset falls back to the slide's
            // main master. Best-effort: any missing piece
            // of this chain (no SlideAtom, no master, no matching
            // TxMasterStyleAtom) just leaves direct formatting as-is.
            if let Some((master_id, _)) = master {
                if let Some(styles) =
                    resolve_master_styles(stream, dir, master_persist_ids, master_id)
                {
                    apply_master_inheritance(&mut text_runs, &styles);
                }
            }
        }
    }
    (
        SlideText {
            text_runs,
            tables,
            image_refs,
            ole_object_refs,
            hidden,
        },
        SlideLinks {
            notes_id_ref,
            headers_footers,
            master,
        },
    )
}

/// A main master's text styles: for each `TextTypeEnum` value (the
/// `TxMasterStyleAtom` `recInstance`, 0-8), its indent levels 0-4, each
/// level already carrying whatever it leaves unset from the level above
/// it ([MS-PPT] `TextMasterStyleAtom`: a level's exceptions override the
/// previous level's).
#[derive(Debug, Clone, Default)]
struct MasterStyles {
    levels: [Vec<(super::style::ParaFormat, super::style::CharFormat)>; 9],
}

impl MasterStyles {
    /// The levels a run of `text_type` inherits from. The placeholder
    /// variants fall back to their base style when the master has none of
    /// their own — `CenterTitle` to `Title`, `CenterBody`/`HalfBody`/
    /// `QuarterBody` to `Body` — as PowerPoint derives them.
    fn for_type(
        &self,
        text_type: TextType,
    ) -> Option<&[(super::style::ParaFormat, super::style::CharFormat)]> {
        let (own, base) = match text_type {
            TextType::Title => (0, None),
            TextType::Body => (1, None),
            TextType::Notes => (2, None),
            TextType::Other => (4, None),
            TextType::CenterBody => (5, Some(1)),
            TextType::CenterTitle => (6, Some(0)),
            TextType::HalfBody => (7, Some(1)),
            TextType::QuarterBody => (8, Some(1)),
        };
        [Some(own), base]
            .into_iter()
            .flatten()
            .map(|i| self.levels[i].as_slice())
            .find(|l| !l.is_empty())
    }
}

/// Resolve `master_id` (a `SlideAtom.masterIdRef`) to its main master via
/// [`resolve_master`] — through a title master to the main master it
/// names, since only main masters carry text styles — then parse every
/// `TxMasterStyleAtom` child: one per text type, up to five indent levels
/// each.
fn resolve_master_styles(
    stream: &[u8],
    dir: &PersistDirectory,
    master_persist_ids: &HashMap<u32, u32>,
    master_id: u32,
) -> Option<MasterStyles> {
    let mut master_id = master_id;
    let mut children = None;
    for _ in 0..MAX_MASTER_CHAIN {
        match resolve_master(stream, dir, master_persist_ids, master_id)? {
            MasterContainer::Main(c) => {
                children = Some(c);
                break;
            },
            MasterContainer::Title(c) => master_id = slide_master_ref(&c)?.0,
        }
    }
    let children = children?;
    let mut styles = MasterStyles::default();
    for rec in RecordIter::new(&children) {
        let Ok(rec) = rec else { break };
        if rec.header.rec_type != RT_TX_MASTER_STYLE_ATOM {
            continue;
        }
        let Some(slot) = styles.levels.get_mut(rec.header.rec_instance as usize) else {
            continue;
        };
        if slot.is_empty() {
            *slot = cascade_levels(style::parse_tx_master_style_atom(
                &rec.data,
                rec.header.rec_instance,
            ));
        }
    }
    if styles.levels.iter().all(Vec::is_empty) {
        return None;
    }
    Some(styles)
}

/// Make each master level carry what it leaves unset from the level above
/// it, so a paragraph at level `n` needs only level `n`.
fn cascade_levels(
    levels: Vec<(super::style::ParaFormat, super::style::CharFormat)>,
) -> Vec<(super::style::ParaFormat, super::style::CharFormat)> {
    let mut out: Vec<(super::style::ParaFormat, super::style::CharFormat)> =
        Vec::with_capacity(levels.len());
    for (pf, cf) in levels {
        let resolved = match out.last() {
            Some((ppf, pcf)) => (pf.inherit_from(ppf), cf.inherit_from(pcf)),
            None => (pf, cf),
        };
        out.push(resolved);
    }
    out
}

/// Fill any unset `CharFormat`/`ParaFormat` field on every `TextRun` from
/// the master style of its text type, at each paragraph's own indent
/// level (`TextPFRun.indentLevel`; level 0 where the run has no paragraph
/// formatting). A run with no direct formatting spans at all gets one
/// synthetic whole-text span carrying pure level-0 master formatting,
/// matching what PowerPoint itself renders.
fn apply_master_inheritance(text_runs: &mut [TextRun], styles: &MasterStyles) {
    for run in text_runs {
        let Some(levels) = styles.for_type(run.text_type) else {
            continue;
        };
        let level_at = |level: Option<u16>| {
            let i = (level.unwrap_or(0) as usize).min(levels.len() - 1);
            &levels[i]
        };
        let text_char_len = run.text.chars().count();
        // Character spans first, while the paragraph spans still carry
        // only the file's own indent levels: split each at the paragraph
        // boundaries it crosses, so each piece inherits from its own
        // paragraph's level.
        if run.char_formats.is_empty() {
            if text_char_len > 0 {
                run.char_formats.push(CharFormatSpan {
                    start: 0,
                    end: text_char_len,
                    format: level_at(None).1.clone(),
                });
            }
        } else {
            let mut split = Vec::with_capacity(run.char_formats.len());
            for span in &run.char_formats {
                let mut cursor = span.start;
                for p in run
                    .para_formats
                    .iter()
                    .filter(|p| p.end > span.start && p.start < span.end)
                {
                    let (a, b) = (p.start.max(span.start), p.end.min(span.end));
                    if a > cursor {
                        split.push(CharFormatSpan {
                            start: cursor,
                            end: a,
                            format: span.format.inherit_from(&level_at(None).1),
                        });
                    }
                    split.push(CharFormatSpan {
                        start: a,
                        end: b,
                        format: span.format.inherit_from(&level_at(p.format.indent_level).1),
                    });
                    cursor = b;
                }
                if cursor < span.end {
                    split.push(CharFormatSpan {
                        start: cursor,
                        end: span.end,
                        format: span.format.inherit_from(&level_at(None).1),
                    });
                }
            }
            run.char_formats = split;
        }
        if run.para_formats.is_empty() {
            if text_char_len > 0 {
                run.para_formats.push(ParaFormatSpan {
                    start: 0,
                    end: text_char_len,
                    format: level_at(None).0.clone(),
                });
            }
        } else {
            for span in &mut run.para_formats {
                let master = &level_at(span.format.indent_level).0;
                span.format = span.format.inherit_from(master);
            }
        }
    }
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
    ole_objects: &HashMap<u32, OleObjectInfo>,
    current_hyperlink: Option<&str>,
    current_placeholder_role: Option<&str>,
    out: &mut Vec<TextRun>,
    tables: &mut Vec<super::table::TableBlock>,
    image_refs: &mut Vec<usize>,
    ole_object_refs: &mut Vec<OleObjectInfo>,
) {
    if depth > MAX_SHAPE_DEPTH {
        return;
    }
    let mut current_type = TextType::Other;
    // Text-run-level hyperlink state: a `MouseClick/
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
                        placeholder_role: current_placeholder_role.map(str::to_string),
                        ..Default::default()
                    });
                    last_text_run_idx = Some(out.len() - 1);
                }
            },
            RT_TEXT_BYTES => {
                let text = decode_text_bytes(&rec.data);
                if !text.is_empty() {
                    out.push(TextRun {
                        text_type: current_type,
                        text,
                        hyperlink: current_hyperlink.map(str::to_string),
                        placeholder_role: current_placeholder_role.map(str::to_string),
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
                            run.placeholder_role = current_placeholder_role.map(str::to_string);
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
                            add_link_range(out, idx, begin, end, url);
                        }
                    }
                }
            },
            RT_SHAPE => {
                // Resolve this shape's own hyperlink (if any) before
                // walking its subtree, so every TextRun produced from it
                // — including nested containers like a group's own child
                // shapes, which are siblings under a group, not children
                // of THIS shape's own text — carries it.
                let shape_hyperlink = resolve_shape_hyperlink(&rec.data, hyperlinks);
                let hyperlink_ref = shape_hyperlink.as_deref().or(current_hyperlink);
                // This shape's own placeholder role, if it has one — same
                // "resolve at the shape actually carrying it, fall back to
                // whatever the enclosing group/shape already had" pattern
                // as the hyperlink just above.
                let shape_placeholder_role = resolve_shape_placeholder_role(&rec.data);
                let placeholder_role_ref = shape_placeholder_role
                    .as_deref()
                    .or(current_placeholder_role);
                // This shape's own picture reference, if it has one
                // — resolved here, at the shape actually
                // carrying it, not inferred from whichever slide
                // happens to be processed last.
                if let Some(idx) = resolve_shape_pib(&rec.data) {
                    image_refs.push(idx);
                }
                // This shape's own embedded/linked/ActiveX OLE object
                // identity, if it has one — the same "resolve at the
                // shape actually carrying it" reasoning as the picture
                // case just above.
                if let Some(info) = resolve_shape_ole_object(&rec.data, ole_objects) {
                    ole_object_refs.push(info);
                }
                extract_shape_text(
                    &rec.data,
                    depth + 1,
                    outline_texts,
                    hyperlinks,
                    ole_objects,
                    hyperlink_ref,
                    placeholder_role_ref,
                    out,
                    tables,
                    image_refs,
                    ole_object_refs,
                );
            },
            RT_SPGR_CONTAINER => {
                // A group whose members form a clean rectangular grid is
                // a reconstructed table; anything less
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
                        ole_objects,
                        current_hyperlink,
                        current_placeholder_role,
                        out,
                        tables,
                        image_refs,
                        ole_object_refs,
                    );
                }
            },
            _ if rec.header.is_container() => {
                extract_shape_text(
                    &rec.data,
                    depth + 1,
                    outline_texts,
                    hyperlinks,
                    ole_objects,
                    current_hyperlink,
                    current_placeholder_role,
                    out,
                    tables,
                    image_refs,
                    ole_object_refs,
                );
            },
            _ => {},
        }
    }
}

/// Try to recognize `spgr_data` (an `OfficeArtSpgrContainer`'s own
/// children) as a table: every `RT_SHAPE` after the group's own leading
/// placeholder shape must carry an `RT_CHILD_ANCHOR`, and the resulting
/// positions must form a clean rectangular grid. Returns
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
        let left = i32::from_le_bytes([
            anchor.data[0],
            anchor.data[1],
            anchor.data[2],
            anchor.data[3],
        ]);
        let top = i32::from_le_bytes([
            anchor.data[4],
            anchor.data[5],
            anchor.data[6],
            anchor.data[7],
        ]);

        let shape_hyperlink = resolve_shape_hyperlink(&rec.data, hyperlinks);
        let hyperlink_ref = shape_hyperlink.as_deref().or(current_hyperlink);
        let mut runs = Vec::new();
        extract_shape_text(
            &rec.data,
            depth + 1,
            outline_texts,
            hyperlinks,
            &HashMap::new(),
            hyperlink_ref,
            None,
            &mut runs,
            &mut Vec::new(),
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
                current = Some(SlideText {
                    text_runs: Vec::new(),
                    ..Default::default()
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
                    let text = decode_text_bytes(&rec.data);
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
/// [`CSTRING_INSTANCE_TARGET`]).
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
/// (`exHyperlinkId`), its `TargetAtom` (the URL/path) and its
/// `LocationAtom`. A hyperlink with only a location jumps inside this
/// deck; see [`internal_location_target`].
fn parse_one_ex_hyperlink(data: &[u8]) -> Option<(u32, String)> {
    let mut id = None;
    let mut target = None;
    let mut location = None;
    for rec in RecordIter::new(data) {
        let Ok(rec) = rec else { break };
        match rec.header.rec_type {
            RT_EXTERNAL_HYPERLINK_ATOM if rec.data.len() >= 4 => {
                id = Some(u32::from_le_bytes([rec.data[0], rec.data[1], rec.data[2], rec.data[3]]));
            },
            RT_CSTRING => {
                let s = decode_utf16le(&rec.data);
                let s = s.trim_end_matches('\0');
                if s.is_empty() {
                    continue;
                }
                match rec.header.rec_instance {
                    CSTRING_INSTANCE_TARGET => target = Some(s.to_string()),
                    CSTRING_INSTANCE_LOCATION => location = Some(s.to_string()),
                    _ => {},
                }
            },
            _ => {},
        }
    }
    let url = match (target, location) {
        (Some(t), _) => t,
        (None, Some(loc)) => internal_location_target(&loc),
        (None, None) => return None,
    };
    Some((id?, url))
}

/// The IR hyperlink for a deck-internal `LocationAtom`. A slide jump is
/// written `"<slideId>,<slide number>,<title>"` (as Apache POI reads it);
/// it becomes `#slide<N>.xml`, the same target the PPTX converter gives a
/// slide jump. Any other location is kept as a `#` fragment.
fn internal_location_target(location: &str) -> String {
    let mut parts = location.splitn(3, ',');
    if let (Some(id), Some(number)) = (parts.next(), parts.next()) {
        if id.trim().parse::<u32>().is_ok() {
            if let Ok(n) = number.trim().parse::<u32>() {
                return format!("#slide{n}.xml");
            }
        }
    }
    format!("#{location}")
}

/// One external object's identity: an embedded/linked/ActiveX OLE object
/// resolved from its `ExOleObjAtom`, or a video/sound object resolved from
/// its `ExMediaAtom` (both share the `ExObjListContainer`'s `exObjId`
/// space, which a shape's `ExObjRefAtom` names).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct OleObjectInfo {
    /// `ExOleObjAtom.subType` — 0-15; see `describe_ole_subtype` in
    /// `convert_ppt.rs` for the enum meaning. 0 for media.
    pub subtype: u32,
    /// `ExOleObjAtom.type` — 0 = embedded, 1 = linked, 2 = ActiveX control
    /// — or [`KIND_MEDIA_VIDEO`](Self::KIND_MEDIA_VIDEO) /
    /// [`KIND_MEDIA_AUDIO`](Self::KIND_MEDIA_AUDIO) for a media object,
    /// which has no `ExOleObjAtom`.
    pub kind: u32,
}

impl OleObjectInfo {
    /// `kind` of a video object (`ExAviMovieContainer`/`ExMCIMovieContainer`).
    pub const KIND_MEDIA_VIDEO: u32 = 0x100;
    /// `kind` of a sound object (MIDI, CD audio, embedded or linked WAV).
    pub const KIND_MEDIA_AUDIO: u32 = 0x101;
}

/// Build the document-wide `objID -> OleObjectInfo` table from the
/// `ExObjListContainer`'s `ExOleObjAtom` records — the same
/// container the hyperlink resolver already partially parses, walked again
/// here for its `ExEmbed`/`ExOleObjAtom` children instead.
fn parse_ex_ole_objects(stream: &[u8]) -> HashMap<u32, OleObjectInfo> {
    let mut out = HashMap::new();
    let Some(ex_obj_list) = find_descendant(stream, RT_EXTERNAL_OBJECT_LIST, 0, 0) else {
        return out;
    };
    collect_ex_ole_objects(&ex_obj_list, 0, &mut out);
    out
}

fn collect_ex_ole_objects(data: &[u8], depth: usize, out: &mut HashMap<u32, OleObjectInfo>) {
    if depth > MAX_SHAPE_DEPTH {
        return;
    }
    for rec in RecordIter::new(data) {
        let Ok(rec) = rec else { break };
        if rec.header.rec_type == RT_EXTERNAL_OLE_OBJECT_ATOM {
            if let Some((id, info)) = parse_one_ex_ole_obj_atom(&rec.data) {
                out.insert(id, info);
            }
            continue; // an ExOleObjAtom's own siblings are never other ExOleObjAtoms
        }
        // A video or sound object: identified by the `exObjId` of the
        // `ExMediaAtom` inside it, and — unlike OLE objects — it left no
        // trace at all before, so a slide's movie or sound vanished.
        let media_kind = match rec.header.rec_type {
            RT_EX_AVI_MOVIE | RT_EX_MCI_MOVIE => Some(OleObjectInfo::KIND_MEDIA_VIDEO),
            RT_EX_MIDI_AUDIO | RT_EX_CD_AUDIO | RT_EX_WAV_AUDIO_EMBEDDED | RT_EX_WAV_AUDIO_LINK => {
                Some(OleObjectInfo::KIND_MEDIA_AUDIO)
            },
            _ => None,
        };
        if let Some(kind) = media_kind {
            let id =
                find_record_any_instance(&rec.data, RT_EX_MEDIA_ATOM, depth + 1).and_then(|b| {
                    b.get(0..4)
                        .map(|b| u32::from_le_bytes([b[0], b[1], b[2], b[3]]))
                });
            if let Some(id) = id {
                out.entry(id).or_insert(OleObjectInfo { subtype: 0, kind });
            }
            continue;
        }
        if rec.header.is_container() {
            collect_ex_ole_objects(&rec.data, depth + 1, out);
        }
    }
}

/// Parse one `ExOleObjAtom`'s fixed body: `drawAspect`(4) + `type`(4) +
/// `objID`(4) + `subType`(4) + ... (`objStgDataRef`/`options` follow but
/// aren't needed for identity), per [MS-PPT] 2.10.20 — byte layout
/// cross-checked against Apache POI's `ExOleObjAtom`.
fn parse_one_ex_ole_obj_atom(data: &[u8]) -> Option<(u32, OleObjectInfo)> {
    if data.len() < 16 {
        return None;
    }
    let kind = u32::from_le_bytes([data[4], data[5], data[6], data[7]]);
    let obj_id = u32::from_le_bytes([data[8], data[9], data[10], data[11]]);
    let subtype = u32::from_le_bytes([data[12], data[13], data[14], data[15]]);
    Some((obj_id, OleObjectInfo { subtype, kind }))
}

/// Resolve a shape's own OLE object reference (if any): its
/// `RT_CLIENT_DATA`'s direct `ExObjRefAtom` child, joined against the
/// document-wide `ole_objects` table by `objID`.
fn resolve_shape_ole_object(
    shape_data: &[u8],
    ole_objects: &HashMap<u32, OleObjectInfo>,
) -> Option<OleObjectInfo> {
    let client_data = find_descendant(shape_data, RT_CLIENT_DATA, 0, 0)?;
    let atom = find_descendant(&client_data, RT_EXTERNAL_OBJECT_REF_ATOM, 0, 0)?;
    if atom.len() < 4 {
        return None;
    }
    let obj_id_ref = u32::from_le_bytes([atom[0], atom[1], atom[2], atom[3]]);
    ole_objects.get(&obj_id_ref).copied()
}

/// Resolve a shape's own placeholder role (if any): its `RT_CLIENT_DATA`'s
/// direct `OEPlaceholderAtom` child's `placeholderId` byte, mapped to the
/// OOXML `ST_PlaceholderType` string vocabulary.
///
/// Scoped to the "regular presentation slide" context only, per Apache
/// POI's `org.apache.poi.sl.usermodel.Placeholder` enum (the authoritative
/// cross-reference this mapping was verified against): the *same* raw
/// `placeholderId` byte means a different role depending on whether the
/// shape lives on a slide, a slide master, a notes slide, or a notes
/// master (e.g. raw `13` is Title on a slide, but raw `1` is Title on a
/// slide master). This function's caller (`collect_slide_containers`) only
/// walks shapes already confirmed to be inside a real `RT_SLIDE`
/// container, so the slide-context mapping is the correct one there;
/// applying this same mapping to a master/notes shape would misattribute
/// its role. Master/notes placeholder roles are not resolved by this
/// pass — a real but bounded simplification of the full 4-context table,
/// consistent with this crate's established "at minimum" completeness bar
/// for this kind of gap.
fn resolve_shape_placeholder_role(shape_data: &[u8]) -> Option<String> {
    let client_data = find_descendant(shape_data, RT_CLIENT_DATA, 0, 0)?;
    let atom = find_descendant(&client_data, RT_OE_PLACEHOLDER_ATOM, 0, 0)?;
    // placementId(4) + placeholderId(1) + ...
    let placeholder_id = *atom.get(4)?;
    placeholder_id_to_ooxml_type(placeholder_id).map(str::to_string)
}

/// Map a `.ppt` slide-context `OEPlaceholderAtom.placeholderId` byte to the
/// OOXML `ST_PlaceholderType` string it corresponds to, verified against
/// Apache POI's `Placeholder` enum's `nativeSlideId`/`ooxmlId` columns.
/// Three roles (`VERTICAL_OBJECT`/`VERTICAL_TEXT_TITLE`/`VERTICAL_TEXT_BODY`,
/// raw `17`/`18`/`25`) have no OOXML equivalent at all (POI itself records
/// their `ooxmlId` as `-2`, "no mapping") and are deliberately left
/// unmapped here rather than inventing a non-standard string.
fn placeholder_id_to_ooxml_type(id: u8) -> Option<&'static str> {
    match id {
        7 => Some("dt"),        // Date
        8 => Some("sldNum"),    // Slide number
        9 => Some("ftr"),       // Footer
        10 => Some("hdr"),      // Header
        11 => Some("sldImg"),   // Slide image
        13 => Some("title"),    // Title
        14 => Some("body"),     // Body
        15 => Some("ctrTitle"), // Centered title
        16 => Some("subTitle"), // Subtitle
        19 => Some("obj"),      // Object
        20 => Some("chart"),    // Chart
        21 => Some("tbl"),      // Table
        22 => Some("clipArt"),  // Clip art
        23 => Some("dgm"),      // Diagram / org chart
        24 => Some("media"),    // Media
        26 => Some("pic"),      // Picture
        _ => None,
    }
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
/// `0x00000000` means "no picture" per [MS-ODRAW].
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

/// Record a text-range hyperlink ([MS-PPT] `TextInteractiveInfoAtom`) on
/// `out[idx]`. `begin`/`end` are `TextPosition` offsets, which count UTF-16
/// code units (the text is stored as UTF-16); they are converted to `char`
/// indices, so an astral-plane character (a surrogate pair) before the
/// range does not shift it.
///
/// The run is not split: splitting it into prefix/link/suffix runs made
/// each piece its own paragraph downstream, breaking the sentence around
/// the link into three lines and dropping the spaces at the joins.
fn add_link_range(out: &mut [TextRun], idx: usize, begin: usize, end: usize, url: &str) {
    let Some(run) = out.get_mut(idx) else { return };
    let len = run.text.chars().count();
    let begin = utf16_to_char_index(&run.text, begin).min(len);
    let end = utf16_to_char_index(&run.text, end).clamp(begin, len);
    if begin >= end {
        return; // empty or invalid range — leave the run untouched
    }
    if begin == 0 && end == len {
        run.hyperlink = Some(url.to_string()); // the whole run is the link
        return;
    }
    run.link_ranges.push(LinkRange {
        start: begin,
        end,
        url: url.to_string(),
    });
}

/// Parse `data` as a `StyleTextPropAtom` body and attach the resulting
/// character-/paragraph-formatting spans to `run`, clamped against
/// `run.text`'s own length.
///
/// The runs' `count`s are in UTF-16 code units, the unit the text is
/// stored in; the spans are parsed in those units and then converted to
/// `char` indices, which is what every consumer slices by.
fn apply_style_text_prop(run: &mut TextRun, data: &[u8]) {
    let text_utf16_len = run.text.encode_utf16().count();
    let (mut para_spans, mut char_spans) = style::parse_style_text_prop(data, text_utf16_len);
    if text_utf16_len != run.text.chars().count() {
        for s in &mut para_spans {
            s.start = utf16_to_char_index(&run.text, s.start);
            s.end = utf16_to_char_index(&run.text, s.end);
        }
        for s in &mut char_spans {
            s.start = utf16_to_char_index(&run.text, s.start);
            s.end = utf16_to_char_index(&run.text, s.end);
        }
    }
    run.para_formats = para_spans;
    run.char_formats = char_spans;
}

/// The `char` index of UTF-16 code unit offset `units` in `text` (an offset
/// inside a surrogate pair rounds up to the next character; past the end
/// clamps to the length).
fn utf16_to_char_index(text: &str, units: usize) -> usize {
    let mut seen = 0usize;
    for (i, c) in text.chars().enumerate() {
        if seen >= units {
            return i;
        }
        seen += c.len_utf16();
    }
    text.chars().count()
}

/// Text content of a single slide.
#[derive(Debug, Clone, Default)]
pub struct SlideText {
    /// All text runs belonging to this slide.
    pub text_runs: Vec<TextRun>,
    /// Shape groups recognized as tables. Rendered after
    /// `text_runs` in the IR — the binary format has no single unified
    /// reading-order concept to interleave them with, so this is a
    /// deliberate simplification, not a claim of true document order.
    pub tables: Vec<super::table::TableBlock>,
    /// 0-based indices into the document's `Pictures`-stream-derived
    /// image list ([`super::images::PptImage::index`]) for every
    /// picture shape resolved on this slide, in shape-tree encounter
    /// order (these used to be silently dumped onto
    /// whichever slide happened to be last, regardless of which slide
    /// actually contains the shape referencing them).
    pub image_refs: Vec<usize>,
    /// Every embedded/linked/ActiveX OLE object resolved on this slide,
    /// in shape-tree encounter order — a slide with an embedded Excel
    /// workbook, Word document, Equation Editor object, etc. previously
    /// surfaced nothing at all indicating the object even existed.
    pub ole_object_refs: Vec<OleObjectInfo>,
    /// Whether the slide is marked hidden (not shown during a slide
    /// show) via a `SlideShowSlideInfoAtom` HIDDEN_BIT sibling of the
    /// `Slide` container. The content is still extracted — a consumer
    /// indexing a deck usually wants it — but a caller can now tell the
    /// author didn't intend it to be seen (the `.ppt`
    /// analogue of the XLSX/PPTX hidden flags).
    pub hidden: bool,
}

/// Look for a `SlideShowSlideInfoAtom` among a `Slide` container's direct
/// children and read its `HIDDEN_BIT` (`0x0004`) out of the
/// `effectTransitionFlags` field.
///
/// Layout ([MS-PPT] 2.13.24): after the 8-byte record header, `slideTime:
/// i32`, `soundIdRef: i32`, `effectDirection: u8`, `effectType: u8`,
/// `effectTransitionFlags: u16` (offset 10 within the atom body), `speed:
/// u8`, 3 unused bytes — 16 bytes total, 24 with the header. Verified
/// against Apache POI's `SSSlideInfoAtom`.
fn slide_is_hidden(children: &[u8]) -> bool {
    for rec in RecordIter::new(children) {
        let Ok(rec) = rec else { break };
        if rec.header.rec_type == RT_SLIDE_SHOW_SLIDE_INFO_ATOM {
            if let Some(flags_bytes) = rec.data.get(10..12) {
                let flags = u16::from_le_bytes([flags_bytes[0], flags_bytes[1]]);
                return flags & 0x0004 != 0;
            }
        }
    }
    false
}

/// Decode a `TextBytesAtom` body ([MS-PPT] `TextBytesAtom`): "each byte is the low
/// byte of a UTF-16 character whose high byte is 0x00" — i.e. Latin-1,
/// whatever the deck's language; PowerPoint writes a `TextCharsAtom`
/// whenever a character needs more.
fn decode_text_bytes(data: &[u8]) -> String {
    data.iter().map(|&b| b as char).collect()
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
    fn test_extract_text_chars() {
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
            &HashMap::new(),
            None,
            None,
            &mut runs,
            &mut Vec::new(),
            &mut Vec::new(),
            &mut Vec::new(),
        );
        assert_eq!(runs.len(), 1);
        assert_eq!(runs[0].text, "Hi");
        assert_eq!(runs[0].text_type, TextType::Title);
    }

    #[test]
    fn test_extract_text_bytes() {
        let mut stream = make_atom(RT_TEXT_HEADER, 0, &1u32.to_le_bytes()); // Body
        stream.extend(make_atom(RT_TEXT_BYTES, 0, b"Hello World"));
        let mut runs = Vec::new();
        extract_shape_text(
            &stream,
            0,
            &[],
            &HashMap::new(),
            &HashMap::new(),
            None,
            None,
            &mut runs,
            &mut Vec::new(),
            &mut Vec::new(),
            &mut Vec::new(),
        );
        assert_eq!(runs.len(), 1);
        assert_eq!(runs[0].text, "Hello World");
        assert_eq!(runs[0].text_type, TextType::Body);
    }

    #[test]
    fn test_extract_multiple_runs() {
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
            &HashMap::new(),
            None,
            None,
            &mut runs,
            &mut Vec::new(),
            &mut Vec::new(),
            &mut Vec::new(),
        );
        assert_eq!(runs.len(), 2);
        assert_eq!(runs[0].text, "Title");
        assert_eq!(runs[1].text, "Body text");
    }

    /// Build one `ExHyperlinkContainer`: `ExHyperlinkAtom` (id) +
    /// `TargetAtom` (a `RT_CSTRING` at `CSTRING_INSTANCE_TARGET`, UTF-16LE).
    /// Slide-jump hyperlinks carry their target in a `LocationAtom`
    /// (`CString` instance 3) with no `TargetAtom`; only the `TargetAtom`
    /// was read, so every internal jump resolved to nothing.
    #[test]
    fn test_location_atom_hyperlink_resolves_to_an_internal_target() {
        let mut children = make_atom(RT_EXTERNAL_HYPERLINK_ATOM, 0, &7u32.to_le_bytes());
        let loc: Vec<u8> = "258,3,Results"
            .encode_utf16()
            .flat_map(|u| u.to_le_bytes())
            .collect();
        children.extend(make_atom(RT_CSTRING, CSTRING_INSTANCE_LOCATION, &loc));
        assert_eq!(parse_one_ex_hyperlink(&children), Some((7, "#slide3.xml".to_string())));
        assert_eq!(internal_location_target("Bookmark1"), "#Bookmark1");
        // A target wins over a location (the location is inside it).
        let with_target = {
            let mut c = children.clone();
            let t: Vec<u8> = "http://example.com/"
                .encode_utf16()
                .flat_map(|u| u.to_le_bytes())
                .collect();
            c.extend(make_atom(RT_CSTRING, CSTRING_INSTANCE_TARGET, &t));
            c
        };
        assert_eq!(
            parse_one_ex_hyperlink(&with_target).map(|(_, u)| u),
            Some("http://example.com/".to_string())
        );
    }

    /// Text-range hyperlink offsets count UTF-16 code units; slicing them
    /// as `char` indices put the link one character late for every emoji
    /// (surrogate pair) before it.
    #[test]
    fn test_hyperlink_range_is_measured_in_utf16_units() {
        let text = "\u{1F600} link here";
        let mut out = vec![TextRun {
            text: text.to_string(),
            ..Default::default()
        }];
        // Emoji = 2 units, space = 1: "link" is units 3..7, chars 2..6.
        add_link_range(&mut out, 0, 3, 7, "http://example.com/");
        assert_eq!((out[0].link_ranges[0].start, out[0].link_ranges[0].end), (2, 6));
        assert_eq!(utf16_to_char_index(text, 2), 1);
        assert_eq!(utf16_to_char_index(text, 1), 1, "inside the pair rounds up");
        assert_eq!(utf16_to_char_index(text, 99), text.chars().count());
    }

    /// `StyleTextPropAtom` run counts are UTF-16 units too: formatting on
    /// the word after an emoji must cover exactly that word.
    #[test]
    fn test_style_runs_are_measured_in_utf16_units() {
        let text = "\u{1F600} bold";
        let mut run = TextRun {
            text: text.to_string(),
            ..Default::default()
        };
        // One paragraph run over all 8 units (+1 mark), then character runs:
        // 3 units plain, 4 units bold (CF_BOLD, fontStyle bold), 1 plain.
        let mut data = Vec::new();
        data.extend_from_slice(&8u32.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        for (count, bold) in [(3u32, false), (4, true), (2, false)] {
            data.extend_from_slice(&count.to_le_bytes());
            data.extend_from_slice(&1u32.to_le_bytes()); // CFMasks: bold
            data.extend_from_slice(&(bold as u16).to_le_bytes()); // fontStyle
        }
        apply_style_text_prop(&mut run, &data);
        let bold: Vec<(usize, usize)> = run
            .char_formats
            .iter()
            .filter(|s| s.format.bold == Some(true))
            .map(|s| (s.start, s.end))
            .collect();
        // chars: emoji(0) space(1) b(2) o(3) l(4) d(5)
        assert_eq!(bold, [(2, 6)]);
    }

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

    // ── picture-shape `pib` resolution ──

    /// Build an `OfficeArtFOPT` record with a single
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
    fn test_resolve_shape_pib_finds_the_pib_property() {
        let shape_data = make_fopt_with_pib(3); // one-based
        assert_eq!(resolve_shape_pib(&shape_data), Some(2)); // zero-based
    }

    #[test]
    fn test_resolve_shape_pib_zero_means_no_picture() {
        let shape_data = make_fopt_with_pib(0);
        assert_eq!(resolve_shape_pib(&shape_data), None);
    }

    #[test]
    fn test_resolve_shape_pib_none_without_fopt() {
        assert_eq!(resolve_shape_pib(&[]), None);
    }

    #[test]
    fn test_resolve_shape_pib_skips_unrelated_properties() {
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
    fn test_shape_with_pib_reaches_image_refs_end_to_end() {
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
            &HashMap::new(),
            None,
            None,
            &mut runs,
            &mut tables,
            &mut image_refs,
            &mut Vec::new(),
        );

        assert_eq!(image_refs, vec![0]);
        assert_eq!(runs.len(), 1, "the shape's own text must still be extracted alongside its pib");
    }

    /// Build one `ExOleObjAtom`'s 24-byte fixed body per [MS-PPT] 2.10.20:
    /// `drawAspect`(4) + `type`(4) + `objID`(4) + `subType`(4) +
    /// `objStgDataRef`(4) + `options`(4).
    fn make_ex_ole_obj_atom(obj_id: u32, subtype: u32, kind: u32) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&1u32.to_le_bytes()); // drawAspect = VISIBLE
        data.extend_from_slice(&kind.to_le_bytes());
        data.extend_from_slice(&obj_id.to_le_bytes());
        data.extend_from_slice(&subtype.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes()); // objStgDataRef
        data.extend_from_slice(&0u32.to_le_bytes()); // options/isBlank
        make_atom(RT_EXTERNAL_OLE_OBJECT_ATOM, 0, &data)
    }

    /// A shape whose `RT_CLIENT_DATA` carries an
    /// `ExObjRefAtom` (`exObjIdRef`), joined against the document's
    /// `ExObjListContainer` -> `ExEmbed` -> `ExOleObjAtom` chain by
    /// `objID`.
    #[test]
    fn test_parse_ex_ole_objects_resolves_id_to_subtype_and_kind() {
        let ole_atom = make_ex_ole_obj_atom(1, 3, 0); // objID=1, Excel, embedded
        let ex_embed = make_container(RT_EXTERNAL_OLE_EMBED, 0, &ole_atom);
        let ex_obj_list = make_container(RT_EXTERNAL_OBJECT_LIST, 0, &ex_embed);

        let map = parse_ex_ole_objects(&ex_obj_list);
        let info = map.get(&1).expect("objID 1 must resolve");
        assert_eq!(info.subtype, 3, "Excel subtype");
        assert_eq!(info.kind, 0, "embedded, not linked");
    }

    /// End-to-end: a shape referencing an OLE object via `ExObjRefAtom`
    /// inside its `RT_CLIENT_DATA` resolves through to `ole_object_refs`,
    /// alongside its own ordinary text (mirrors
    /// `test_shape_with_pib_reaches_image_refs_end_to_end` for the OLE case).
    #[test]
    fn test_shape_with_ex_obj_ref_reaches_ole_object_refs_end_to_end() {
        let mut ole_objects = HashMap::new();
        ole_objects.insert(
            1u32,
            super::OleObjectInfo {
                subtype: 3,
                kind: 0,
            },
        );

        let ex_obj_ref_atom = make_atom(RT_EXTERNAL_OBJECT_REF_ATOM, 0, &1u32.to_le_bytes());
        let client_data = make_container(RT_CLIENT_DATA, 0, &ex_obj_ref_atom);
        let mut shape_children = client_data;
        let mut textbox_children = make_atom(RT_TEXT_HEADER, 0, &4u32.to_le_bytes());
        textbox_children.extend(make_atom(RT_TEXT_BYTES, 0, b"caption"));
        shape_children.extend(make_container(0xF00D, 0, &textbox_children));
        let shape = make_container(RT_SHAPE, 0, &shape_children);

        let mut runs = Vec::new();
        let mut ole_object_refs = Vec::new();
        extract_shape_text(
            &shape,
            0,
            &[],
            &HashMap::new(),
            &ole_objects,
            None,
            None,
            &mut runs,
            &mut Vec::new(),
            &mut Vec::new(),
            &mut ole_object_refs,
        );

        assert_eq!(ole_object_refs.len(), 1);
        assert_eq!(ole_object_refs[0].subtype, 3);
        assert_eq!(
            runs.len(),
            1,
            "the shape's own text must still be extracted alongside its OLE ref"
        );
    }

    /// Video and sound objects ([MS-PPT] `ExAviMovieContainer`,
    /// `ExWAVAudioEmbeddedContainer`, …) are named by their `ExMediaAtom`'s
    /// `exObjId`; a shape's `ExObjRefAtom` referencing one resolved to
    /// nothing, so a slide's movie or sound left no trace.
    #[test]
    fn test_media_objects_resolve_like_ole_objects() {
        let media =
            |id: u32| make_atom(RT_EX_MEDIA_ATOM, 0, &[&id.to_le_bytes()[..], &[0u8; 4]].concat());
        let video = make_container(
            RT_EX_AVI_MOVIE,
            0,
            &make_container(0x1005, 0, &media(4)), // ExVideoContainer
        );
        let sound = make_container(RT_EX_WAV_AUDIO_EMBEDDED, 0, &media(5));
        let ex_obj_list = make_container(RT_EXTERNAL_OBJECT_LIST, 0, &[video, sound].concat());
        let map = parse_ex_ole_objects(&ex_obj_list);
        assert_eq!(map.get(&4).map(|i| i.kind), Some(OleObjectInfo::KIND_MEDIA_VIDEO));
        assert_eq!(map.get(&5).map(|i| i.kind), Some(OleObjectInfo::KIND_MEDIA_AUDIO));
    }

    /// A shape with no `ExObjRefAtom` at all must not resolve anything,
    /// even when the document has real OLE objects elsewhere.
    #[test]
    fn test_resolve_shape_ole_object_none_without_ex_obj_ref() {
        let mut ole_objects = HashMap::new();
        ole_objects.insert(
            1u32,
            super::OleObjectInfo {
                subtype: 3,
                kind: 0,
            },
        );

        let client_data = make_container(RT_CLIENT_DATA, 0, &[]);
        let shape_children = client_data;

        assert!(resolve_shape_ole_object(&shape_children, &ole_objects).is_none());
    }

    /// Build an `OEPlaceholderAtom` body: `placementId(4)` +
    /// `placeholderId(1)` + `placeholderSize(1)` + `unusedShort(2)`.
    fn make_oe_placeholder_atom_data(placeholder_id: u8) -> Vec<u8> {
        let mut data = vec![0u8; 8];
        data[4] = placeholder_id;
        data
    }

    #[test]
    fn test_placeholder_id_to_ooxml_type_maps_known_slide_context_values() {
        assert_eq!(placeholder_id_to_ooxml_type(13), Some("title"));
        assert_eq!(placeholder_id_to_ooxml_type(14), Some("body"));
        assert_eq!(placeholder_id_to_ooxml_type(15), Some("ctrTitle"));
        assert_eq!(placeholder_id_to_ooxml_type(16), Some("subTitle"));
        assert_eq!(placeholder_id_to_ooxml_type(7), Some("dt"));
        assert_eq!(placeholder_id_to_ooxml_type(8), Some("sldNum"));
        assert_eq!(placeholder_id_to_ooxml_type(9), Some("ftr"));
        assert_eq!(placeholder_id_to_ooxml_type(10), Some("hdr"));
        assert_eq!(placeholder_id_to_ooxml_type(22), Some("clipArt"));
    }

    /// Raw values with no OOXML equivalent (vertical title/body/object —
    /// POI's own table records their `ooxmlId` as -2) must not be
    /// invented a non-standard string.
    #[test]
    fn test_placeholder_id_to_ooxml_type_leaves_vertical_roles_unmapped() {
        assert_eq!(placeholder_id_to_ooxml_type(17), None);
        assert_eq!(placeholder_id_to_ooxml_type(18), None);
        assert_eq!(placeholder_id_to_ooxml_type(25), None);
        assert_eq!(placeholder_id_to_ooxml_type(0), None);
        assert_eq!(placeholder_id_to_ooxml_type(255), None);
    }

    #[test]
    fn test_resolve_shape_placeholder_role_reads_the_atom() {
        let atom = make_atom(RT_OE_PLACEHOLDER_ATOM, 0, &make_oe_placeholder_atom_data(9)); // Footer
        let client_data = make_container(RT_CLIENT_DATA, 0, &atom);
        assert_eq!(resolve_shape_placeholder_role(&client_data).as_deref(), Some("ftr"));
    }

    #[test]
    fn test_resolve_shape_placeholder_role_none_without_the_atom() {
        let client_data = make_container(RT_CLIENT_DATA, 0, &[]);
        assert!(resolve_shape_placeholder_role(&client_data).is_none());
    }

    /// End-to-end: a shape carrying an `OEPlaceholderAtom`
    /// for "Footer" produces a `TextRun` whose `placeholder_role` is
    /// `Some("ftr")`.
    #[test]
    fn test_shape_with_oe_placeholder_atom_reaches_placeholder_role_end_to_end() {
        let mut text_header = make_atom(RT_TEXT_HEADER, 0, &4u32.to_le_bytes()); // TextType::Other
        text_header.extend(make_atom(RT_TEXT_CHARS, 0, &[0x46, 0x00, 0x74, 0x00])); // "Ft" UTF-16LE
        let placeholder_atom =
            make_atom(RT_OE_PLACEHOLDER_ATOM, 0, &make_oe_placeholder_atom_data(9)); // Footer
        let client_data = make_container(RT_CLIENT_DATA, 0, &placeholder_atom);
        let mut shape_data = text_header;
        shape_data.extend(client_data);
        let shape = make_container(RT_SHAPE, 0, &shape_data);

        let mut runs = Vec::new();
        extract_shape_text(
            &shape,
            0,
            &[],
            &HashMap::new(),
            &HashMap::new(),
            None,
            None,
            &mut runs,
            &mut Vec::new(),
            &mut Vec::new(),
            &mut Vec::new(),
        );
        assert_eq!(runs.len(), 1);
        assert_eq!(runs[0].placeholder_role.as_deref(), Some("ftr"));
    }

    #[test]
    fn test_parse_ex_hyperlinks_resolves_id_to_target_url() {
        let ex_obj_list =
            make_container(RT_EXTERNAL_OBJECT_LIST, 0, &make_ex_hyperlink(1, "http://example.com"));
        let map = parse_ex_hyperlinks(&ex_obj_list);
        assert_eq!(map.get(&1).map(String::as_str), Some("http://example.com"));
    }

    #[test]
    fn test_parse_ex_hyperlinks_multiple_entries() {
        let mut ex_obj_list_children = make_ex_hyperlink(1, "http://a.example/");
        ex_obj_list_children.extend(make_ex_hyperlink(2, "http://b.example/"));
        let ex_obj_list = make_container(RT_EXTERNAL_OBJECT_LIST, 0, &ex_obj_list_children);
        let map = parse_ex_hyperlinks(&ex_obj_list);
        assert_eq!(map.len(), 2);
        assert_eq!(map.get(&1).map(String::as_str), Some("http://a.example/"));
        assert_eq!(map.get(&2).map(String::as_str), Some("http://b.example/"));
    }

    #[test]
    fn test_empty_stream_yields_no_hyperlinks() {
        assert!(parse_ex_hyperlinks(&[]).is_empty());
    }

    /// The real end-to-end chain: a shape's own
    /// `InteractiveInfoAtom` (`exHyperlinkIdRef` + `II_HyperlinkAction`)
    /// resolves through the document's `ExObjList` to a real URL, which
    /// ends up on the shape's own `TextRun::hyperlink`.
    #[test]
    fn test_shape_with_interactive_info_resolves_its_hyperlink() {
        let mut hyperlinks = HashMap::new();
        hyperlinks.insert(1u32, "http://testuri.org/".to_string());

        let shape = make_hyperlinked_shape(1, b"Click here", 1); // Body, exHyperlinkIdRef=1
        let mut runs = Vec::new();
        extract_shape_text(
            &shape,
            0,
            &[],
            &hyperlinks,
            &HashMap::new(),
            None,
            None,
            &mut runs,
            &mut Vec::new(),
            &mut Vec::new(),
            &mut Vec::new(),
        );

        assert_eq!(runs.len(), 1);
        assert_eq!(runs[0].text, "Click here");
        assert_eq!(runs[0].hyperlink.as_deref(), Some("http://testuri.org/"));
    }

    /// The far more common, real-world shape: a hyperlink
    /// covering only PART of a text run's characters, via a sibling
    /// `MouseClickInteractiveInfoContainer` + `MouseClickTextInteractiveInfoAtom`
    /// pair in the `ClientTextbox` (not nested in `RT_CLIENT_DATA` at all —
    /// confirmed against real corpus bytes, not just the spec). The link
    /// is recorded on the run over exactly its characters; the run itself
    /// stays whole.
    #[test]
    fn test_text_range_hyperlink_is_recorded_on_the_run() {
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
            &HashMap::new(),
            None,
            None,
            &mut runs,
            &mut Vec::new(),
            &mut Vec::new(),
            &mut Vec::new(),
        );

        assert_eq!(runs.len(), 1, "{runs:?}");
        assert_eq!(runs[0].text, "See http://example.com/ here");
        assert_eq!(runs[0].hyperlink, None);
        assert_eq!(
            runs[0].link_ranges,
            [LinkRange {
                start: 4,
                end: 23,
                url: "http://example.com/".to_string()
            }]
        );
    }

    /// The hyperlinked range can cover the WHOLE run (no unlinked prefix
    /// or suffix) — must produce exactly one run, not empty placeholders.
    #[test]
    fn test_text_range_hyperlink_covering_the_whole_run_produces_one_run() {
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
            &HashMap::new(),
            None,
            None,
            &mut runs,
            &mut Vec::new(),
            &mut Vec::new(),
            &mut Vec::new(),
        );

        assert_eq!(runs.len(), 1);
        assert_eq!(runs[0].text, "clickme");
        assert_eq!(runs[0].hyperlink.as_deref(), Some("http://example.com/"));
    }

    // ── grid-of-shapes table reconstruction ──

    fn make_child_anchor(left: i32, top: i32, right: i32, bottom: i32) -> Vec<u8> {
        let mut body = Vec::new();
        body.extend_from_slice(&left.to_le_bytes());
        body.extend_from_slice(&top.to_le_bytes());
        body.extend_from_slice(&right.to_le_bytes());
        body.extend_from_slice(&bottom.to_le_bytes());
        make_atom(RT_CHILD_ANCHOR, 0, &body)
    }

    fn make_table_cell_shape(left: i32, top: i32, text: &[u8]) -> Vec<u8> {
        let mut children =
            make_child_anchor(left, top, left.saturating_add(100), top.saturating_add(50));
        // Tx_TYPE_OTHER
        let mut textbox_children = make_atom(RT_TEXT_HEADER, 0, &4u32.to_le_bytes());
        textbox_children.extend(make_atom(RT_TEXT_BYTES, 0, text));
        children.extend(make_container(0xF00D, 0, &textbox_children));
        make_container(RT_SHAPE, 0, &children)
    }

    /// A group whose members form a clean 2x2 grid must
    /// become one `TableBlock`, and the cell text must NOT also appear as
    /// flat paragraphs (that would duplicate it in the IR).
    #[test]
    fn test_spgr_container_with_clean_grid_becomes_a_table() {
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
            &HashMap::new(),
            None,
            None,
            &mut runs,
            &mut tables,
            &mut Vec::new(),
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

    /// Child anchors at the `i32` extremes reach the table heuristic
    /// straight from the file; they must neither overflow nor lose text.
    #[test]
    fn test_spgr_container_with_extreme_child_anchors_does_not_overflow() {
        let mut spgr_children = make_container(RT_SHAPE, 0, &[]);
        spgr_children.extend(make_table_cell_shape(i32::MIN, i32::MIN, b"A1"));
        spgr_children.extend(make_table_cell_shape(i32::MAX, i32::MIN, b"B1"));
        spgr_children.extend(make_table_cell_shape(i32::MIN, i32::MAX, b"A2"));
        spgr_children.extend(make_table_cell_shape(i32::MAX, i32::MAX, b"B2"));
        let spgr = make_container(RT_SPGR_CONTAINER, 0, &spgr_children);

        let mut runs = Vec::new();
        let mut tables = Vec::new();
        extract_shape_text(
            &spgr,
            0,
            &[],
            &HashMap::new(),
            &HashMap::new(),
            None,
            None,
            &mut runs,
            &mut tables,
            &mut Vec::new(),
            &mut Vec::new(),
        );
        assert_eq!(tables.len(), 1);
        assert_eq!(tables[0].rows[1][1][0].text, "B2");
    }

    /// A group that ISN'T a clean grid (here: only 3
    /// members, one short of a 2x2) must fall back to the ordinary
    /// flat-paragraph group walk, unchanged from before this feature
    /// existed — no text lost, just no table structure.
    #[test]
    fn test_spgr_container_that_is_not_a_grid_falls_back_to_flat_paragraphs() {
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
            &HashMap::new(),
            None,
            None,
            &mut runs,
            &mut tables,
            &mut Vec::new(),
            &mut Vec::new(),
        );

        assert!(tables.is_empty());
        assert_eq!(runs.len(), 3);
    }

    /// A group member with no `RT_CHILD_ANCHOR` at all (an
    /// odd/unexpected shape) must also bail to the flat-paragraph
    /// fallback rather than guessing a position for it.
    #[test]
    fn test_spgr_container_member_without_child_anchor_falls_back() {
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
            &HashMap::new(),
            None,
            None,
            &mut runs,
            &mut tables,
            &mut Vec::new(),
            &mut Vec::new(),
        );

        assert!(tables.is_empty());
        assert_eq!(runs.len(), 4);
    }

    /// A shape with no `InteractiveInfo` at all must never get a
    /// hyperlink, even when the document has some hyperlinks elsewhere.
    #[test]
    fn test_shape_without_interactive_info_has_no_hyperlink() {
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
            &HashMap::new(),
            None,
            None,
            &mut runs,
            &mut Vec::new(),
            &mut Vec::new(),
            &mut Vec::new(),
        );
        assert_eq!(runs[0].hyperlink, None);
    }

    /// `II_NoAction` (0x00) must never resolve a hyperlink, even when
    /// `exHyperlinkIdRef` happens to name a real entry — the action type
    /// gates whether the id is meaningful at all ([MS-PPT] 2.6.10).
    #[test]
    fn test_non_hyperlink_action_does_not_resolve_a_hyperlink() {
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
            &HashMap::new(),
            None,
            None,
            &mut runs,
            &mut Vec::new(),
            &mut Vec::new(),
            &mut Vec::new(),
        );
        assert_eq!(runs[0].hyperlink, None);
    }

    /// Full public pipeline: `extract_slides_text` on a
    /// synthetic "PowerPoint Document" stream carrying both an
    /// `ExObjListContainer` and a hyperlinked shape (with no persist
    /// directory or `SlideListWithText` — the "last resort" fallback,
    /// which still builds `hyperlinks` from the *whole* stream first)
    /// resolves the shape's `TextRun::hyperlink` end to end.
    #[test]
    fn test_extract_slides_text_resolves_hyperlinks_end_to_end() {
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
    fn test_extract_text_from_nested_shape_containers() {
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
            &HashMap::new(),
            None,
            None,
            &mut runs,
            &mut Vec::new(),
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
    fn test_slide_list_cache_fallback_without_persist_directory() {
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

    /// A `SlideShowSlideInfoAtom` HIDDEN_BIT sibling of the
    /// `Slide` container's real content must reach `SlideText::hidden`.
    #[test]
    fn test_hidden_slide_is_flagged() {
        let stream = hidden_slide_container_bytes("Hidden Slide");
        let slides = extract_slides_text(&stream, None);
        assert_eq!(slides.len(), 1);
        assert_eq!(slides[0].text_runs[0].text, "Hidden Slide");
        assert!(slides[0].hidden, "slide must be flagged hidden");
    }

    /// A slide with no `SlideShowSlideInfoAtom` at all — the overwhelming
    /// majority of real slides — must not be flagged hidden.
    #[test]
    fn test_ordinary_slide_is_not_hidden() {
        let stream = slide_container_bytes("Ordinary Slide");
        let slides = extract_slides_text(&stream, None);
        assert_eq!(slides.len(), 1);
        assert!(!slides[0].hidden);
    }

    /// A `SlideShowSlideInfoAtom` present but with HIDDEN_BIT clear (a
    /// slide that merely has a custom transition) must not be flagged
    /// hidden either.
    #[test]
    fn test_slide_info_atom_without_hidden_bit_is_not_hidden() {
        let header = make_atom(RT_TEXT_HEADER, 0, &0u32.to_le_bytes());
        let mut textbox_children = header;
        textbox_children.extend(make_atom(RT_TEXT_BYTES, 0, b"Visible Slide"));
        let textbox = make_container(0xF00D, 0, &textbox_children);

        let mut info_body = 0i32.to_le_bytes().to_vec();
        info_body.extend_from_slice(&0i32.to_le_bytes());
        info_body.push(0);
        info_body.push(5); // effectType = fade, unrelated to hidden
        info_body.extend_from_slice(&0x0001u16.to_le_bytes()); // MANUAL_ADVANCE_BIT only
        info_body.push(0);
        info_body.extend_from_slice(&[0, 0, 0]);
        let info_atom = make_atom(RT_SLIDE_SHOW_SLIDE_INFO_ATOM, 0, &info_body);

        let mut children = textbox;
        children.extend(info_atom);
        let stream = make_container(RT_SLIDE, 0, &children);

        let slides = extract_slides_text(&stream, None);
        assert_eq!(slides.len(), 1);
        assert!(!slides[0].hidden);
    }

    #[test]
    fn test_text_type_variants() {
        assert_eq!(TextType::from_u32(0), TextType::Title);
        assert_eq!(TextType::from_u32(1), TextType::Body);
        assert_eq!(TextType::from_u32(2), TextType::Notes);
        assert_eq!(TextType::from_u32(99), TextType::Other);
    }

    /// Every `TextTypeEnum` value from 3 upward was shifted
    /// one slot low (5 misread as `CenterTitle`, 6 as `HalfBody`, …),
    /// swapping a title slide's real title and subtitle. Values verified
    /// against the published [MS-PPT] `TextTypeEnum` spec page directly:
    /// 3 is genuinely undefined (the enum jumps from 2 to 4).
    #[test]
    fn test_text_type_values_3_and_up_match_the_spec_not_the_old_shifted_mapping() {
        assert_eq!(TextType::from_u32(3), TextType::Other, "3 is undefined in the spec");
        assert_eq!(TextType::from_u32(4), TextType::Other, "4 = Tx_TYPE_OTHER");
        assert_eq!(TextType::from_u32(5), TextType::CenterBody, "5 = Tx_TYPE_CENTERBODY");
        assert_eq!(TextType::from_u32(6), TextType::CenterTitle, "6 = Tx_TYPE_CENTERTITLE");
        assert_eq!(TextType::from_u32(7), TextType::HalfBody, "7 = Tx_TYPE_HALFBODY");
        assert_eq!(TextType::from_u32(8), TextType::QuarterBody, "8 = Tx_TYPE_QUARTERBODY");
    }

    #[test]
    fn test_decode_utf16le_basic() {
        let data = [0x41, 0x00, 0x42, 0x00, 0x43, 0x00]; // "ABC"
        assert_eq!(decode_utf16le(&data), "ABC");
    }

    #[test]
    fn test_fallback_when_no_slide_list() {
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

    fn level(
        bold: Option<bool>,
        italic: Option<bool>,
    ) -> (super::style::ParaFormat, super::style::CharFormat) {
        (
            super::style::ParaFormat::default(),
            super::style::CharFormat {
                bold,
                italic,
                ..Default::default()
            },
        )
    }

    /// Master inheritance covered only Title and Body at indent level 0: a
    /// nested bullet (level 1+) and every other text type (notes, other,
    /// the centred/half/quarter placeholder variants) got nothing from the
    /// master. Each paragraph now inherits from its own level, a level
    /// carries what it leaves unset from the level above it, and the
    /// placeholder variants fall back to their base style.
    #[test]
    fn test_master_inheritance_follows_indent_level_and_every_text_type() {
        let mut styles = MasterStyles::default();
        styles.levels[1] = cascade_levels(vec![level(Some(true), None), level(None, Some(true))]);
        styles.levels[0] = cascade_levels(vec![level(None, Some(true))]);
        assert_eq!(styles.levels[1][1].1.bold, Some(true), "level 1 cascades from level 0");

        let para = |start, end, indent| ParaFormatSpan {
            start,
            end,
            format: super::style::ParaFormat {
                indent_level: Some(indent),
                ..Default::default()
            },
        };
        let mut runs = vec![
            TextRun {
                text_type: TextType::Body,
                text: "Top\rNested".into(),
                para_formats: vec![para(0, 4, 0), para(4, 10, 1)],
                char_formats: vec![CharFormatSpan {
                    start: 0,
                    end: 10,
                    format: Default::default(),
                }],
                ..Default::default()
            },
            TextRun {
                text_type: TextType::CenterTitle,
                text: "Centred".into(),
                ..Default::default()
            },
            TextRun {
                text_type: TextType::HalfBody,
                text: "Half".into(),
                ..Default::default()
            },
        ];
        apply_master_inheritance(&mut runs, &styles);

        let body = &runs[0].char_formats;
        let at = |i: usize| body.iter().find(|s| s.start <= i && i < s.end).unwrap();
        assert_eq!((at(0).format.bold, at(0).format.italic), (Some(true), None));
        assert_eq!((at(5).format.bold, at(5).format.italic), (Some(true), Some(true)));
        // CenterTitle has no master style of its own: Title's applies.
        assert_eq!(runs[1].char_formats[0].format.italic, Some(true));
        // HalfBody falls back to Body level 0.
        assert_eq!(runs[2].char_formats[0].format.bold, Some(true));
    }

    fn slide_container_bytes(title: &str) -> Vec<u8> {
        let header = make_atom(RT_TEXT_HEADER, 0, &0u32.to_le_bytes());
        let mut textbox_children = header;
        textbox_children.extend(make_atom(RT_TEXT_BYTES, 0, title.as_bytes()));
        let textbox = make_container(0xF00D, 0, &textbox_children); // ClientTextbox
        make_container(RT_SLIDE, 0, &textbox)
    }

    /// A slide's own text that starts like an English master prompt was
    /// deleted on the degraded (no persist directory) paths; the masters
    /// are excluded by structure now, so slide text is never string-matched.
    #[test]
    fn test_slide_text_that_reads_like_a_prompt_is_kept() {
        let stream = slide_container_bytes("Click to add title: our roadmap");
        let slides = extract_slides_text(&stream, None);
        assert_eq!(slides[0].text_runs[0].text, "Click to add title: our roadmap");
    }

    /// The last-resort whole-stream dump skipped only *English* master
    /// prompts; a localized master's prompts leaked as content. Master
    /// containers are skipped whatever language their prompts are in.
    #[test]
    fn test_last_resort_dump_skips_master_containers_in_any_language() {
        let prompt = "Klicken Sie, um das Titelformat zu bearbeiten";
        let mut master_tb = make_atom(RT_TEXT_HEADER, 0, &0u32.to_le_bytes());
        master_tb.extend(make_atom(RT_TEXT_BYTES, 0, prompt.as_bytes()));
        let mut stream = make_container(RT_MAIN_MASTER, 0, &make_container(0xF00D, 0, &master_tb));
        // Loose text outside any slide container (the degraded shape).
        let mut loose = make_atom(RT_TEXT_HEADER, 0, &4u32.to_le_bytes());
        loose.extend(make_atom(RT_TEXT_BYTES, 0, b"Loose text"));
        stream.extend(make_container(0xF00D, 0, &loose));
        let slides = extract_slides_text(&stream, None);
        let all: Vec<&str> = slides
            .iter()
            .flat_map(|s| &s.text_runs)
            .map(|r| r.text.as_str())
            .collect();
        assert_eq!(all, ["Loose text"]);
    }

    /// Same shape as `slide_container_bytes`, plus a
    /// `SlideShowSlideInfoAtom` sibling of the `ClientTextbox` with
    /// `HIDDEN_BIT` (0x0004) set — the real byte shape [MS-PPT] 2.13.24
    /// describes, cross-checked against Apache POI's `SSSlideInfoAtom`.
    fn hidden_slide_container_bytes(title: &str) -> Vec<u8> {
        let header = make_atom(RT_TEXT_HEADER, 0, &0u32.to_le_bytes());
        let mut textbox_children = header;
        textbox_children.extend(make_atom(RT_TEXT_BYTES, 0, title.as_bytes()));
        let textbox = make_container(0xF00D, 0, &textbox_children); // ClientTextbox

        let mut info_body = 0i32.to_le_bytes().to_vec(); // slideTime
        info_body.extend_from_slice(&0i32.to_le_bytes()); // soundIdRef
        info_body.push(0); // effectDirection
        info_body.push(0); // effectType
        info_body.extend_from_slice(&0x0004u16.to_le_bytes()); // effectTransitionFlags: HIDDEN_BIT
        info_body.push(0); // speed
        info_body.extend_from_slice(&[0, 0, 0]); // unused
        let info_atom = make_atom(RT_SLIDE_SHOW_SLIDE_INFO_ATOM, 0, &info_body);

        let mut children = textbox;
        children.extend(info_atom);
        make_container(RT_SLIDE, 0, &children)
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
    fn test_persist_resolution_ignores_stale_orphaned_slide_copy() {
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

    /// `ExObjListContainer` resolution must go through the
    /// same persist-directory mechanism slides do, not a raw whole-stream
    /// scan: a stale/orphaned `ExObjListContainer` left behind by an
    /// earlier incremental save (mapping the same `exHyperlinkId` to a
    /// *different*, superseded URL) must never win over the current,
    /// persist-resolved one.
    #[test]
    fn test_persist_resolution_uses_the_current_exobjlist_not_a_stale_one() {
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
    fn test_persist_resolution_works_without_current_user_stream() {
        let (stream, _current_user) = build_persist_regression_fixture();
        let slides = extract_slides_text(&stream, None);

        assert_eq!(slides.len(), 1);
        assert_eq!(slides[0].text_runs[0].text, "REAL SLIDE TEXT");
    }

    /// Real byte shape confirmed against a live corpus file's
    /// `DocumentContainer`: a `HeadersFootersContainer` (0x0FD9) holding
    /// a `HeadersFootersAtom` (0x0FDA) plus `CString` (0x0FBA) children
    /// for the user date (instance 0) and footer (instance 2) text.
    fn header_footer_container_bytes(date: &str, footer: &str) -> Vec<u8> {
        // formatId 0; flags fHasDate | fHasUserDate | fHasFooter.
        let flags = HF_HAS_DATE | HF_HAS_USER_DATE | HF_HAS_FOOTER;
        let mut atom = 0u16.to_le_bytes().to_vec();
        atom.extend_from_slice(&flags.to_le_bytes());
        let mut children = make_atom(RT_HEADER_FOOTER_ATOM, 0, &atom);
        let date_bytes: Vec<u8> = date.encode_utf16().flat_map(|u| u.to_le_bytes()).collect();
        children.extend(make_atom(RT_CSTRING, 0, &date_bytes));
        let footer_bytes: Vec<u8> = footer
            .encode_utf16()
            .flat_map(|u| u.to_le_bytes())
            .collect();
        children.extend(make_atom(RT_CSTRING, 2, &footer_bytes));
        make_container(RT_HEADER_FOOTER, 3, &children)
    }

    /// A `DocumentContainer`-level `HeadersFootersContainer`'s
    /// date/footer text must reach every slide's text runs, matching the
    /// real corpus file this was verified against
    /// (`26 August 2004` / `Transport CDM Workshop`, repeated per slide).
    #[test]
    fn test_header_footer_text_reaches_every_slide() {
        let mut stream = Vec::new();

        let hf = header_footer_container_bytes("26 August 2004", "Transport CDM Workshop");
        let slide_persist = slide_persist_atom_bytes(2, 256);
        let slide_list = make_container(RT_SLIDE_LIST_WITH_TEXT, SLWT_SLIDES, &slide_persist);
        let mut doc_children = hf;
        doc_children.extend(slide_list);
        let doc_offset = stream.len() as u32;
        stream.extend(make_container(RT_DOCUMENT, 0, &doc_children));

        let real_slide_offset = stream.len() as u32;
        stream.extend(slide_container_bytes("Slide Title"));

        let pd_offset = stream.len() as u32;
        stream.extend(persist_directory_bytes(&[(1, doc_offset), (2, real_slide_offset)]));
        let edit_offset = stream.len() as u32;
        stream.extend(user_edit_atom_bytes(0, pd_offset, 1));
        let current_user = current_user_bytes(edit_offset);

        let slides = extract_slides_text(&stream, Some(&current_user));
        assert_eq!(slides.len(), 1);
        let texts: Vec<&str> = slides[0]
            .text_runs
            .iter()
            .map(|r| r.text.as_str())
            .collect();
        assert!(texts.contains(&"Slide Title"), "{texts:?}");
        assert!(texts.contains(&"26 August 2004"), "{texts:?}");
        assert!(texts.contains(&"Transport CDM Workshop"), "{texts:?}");
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
    fn test_corrupted_record_length_before_real_content_does_not_lose_it() {
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
    fn test_outline_text_ref_atom_resolves_indexed_placeholder_text() {
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
            &HashMap::new(),
            None,
            None,
            &mut runs,
            &mut Vec::new(),
            &mut Vec::new(),
            &mut Vec::new(),
        );
        assert_eq!(runs.len(), 1);
        assert_eq!(runs[0].text, "second");
        assert_eq!(runs[0].text_type, TextType::Body);
    }

    /// A `SlideAtom` body: 12 opaque bytes of embedded `SSlideLayoutAtom`,
    /// then `masterIdRef: i32`, `notesIdRef: i32`, `flags: u16`.
    fn slide_atom_bytes(master_id_ref: u32) -> Vec<u8> {
        let mut body = vec![0u8; 12];
        body.extend_from_slice(&master_id_ref.to_le_bytes());
        body.extend_from_slice(&0u32.to_le_bytes());
        body.extend_from_slice(&0u16.to_le_bytes());
        make_atom(RT_SLIDE_ATOM, 0, &body)
    }

    /// A `TxMasterStyleAtom` body with exactly one indent level (0),
    /// setting only `alignment` (`PFMasks` bit 11) and `font_size`
    /// (`CFMasks` bit 17) — the same two bits `style.rs`'s own
    /// `parse_pf_body`/`parse_cf_body` decode. `text_type` `0`
    /// (`Title`)/`1` (`Body`) are both `< 5`, so no per-level
    /// `indentLevel` field is present.
    fn master_style_bytes(text_type: u16, alignment: u16, font_size: i16) -> Vec<u8> {
        let mut body = Vec::new();
        body.extend_from_slice(&1u16.to_le_bytes()); // levels = 1
        const PF_ALIGN: u32 = 1 << 11;
        body.extend_from_slice(&PF_ALIGN.to_le_bytes());
        body.extend_from_slice(&alignment.to_le_bytes());
        const CF_SIZE: u32 = 1 << 17;
        body.extend_from_slice(&CF_SIZE.to_le_bytes());
        body.extend_from_slice(&font_size.to_le_bytes());
        make_atom(RT_TX_MASTER_STYLE_ATOM, text_type, &body)
    }

    #[test]
    fn test_placeholder_formatting_inherits_from_master_when_direct_formatting_is_absent() {
        // A Title placeholder with no StyleTextPropAtom of its own at
        // all (no direct formatting) whose slide's SlideAtom.masterIdRef
        // names a MainMaster carrying a Title TxMasterStyleAtom. The run
        // must end up with the master's alignment/font_size.
        //
        // The reference is a `masterId` (0x8000000C here, as in a real
        // Apache Tika sample), mapped to the master's persist id (3)
        // through the MasterListWithTextContainer — not a persist id with
        // its high bit masked, which this test used to encode and which
        // resolves to nothing on real decks.
        let mut stream = Vec::new();

        let doc_offset = stream.len() as u32;
        let mut slide_list_children = slide_persist_atom_bytes(2, 256);
        slide_list_children.extend(make_atom(RT_TEXT_HEADER, 0, &0u32.to_le_bytes())); // type=Title
        slide_list_children.extend(make_atom(RT_TEXT_BYTES, 0, b"Title Text"));
        let mut doc_children = make_container(
            RT_SLIDE_LIST_WITH_TEXT,
            SLWT_MASTERS,
            &slide_persist_atom_bytes(3, 0x8000_000C),
        );
        doc_children.extend(make_container(
            RT_SLIDE_LIST_WITH_TEXT,
            SLWT_SLIDES,
            &slide_list_children,
        ));
        stream.extend(make_container(RT_DOCUMENT, 0, &doc_children));

        let slide_offset = stream.len() as u32;
        let index_bytes = 0i32.to_le_bytes();
        let outline_ref = make_atom(RT_OUTLINE_TEXT_REF_ATOM, 0, &index_bytes);
        let textbox = make_container(0xF00D, 0, &outline_ref);
        let shape = make_container(0xF004, 0, &textbox);
        let mut slide_children = slide_atom_bytes(0x8000_000C);
        slide_children.extend(shape);
        stream.extend(make_container(RT_SLIDE, 0, &slide_children));

        let master_offset = stream.len() as u32;
        let master_children = master_style_bytes(0, 1, 44); // Title, center, 44pt
        stream.extend(make_container(RT_MAIN_MASTER, 0, &master_children));

        let pd_offset = stream.len() as u32;
        stream.extend(persist_directory_bytes(&[
            (1, doc_offset),
            (2, slide_offset),
            (3, master_offset),
        ]));
        let edit_offset = stream.len() as u32;
        stream.extend(user_edit_atom_bytes(0, pd_offset, 1));
        let current_user = current_user_bytes(edit_offset);

        let slides = extract_slides_text(&stream, Some(&current_user));
        assert_eq!(slides.len(), 1);
        let run = &slides[0].text_runs[0];
        assert_eq!(run.text, "Title Text");
        assert_eq!(
            run.char_formats.len(),
            1,
            "expected one synthetic master-inherited span: {:?}",
            run.char_formats
        );
        assert_eq!(
            run.char_formats[0].format.font_size,
            Some(44),
            "font size must inherit from the master"
        );
        assert_eq!(run.para_formats.len(), 1);
        assert_eq!(
            run.para_formats[0].format.alignment,
            Some(1),
            "alignment must inherit from the master"
        );
    }

    #[test]
    fn test_persist_resolution_follows_outline_text_ref_atom() {
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
    fn test_falls_back_to_inline_cache_when_persist_resolved_slides_are_all_textless() {
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

    /// A `SlideAtom` naming `master_id` with `fMasterObjects` as given.
    fn slide_atom_following(master_id: u32, follows: bool) -> Vec<u8> {
        let mut body = vec![0u8; 24];
        body[12..16].copy_from_slice(&master_id.to_le_bytes());
        if follows {
            body[20..22].copy_from_slice(&SLIDE_FLAG_MASTER_OBJECTS.to_le_bytes());
        }
        make_atom(RT_SLIDE_ATOM, 2, &body)
    }

    /// A shape holding one `Tx_TYPE_OTHER` text, with `client_data`
    /// children (a placeholder atom) when given.
    fn other_text_shape(text: &str, client_data: Option<Vec<u8>>) -> Vec<u8> {
        let mut children = Vec::new();
        if let Some(cd) = client_data {
            children.extend(make_container(RT_CLIENT_DATA, 0, &cd));
        }
        let mut tb = make_atom(RT_TEXT_HEADER, 0, &4u32.to_le_bytes());
        tb.extend(make_atom(RT_TEXT_BYTES, 0, text.as_bytes()));
        children.extend(make_container(0xF00D, 0, &tb));
        make_container(RT_SHAPE, 0, &children)
    }

    /// A title slide names a title master (a `SlideContainer` in the
    /// master list), which names its main master: both masters' static
    /// text is shown and extracted, each once, while header/footer
    /// placeholders (`RoundTripHFPlaceholder12Atom`) are not static text.
    /// `masterIdRef` values are `masterId`s from the
    /// `MasterListWithTextContainer`, not persist ids.
    #[test]
    fn test_title_master_chain_yields_static_text_of_both_masters() {
        const TITLE_MASTER: u32 = 0x8000_0010;
        const MAIN_MASTER: u32 = 0x8000_0020;
        let mut stream = Vec::new();

        let doc_offset = stream.len() as u32;
        let mut doc = make_container(
            RT_SLIDE_LIST_WITH_TEXT,
            SLWT_MASTERS,
            &[
                slide_persist_atom_bytes(3, TITLE_MASTER),
                slide_persist_atom_bytes(4, MAIN_MASTER),
            ]
            .concat(),
        );
        doc.extend(make_container(
            RT_SLIDE_LIST_WITH_TEXT,
            SLWT_SLIDES,
            &[
                slide_persist_atom_bytes(2, 256),
                slide_persist_atom_bytes(5, 257),
            ]
            .concat(),
        ));
        stream.extend(make_container(RT_DOCUMENT, 0, &doc));

        let slide_offset = stream.len() as u32;
        let mut slide = slide_atom_following(TITLE_MASTER, true);
        slide.extend(other_text_shape("SLIDE TEXT", None));
        stream.extend(make_container(RT_SLIDE, 0, &slide));

        let title_master_offset = stream.len() as u32;
        let mut title_master = slide_atom_following(MAIN_MASTER, true);
        title_master.extend(other_text_shape("TITLE MASTER TEXT", None));
        title_master.extend(other_text_shape("SHARED TEXT", None));
        stream.extend(make_container(RT_SLIDE, 0, &title_master));

        let main_master_offset = stream.len() as u32;
        let mut main_master = slide_atom_following(0, false);
        // Tx_TYPE_OTHER style: right-aligned, 30pt.
        main_master.extend(master_style_bytes(4, 2, 30));
        main_master.extend(other_text_shape("MAIN MASTER TEXT", None));
        main_master.extend(other_text_shape("SHARED TEXT", None));
        main_master.extend(other_text_shape(
            "FOOTER PROMPT",
            Some(make_atom(RT_ROUND_TRIP_HF_PLACEHOLDER12_ATOM, 0, &[9])),
        ));
        stream.extend(make_container(RT_MAIN_MASTER, 0, &main_master));

        // A second slide on the main master, hiding its objects.
        let hidden_offset = stream.len() as u32;
        let mut hidden = slide_atom_following(MAIN_MASTER, false);
        hidden.extend(other_text_shape("SECOND SLIDE", None));
        stream.extend(make_container(RT_SLIDE, 0, &hidden));

        let pd_offset = stream.len() as u32;
        stream.extend(persist_directory_bytes(&[
            (1, doc_offset),
            (2, slide_offset),
            (3, title_master_offset),
            (4, main_master_offset),
            (5, hidden_offset),
        ]));
        let edit_offset = stream.len() as u32;
        stream.extend(user_edit_atom_bytes(0, pd_offset, 1));
        let current_user = current_user_bytes(edit_offset);

        let deck = extract_deck_text(&stream, Some(&current_user));
        assert_eq!(deck.slides.len(), 2);
        let master: Vec<&str> = deck.master_text.iter().map(|r| r.text.as_str()).collect();
        assert_eq!(master, ["TITLE MASTER TEXT", "SHARED TEXT", "MAIN MASTER TEXT"]);
        // Text styles come from the main master behind the title master.
        let slide_run = &deck.slides[0].text_runs[0];
        assert_eq!(slide_run.text, "SLIDE TEXT");
        assert_eq!(slide_run.char_formats[0].format.font_size, Some(30));
        assert_eq!(slide_run.para_formats[0].format.alignment, Some(2));
    }

    /// A master reference cycle (title master naming itself through a
    /// second title master) terminates.
    #[test]
    fn test_master_reference_cycle_terminates() {
        const A: u32 = 0x8000_0001;
        const B: u32 = 0x8000_0002;
        let mut stream = Vec::new();
        let doc_offset = stream.len() as u32;
        let mut doc = make_container(
            RT_SLIDE_LIST_WITH_TEXT,
            SLWT_MASTERS,
            &[
                slide_persist_atom_bytes(3, A),
                slide_persist_atom_bytes(4, B),
            ]
            .concat(),
        );
        doc.extend(make_container(
            RT_SLIDE_LIST_WITH_TEXT,
            SLWT_SLIDES,
            &slide_persist_atom_bytes(2, 256),
        ));
        stream.extend(make_container(RT_DOCUMENT, 0, &doc));
        let mut offsets = vec![(1, doc_offset)];
        for (pid, next, text) in [(2, A, "SLIDE"), (3, B, "A TEXT"), (4, A, "B TEXT")] {
            offsets.push((pid, stream.len() as u32));
            let mut c = slide_atom_following(next, true);
            c.extend(other_text_shape(text, None));
            stream.extend(make_container(RT_SLIDE, 0, &c));
        }
        let pd_offset = stream.len() as u32;
        stream.extend(persist_directory_bytes(&offsets));
        let edit_offset = stream.len() as u32;
        stream.extend(user_edit_atom_bytes(0, pd_offset, 1));
        let deck = extract_deck_text(&stream, Some(&current_user_bytes(edit_offset)));
        let master: Vec<&str> = deck.master_text.iter().map(|r| r.text.as_str()).collect();
        assert_eq!(master, ["A TEXT", "B TEXT"]);
    }
}
