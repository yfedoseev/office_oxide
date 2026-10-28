//! Parts of a spreadsheet package that several sheets can name.
//!
//! A drawing, a comments part or a table definition is stored once in the
//! package, and every sheet whose relationships point at it uses it. Reading
//! it for each sheet that names it multiplied the work by the sheet count (a
//! few hundred kilobytes of package named by hundreds of sheets was gigabytes
//! of copies), so each is read and parsed once, by the zip entry its path
//! resolves to, and the sheets share the result.
//!
//! What a sheet's *use* of a shared part adds to the IR and to the renderings
//! is the converters' to charge, before they copy it.

use std::collections::HashMap;
use std::io::{Read, Seek};
use std::sync::Arc;

use zip::read::ZipArchive;

use super::worksheet::{self, SheetComment, SheetTable, WorksheetPicture, WorksheetTextShape};
use super::{
    MediaCache, XlsxDocument, guess_image_format_from_bytes, parse_drawing_anchors, part_key,
    read_media_cached, read_rels_reporting, resolve_relative_zip_path, sheet_rels_path,
};
use crate::core::opc;
use crate::core::relationships::{Relationships, rel_types};

/// Most reports kept in a [`ReportList`]; the rest are counted.
pub(super) const MAX_REPORTS: usize = 100;
/// Longest name or reason kept in a report, in characters: both can echo
/// values the file controls.
const MAX_REPORT_CHARS: usize = 256;

/// What could not be read, as `(name, reason)`: each item once, at most
/// [`MAX_REPORTS`] kept, each name and reason cut to a fixed length, and the
/// number of the rest. One sink for the relationship and other parts that
/// could not be read and for the sheets that could not, so that however many
/// the file names, and however long it makes their names, what is kept (and
/// every renderer echoes) is bounded.
#[derive(Default)]
pub(crate) struct ReportList {
    pub(super) listed: Vec<(String, String)>,
    pub(super) omitted: usize,
    /// Hashes of the parts recorded, so that the names of parts the file
    /// makes up (a missing part has none but its path) are not all kept.
    seen: std::collections::HashSet<u64>,
}

impl ReportList {
    /// Record that the part at `path` could not be read, for `reason`. A part
    /// already recorded (by the zip entry its path resolves to) is not
    /// recorded again, whether or not it was kept.
    pub(super) fn record(&mut self, entries: &opc::ZipEntryIndex, path: &str, reason: &str) {
        use std::hash::{Hash, Hasher};
        let mut hasher = std::collections::hash_map::DefaultHasher::new();
        part_key(entries, path).hash(&mut hasher);
        if !self.seen.insert(hasher.finish()) {
            return;
        }
        self.push(path, reason);
    }

    /// Record that the item `name` could not be read, for `reason`. Every
    /// call is one item: two sheets of the same name are two.
    pub(super) fn record_item(&mut self, name: &str, reason: &str) {
        self.push(name, reason);
    }

    fn push(&mut self, name: &str, reason: &str) {
        if self.listed.len() >= MAX_REPORTS {
            self.omitted += 1;
            return;
        }
        let cut = |s: &str| s.chars().take(MAX_REPORT_CHARS).collect::<String>();
        self.listed.push((cut(name), cut(reason)));
    }
}

/// Whether a relationship names a part of the package: a target outside it
/// ([ECMA-376] Part 2 §9.3.2, `TargetMode="External"`) is never looked up in it.
pub(super) fn is_internal(rel: &crate::core::relationships::Relationship) -> bool {
    rel.target_mode != crate::core::relationships::TargetMode::External
}

/// A drawing part, parsed: the pictures and text shapes its anchors place.
pub(super) struct SharedDrawing {
    pub(super) pictures: Arc<[WorksheetPicture]>,
    pub(super) text_shapes: Arc<[WorksheetTextShape]>,
}

/// The parts already read, by the zip entry each is stored in. A part that
/// could not be read or parsed is remembered too, so that it is not tried
/// again by every sheet that names it.
#[derive(Default)]
pub(super) struct SharedParts {
    drawings: HashMap<String, Option<Arc<SharedDrawing>>>,
    comments: HashMap<String, Option<Arc<[SheetComment]>>>,
    threaded: HashMap<String, Option<Arc<worksheet::ThreadedComments>>>,
    /// Merged comments by the pair of parts merged.
    merged: HashMap<(String, String), (Arc<[SheetComment]>, bool)>,
    tables: HashMap<String, Option<Arc<SheetTable>>>,
    /// The text of chart parts, for the chartsheets that display them.
    charts: HashMap<String, Option<Arc<str>>>,
    /// The chart parts each drawing names, for the chartsheets that display
    /// them, by drawing: the drawing's relationships are read once however
    /// many chartsheets name it.
    drawing_charts: HashMap<String, Arc<[String]>>,
    /// Relationship parts met on the way that could not be read, as
    /// `(part, reason)`.
    pub(super) unreadable: ReportList,
}

impl SharedParts {
    /// The drawing `sheet_rels` names for the sheet at `sheet_path`, with its
    /// pictures' bytes resolved. `None` when the sheet has no drawing or the
    /// part cannot be read: drawings are best-effort extras and do not fail
    /// the worksheet.
    pub(super) fn drawing_for_sheet<R: Read + Seek>(
        &mut self,
        archive: &mut ZipArchive<R>,
        entries: &opc::ZipEntryIndex,
        media: &mut MediaCache,
        sheet_path: &str,
        sheet_rels: &Relationships,
    ) -> Option<Arc<SharedDrawing>> {
        let rel = sheet_rels
            .first_by_type(rel_types::DRAWING)
            .filter(|rel| is_internal(rel))?;
        let drawing_path = resolve_relative_zip_path(sheet_path, &rel.target);
        let key = part_key(entries, &drawing_path);
        if let Some(hit) = self.drawings.get(&key) {
            return hit.clone();
        }
        let read = read_drawing(archive, entries, media, &drawing_path, &mut self.unreadable)
            .map(Arc::new);
        self.drawings.insert(key, read.clone());
        read
    }

    /// The comments of the sheet at `sheet_path`, legacy and threaded merged,
    /// and whether the merge stopped at `budget`. Both parts are read once
    /// however many sheets name them, and so is each pair merged.
    pub(super) fn comments_for_sheet<R: Read + Seek>(
        &mut self,
        archive: &mut ZipArchive<R>,
        entries: &opc::ZipEntryIndex,
        sheet_path: &str,
        sheet_rels: &Relationships,
        persons: &HashMap<String, Arc<str>>,
        budget: &mut crate::limits::TextBudget,
    ) -> (Arc<[SheetComment]>, bool) {
        let legacy_path = sheet_rels
            .first_by_type(rel_types::COMMENTS)
            .filter(|rel| is_internal(rel))
            .map(|rel| resolve_relative_zip_path(sheet_path, &rel.target));
        let threaded_path = sheet_rels
            .first_by_type(rel_types::THREADED_COMMENTS)
            .filter(|rel| is_internal(rel))
            .map(|rel| resolve_relative_zip_path(sheet_path, &rel.target));
        let legacy_key = legacy_path.as_deref().map(|p| part_key(entries, p));
        let threaded_key = threaded_path.as_deref().map(|p| part_key(entries, p));
        let pair =
            (legacy_key.clone().unwrap_or_default(), threaded_key.clone().unwrap_or_default());
        if let Some(hit) = self.merged.get(&pair) {
            return hit.clone();
        }

        let legacy: Arc<[SheetComment]> = match (legacy_path, legacy_key) {
            (Some(path), Some(key)) => {
                let unreadable = &mut self.unreadable;
                self.comments
                    .entry(key)
                    .or_insert_with(|| {
                        read_parsed(archive, entries, &path, unreadable, worksheet::parse_comments)
                            .map(Arc::from)
                    })
                    .clone()
                    .unwrap_or_else(|| Arc::from([]))
            },
            _ => Arc::from([]),
        };
        let threaded: Arc<worksheet::ThreadedComments> = match (threaded_path, threaded_key) {
            (Some(path), Some(key)) => {
                let unreadable = &mut self.unreadable;
                self.threaded
                    .entry(key)
                    .or_insert_with(|| {
                        read_parsed(
                            archive,
                            entries,
                            &path,
                            unreadable,
                            worksheet::parse_threaded_comments,
                        )
                        .map(|messages| Arc::new(worksheet::ThreadedComments::new(messages)))
                    })
                    .clone()
                    .unwrap_or_default()
            },
            _ => Arc::default(),
        };
        let merged = if threaded.is_empty() {
            (legacy, false)
        } else {
            let (comments, truncated) =
                worksheet::merge_threaded_comments(&legacy, &threaded, persons, budget);
            (Arc::from(comments), truncated)
        };
        self.merged.insert(pair, merged.clone());
        merged
    }

    /// The chart parts a chartsheet displays: its drawing's chart
    /// relationships ([ECMA-376] §12.3.2 — a chartsheet shows one chart
    /// through a drawing part). A part that is missing or unrelated yields
    /// nothing; the sheet is then its name alone. Each chart is listed once.
    pub(super) fn chartsheet_chart_parts<R: Read + Seek>(
        &mut self,
        archive: &mut ZipArchive<R>,
        entries: &opc::ZipEntryIndex,
        sheet_path: &str,
    ) -> Vec<String> {
        let sheet_rels = read_rels_reporting(
            archive,
            entries,
            &sheet_rels_path(sheet_path),
            &mut self.unreadable,
        );
        let mut charts: Vec<String> = Vec::new();
        let mut listed = std::collections::HashSet::new();
        for drawing in sheet_rels
            .get_by_type(rel_types::DRAWING)
            .into_iter()
            .filter(|rel| is_internal(rel))
        {
            let drawing_path = resolve_relative_zip_path(sheet_path, &drawing.target);
            let key = part_key(entries, &drawing_path);
            let named = match self.drawing_charts.get(&key) {
                Some(hit) => hit.clone(),
                None => {
                    let drawing_rels = read_rels_reporting(
                        archive,
                        entries,
                        &sheet_rels_path(&drawing_path),
                        &mut self.unreadable,
                    );
                    let named: Arc<[String]> = drawing_rels
                        .get_by_type(rel_types::CHART)
                        .into_iter()
                        .filter(|rel| is_internal(rel))
                        .map(|chart| resolve_relative_zip_path(&drawing_path, &chart.target))
                        .collect();
                    self.drawing_charts.insert(key, named.clone());
                    named
                },
            };
            for path in named.iter() {
                if listed.insert(part_key(entries, path)) {
                    charts.push(path.clone());
                }
            }
        }
        charts
    }

    /// The text of the chart part at `path`, read once; `None` when the part
    /// cannot be read or holds no text.
    pub(super) fn chart_text<R: Read + Seek>(
        &mut self,
        archive: &mut ZipArchive<R>,
        entries: &opc::ZipEntryIndex,
        path: &str,
    ) -> Option<Arc<str>> {
        let unreadable = &mut self.unreadable;
        self.charts
            .entry(part_key(entries, path))
            .or_insert_with(|| {
                let data = read_reporting(archive, entries, path, unreadable)?;
                let text = super::extract_chart_text(&data);
                (!text.is_empty()).then(|| Arc::from(text))
            })
            .clone()
    }

    /// The table part at `path`, parsed once. `None` (and a report, once)
    /// when it is missing or does not parse: the cells it describes are in
    /// the sheet either way.
    pub(super) fn table<R: Read + Seek>(
        &mut self,
        archive: &mut ZipArchive<R>,
        entries: &opc::ZipEntryIndex,
        path: &str,
    ) -> Option<Arc<SheetTable>> {
        let unreadable = &mut self.unreadable;
        self.tables
            .entry(part_key(entries, path))
            .or_insert_with(|| {
                read_parsed(archive, entries, path, unreadable, worksheet::parse_table_part)
                    .map(Arc::new)
            })
            .clone()
    }
}

/// The part at `path`, read; `None` when it cannot be, with the reason
/// recorded (once per part) in `unreadable`. Callers keep the outcome, so the
/// part is not tried again by every reference to it.
pub(super) fn read_reporting<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    entries: &opc::ZipEntryIndex,
    path: &str,
    unreadable: &mut ReportList,
) -> Option<Vec<u8>> {
    match XlsxDocument::read_xml_entry(archive, entries, path) {
        Ok(data) => Some(data),
        // A relationship to a part the package lacks is dangling, which is
        // not a part that cannot be read: only a part that is there and
        // cannot be read or parsed is reported.
        Err(crate::core::Error::MissingPart(_)) => None,
        Err(e) => {
            log::warn!("xlsx: part {path} unreadable: {e}");
            unreadable.record(entries, path, &e.to_string());
            None
        },
    }
}

/// The part at `path`, read and parsed by `parse`; `None` with the reason
/// recorded when either fails.
pub(super) fn read_parsed<R: Read + Seek, T>(
    archive: &mut ZipArchive<R>,
    entries: &opc::ZipEntryIndex,
    path: &str,
    unreadable: &mut ReportList,
    parse: impl FnOnce(&[u8]) -> crate::core::Result<T>,
) -> Option<T> {
    let data = read_reporting(archive, entries, path, unreadable)?;
    match parse(&data) {
        Ok(parsed) => Some(parsed),
        Err(e) => {
            log::warn!("xlsx: part {path} unreadable: {e}");
            unreadable.record(entries, path, &e.to_string());
            None
        },
    }
}

/// Read and parse the drawing part at `drawing_path`: its `<xdr:pic>` and
/// `<xdr:sp>` anchors, and each picture's underlying media bytes.
fn read_drawing<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    entries: &opc::ZipEntryIndex,
    media: &mut MediaCache,
    drawing_path: &str,
    unreadable: &mut ReportList,
) -> Option<SharedDrawing> {
    let drawing_xml = read_reporting(archive, entries, drawing_path, unreadable)?;

    let drawing_rels =
        read_rels_reporting(archive, entries, &sheet_rels_path(drawing_path), unreadable);

    let parsed = match parse_drawing_anchors(&drawing_xml) {
        Ok(a) => a,
        Err(e) => {
            log::warn!("xlsx: drawing {drawing_path} failed to parse ({e}); dropping anchors");
            unreadable.record(entries, drawing_path, &e.to_string());
            return None;
        },
    };

    // Resolve picture anchors → bytes.
    let mut pictures = Vec::with_capacity(parsed.pictures.len());
    for a in parsed.pictures {
        let Some(rel) = drawing_rels.get_by_id(&a.embed_rid) else {
            continue;
        };
        let media_path = resolve_relative_zip_path(drawing_path, &rel.target);
        // A linked picture names a target outside the package.
        if rel.target_mode == crate::core::relationships::TargetMode::External {
            continue;
        }
        let Some(bytes) = read_media_cached(archive, entries, &media_path, media, unreadable)
        else {
            continue;
        };
        let ext = std::path::Path::new(&*rel.target)
            .extension()
            .and_then(|s| s.to_str())
            .map(|s| s.to_ascii_lowercase())
            .unwrap_or_else(|| guess_image_format_from_bytes(&bytes).to_string());

        pictures.push(WorksheetPicture {
            data: bytes,
            format: ext,
            x_emu: a.x_emu,
            y_emu: a.y_emu,
            cx_emu: a.cx_emu,
            cy_emu: a.cy_emu,
            alt_text: a.alt_text,
        });
    }

    let text_shapes: Vec<WorksheetTextShape> = parsed
        .text_shapes
        .into_iter()
        .map(|t| WorksheetTextShape {
            text: t.text,
            font_name: t.font_name,
            font_size_pt: t.font_size_pt,
            bold: t.bold,
            italic: t.italic,
            color_hex: t.color_hex,
            x_emu: t.x_emu,
            y_emu: t.y_emu,
            cx_emu: t.cx_emu,
            cy_emu: t.cy_emu,
        })
        .collect();

    Some(SharedDrawing {
        pictures: pictures.into(),
        text_shapes: text_shapes.into(),
    })
}
