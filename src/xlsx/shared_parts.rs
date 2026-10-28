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

/// Most relationship parts recorded as unreadable; the rest are counted.
pub(super) const MAX_UNREADABLE_PARTS: usize = 100;
/// Longest part name or reason kept in a record, in characters: both can
/// echo values the file controls.
const MAX_REPORT_CHARS: usize = 256;

/// The relationship parts that could not be read: each part once, at most
/// [`MAX_UNREADABLE_PARTS`] kept, and the number of the rest.
#[derive(Default)]
pub(crate) struct UnreadableParts {
    pub(super) listed: Vec<(String, String)>,
    pub(super) omitted: usize,
    seen: std::collections::HashSet<String>,
}

impl UnreadableParts {
    /// Record that the part at `path` could not be read, for `reason`. A part
    /// already recorded (by the zip entry its path resolves to) is not
    /// recorded again.
    pub(super) fn record(&mut self, entries: &opc::ZipEntryIndex, path: &str, reason: &str) {
        if !self.seen.insert(part_key(entries, path)) {
            return;
        }
        if self.listed.len() >= MAX_UNREADABLE_PARTS {
            self.omitted += 1;
            return;
        }
        let cut = |s: &str| s.chars().take(MAX_REPORT_CHARS).collect::<String>();
        self.listed.push((cut(path), cut(reason)));
    }
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
    pub(super) unreadable: UnreadableParts,
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
        let rel = sheet_rels.first_by_type(rel_types::DRAWING)?;
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
            .map(|rel| resolve_relative_zip_path(sheet_path, &rel.target));
        let threaded_path = sheet_rels
            .first_by_type(rel_types::THREADED_COMMENTS)
            .map(|rel| resolve_relative_zip_path(sheet_path, &rel.target));
        let legacy_key = legacy_path.as_deref().map(|p| part_key(entries, p));
        let threaded_key = threaded_path.as_deref().map(|p| part_key(entries, p));
        let pair =
            (legacy_key.clone().unwrap_or_default(), threaded_key.clone().unwrap_or_default());
        if let Some(hit) = self.merged.get(&pair) {
            return hit.clone();
        }

        let legacy: Arc<[SheetComment]> = match (legacy_path, legacy_key) {
            (Some(path), Some(key)) => self
                .comments
                .entry(key)
                .or_insert_with(|| {
                    XlsxDocument::read_xml_entry(archive, entries, &path)
                        .ok()
                        .and_then(|d| worksheet::parse_comments(&d).ok())
                        .map(Arc::from)
                })
                .clone()
                .unwrap_or_else(|| Arc::from([])),
            _ => Arc::from([]),
        };
        let threaded: Arc<worksheet::ThreadedComments> = match (threaded_path, threaded_key) {
            (Some(path), Some(key)) => self
                .threaded
                .entry(key)
                .or_insert_with(|| {
                    XlsxDocument::read_xml_entry(archive, entries, &path)
                        .ok()
                        .and_then(|d| worksheet::parse_threaded_comments(&d).ok())
                        .map(|messages| Arc::new(worksheet::ThreadedComments::new(messages)))
                })
                .clone()
                .unwrap_or_default(),
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
        for drawing in sheet_rels.get_by_type(rel_types::DRAWING) {
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
        self.charts
            .entry(part_key(entries, path))
            .or_insert_with(|| {
                let data = XlsxDocument::read_xml_entry(archive, entries, path).ok()?;
                let text = super::extract_chart_text(&data);
                (!text.is_empty()).then(|| Arc::from(text))
            })
            .clone()
    }

    /// The table part at `path`, parsed once. `None` (and a log line, once)
    /// when it is missing or does not parse: the cells it describes are in
    /// the sheet either way.
    pub(super) fn table<R: Read + Seek>(
        &mut self,
        archive: &mut ZipArchive<R>,
        entries: &opc::ZipEntryIndex,
        path: &str,
    ) -> Option<Arc<SheetTable>> {
        self.tables
            .entry(part_key(entries, path))
            .or_insert_with(|| {
                match XlsxDocument::read_xml_entry(archive, entries, path)
                    .map_err(|e| e.to_string())
                    .and_then(|d| worksheet::parse_table_part(&d).map_err(|e| e.to_string()))
                {
                    Ok(t) => Some(Arc::new(t)),
                    Err(e) => {
                        log::warn!("xlsx: table part {path} skipped: {e}");
                        None
                    },
                }
            })
            .clone()
    }
}

/// Read and parse the drawing part at `drawing_path`: its `<xdr:pic>` and
/// `<xdr:sp>` anchors, and each picture's underlying media bytes.
fn read_drawing<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    entries: &opc::ZipEntryIndex,
    media: &mut MediaCache,
    drawing_path: &str,
    unreadable: &mut UnreadableParts,
) -> Option<SharedDrawing> {
    let drawing_xml = match XlsxDocument::read_xml_entry(archive, entries, drawing_path) {
        Ok(d) => d,
        Err(e) => {
            log::debug!("XlsxDocument: drawing part {drawing_path} unreadable ({e}); skipping");
            return None;
        },
    };

    let drawing_rels =
        read_rels_reporting(archive, entries, &sheet_rels_path(drawing_path), unreadable);

    let parsed = match parse_drawing_anchors(&drawing_xml) {
        Ok(a) => a,
        Err(e) => {
            log::debug!(
                "XlsxDocument: drawing {drawing_path} failed to parse ({e}); dropping anchors"
            );
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
        let Some(bytes) = read_media_cached(archive, entries, &media_path, media) else {
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
