//! High-level DOC document API.

use std::io::{Read, Seek};

use crate::cfb::{CfbReader, SummaryProperties, parse_summary_information};

use super::chpx::{parse_chpx_runs, resolve_deleted_cp_ranges_from_runs};
use super::error::{DocError, Result};
use super::fib::Fib;
use super::images::{DocImage, extract_images};
use super::list_format::ListFormatting;
use super::papx::{DocParagraph, build_paragraphs, parse_papx_paragraphs};
use super::piece_table::{
    covers_declared_length, extract_text_range_excluding, parse_clx, sanitize_text,
};

/// A parsed legacy Word document.
#[derive(Debug)]
pub struct DocDocument {
    /// The raw extracted text (after sanitization).
    text: String,
    /// Extracted images from the Data stream.
    images: Vec<DocImage>,
    /// Structured main-text paragraphs with PAP (paragraph property) flags.
    /// Populated only when the FIB advertises a PlcfBtePapx (PAPX FKP index);
    /// empty for very old or minimal files, in which case `doc_to_ir` falls
    /// back to the line-based heuristic on `text`.
    paragraphs: Vec<DocParagraph>,
    /// Text of the subdocuments that follow the main document in the piece
    /// table's character space: footnotes, headers/footers, comments,
    /// endnotes and text boxes. The FIB's `ccp*` lengths that delimit them
    /// were parsed and then never used, so none of this reached any
    /// consumer.
    subdocuments: Vec<SubDocument>,
    /// `true` when the CFB container has a top-level `_VBA_PROJECT`
    /// storage — a cheap macro-presence signal, no VBA interpretation.
    has_macros: bool,
    /// `false` when the piece table has a gap before the FIB's declared
    /// `ccpText` — text in that gap is silently absent from `plain_text()`/
    /// `paragraphs()` with no other signal, so a caller who cares can at
    /// least tell "genuinely short document" apart from "84% missing".
    text_complete: bool,
    /// Title/author/subject/keywords/comments/dates from the
    /// `\x05SummaryInformation` OLE property-set stream every real
    /// `.doc` carries by default — parsed and then never read anywhere
    /// in the crate before.
    summary_properties: Option<SummaryProperties>,
    /// `PlfLst`/`PlfLfo` list definitions, resolving a paragraph's
    /// `(ilfo, ilvl)` to its declared start-at value and number format —
    /// parsed from a FIB pointer that was previously read and then never
    /// used anywhere in the crate.
    list_formatting: ListFormatting,
    /// Comment author names, parsed from `GrpXstAtnOwners`. The FIB
    /// fields locating this array (`fcGrpXstAtnOwners`/
    /// `lcbGrpXstAtnOwners`) were never parsed at all before, so every
    /// `.doc` comment's authorship was unrecoverable.
    comment_authors: Vec<String>,
    /// The first section's header/footer content, parsed from `PlcfHdd`.
    /// Before this, the entire header document collapsed into one
    /// unlabeled, duplicated blob with no way to tell header from
    /// footer, first-page from default, or one section from another.
    header_footer: HeaderFooterStories,
    /// Individual comments, split from the merged Comments substory using
    /// `PlcfandTxt`'s CP boundaries and attributed via `PlcfandRef`'s
    /// `ATRDPre10.ibst` index into `comment_authors`. Empty when either
    /// PLC is absent, malformed, or the two disagree on comment count —
    /// callers fall back to the merged-substory behavior in that case.
    comments: Vec<ParsedComment>,
    /// Embedded OLE objects (Excel workbooks, Equation Editor/MathType,
    /// OLE Package, embedded Word/PowerPoint, etc.) found under the
    /// root's `ObjectPool` storage, identified by presence of a
    /// well-known stream name. Empty when there's no `ObjectPool` at
    /// all.
    ole_objects: Vec<super::ole_objects::EmbeddedOleObject>,
}

/// One of the subdocuments stored after the main text in a `.doc`.
#[derive(Debug, Clone)]
pub struct SubDocument {
    /// Which subdocument this is.
    pub kind: SubDocumentKind,
    /// Sanitised text of the subdocument.
    pub text: String,
}

/// One comment, split out of the merged Comments substory by `PlcfandTxt`
/// and attributed by the matching `PlcfandRef` entry.
#[derive(Debug, Clone)]
pub struct ParsedComment {
    /// Sanitised body text of this one comment.
    pub text: String,
    /// The comment's author, resolved from `GrpXstAtnOwners` via the
    /// matching `ATRDPre10.ibst`. `None` when `ibst` is out of range.
    pub author: Option<String>,
}

/// The first document section's header/footer content, split out of the
/// merged header-document blob using `PlcfHdd`'s story boundaries. Only
/// the first section's 6 stories are captured — matches this crate's DOC
/// model, which builds exactly one `ir::Section` for the whole document.
#[derive(Debug, Clone, Default)]
pub struct HeaderFooterStories {
    /// Even-page header (story 0 of a section's group).
    pub even_header: Option<String>,
    /// Odd-page header — used on every page when even/odd headers
    /// aren't separately enabled (story 1).
    pub odd_header: Option<String>,
    /// Even-page footer (story 2).
    pub even_footer: Option<String>,
    /// Odd-page footer — used on every page when even/odd footers
    /// aren't separately enabled (story 3).
    pub odd_footer: Option<String>,
    /// First-page header (story 4).
    pub first_header: Option<String>,
    /// First-page footer (story 5).
    pub first_footer: Option<String>,
}

impl HeaderFooterStories {
    pub(crate) fn is_empty(&self) -> bool {
        self.even_header.is_none()
            && self.odd_header.is_none()
            && self.even_footer.is_none()
            && self.odd_footer.is_none()
            && self.first_header.is_none()
            && self.first_footer.is_none()
    }
}

/// The subdocument kinds `.doc` stores after the main text, in the fixed
/// order [MS-DOC] §2.5.1 defines for the `ccp*` lengths.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SubDocumentKind {
    /// Footnote bodies (`ccpFtn`).
    Footnotes,
    /// Header and footer bodies (`ccpHdd`).
    HeadersFooters,
    /// Comment bodies (`ccpAtn`).
    Comments,
    /// Endnote bodies (`ccpEdn`).
    Endnotes,
    /// Text-box bodies (`ccpTxbx`).
    TextBoxes,
    /// Header text-box bodies (`ccpHdrTxbx`).
    HeaderTextBoxes,
}

impl DocDocument {
    /// Open a DOC file from a reader.
    pub fn from_reader<R: Read + Seek>(reader: R) -> Result<Self> {
        let mut cfb = CfbReader::new(reader)?;

        let word_doc = cfb
            .open_stream("WordDocument")
            .map_err(|_| DocError::MissingStream("WordDocument stream not found".into()))?;

        // Propagate FIB errors. Swallowing them into an empty document with
        // `Ok` is what made an encrypted file and a Word 6.0/95 file both
        // look like documents that simply had no text.
        let fib = Fib::parse(&word_doc)?;

        // Open the appropriate table stream; try preferred first, then fallback.
        let table_stream = if fib.use_table1 {
            cfb.open_stream("1Table")
                .or_else(|_| cfb.open_stream("0Table"))
        } else {
            cfb.open_stream("0Table")
                .or_else(|_| cfb.open_stream("1Table"))
        };
        // Each of the three failures below used to return an empty document
        // with `Ok`, which is indistinguishable from a document that has no
        // text. A file that cannot be read must say so.
        let table_stream = table_stream.map_err(|_| {
            DocError::MissingStream("neither 0Table nor 1Table stream is present".into())
        })?;

        // Extract CLX from the table stream.
        let clx_start = fib.clx_offset as usize;
        let clx_end = clx_start + fib.clx_size as usize;

        if clx_start >= table_stream.len()
            || clx_size_zero_or_oob(fib.clx_size, clx_start, table_stream.len())
        {
            return Err(DocError::InvalidPieceTable(format!(
                "CLX at offset {clx_start} size {} is outside the {}-byte table stream",
                fib.clx_size,
                table_stream.len()
            )));
        }

        let clx_end = clx_end.min(table_stream.len());
        let clx_data = &table_stream[clx_start..clx_end];
        let pieces = parse_clx(clx_data)?;
        let text_complete = covers_declared_length(&pieces, fib.text_len);

        // The CHPX FKP is parsed once here (rather than separately by each
        // consumer) since both the deleted-revision-mark filter below and
        // `build_paragraphs`'s per-run character formatting
        // need it, and a CHPX FKP walk is not cheap to repeat.
        let chpx_runs = if fib.fc_plcf_bte_chpx != 0 && fib.lcb_plcf_bte_chpx != 0 {
            parse_chpx_runs(&word_doc, &table_stream, fib.fc_plcf_bte_chpx, fib.lcb_plcf_bte_chpx)
        } else {
            Vec::new()
        };

        // Deleted revision-mark text (`sprmCFRMarkDel`) is excluded from the
        // main flat text up front, at extraction time — the same "accepted
        // view" policy already applied to DOCX's `w:del`.
        // Structured paragraph text (`paragraphs()`, used by `doc_to_ir`) is
        // left unfiltered: splicing deletions out of a multi-run paragraph
        // while preserving field-code (`HYPERLINK`) boundaries, on top of
        // the per-run character formatting `build_paragraphs` now also
        // carries, is more than this fix attempts — it stays a
        // deliberately separate, still-open piece of revision-mark handling.
        let deleted_ranges = resolve_deleted_cp_ranges_from_runs(&chpx_runs, &pieces, fib.text_len);
        let raw_text = extract_text_range_excluding(
            &word_doc,
            &pieces,
            0,
            fib.text_len,
            fib.lid,
            &deleted_ranges,
        );
        let text = sanitize_text(&raw_text);

        // Needed inside the loop below to resolve each comment's author by
        // `ibst` index — parsed here (it only needs `table_stream`/`fib`,
        // not anything the loop computes) rather than after it.
        let comment_authors = parse_grp_xst_atn_owners(
            &table_stream,
            fib.fc_grp_xst_atn_owners,
            fib.lcb_grp_xst_atn_owners,
        );

        // The subdocuments follow the main text contiguously in the piece
        // table's character space, each delimited by its own `ccp*` length.
        let mut subdocuments = Vec::new();
        let mut header_footer = HeaderFooterStories::default();
        let mut comments = Vec::new();
        let mut cp = fib.text_len;
        for (kind, len) in [
            (SubDocumentKind::Footnotes, fib.footnote_len),
            (SubDocumentKind::HeadersFooters, fib.header_len),
            (SubDocumentKind::Comments, fib.comment_len),
            (SubDocumentKind::Endnotes, fib.endnote_len),
            (SubDocumentKind::TextBoxes, fib.textbox_len),
            (SubDocumentKind::HeaderTextBoxes, fib.header_textbox_len),
        ] {
            if len == 0 {
                continue;
            }
            let end = cp.saturating_add(len);
            let raw = super::piece_table::extract_text_range(&word_doc, &pieces, cp, end, fib.lid);
            // The header document opens with six fixed stories — the
            // footnote/endnote separators and "continued" notices Word
            // shows only when a note spans pages ([MS-DOC] "Headers").
            // They are not header text; keep them out of the flat text
            // the way the split below already keeps them out of the IR.
            let mut text_start = 0;
            if kind == SubDocumentKind::HeadersFooters {
                header_footer =
                    parse_plcf_hdd_stories(&table_stream, &raw, fib.fc_plcf_hdd, fib.lcb_plcf_hdd);
                text_start =
                    plcf_hdd_first_section_cp(&table_stream, fib.fc_plcf_hdd, fib.lcb_plcf_hdd)
                        .unwrap_or(0);
            }
            if kind == SubDocumentKind::Comments {
                comments = parse_comments(
                    &table_stream,
                    &raw,
                    fib.fc_plcf_and_txt,
                    fib.lcb_plcf_and_txt,
                    fib.fc_plcf_and_ref,
                    fib.lcb_plcf_and_ref,
                    &comment_authors,
                );
            }
            let sub = if text_start > 0 {
                sanitize_text(&raw.chars().skip(text_start).collect::<String>())
            } else {
                sanitize_text(&raw)
            };
            if !sub.trim().is_empty() {
                subdocuments.push(SubDocument { kind, text: sub });
            }
            cp = end;
        }

        // Build structured paragraphs (with table / list PAP flags) from the
        // PAPX FKP, when the FIB advertises one. Without it we cannot detect
        // tables or lists, so `doc_to_ir` falls back to the line heuristic.
        let paragraphs = if fib.fc_plcf_bte_papx != 0 && fib.lcb_plcf_bte_papx != 0 {
            let fkp = parse_papx_paragraphs(
                &word_doc,
                &table_stream,
                fib.fc_plcf_bte_papx,
                fib.lcb_plcf_bte_papx,
            );
            build_paragraphs(&word_doc, &pieces, &fkp, fib.text_len, fib.lid, &chpx_runs)
        } else {
            Vec::new()
        };

        let list_formatting = ListFormatting::parse(
            &table_stream,
            fib.fc_plcf_lst,
            fib.lcb_plcf_lst,
            fib.fc_plf_lfo,
            fib.lcb_plf_lfo,
        );

        // Extract images from the Data stream (if present).
        let images = match cfb.open_stream("Data") {
            Ok(data_stream) => extract_images(&data_stream),
            Err(_) => Vec::new(),
        };
        let has_macros = cfb.has_root_entry("_VBA_PROJECT");
        // At minimum, recognize an embedded OLE object exists and
        // surface its identity — before this, `ObjectPool` was never
        // traversed at all, so an embedded Excel workbook, Equation
        // Editor object, etc. left no trace anywhere.
        let ole_objects = super::ole_objects::extract_ole_objects(&cfb);
        let summary_properties = cfb
            .open_stream("\u{5}SummaryInformation")
            .ok()
            .and_then(|data| parse_summary_information(&data));

        Ok(Self {
            text,
            images,
            paragraphs,
            subdocuments,
            has_macros,
            text_complete,
            summary_properties,
            list_formatting,
            comment_authors,
            header_footer,
            comments,
            ole_objects,
        })
    }

    /// Open a DOC file from a path.
    pub fn open<P: AsRef<std::path::Path>>(path: P) -> Result<Self> {
        let file = std::fs::File::open(path)?;
        Self::from_reader(file)
    }

    /// Get all extracted images.
    pub fn images(&self) -> &[DocImage] {
        &self.images
    }

    /// Footnotes, headers, comments, endnotes and text boxes, in the order
    /// the file stores them.
    pub fn subdocuments(&self) -> &[SubDocument] {
        &self.subdocuments
    }

    /// Comment author names declared by `GrpXstAtnOwners`, in file order.
    /// Empty when the document has no comments. There is no per-comment
    /// author correlation here (that needs `PlcfAtn`/`ATRD`, tracked
    /// separately) — when this holds exactly one name, every comment in
    /// the document was written by that single author.
    pub(crate) fn comment_authors(&self) -> &[String] {
        &self.comment_authors
    }

    /// The first section's header/footer content, split from `PlcfHdd`.
    /// All fields are `None` when the document has no header document at
    /// all, or its `PlcfHdd` doesn't cover a full section's worth of
    /// stories.
    pub(crate) fn header_footer(&self) -> &HeaderFooterStories {
        &self.header_footer
    }

    /// Individual comments split from the merged Comments substory, each
    /// with its own resolved author. Empty when `PlcfandTxt`/`PlcfandRef`
    /// couldn't be parsed or disagreed on comment count — callers must
    /// fall back to the merged-substory behavior in that case.
    pub(crate) fn comments(&self) -> &[ParsedComment] {
        &self.comments
    }

    /// Embedded OLE objects found under `ObjectPool`, identified by
    /// presence of a well-known stream name. Empty when there's no
    /// `ObjectPool` at all, or it's empty/unrecognizable.
    pub(crate) fn ole_objects(&self) -> &[super::ole_objects::EmbeddedOleObject] {
        &self.ole_objects
    }

    /// `true` when the file carries a `_VBA_PROJECT` storage — a cheap
    /// macro-presence signal, no VBA interpretation.
    pub fn has_macros(&self) -> bool {
        self.has_macros
    }

    /// `false` when the piece table has a gap before the FIB's declared
    /// text length, meaning `plain_text()`/`paragraphs()` are missing real
    /// content that could not be safely recovered.
    pub fn text_complete(&self) -> bool {
        self.text_complete
    }

    /// Title/author/subject/keywords/comments/dates from the file's
    /// `\x05SummaryInformation` OLE property set, when present and
    /// well-formed.
    pub fn summary_properties(&self) -> Option<&crate::cfb::SummaryProperties> {
        self.summary_properties.as_ref()
    }

    /// `PlfLst`/`PlfLfo` list definitions, resolving a paragraph's
    /// `(ilfo, ilvl)` to its declared start-at value and number format.
    pub(crate) fn list_formatting(&self) -> &ListFormatting {
        &self.list_formatting
    }

    /// Get the extracted plain text.
    ///
    /// Includes footnote/endnote/comment/textbox bodies — `to_ir()` (via
    /// `doc_to_ir`) already carries this content as its own elements, and
    /// leaving it out here made this renderer disagree with that one, the
    /// same gap already fixed for DOCX.
    pub fn plain_text(&self) -> String {
        let mut out = self.text.clone();
        for sub in &self.subdocuments {
            let text = sub.text.trim();
            if text.is_empty() {
                continue;
            }
            if !out.is_empty() && !out.ends_with('\n') {
                out.push('\n');
            }
            out.push_str(text);
            out.push('\n');
        }
        out
    }

    /// Get a reference to the extracted plain text.
    pub fn plain_text_ref(&self) -> &str {
        &self.text
    }

    /// Structured main-text paragraphs with PAP flags (table / list).
    ///
    /// Empty when the document has no PAPX FKP, in which case callers fall
    /// back to the line-based heuristic on [`Self::plain_text_ref`].
    pub(crate) fn paragraphs(&self) -> &[DocParagraph] {
        &self.paragraphs
    }

    /// Convert to markdown (basic: paragraphs separated by blank lines).
    ///
    /// Includes footnote/endnote/comment/textbox bodies — see the
    /// identical note on `plain_text()`.
    pub fn to_markdown(&self) -> String {
        let mut result = text_to_markdown_blocks(&self.text);
        for sub in &self.subdocuments {
            let block = text_to_markdown_blocks(&sub.text);
            let block = block.trim();
            if block.is_empty() {
                continue;
            }
            if !result.ends_with("\n\n") && !result.is_empty() {
                result.push('\n');
            }
            result.push_str(block);
            result.push_str("\n\n");
        }
        result
    }
}

/// Split `text` into markdown paragraphs, each separated by a blank line —
/// the shared rendering shape `to_markdown()` uses for both the main text
/// and every subdocument body.
fn text_to_markdown_blocks(text: &str) -> String {
    let mut result = String::new();
    let mut prev_empty = false;

    for line in text.lines() {
        let trimmed = line.trim();
        if trimmed.is_empty() {
            if !prev_empty {
                result.push('\n');
            }
            prev_empty = true;
        } else {
            result.push_str(trimmed);
            result.push_str("\n\n");
            prev_empty = false;
        }
    }

    result
}

fn clx_size_zero_or_oob(clx_size: u32, clx_start: usize, stream_len: usize) -> bool {
    clx_size == 0 || clx_start + clx_size as usize > stream_len + 1024 // allow some slack
}

/// Parse `GrpXstAtnOwners`: an array of XSTs (comment author names) packed
/// back-to-back at `fc` for `lcb` bytes in the Table stream. Each entry is
/// a `u16` character count `cch` followed by `cch` UTF-16LE code units —
/// no STTBF-style count/extra-data header, per [MS-DOC] §2.5.5's
/// description of `fcGrpXstAtnOwners`.
fn parse_grp_xst_atn_owners(table_stream: &[u8], fc: u32, lcb: u32) -> Vec<String> {
    if lcb == 0 {
        return Vec::new();
    }
    let start = fc as usize;
    let end = start.saturating_add(lcb as usize).min(table_stream.len());
    if start >= end {
        return Vec::new();
    }

    let mut names = Vec::new();
    let mut pos = start;
    while pos + 2 <= end {
        let cch = u16::from_le_bytes([table_stream[pos], table_stream[pos + 1]]) as usize;
        pos += 2;
        let byte_len = cch * 2;
        if pos + byte_len > end {
            break;
        }
        let units: Vec<u16> = table_stream[pos..pos + byte_len]
            .as_chunks::<2>()
            .0
            .iter()
            .map(|c| u16::from_le_bytes(*c))
            .collect();
        pos += byte_len;
        names.push(String::from_utf16_lossy(&units));
    }
    names
}

/// Parse `PlcfHdd` (the header/footer story delimiter PLC) and split
/// `raw_header_text` — the header document's own raw (pre-sanitize) text,
/// where char index `i` corresponds to CP `i` within the header document
/// itself — into the first section's 6 stories.
///
/// `PlcfHdd` is a PLC with no data elements, just an array of CPs
/// (`aCP`). If the header document exists, its first 6 stories are fixed
/// footnote/endnote separators; every subsequent group of 6 stories
/// belongs to one document section, in order: even header, odd header,
/// even footer, odd footer, first-page header, first-page footer
/// ([MS-DOC] "Headers"). `aCP` has `n + 2` entries for `n` stories: the
/// first `n` mark story starts, entry `n` marks the end of the last
/// story (`== ccpHdd - 1`), and the final entry is undefined/ignored —
/// so `aCP[0..=n]` (`n + 1` values) are the usual PLC boundary CPs for
/// `n` elements, per [MS-DOC] "Plcfhdd". This crate models only one
/// `ir::Section` for the whole document, so only the first section's
/// group (stories 6..12) is extracted.
fn parse_plcf_hdd_stories(
    table_stream: &[u8],
    raw_header_text: &str,
    fc: u32,
    lcb: u32,
) -> HeaderFooterStories {
    let mut result = HeaderFooterStories::default();
    let Some(cps) = plcf_hdd_cps(table_stream, fc, lcb) else {
        return result;
    };

    let char_range = |lo: usize, hi: usize| -> Option<String> {
        if hi <= lo {
            return None;
        }
        let text: String = raw_header_text.chars().skip(lo).take(hi - lo).collect();
        let sanitized = sanitize_text(&text);
        let trimmed = sanitized.trim();
        if trimmed.is_empty() {
            None
        } else {
            Some(trimmed.to_string())
        }
    };

    // First section's group starts right after the 6 fixed separator
    // stories, at story index 6.
    let base = 6;
    result.even_header = char_range(cps[base], cps[base + 1]);
    result.odd_header = char_range(cps[base + 1], cps[base + 2]);
    result.even_footer = char_range(cps[base + 2], cps[base + 3]);
    result.odd_footer = char_range(cps[base + 3], cps[base + 4]);
    result.first_header = char_range(cps[base + 4], cps[base + 5]);
    result.first_footer = char_range(cps[base + 5], cps[base + 6]);
    result
}

/// The `aCP[0..=n]` boundary CPs of `PlcfHdd`, or `None` when the PLC is
/// absent, malformed, or shorter than the 6 fixed separator stories plus
/// one full section's group of 6 header/footer stories.
fn plcf_hdd_cps(table_stream: &[u8], fc: u32, lcb: u32) -> Option<Vec<usize>> {
    if lcb == 0 {
        return None;
    }
    let start = fc as usize;
    let end = start.saturating_add(lcb as usize).min(table_stream.len());
    if start >= end || !(end - start).is_multiple_of(4) {
        return None;
    }
    let total_cps = (end - start) / 4;
    if total_cps < 2 {
        return None;
    }
    let n = total_cps - 2;
    if n < 12 {
        return None;
    }
    Some(
        (0..=n)
            .map(|i| {
                u32::from_le_bytes([
                    table_stream[start + i * 4],
                    table_stream[start + i * 4 + 1],
                    table_stream[start + i * 4 + 2],
                    table_stream[start + i * 4 + 3],
                ]) as usize
            })
            .collect(),
    )
}

/// The header-document CP where the first section's own stories begin —
/// i.e. the end of the 6 fixed footnote/endnote separator stories.
fn plcf_hdd_first_section_cp(table_stream: &[u8], fc: u32, lcb: u32) -> Option<usize> {
    plcf_hdd_cps(table_stream, fc, lcb).map(|cps| cps[6])
}

/// Split the merged Comments substory into individual comments using
/// `PlcfandTxt`'s CP boundaries, and attribute each to an author using
/// the matching `PlcfandRef`/`ATRDPre10.ibst` entry. Returns an empty
/// `Vec` (meaning "fall back to the merged substory") whenever either
/// PLC is absent/malformed, or the two disagree on comment count — that
/// mismatch means the file doesn't match the fixed-size assumptions
/// here closely enough to trust a per-comment split.
fn parse_comments(
    table_stream: &[u8],
    raw_comment_text: &str,
    fc_txt: u32,
    lcb_txt: u32,
    fc_ref: u32,
    lcb_ref: u32,
    authors: &[String],
) -> Vec<ParsedComment> {
    let Some(bodies) = parse_plcf_and_txt_ranges(table_stream, raw_comment_text, fc_txt, lcb_txt)
    else {
        return Vec::new();
    };
    let ibsts = parse_plcf_and_ref_ibsts(table_stream, fc_ref, lcb_ref);
    if bodies.len() != ibsts.len() {
        return Vec::new();
    }
    bodies
        .into_iter()
        .zip(ibsts)
        .filter_map(|(text, ibst)| {
            if text.is_empty() {
                return None;
            }
            Some(ParsedComment {
                text,
                author: authors.get(ibst).cloned(),
            })
        })
        .collect()
}

/// `PlcfandTxt` — a PLC of pure CPs (no data elements) whose `aCP` has
/// `n + 2` entries for `n` comment-text ranges within the Comments
/// substory's own character space: `aCP[0..=n]` (`n + 1` values) are the
/// usual PLC boundary CPs, and the final entry is undefined/ignored, per
/// [MS-DOC] "PlcfandTxt". Each range's leading `0x0005` reference-mark
/// character is stripped before returning.
fn parse_plcf_and_txt_ranges(
    table_stream: &[u8],
    raw_comment_text: &str,
    fc: u32,
    lcb: u32,
) -> Option<Vec<String>> {
    if lcb == 0 {
        return None;
    }
    let start = fc as usize;
    let end = start.saturating_add(lcb as usize).min(table_stream.len());
    if start >= end || !(end - start).is_multiple_of(4) {
        return None;
    }
    let total_cps = (end - start) / 4;
    if total_cps < 2 {
        return None;
    }
    let n = total_cps - 2;
    if n == 0 {
        return None;
    }

    let cps: Vec<usize> = (0..=n)
        .map(|i| {
            u32::from_le_bytes([
                table_stream[start + i * 4],
                table_stream[start + i * 4 + 1],
                table_stream[start + i * 4 + 2],
                table_stream[start + i * 4 + 3],
            ]) as usize
        })
        .collect();

    Some(
        (0..n)
            .map(|i| {
                let (lo, hi) = (cps[i], cps[i + 1]);
                if hi <= lo {
                    return String::new();
                }
                let text: String = raw_comment_text.chars().skip(lo).take(hi - lo).collect();
                let text = text.strip_prefix('\u{5}').unwrap_or(&text);
                sanitize_text(text).trim().to_string()
            })
            .collect(),
    )
}

/// `PlcfandRef`'s data elements are `ATRDPre10` structures (30 bytes
/// each: a 20-byte `xstUsrInitl`, then a 2-byte `ibst` index into
/// `GrpXstAtnOwners`, then 8 bytes this crate doesn't need). Returns one
/// `ibst` per comment, in the same document order as `PlcfandTxt`'s
/// ranges (both PLCs are populated in comment-creation order, per
/// [MS-DOC] "PlcfandTxt"'s own cross-reference to "PlcfandRef").
fn parse_plcf_and_ref_ibsts(table_stream: &[u8], fc: u32, lcb: u32) -> Vec<usize> {
    const ATRD_PRE10_SIZE: usize = 30;
    if lcb == 0 {
        return Vec::new();
    }
    let start = fc as usize;
    let end = start.saturating_add(lcb as usize);
    if end > table_stream.len() || start >= end {
        return Vec::new();
    }
    let cb = end - start;
    if cb < 4 {
        return Vec::new();
    }
    // PLC element count: iMac = (cb - 4) / (4 + cbStruct).
    let n = (cb - 4) / (4 + ATRD_PRE10_SIZE);
    let data_start = start + 4 * (n + 1);
    (0..n)
        .filter_map(|i| {
            let atrd_start = data_start + i * ATRD_PRE10_SIZE;
            let ibst_offset = atrd_start + 20;
            if ibst_offset + 2 > table_stream.len() {
                return None;
            }
            Some(u16::from_le_bytes([table_stream[ibst_offset], table_stream[ibst_offset + 1]])
                as usize)
        })
        .collect()
}

impl crate::core::OfficeDocument for DocDocument {
    fn plain_text(&self) -> String {
        self.plain_text()
    }

    fn to_markdown(&self) -> String {
        self.to_markdown()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_markdown_double_spacing() {
        let doc = DocDocument {
            subdocuments: Vec::new(),
            has_macros: false,
            text_complete: true,
            summary_properties: None,
            list_formatting: crate::doc::list_format::ListFormatting::default(),
            comment_authors: Vec::new(),
            comments: Vec::new(),
            ole_objects: Vec::new(),
            header_footer: HeaderFooterStories::default(),
            images: Vec::new(),
            text: "First paragraph\nSecond paragraph\n\nAfter gap".into(),
            paragraphs: Vec::new(),
        };
        let md = doc.to_markdown();
        assert!(md.contains("First paragraph\n\n"));
        assert!(md.contains("Second paragraph\n\n"));
        assert!(md.contains("After gap\n\n"));
    }

    #[test]
    fn test_plain_text_access() {
        let doc = DocDocument {
            subdocuments: Vec::new(),
            has_macros: false,
            text_complete: true,
            summary_properties: None,
            list_formatting: crate::doc::list_format::ListFormatting::default(),
            comment_authors: Vec::new(),
            comments: Vec::new(),
            ole_objects: Vec::new(),
            header_footer: HeaderFooterStories::default(),
            images: Vec::new(),
            text: "Hello World".into(),
            paragraphs: Vec::new(),
        };
        assert_eq!(doc.plain_text(), "Hello World");
    }

    /// `plain_text()`/`to_markdown()` only ever walked
    /// `self.text` (the main body), never `self.subdocuments`, so a
    /// footnote/endnote/comment/textbox-only document silently vanished
    /// from both renderers even though `to_ir()` (via `doc_to_ir`) already
    /// carried the content correctly.
    #[test]
    fn test_plain_text_and_markdown_include_subdocument_bodies() {
        let doc = DocDocument {
            subdocuments: vec![
                SubDocument {
                    kind: SubDocumentKind::Footnotes,
                    text: "FOOTNOTE ONE".into(),
                },
                SubDocument {
                    kind: SubDocumentKind::Comments,
                    text: "REVIEW NOTE".into(),
                },
                SubDocument {
                    kind: SubDocumentKind::HeaderTextBoxes,
                    text: "SIDEBAR".into(),
                },
            ],
            has_macros: false,
            text_complete: true,
            summary_properties: None,
            list_formatting: crate::doc::list_format::ListFormatting::default(),
            comment_authors: Vec::new(),
            comments: Vec::new(),
            ole_objects: Vec::new(),
            header_footer: HeaderFooterStories::default(),
            images: Vec::new(),
            text: "Main body text".into(),
            paragraphs: Vec::new(),
        };
        let text = doc.plain_text();
        for token in ["Main body text", "FOOTNOTE ONE", "REVIEW NOTE", "SIDEBAR"] {
            assert!(text.contains(token), "{token} missing from plain_text(): {text:?}");
        }
        let md = doc.to_markdown();
        for token in ["Main body text", "FOOTNOTE ONE", "REVIEW NOTE", "SIDEBAR"] {
            assert!(md.contains(token), "{token} missing from to_markdown(): {md:?}");
        }
    }

    /// An empty subdocument body must contribute nothing — no stray blank
    /// paragraphs or extra separators.
    #[test]
    fn test_empty_subdocuments_are_skipped_in_both_renderers() {
        let doc = DocDocument {
            subdocuments: vec![SubDocument {
                kind: SubDocumentKind::Comments,
                text: "  \n ".into(),
            }],
            has_macros: false,
            text_complete: true,
            summary_properties: None,
            list_formatting: crate::doc::list_format::ListFormatting::default(),
            comment_authors: Vec::new(),
            comments: Vec::new(),
            ole_objects: Vec::new(),
            header_footer: HeaderFooterStories::default(),
            images: Vec::new(),
            text: "Body".into(),
            paragraphs: Vec::new(),
        };
        assert_eq!(doc.plain_text(), "Body");
        assert_eq!(doc.to_markdown(), "Body\n\n");
    }

    /// `text_complete()` reaches `to_ir()`'s `Metadata::text_truncated` so
    /// a caller who never inspects `DocDocument` directly can still tell a
    /// piece-table gap apart from a genuinely short document.
    #[test]
    fn test_incomplete_text_reaches_metadata_as_truncated() {
        let doc = DocDocument {
            subdocuments: Vec::new(),
            has_macros: false,
            text_complete: false,
            summary_properties: None,
            list_formatting: crate::doc::list_format::ListFormatting::default(),
            comment_authors: Vec::new(),
            comments: Vec::new(),
            ole_objects: Vec::new(),
            header_footer: HeaderFooterStories::default(),
            images: Vec::new(),
            text: "only the recovered fragment".into(),
            paragraphs: Vec::new(),
        };
        assert!(!doc.text_complete());
        let ir = crate::convert_doc::doc_to_ir(&doc);
        assert!(
            ir.metadata.text_truncated,
            "a piece-table gap must be visible on Metadata::text_truncated"
        );
    }

    /// The common case: a complete piece table must not be flagged.
    #[test]
    fn test_complete_text_is_not_flagged_truncated() {
        let doc = make_doc("Hello World");
        assert!(doc.text_complete());
        let ir = crate::convert_doc::doc_to_ir(&doc);
        assert!(!ir.metadata.text_truncated);
    }

    fn make_doc(text: &str) -> DocDocument {
        DocDocument {
            subdocuments: Vec::new(),
            has_macros: false,
            text_complete: true,
            summary_properties: None,
            list_formatting: crate::doc::list_format::ListFormatting::default(),
            comment_authors: Vec::new(),
            comments: Vec::new(),
            ole_objects: Vec::new(),
            header_footer: HeaderFooterStories::default(),
            images: Vec::new(),
            text: text.to_string(),
            paragraphs: Vec::new(),
        }
    }

    /// Build a `DocDocument` whose IR comes from structured paragraphs
    /// (the PAPX path) rather than the line heuristic. Used to TDD the
    /// table / list walkers without a binary `.doc` fixture.
    fn make_doc_with_paragraphs(paras: Vec<DocParagraph>) -> DocDocument {
        DocDocument {
            subdocuments: Vec::new(),
            has_macros: false,
            text_complete: true,
            summary_properties: None,
            list_formatting: crate::doc::list_format::ListFormatting::default(),
            comment_authors: Vec::new(),
            comments: Vec::new(),
            ole_objects: Vec::new(),
            header_footer: HeaderFooterStories::default(),
            images: Vec::new(),
            text: String::new(),
            paragraphs: paras,
        }
    }

    /// Construct a paragraph with the given PAP flags and terminator.
    fn pap(text: &str, props: crate::doc::sprm::PapProps) -> DocParagraph {
        DocParagraph {
            text: text.to_string(),
            terminator: '\r',
            props,
            hyperlinks: Vec::new(),
            chp_runs: Vec::new(),
        }
    }

    fn list_props(level: u8) -> crate::doc::sprm::PapProps {
        crate::doc::sprm::PapProps {
            ilvl: Some(level),
            // A real list item carries a valid `ilfo` (0x0001–0x07FE); without
            // it the paragraph is not in a list per [MS-DOC] §2.4.6.3.
            ilfo: Some(1),
            ..Default::default()
        }
    }

    #[test]
    fn test_ir_list_emits_nested_list_from_ilvl_paragraphs() {
        use crate::ir::Element;
        let doc = make_doc_with_paragraphs(vec![
            pap("Intro.", Default::default()),
            pap("First", list_props(0)),
            pap("Second", list_props(0)),
            pap("Nested", list_props(1)),
            pap("After.", Default::default()),
        ]);
        let ir = crate::convert_doc::doc_to_ir(&doc);
        let elements = &ir.sections[0].elements;

        // [Paragraph, List, Paragraph]
        assert_eq!(elements.len(), 3, "expected intro, list, outro");
        assert!(matches!(elements[0], Element::Paragraph(_)));
        assert!(matches!(elements[2], Element::Paragraph(_)));

        let list = match &elements[1] {
            Element::List(l) => l,
            _ => panic!("expected a List element"),
        };
        assert_eq!(list.items.len(), 2, "two top-level items");
        // Second item nests the level-1 paragraph.
        assert!(list.items[1].nested.is_some(), "second item must nest");
        let nested = list.items[1].nested.as_ref().unwrap();
        assert_eq!(nested.items.len(), 1);
    }

    #[test]
    fn test_ir_consecutive_list_runs_split_on_prose() {
        use crate::ir::Element;
        let doc = make_doc_with_paragraphs(vec![
            pap("A1", list_props(0)),
            pap("A2", list_props(0)),
            pap("gap", Default::default()),
            pap("B1", list_props(0)),
        ]);
        let ir = crate::convert_doc::doc_to_ir(&doc);
        let elements = &ir.sections[0].elements;
        // [List(A), Paragraph(gap), List(B)]
        let lists: Vec<_> = elements
            .iter()
            .filter(|e| matches!(e, Element::List(_)))
            .collect();
        assert_eq!(lists.len(), 2, "the prose gap must split the run");
    }

    #[test]
    fn test_ir_empty_doc_produces_empty_section() {
        let ir = crate::convert_doc::doc_to_ir(&make_doc(""));
        assert!(ir.sections[0].elements.is_empty());
        assert!(ir.metadata.title.is_none());
    }

    #[test]
    fn test_ir_allcaps_first_line_becomes_h1() {
        use crate::ir::Element;
        let ir = crate::convert_doc::doc_to_ir(&make_doc("INTRODUCTION\nSome text here."));
        assert_eq!(ir.metadata.title.as_deref(), Some("INTRODUCTION"));
        assert!(matches!(ir.sections[0].elements[0], Element::Heading(ref h) if h.level == 1));
    }

    #[test]
    fn test_ir_first_short_line_no_punct_becomes_h1() {
        use crate::ir::Element;
        let ir = crate::convert_doc::doc_to_ir(&make_doc("My Document Title\nThis is body text."));
        assert!(matches!(ir.sections[0].elements[0], Element::Heading(ref h) if h.level == 1));
    }

    #[test]
    fn test_ir_allcaps_non_first_line_becomes_h2() {
        use crate::ir::Element;
        let ir = crate::convert_doc::doc_to_ir(&make_doc("Title\nSECTION TWO\nBody text."));
        assert!(matches!(ir.sections[0].elements[1], Element::Heading(ref h) if h.level == 2));
    }

    #[test]
    fn test_ir_line_ending_with_period_becomes_paragraph() {
        use crate::ir::Element;
        let ir = crate::convert_doc::doc_to_ir(&make_doc("This is a sentence."));
        assert!(matches!(ir.sections[0].elements[0], Element::Paragraph(_)));
    }

    #[test]
    fn test_ir_blank_lines_are_skipped() {
        let ir = crate::convert_doc::doc_to_ir(&make_doc("Title\n\n\nText"));
        assert_eq!(ir.sections[0].elements.len(), 2);
    }

    #[test]
    fn test_ir_list_run_with_nonzero_base_level_keeps_every_item() {
        // Regression: `.doc` list levels are not guaranteed to start at 0.
        // Word's `simple-list.doc` fixture writes `ilvl = 1` for a flat list,
        // which used to collapse the run to a single item because
        // `build_nested_list` was called with `base_level = 0`.
        use crate::ir::Element;
        let doc = make_doc_with_paragraphs(vec![
            pap("First", list_props(1)),
            pap("Second", list_props(1)),
            pap("Third", list_props(1)),
        ]);
        let ir = crate::convert_doc::doc_to_ir(&doc);
        let elements = &ir.sections[0].elements;
        assert_eq!(elements.len(), 1, "a single list run");
        let list = match &elements[0] {
            Element::List(l) => l,
            _ => panic!("expected a List element"),
        };
        assert_eq!(list.items.len(), 3, "all three items must survive");
    }

    #[test]
    fn test_ir_format_is_doc() {
        let ir = crate::convert_doc::doc_to_ir(&make_doc("content"));
        assert_eq!(ir.metadata.format, crate::format::DocumentFormat::Doc);
    }

    /// `SummaryInformation` fields must reach `Metadata`, and
    /// the declared title must beat the heading-guess title.
    #[test]
    fn test_ir_summary_properties_reach_metadata() {
        let mut doc = make_doc("SOME ALL-CAPS HEADING\nBody text follows.");
        doc.summary_properties = Some(SummaryProperties {
            title: Some("Declared Title".to_string()),
            subject: Some("Declared Subject".to_string()),
            author: Some("Declared Author".to_string()),
            keywords: Some("alpha, beta".to_string()),
            comments: Some("Declared Comment".to_string()),
            created: Some("2020-01-02T03:04:05Z".to_string()),
            modified: Some("2021-06-07T08:09:10Z".to_string()),
        });
        let ir = crate::convert_doc::doc_to_ir(&doc);
        assert_eq!(ir.metadata.title.as_deref(), Some("Declared Title"));
        assert_eq!(ir.metadata.author.as_deref(), Some("Declared Author"));
        assert_eq!(ir.metadata.subject.as_deref(), Some("Declared Subject"));
        assert_eq!(ir.metadata.keywords, vec!["alpha".to_string(), "beta".to_string()]);
        assert_eq!(ir.metadata.description.as_deref(), Some("Declared Comment"));
        assert_eq!(ir.metadata.created.as_deref(), Some("2020-01-02T03:04:05Z"));
        assert_eq!(ir.metadata.modified.as_deref(), Some("2021-06-07T08:09:10Z"));
    }

    /// A missing/empty title in `SummaryInformation` must not shadow the
    /// heading-guess fallback (metadata must not regress the heading guess).
    #[test]
    fn test_ir_empty_summary_title_falls_back_to_heading_guess() {
        let mut doc = make_doc("A HEADING LINE\nBody text follows.");
        doc.summary_properties = Some(SummaryProperties {
            title: Some(String::new()),
            ..Default::default()
        });
        let ir = crate::convert_doc::doc_to_ir(&doc);
        assert_eq!(ir.metadata.title.as_deref(), Some("A HEADING LINE"));
    }

    /// `GrpXstAtnOwners` is an array of XSTs packed
    /// back-to-back with no STTBF-style header: each entry is a `cch`
    /// `u16` followed by `cch` UTF-16LE code units.
    #[test]
    fn test_grp_xst_atn_owners_parses_multiple_packed_entries() {
        let mut data = Vec::new();
        for name in ["Michael McCandless", "Miklos Vajna"] {
            let units: Vec<u16> = name.encode_utf16().collect();
            data.extend_from_slice(&(units.len() as u16).to_le_bytes());
            for u in units {
                data.extend_from_slice(&u.to_le_bytes());
            }
        }
        let names = parse_grp_xst_atn_owners(&data, 0, data.len() as u32);
        assert_eq!(names, vec!["Michael McCandless".to_string(), "Miklos Vajna".to_string()]);
    }

    #[test]
    fn test_grp_xst_atn_owners_zero_length_is_empty() {
        let data = vec![0u8; 32];
        assert!(parse_grp_xst_atn_owners(&data, 4, 0).is_empty());
    }

    /// A single declared comment author unambiguously
    /// attributes every comment in the document; `doc_to_ir` must carry
    /// it onto the merged Comments `Note` via the new `author` field.
    #[test]
    fn test_a_single_comment_author_reaches_the_comments_note() {
        use crate::ir::Element;
        let mut doc = make_doc("Body text.");
        doc.subdocuments = vec![SubDocument {
            kind: SubDocumentKind::Comments,
            text: "Here is a comment".into(),
        }];
        doc.comment_authors = vec!["Michael McCandless".to_string()];

        let ir = crate::convert_doc::doc_to_ir(&doc);
        let note = ir.sections[0]
            .elements
            .iter()
            .find_map(|e| {
                if let Element::Endnote(n) = e {
                    Some(n)
                } else {
                    None
                }
            })
            .expect("expected a comments Endnote");
        assert_eq!(note.author.as_deref(), Some("Michael McCandless"));
    }

    /// Two or more declared authors cannot be attributed to this merged,
    /// per-subdocument (not per-comment) `Note` without `PlcfAtn`/`ATRD`
    /// correlation — leave `author` unset rather than guess wrong.
    #[test]
    fn test_multiple_comment_authors_leave_the_note_author_unset() {
        use crate::ir::Element;
        let mut doc = make_doc("Body text.");
        doc.subdocuments = vec![SubDocument {
            kind: SubDocumentKind::Comments,
            text: "Inner\nOuter".into(),
        }];
        doc.comment_authors = vec!["vmiklos".to_string(), "Miklos Vajna".to_string()];

        let ir = crate::convert_doc::doc_to_ir(&doc);
        let note = ir.sections[0]
            .elements
            .iter()
            .find_map(|e| {
                if let Element::Endnote(n) = e {
                    Some(n)
                } else {
                    None
                }
            })
            .expect("expected a comments Endnote");
        assert_eq!(note.author, None);
    }

    /// Every footnote in a document used to collapse into a
    /// single `Element::Footnote` holding all footnotes concatenated.
    /// Each footnote is self-delimited in its own substory text by a
    /// leading `\u{2}` (auto-number reference-mark) character — verified
    /// present in 49/51 real-corpus files with footnotes, 5/5 with
    /// endnotes, and absent in all 12 real-corpus files with comments
    /// (comments' reference point lives only in the main text via
    /// `PlcfAtn`, not duplicated into the substory).
    #[test]
    fn test_footnotes_split_into_one_element_per_reference_mark() {
        use crate::ir::{Element, InlineContent, Note};
        let mut doc = make_doc("Body text.");
        doc.subdocuments = vec![SubDocument {
            kind: SubDocumentKind::Footnotes,
            text: "\u{2} First footnote.\n\u{2} Second footnote.\n\u{2} Third footnote.\n".into(),
        }];

        let ir = crate::convert_doc::doc_to_ir(&doc);
        let footnotes: Vec<&Note> = ir.sections[0]
            .elements
            .iter()
            .filter_map(|e| {
                if let Element::Footnote(n) = e {
                    Some(n)
                } else {
                    None
                }
            })
            .collect();
        assert_eq!(footnotes.len(), 3, "expected one Footnote element per reference mark");
        let text_of = |n: &Note| -> String {
            n.content
                .iter()
                .filter_map(|e| match e {
                    Element::Paragraph(p) => p.content.first().map(|c| match c {
                        InlineContent::Text(t) => t.text.clone(),
                        _ => String::new(),
                    }),
                    _ => None,
                })
                .collect::<Vec<_>>()
                .join(" ")
        };
        assert_eq!(text_of(footnotes[0]), "First footnote.");
        assert_eq!(text_of(footnotes[1]), "Second footnote.");
        assert_eq!(text_of(footnotes[2]), "Third footnote.");
        assert_eq!(footnotes[0].id, 0);
        assert_eq!(footnotes[1].id, 1);
        assert_eq!(footnotes[2].id, 2);
    }

    /// Comments never carry the `\u{2}` marker in their own substory (see
    /// above) — real per-comment splitting comes from `PlcfandTxt`
    /// instead (tested separately). When that PLC is absent
    /// (as here, with `doc.comments()` at its default empty `Vec`), the
    /// old merged-into-one-Note behavior is the only safe fallback.
    #[test]
    fn test_comments_stay_merged_into_a_single_note() {
        use crate::ir::{Element, Note};
        let mut doc = make_doc("Body text.");
        doc.subdocuments = vec![SubDocument {
            kind: SubDocumentKind::Comments,
            text: "First comment.\nSecond comment.\n".into(),
        }];

        let ir = crate::convert_doc::doc_to_ir(&doc);
        let comments: Vec<&Note> = ir.sections[0]
            .elements
            .iter()
            .filter_map(|e| {
                if let Element::Endnote(n) = e {
                    Some(n)
                } else {
                    None
                }
            })
            .collect();
        assert_eq!(comments.len(), 1, "comments must stay merged until PlcfAtn is parsed");
        assert_eq!(comments[0].content.len(), 2, "both lines must still reach the one Note");
    }

    /// byte-level `PlcfHdd` parsing for a single-section
    /// document. `aCP` has `n + 2 = 14` entries for `n = 12` stories (the
    /// 6 fixed separators + this one section's 6 header/footer stories):
    /// the first 6 (indices 0..6) are all-empty separators, story 6..12
    /// hold the real content, `aCP[12]` closes the last story, and
    /// `aCP[13]` is the spec's own "undefined, must be ignored" filler.
    #[test]
    fn test_plcf_hdd_splits_the_first_sections_six_stories() {
        let raw = "EVEN HEADERODD HEADEREVEN FOOTERODD FOOTERFIRST HEADERFIRST FOOTERX";
        assert_eq!(raw.chars().count(), 67, "fixture text must match the cps below exactly");

        let cps: [u32; 14] = [0, 0, 0, 0, 0, 0, 0, 11, 21, 32, 42, 54, 66, 67];
        let mut table_stream = Vec::new();
        for cp in cps {
            table_stream.extend_from_slice(&cp.to_le_bytes());
        }

        let stories = parse_plcf_hdd_stories(&table_stream, raw, 0, table_stream.len() as u32);
        assert_eq!(stories.even_header.as_deref(), Some("EVEN HEADER"));
        assert_eq!(stories.odd_header.as_deref(), Some("ODD HEADER"));
        assert_eq!(stories.even_footer.as_deref(), Some("EVEN FOOTER"));
        assert_eq!(stories.odd_footer.as_deref(), Some("ODD FOOTER"));
        assert_eq!(stories.first_header.as_deref(), Some("FIRST HEADER"));
        assert_eq!(stories.first_footer.as_deref(), Some("FIRST FOOTER"));
    }

    /// Regression: the six fixed stories ahead of the first section's
    /// group are Word's footnote/endnote separators and "(continued from
    /// previous page)" notices. `to_ir()` never carried them; the flat
    /// text did, so `plain_text()` showed a notice the document never
    /// displays and the two surfaces disagreed.
    #[test]
    fn test_plcf_hdd_first_section_cp_skips_the_six_separator_stories() {
        // Stories 0..6 hold a continuation notice (chars 0..9); the real
        // headers start at CP 9.
        let cps: [u32; 14] = [0, 0, 0, 9, 9, 9, 9, 20, 30, 41, 51, 63, 75, 76];
        let mut table_stream = Vec::new();
        for cp in cps {
            table_stream.extend_from_slice(&cp.to_le_bytes());
        }
        assert_eq!(plcf_hdd_first_section_cp(&table_stream, 0, table_stream.len() as u32), Some(9));
        assert_eq!(plcf_hdd_first_section_cp(&[0u8; 8], 0, 0), None);
    }

    #[test]
    fn test_plcf_hdd_zero_length_yields_no_stories() {
        let stories = parse_plcf_hdd_stories(&[0u8; 8], "", 0, 0);
        assert!(stories.is_empty());
    }

    #[test]
    fn test_plcf_hdd_shorter_than_one_section_yields_no_stories() {
        // Only the 6 fixed separators (n=6, needs n+2=8 CPs) — no
        // section's worth of header/footer stories at all.
        let table_stream = vec![0u8; 8 * 4];
        let stories = parse_plcf_hdd_stories(&table_stream, "", 0, table_stream.len() as u32);
        assert!(stories.is_empty());
    }

    /// `doc_to_ir` must attach `PlcfHdd`-derived stories to
    /// the real `Section.header`/`.footer`/etc fields instead of dumping
    /// the merged blob as a generic `TextBox`, and must still fall back
    /// to the old behavior when no structured data is available.
    #[test]
    fn test_header_footer_stories_reach_the_section_fields() {
        use crate::ir::Element;
        let mut doc = make_doc("Body text.");
        doc.subdocuments = vec![SubDocument {
            kind: SubDocumentKind::HeadersFooters,
            text: "irrelevant merged blob".into(),
        }];
        doc.header_footer = HeaderFooterStories {
            odd_header: Some("The Odd Header".to_string()),
            first_footer: Some("The First Footer".to_string()),
            ..Default::default()
        };

        let ir = crate::convert_doc::doc_to_ir(&doc);
        let section = &ir.sections[0];
        assert!(section.header.is_some(), "odd_header must reach Section.header");
        assert!(
            section.first_page_footer.is_some(),
            "first_footer must reach Section.first_page_footer"
        );
        assert!(section.footer.is_none());
        assert!(
            !section
                .elements
                .iter()
                .any(|e| matches!(e, Element::TextBox(_))),
            "must not also dump the merged blob as a generic TextBox once structured: {:?}",
            section.elements
        );
    }

    #[test]
    fn test_header_footer_falls_back_to_a_textbox_when_plcf_hdd_is_absent() {
        use crate::ir::Element;
        let mut doc = make_doc("Body text.");
        doc.subdocuments = vec![SubDocument {
            kind: SubDocumentKind::HeadersFooters,
            text: "OLD MERGED HEADER BLOB".into(),
        }];
        // doc.header_footer left at its default (empty) — no PlcfHdd data.

        let ir = crate::convert_doc::doc_to_ir(&doc);
        let section = &ir.sections[0];
        assert!(section.header.is_none());
        assert!(
            section
                .elements
                .iter()
                .any(|e| matches!(e, Element::TextBox(_))),
            "must fall back to the old merged TextBox when PlcfHdd yielded nothing: {:?}",
            section.elements
        );
    }

    /// byte-level `PlcfandTxt`/`PlcfandRef` parsing for a
    /// 2-comment document. `PlcfandTxt.aCP` has `n + 2 = 4` entries for
    /// `n = 2` comment ranges within the Comments substory's own
    /// character space; `PlcfandRef`'s data elements are 30-byte
    /// `ATRDPre10`s whose `ibst` (offset 20) indexes `comment_authors`.
    #[test]
    fn test_parse_comments_splits_and_attributes_two_comments() {
        let raw = "\u{5}First comment\r\u{5}Second comment\r";
        let first_len = "\u{5}First comment\r".chars().count();
        let total_len = raw.chars().count();
        assert_eq!(first_len, 15);
        assert_eq!(total_len, 31);

        // PlcfandTxt: pure-CP PLC, comments-substory-local CPs.
        let and_txt_cps: [u32; 4] = [0, first_len as u32, total_len as u32, 0];
        let mut and_txt = Vec::new();
        for cp in and_txt_cps {
            and_txt.extend_from_slice(&cp.to_le_bytes());
        }

        // PlcfandRef: 3 main-doc CPs (n=2 comments) + 2 ATRDPre10s (30
        // bytes each), ibst at offset 20 within each.
        let and_ref_cps: [u32; 3] = [0, 1, 2];
        let mut and_ref = Vec::new();
        for cp in and_ref_cps {
            and_ref.extend_from_slice(&cp.to_le_bytes());
        }
        let mut atrd0 = vec![0u8; 30];
        atrd0[20..22].copy_from_slice(&1u16.to_le_bytes()); // ibst=1
        let mut atrd1 = vec![0u8; 30];
        atrd1[20..22].copy_from_slice(&0u16.to_le_bytes()); // ibst=0
        and_ref.extend_from_slice(&atrd0);
        and_ref.extend_from_slice(&atrd1);

        // Lay both PLCs out in one fake table stream, back to back.
        let and_txt_fc = 0u32;
        let and_ref_fc = and_txt.len() as u32;
        let mut table_stream = and_txt.clone();
        table_stream.extend_from_slice(&and_ref);

        let authors = vec!["Miklos Vajna".to_string(), "vmiklos".to_string()];
        let comments = parse_comments(
            &table_stream,
            raw,
            and_txt_fc,
            and_txt.len() as u32,
            and_ref_fc,
            and_ref.len() as u32,
            &authors,
        );

        assert_eq!(comments.len(), 2, "expected 2 split comments: {comments:?}");
        assert_eq!(comments[0].text, "First comment");
        assert_eq!(comments[0].author.as_deref(), Some("vmiklos"), "ibst=1 -> authors[1]");
        assert_eq!(comments[1].text, "Second comment");
        assert_eq!(comments[1].author.as_deref(), Some("Miklos Vajna"), "ibst=0 -> authors[0]");
    }

    #[test]
    fn test_parse_comments_falls_back_to_empty_on_a_count_mismatch() {
        // A PlcfandTxt describing 2 ranges but a PlcfandRef describing
        // only 1 ATRDPre10 must not be trusted at all.
        let and_txt_cps: [u32; 4] = [0, 5, 10, 0];
        let mut and_txt = Vec::new();
        for cp in and_txt_cps {
            and_txt.extend_from_slice(&cp.to_le_bytes());
        }
        let and_ref_cps: [u32; 2] = [0, 1];
        let mut and_ref = Vec::new();
        for cp in and_ref_cps {
            and_ref.extend_from_slice(&cp.to_le_bytes());
        }
        and_ref.extend_from_slice(&[0u8; 30]);

        let mut table_stream = and_txt.clone();
        table_stream.extend_from_slice(&and_ref);
        let comments = parse_comments(
            &table_stream,
            "helloworld",
            0,
            and_txt.len() as u32,
            and_txt.len() as u32,
            and_ref.len() as u32,
            &[],
        );
        assert!(comments.is_empty(), "count mismatch must fall back to empty: {comments:?}");
    }

    /// `doc_to_ir` must emit one `Element::Endnote` per
    /// `doc.comments()` entry (each with its own resolved author) instead
    /// of falling through to the generic merged-substory path, whenever
    /// `PlcfandTxt`/`PlcfandRef` successfully split the document.
    #[test]
    fn test_split_comments_reach_the_ir_as_separate_notes() {
        use crate::ir::Element;
        let mut doc = make_doc("Body text.");
        doc.subdocuments = vec![SubDocument {
            kind: SubDocumentKind::Comments,
            text: "Inner\nOuter\nAs in non-range.".into(),
        }];
        doc.comments = vec![
            ParsedComment {
                text: "Inner".into(),
                author: Some("vmiklos".into()),
            },
            ParsedComment {
                text: "Outer".into(),
                author: Some("vmiklos".into()),
            },
            ParsedComment {
                text: "As in non-range.".into(),
                author: Some("Miklos Vajna".into()),
            },
        ];

        let ir = crate::convert_doc::doc_to_ir(&doc);
        let notes: Vec<&crate::ir::Note> = ir.sections[0]
            .elements
            .iter()
            .filter_map(|e| {
                if let Element::Endnote(n) = e {
                    Some(n)
                } else {
                    None
                }
            })
            .collect();
        assert_eq!(notes.len(), 3, "expected one Endnote per split comment: {notes:?}");
        assert_eq!(notes[0].author.as_deref(), Some("vmiklos"));
        assert_eq!(notes[1].author.as_deref(), Some("vmiklos"));
        assert_eq!(notes[2].author.as_deref(), Some("Miklos Vajna"));
        assert_eq!(notes[0].id, 0);
        assert_eq!(notes[1].id, 1);
        assert_eq!(notes[2].id, 2);
    }
}
