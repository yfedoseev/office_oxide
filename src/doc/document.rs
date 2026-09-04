//! High-level DOC document API.

use std::io::{Read, Seek};

use crate::cfb::CfbReader;

use super::error::{DocError, Result};
use super::fib::Fib;
use super::images::{DocImage, extract_images};
use super::papx::{DocParagraph, build_paragraphs, parse_papx_paragraphs};
use super::piece_table::{extract_text, parse_clx, sanitize_text};
use super::styles::{StyleDef, parse_style_sheet};

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
}

impl DocDocument {
    /// Open a DOC file from a reader.
    pub fn from_reader<R: Read + Seek>(reader: R) -> Result<Self> {
        let mut cfb = CfbReader::new(reader)?;

        let word_doc = cfb
            .open_stream("WordDocument")
            .map_err(|_| DocError::MissingStream("WordDocument stream not found".into()))?;

        let fib = match Fib::parse(&word_doc) {
            Ok(f) => f,
            Err(_) => {
                return Ok(Self {
                    text: String::new(),
                    images: Vec::new(),
                    paragraphs: Vec::new(),
                });
            }, // Unsupported Word version
        };

        // Open the appropriate table stream; try preferred first, then fallback.
        let table_stream = if fib.use_table1 {
            cfb.open_stream("1Table")
                .or_else(|_| cfb.open_stream("0Table"))
        } else {
            cfb.open_stream("0Table")
                .or_else(|_| cfb.open_stream("1Table"))
        };
        let table_stream = match table_stream {
            Ok(s) => s,
            Err(_) => {
                return Ok(Self {
                    text: String::new(),
                    images: Vec::new(),
                    paragraphs: Vec::new(),
                });
            }, // Word 6/95 or corrupted
        };

        // Extract CLX from the table stream.
        let clx_start = fib.clx_offset as usize;
        let clx_end = clx_start + fib.clx_size as usize;

        if clx_start >= table_stream.len()
            || clx_size_zero_or_oob(fib.clx_size, clx_start, table_stream.len())
        {
            // CLX not available — return empty document.
            return Ok(Self {
                text: String::new(),
                images: Vec::new(),
                paragraphs: Vec::new(),
            });
        }

        let clx_end = clx_end.min(table_stream.len());
        let clx_data = &table_stream[clx_start..clx_end];
        let pieces = match parse_clx(clx_data) {
            Ok(p) => p,
            Err(_) => {
                return Ok(Self {
                    text: String::new(),
                    images: Vec::new(),
                    paragraphs: Vec::new(),
                });
            },
        };

        // Extract main document text only (not footnotes, headers, etc.).
        let raw_text = extract_text(&word_doc, &pieces, fib.text_len);
        let text = sanitize_text(&raw_text);

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
            // Parse the style sheet so paragraph styles (incl. built-in
            // Heading 1-9 and user-defined "Heading N") can be resolved to a
            // real heading level. A malformed/absent sheet yields an empty
            // list and headings fall back to the line heuristic.
            let styles: Vec<StyleDef> = parse_style_sheet(&table_stream, &fib);
            build_paragraphs(&word_doc, &pieces, &fkp, fib.text_len, &styles)
        } else {
            Vec::new()
        };

        // Extract images from the Data stream (if present).
        let images = match cfb.open_stream("Data") {
            Ok(data_stream) => extract_images(&data_stream),
            Err(_) => Vec::new(),
        };

        Ok(Self {
            text,
            images,
            paragraphs,
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

    /// Get the extracted plain text.
    pub fn plain_text(&self) -> String {
        self.text.clone()
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
    pub fn to_markdown(&self) -> String {
        let mut result = String::new();
        let mut prev_empty = false;

        for line in self.text.lines() {
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
}

fn clx_size_zero_or_oob(clx_size: u32, clx_start: usize, stream_len: usize) -> bool {
    clx_size == 0 || clx_start + clx_size as usize > stream_len + 1024 // allow some slack
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
    fn markdown_double_spacing() {
        let doc = DocDocument {
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
    fn plain_text_access() {
        let doc = DocDocument {
            images: Vec::new(),
            text: "Hello World".into(),
            paragraphs: Vec::new(),
        };
        assert_eq!(doc.plain_text(), "Hello World");
    }

    fn make_doc(text: &str) -> DocDocument {
        DocDocument {
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
    fn ir_list_emits_nested_list_from_ilvl_paragraphs() {
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
    fn ir_consecutive_list_runs_split_on_prose() {
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
    fn ir_empty_doc_produces_empty_section() {
        let ir = crate::convert_doc::doc_to_ir(&make_doc(""));
        assert!(ir.sections[0].elements.is_empty());
        assert!(ir.metadata.title.is_none());
    }

    #[test]
    fn ir_allcaps_first_line_becomes_h1() {
        use crate::ir::Element;
        let ir = crate::convert_doc::doc_to_ir(&make_doc("INTRODUCTION\nSome text here."));
        assert_eq!(ir.metadata.title.as_deref(), Some("INTRODUCTION"));
        assert!(matches!(ir.sections[0].elements[0], Element::Heading(ref h) if h.level == 1));
    }

    #[test]
    fn ir_first_short_line_no_punct_becomes_h1() {
        use crate::ir::Element;
        let ir = crate::convert_doc::doc_to_ir(&make_doc("My Document Title\nThis is body text."));
        assert!(matches!(ir.sections[0].elements[0], Element::Heading(ref h) if h.level == 1));
    }

    #[test]
    fn ir_allcaps_non_first_line_becomes_h2() {
        use crate::ir::Element;
        let ir = crate::convert_doc::doc_to_ir(&make_doc("Title\nSECTION TWO\nBody text."));
        assert!(matches!(ir.sections[0].elements[1], Element::Heading(ref h) if h.level == 2));
    }

    #[test]
    fn ir_styled_heading_uses_real_level() {
        use crate::ir::Element;
        // A styled "Heading 3" paragraph that is NOT the first element must
        // keep its real level (3) instead of collapsing to the heuristic level
        // 2 (the `elements.is_empty() ? 1 : 2` rule in `emit_prose`). This is
        // the regression test for deriving heading levels from paragraph style
        // rather than the line heuristic.
        let doc = make_doc_with_paragraphs(vec![
            pap("Intro paragraph.", Default::default()),
            pap(
                "Subsection",
                crate::doc::sprm::PapProps {
                    heading_level: Some(3),
                    ..Default::default()
                },
            ),
        ]);
        let ir = crate::convert_doc::doc_to_ir(&doc);
        let elements = &ir.sections[0].elements;
        assert!(matches!(elements[0], Element::Paragraph(_)));
        match &elements[1] {
            Element::Heading(h) => {
                assert_eq!(h.level, 3, "styled heading must keep its real level")
            },
            other => panic!("expected a Heading, got {:?}", other),
        }
    }

    /// Regression: MS-DOC outline levels run to `MAX_OUTLINE_LEVEL` but
    /// `Heading::level` is a 1..=MAX_HEADING_DEPTH markdown depth, so a
    /// deeply-nested heading must clamp at the IR boundary rather than emitting
    /// an out-of-contract level.
    #[test]
    fn ir_deep_outline_level_clamps_to_ir_max_depth() {
        use crate::doc::MAX_OUTLINE_LEVEL;
        use crate::ir::Element;
        use crate::ir::MAX_HEADING_DEPTH;
        let doc = make_doc_with_paragraphs(vec![pap(
            "Deep section",
            crate::doc::sprm::PapProps {
                heading_level: Some(MAX_OUTLINE_LEVEL),
                ..Default::default()
            },
        )]);
        let ir = crate::convert_doc::doc_to_ir(&doc);
        match &ir.sections[0].elements[0] {
            Element::Heading(h) => assert_eq!(
                h.level, MAX_HEADING_DEPTH,
                "outline level {MAX_OUTLINE_LEVEL} must clamp to {MAX_HEADING_DEPTH}"
            ),
            other => panic!("expected a Heading, got {:?}", other),
        }
    }

    #[test]
    fn ir_line_ending_with_period_becomes_paragraph() {
        use crate::ir::Element;
        let ir = crate::convert_doc::doc_to_ir(&make_doc("This is a sentence."));
        assert!(matches!(ir.sections[0].elements[0], Element::Paragraph(_)));
    }

    #[test]
    fn ir_blank_lines_are_skipped() {
        let ir = crate::convert_doc::doc_to_ir(&make_doc("Title\n\n\nText"));
        assert_eq!(ir.sections[0].elements.len(), 2);
    }

    #[test]
    fn ir_list_run_with_nonzero_base_level_keeps_every_item() {
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
    fn ir_format_is_doc() {
        let ir = crate::convert_doc::doc_to_ir(&make_doc("content"));
        assert_eq!(ir.metadata.format, crate::format::DocumentFormat::Doc);
    }

    // --------------------------------------------------------------------
    // End-to-end regression for the style-sheet heading path (PR #127).
    //
    // Every real `.doc` in the POI / Tika / LibreOffice corpora reports
    // `fc_stshf == 0` (no style sheet), so the new path that resolves a
    // heading's *real* level from the paragraph style (see `build_paragraphs`
    // → `heading_level_for_istd`) is never exercised by a corpus run. This test
    // synthesises a complete CFB `.doc` in code — per AGENTS.md #4 no
    // third-party fixture is committed — whose `1Table` carries a style sheet
    // mapping a paragraph's `istd` to built-in `Heading 3`, and asserts the IR
    // heading keeps level 3. The line heuristic tops out at level 2, so a
    // level-3 heading can ONLY come from the style sheet: the output
    // simultaneously proves the path fires and that it is not silently
    // collapsing to the heuristic level.
    // --------------------------------------------------------------------

    // CFB v3 sector sentinels (mirror `cfb::header`).
    const CFB_EOC: u32 = 0xFFFF_FFFE;
    const CFB_FREE: u32 = 0xFFFF_FFFF;
    const CFB_FATSECT: u32 = 0xFFFF_FFFD;

    /// Build a complete, minimal CFB `.doc` driving the style-sheet heading
    /// path end to end: an "Introduction." body paragraph (Normal style)
    /// followed by a "Subsection Three" paragraph styled `Heading 3`. The bytes
    /// are synthesised here, not read from any document file.
    fn build_synthetic_styled_doc() -> Vec<u8> {
        const SECTOR: usize = 512;
        const WORD_DOC_SECTORS: usize = 4; // 2048 bytes
        const TABLE_SECTORS: usize = 2; // 1024 bytes
        const TOTAL_SECTORS: usize = 1 + 1 + WORD_DOC_SECTORS + TABLE_SECTORS;

        let mut file = vec![0u8; SECTOR + TOTAL_SECTORS * SECTOR];

        // ── Header ──
        file[0..8].copy_from_slice(&[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
        file[0x18..0x1A].copy_from_slice(&0x003Eu16.to_le_bytes());
        file[0x1A..0x1C].copy_from_slice(&3u16.to_le_bytes());
        file[0x1C..0x1E].copy_from_slice(&0xFFFEu16.to_le_bytes());
        file[0x1E..0x20].copy_from_slice(&9u16.to_le_bytes());
        file[0x20..0x22].copy_from_slice(&6u16.to_le_bytes());
        file[0x2C..0x30].copy_from_slice(&1u32.to_le_bytes()); // FAT count = 1
        file[0x30..0x34].copy_from_slice(&0u32.to_le_bytes()); // first dir = 0
        file[0x38..0x3C].copy_from_slice(&4096u32.to_le_bytes()); // mini cutoff
        file[0x3C..0x40].copy_from_slice(&CFB_EOC.to_le_bytes()); // no mini-FAT
        file[0x40..0x44].copy_from_slice(&0u32.to_le_bytes());
        file[0x44..0x48].copy_from_slice(&CFB_EOC.to_le_bytes()); // no DIFAT chain
        file[0x48..0x4C].copy_from_slice(&0u32.to_le_bytes());
        file[0x4C..0x50].copy_from_slice(&1u32.to_le_bytes()); // DIFAT[0] = FAT sector 1
        for i in 1..109 {
            let off = 0x4C + i * 4;
            file[off..off + 4].copy_from_slice(&CFB_FREE.to_le_bytes());
        }

        // ── Directory (sector 0): Root Entry -> {WordDocument, 1Table} ──
        let dir = SECTOR;
        write_dir_entry(&mut file[dir..dir + 128], "Root Entry", 5, 1, CFB_EOC, 0);
        write_dir_entry(
            &mut file[dir + 128..dir + 256],
            "WordDocument",
            2,
            CFB_FREE,
            2,
            (WORD_DOC_SECTORS * SECTOR) as u32,
        );
        write_dir_entry(
            &mut file[dir + 256..dir + 384],
            "1Table",
            2,
            CFB_FREE,
            6,
            (TABLE_SECTORS * SECTOR) as u32,
        );
        // Entry 3: empty (type 0) — fills out the 512-byte directory sector.

        // ── FAT (sector 1) ──
        let fat = SECTOR + SECTOR; // byte offset 1024
        let mut fat_entry = |idx: usize, val: u32| {
            let off = fat + idx * 4;
            file[off..off + 4].copy_from_slice(&val.to_le_bytes());
        };
        fat_entry(0, CFB_EOC); // directory, single sector
        fat_entry(1, CFB_FATSECT); // this sector is a FAT sector
        fat_entry(2, 3); // WordDocument chain: 2 -> 3 -> 4 -> 5 -> EOC
        fat_entry(3, 4);
        fat_entry(4, 5);
        fat_entry(5, CFB_EOC);
        fat_entry(6, 7); // 1Table chain: 6 -> 7 -> EOC
        fat_entry(7, CFB_EOC);
        for i in 8..128 {
            fat_entry(i, CFB_FREE);
        }

        // ── WordDocument stream (sectors 2..5) ──
        let wd_off = SECTOR + 2 * SECTOR; // byte offset 1536
        let mut wd = vec![0u8; WORD_DOC_SECTORS * SECTOR];

        // FIB.
        wd[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes()); // wIdent (Word 97+)
        wd[2..4].copy_from_slice(&0x00C1u16.to_le_bytes()); // nFib
        wd[0x0A..0x0C].copy_from_slice(&(1u16 << 9).to_le_bytes()); // flags: use 1Table

        let raw = "Introduction.\rSubsection Three\r";
        let text_bytes: Vec<u8> = raw.encode_utf16().flat_map(|u| u.to_le_bytes()).collect();
        let ccp = (text_bytes.len() / 2) as u32; // main-text character count
        wd[0x4C..0x50].copy_from_slice(&ccp.to_le_bytes()); // ccpText

        let clx = build_clx(ccp);
        let stsh = build_stsh_heading3();
        wd[0xA2..0xA6].copy_from_slice(&288u32.to_le_bytes()); // fcStshf = 0x120
        wd[0xA6..0xAA].copy_from_slice(&(stsh.len() as u32).to_le_bytes()); // lcbStshf
        wd[0x102..0x106].copy_from_slice(&256u32.to_le_bytes()); // fcPlcfBtePapx
        wd[0x106..0x10A].copy_from_slice(&12u32.to_le_bytes()); // lcbPlcfBtePapx
        wd[0x1A2..0x1A6].copy_from_slice(&0u32.to_le_bytes()); // fcClx
        wd[0x1A6..0x1AA].copy_from_slice(&(clx.len() as u32).to_le_bytes()); // lcbClx

        // Document text at fc 0x300 (Unicode).
        wd[0x300..0x300 + text_bytes.len()].copy_from_slice(&text_bytes);

        // PAPX FKP page at page number 2 (byte offset 0x400). Two paragraphs:
        // a Normal body paragraph and a Heading-3 paragraph.
        let para0_end_fc = 0x300 + ("Introduction.\r".encode_utf16().count() * 2) as u32;
        let para1_end_fc = 0x300 + (raw.encode_utf16().count() * 2) as u32;
        wd[0x400..0x404].copy_from_slice(&0x300u32.to_le_bytes()); // rgfc[0]
        wd[0x404..0x408].copy_from_slice(&para0_end_fc.to_le_bytes()); // rgfc[1]
        wd[0x408..0x40C].copy_from_slice(&para1_end_fc.to_le_bytes()); // rgfc[2]
        // rgbx: byte 0 of each BX is the word offset of that paragraph's PAPX.
        wd[0x40C] = 19; // para0 PAPX at byte 38 (word 19)
        wd[0x419] = 21; // para1 PAPX at byte 42 (word 21)
        // para0 PAPX: cw=2, istd=0 (Normal), empty grpprl.
        wd[0x426] = 0x02; // cw
        // 0x427..0x429 = istd 0x0000, 0x429 = grpprl (0x00)
        // para1 PAPX: cw=2, istd=3 (Heading 3), empty grpprl.
        wd[0x42A] = 0x02; // cw
        wd[0x42B..0x42D].copy_from_slice(&0x0003u16.to_le_bytes()); // istd = 3
        // 0x42D = grpprl (0x00)
        wd[0x5FF] = 2; // crun = 2

        file[wd_off..wd_off + wd.len()].copy_from_slice(&wd);

        // ── 1Table stream (sectors 6..7) ──
        let tbl_off = SECTOR + 6 * SECTOR; // byte offset 3584
        let mut tbl = vec![0u8; TABLE_SECTORS * SECTOR];
        tbl[0..clx.len()].copy_from_slice(&clx); // CLX at offset 0
        // PlcfBtePapx at 0x100: n=1 -> [FC0=0][FC1=ccp][BTE=pn 2].
        tbl[0x100..0x104].copy_from_slice(&0u32.to_le_bytes());
        tbl[0x104..0x108].copy_from_slice(&ccp.to_le_bytes());
        tbl[0x108..0x10C].copy_from_slice(&2u32.to_le_bytes());
        // STSH at 0x120 (0x300 would overflow the 2-sector table stream).
        tbl[0x120..0x120 + stsh.len()].copy_from_slice(&stsh);

        file[tbl_off..tbl_off + tbl.len()].copy_from_slice(&tbl);

        file
    }

    /// Like `build_synthetic_styled_doc`, but paragraph 1 carries
    /// `sprmPOutlineLvl` (0x6412) in its PAPX grpprl (level 5) instead of
    /// relying on a style. Its PAPX `istd` is 0 (Normal), so the Heading level
    /// comes solely from the outline SPRM — exercising the SPRM path end to end
    /// (CFB → FKP → `build_paragraphs` → `doc_to_ir`).
    fn build_synthetic_outline_doc() -> Vec<u8> {
        const SECTOR: usize = 512;
        const WORD_DOC_SECTORS: usize = 4; // 2048 bytes
        const TABLE_SECTORS: usize = 2; // 1024 bytes
        const TOTAL_SECTORS: usize = 1 + 1 + WORD_DOC_SECTORS + TABLE_SECTORS;

        let mut file = vec![0u8; SECTOR + TOTAL_SECTORS * SECTOR];

        // ── Header ──
        file[0..8].copy_from_slice(&[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
        file[0x18..0x1A].copy_from_slice(&0x003Eu16.to_le_bytes());
        file[0x1A..0x1C].copy_from_slice(&3u16.to_le_bytes());
        file[0x1C..0x1E].copy_from_slice(&0xFFFEu16.to_le_bytes());
        file[0x1E..0x20].copy_from_slice(&9u16.to_le_bytes());
        file[0x20..0x22].copy_from_slice(&6u16.to_le_bytes());
        file[0x2C..0x30].copy_from_slice(&1u32.to_le_bytes()); // FAT count = 1
        file[0x30..0x34].copy_from_slice(&0u32.to_le_bytes()); // first dir = 0
        file[0x38..0x3C].copy_from_slice(&4096u32.to_le_bytes()); // mini cutoff
        file[0x3C..0x40].copy_from_slice(&CFB_EOC.to_le_bytes()); // no mini-FAT
        file[0x40..0x44].copy_from_slice(&0u32.to_le_bytes());
        file[0x44..0x48].copy_from_slice(&CFB_EOC.to_le_bytes()); // no DIFAT chain
        file[0x48..0x4C].copy_from_slice(&0u32.to_le_bytes());
        file[0x4C..0x50].copy_from_slice(&1u32.to_le_bytes()); // DIFAT[0] = FAT sector 1
        for i in 1..109 {
            let off = 0x4C + i * 4;
            file[off..off + 4].copy_from_slice(&CFB_FREE.to_le_bytes());
        }

        // ── Directory (sector 0): Root Entry -> {WordDocument, 1Table} ──
        let dir = SECTOR;
        write_dir_entry(&mut file[dir..dir + 128], "Root Entry", 5, 1, CFB_EOC, 0);
        write_dir_entry(
            &mut file[dir + 128..dir + 256],
            "WordDocument",
            2,
            CFB_FREE,
            2,
            (WORD_DOC_SECTORS * SECTOR) as u32,
        );
        write_dir_entry(
            &mut file[dir + 256..dir + 384],
            "1Table",
            2,
            CFB_FREE,
            6,
            (TABLE_SECTORS * SECTOR) as u32,
        );

        // ── FAT (sector 1) ──
        let fat = SECTOR + SECTOR; // byte offset 1024
        let mut fat_entry = |idx: usize, val: u32| {
            let off = fat + idx * 4;
            file[off..off + 4].copy_from_slice(&val.to_le_bytes());
        };
        fat_entry(0, CFB_EOC); // directory, single sector
        fat_entry(1, CFB_FATSECT); // this sector is a FAT sector
        fat_entry(2, 3); // WordDocument chain: 2 -> 3 -> 4 -> 5 -> EOC
        fat_entry(3, 4);
        fat_entry(4, 5);
        fat_entry(5, CFB_EOC);
        fat_entry(6, 7); // 1Table chain: 6 -> 7 -> EOC
        fat_entry(7, CFB_EOC);
        for i in 8..128 {
            fat_entry(i, CFB_FREE);
        }

        // ── WordDocument stream (sectors 2..5) ──
        let wd_off = SECTOR + 2 * SECTOR; // byte offset 1536
        let mut wd = vec![0u8; WORD_DOC_SECTORS * SECTOR];

        // FIB.
        wd[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes()); // wIdent (Word 97+)
        wd[2..4].copy_from_slice(&0x00C1u16.to_le_bytes()); // nFib
        wd[0x0A..0x0C].copy_from_slice(&(1u16 << 9).to_le_bytes()); // flags: use 1Table

        let raw = "Introduction.\rSubsection Three\r";
        let text_bytes: Vec<u8> = raw.encode_utf16().flat_map(|u| u.to_le_bytes()).collect();
        let ccp = (text_bytes.len() / 2) as u32; // main-text character count
        wd[0x4C..0x50].copy_from_slice(&ccp.to_le_bytes()); // ccpText

        let clx = build_clx(ccp);
        let stsh = build_stsh_heading3();
        wd[0xA2..0xA6].copy_from_slice(&288u32.to_le_bytes()); // fcStshf = 0x120
        wd[0xA6..0xAA].copy_from_slice(&(stsh.len() as u32).to_le_bytes()); // lcbStshf
        wd[0x102..0x106].copy_from_slice(&256u32.to_le_bytes()); // fcPlcfBtePapx
        wd[0x106..0x10A].copy_from_slice(&12u32.to_le_bytes()); // lcbPlcfBtePapx
        wd[0x1A2..0x1A6].copy_from_slice(&0u32.to_le_bytes()); // fcClx
        wd[0x1A6..0x1AA].copy_from_slice(&(clx.len() as u32).to_le_bytes()); // lcbClx

        // Document text at fc 0x300 (Unicode).
        wd[0x300..0x300 + text_bytes.len()].copy_from_slice(&text_bytes);

        // PAPX FKP page at page number 2 (byte offset 0x400). Two paragraphs:
        // a Normal body paragraph and an outline-SPRM Heading-5 paragraph.
        let para0_end_fc = 0x300 + ("Introduction.\r".encode_utf16().count() * 2) as u32;
        let para1_end_fc = 0x300 + (raw.encode_utf16().count() * 2) as u32;
        wd[0x400..0x404].copy_from_slice(&0x300u32.to_le_bytes()); // rgfc[0]
        wd[0x404..0x408].copy_from_slice(&para0_end_fc.to_le_bytes()); // rgfc[1]
        wd[0x408..0x40C].copy_from_slice(&para1_end_fc.to_le_bytes()); // rgfc[2]
        // rgbx: byte 0 of each BX is the word offset of that paragraph's PAPX.
        wd[0x40C] = 19; // para0 PAPX at byte 38 (word 19)
        wd[0x419] = 21; // para1 PAPX at byte 42 (word 21)
        // para0 PAPX: cw=2, istd=0 (Normal), empty grpprl.
        wd[0x426] = 0x02; // cw
        // 0x427..0x429 = istd 0x0000, 0x429 = grpprl (0x00)
        // para1 PAPX: cw=5, istd=0 (Normal) — but the grpprl carries
        // `sprmPOutlineLvl` (0x6412, spra-3 => 4-byte operand) with outline
        // level 5. The PAPX is 1 (cw) + 2 (istd) + 7 (grpprl) = 10 bytes = cb.
        wd[0x42A] = 0x05; // cw
        wd[0x42B..0x42D].copy_from_slice(&0x0000u16.to_le_bytes()); // istd = 0 (Normal)
        // grpprl: 0x6412 (LE) + 4-byte operand (low byte = 5) + 1 pad byte.
        wd[0x42D..0x434].copy_from_slice(&[0x12, 0x64, 0x05, 0x00, 0x00, 0x00, 0x00]);
        wd[0x5FF] = 2; // crun = 2

        file[wd_off..wd_off + wd.len()].copy_from_slice(&wd);

        // ── 1Table stream (sectors 6..7) ──
        let tbl_off = SECTOR + 6 * SECTOR; // byte offset 3584
        let mut tbl = vec![0u8; TABLE_SECTORS * SECTOR];
        tbl[0..clx.len()].copy_from_slice(&clx); // CLX at offset 0
        // PlcfBtePapx at 0x100: n=1 -> [FC0=0][FC1=ccp][BTE=pn 2].
        tbl[0x100..0x104].copy_from_slice(&0u32.to_le_bytes());
        tbl[0x104..0x108].copy_from_slice(&ccp.to_le_bytes());
        tbl[0x108..0x10C].copy_from_slice(&2u32.to_le_bytes());
        // STSH at 0x120 (0x300 would overflow the 2-sector table stream).
        tbl[0x120..0x120 + stsh.len()].copy_from_slice(&stsh);

        file[tbl_off..tbl_off + tbl.len()].copy_from_slice(&tbl);

        file
    }

    /// Write a 128-byte CFB directory entry.
    fn write_dir_entry(
        buf: &mut [u8],
        name: &str,
        entry_type: u8,
        child: u32,
        start_sector: u32,
        stream_size: u32,
    ) {
        let utf16: Vec<u16> = name.encode_utf16().collect();
        for (i, &ch) in utf16.iter().enumerate() {
            let bytes = ch.to_le_bytes();
            buf[i * 2] = bytes[0];
            buf[i * 2 + 1] = bytes[1];
        }
        let name_size = ((utf16.len() + 1) * 2) as u16;
        buf[0x40..0x42].copy_from_slice(&name_size.to_le_bytes());
        buf[0x42] = entry_type;
        buf[0x43] = 1; // color: black
        buf[0x44..0x48].copy_from_slice(&CFB_FREE.to_le_bytes()); // left sibling
        buf[0x48..0x4C].copy_from_slice(&CFB_FREE.to_le_bytes()); // right sibling
        buf[0x4C..0x50].copy_from_slice(&child.to_le_bytes()); // child
        buf[0x74..0x78].copy_from_slice(&start_sector.to_le_bytes()); // start sector
        buf[0x78..0x7C].copy_from_slice(&stream_size.to_le_bytes()); // stream size
    }

    /// Build a CLX with a single Unicode piece whose `fc` points at 0x300 in the
    /// WordDocument stream (where the document text is placed). `ccp` is the
    /// total main-text character count.
    fn build_clx(ccp: u32) -> Vec<u8> {
        let mut clx = Vec::new();
        clx.push(0x02); // Pcdt marker
        clx.extend_from_slice(&16u32.to_le_bytes()); // PlcPcd size = 16
        clx.extend_from_slice(&0u32.to_le_bytes()); // CP[0]
        clx.extend_from_slice(&ccp.to_le_bytes()); // CP[1]
        clx.extend_from_slice(&0u16.to_le_bytes()); // PCD: unused u16
        clx.extend_from_slice(&0x300u32.to_le_bytes()); // PCD: fc = 0x300 (Unicode)
        clx.extend_from_slice(&0u16.to_le_bytes()); // PCD: prm
        clx
    }

    /// Build one `LPStd` (`cbStd` + `STD` = `StdfBase` + `xstzName`), per
    /// MS-DOC §2.9.135 / §2.9.258 / §2.9.354.
    fn lpstd(sti: u16, name: &str) -> Vec<u8> {
        let mut std = vec![0u8; 10]; // StdfBase
        std[0..2].copy_from_slice(&sti.to_le_bytes());
        let units: Vec<u16> = name.encode_utf16().collect();
        std.extend_from_slice(&(units.len() as u16).to_le_bytes()); // cch
        for u in &units {
            std.extend_from_slice(&u.to_le_bytes());
        }
        std.extend_from_slice(&0u16.to_le_bytes()); // chTerm
        let mut out = (std.len() as u16).to_le_bytes().to_vec(); // cbStd
        out.extend_from_slice(&std);
        out
    }

    /// A spec-conformant 15-style STSH (§2.9.271): `cbStshi`(18) then `Stshif`(18),
    /// followed directly by `rglpstd`, using the fixed-index table — istd 0 is
    /// Normal (sti 0), istd 1–9 are Heading 1–9 (sti 1–9), and istd 10–14 are
    /// empty. So `istd 3` is `Heading 3`. Built from the spec, not from our
    /// parser.
    fn build_stsh_heading3() -> Vec<u8> {
        let mut d = 18u16.to_le_bytes().to_vec(); // cbStshi
        d.extend_from_slice(&15u16.to_le_bytes()); // cstd
        d.extend_from_slice(&0x000Au16.to_le_bytes()); // cbSTDBaseInFile
        d.extend_from_slice(&[0u8; 14]); // remainder of the 18-byte Stshif
        d.extend_from_slice(&lpstd(0, "Normal"));
        for lvl in 1..=9u16 {
            d.extend_from_slice(&lpstd(lvl, &format!("Heading {lvl}")));
        }
        for _ in 0..5 {
            d.extend_from_slice(&[0u8; 2]); // empty LPStd (cbStd = 0)
        }
        d
    }

    #[test]
    fn synthetic_doc_styled_heading_uses_style_sheet_level() {
        use crate::ir::Element;

        let doc_bytes = build_synthetic_styled_doc();
        // Round-trip through the full DOC pipeline: CFB parse → FIB → piece
        // table → PAPX FKP → style sheet → IR. No third-party fixture.
        let doc = DocDocument::from_reader(std::io::Cursor::new(doc_bytes))
            .expect("synthetic .doc must parse");
        let ir = crate::convert_doc::doc_to_ir(&doc);

        let elements = &ir.sections[0].elements;
        assert_eq!(elements.len(), 2, "expected intro paragraph + styled heading");
        assert!(
            matches!(elements[0], Element::Paragraph(_)),
            "first element must be ordinary body prose"
        );
        match &elements[1] {
            Element::Heading(h) => {
                // Heading 3 can ONLY come from the style sheet: the line
                // heuristic tops out at level 2, so level 3 proves the real
                // style-sheet level reached the IR.
                assert_eq!(h.level, 3, "styled Heading 3 must keep its real level");
            },
            other => panic!("second element must be a Heading, got {:?}", other),
        }
        assert_eq!(
            ir.metadata.title.as_deref(),
            Some("Subsection Three"),
            "the styled heading becomes the document title"
        );
    }

    #[test]
    fn synthetic_doc_outline_sprm_heading() {
        use crate::ir::Element;

        let doc_bytes = build_synthetic_outline_doc();
        // Full DOC pipeline: CFB → FIB → piece table → PAPX FKP → IR. The
        // second paragraph's PAPX carries `sprmPOutlineLvl` (0x6412, level 5)
        // in its grpprl with `istd` 0 (Normal), so the Heading level comes
        // solely from the SPRM — the style-sheet path would resolve istd 0
        // (Normal) to no heading at all, and the line heuristic tops out at
        // level 2. Level 5 therefore proves the `sprmPOutlineLvl` path reached
        // the IR end to end.
        let doc = DocDocument::from_reader(std::io::Cursor::new(doc_bytes))
            .expect("synthetic .doc must parse");
        let ir = crate::convert_doc::doc_to_ir(&doc);

        let elements = &ir.sections[0].elements;
        assert_eq!(elements.len(), 2, "expected intro paragraph + SPRM heading");
        assert!(
            matches!(elements[0], Element::Paragraph(_)),
            "first element must be ordinary body prose"
        );
        match &elements[1] {
            Element::Heading(h) => {
                assert_eq!(h.level, 5, "outline SPRM must set the real level 5");
            },
            other => panic!("second element must be a Heading, got {:?}", other),
        }
        assert_eq!(
            ir.metadata.title.as_deref(),
            Some("Subsection Three"),
            "the SPRM heading becomes the document title"
        );
    }
}
