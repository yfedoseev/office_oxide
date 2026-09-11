//! DOCX creation (write) module.
//!
//! Provides a builder API for creating DOCX files from scratch.
//!
//! # Example
//!
//! ```rust,no_run
//! use office_oxide::docx::write::{DocxWriter, Run, Alignment};
//!
//! let mut doc = DocxWriter::new();
//! doc.add_heading("Report", 1)
//!    .add_paragraph("This is a paragraph.")
//!    .add_rich_paragraph(&[
//!        Run::new("Bold text").bold(),
//!        Run::new(" and ").into(),
//!        Run::new("red italic").italic().color("FF0000"),
//!    ])
//!    .add_paragraph_aligned("Centred text", Alignment::Center)
//!    .add_list(&["First", "Second", "Third"], false)
//!    .add_table(&[
//!        vec!["Name", "Age"],
//!        vec!["Alice", "30"],
//!    ]);
//! doc.save("report.docx").unwrap();
//! ```

use std::io::{Seek, Write};
use std::path::Path;

use quick_xml::Writer;
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};

use crate::core::opc::{OpcWriter, PartName};
use crate::core::relationships::rel_types;
use crate::ir::{
    BorderLine, BorderStyle, CellVerticalAlign, ColumnLayout, ImageFormat, ImagePositioning,
    LineSpacing, ListStyle, PageSetup, ParagraphAlignment, SectionBreakType, TableAlignment,
    UnderlineStyle, VerticalAlign,
};

use super::Result;

// ---------------------------------------------------------------------------
// Content types
// ---------------------------------------------------------------------------

const CT_DOCUMENT: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const CT_STYLES: &str = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
const CT_FONT_TABLE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml";
const CT_SETTINGS: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";
const CT_NUMBERING: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
const CT_HEADER: &str = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
const CT_FOOTER: &str = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
const CT_FOOTNOTES: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml";
const CT_ENDNOTES: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml";

use crate::core::xml::ns::{R_STR as R_NS, WML_STR as WML_NS};

const DRAWING_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const DML_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PIC_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const WPS_NS: &str = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";

// ---------------------------------------------------------------------------
// Public data types
// ---------------------------------------------------------------------------

/// Paragraph text alignment.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Alignment {
    /// Left-aligned.
    Left,
    /// Centered.
    Center,
    /// Right-aligned.
    Right,
    /// Justified (both edges).
    Justify,
}

impl Alignment {
    fn as_wml_val(&self) -> &'static str {
        match self {
            Self::Left => "left",
            Self::Center => "center",
            Self::Right => "right",
            Self::Justify => "both",
        }
    }
}

/// A styled text run within a paragraph.
///
/// Build with the builder methods; plain text uses `Run::new("text")`.
///
/// # Example
/// ```rust,no_run
/// use office_oxide::docx::write::Run;
///
/// let r = Run::new("Hello, world!")
///     .bold()
///     .font_size(14.0)
///     .color("1F497D");
/// ```
#[derive(Debug, Clone, Default)]
pub struct Run {
    /// The text content of this run.
    pub text: String,
    /// Bold formatting.
    pub bold: bool,
    /// Italic formatting.
    pub italic: bool,
    /// Underline formatting (legacy; `underline_style` takes priority if set).
    pub underline: bool,
    /// Strikethrough formatting.
    pub strikethrough: bool,
    /// 6-char hex color string, e.g. `"FF0000"` (no leading `#`). `color_rgb` takes priority.
    pub color: Option<String>,
    /// Font size in points, e.g. `12.0`. `font_size_half_pt` takes priority.
    pub font_size_pt: Option<f64>,
    /// Font name, e.g. `"Arial"`.
    pub font_name: Option<String>,
    // --- new rich fields ---
    /// Underline style (takes priority over `underline: bool`).
    pub underline_style: Option<UnderlineStyle>,
    /// Font size in half-points (takes priority over `font_size_pt`).
    pub font_size_half_pt: Option<u32>,
    /// RGB color (takes priority over `color`).
    pub color_rgb: Option<[u8; 3]>,
    /// Highlight / shading background color.
    pub highlight: Option<[u8; 3]>,
    /// Vertical alignment (superscript / subscript).
    pub vertical_align: Option<VerticalAlign>,
    /// All-caps text transform.
    pub all_caps: bool,
    /// Hyperlink target URL. Emitted as a `w:hyperlink` wrapper with an
    /// external relationship; previously read and thrown away.
    pub hyperlink: Option<String>,
    /// Small-caps text transform.
    pub small_caps: bool,
    /// Character spacing in half-points (positive = expand, negative = condense).
    pub char_spacing_half_pt: Option<i32>,
    /// Footnote reference (special run that emits a `<w:footnoteReference>`).
    pub footnote_ref: Option<u32>,
    /// Endnote reference (special run that emits a `<w:endnoteReference>`).
    pub endnote_ref: Option<u32>,
}

impl Run {
    /// Create a run with the given text and default (unstyled) properties.
    pub fn new(text: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            ..Default::default()
        }
    }

    /// Apply bold formatting.
    pub fn bold(mut self) -> Self {
        self.bold = true;
        self
    }
    /// Apply italic formatting.
    pub fn italic(mut self) -> Self {
        self.italic = true;
        self
    }
    /// Apply underline formatting.
    pub fn underline(mut self) -> Self {
        self.underline = true;
        self
    }
    /// Apply strikethrough formatting.
    pub fn strikethrough(mut self) -> Self {
        self.strikethrough = true;
        self
    }

    /// Set the font color. `hex` is a 6-character hex string without `#`.
    pub fn color(mut self, hex: impl Into<String>) -> Self {
        self.color = Some(hex.into());
        self
    }

    /// Set font size in points (e.g. `12.0`).
    pub fn font_size(mut self, pt: f64) -> Self {
        self.font_size_pt = Some(pt);
        self
    }

    /// Set the font family name (e.g. `"Arial"`).
    pub fn font(mut self, name: impl Into<String>) -> Self {
        self.font_name = Some(name.into());
        self
    }

    fn has_rpr(&self) -> bool {
        self.bold
            || self.italic
            || self.underline
            || self.underline_style.is_some()
            || self.strikethrough
            || self.color.is_some()
            || self.color_rgb.is_some()
            || self.font_size_pt.is_some()
            || self.font_size_half_pt.is_some()
            || self.font_name.is_some()
            || self.highlight.is_some()
            || self.vertical_align.is_some()
            || self.all_caps
            || self.small_caps
            || self.char_spacing_half_pt.is_some()
            || self.footnote_ref.is_some()
            || self.endnote_ref.is_some()
    }
}

impl From<&str> for Run {
    fn from(s: &str) -> Self {
        Self::new(s)
    }
}

impl From<String> for Run {
    fn from(s: String) -> Self {
        Self::new(s)
    }
}

// ---------------------------------------------------------------------------
// Paragraph properties (IR-level, used by add_ir_paragraph)
// ---------------------------------------------------------------------------

/// Rich paragraph layout properties (used by `add_ir_paragraph`).
#[derive(Debug, Clone, Default)]
pub struct IrParaProps {
    /// Text alignment for the paragraph.
    pub alignment: Option<ParagraphAlignment>,
    /// Left indent in twips.
    pub indent_left_twips: Option<i32>,
    /// Right indent in twips.
    pub indent_right_twips: Option<i32>,
    /// First-line indent in twips (positive = indent, negative = hanging).
    pub first_line_indent_twips: Option<i32>,
    /// Space before the paragraph in twips.
    pub space_before_twips: Option<u32>,
    /// Space after the paragraph in twips.
    pub space_after_twips: Option<u32>,
    /// Line spacing specification.
    pub line_spacing: Option<LineSpacing>,
    /// Named paragraph style (e.g. `"Heading1"`).
    pub style: Option<String>,
    /// Numbering list reference `(numId, indentLevel)`.
    pub numbering: Option<(u32, u8)>,
    /// Keep this paragraph on the same page as the next paragraph.
    pub keep_with_next: bool,
    /// Prevent a page break within the paragraph.
    pub keep_together: bool,
    /// Force a page break before the paragraph.
    pub page_break_before: bool,
    /// Background shading color as `[R, G, B]`.
    pub background_color: Option<[u8; 3]>,
    /// Outline level in ECMA-376 §17.3.1.20's value space: `0` = Heading 1,
    /// … `9` = no outline level (body text).
    pub outline_level: Option<u8>,
    /// Tab stops for this paragraph. Dot-leader tables of contents lose both
    /// their leaders and their alignment when these are dropped.
    pub tabs: Vec<crate::ir::TabStop>,
    /// Absolute frame position (`w:framePr`). Used by the layout-preserving
    /// path; the field existed in the IR and reached no writer.
    pub frame_position: Option<crate::ir::FramePosition>,
    /// Paragraph border definition.
    pub border: Option<crate::ir::ParagraphBorder>,
}

// ---------------------------------------------------------------------------
// Internal data model
// ---------------------------------------------------------------------------

struct DocxParagraph {
    style: Option<String>,
    runs: Vec<Run>,
    numbering: Option<(u32, u8)>,
    alignment: Option<Alignment>,
}

impl DocxParagraph {
    fn plain(text: &str, style: Option<String>, numbering: Option<(u32, u8)>) -> Self {
        Self {
            style,
            runs: vec![Run::new(text)],
            numbering,
            alignment: None,
        }
    }
}

struct DocxTable {
    rows: Vec<Vec<String>>,
}

struct DocxRichParagraph {
    runs: Vec<Run>,
    props: IrParaProps,
}

struct DocxRichTable {
    column_widths_twips: Vec<u32>,
    border: Option<crate::ir::TableBorder>,
    alignment: Option<TableAlignment>,
    cell_padding_twips: Option<u32>,
    rows: Vec<DocxRichRow>,
    width_twips: Option<u32>,
    indent_left_twips: Option<i32>,
    caption: Option<String>,
}

struct DocxRichRow {
    height_twips: Option<u32>,
    allow_break: bool,
    repeat_as_header: bool,
    cells: Vec<DocxRichCell>,
}

struct DocxRichCell {
    content: Vec<DocxElement>,
    col_span: u32,
    row_span: u32,
    background_color: Option<[u8; 3]>,
    border: Option<crate::ir::TableBorder>,
    vertical_align: Option<CellVerticalAlign>,
    text_align: Option<ParagraphAlignment>,
    text_direction: Option<crate::ir::TextDirection>,
    width_twips: Option<u32>,
    padding: Option<crate::ir::CellPadding>,
    is_vmerge_continue: bool,
}

struct DocxImage {
    data: Vec<u8>,
    format: ImageFormat,
    display_width_emu: u64,
    display_height_emu: u64,
    alt_text: Option<String>,
    decorative: bool,
    positioning: ImagePositioning,
}

#[allow(dead_code)]
struct DocxSectPr {
    page_setup: Option<PageSetup>,
    columns: Option<ColumnLayout>,
    break_type: SectionBreakType,
    header_rids: Vec<(String, String)>,
    footer_rids: Vec<(String, String)>,
    footnote_rid: Option<String>,
    endnote_rid: Option<String>,
}

/// The type of header or footer section to add.
#[derive(Clone, Copy, PartialEq, Eq)]
pub enum HfType {
    /// Default (odd-page) header.
    DefaultHeader,
    /// Default (odd-page) footer.
    DefaultFooter,
    /// First-page header.
    FirstPageHeader,
    /// First-page footer.
    FirstPageFooter,
    /// Even-page header.
    EvenPageHeader,
    /// Even-page footer.
    EvenPageFooter,
}

struct DocxHf {
    hf_type: HfType,
    elements: Vec<DocxElement>,
}

struct DocxNote {
    id: u32,
    elements: Vec<DocxElement>,
}

struct DocxRichList {
    ordered: bool,
    items: Vec<Vec<DocxElement>>,
    start_number: Option<u32>,
    style: Option<ListStyle>,
    level: u8,
    num_id: u32,
}

struct DocxTextBox {
    content: Vec<DocxElement>,
    width_emu: u64,
    height_emu: u64,
    x_emu: i64,
    y_emu: i64,
    h_anchor: crate::ir::FloatAnchor,
    v_anchor: crate::ir::FloatAnchor,
    wrap: crate::ir::TextWrap,
}

enum DocxElement {
    Paragraph(DocxParagraph),
    RichParagraph(DocxRichParagraph),
    Table(DocxTable),
    RichTable(DocxRichTable),
    Image(usize),
    SectPr(DocxSectPr),
    PageBreak,
    ColumnBreak,
    RichList(DocxRichList),
    CodeBlock(String),
    TextBox(DocxTextBox),
}

struct CoreProps {
    title: Option<String>,
    author: Option<String>,
    subject: Option<String>,
    keywords: Option<String>,
    description: Option<String>,
    created: Option<String>,
    modified: Option<String>,
}

/// Resolved `r:id` for each hyperlink URL used anywhere in the package.
type HyperlinkRids = std::collections::HashMap<String, String>;

/// Every distinct hyperlink URL reachable from `elements`, in first-seen order.
fn collect_hyperlinks(elements: &[DocxElement], out: &mut Vec<String>) {
    fn push_runs(runs: &[Run], out: &mut Vec<String>) {
        for r in runs {
            if let Some(ref url) = r.hyperlink {
                if !url.is_empty() && !out.contains(url) {
                    out.push(url.clone());
                }
            }
        }
    }
    for e in elements {
        match e {
            DocxElement::RichParagraph(p) => push_runs(&p.runs, out),
            DocxElement::RichList(l) => {
                for item in &l.items {
                    collect_hyperlinks(item, out);
                }
            },
            DocxElement::RichTable(t) => {
                for row in &t.rows {
                    for c in &row.cells {
                        collect_hyperlinks(&c.content, out);
                    }
                }
            },
            DocxElement::TextBox(tb) => collect_hyperlinks(&tb.content, out),
            _ => {},
        }
    }
}

struct ImageInfo {
    idx: usize,
    rid: String,
    width_emu: u64,
    height_emu: u64,
    alt_text: Option<String>,
    decorative: bool,
    positioning: ImagePositioning,
}

// ---------------------------------------------------------------------------
// Builder API
// ---------------------------------------------------------------------------

/// Builder for creating DOCX files from scratch.
pub struct DocxWriter {
    elements: Vec<DocxElement>,
    images: Vec<DocxImage>,
    /// Page background colour, written as `w:background`.
    background_rgb: Option<[u8; 3]>,
    headers_footers: Vec<DocxHf>,
    footnotes: Vec<DocxNote>,
    endnotes: Vec<DocxNote>,
    core_props: Option<CoreProps>,
    next_num_id: u32,
    /// Embedded font programs to ship inside the package under `word/fonts/`.
    /// Each entry is `(font_name, ttf_or_otf_bytes)`. The reader recognizes
    /// these and re-uses them to render any downstream conversion (notably
    /// PDF) so a PDF→DOCX→PDF round-trip preserves typeface fidelity.
    embedded_fonts: Vec<(String, Vec<u8>)>,
}

impl DocxWriter {
    /// Create a new empty DOCX builder.
    pub fn new() -> Self {
        Self {
            elements: Vec::new(),
            images: Vec::new(),
            background_rgb: None,
            headers_footers: Vec::new(),
            footnotes: Vec::new(),
            endnotes: Vec::new(),
            core_props: None,
            next_num_id: 3,
            embedded_fonts: Vec::new(),
        }
    }

    /// Embed a font program (TrueType / OpenType bytes) under `word/fonts/`.
    /// `name` is used for the file name and as the human-readable font name.
    /// Subsequent calls with the same name are deduplicated.
    pub fn embed_font(&mut self, name: impl Into<String>, data: Vec<u8>) -> &mut Self {
        let name = name.into();
        if !self.embedded_fonts.iter().any(|(n, _)| n == &name) {
            self.embedded_fonts.push((name, data));
        }
        self
    }

    /// Add a plain paragraph with the given text.
    pub fn add_paragraph(&mut self, text: &str) -> &mut Self {
        self.elements
            .push(DocxElement::Paragraph(DocxParagraph::plain(text, None, None)));
        self
    }

    /// Add a paragraph with explicit alignment.
    pub fn add_paragraph_aligned(&mut self, text: &str, alignment: Alignment) -> &mut Self {
        self.elements.push(DocxElement::Paragraph(DocxParagraph {
            style: None,
            runs: vec![Run::new(text)],
            numbering: None,
            alignment: Some(alignment),
        }));
        self
    }

    /// Add a paragraph built from styled [`Run`]s.
    pub fn add_rich_paragraph(&mut self, runs: &[Run]) -> &mut Self {
        self.elements.push(DocxElement::Paragraph(DocxParagraph {
            style: None,
            runs: runs.to_vec(),
            numbering: None,
            alignment: None,
        }));
        self
    }

    /// Add a rich paragraph with both custom runs and explicit alignment.
    pub fn add_rich_paragraph_aligned(&mut self, runs: &[Run], alignment: Alignment) -> &mut Self {
        self.elements.push(DocxElement::Paragraph(DocxParagraph {
            style: None,
            runs: runs.to_vec(),
            numbering: None,
            alignment: Some(alignment),
        }));
        self
    }

    /// Add a heading paragraph at the given level (1-6).
    ///
    /// Level is clamped to `1..=6`.
    pub fn add_heading(&mut self, text: &str, level: u8) -> &mut Self {
        let level = level.clamp(1, 6);
        self.elements
            .push(DocxElement::Paragraph(DocxParagraph::plain(
                text,
                Some(format!("Heading{level}")),
                None,
            )));
        self
    }

    /// Add a table. The first row is treated as the header row.
    pub fn add_table(&mut self, rows: &[Vec<&str>]) -> &mut Self {
        let owned: Vec<Vec<String>> = rows
            .iter()
            .map(|row| row.iter().map(|s| s.to_string()).collect())
            .collect();
        self.elements
            .push(DocxElement::Table(DocxTable { rows: owned }));
        self
    }

    /// Add a list of items.
    ///
    /// `ordered = false` → bullet list; `ordered = true` → numbered list.
    pub fn add_list(&mut self, items: &[&str], ordered: bool) -> &mut Self {
        let num_id: u32 = if ordered { 2 } else { 1 };
        for item in items {
            self.elements
                .push(DocxElement::Paragraph(DocxParagraph::plain(
                    item,
                    Some("ListParagraph".to_string()),
                    Some((num_id, 0)),
                )));
        }
        self
    }

    /// Insert a page break.
    pub fn add_page_break(&mut self) -> &mut Self {
        self.elements.push(DocxElement::PageBreak);
        self
    }

    /// Insert a column break.
    pub fn add_column_break(&mut self) -> &mut Self {
        self.elements.push(DocxElement::ColumnBreak);
        self
    }

    // --- Rich IR methods ---

    /// Add a paragraph from IR runs and rich paragraph properties.
    pub fn add_ir_paragraph(&mut self, runs: &[Run], props: Option<IrParaProps>) -> &mut Self {
        self.elements
            .push(DocxElement::RichParagraph(DocxRichParagraph {
                runs: runs.to_vec(),
                props: props.unwrap_or_default(),
            }));
        self
    }

    /// Add a full IR table with borders, column widths, and cell styling.
    pub fn add_ir_table(&mut self, table: &crate::ir::Table) -> &mut Self {
        let rich = convert_ir_table(table);
        self.elements.push(DocxElement::RichTable(rich));
        self
    }

    /// Add an inline image from IR. Skips silently if `image.data` is None.
    pub fn add_ir_image(&mut self, image: &crate::ir::Image) -> &mut Self {
        let data = match &image.data {
            Some(d) => d.clone(),
            None => return self,
        };
        let format = image.format.clone().unwrap_or(ImageFormat::Png);

        let (w_emu, h_emu) =
            if let (Some(w), Some(h)) = (image.display_width_emu, image.display_height_emu) {
                (w, h)
            } else if let (Some(pw), Some(ph)) = (image.pixel_width, image.pixel_height) {
                (px_to_emu(pw), px_to_emu(ph))
            } else {
                (914400u64, 685800u64)
            };

        let idx = self.images.len();
        self.images.push(DocxImage {
            data,
            format,
            display_width_emu: w_emu,
            display_height_emu: h_emu,
            alt_text: image.alt_text.clone(),
            decorative: image.decorative,
            positioning: image.positioning.clone(),
        });
        self.elements.push(DocxElement::Image(idx));
        self
    }

    /// Set section page setup and column layout (appended as `<w:sectPr>` at end of body).
    /// Set the page background colour, written as `w:background`.
    pub fn set_background_rgb(&mut self, rgb: [u8; 3]) -> &mut Self {
        self.background_rgb = Some(rgb);
        self
    }

    /// Set section properties: page geometry, column layout and break type.
    pub fn set_section_props(
        &mut self,
        page_setup: Option<PageSetup>,
        columns: Option<ColumnLayout>,
        break_type: SectionBreakType,
    ) -> &mut Self {
        self.elements.push(DocxElement::SectPr(DocxSectPr {
            page_setup,
            columns,
            break_type,
            header_rids: Vec::new(),
            footer_rids: Vec::new(),
            footnote_rid: None,
            endnote_rid: None,
        }));
        self
    }

    /// Add an IR list with rich style information.
    pub fn add_ir_list(&mut self, list: &crate::ir::List) -> &mut Self {
        self.add_ir_list_at(list, list.level);
        self
    }

    /// Emit `list` and, recursively, every sub-list hanging off its items.
    ///
    /// `ListItem::nested` was read by no writer at all, so everything below
    /// level 0 vanished from the output while the API reported success.
    fn add_ir_list_at(&mut self, list: &crate::ir::List, level: u8) {
        let num_id = self.next_num_id;
        self.next_num_id += 1;
        let start_number = list.start_number.unwrap_or(1);
        let style = list.style.clone();

        let items: Vec<Vec<DocxElement>> = list
            .items
            .iter()
            .map(|item| {
                let mut elems: Vec<DocxElement> = Vec::new();
                for content_elem in &item.content {
                    convert_ir_element_to_docx_elements(content_elem, &mut elems);
                }
                elems
            })
            .collect();

        self.elements.push(DocxElement::RichList(DocxRichList {
            ordered: list.ordered,
            items,
            start_number: if start_number != 1 {
                Some(start_number)
            } else {
                None
            },
            style,
            level,
            num_id,
        }));

        for item in &list.items {
            if let Some(ref nested) = item.nested {
                self.add_ir_list_at(nested, level.saturating_add(1).min(8));
            }
        }
    }

    /// Add a code block.
    pub fn add_code_block(&mut self, content: &str) -> &mut Self {
        self.elements
            .push(DocxElement::CodeBlock(content.to_string()));
        self
    }

    /// Add a text box with IR content.
    pub fn add_text_box(&mut self, tb: &crate::ir::TextBox) -> &mut Self {
        let mut inner: Vec<DocxElement> = Vec::new();
        for elem in &tb.content {
            convert_ir_element_to_docx_elements(elem, &mut inner);
        }
        let width_emu = tb.width_emu.unwrap_or(914400);
        let height_emu = tb.height_emu.unwrap_or(685800);
        self.elements.push(DocxElement::TextBox(DocxTextBox {
            content: inner,
            width_emu,
            height_emu,
            x_emu: tb.x_emu.unwrap_or(0),
            y_emu: tb.y_emu.unwrap_or(0),
            h_anchor: tb.h_anchor.clone(),
            v_anchor: tb.v_anchor.clone(),
            wrap: tb.wrap.clone(),
        }));
        self
    }

    /// Add a footnote with the given ID and IR content.
    pub fn add_footnote(&mut self, id: u32, content: &[crate::ir::Element]) -> &mut Self {
        let mut elems: Vec<DocxElement> = Vec::new();
        for elem in content {
            convert_ir_element_to_docx_elements(elem, &mut elems);
        }
        self.footnotes.push(DocxNote {
            id,
            elements: elems,
        });
        self
    }

    /// Add an endnote with the given ID and IR content.
    pub fn add_endnote(&mut self, id: u32, content: &[crate::ir::Element]) -> &mut Self {
        let mut elems: Vec<DocxElement> = Vec::new();
        for elem in content {
            convert_ir_element_to_docx_elements(elem, &mut elems);
        }
        self.endnotes.push(DocxNote {
            id,
            elements: elems,
        });
        self
    }

    /// Set document metadata (written to `docProps/core.xml`).
    pub fn set_metadata(&mut self, meta: &crate::ir::Metadata) -> &mut Self {
        let keywords = if meta.keywords.is_empty() {
            None
        } else {
            Some(meta.keywords.join(", "))
        };
        self.core_props = Some(CoreProps {
            title: meta.title.clone(),
            author: meta.author.clone(),
            subject: meta.subject.clone(),
            keywords,
            description: meta.description.clone(),
            created: meta.created.clone(),
            modified: meta.modified.clone(),
        });
        self
    }

    /// Add a section header.
    pub fn add_section_header(
        &mut self,
        hf_type: HfType,
        elements: Vec<crate::ir::Element>,
    ) -> &mut Self {
        let mut docx_elems: Vec<DocxElement> = Vec::new();
        for elem in &elements {
            convert_ir_element_to_docx_elements(elem, &mut docx_elems);
        }
        self.headers_footers.push(DocxHf {
            hf_type,
            elements: docx_elems,
        });
        self
    }

    /// Save the document to a file at `path`.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        let opc = OpcWriter::create(path)?;
        self.write_package(opc)?;
        Ok(())
    }

    /// Write the document to an arbitrary `Write + Seek` destination.
    pub fn write_to<W: Write + Seek>(&self, writer: W) -> Result<()> {
        let opc = OpcWriter::new(writer)?;
        self.write_package(opc)?;
        Ok(())
    }

    // -----------------------------------------------------------------------
    // Internals
    // -----------------------------------------------------------------------

    fn write_package<W: Write + Seek>(&self, mut opc: OpcWriter<W>) -> Result<()> {
        let doc_part = PartName::new("/word/document.xml")?;
        let styles_part = PartName::new("/word/styles.xml")?;

        opc.add_package_rel(rel_types::OFFICE_DOCUMENT, "word/document.xml");
        opc.add_part_rel(&doc_part, rel_types::STYLES, "styles.xml");

        // --- Register images ---
        let mut image_rids: Vec<ImageInfo> = Vec::new();
        for (idx, img) in self.images.iter().enumerate() {
            let n = idx + 1;
            let ext = img.format.extension();
            let target = format!("media/image{n}.{ext}");
            let rid = opc.add_part_rel(&doc_part, rel_types::IMAGE, &target);

            let ct = img.format.content_type();
            let part_name = format!("/word/media/image{n}.{ext}");
            let img_part = PartName::new(&part_name)?;
            opc.add_part(&img_part, ct, &img.data)?;

            image_rids.push(ImageInfo {
                idx,
                rid,
                width_emu: img.display_width_emu,
                height_emu: img.display_height_emu,
                alt_text: img.alt_text.clone(),
                decorative: img.decorative,
                positioning: img.positioning.clone(),
            });
        }

        // --- Embed fonts ---
        // Three pieces have to land together so Word/LibreOffice
        // actually pick up the font programs:
        //
        //   1. The TTF/OTF parts under `/word/fonts/font_<n>_<safe>.ttf`.
        //   2. `/word/fontTable.xml` listing each font name with an
        //      `<w:embedRegular r:id="rIdN"/>` reference.
        //   3. `/word/_rels/fontTable.xml.rels` mapping each rId from
        //      step 2 to the matching font part.
        //   4. A relationship in `word/_rels/document.xml.rels` of type
        //      `…/fontTable` so Word knows where to find fontTable.xml.
        //
        // Without all four, the in-process reader still finds the TTFs
        // by directory scan, but Word silently substitutes Calibri.
        if !self.embedded_fonts.is_empty() {
            let font_table_part = PartName::new("/word/fontTable.xml")?;
            opc.add_part_rel(&doc_part, rel_types::FONT_TABLE, "fontTable.xml");

            // Each font part + the fontTable→font rel.
            let mut font_entries: Vec<(String, String)> =
                Vec::with_capacity(self.embedded_fonts.len());
            for (idx, (name, data)) in self.embedded_fonts.iter().enumerate() {
                let n = idx + 1;
                let safe = crate::core::embedded_fonts::sanitize_font_filename(name);
                let target_rel = format!("fonts/font_{n}_{safe}.ttf");
                let target_abs = format!("/word/fonts/font_{n}_{safe}.ttf");
                let part = PartName::new(&target_abs)?;
                opc.add_part(&part, "application/x-font-ttf", data)?;
                let rid = opc.add_part_rel(&font_table_part, rel_types::FONT, &target_rel);
                font_entries.push((name.clone(), rid));
            }

            let xml = generate_font_table_xml(&font_entries);
            opc.add_part(&font_table_part, CT_FONT_TABLE, &xml)?;
        }

        // --- Register headers/footers ---
        // Hyperlinks need one external relationship each, shared by every run
        // that points at the same URL. Collected before any part is written
        // because headers, footers and notes can carry links too.
        let mut hyperlink_rids: HyperlinkRids = HyperlinkRids::new();
        {
            let mut urls: Vec<String> = Vec::new();
            collect_hyperlinks(&self.elements, &mut urls);
            for hf in &self.headers_footers {
                collect_hyperlinks(&hf.elements, &mut urls);
            }
            for n in self.footnotes.iter().chain(self.endnotes.iter()) {
                collect_hyperlinks(&n.elements, &mut urls);
            }
            for url in urls {
                if hyperlink_rids.contains_key(&url) {
                    continue;
                }
                let rid = opc.add_part_rel_with_mode(
                    &doc_part,
                    rel_types::HYPERLINK,
                    &url,
                    crate::core::relationships::TargetMode::External,
                );
                hyperlink_rids.insert(url, rid);
            }
        }

        let mut hf_rids: Vec<(HfType, String)> = Vec::new();
        for (i, hf) in self.headers_footers.iter().enumerate() {
            let n = i + 1;
            let (kind, ct, rel_type) = match hf.hf_type {
                HfType::DefaultHeader | HfType::FirstPageHeader | HfType::EvenPageHeader => {
                    ("header", CT_HEADER, rel_types::HEADER)
                },
                HfType::DefaultFooter | HfType::FirstPageFooter | HfType::EvenPageFooter => {
                    ("footer", CT_FOOTER, rel_types::FOOTER)
                },
            };
            let target = format!("{kind}{n}.xml");
            let rid = opc.add_part_rel(&doc_part, rel_type, &target);
            let part_name = format!("/word/{target}");
            let hf_part = PartName::new(&part_name)?;
            let hf_xml = generate_hf_xml(
                &hf.elements,
                &image_rids,
                matches!(
                    hf.hf_type,
                    HfType::DefaultHeader | HfType::FirstPageHeader | HfType::EvenPageHeader
                ),
                &hyperlink_rids,
            );
            opc.add_part(&hf_part, ct, &hf_xml)?;
            hf_rids.push((hf.hf_type, rid));
        }

        // --- Register footnotes/endnotes ---
        let footnote_rid = if !self.footnotes.is_empty() {
            let notes_part = PartName::new("/word/footnotes.xml")?;
            let rid = opc.add_part_rel(&doc_part, rel_types::FOOTNOTES, "footnotes.xml");
            let xml = generate_footnotes_xml(&self.footnotes, &image_rids, &hyperlink_rids);
            opc.add_part(&notes_part, CT_FOOTNOTES, &xml)?;
            Some(rid)
        } else {
            None
        };

        let endnote_rid = if !self.endnotes.is_empty() {
            let notes_part = PartName::new("/word/endnotes.xml")?;
            let rid = opc.add_part_rel(&doc_part, rel_types::ENDNOTES, "endnotes.xml");
            let xml = generate_endnotes_xml(&self.endnotes, &image_rids, &hyperlink_rids);
            opc.add_part(&notes_part, CT_ENDNOTES, &xml)?;
            Some(rid)
        } else {
            None
        };

        // --- Core properties ---
        if let Some(ref props) = self.core_props {
            let core_part = PartName::new("/docProps/core.xml")?;
            opc.add_package_rel(rel_types::CORE_PROPERTIES, "docProps/core.xml");
            let xml = generate_core_props_xml(props);
            opc.add_part(
                &core_part,
                "application/vnd.openxmlformats-package.core-properties+xml",
                &xml,
            )?;
        }

        // --- Gather sectPr info ---
        let mut sectpr_info: Option<SectPrInfo> = None;
        for elem in &self.elements {
            if let DocxElement::SectPr(sp) = elem {
                sectpr_info = Some(SectPrInfo {
                    page_setup: sp.page_setup.clone(),
                    columns: sp.columns.clone(),
                    break_type: sp.break_type.clone(),
                    hf_rids: hf_rids.clone(),
                    footnote_rid: footnote_rid.clone(),
                    endnote_rid: endnote_rid.clone(),
                });
            }
        }
        if sectpr_info.is_none()
            && (!hf_rids.is_empty() || footnote_rid.is_some() || endnote_rid.is_some())
        {
            sectpr_info = Some(SectPrInfo {
                page_setup: None,
                columns: None,
                break_type: SectionBreakType::Continuous,
                hf_rids: hf_rids.clone(),
                footnote_rid: footnote_rid.clone(),
                endnote_rid: endnote_rid.clone(),
            });
        }

        // --- Generate document ---
        let document_xml = self.generate_document_xml(
            &image_rids,
            sectpr_info.as_ref(),
            !self.images.is_empty(),
            self.has_text_boxes(),
            &hyperlink_rids,
        );
        opc.add_part(&doc_part, CT_DOCUMENT, &document_xml)?;

        let needs_numbering = self.has_lists();
        let styles_xml = generate_styles_xml(
            needs_numbering,
            !self.footnotes.is_empty() || !self.endnotes.is_empty(),
        );
        opc.add_part(&styles_part, CT_STYLES, &styles_xml)?;

        if needs_numbering {
            let numbering_part = PartName::new("/word/numbering.xml")?;
            opc.add_part_rel(&doc_part, rel_types::NUMBERING, "numbering.xml");
            let numbering_xml = self.generate_numbering_xml();
            opc.add_part(&numbering_part, CT_NUMBERING, &numbering_xml)?;
        }

        // settings.xml carries two switches other parts depend on:
        // w:evenAndOddHeaders, without which a w:type="even" header is never
        // shown, and w:embedTrueTypeFonts, without which fontTable.xml's
        // embedded font references are inert.
        let has_even_hf = self
            .headers_footers
            .iter()
            .any(|hf| matches!(hf.hf_type, HfType::EvenPageHeader | HfType::EvenPageFooter));
        let has_fonts = !self.embedded_fonts.is_empty();
        if has_even_hf || has_fonts {
            let settings_part = PartName::new("/word/settings.xml")?;
            opc.add_part_rel(&doc_part, rel_types::SETTINGS, "settings.xml");
            let xml = generate_settings_xml(has_even_hf, has_fonts);
            opc.add_part(&settings_part, CT_SETTINGS, &xml)?;
        }

        opc.finish()?;
        Ok(())
    }

    fn has_text_boxes(&self) -> bool {
        fn check(elements: &[DocxElement]) -> bool {
            elements
                .iter()
                .any(|e| matches!(e, DocxElement::TextBox(_)))
        }
        check(&self.elements)
            || self.headers_footers.iter().any(|hf| check(&hf.elements))
            || self.footnotes.iter().any(|n| check(&n.elements))
            || self.endnotes.iter().any(|n| check(&n.elements))
    }

    fn has_lists(&self) -> bool {
        fn check_elements(elements: &[DocxElement]) -> bool {
            elements.iter().any(|e| match e {
                DocxElement::Paragraph(p) => p.numbering.is_some(),
                DocxElement::RichParagraph(p) => p.props.numbering.is_some(),
                DocxElement::RichList(_) => true,
                DocxElement::RichTable(t) => t
                    .rows
                    .iter()
                    .any(|r| r.cells.iter().any(|c| check_elements(&c.content))),
                DocxElement::TextBox(tb) => check_elements(&tb.content),
                _ => false,
            })
        }
        // A list in a header, footer, footnote or endnote emits ListParagraph
        // and numId=1 just like one in the body. Checking only `elements`
        // meant those parts referenced a numbering part that was never
        // written and a style that was never defined.
        check_elements(&self.elements)
            || self
                .headers_footers
                .iter()
                .any(|hf| check_elements(&hf.elements))
            || self.footnotes.iter().any(|n| check_elements(&n.elements))
            || self.endnotes.iter().any(|n| check_elements(&n.elements))
    }

    fn generate_document_xml(
        &self,
        image_rids: &[ImageInfo],
        sect_pr: Option<&SectPrInfo>,
        has_images: bool,
        has_text_boxes: bool,
        links: &HyperlinkRids,
    ) -> Vec<u8> {
        let mut w = Writer::new_with_indent(Vec::new(), b' ', 2);

        w.write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
            .expect("write decl");

        let mut root = BytesStart::new("w:document");
        root.push_attribute(("xmlns:w", WML_NS));
        root.push_attribute(("xmlns:r", R_NS));
        // Declared unconditionally. Gating these on has_images/has_text_boxes
        // meant a drawing the scan did not reach — one nested in a table cell,
        // for instance — emitted wp:/a:/pic:/wps: with no declaration, making
        // document.xml not well-formed. A content scan that has to stay in
        // step with every nesting site is the wrong shape for this.
        let _ = (has_images, has_text_boxes);
        root.push_attribute(("xmlns:wp", DRAWING_NS));
        root.push_attribute(("xmlns:a", DML_NS));
        root.push_attribute(("xmlns:pic", PIC_NS));
        root.push_attribute(("xmlns:wps", WPS_NS));
        w.write_event(Event::Start(root))
            .expect("write document start");

        // w:background must be the first child of w:document. The IR carried
        // a section background colour that reached no writer, so page colour
        // was silently lost.
        if let Some(rgb) = self.background_rgb {
            let mut bg = BytesStart::new("w:background");
            bg.push_attribute(("w:color", rgb_to_hex(rgb).as_str()));
            w.write_event(Event::Empty(bg)).expect("write background");
        }

        w.write_event(Event::Start(BytesStart::new("w:body")))
            .expect("write body start");

        // Multi-section DOCX: each non-final `<w:sectPr>` lives inside the
        // `<w:pPr>` of a paragraph that terminates that section. Only the
        // final sectPr sits at body level. The previous implementation
        // dropped every non-final SectPr on the floor, so a multi-section
        // IR (e.g. one section per source PDF page from `pdf_to_ir`)
        // collapsed into a single section on the read side and lost all
        // per-page geometry.
        //
        // Find the last `DocxElement::SectPr` index — that's the final
        // section, written at body level. Every earlier SectPr is emitted
        // as a synthetic empty paragraph carrying just `<w:pPr><w:sectPr>…</w:sectPr></w:pPr>`,
        // which `parse_paragraph_properties_fast` recognises and pushes
        // into `body.section_breaks`. `docx_to_ir` then walks
        // `section_breaks` to slice elements into per-section windows.
        let last_sectpr_idx: Option<usize> = self
            .elements
            .iter()
            .rposition(|e| matches!(e, DocxElement::SectPr(_)));

        let mut image_counter = 0u32;
        for (idx, element) in self.elements.iter().enumerate() {
            if let DocxElement::SectPr(sp) = element {
                if Some(idx) == last_sectpr_idx {
                    // Final section is rendered as the body-level sectPr
                    // below (uses the `sect_pr` info already gathered).
                    continue;
                }
                write_inline_section_break_paragraph(&mut w, sp);
                continue;
            }
            write_docx_element(&mut w, element, image_rids, &mut image_counter, links);
        }

        if let Some(sp) = sect_pr {
            write_body_sect_pr(&mut w, sp);
        }

        w.write_event(Event::End(BytesEnd::new("w:body")))
            .expect("write body end");
        w.write_event(Event::End(BytesEnd::new("w:document")))
            .expect("write document end");

        w.into_inner()
    }

    fn generate_numbering_xml(&self) -> Vec<u8> {
        let mut w = Writer::new_with_indent(Vec::new(), b' ', 2);

        w.write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
            .expect("write decl");

        let mut root = BytesStart::new("w:numbering");
        root.push_attribute(("xmlns:w", WML_NS));
        w.write_event(Event::Start(root))
            .expect("write numbering start");

        // CT_Numbering is `numPicBullet*, abstractNum*, num*`: every abstract
        // definition must precede every instance, so these run as two passes
        // rather than one pair per list.
        write_abstract_num(&mut w, 0, "bullet", "\u{2022}");
        write_abstract_num(&mut w, 1, "decimal", "%1.");
        for elem in &self.elements {
            if let DocxElement::RichList(rl) = elem {
                let abstract_id = rl.num_id - 3 + 2;
                let (fmt, lvl_text) = list_style_to_fmt(rl.style.as_ref(), rl.ordered);
                write_abstract_num(&mut w, abstract_id, fmt, lvl_text);
            }
        }

        write_num(&mut w, 1, 0, None);
        write_num(&mut w, 2, 1, None);
        for elem in &self.elements {
            if let DocxElement::RichList(rl) = elem {
                let abstract_id = rl.num_id - 3 + 2;
                write_num(&mut w, rl.num_id, abstract_id, rl.start_number);
            }
        }

        w.write_event(Event::End(BytesEnd::new("w:numbering")))
            .expect("write numbering end");

        w.into_inner()
    }
}

impl Default for DocxWriter {
    fn default() -> Self {
        Self::new()
    }
}

// ---------------------------------------------------------------------------
// Helper to expose HfType to consumers (needed for add_section_header)
// ---------------------------------------------------------------------------

impl HfType {
    /// Returns the default (odd-page) header variant.
    pub fn default_header() -> Self {
        Self::DefaultHeader
    }
    /// Returns the default (odd-page) footer variant.
    pub fn default_footer() -> Self {
        Self::DefaultFooter
    }
    /// Returns the first-page header variant.
    pub fn first_page_header() -> Self {
        Self::FirstPageHeader
    }
    /// Returns the first-page footer variant.
    pub fn first_page_footer() -> Self {
        Self::FirstPageFooter
    }
    /// Returns the even-page header variant.
    pub fn even_page_header() -> Self {
        Self::EvenPageHeader
    }
    /// Returns the even-page footer variant.
    pub fn even_page_footer() -> Self {
        Self::EvenPageFooter
    }
}

// ---------------------------------------------------------------------------
// Convert IR Table → DocxRichTable with vMerge expansion
// ---------------------------------------------------------------------------

fn convert_ir_table(table: &crate::ir::Table) -> DocxRichTable {
    let num_rows = table.rows.len();
    let num_cols = table.rows.iter().map(|r| r.cells.len()).max().unwrap_or(0);

    // Grid to track occupied cells (for vMerge continuation)
    let mut grid: Vec<Vec<bool>> = vec![vec![false; num_cols]; num_rows];

    let mut rich_rows: Vec<DocxRichRow> = Vec::new();

    for (row_idx, row) in table.rows.iter().enumerate() {
        let mut rich_cells: Vec<DocxRichCell> = Vec::new();

        let mut col_cursor = 0usize;
        for cell in &row.cells {
            // Skip occupied grid cells
            while col_cursor < num_cols && grid[row_idx][col_cursor] {
                // Emit a vMerge continuation placeholder
                rich_cells.push(DocxRichCell {
                    content: vec![DocxElement::Paragraph(DocxParagraph::plain("", None, None))],
                    col_span: 1,
                    row_span: 1,
                    background_color: None,
                    border: None,
                    vertical_align: None,
                    text_align: None,
                    text_direction: None,
                    width_twips: None,
                    padding: None,
                    is_vmerge_continue: true,
                });
                col_cursor += 1;
            }

            if col_cursor >= num_cols {
                break;
            }

            // Mark occupied cells
            // Clamp to the grid BEFORE looping. col_span/row_span are u32 and
            // come straight from a parsed document, so with the bounds check
            // inside the loop body the iteration still ran row_span * col_span
            // times — up to 1.8e19 — doing nothing. No allocation, so nothing
            // ever stopped it: the writer simply never returned.
            let col_span = (cell.col_span.max(1) as usize).min(num_cols.saturating_sub(col_cursor));
            let row_span = (cell.row_span.max(1) as usize).min(num_rows.saturating_sub(row_idx));
            for dr in 0..row_span {
                for dc in 0..col_span {
                    grid[row_idx + dr][col_cursor + dc] = true;
                }
            }

            // Convert cell content
            let mut content_elems: Vec<DocxElement> = Vec::new();
            for elem in &cell.content {
                convert_ir_element_to_docx_elements(elem, &mut content_elems);
            }
            if content_elems.is_empty() {
                content_elems.push(DocxElement::Paragraph(DocxParagraph::plain("", None, None)));
            }
            // Apply cell-level text alignment to paragraphs that have no alignment of their own.
            if let Some(ref align) = cell.text_align {
                for elem in &mut content_elems {
                    if let DocxElement::RichParagraph(rp) = elem {
                        if rp.props.alignment.is_none() {
                            rp.props.alignment = Some(align.clone());
                        }
                    }
                }
            }

            rich_cells.push(DocxRichCell {
                content: content_elems,
                col_span: cell.col_span.max(1),
                row_span: cell.row_span.max(1),
                background_color: cell.background_color,
                border: cell.border.clone(),
                vertical_align: cell.vertical_align.clone(),
                text_align: cell.text_align.clone(),
                text_direction: cell.text_direction.clone(),
                width_twips: cell.width_twips,
                padding: cell.padding.clone(),
                is_vmerge_continue: false,
            });

            col_cursor += col_span;
        }

        rich_rows.push(DocxRichRow {
            height_twips: row.height_twips,
            allow_break: row.allow_break,
            repeat_as_header: row.repeat_as_header,
            cells: rich_cells,
        });
    }

    DocxRichTable {
        column_widths_twips: table.column_widths_twips.clone(),
        border: table.border.clone(),
        alignment: table.alignment.clone(),
        cell_padding_twips: table.cell_padding_twips,
        rows: rich_rows,
        width_twips: table.width_twips,
        indent_left_twips: table.indent_left_twips,
        caption: table.caption.clone(),
    }
}

fn convert_ir_element_to_docx_elements(elem: &crate::ir::Element, out: &mut Vec<DocxElement>) {
    // The readers bound nesting with DepthGuard (MAX_NESTING_DEPTH); the
    // writers never did, so a deeply nested IR — and DocumentIR is
    // Deserialize, so it can come from anywhere — overflowed the stack and
    // aborted. An abort is not catchable, so no caller could defend.
    let Some(_guard) = crate::core::xml::DepthGuard::enter() else {
        log::warn!("docx: element nesting exceeds the depth limit; subtree skipped");
        return;
    };
    use crate::ir::Element as E;
    match elem {
        E::Paragraph(p) => {
            let runs = ir_paragraph_to_runs(p);
            let props = ir_paragraph_to_props(p);
            out.push(DocxElement::RichParagraph(DocxRichParagraph { runs, props }));
        },
        E::Heading(h) => {
            let level = h.clamped_level();
            let runs = ir_inline_to_runs(&h.content);
            let props = IrParaProps {
                style: Some(format!("Heading{level}")),
                // A heading nested in a cell, text box, header or note kept
                // its style but lost its alignment, unlike the same heading
                // at top level.
                alignment: h.alignment.clone(),
                ..Default::default()
            };
            out.push(DocxElement::RichParagraph(DocxRichParagraph { runs, props }));
        },
        E::Table(t) => out.push(DocxElement::RichTable(convert_ir_table(t))),
        E::List(l) => {
            // numId 1 -> abstract 0 (bullet), numId 2 -> abstract 1 (decimal).
            // Hardcoding 1 made every ordered list nested in a cell, text box
            // or header render as bullets, disagreeing with the same list at
            // top level.
            let num_id = if l.ordered { 2u32 } else { 1u32 };
            for item in &l.items {
                for content_elem in &item.content {
                    if let E::Paragraph(p) = content_elem {
                        let runs = ir_paragraph_to_runs(p);
                        let mut props = ir_paragraph_to_props(p);
                        props.style = Some("ListParagraph".to_string());
                        props.numbering = Some((num_id, l.level));
                        out.push(DocxElement::RichParagraph(DocxRichParagraph { runs, props }));
                    }
                }
                // Sub-lists were dropped entirely: everything below level 0
                // never reached the file.
                if let Some(ref nested) = item.nested {
                    convert_ir_element_to_docx_elements(&E::List(nested.clone()), out);
                }
            }
        },
        E::Image(_img) => {
            // Image embedding requires the outer DocxWriter context for index tracking.
            // Skip in nested contexts (table cells, text boxes, headers).
        },
        E::ThematicBreak => {},
        E::PageBreak => out.push(DocxElement::PageBreak),
        E::ColumnBreak => out.push(DocxElement::ColumnBreak),
        E::TextBox(tb) => {
            let mut inner: Vec<DocxElement> = Vec::new();
            for e in &tb.content {
                convert_ir_element_to_docx_elements(e, &mut inner);
            }
            out.push(DocxElement::TextBox(DocxTextBox {
                content: inner,
                width_emu: tb.width_emu.unwrap_or(914400),
                height_emu: tb.height_emu.unwrap_or(685800),
                x_emu: tb.x_emu.unwrap_or(0),
                y_emu: tb.y_emu.unwrap_or(0),
                h_anchor: tb.h_anchor.clone(),
                v_anchor: tb.v_anchor.clone(),
                wrap: tb.wrap.clone(),
            }));
        },
        E::Footnote(_) | E::Endnote(_) => {},
        E::CodeBlock(cb) => out.push(DocxElement::CodeBlock(cb.content.clone())),
        E::Shape(_) => {
            // Vector shapes are emitted by the layout-preserving DOCX
            // writer in pdf_oxide directly; the markdown-driven IR
            // writer doesn't have anywhere to put them yet.
        },
    }
}

fn ir_paragraph_to_runs(p: &crate::ir::Paragraph) -> Vec<Run> {
    ir_inline_to_runs(&p.content)
}

fn ir_inline_to_runs(content: &[crate::ir::InlineContent]) -> Vec<Run> {
    use crate::ir::InlineContent;
    let mut runs: Vec<Run> = Vec::new();
    for item in content {
        match item {
            InlineContent::Text(span) => {
                let mut run = Run::new(&span.text);
                run.bold = span.bold;
                run.italic = span.italic;
                run.strikethrough = span.strikethrough;
                run.font_name = span.font_name.clone();
                run.font_size_half_pt = span.font_size_half_pt;
                run.color_rgb = span.color;
                run.underline_style = span.underline.clone();
                run.highlight = span.highlight;
                run.vertical_align = span.vertical_align.clone();
                run.all_caps = span.all_caps;
                run.small_caps = span.small_caps;
                run.char_spacing_half_pt = span.char_spacing_half_pt;
                run.hyperlink = span.hyperlink.clone();
                runs.push(run);
            },
            InlineContent::LineBreak => {
                runs.push(Run {
                    text: "\n".to_string(),
                    ..Default::default()
                });
            },
            InlineContent::FootnoteRef(r) => {
                let run = Run {
                    footnote_ref: Some(r.note_id),
                    ..Default::default()
                };
                runs.push(run);
            },
            InlineContent::EndnoteRef(r) => {
                let run = Run {
                    endnote_ref: Some(r.note_id),
                    ..Default::default()
                };
                runs.push(run);
            },
        }
    }
    runs
}

fn ir_paragraph_to_props(p: &crate::ir::Paragraph) -> IrParaProps {
    IrParaProps {
        alignment: p.alignment.clone(),
        indent_left_twips: p.indent_left_twips,
        indent_right_twips: p.indent_right_twips,
        first_line_indent_twips: p.first_line_indent_twips,
        space_before_twips: p.space_before_twips,
        space_after_twips: p.space_after_twips,
        line_spacing: p.line_spacing.clone(),
        style: None,
        numbering: None,
        keep_with_next: p.keep_with_next,
        keep_together: p.keep_together,
        page_break_before: p.page_break_before,
        background_color: p.background_color,
        outline_level: p.outline_level,
        border: p.border.clone(),
        tabs: p.tabs.clone(),
        frame_position: p.frame_position.clone(),
    }
}

// ---------------------------------------------------------------------------
// SectPr info helper
// ---------------------------------------------------------------------------

#[allow(dead_code)]
struct SectPrInfo {
    page_setup: Option<PageSetup>,
    columns: Option<ColumnLayout>,
    break_type: SectionBreakType,
    hf_rids: Vec<(HfType, String)>,
    footnote_rid: Option<String>,
    endnote_rid: Option<String>,
}

// ---------------------------------------------------------------------------
// XML serialisation helpers
// ---------------------------------------------------------------------------

fn write_docx_element(
    w: &mut Writer<Vec<u8>>,
    elem: &DocxElement,
    image_rids: &[ImageInfo],
    image_counter: &mut u32,
    links: &HyperlinkRids,
) {
    let Some(_guard) = crate::core::xml::DepthGuard::enter() else {
        log::warn!("docx: element nesting exceeds the depth limit; subtree skipped");
        return;
    };
    match elem {
        DocxElement::Paragraph(p) => write_paragraph(w, p),
        DocxElement::RichParagraph(p) => write_rich_paragraph(w, p, links),
        DocxElement::Table(t) => write_table(w, t),
        DocxElement::RichTable(t) => write_rich_table(w, t, image_rids, image_counter, links),
        DocxElement::Image(idx) => {
            *image_counter += 1;
            if let Some(info) = image_rids.iter().find(|i| i.idx == *idx) {
                match &info.positioning {
                    ImagePositioning::Inline => {
                        write_inline_image_run(
                            w,
                            &info.rid,
                            info.width_emu,
                            info.height_emu,
                            info.alt_text.as_deref(),
                            info.decorative,
                            *image_counter,
                        );
                    },
                    ImagePositioning::Floating(fi) => {
                        write_floating_image_run(
                            w,
                            &info.rid,
                            fi,
                            info.alt_text.as_deref(),
                            info.decorative,
                            *image_counter,
                        );
                    },
                }
            }
        },
        DocxElement::SectPr(_) => {},
        DocxElement::PageBreak => write_page_break(w),
        DocxElement::ColumnBreak => write_column_break(w),
        DocxElement::RichList(rl) => write_rich_list(w, rl, image_rids, image_counter, links),
        DocxElement::CodeBlock(content) => write_code_block(w, content),
        DocxElement::TextBox(tb) => write_text_box(w, tb, image_rids, image_counter, links),
    }
}

fn write_paragraph(w: &mut Writer<Vec<u8>>, p: &DocxParagraph) {
    w.write_event(Event::Start(BytesStart::new("w:p")))
        .expect("write p start");

    let has_ppr = p.style.is_some() || p.numbering.is_some() || p.alignment.is_some();
    if has_ppr {
        w.write_event(Event::Start(BytesStart::new("w:pPr")))
            .expect("write pPr start");

        if let Some(ref style) = p.style {
            let mut elem = BytesStart::new("w:pStyle");
            elem.push_attribute(("w:val", style.as_str()));
            w.write_event(Event::Empty(elem)).expect("write pStyle");
        }

        if let Some(ref align) = p.alignment {
            let mut elem = BytesStart::new("w:jc");
            elem.push_attribute(("w:val", align.as_wml_val()));
            w.write_event(Event::Empty(elem)).expect("write jc");
        }

        if let Some((num_id, ilvl)) = p.numbering {
            write_num_pr(w, num_id, ilvl);
        }

        w.write_event(Event::End(BytesEnd::new("w:pPr")))
            .expect("write pPr end");
    }

    for run in &p.runs {
        write_run(w, run);
    }

    w.write_event(Event::End(BytesEnd::new("w:p")))
        .expect("write p end");
}

/// `w:val` for a tab stop's alignment.
fn tab_align_val(a: &crate::ir::TabAlignment) -> &'static str {
    use crate::ir::TabAlignment;
    match a {
        TabAlignment::Left => "left",
        TabAlignment::Center => "center",
        TabAlignment::Right => "right",
        TabAlignment::Decimal => "decimal",
        TabAlignment::Bar => "bar",
    }
}

/// `w:leader`, or `None` when the stop has no leader (the default).
fn tab_leader_val(l: &crate::ir::TabLeader) -> Option<&'static str> {
    use crate::ir::TabLeader;
    match l {
        TabLeader::None => None,
        TabLeader::Dot => Some("dot"),
        TabLeader::Hyphen => Some("hyphen"),
        TabLeader::Underscore => Some("underscore"),
        TabLeader::Heavy => Some("heavy"),
        TabLeader::MiddleDot => Some("middleDot"),
    }
}

fn write_rich_paragraph(w: &mut Writer<Vec<u8>>, p: &DocxRichParagraph, links: &HyperlinkRids) {
    w.write_event(Event::Start(BytesStart::new("w:p")))
        .expect("write p start");

    let props = &p.props;
    let has_ppr = props.style.is_some()
        || props.numbering.is_some()
        || props.alignment.is_some()
        || props.indent_left_twips.is_some()
        || props.indent_right_twips.is_some()
        || props.first_line_indent_twips.is_some()
        || props.space_before_twips.is_some()
        || props.space_after_twips.is_some()
        || props.line_spacing.is_some()
        || props.keep_with_next
        || props.keep_together
        || props.page_break_before
        || props.background_color.is_some()
        || props.outline_level.is_some()
        || props.border.is_some()
        || !props.tabs.is_empty()
        || props.frame_position.is_some();

    if has_ppr {
        w.write_event(Event::Start(BytesStart::new("w:pPr")))
            .expect("write pPr start");

        if let Some(ref style) = props.style {
            let mut elem = BytesStart::new("w:pStyle");
            elem.push_attribute(("w:val", style.as_str()));
            w.write_event(Event::Empty(elem)).expect("write pStyle");
        }

        if props.keep_with_next {
            w.write_event(Event::Empty(BytesStart::new("w:keepNext")))
                .expect("write keepNext");
        }
        if props.keep_together {
            w.write_event(Event::Empty(BytesStart::new("w:keepLines")))
                .expect("write keepLines");
        }
        if props.page_break_before {
            w.write_event(Event::Empty(BytesStart::new("w:pageBreakBefore")))
                .expect("write pageBreakBefore");
        }
        // CT_PPrBase puts framePr straight after pageBreakBefore.
        if let Some(ref fp) = props.frame_position {
            let mut frame = BytesStart::new("w:framePr");
            frame.push_attribute(("w:w", fp.width_twips.to_string().as_str()));
            frame.push_attribute(("w:h", fp.height_twips.to_string().as_str()));
            frame.push_attribute(("w:hAnchor", "page"));
            frame.push_attribute(("w:vAnchor", "page"));
            frame.push_attribute(("w:x", fp.x_twips.to_string().as_str()));
            frame.push_attribute(("w:y", fp.y_twips.to_string().as_str()));
            w.write_event(Event::Empty(frame)).expect("write framePr");
        }

        // CT_PPrBase is a strict sequence. The order below follows it:
        // numPr, pBdr, shd, tabs, spacing, ind, jc, outlineLvl. Emitting
        // these in a different order produces a schema-invalid document.
        if let Some((num_id, ilvl)) = props.numbering {
            write_num_pr(w, num_id, ilvl);
        }

        if let Some(ref pbdr) = props.border {
            write_paragraph_borders(w, pbdr);
        }

        if let Some(ref color) = props.background_color {
            let mut shd = BytesStart::new("w:shd");
            shd.push_attribute(("w:val", "clear"));
            shd.push_attribute(("w:fill", rgb_to_hex(*color).as_str()));
            shd.push_attribute(("w:color", "auto"));
            w.write_event(Event::Empty(shd)).expect("write pShd");
        }

        if !props.tabs.is_empty() {
            w.write_event(Event::Start(BytesStart::new("w:tabs")))
                .expect("write tabs start");
            for stop in &props.tabs {
                let mut t = BytesStart::new("w:tab");
                t.push_attribute(("w:val", tab_align_val(&stop.alignment)));
                t.push_attribute(("w:pos", stop.position_twips.to_string().as_str()));
                if let Some(leader) = tab_leader_val(&stop.leader) {
                    t.push_attribute(("w:leader", leader));
                }
                w.write_event(Event::Empty(t)).expect("write tab");
            }
            w.write_event(Event::End(BytesEnd::new("w:tabs")))
                .expect("write tabs end");
        }

        // Spacing
        let has_spacing = props.space_before_twips.is_some()
            || props.space_after_twips.is_some()
            || props.line_spacing.is_some();
        if has_spacing {
            let mut sp = BytesStart::new("w:spacing");
            if let Some(v) = props.space_before_twips {
                sp.push_attribute(("w:before", v.to_string().as_str()));
            }
            if let Some(v) = props.space_after_twips {
                sp.push_attribute(("w:after", v.to_string().as_str()));
            }
            match &props.line_spacing {
                Some(LineSpacing::Auto(v)) | Some(LineSpacing::Multiple(v)) => {
                    sp.push_attribute(("w:line", v.to_string().as_str()));
                    sp.push_attribute(("w:lineRule", "auto"));
                },
                Some(LineSpacing::Exact(v)) => {
                    sp.push_attribute(("w:line", v.to_string().as_str()));
                    sp.push_attribute(("w:lineRule", "exact"));
                },
                Some(LineSpacing::AtLeast(v)) => {
                    sp.push_attribute(("w:line", v.to_string().as_str()));
                    sp.push_attribute(("w:lineRule", "atLeast"));
                },
                None => {},
            }
            w.write_event(Event::Empty(sp)).expect("write spacing");
        }

        // Indent
        let has_indent = props.indent_left_twips.is_some()
            || props.indent_right_twips.is_some()
            || props.first_line_indent_twips.is_some();
        if has_indent {
            let mut ind = BytesStart::new("w:ind");
            if let Some(v) = props.indent_left_twips {
                ind.push_attribute(("w:left", v.to_string().as_str()));
            }
            if let Some(v) = props.indent_right_twips {
                ind.push_attribute(("w:right", v.to_string().as_str()));
            }
            if let Some(v) = props.first_line_indent_twips {
                if v >= 0 {
                    ind.push_attribute(("w:firstLine", v.to_string().as_str()));
                } else {
                    // unsigned_abs, not -v: negating i32::MIN overflows, and
                    // w:hanging is ST_TwipsMeasure (unsigned) anyway, so the
                    // magnitude is what the attribute wants.
                    ind.push_attribute(("w:hanging", v.unsigned_abs().to_string().as_str()));
                }
            }
            w.write_event(Event::Empty(ind)).expect("write ind");
        }

        if let Some(align) = &props.alignment {
            let mut elem = BytesStart::new("w:jc");
            elem.push_attribute(("w:val", para_align_val(align)));
            w.write_event(Event::Empty(elem)).expect("write jc");
        }

        if let Some(level) = props.outline_level {
            let mut lvl = BytesStart::new("w:outlineLvl");
            lvl.push_attribute(("w:val", level.to_string().as_str()));
            w.write_event(Event::Empty(lvl)).expect("write outlineLvl");
        }
        w.write_event(Event::End(BytesEnd::new("w:pPr")))
            .expect("write pPr end");
    }

    // A run carrying a URL is wrapped in w:hyperlink pointing at the external
    // relationship collected for it. Adjacent runs sharing a URL share one
    // wrapper, so a styled link stays a single link.
    let mut i = 0usize;
    while i < p.runs.len() {
        let url = p.runs[i].hyperlink.as_deref().filter(|u| !u.is_empty());
        match url.and_then(|u| links.get(u)) {
            Some(rid) => {
                let mut link = BytesStart::new("w:hyperlink");
                link.push_attribute(("r:id", rid.as_str()));
                w.write_event(Event::Start(link)).expect("write hyperlink");
                while i < p.runs.len()
                    && p.runs[i].hyperlink.as_deref().filter(|u| !u.is_empty()) == url
                {
                    write_run(w, &p.runs[i]);
                    i += 1;
                }
                w.write_event(Event::End(BytesEnd::new("w:hyperlink")))
                    .expect("write hyperlink end");
            },
            None => {
                write_run(w, &p.runs[i]);
                i += 1;
            },
        }
    }

    w.write_event(Event::End(BytesEnd::new("w:p")))
        .expect("write p end");
}

fn write_num_pr(w: &mut Writer<Vec<u8>>, num_id: u32, ilvl: u8) {
    w.write_event(Event::Start(BytesStart::new("w:numPr")))
        .expect("write numPr start");
    let mut ilvl_elem = BytesStart::new("w:ilvl");
    ilvl_elem.push_attribute(("w:val", ilvl.to_string().as_str()));
    w.write_event(Event::Empty(ilvl_elem)).expect("write ilvl");
    let mut num_id_elem = BytesStart::new("w:numId");
    num_id_elem.push_attribute(("w:val", num_id.to_string().as_str()));
    w.write_event(Event::Empty(num_id_elem))
        .expect("write numId");
    w.write_event(Event::End(BytesEnd::new("w:numPr")))
        .expect("write numPr end");
}

fn write_rpr(w: &mut Writer<Vec<u8>>, run: &Run) {
    w.write_event(Event::Start(BytesStart::new("w:rPr")))
        .expect("write rPr start");
    write_rpr_content(w, run);
    w.write_event(Event::End(BytesEnd::new("w:rPr")))
        .expect("write rPr end");
}

fn write_rpr_content(w: &mut Writer<Vec<u8>>, run: &Run) {
    if let Some(ref name) = run.font_name {
        let mut elem = BytesStart::new("w:rFonts");
        elem.push_attribute(("w:ascii", name.as_str()));
        elem.push_attribute(("w:hAnsi", name.as_str()));
        elem.push_attribute(("w:cs", name.as_str()));
        elem.push_attribute(("w:eastAsia", name.as_str()));
        w.write_event(Event::Empty(elem)).expect("write rFonts");
    }
    if run.bold {
        w.write_event(Event::Empty(BytesStart::new("w:b")))
            .expect("write bold");
    }
    if run.italic {
        w.write_event(Event::Empty(BytesStart::new("w:i")))
            .expect("write italic");
    }
    if let Some(ref us) = run.underline_style {
        let val = underline_style_val(us);
        let mut elem = BytesStart::new("w:u");
        elem.push_attribute(("w:val", val));
        w.write_event(Event::Empty(elem)).expect("write u style");
    } else if run.underline {
        let mut elem = BytesStart::new("w:u");
        elem.push_attribute(("w:val", "single"));
        w.write_event(Event::Empty(elem)).expect("write underline");
    }
    if run.strikethrough {
        w.write_event(Event::Empty(BytesStart::new("w:strike")))
            .expect("write strike");
    }
    if let Some(ref rgb) = run.color_rgb {
        let mut elem = BytesStart::new("w:color");
        elem.push_attribute(("w:val", rgb_to_hex(*rgb).as_str()));
        w.write_event(Event::Empty(elem)).expect("write color rgb");
    } else if let Some(ref hex) = run.color {
        let mut elem = BytesStart::new("w:color");
        elem.push_attribute(("w:val", hex.as_str()));
        w.write_event(Event::Empty(elem)).expect("write color");
    }
    if let Some(hp) = run.font_size_half_pt {
        let val = hp.to_string();
        let mut sz = BytesStart::new("w:sz");
        sz.push_attribute(("w:val", val.as_str()));
        w.write_event(Event::Empty(sz)).expect("write sz");
        let mut sz_cs = BytesStart::new("w:szCs");
        sz_cs.push_attribute(("w:val", val.as_str()));
        w.write_event(Event::Empty(sz_cs)).expect("write szCs");
    } else if let Some(pt) = run.font_size_pt {
        let half_pts = (pt * 2.0).round() as u32;
        let val = half_pts.to_string();
        let mut sz = BytesStart::new("w:sz");
        sz.push_attribute(("w:val", val.as_str()));
        w.write_event(Event::Empty(sz)).expect("write sz");
        let mut sz_cs = BytesStart::new("w:szCs");
        sz_cs.push_attribute(("w:val", val.as_str()));
        w.write_event(Event::Empty(sz_cs)).expect("write szCs");
    }
    if let Some(ref hl) = run.highlight {
        let mut shd = BytesStart::new("w:shd");
        shd.push_attribute(("w:val", "clear"));
        shd.push_attribute(("w:fill", rgb_to_hex(*hl).as_str()));
        shd.push_attribute(("w:color", "auto"));
        w.write_event(Event::Empty(shd)).expect("write rShd");
    }
    if let Some(ref va) = run.vertical_align {
        let val = match va {
            VerticalAlign::Superscript => "superscript",
            VerticalAlign::Subscript => "subscript",
            VerticalAlign::Baseline => "baseline",
        };
        let mut elem = BytesStart::new("w:vertAlign");
        elem.push_attribute(("w:val", val));
        w.write_event(Event::Empty(elem)).expect("write vertAlign");
    }
    if run.all_caps {
        w.write_event(Event::Empty(BytesStart::new("w:caps")))
            .expect("write caps");
    }
    if run.small_caps {
        w.write_event(Event::Empty(BytesStart::new("w:smallCaps")))
            .expect("write smallCaps");
    }
    if let Some(spacing) = run.char_spacing_half_pt {
        let mut elem = BytesStart::new("w:spacing");
        elem.push_attribute(("w:val", spacing.to_string().as_str()));
        w.write_event(Event::Empty(elem))
            .expect("write char spacing");
    }
}

fn write_field_run(w: &mut Writer<Vec<u8>>, run: &Run, instr: &str) {
    // begin run
    w.write_event(Event::Start(BytesStart::new("w:r")))
        .expect("write r start");
    if run.has_rpr() {
        write_rpr(w, run);
    }
    let mut fc = BytesStart::new("w:fldChar");
    fc.push_attribute(("w:fldCharType", "begin"));
    w.write_event(Event::Empty(fc))
        .expect("write fldChar begin");
    w.write_event(Event::End(BytesEnd::new("w:r")))
        .expect("write r end");
    // instrText run
    w.write_event(Event::Start(BytesStart::new("w:r")))
        .expect("write r start");
    if run.has_rpr() {
        write_rpr(w, run);
    }
    let mut it = BytesStart::new("w:instrText");
    it.push_attribute(("xml:space", "preserve"));
    w.write_event(Event::Start(it))
        .expect("write instrText start");
    w.write_event(Event::Text(BytesText::new(&crate::core::xml::sanitize_xml_text(instr))))
        .expect("write instrText");
    w.write_event(Event::End(BytesEnd::new("w:instrText")))
        .expect("write instrText end");
    w.write_event(Event::End(BytesEnd::new("w:r")))
        .expect("write r end");
    // end run
    w.write_event(Event::Start(BytesStart::new("w:r")))
        .expect("write r start");
    if run.has_rpr() {
        write_rpr(w, run);
    }
    let mut fc = BytesStart::new("w:fldChar");
    fc.push_attribute(("w:fldCharType", "end"));
    w.write_event(Event::Empty(fc)).expect("write fldChar end");
    w.write_event(Event::End(BytesEnd::new("w:r")))
        .expect("write r end");
}

fn write_run(w: &mut Writer<Vec<u8>>, run: &Run) {
    if let Some(note_id) = run.footnote_ref {
        write_footnote_ref_run(w, note_id, false);
        return;
    }
    if let Some(note_id) = run.endnote_ref {
        write_footnote_ref_run(w, note_id, true);
        return;
    }

    let instr = match run.text.as_str() {
        "{PAGE}" => Some(" PAGE "),
        "{NUMPAGES}" => Some(" NUMPAGES "),
        _ => None,
    };
    if let Some(instr_text) = instr {
        write_field_run(w, run, instr_text);
        return;
    }

    w.write_event(Event::Start(BytesStart::new("w:r")))
        .expect("write r start");
    if run.has_rpr() {
        write_rpr(w, run);
    }
    if run.text == "\n" {
        w.write_event(Event::Empty(BytesStart::new("w:br")))
            .expect("write br");
    } else {
        let text = &run.text;
        let mut t_elem = BytesStart::new("w:t");
        if text.starts_with(' ') || text.ends_with(' ') || text.contains("  ") {
            t_elem.push_attribute(("xml:space", "preserve"));
        }
        w.write_event(Event::Start(t_elem)).expect("write t start");
        w.write_event(Event::Text(BytesText::new(&crate::core::xml::sanitize_xml_text(text))))
            .expect("write text");
        w.write_event(Event::End(BytesEnd::new("w:t")))
            .expect("write t end");
    }
    w.write_event(Event::End(BytesEnd::new("w:r")))
        .expect("write r end");
}

fn write_footnote_ref_run(w: &mut Writer<Vec<u8>>, note_id: u32, is_endnote: bool) {
    w.write_event(Event::Start(BytesStart::new("w:r")))
        .expect("write r start");
    w.write_event(Event::Start(BytesStart::new("w:rPr")))
        .expect("write rPr start");
    let style_name = if is_endnote {
        "EndnoteReference"
    } else {
        "FootnoteReference"
    };
    let mut r_style = BytesStart::new("w:rStyle");
    r_style.push_attribute(("w:val", style_name));
    w.write_event(Event::Empty(r_style)).expect("write rStyle");
    w.write_event(Event::End(BytesEnd::new("w:rPr")))
        .expect("write rPr end");
    let tag = if is_endnote {
        "w:endnoteReference"
    } else {
        "w:footnoteReference"
    };
    let mut ref_elem = BytesStart::new(tag);
    ref_elem.push_attribute(("w:id", note_id.to_string().as_str()));
    w.write_event(Event::Empty(ref_elem))
        .expect("write note ref");
    w.write_event(Event::End(BytesEnd::new("w:r")))
        .expect("write r end");
}

fn write_page_break(w: &mut Writer<Vec<u8>>) {
    w.write_event(Event::Start(BytesStart::new("w:p")))
        .expect("write p start");
    w.write_event(Event::Start(BytesStart::new("w:r")))
        .expect("write r start");
    let mut br = BytesStart::new("w:br");
    br.push_attribute(("w:type", "page"));
    w.write_event(Event::Empty(br)).expect("write br");
    w.write_event(Event::End(BytesEnd::new("w:r")))
        .expect("write r end");
    w.write_event(Event::End(BytesEnd::new("w:p")))
        .expect("write p end");
}

fn write_column_break(w: &mut Writer<Vec<u8>>) {
    w.write_event(Event::Start(BytesStart::new("w:p")))
        .expect("write p start");
    w.write_event(Event::Start(BytesStart::new("w:r")))
        .expect("write r start");
    let mut br = BytesStart::new("w:br");
    br.push_attribute(("w:type", "column"));
    w.write_event(Event::Empty(br)).expect("write column br");
    w.write_event(Event::End(BytesEnd::new("w:r")))
        .expect("write r end");
    w.write_event(Event::End(BytesEnd::new("w:p")))
        .expect("write p end");
}

/// Write `w:tblGrid`, which `CT_Tbl` requires. Uses `widths` when known and
/// otherwise emits `col_count` auto-width columns.
fn write_tbl_grid(w: &mut Writer<Vec<u8>>, widths: &[u32], col_count: usize) {
    w.write_event(Event::Start(BytesStart::new("w:tblGrid")))
        .expect("write tblGrid start");
    if widths.is_empty() {
        for _ in 0..col_count {
            let mut gc = BytesStart::new("w:gridCol");
            gc.push_attribute(("w:w", "0"));
            w.write_event(Event::Empty(gc)).expect("write gridCol");
        }
    } else {
        for &cw in widths {
            let mut gc = BytesStart::new("w:gridCol");
            gc.push_attribute(("w:w", cw.to_string().as_str()));
            w.write_event(Event::Empty(gc)).expect("write gridCol");
        }
    }
    w.write_event(Event::End(BytesEnd::new("w:tblGrid")))
        .expect("write tblGrid end");
}

fn write_table(w: &mut Writer<Vec<u8>>, table: &DocxTable) {
    w.write_event(Event::Start(BytesStart::new("w:tbl")))
        .expect("write tbl start");

    // CT_Tbl requires tblPr and tblGrid before the rows.
    w.write_event(Event::Start(BytesStart::new("w:tblPr")))
        .expect("write tblPr start");
    let mut tbl_w = BytesStart::new("w:tblW");
    tbl_w.push_attribute(("w:w", "0"));
    tbl_w.push_attribute(("w:type", "auto"));
    w.write_event(Event::Empty(tbl_w)).expect("write tblW");
    w.write_event(Event::End(BytesEnd::new("w:tblPr")))
        .expect("write tblPr end");
    write_tbl_grid(w, &[], table.rows.iter().map(|r| r.len()).max().unwrap_or(0));

    for row in &table.rows {
        w.write_event(Event::Start(BytesStart::new("w:tr")))
            .expect("write tr start");

        for cell_text in row {
            w.write_event(Event::Start(BytesStart::new("w:tc")))
                .expect("write tc start");
            let p = DocxParagraph::plain(cell_text, None, None);
            write_paragraph(w, &p);
            w.write_event(Event::End(BytesEnd::new("w:tc")))
                .expect("write tc end");
        }

        w.write_event(Event::End(BytesEnd::new("w:tr")))
            .expect("write tr end");
    }

    w.write_event(Event::End(BytesEnd::new("w:tbl")))
        .expect("write tbl end");
}

fn write_rich_table(
    w: &mut Writer<Vec<u8>>,
    table: &DocxRichTable,
    image_rids: &[ImageInfo],
    image_counter: &mut u32,
    links: &HyperlinkRids,
) {
    if let Some(ref caption) = table.caption {
        let p = DocxParagraph::plain(caption, Some("Caption".to_string()), None);
        write_paragraph(w, &p);
    }

    w.write_event(Event::Start(BytesStart::new("w:tbl")))
        .expect("write tbl start");

    // tblPr
    w.write_event(Event::Start(BytesStart::new("w:tblPr")))
        .expect("write tblPr start");

    // CT_TblPrBase is a strict sequence: tblW, jc, tblInd, tblBorders,
    // shd, tblCellMar, tblCaption. Emitting tblCaption first and jc after
    // the borders made every table carrying either one invalid.
    let mut tbl_w = BytesStart::new("w:tblW");
    if let Some(w_twips) = table.width_twips {
        tbl_w.push_attribute(("w:w", w_twips.to_string().as_str()));
        tbl_w.push_attribute(("w:type", "dxa"));
    } else {
        tbl_w.push_attribute(("w:w", "0"));
        tbl_w.push_attribute(("w:type", "auto"));
    }
    w.write_event(Event::Empty(tbl_w)).expect("write tblW");

    if let Some(align) = &table.alignment {
        let val = match align {
            TableAlignment::Left => "left",
            TableAlignment::Center => "center",
            TableAlignment::Right => "right",
        };
        let mut jc = BytesStart::new("w:jc");
        jc.push_attribute(("w:val", val));
        w.write_event(Event::Empty(jc)).expect("write tbl jc");
    }

    if let Some(ind) = table.indent_left_twips {
        let mut tbl_ind = BytesStart::new("w:tblInd");
        tbl_ind.push_attribute(("w:w", ind.to_string().as_str()));
        tbl_ind.push_attribute(("w:type", "dxa"));
        w.write_event(Event::Empty(tbl_ind)).expect("write tblInd");
    }

    if let Some(ref border) = table.border {
        write_table_borders(w, border, "w:tblBorders");
    }

    if let Some(pad) = table.cell_padding_twips {
        let pad_str = pad.to_string();
        w.write_event(Event::Start(BytesStart::new("w:tblCellMar")))
            .expect("write tblCellMar start");
        for side in &["w:top", "w:left", "w:bottom", "w:right"] {
            let mut elem = BytesStart::new(*side);
            elem.push_attribute(("w:w", pad_str.as_str()));
            elem.push_attribute(("w:type", "dxa"));
            w.write_event(Event::Empty(elem))
                .expect("write cell margin");
        }
        w.write_event(Event::End(BytesEnd::new("w:tblCellMar")))
            .expect("write tblCellMar end");
    }
    // w:tblCaption is what the reader looks for. Emitting the caption only as
    // a Caption-styled paragraph lost it on round-trip and accumulated a
    // phantom body paragraph on every cycle.
    if let Some(ref caption) = table.caption {
        let mut cap = BytesStart::new("w:tblCaption");
        cap.push_attribute(("w:val", caption.as_str()));
        w.write_event(Event::Empty(cap)).expect("write tblCaption");
    }
    w.write_event(Event::End(BytesEnd::new("w:tblPr")))
        .expect("write tblPr end");

    // tblGrid is required by CT_Tbl (minOccurs=1), so it is written even when
    // no explicit widths are known — markdown tables carry none. A zero width
    // means "auto", which is what Word writes for an unsized column.
    write_tbl_grid(
        w,
        &table.column_widths_twips,
        table.rows.iter().map(|r| r.cells.len()).max().unwrap_or(0),
    );

    for row in &table.rows {
        w.write_event(Event::Start(BytesStart::new("w:tr")))
            .expect("write tr start");

        let has_tr_pr = row.height_twips.is_some() || row.repeat_as_header || !row.allow_break;
        if has_tr_pr {
            w.write_event(Event::Start(BytesStart::new("w:trPr")))
                .expect("write trPr start");
            if let Some(h) = row.height_twips {
                let mut trh = BytesStart::new("w:trHeight");
                trh.push_attribute(("w:val", h.to_string().as_str()));
                w.write_event(Event::Empty(trh)).expect("write trHeight");
            }
            if row.repeat_as_header {
                w.write_event(Event::Empty(BytesStart::new("w:tblHeader")))
                    .expect("write tblHeader");
            }
            if !row.allow_break {
                w.write_event(Event::Empty(BytesStart::new("w:cantSplit")))
                    .expect("write cantSplit");
            }
            w.write_event(Event::End(BytesEnd::new("w:trPr")))
                .expect("write trPr end");
        }

        for cell in &row.cells {
            w.write_event(Event::Start(BytesStart::new("w:tc")))
                .expect("write tc start");

            // tcPr
            let has_tc_pr = cell.col_span > 1
                || cell.row_span > 1
                || cell.is_vmerge_continue
                || cell.background_color.is_some()
                || cell.border.is_some()
                || cell.vertical_align.is_some()
                || cell.text_align.is_some()
                || cell.text_direction.is_some()
                || cell.width_twips.is_some()
                || cell.padding.is_some();
            if has_tc_pr {
                w.write_event(Event::Start(BytesStart::new("w:tcPr")))
                    .expect("write tcPr start");

                // CT_TcPrBase is a strict sequence: tcW, gridSpan, vMerge,
                // tcBorders, shd, tcMar, textDirection, vAlign. This is the
                // third element-ordering defect of the same shape in this
                // release (w:pPr and w:tblPr were the others).
                if let Some(width) = cell.width_twips {
                    let mut tcw = BytesStart::new("w:tcW");
                    tcw.push_attribute(("w:w", width.to_string().as_str()));
                    tcw.push_attribute(("w:type", "dxa"));
                    w.write_event(Event::Empty(tcw)).expect("write tcW");
                }

                if cell.col_span > 1 {
                    let mut gs = BytesStart::new("w:gridSpan");
                    gs.push_attribute(("w:val", cell.col_span.to_string().as_str()));
                    w.write_event(Event::Empty(gs)).expect("write gridSpan");
                }

                if cell.is_vmerge_continue {
                    w.write_event(Event::Empty(BytesStart::new("w:vMerge")))
                        .expect("write vMerge cont");
                } else if cell.row_span > 1 {
                    let mut vm = BytesStart::new("w:vMerge");
                    vm.push_attribute(("w:val", "restart"));
                    w.write_event(Event::Empty(vm))
                        .expect("write vMerge restart");
                }

                if let Some(ref border) = cell.border {
                    write_table_borders(w, border, "w:tcBorders");
                }

                // CT_TcPrBase orders tcMar before textDirection and vAlign.
                if let Some(ref bg) = cell.background_color {
                    let mut shd = BytesStart::new("w:shd");
                    shd.push_attribute(("w:val", "clear"));
                    shd.push_attribute(("w:fill", rgb_to_hex(*bg).as_str()));
                    shd.push_attribute(("w:color", "auto"));
                    w.write_event(Event::Empty(shd)).expect("write cell shd");
                }

                if let Some(ref pad) = cell.padding {
                    w.write_event(Event::Start(BytesStart::new("w:tcMar")))
                        .expect("write tcMar start");
                    for (side, val) in [
                        ("w:top", pad.top_twips),
                        ("w:left", pad.left_twips),
                        ("w:bottom", pad.bottom_twips),
                        ("w:right", pad.right_twips),
                    ] {
                        if let Some(v) = val {
                            let mut elem = BytesStart::new(side);
                            elem.push_attribute(("w:w", v.to_string().as_str()));
                            elem.push_attribute(("w:type", "dxa"));
                            w.write_event(Event::Empty(elem)).expect("write tcMar side");
                        }
                    }
                    w.write_event(Event::End(BytesEnd::new("w:tcMar")))
                        .expect("write tcMar end");
                }

                if let Some(ref td) = cell.text_direction {
                    let val = match td {
                        crate::ir::TextDirection::LrTb => "lrTb",
                        crate::ir::TextDirection::TbRl => "tbRl",
                        crate::ir::TextDirection::BtLr => "btLr",
                    };
                    let mut td_elem = BytesStart::new("w:textDirection");
                    td_elem.push_attribute(("w:val", val));
                    w.write_event(Event::Empty(td_elem))
                        .expect("write textDirection");
                }
                if let Some(ref va) = cell.vertical_align {
                    let val = match va {
                        CellVerticalAlign::Top => "top",
                        CellVerticalAlign::Center => "center",
                        CellVerticalAlign::Bottom => "bottom",
                    };
                    let mut v_align = BytesStart::new("w:vAlign");
                    v_align.push_attribute(("w:val", val));
                    w.write_event(Event::Empty(v_align)).expect("write vAlign");
                }
                w.write_event(Event::End(BytesEnd::new("w:tcPr")))
                    .expect("write tcPr end");
            }

            for elem in &cell.content {
                write_docx_element(w, elem, image_rids, image_counter, links);
            }

            w.write_event(Event::End(BytesEnd::new("w:tc")))
                .expect("write tc end");
        }

        w.write_event(Event::End(BytesEnd::new("w:tr")))
            .expect("write tr end");
    }

    w.write_event(Event::End(BytesEnd::new("w:tbl")))
        .expect("write tbl end");
}

fn write_paragraph_borders(w: &mut Writer<Vec<u8>>, border: &crate::ir::ParagraphBorder) {
    w.write_event(Event::Start(BytesStart::new("w:pBdr")))
        .expect("write pBdr start");
    for (side_tag, side) in &[
        ("w:top", &border.top),
        ("w:left", &border.left),
        ("w:bottom", &border.bottom),
        ("w:right", &border.right),
        ("w:between", &border.between),
    ] {
        if let Some(bl) = side {
            write_border_line(w, side_tag, bl);
        }
    }
    w.write_event(Event::End(BytesEnd::new("w:pBdr")))
        .expect("write pBdr end");
}

fn write_table_borders(w: &mut Writer<Vec<u8>>, border: &crate::ir::TableBorder, tag: &str) {
    w.write_event(Event::Start(BytesStart::new(tag.to_owned())))
        .expect("write border start");
    for (side_tag, side) in [
        ("w:top", &border.top),
        ("w:left", &border.left),
        ("w:bottom", &border.bottom),
        ("w:right", &border.right),
        ("w:insideH", &border.inside_h),
        ("w:insideV", &border.inside_v),
    ] {
        if let Some(bl) = side {
            write_border_line(w, side_tag, bl);
        }
    }
    w.write_event(Event::End(BytesEnd::new(tag.to_owned())))
        .expect("write border end");
}

fn write_border_line(w: &mut Writer<Vec<u8>>, tag: &str, bl: &BorderLine) {
    let val = border_style_val(&bl.style);
    let mut elem = BytesStart::new(tag.to_owned());
    elem.push_attribute(("w:val", val));
    if let Some(sz) = bl.size {
        elem.push_attribute(("w:sz", sz.to_string().as_str()));
    } else {
        elem.push_attribute(("w:sz", "4"));
    }
    if let Some(sp) = bl.space {
        elem.push_attribute(("w:space", sp.to_string().as_str()));
    } else {
        elem.push_attribute(("w:space", "0"));
    }
    if let Some(ref color) = bl.color {
        elem.push_attribute(("w:color", rgb_to_hex(*color).as_str()));
    } else {
        elem.push_attribute(("w:color", "000000"));
    }
    w.write_event(Event::Empty(elem))
        .expect("write border line");
}

fn write_rich_list(
    w: &mut Writer<Vec<u8>>,
    rl: &DocxRichList,
    image_rids: &[ImageInfo],
    image_counter: &mut u32,
    links: &HyperlinkRids,
) {
    for item_elems in &rl.items {
        // Wrap item elements in a ListParagraph with numbering
        for (i, elem) in item_elems.iter().enumerate() {
            match elem {
                DocxElement::RichParagraph(rp) => {
                    let mut new_props = rp.props.clone();
                    if i == 0 {
                        new_props.style = Some("ListParagraph".to_string());
                        new_props.numbering = Some((rl.num_id, rl.level));
                    }
                    let new_p = DocxRichParagraph {
                        runs: rp.runs.clone(),
                        props: new_props,
                    };
                    write_rich_paragraph(w, &new_p, links);
                },
                other => write_docx_element(w, other, image_rids, image_counter, links),
            }
        }
    }
}

fn write_code_block(w: &mut Writer<Vec<u8>>, content: &str) {
    w.write_event(Event::Start(BytesStart::new("w:p")))
        .expect("write p start");

    w.write_event(Event::Start(BytesStart::new("w:pPr")))
        .expect("write pPr start");
    let mut ps = BytesStart::new("w:pStyle");
    ps.push_attribute(("w:val", "Code"));
    w.write_event(Event::Empty(ps)).expect("write code pStyle");
    w.write_event(Event::End(BytesEnd::new("w:pPr")))
        .expect("write pPr end");

    for (i, line) in content.lines().enumerate() {
        if i > 0 {
            // Line break between lines
            w.write_event(Event::Start(BytesStart::new("w:r")))
                .expect("write r start");
            w.write_event(Event::Empty(BytesStart::new("w:br")))
                .expect("write br");
            w.write_event(Event::End(BytesEnd::new("w:r")))
                .expect("write r end");
        }
        let run = Run::new(line);
        write_run(w, &run);
    }

    w.write_event(Event::End(BytesEnd::new("w:p")))
        .expect("write p end");
}

fn write_text_box(
    w: &mut Writer<Vec<u8>>,
    tb: &DocxTextBox,
    image_rids: &[ImageInfo],
    image_counter: &mut u32,
    links: &HyperlinkRids,
) {
    let float_anchor_val = |a: &crate::ir::FloatAnchor| match a {
        crate::ir::FloatAnchor::Page => "page",
        crate::ir::FloatAnchor::Margin => "margin",
        crate::ir::FloatAnchor::Column => "column",
        crate::ir::FloatAnchor::Paragraph => "paragraph",
    };

    w.write_event(Event::Start(BytesStart::new("w:p")))
        .expect("write p start");
    w.write_event(Event::Start(BytesStart::new("w:r")))
        .expect("write r start");
    w.write_event(Event::Start(BytesStart::new("w:drawing")))
        .expect("write drawing start");

    let mut anchor = BytesStart::new("wp:anchor");
    anchor.push_attribute(("distT", "0"));
    anchor.push_attribute(("distB", "0"));
    anchor.push_attribute(("distL", "114300"));
    anchor.push_attribute(("distR", "114300"));
    anchor.push_attribute(("simplePos", "0"));
    anchor.push_attribute(("relativeHeight", "251659264"));
    anchor.push_attribute(("behindDoc", "0"));
    anchor.push_attribute(("locked", "0"));
    anchor.push_attribute(("layoutInCell", "1"));
    anchor.push_attribute(("allowOverlap", "1"));
    w.write_event(Event::Start(anchor))
        .expect("write anchor start");

    let mut spos = BytesStart::new("wp:simplePos");
    spos.push_attribute(("x", "0"));
    spos.push_attribute(("y", "0"));
    w.write_event(Event::Empty(spos)).expect("write simplePos");

    let mut pos_h = BytesStart::new("wp:positionH");
    pos_h.push_attribute(("relativeFrom", float_anchor_val(&tb.h_anchor)));
    w.write_event(Event::Start(pos_h))
        .expect("write positionH start");
    let x_str = tb.x_emu.to_string();
    w.write_event(Event::Start(BytesStart::new("wp:posOffset")))
        .expect("write posOffset start");
    w.write_event(Event::Text(BytesText::new(&crate::core::xml::sanitize_xml_text(&x_str))))
        .expect("write posOffset text");
    w.write_event(Event::End(BytesEnd::new("wp:posOffset")))
        .expect("write posOffset end");
    w.write_event(Event::End(BytesEnd::new("wp:positionH")))
        .expect("write positionH end");

    let mut pos_v = BytesStart::new("wp:positionV");
    pos_v.push_attribute(("relativeFrom", float_anchor_val(&tb.v_anchor)));
    w.write_event(Event::Start(pos_v))
        .expect("write positionV start");
    let y_str = tb.y_emu.to_string();
    w.write_event(Event::Start(BytesStart::new("wp:posOffset")))
        .expect("write posOffset start");
    w.write_event(Event::Text(BytesText::new(&crate::core::xml::sanitize_xml_text(&y_str))))
        .expect("write posOffset text");
    w.write_event(Event::End(BytesEnd::new("wp:posOffset")))
        .expect("write posOffset end");
    w.write_event(Event::End(BytesEnd::new("wp:positionV")))
        .expect("write positionV end");

    let mut extent = BytesStart::new("wp:extent");
    extent.push_attribute(("cx", tb.width_emu.to_string().as_str()));
    extent.push_attribute(("cy", tb.height_emu.to_string().as_str()));
    w.write_event(Event::Empty(extent)).expect("write extent");

    *image_counter += 1;

    match &tb.wrap {
        crate::ir::TextWrap::Square => {
            let mut ws = BytesStart::new("wp:wrapSquare");
            ws.push_attribute(("wrapText", "bothSides"));
            w.write_event(Event::Empty(ws)).expect("write wrapSquare");
        },
        crate::ir::TextWrap::Tight => {
            w.write_event(Event::Empty(BytesStart::new("wp:wrapTight")))
                .expect("write wrapTight");
        },
        crate::ir::TextWrap::TopAndBottom => {
            w.write_event(Event::Empty(BytesStart::new("wp:wrapTopAndBottom")))
                .expect("write wrapTopAndBottom");
        },
        crate::ir::TextWrap::Behind | crate::ir::TextWrap::InFront => {
            w.write_event(Event::Empty(BytesStart::new("wp:wrapNone")))
                .expect("write wrapNone");
        },
        crate::ir::TextWrap::Through => {
            w.write_event(Event::Empty(BytesStart::new("wp:wrapThrough")))
                .expect("write wrapThrough");
        },
    }

    // CT_Anchor orders EG_WrapType before docPr; emitting docPr first
    // produces a schema-invalid drawing.
    let mut doc_pr = BytesStart::new("wp:docPr");
    doc_pr.push_attribute(("id", image_counter.to_string().as_str()));
    doc_pr.push_attribute(("name", format!("TextBox{}", *image_counter).as_str()));
    w.write_event(Event::Empty(doc_pr)).expect("write docPr");

    w.write_event(Event::Start(BytesStart::new("a:graphic")))
        .expect("write graphic start");
    let mut gdata = BytesStart::new("a:graphicData");
    gdata.push_attribute(("uri", WPS_NS));
    w.write_event(Event::Start(gdata))
        .expect("write graphicData start");

    w.write_event(Event::Start(BytesStart::new("wps:wsp")))
        .expect("write wsp start");

    let mut cnv_sp_pr = BytesStart::new("wps:cNvSpPr");
    cnv_sp_pr.push_attribute(("txBx", "1"));
    w.write_event(Event::Empty(cnv_sp_pr))
        .expect("write cNvSpPr");

    w.write_event(Event::Start(BytesStart::new("wps:spPr")))
        .expect("write spPr start");
    w.write_event(Event::Start(BytesStart::new("a:xfrm")))
        .expect("write xfrm start");
    let mut off = BytesStart::new("a:off");
    off.push_attribute(("x", "0"));
    off.push_attribute(("y", "0"));
    w.write_event(Event::Empty(off)).expect("write off");
    let mut ext = BytesStart::new("a:ext");
    ext.push_attribute(("cx", tb.width_emu.to_string().as_str()));
    ext.push_attribute(("cy", tb.height_emu.to_string().as_str()));
    w.write_event(Event::Empty(ext)).expect("write ext");
    w.write_event(Event::End(BytesEnd::new("a:xfrm")))
        .expect("write xfrm end");
    let mut geom = BytesStart::new("a:prstGeom");
    geom.push_attribute(("prst", "rect"));
    w.write_event(Event::Start(geom))
        .expect("write prstGeom start");
    w.write_event(Event::Empty(BytesStart::new("a:avLst")))
        .expect("write avLst");
    w.write_event(Event::End(BytesEnd::new("a:prstGeom")))
        .expect("write prstGeom end");
    w.write_event(Event::End(BytesEnd::new("wps:spPr")))
        .expect("write spPr end");

    w.write_event(Event::Start(BytesStart::new("wps:txbx")))
        .expect("write txbx start");
    w.write_event(Event::Start(BytesStart::new("w:txbxContent")))
        .expect("write txbxContent start");
    let mut txb_ic = 0u32;
    for elem in &tb.content {
        write_docx_element(w, elem, image_rids, &mut txb_ic, links);
    }
    w.write_event(Event::End(BytesEnd::new("w:txbxContent")))
        .expect("write txbxContent end");
    w.write_event(Event::End(BytesEnd::new("wps:txbx")))
        .expect("write txbx end");

    w.write_event(Event::Empty(BytesStart::new("wps:bodyPr")))
        .expect("write bodyPr");
    w.write_event(Event::End(BytesEnd::new("wps:wsp")))
        .expect("write wsp end");
    w.write_event(Event::End(BytesEnd::new("a:graphicData")))
        .expect("write graphicData end");
    w.write_event(Event::End(BytesEnd::new("a:graphic")))
        .expect("write graphic end");
    w.write_event(Event::End(BytesEnd::new("wp:anchor")))
        .expect("write anchor end");
    w.write_event(Event::End(BytesEnd::new("w:drawing")))
        .expect("write drawing end");
    w.write_event(Event::End(BytesEnd::new("w:r")))
        .expect("write r end");
    w.write_event(Event::End(BytesEnd::new("w:p")))
        .expect("write p end");
}

fn write_inline_image_run(
    w: &mut Writer<Vec<u8>>,
    rid: &str,
    width_emu: u64,
    height_emu: u64,
    alt_text: Option<&str>,
    _decorative: bool,
    pic_id: u32,
) {
    w.write_event(Event::Start(BytesStart::new("w:p")))
        .expect("write p start");
    w.write_event(Event::Start(BytesStart::new("w:r")))
        .expect("write r start");
    w.write_event(Event::Start(BytesStart::new("w:drawing")))
        .expect("write drawing start");

    // wp:inline
    w.write_event(Event::Start(BytesStart::new("wp:inline")))
        .expect("write inline start");

    // wp:extent
    let mut extent = BytesStart::new("wp:extent");
    extent.push_attribute(("cx", width_emu.to_string().as_str()));
    extent.push_attribute(("cy", height_emu.to_string().as_str()));
    w.write_event(Event::Empty(extent)).expect("write extent");

    // wp:docPr
    let mut doc_pr = BytesStart::new("wp:docPr");
    doc_pr.push_attribute(("id", pic_id.to_string().as_str()));
    doc_pr.push_attribute(("name", format!("Image{pic_id}").as_str()));
    if let Some(alt) = alt_text {
        doc_pr.push_attribute(("descr", alt));
    }
    w.write_event(Event::Empty(doc_pr)).expect("write docPr");

    // a:graphic
    w.write_event(Event::Start(BytesStart::new("a:graphic")))
        .expect("write graphic start");

    let mut gdata = BytesStart::new("a:graphicData");
    gdata.push_attribute(("uri", PIC_NS));
    w.write_event(Event::Start(gdata))
        .expect("write graphicData start");

    // pic:pic
    w.write_event(Event::Start(BytesStart::new("pic:pic")))
        .expect("write pic start");

    // pic:nvPicPr
    w.write_event(Event::Start(BytesStart::new("pic:nvPicPr")))
        .expect("write nvPicPr start");
    let mut cnv_pr = BytesStart::new("pic:cNvPr");
    cnv_pr.push_attribute(("id", pic_id.to_string().as_str()));
    cnv_pr.push_attribute(("name", format!("Image{pic_id}").as_str()));
    w.write_event(Event::Empty(cnv_pr)).expect("write cNvPr");
    w.write_event(Event::Empty(BytesStart::new("pic:cNvPicPr")))
        .expect("write cNvPicPr");
    w.write_event(Event::End(BytesEnd::new("pic:nvPicPr")))
        .expect("write nvPicPr end");

    // pic:blipFill
    w.write_event(Event::Start(BytesStart::new("pic:blipFill")))
        .expect("write blipFill start");
    let mut blip = BytesStart::new("a:blip");
    blip.push_attribute(("r:embed", rid));
    w.write_event(Event::Empty(blip)).expect("write blip");
    w.write_event(Event::Start(BytesStart::new("a:stretch")))
        .expect("write stretch start");
    w.write_event(Event::Empty(BytesStart::new("a:fillRect")))
        .expect("write fillRect");
    w.write_event(Event::End(BytesEnd::new("a:stretch")))
        .expect("write stretch end");
    w.write_event(Event::End(BytesEnd::new("pic:blipFill")))
        .expect("write blipFill end");

    // pic:spPr
    w.write_event(Event::Start(BytesStart::new("pic:spPr")))
        .expect("write spPr start");
    w.write_event(Event::Start(BytesStart::new("a:xfrm")))
        .expect("write xfrm start");
    let mut off = BytesStart::new("a:off");
    off.push_attribute(("x", "0"));
    off.push_attribute(("y", "0"));
    w.write_event(Event::Empty(off)).expect("write off");
    let mut ext = BytesStart::new("a:ext");
    ext.push_attribute(("cx", width_emu.to_string().as_str()));
    ext.push_attribute(("cy", height_emu.to_string().as_str()));
    w.write_event(Event::Empty(ext)).expect("write ext");
    w.write_event(Event::End(BytesEnd::new("a:xfrm")))
        .expect("write xfrm end");
    let mut geom = BytesStart::new("a:prstGeom");
    geom.push_attribute(("prst", "rect"));
    w.write_event(Event::Start(geom))
        .expect("write prstGeom start");
    w.write_event(Event::Empty(BytesStart::new("a:avLst")))
        .expect("write avLst");
    w.write_event(Event::End(BytesEnd::new("a:prstGeom")))
        .expect("write prstGeom end");
    w.write_event(Event::End(BytesEnd::new("pic:spPr")))
        .expect("write spPr end");

    w.write_event(Event::End(BytesEnd::new("pic:pic")))
        .expect("write pic end");
    w.write_event(Event::End(BytesEnd::new("a:graphicData")))
        .expect("write graphicData end");
    w.write_event(Event::End(BytesEnd::new("a:graphic")))
        .expect("write graphic end");
    w.write_event(Event::End(BytesEnd::new("wp:inline")))
        .expect("write inline end");
    w.write_event(Event::End(BytesEnd::new("w:drawing")))
        .expect("write drawing end");
    w.write_event(Event::End(BytesEnd::new("w:r")))
        .expect("write r end");
    w.write_event(Event::End(BytesEnd::new("w:p")))
        .expect("write p end");
}

fn write_floating_image_run(
    w: &mut Writer<Vec<u8>>,
    rid: &str,
    fi: &crate::ir::FloatingImage,
    alt_text: Option<&str>,
    _decorative: bool,
    pic_id: u32,
) {
    let float_anchor_val = |a: &crate::ir::FloatAnchor| match a {
        crate::ir::FloatAnchor::Page => "page",
        crate::ir::FloatAnchor::Margin => "margin",
        crate::ir::FloatAnchor::Column => "column",
        crate::ir::FloatAnchor::Paragraph => "paragraph",
    };

    w.write_event(Event::Start(BytesStart::new("w:p")))
        .expect("write p start");
    w.write_event(Event::Start(BytesStart::new("w:r")))
        .expect("write r start");
    w.write_event(Event::Start(BytesStart::new("w:drawing")))
        .expect("write drawing start");

    let mut anchor = BytesStart::new("wp:anchor");
    anchor.push_attribute((
        "behindDoc",
        if matches!(fi.text_wrap, crate::ir::TextWrap::Behind) {
            "1"
        } else {
            "0"
        },
    ));
    anchor.push_attribute(("distT", "0"));
    anchor.push_attribute(("distB", "0"));
    anchor.push_attribute(("distL", "114300"));
    anchor.push_attribute(("distR", "114300"));
    anchor.push_attribute(("simplePos", "0"));
    anchor.push_attribute(("relativeHeight", "251659264"));
    // locked and layoutInCell are required by CT_Anchor. The text box writer
    // sets them; this one never did. behindDoc is already set above — adding
    // it a second time makes the element not well-formed, which no schema
    // check can even reach because the parse fails first.
    anchor.push_attribute(("locked", "0"));
    anchor.push_attribute(("layoutInCell", "1"));
    anchor.push_attribute(("allowOverlap", if fi.allow_overlap { "1" } else { "0" }));
    w.write_event(Event::Start(anchor))
        .expect("write anchor start");

    // CT_Point2D requires both x and y; a bare element is invalid.
    let mut spos = BytesStart::new("wp:simplePos");
    spos.push_attribute(("x", "0"));
    spos.push_attribute(("y", "0"));
    w.write_event(Event::Empty(spos)).expect("write simplePos");

    let mut pos_h = BytesStart::new("wp:positionH");
    pos_h.push_attribute(("relativeFrom", float_anchor_val(&fi.h_anchor)));
    w.write_event(Event::Start(pos_h))
        .expect("write positionH start");
    let x_str = fi.x_emu.to_string();
    w.write_event(Event::Start(BytesStart::new("wp:posOffset")))
        .expect("write posOffset start");
    w.write_event(Event::Text(BytesText::new(&crate::core::xml::sanitize_xml_text(&x_str))))
        .expect("write posOffset text");
    w.write_event(Event::End(BytesEnd::new("wp:posOffset")))
        .expect("write posOffset end");
    w.write_event(Event::End(BytesEnd::new("wp:positionH")))
        .expect("write positionH end");

    let mut pos_v = BytesStart::new("wp:positionV");
    pos_v.push_attribute(("relativeFrom", float_anchor_val(&fi.v_anchor)));
    w.write_event(Event::Start(pos_v))
        .expect("write positionV start");
    let y_str = fi.y_emu.to_string();
    w.write_event(Event::Start(BytesStart::new("wp:posOffset")))
        .expect("write posOffset start");
    w.write_event(Event::Text(BytesText::new(&crate::core::xml::sanitize_xml_text(&y_str))))
        .expect("write posOffset text");
    w.write_event(Event::End(BytesEnd::new("wp:posOffset")))
        .expect("write posOffset end");
    w.write_event(Event::End(BytesEnd::new("wp:positionV")))
        .expect("write positionV end");

    let mut extent = BytesStart::new("wp:extent");
    extent.push_attribute(("cx", fi.width_emu.to_string().as_str()));
    extent.push_attribute(("cy", fi.height_emu.to_string().as_str()));
    w.write_event(Event::Empty(extent)).expect("write extent");

    match &fi.text_wrap {
        crate::ir::TextWrap::Square => {
            let mut ws = BytesStart::new("wp:wrapSquare");
            ws.push_attribute(("wrapText", "bothSides"));
            w.write_event(Event::Empty(ws)).expect("write wrapSquare");
        },
        crate::ir::TextWrap::Tight => {
            w.write_event(Event::Empty(BytesStart::new("wp:wrapTight")))
                .expect("write wrapTight");
        },
        crate::ir::TextWrap::TopAndBottom => {
            w.write_event(Event::Empty(BytesStart::new("wp:wrapTopAndBottom")))
                .expect("write wrapTopAndBottom");
        },
        crate::ir::TextWrap::Behind => {
            w.write_event(Event::Empty(BytesStart::new("wp:wrapNone")))
                .expect("write wrapNone behind");
        },
        crate::ir::TextWrap::InFront => {
            w.write_event(Event::Empty(BytesStart::new("wp:wrapNone")))
                .expect("write wrapNone infront");
        },
        crate::ir::TextWrap::Through => {
            w.write_event(Event::Empty(BytesStart::new("wp:wrapThrough")))
                .expect("write wrapThrough");
        },
    }

    // CT_Anchor orders EG_WrapType before docPr; emitting docPr first
    // produces a schema-invalid drawing.
    let mut doc_pr = BytesStart::new("wp:docPr");
    doc_pr.push_attribute(("id", pic_id.to_string().as_str()));
    doc_pr.push_attribute(("name", format!("Image{pic_id}").as_str()));
    if let Some(alt) = alt_text {
        doc_pr.push_attribute(("descr", alt));
    }
    w.write_event(Event::Empty(doc_pr)).expect("write docPr");

    // a:graphic (same pic:pic structure as inline)
    w.write_event(Event::Start(BytesStart::new("a:graphic")))
        .expect("write graphic start");
    let mut gdata = BytesStart::new("a:graphicData");
    gdata.push_attribute(("uri", PIC_NS));
    w.write_event(Event::Start(gdata))
        .expect("write graphicData start");

    w.write_event(Event::Start(BytesStart::new("pic:pic")))
        .expect("write pic start");
    w.write_event(Event::Start(BytesStart::new("pic:nvPicPr")))
        .expect("write nvPicPr start");
    let mut cnv_pr = BytesStart::new("pic:cNvPr");
    cnv_pr.push_attribute(("id", pic_id.to_string().as_str()));
    cnv_pr.push_attribute(("name", format!("Image{pic_id}").as_str()));
    w.write_event(Event::Empty(cnv_pr)).expect("write cNvPr");
    w.write_event(Event::Empty(BytesStart::new("pic:cNvPicPr")))
        .expect("write cNvPicPr");
    w.write_event(Event::End(BytesEnd::new("pic:nvPicPr")))
        .expect("write nvPicPr end");

    w.write_event(Event::Start(BytesStart::new("pic:blipFill")))
        .expect("write blipFill start");
    let mut blip = BytesStart::new("a:blip");
    blip.push_attribute(("r:embed", rid));
    w.write_event(Event::Empty(blip)).expect("write blip");
    w.write_event(Event::Start(BytesStart::new("a:stretch")))
        .expect("write stretch start");
    w.write_event(Event::Empty(BytesStart::new("a:fillRect")))
        .expect("write fillRect");
    w.write_event(Event::End(BytesEnd::new("a:stretch")))
        .expect("write stretch end");
    w.write_event(Event::End(BytesEnd::new("pic:blipFill")))
        .expect("write blipFill end");

    w.write_event(Event::Start(BytesStart::new("pic:spPr")))
        .expect("write spPr start");
    w.write_event(Event::Start(BytesStart::new("a:xfrm")))
        .expect("write xfrm start");
    let mut off = BytesStart::new("a:off");
    off.push_attribute(("x", "0"));
    off.push_attribute(("y", "0"));
    w.write_event(Event::Empty(off)).expect("write off");
    let mut ext = BytesStart::new("a:ext");
    ext.push_attribute(("cx", fi.width_emu.to_string().as_str()));
    ext.push_attribute(("cy", fi.height_emu.to_string().as_str()));
    w.write_event(Event::Empty(ext)).expect("write ext");
    w.write_event(Event::End(BytesEnd::new("a:xfrm")))
        .expect("write xfrm end");
    let mut geom = BytesStart::new("a:prstGeom");
    geom.push_attribute(("prst", "rect"));
    w.write_event(Event::Start(geom))
        .expect("write prstGeom start");
    w.write_event(Event::Empty(BytesStart::new("a:avLst")))
        .expect("write avLst");
    w.write_event(Event::End(BytesEnd::new("a:prstGeom")))
        .expect("write prstGeom end");
    w.write_event(Event::End(BytesEnd::new("pic:spPr")))
        .expect("write spPr end");

    w.write_event(Event::End(BytesEnd::new("pic:pic")))
        .expect("write pic end");
    w.write_event(Event::End(BytesEnd::new("a:graphicData")))
        .expect("write graphicData end");
    w.write_event(Event::End(BytesEnd::new("a:graphic")))
        .expect("write graphic end");
    w.write_event(Event::End(BytesEnd::new("wp:anchor")))
        .expect("write anchor end");
    w.write_event(Event::End(BytesEnd::new("w:drawing")))
        .expect("write drawing end");
    w.write_event(Event::End(BytesEnd::new("w:r")))
        .expect("write r end");
    w.write_event(Event::End(BytesEnd::new("w:p")))
        .expect("write p end");
}

/// Emit a synthetic empty paragraph that carries an inline `<w:sectPr>`
/// inside its `<w:pPr>`. Used for non-final section boundaries — the
/// paragraph is what marks the section break for the reader; its
/// `<w:sectPr>` describes the section ending at this point. We don't
/// emit hf / footnote references on inline sectPr (they're document-wide
/// and live on the body-level final sectPr only).
fn write_inline_section_break_paragraph(w: &mut Writer<Vec<u8>>, sp: &DocxSectPr) {
    w.write_event(Event::Start(BytesStart::new("w:p")))
        .expect("write inline-section p start");
    w.write_event(Event::Start(BytesStart::new("w:pPr")))
        .expect("write inline-section pPr start");
    write_section_pr_body(w, sp.page_setup.as_ref(), sp.columns.as_ref(), &sp.break_type);
    w.write_event(Event::End(BytesEnd::new("w:pPr")))
        .expect("write inline-section pPr end");
    w.write_event(Event::End(BytesEnd::new("w:p")))
        .expect("write inline-section p end");
}

/// Shared `<w:sectPr>...</w:sectPr>` body writer — used by both the
/// body-level final sectPr and inline (per-paragraph) section breaks.
/// Caller writes the surrounding `<w:sectPr>`/`</w:sectPr>` tags.
fn write_section_pr_body(
    w: &mut Writer<Vec<u8>>,
    page_setup: Option<&PageSetup>,
    columns: Option<&ColumnLayout>,
    break_type: &SectionBreakType,
) {
    w.write_event(Event::Start(BytesStart::new("w:sectPr")))
        .expect("write sectPr start");

    match break_type {
        SectionBreakType::Continuous => {
            // Continuous is the default; emit it explicitly so the reader
            // doesn't pick up a stale value from a sibling section.
            let mut t = BytesStart::new("w:type");
            t.push_attribute(("w:val", "continuous"));
            w.write_event(Event::Empty(t)).expect("write sect type");
        },
        SectionBreakType::NextPage => {
            let mut t = BytesStart::new("w:type");
            t.push_attribute(("w:val", "nextPage"));
            w.write_event(Event::Empty(t)).expect("write sect type");
        },
        SectionBreakType::EvenPage => {
            let mut t = BytesStart::new("w:type");
            t.push_attribute(("w:val", "evenPage"));
            w.write_event(Event::Empty(t)).expect("write sect type");
        },
        SectionBreakType::OddPage => {
            let mut t = BytesStart::new("w:type");
            t.push_attribute(("w:val", "oddPage"));
            w.write_event(Event::Empty(t)).expect("write sect type");
        },
    }

    if let Some(ps) = page_setup {
        let mut pg_sz = BytesStart::new("w:pgSz");
        pg_sz.push_attribute(("w:w", ps.width_twips.to_string().as_str()));
        pg_sz.push_attribute(("w:h", ps.height_twips.to_string().as_str()));
        if ps.landscape {
            pg_sz.push_attribute(("w:orient", "landscape"));
        }
        w.write_event(Event::Empty(pg_sz)).expect("write pgSz");

        let mut pg_mar = BytesStart::new("w:pgMar");
        pg_mar.push_attribute(("w:top", ps.margin_top_twips.to_string().as_str()));
        pg_mar.push_attribute(("w:bottom", ps.margin_bottom_twips.to_string().as_str()));
        pg_mar.push_attribute(("w:left", ps.margin_left_twips.to_string().as_str()));
        pg_mar.push_attribute(("w:right", ps.margin_right_twips.to_string().as_str()));
        pg_mar.push_attribute(("w:header", ps.header_distance_twips.to_string().as_str()));
        pg_mar.push_attribute(("w:footer", ps.footer_distance_twips.to_string().as_str()));
        // CT_PageMar declares all seven attributes as use="required".
        pg_mar.push_attribute(("w:gutter", "0"));
        w.write_event(Event::Empty(pg_mar)).expect("write pgMar");
    }

    if let Some(cols) = columns {
        if cols.column_widths_twips.is_empty() {
            let mut c = BytesStart::new("w:cols");
            c.push_attribute(("w:num", cols.count.to_string().as_str()));
            if let Some(sp) = cols.space_twips {
                c.push_attribute(("w:space", sp.to_string().as_str()));
            }
            if cols.separator {
                c.push_attribute(("w:sep", "1"));
            }
            w.write_event(Event::Empty(c)).expect("write cols");
        } else {
            let mut c = BytesStart::new("w:cols");
            c.push_attribute(("w:num", cols.count.to_string().as_str()));
            if cols.separator {
                c.push_attribute(("w:sep", "1"));
            }
            w.write_event(Event::Start(c)).expect("write cols start");
            let default_space = cols.space_twips.unwrap_or(720);
            for &cw in &cols.column_widths_twips {
                let mut col = BytesStart::new("w:col");
                col.push_attribute(("w:w", cw.to_string().as_str()));
                col.push_attribute(("w:space", default_space.to_string().as_str()));
                w.write_event(Event::Empty(col)).expect("write col");
            }
            w.write_event(Event::End(BytesEnd::new("w:cols")))
                .expect("write cols end");
        }
    }

    w.write_event(Event::End(BytesEnd::new("w:sectPr")))
        .expect("write sectPr end");
}

fn write_body_sect_pr(w: &mut Writer<Vec<u8>>, sp: &SectPrInfo) {
    w.write_event(Event::Start(BytesStart::new("w:sectPr")))
        .expect("write sectPr start");

    // ECMA-376 §17.10.5 allows at most one reference of each type per
    // section. headers_footers is a single flat list shared by every
    // section, so without this a multi-section document emitted duplicate
    // w:type values in the final sectPr.
    let mut seen_types: Vec<HfType> = Vec::new();
    let mut has_first_page = false;
    for (hf_type, rid) in &sp.hf_rids {
        if seen_types.contains(hf_type) {
            continue;
        }
        seen_types.push(*hf_type);
        if matches!(hf_type, HfType::FirstPageHeader | HfType::FirstPageFooter) {
            has_first_page = true;
        }
        let (tag, type_val) = match hf_type {
            HfType::DefaultHeader => ("w:headerReference", "default"),
            HfType::FirstPageHeader => ("w:headerReference", "first"),
            HfType::EvenPageHeader => ("w:headerReference", "even"),
            HfType::DefaultFooter => ("w:footerReference", "default"),
            HfType::FirstPageFooter => ("w:footerReference", "first"),
            HfType::EvenPageFooter => ("w:footerReference", "even"),
        };
        let mut elem = BytesStart::new(tag);
        elem.push_attribute(("w:type", type_val));
        elem.push_attribute(("r:id", rid.as_str()));
        w.write_event(Event::Empty(elem)).expect("write hfRef");
    }

    if let Some(ref rid) = sp.footnote_rid {
        let elem = BytesStart::new("w:footnotePr");
        let _ = rid;
        w.write_event(Event::Empty(elem)).expect("write footnotePr");
    }

    match sp.break_type {
        SectionBreakType::Continuous => {},
        SectionBreakType::NextPage => {
            let mut t = BytesStart::new("w:type");
            t.push_attribute(("w:val", "nextPage"));
            w.write_event(Event::Empty(t)).expect("write sect type");
        },
        SectionBreakType::EvenPage => {
            let mut t = BytesStart::new("w:type");
            t.push_attribute(("w:val", "evenPage"));
            w.write_event(Event::Empty(t)).expect("write sect type");
        },
        SectionBreakType::OddPage => {
            let mut t = BytesStart::new("w:type");
            t.push_attribute(("w:val", "oddPage"));
            w.write_event(Event::Empty(t)).expect("write sect type");
        },
    }

    if let Some(ref ps) = sp.page_setup {
        let mut pg_sz = BytesStart::new("w:pgSz");
        pg_sz.push_attribute(("w:w", ps.width_twips.to_string().as_str()));
        pg_sz.push_attribute(("w:h", ps.height_twips.to_string().as_str()));
        if ps.landscape {
            pg_sz.push_attribute(("w:orient", "landscape"));
        }
        w.write_event(Event::Empty(pg_sz)).expect("write pgSz");

        let mut pg_mar = BytesStart::new("w:pgMar");
        pg_mar.push_attribute(("w:top", ps.margin_top_twips.to_string().as_str()));
        pg_mar.push_attribute(("w:bottom", ps.margin_bottom_twips.to_string().as_str()));
        pg_mar.push_attribute(("w:left", ps.margin_left_twips.to_string().as_str()));
        pg_mar.push_attribute(("w:right", ps.margin_right_twips.to_string().as_str()));
        pg_mar.push_attribute(("w:header", ps.header_distance_twips.to_string().as_str()));
        pg_mar.push_attribute(("w:footer", ps.footer_distance_twips.to_string().as_str()));
        // CT_PageMar declares all seven attributes as use="required".
        pg_mar.push_attribute(("w:gutter", "0"));
        w.write_event(Event::Empty(pg_mar)).expect("write pgMar");
    }

    if let Some(ref cols) = sp.columns {
        if cols.column_widths_twips.is_empty() {
            let mut c = BytesStart::new("w:cols");
            c.push_attribute(("w:num", cols.count.to_string().as_str()));
            if let Some(sp) = cols.space_twips {
                c.push_attribute(("w:space", sp.to_string().as_str()));
            }
            if cols.separator {
                c.push_attribute(("w:sep", "1"));
            }
            w.write_event(Event::Empty(c)).expect("write cols");
        } else {
            let mut c = BytesStart::new("w:cols");
            c.push_attribute(("w:num", cols.count.to_string().as_str()));
            if cols.separator {
                c.push_attribute(("w:sep", "1"));
            }
            w.write_event(Event::Start(c)).expect("write cols start");
            let default_space = cols.space_twips.unwrap_or(720);
            for &cw in &cols.column_widths_twips {
                let mut col = BytesStart::new("w:col");
                col.push_attribute(("w:w", cw.to_string().as_str()));
                col.push_attribute(("w:space", default_space.to_string().as_str()));
                w.write_event(Event::Empty(col)).expect("write col");
            }
            w.write_event(Event::End(BytesEnd::new("w:cols")))
                .expect("write cols end");
        }
    }

    // CT_SectPr puts titlePg after cols, not next to the header
    // references. A w:type="first" reference does nothing without it — the
    // distinct first-page header is simply never shown.
    if has_first_page {
        w.write_event(Event::Empty(BytesStart::new("w:titlePg")))
            .expect("write titlePg");
    }

    w.write_event(Event::End(BytesEnd::new("w:sectPr")))
        .expect("write sectPr end");
}

// ---------------------------------------------------------------------------
// Header/footer XML generation
// ---------------------------------------------------------------------------

fn generate_hf_xml(
    elements: &[DocxElement],
    image_rids: &[ImageInfo],
    is_header: bool,
    links: &HyperlinkRids,
) -> Vec<u8> {
    let mut w = Writer::new_with_indent(Vec::new(), b' ', 2);
    w.write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
        .expect("write decl");

    let tag = if is_header { "w:hdr" } else { "w:ftr" };
    let mut root = BytesStart::new(tag);
    root.push_attribute(("xmlns:w", WML_NS));
    root.push_attribute(("xmlns:r", R_NS));
    // A header or footer can hold a drawing (a text box or an image), which
    // emits the wp:/a:/pic:/wps: prefixes. The document root declares these
    // and this one never did, so such a part was not well-formed XML at all.
    // Declared unconditionally: the cost is four attributes, and gating them
    // on a content scan is what left the gap in the first place.
    root.push_attribute(("xmlns:wp", DRAWING_NS));
    root.push_attribute(("xmlns:a", DML_NS));
    root.push_attribute(("xmlns:pic", PIC_NS));
    root.push_attribute(("xmlns:wps", WPS_NS));
    w.write_event(Event::Start(root)).expect("write hf root");

    let mut ic = 0u32;
    for elem in elements {
        write_docx_element(&mut w, elem, image_rids, &mut ic, links);
    }
    if elements.is_empty() {
        w.write_event(Event::Start(BytesStart::new("w:p")))
            .expect("write p");
        w.write_event(Event::End(BytesEnd::new("w:p")))
            .expect("write p end");
    }

    w.write_event(Event::End(BytesEnd::new(tag)))
        .expect("write hf end");
    w.into_inner()
}

// ---------------------------------------------------------------------------
// Footnotes/endnotes XML generation
// ---------------------------------------------------------------------------

fn generate_footnotes_xml(
    notes: &[DocxNote],
    image_rids: &[ImageInfo],
    links: &HyperlinkRids,
) -> Vec<u8> {
    generate_notes_xml(notes, image_rids, false, links)
}

fn generate_endnotes_xml(
    notes: &[DocxNote],
    image_rids: &[ImageInfo],
    links: &HyperlinkRids,
) -> Vec<u8> {
    generate_notes_xml(notes, image_rids, true, links)
}

/// `word/settings.xml`, written only when a part depends on it.
fn generate_settings_xml(even_and_odd_headers: bool, embed_fonts: bool) -> Vec<u8> {
    let mut w = Writer::new_with_indent(Vec::new(), b' ', 2);
    w.write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
        .expect("write decl");
    let mut root = BytesStart::new("w:settings");
    root.push_attribute(("xmlns:w", WML_NS));
    w.write_event(Event::Start(root)).expect("write settings");
    // CT_Settings is a sequence; embedTrueTypeFonts precedes evenAndOddHeaders.
    if embed_fonts {
        w.write_event(Event::Empty(BytesStart::new("w:embedTrueTypeFonts")))
            .expect("write embedTrueTypeFonts");
    }
    if even_and_odd_headers {
        w.write_event(Event::Empty(BytesStart::new("w:evenAndOddHeaders")))
            .expect("write evenAndOddHeaders");
    }
    w.write_event(Event::End(BytesEnd::new("w:settings")))
        .expect("write settings end");
    w.into_inner()
}

fn generate_notes_xml(
    notes: &[DocxNote],
    image_rids: &[ImageInfo],
    is_endnote: bool,
    links: &HyperlinkRids,
) -> Vec<u8> {
    let mut w = Writer::new_with_indent(Vec::new(), b' ', 2);
    w.write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
        .expect("write decl");

    let root_tag = if is_endnote {
        "w:endnotes"
    } else {
        "w:footnotes"
    };
    let note_tag = if is_endnote {
        "w:endnote"
    } else {
        "w:footnote"
    };

    let mut root = BytesStart::new(root_tag);
    root.push_attribute(("xmlns:w", WML_NS));
    // A note can carry a hyperlink, which emits r:id. Without xmlns:r the
    // part is not even well-formed XML.
    root.push_attribute(("xmlns:r", R_NS));
    w.write_event(Event::Start(root)).expect("write notes root");

    // Word expects the separator (id -1) and continuationSeparator (id 0)
    // notes in every notes part; sectPr's footnotePr refers to them
    // implicitly. Without them there is no rule above the notes.
    for (id, kind) in [("-1", "separator"), ("0", "continuationSeparator")] {
        let mut sep = BytesStart::new(note_tag);
        sep.push_attribute(("w:type", kind));
        sep.push_attribute(("w:id", id));
        w.write_event(Event::Start(sep))
            .expect("write separator note");
        w.write_event(Event::Start(BytesStart::new("w:p")))
            .expect("write p");
        w.write_event(Event::Start(BytesStart::new("w:r")))
            .expect("write r");
        let mark = if is_endnote {
            "w:endnoteRef"
        } else {
            "w:footnoteRef"
        };
        let _ = mark;
        w.write_event(Event::Empty(BytesStart::new(if kind == "separator" {
            "w:separator"
        } else {
            "w:continuationSeparator"
        })))
        .expect("write separator mark");
        w.write_event(Event::End(BytesEnd::new("w:r")))
            .expect("write r end");
        w.write_event(Event::End(BytesEnd::new("w:p")))
            .expect("write p end");
        w.write_event(Event::End(BytesEnd::new(note_tag)))
            .expect("write separator note end");
    }

    for note in notes {
        let mut note_elem = BytesStart::new(note_tag);
        note_elem.push_attribute(("w:id", note.id.to_string().as_str()));
        w.write_event(Event::Start(note_elem))
            .expect("write note start");

        let mut ic = 0u32;
        for elem in &note.elements {
            write_docx_element(&mut w, elem, image_rids, &mut ic, links);
        }
        if note.elements.is_empty() {
            w.write_event(Event::Start(BytesStart::new("w:p")))
                .expect("write p");
            w.write_event(Event::End(BytesEnd::new("w:p")))
                .expect("write p end");
        }

        w.write_event(Event::End(BytesEnd::new(note_tag)))
            .expect("write note end");
    }

    w.write_event(Event::End(BytesEnd::new(root_tag)))
        .expect("write notes end");
    w.into_inner()
}

// ---------------------------------------------------------------------------
// Core properties XML
// ---------------------------------------------------------------------------

fn generate_core_props_xml(props: &CoreProps) -> Vec<u8> {
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
    root.push_attribute(("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"));
    w.write_event(Event::Start(root)).expect("write core root");

    if let Some(ref v) = props.title {
        w.write_event(Event::Start(BytesStart::new("dc:title")))
            .expect("write title start");
        w.write_event(Event::Text(BytesText::new(&crate::core::xml::sanitize_xml_text(v))))
            .expect("write title text");
        w.write_event(Event::End(BytesEnd::new("dc:title")))
            .expect("write title end");
    }
    if let Some(ref v) = props.subject {
        w.write_event(Event::Start(BytesStart::new("dc:subject")))
            .expect("write subject start");
        w.write_event(Event::Text(BytesText::new(&crate::core::xml::sanitize_xml_text(v))))
            .expect("write subject text");
        w.write_event(Event::End(BytesEnd::new("dc:subject")))
            .expect("write subject end");
    }
    if let Some(ref v) = props.author {
        w.write_event(Event::Start(BytesStart::new("dc:creator")))
            .expect("write creator start");
        w.write_event(Event::Text(BytesText::new(&crate::core::xml::sanitize_xml_text(v))))
            .expect("write creator text");
        w.write_event(Event::End(BytesEnd::new("dc:creator")))
            .expect("write creator end");
    }
    if let Some(ref v) = props.description {
        w.write_event(Event::Start(BytesStart::new("dc:description")))
            .expect("write desc start");
        w.write_event(Event::Text(BytesText::new(&crate::core::xml::sanitize_xml_text(v))))
            .expect("write desc text");
        w.write_event(Event::End(BytesEnd::new("dc:description")))
            .expect("write desc end");
    }
    if let Some(ref v) = props.keywords {
        w.write_event(Event::Start(BytesStart::new("cp:keywords")))
            .expect("write kw start");
        w.write_event(Event::Text(BytesText::new(&crate::core::xml::sanitize_xml_text(v))))
            .expect("write kw text");
        w.write_event(Event::End(BytesEnd::new("cp:keywords")))
            .expect("write kw end");
    }
    if let Some(ref v) = props.created {
        let mut elem = BytesStart::new("dcterms:created");
        elem.push_attribute(("xsi:type", "dcterms:W3CDTF"));
        w.write_event(Event::Start(elem))
            .expect("write created start");
        w.write_event(Event::Text(BytesText::new(&crate::core::xml::sanitize_xml_text(v))))
            .expect("write created text");
        w.write_event(Event::End(BytesEnd::new("dcterms:created")))
            .expect("write created end");
    }
    if let Some(ref v) = props.modified {
        let mut elem = BytesStart::new("dcterms:modified");
        elem.push_attribute(("xsi:type", "dcterms:W3CDTF"));
        w.write_event(Event::Start(elem))
            .expect("write modified start");
        w.write_event(Event::Text(BytesText::new(&crate::core::xml::sanitize_xml_text(v))))
            .expect("write modified text");
        w.write_event(Event::End(BytesEnd::new("dcterms:modified")))
            .expect("write modified end");
    }

    w.write_event(Event::End(BytesEnd::new("cp:coreProperties")))
        .expect("write core end");
    w.into_inner()
}

// ---------------------------------------------------------------------------
// fontTable.xml generator
// ---------------------------------------------------------------------------

/// Build `word/fontTable.xml` listing each embedded font with an
/// `<w:embedRegular r:id="…"/>` reference. Word looks up `<w:rFonts
/// w:ascii="…"/>` names against this table and uses the embedded
/// program when there's a match. Without it, Word silently
/// substitutes Calibri / Cambria for everything regardless of how
/// many TTFs we ship under `/word/fonts/`.
///
/// **Known limitation**: each entry is emitted as `<w:embedRegular>`
/// regardless of whether the underlying program is a regular, bold,
/// italic, or bold-italic face — we don't introspect the font binary
/// to detect the style. If a caller wants Word to pick up a bold-only
/// face, they should embed it under a distinct family name (e.g.
/// `Calibri-Bold`) and reference that name from runs explicitly.
fn generate_font_table_xml(entries: &[(String, String)]) -> Vec<u8> {
    let mut w = Writer::new_with_indent(Vec::new(), b' ', 2);
    w.write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
        .expect("decl");

    let mut fonts = BytesStart::new("w:fonts");
    fonts.push_attribute(("xmlns:w", crate::core::xml::ns::WML_STR));
    fonts.push_attribute(("xmlns:r", crate::core::xml::ns::R_STR));
    w.write_event(Event::Start(fonts)).expect("fonts start");

    for (name, rid) in entries {
        let mut font = BytesStart::new("w:font");
        font.push_attribute(("w:name", name.as_str()));
        w.write_event(Event::Start(font)).expect("font start");

        // <w:embedRegular r:id="rIdN"/> — Word treats this as the regular-weight
        // glyph source for the named font face.
        let mut embed = BytesStart::new("w:embedRegular");
        embed.push_attribute(("r:id", rid.as_str()));
        w.write_event(Event::Empty(embed)).expect("embedRegular");

        w.write_event(Event::End(BytesEnd::new("w:font")))
            .expect("font end");
    }

    w.write_event(Event::End(BytesEnd::new("w:fonts")))
        .expect("fonts end");
    w.into_inner()
}

// ---------------------------------------------------------------------------
// Styles and numbering generators
// ---------------------------------------------------------------------------

fn generate_styles_xml(has_numbering: bool, has_notes: bool) -> Vec<u8> {
    let mut w = Writer::new_with_indent(Vec::new(), b' ', 2);

    w.write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
        .expect("write decl");

    let mut root = BytesStart::new("w:styles");
    root.push_attribute(("xmlns:w", WML_NS));
    w.write_event(Event::Start(root))
        .expect("write styles start");

    write_paragraph_style(&mut w, "Normal", "Normal", None);
    for level in 1u8..=6 {
        let style_id = format!("Heading{level}");
        let name = format!("heading {level}");
        write_paragraph_style(&mut w, &style_id, &name, Some(level - 1));
    }
    if has_numbering {
        write_paragraph_style(&mut w, "ListParagraph", "List Paragraph", None);
    }
    write_code_style(&mut w);

    // These are written unconditionally: a run may carry a footnote_ref
    // without a matching note part, and a dangling w:rStyle is exactly the
    // kind of unresolved reference Word refuses to open.
    let _ = has_notes;
    write_character_style(&mut w, "FootnoteReference", "footnote reference");
    write_character_style(&mut w, "EndnoteReference", "endnote reference");

    w.write_event(Event::End(BytesEnd::new("w:styles")))
        .expect("write styles end");

    w.into_inner()
}

fn write_paragraph_style(
    w: &mut Writer<Vec<u8>>,
    style_id: &str,
    name: &str,
    outline_level: Option<u8>,
) {
    let mut elem = BytesStart::new("w:style");
    elem.push_attribute(("w:type", "paragraph"));
    elem.push_attribute(("w:styleId", style_id));
    w.write_event(Event::Start(elem))
        .expect("write style start");

    let mut name_elem = BytesStart::new("w:name");
    name_elem.push_attribute(("w:val", name));
    w.write_event(Event::Empty(name_elem))
        .expect("write style name");

    // basedOn Normal so heading styles inherit body defaults.
    if outline_level.is_some() {
        let mut based = BytesStart::new("w:basedOn");
        based.push_attribute(("w:val", "Normal"));
        w.write_event(Event::Empty(based)).expect("write basedOn");
    }

    if let Some(level) = outline_level {
        w.write_event(Event::Start(BytesStart::new("w:pPr")))
            .expect("write pPr start");
        // Spacing-before for visual breathing room above the heading.
        let mut sp = BytesStart::new("w:spacing");
        sp.push_attribute(("w:before", "240")); // 12 pt
        sp.push_attribute(("w:after", "120")); //  6 pt
        w.write_event(Event::Empty(sp)).expect("write spacing");
        let mut lvl = BytesStart::new("w:outlineLvl");
        lvl.push_attribute(("w:val", level.to_string().as_str()));
        w.write_event(Event::Empty(lvl)).expect("write outlineLvl");
        w.write_event(Event::End(BytesEnd::new("w:pPr")))
            .expect("write pPr end");

        // Run properties — size & bold per Word's default heading scale.
        // Without this, every <w:pStyle val="HeadingN"/> in the body
        // renders as plain Normal — the headings disappear visually.
        let (sz_half_pt, bold, italic, color) = match level {
            0 => (56, true, false, "2F5496"), // Heading 1: 28 pt
            1 => (44, true, false, "2F5496"), // Heading 2: 22 pt
            2 => (32, true, false, "1F3864"), // Heading 3: 16 pt
            3 => (28, true, true, "2F5496"),  // Heading 4: 14 pt italic
            4 => (24, true, false, "2F5496"), // Heading 5: 12 pt
            _ => (22, true, true, "1F3864"),  // Heading 6: 11 pt italic
        };
        w.write_event(Event::Start(BytesStart::new("w:rPr")))
            .expect("write rPr start");
        if bold {
            w.write_event(Event::Empty(BytesStart::new("w:b")))
                .expect("write b");
        }
        if italic {
            w.write_event(Event::Empty(BytesStart::new("w:i")))
                .expect("write i");
        }
        let mut col = BytesStart::new("w:color");
        col.push_attribute(("w:val", color));
        w.write_event(Event::Empty(col)).expect("write color");
        let sz_str = sz_half_pt.to_string();
        let mut sz = BytesStart::new("w:sz");
        sz.push_attribute(("w:val", sz_str.as_str()));
        w.write_event(Event::Empty(sz)).expect("write sz");
        let mut sz_cs = BytesStart::new("w:szCs");
        sz_cs.push_attribute(("w:val", sz_str.as_str()));
        w.write_event(Event::Empty(sz_cs)).expect("write szCs");
        w.write_event(Event::End(BytesEnd::new("w:rPr")))
            .expect("write rPr end");
    }

    w.write_event(Event::End(BytesEnd::new("w:style")))
        .expect("write style end");
}

fn write_code_style(w: &mut Writer<Vec<u8>>) {
    let mut elem = BytesStart::new("w:style");
    elem.push_attribute(("w:type", "paragraph"));
    elem.push_attribute(("w:styleId", "Code"));
    w.write_event(Event::Start(elem))
        .expect("write code style start");

    let mut name_elem = BytesStart::new("w:name");
    name_elem.push_attribute(("w:val", "Code"));
    w.write_event(Event::Empty(name_elem))
        .expect("write code name");

    // pPr: shading
    w.write_event(Event::Start(BytesStart::new("w:pPr")))
        .expect("write pPr");
    let mut shd = BytesStart::new("w:shd");
    shd.push_attribute(("w:val", "clear"));
    shd.push_attribute(("w:fill", "F0F0F0"));
    shd.push_attribute(("w:color", "auto"));
    w.write_event(Event::Empty(shd)).expect("write code shd");
    w.write_event(Event::End(BytesEnd::new("w:pPr")))
        .expect("write pPr end");

    // rPr: Courier New 10pt
    w.write_event(Event::Start(BytesStart::new("w:rPr")))
        .expect("write rPr");
    let mut fonts = BytesStart::new("w:rFonts");
    fonts.push_attribute(("w:ascii", "Courier New"));
    fonts.push_attribute(("w:hAnsi", "Courier New"));
    w.write_event(Event::Empty(fonts))
        .expect("write code fonts");
    let mut sz = BytesStart::new("w:sz");
    sz.push_attribute(("w:val", "20")); // 10pt = 20 half-points
    w.write_event(Event::Empty(sz)).expect("write code sz");
    w.write_event(Event::End(BytesEnd::new("w:rPr")))
        .expect("write rPr end");

    w.write_event(Event::End(BytesEnd::new("w:style")))
        .expect("write code style end");
}

fn write_character_style(w: &mut Writer<Vec<u8>>, style_id: &str, name: &str) {
    let mut elem = BytesStart::new("w:style");
    elem.push_attribute(("w:type", "character"));
    elem.push_attribute(("w:styleId", style_id));
    w.write_event(Event::Start(elem))
        .expect("write char style start");

    let mut name_elem = BytesStart::new("w:name");
    name_elem.push_attribute(("w:val", name));
    w.write_event(Event::Empty(name_elem))
        .expect("write char style name");

    w.write_event(Event::End(BytesEnd::new("w:style")))
        .expect("write char style end");
}

fn write_abstract_num(
    w: &mut Writer<Vec<u8>>,
    abstract_num_id: u32,
    num_fmt: &str,
    lvl_text: &str,
) {
    let mut elem = BytesStart::new("w:abstractNum");
    elem.push_attribute(("w:abstractNumId", abstract_num_id.to_string().as_str()));
    w.write_event(Event::Start(elem))
        .expect("write abstractNum start");

    let mut lvl = BytesStart::new("w:lvl");
    lvl.push_attribute(("w:ilvl", "0"));
    w.write_event(Event::Start(lvl)).expect("write lvl start");

    let mut fmt = BytesStart::new("w:numFmt");
    fmt.push_attribute(("w:val", num_fmt));
    w.write_event(Event::Empty(fmt)).expect("write numFmt");

    let mut text = BytesStart::new("w:lvlText");
    text.push_attribute(("w:val", lvl_text));
    w.write_event(Event::Empty(text)).expect("write lvlText");

    w.write_event(Event::End(BytesEnd::new("w:lvl")))
        .expect("write lvl end");
    w.write_event(Event::End(BytesEnd::new("w:abstractNum")))
        .expect("write abstractNum end");
}

fn write_num(
    w: &mut Writer<Vec<u8>>,
    num_id: u32,
    abstract_num_id: u32,
    start_override: Option<u32>,
) {
    let mut elem = BytesStart::new("w:num");
    elem.push_attribute(("w:numId", num_id.to_string().as_str()));
    w.write_event(Event::Start(elem)).expect("write num start");

    let mut abs = BytesStart::new("w:abstractNumId");
    abs.push_attribute(("w:val", abstract_num_id.to_string().as_str()));
    w.write_event(Event::Empty(abs))
        .expect("write abstractNumId");

    if let Some(start) = start_override {
        let mut lvl_override = BytesStart::new("w:lvlOverride");
        lvl_override.push_attribute(("w:ilvl", "0"));
        w.write_event(Event::Start(lvl_override))
            .expect("write lvlOverride start");
        let mut so = BytesStart::new("w:startOverride");
        so.push_attribute(("w:val", start.to_string().as_str()));
        w.write_event(Event::Empty(so))
            .expect("write startOverride");
        w.write_event(Event::End(BytesEnd::new("w:lvlOverride")))
            .expect("write lvlOverride end");
    }

    w.write_event(Event::End(BytesEnd::new("w:num")))
        .expect("write num end");
}

// ---------------------------------------------------------------------------
// Value mapping helpers
// ---------------------------------------------------------------------------

fn rgb_to_hex(rgb: [u8; 3]) -> String {
    format!("{:02X}{:02X}{:02X}", rgb[0], rgb[1], rgb[2])
}

fn px_to_emu(px: u32) -> u64 {
    px as u64 * 914400 / 96
}

fn para_align_val(align: &ParagraphAlignment) -> &'static str {
    match align {
        ParagraphAlignment::Left => "left",
        ParagraphAlignment::Center => "center",
        ParagraphAlignment::Right => "right",
        ParagraphAlignment::Justify => "both",
        ParagraphAlignment::Distribute => "distribute",
    }
}

fn underline_style_val(us: &UnderlineStyle) -> &'static str {
    match us {
        UnderlineStyle::Single => "single",
        UnderlineStyle::Double => "double",
        UnderlineStyle::Thick => "thick",
        UnderlineStyle::Dotted => "dotted",
        UnderlineStyle::Dash => "dash",
        UnderlineStyle::DotDash => "dotDash",
        UnderlineStyle::DotDotDash => "dotDotDash",
        UnderlineStyle::Wave => "wave",
        UnderlineStyle::Words => "words",
        UnderlineStyle::None => "none",
    }
}

fn border_style_val(style: &BorderStyle) -> &'static str {
    match style {
        BorderStyle::None => "none",
        BorderStyle::Single => "single",
        BorderStyle::Thick => "thick",
        BorderStyle::Double => "double",
        BorderStyle::Dotted => "dotted",
        BorderStyle::Dashed => "dashed",
        BorderStyle::Wave => "wave",
        BorderStyle::DashSmallGap => "dashSmallGap",
        BorderStyle::Outset => "outset",
        BorderStyle::Inset => "inset",
    }
}

fn list_style_to_fmt(style: Option<&ListStyle>, ordered: bool) -> (&'static str, &'static str) {
    match style {
        Some(ListStyle::Bullet) => ("bullet", "\u{2022}"),
        Some(ListStyle::Decimal) => ("decimal", "%1."),
        Some(ListStyle::LowerRoman) => ("lowerRoman", "%1."),
        Some(ListStyle::UpperRoman) => ("upperRoman", "%1."),
        Some(ListStyle::LowerAlpha) => ("lowerLetter", "%1."),
        Some(ListStyle::UpperAlpha) => ("upperLetter", "%1."),
        Some(ListStyle::Dash) => ("bullet", "\u{2013}"),
        Some(ListStyle::Square) => ("bullet", "\u{25AA}"),
        Some(ListStyle::Circle) => ("bullet", "\u{25CB}"),
        None => {
            if ordered {
                ("decimal", "%1.")
            } else {
                ("bullet", "\u{2022}")
            }
        },
    }
}

// ---------------------------------------------------------------------------
// OpcWriter extension — add_part_raw (raw bytes, no encoding)
// ---------------------------------------------------------------------------

// The existing add_part takes &[u8] and re-encodes; for images we need raw bytes.
// We use the same add_part since it just stores bytes.

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;
    use crate::docx::DocxDocument;
    use std::io::Cursor;

    fn roundtrip(doc: DocxWriter) -> DocxDocument {
        let mut buf = Cursor::new(Vec::new());
        doc.write_to(&mut buf).unwrap();
        buf.set_position(0);
        DocxDocument::from_reader(buf).unwrap()
    }

    /// Read a part from a written package.
    fn part_xml(doc: DocxWriter, name: &str) -> String {
        let mut buf = Cursor::new(Vec::new());
        doc.write_to(&mut buf).unwrap();
        buf.set_position(0);
        let mut zip = zip::ZipArchive::new(buf).unwrap();
        let mut entry = zip.by_name(name).unwrap();
        let mut xml = String::new();
        std::io::Read::read_to_string(&mut entry, &mut xml).unwrap();
        xml
    }

    fn ir_cell(text: &str) -> crate::ir::TableCell {
        crate::ir::TableCell {
            content: vec![crate::ir::Element::Paragraph(crate::ir::Paragraph {
                content: vec![crate::ir::InlineContent::Text(crate::ir::TextSpan {
                    text: text.into(),
                    ..Default::default()
                })],
                ..Default::default()
            })],
            ..Default::default()
        }
    }

    /// `CT_Tbl` is `tblPr, tblGrid, rows` — `tblGrid` has `minOccurs=1`. It was
    /// emitted only when explicit column widths were known, so every table
    /// built from markdown (which carries none) was schema-invalid.
    #[test]
    fn table_without_explicit_widths_still_carries_a_grid() {
        let mut doc = DocxWriter::new();
        let table = crate::ir::Table {
            rows: vec![crate::ir::TableRow {
                cells: vec![ir_cell("a"), ir_cell("b")],
                ..Default::default()
            }],
            ..Default::default()
        };
        doc.add_ir_table(&table);
        let xml = part_xml(doc, "word/document.xml");

        let grid = xml
            .find("<w:tblGrid")
            .expect("w:tblGrid is required by CT_Tbl");
        let tr = xml.find("<w:tr").expect("table must have a row");
        assert!(grid < tr, "tblGrid must precede the rows");
        assert_eq!(xml.matches("<w:gridCol").count(), 2, "one gridCol per column");
    }

    /// The simple `add_table` writer emitted neither `tblPr` nor `tblGrid`.
    #[test]
    fn simple_table_carries_properties_and_grid() {
        let mut doc = DocxWriter::new();
        doc.add_table(&[vec!["a", "b"], vec!["c", "d"]]);
        let xml = part_xml(doc, "word/document.xml");

        let pr = xml.find("<w:tblPr").expect("w:tblPr is required by CT_Tbl");
        let grid = xml
            .find("<w:tblGrid")
            .expect("w:tblGrid is required by CT_Tbl");
        let tr = xml.find("<w:tr").expect("table must have a row");
        assert!(pr < grid && grid < tr, "order must be tblPr, tblGrid, rows");
        assert_eq!(xml.matches("<w:gridCol").count(), 2);
    }

    /// `CT_Numbering` is `numPicBullet*, abstractNum*, num*` — every
    /// `abstractNum` must precede every `num`. Emitting them interleaved,
    /// one pair per list, put an `abstractNum` after a `num`.
    #[test]
    fn numbering_puts_every_abstract_definition_before_every_instance() {
        let mut doc = DocxWriter::new();
        for ordered in [true, false, true] {
            doc.add_ir_list(&crate::ir::List {
                ordered,
                items: vec![crate::ir::ListItem {
                    content: crate::ir::inline_to_element_block(vec![
                        crate::ir::InlineContent::Text(crate::ir::TextSpan {
                            text: "item".into(),
                            ..Default::default()
                        }),
                    ]),
                    nested: None,
                }],
                ..Default::default()
            });
        }
        let xml = part_xml(doc, "word/numbering.xml");

        let last_abstract = xml.rfind("<w:abstractNum ").expect("abstractNum");
        let first_num = xml.find("<w:num ").expect("num");
        assert!(
            last_abstract < first_num,
            "every abstractNum must precede every num; got:\n{xml}"
        );
    }

    /// `CT_Anchor` orders the wrap group before `docPr`. Both anchor writers
    /// emitted `docPr` first.
    #[test]
    fn floating_anchor_puts_the_wrap_before_doc_pr() {
        let mut doc = DocxWriter::new();
        doc.add_text_box(&crate::ir::TextBox::default());
        let xml = part_xml(doc, "word/document.xml");

        let wrap = xml
            .find("<wp:wrap")
            .expect("CT_Anchor requires an EG_WrapType element");
        let doc_pr = xml.find("<wp:docPr").expect("CT_Anchor requires docPr");
        assert!(wrap < doc_pr, "the wrap element must precede docPr");
    }

    /// `CT_PageMar` declares seven attributes, all `use="required"`.
    #[test]
    fn page_margins_carry_every_required_attribute() {
        let mut doc = DocxWriter::new();
        doc.add_paragraph("x");
        doc.set_section_props(
            Some(crate::ir::PageSetup::default()),
            None,
            crate::ir::SectionBreakType::NextPage,
        );
        let xml = part_xml(doc, "word/document.xml");

        let mar = xml.find("<w:pgMar").expect("w:pgMar");
        let end = xml[mar..].find("/>").unwrap() + mar;
        let tag = &xml[mar..end];
        for attr in [
            "w:top", "w:right", "w:bottom", "w:left", "w:header", "w:footer", "w:gutter",
        ] {
            assert!(tag.contains(attr), "pgMar missing required {attr}: {tag}");
        }
    }

    /// `CT_PPrBase` orders `spacing` before `ind`.
    #[test]
    fn paragraph_properties_put_spacing_before_indent() {
        let mut doc = DocxWriter::new();
        doc.add_ir_paragraph(
            &[Run::new("x")],
            Some(IrParaProps {
                indent_left_twips: Some(720),
                space_after_twips: Some(240),
                ..Default::default()
            }),
        );
        let xml = part_xml(doc, "word/document.xml");

        let spacing = xml.find("<w:spacing").expect("w:spacing");
        let ind = xml.find("<w:ind").expect("w:ind");
        assert!(spacing < ind, "CT_PPrBase requires spacing before ind");
    }

    fn ir_list(ordered: bool, text: &str) -> crate::ir::Element {
        crate::ir::Element::List(crate::ir::List {
            ordered,
            items: vec![crate::ir::ListItem {
                content: crate::ir::inline_to_element_block(vec![crate::ir::InlineContent::Text(
                    crate::ir::TextSpan {
                        text: text.into(),
                        ..Default::default()
                    },
                )]),
                nested: None,
            }],
            ..Default::default()
        })
    }

    fn all_parts(doc: DocxWriter) -> std::collections::HashMap<String, String> {
        let mut buf = Cursor::new(Vec::new());
        doc.write_to(&mut buf).unwrap();
        buf.set_position(0);
        let mut zip = zip::ZipArchive::new(buf).unwrap();
        let mut out = std::collections::HashMap::new();
        for i in 0..zip.len() {
            let mut e = zip.by_index(i).unwrap();
            let name = e.name().to_string();
            let mut xml = String::new();
            if std::io::Read::read_to_string(&mut e, &mut xml).is_ok() {
                out.insert(name, xml);
            }
        }
        out
    }

    /// A list in a header, footer or note emits `ListParagraph` and
    /// `numId=1` exactly like one in the body, but the numbering part and
    /// the style were gated on the body alone. The result is schema-valid
    /// and has a `w:numId` pointing at a part that does not exist.
    #[test]
    fn a_list_outside_the_body_still_gets_its_numbering_and_style() {
        let mut doc = DocxWriter::new();
        doc.add_paragraph("body text, no lists at top level");
        doc.add_section_header(HfType::DefaultHeader, vec![ir_list(false, "hdr item")]);
        let parts = all_parts(doc);

        let header = parts
            .iter()
            .find(|(k, _)| k.starts_with("word/header"))
            .map(|(_, v)| v.clone())
            .expect("header part");
        assert!(header.contains("w:numId"), "the header list should carry numbering");
        assert!(
            parts.contains_key("word/numbering.xml"),
            "a numId with no numbering.xml is a dangling reference; parts: {:?}",
            parts.keys().collect::<Vec<_>>()
        );
        let styles = &parts["word/styles.xml"];
        assert!(
            styles.contains(r#"w:styleId="ListParagraph""#),
            "ListParagraph referenced but not defined"
        );
    }

    /// A run may carry a footnote reference with no matching note part, so
    /// the reference character styles must always be defined.
    #[test]
    fn note_reference_styles_are_always_defined() {
        let mut doc = DocxWriter::new();
        doc.add_ir_paragraph(
            &[
                Run::new("dangling"),
                Run {
                    footnote_ref: Some(7),
                    ..Default::default()
                },
            ],
            None,
        );
        let parts = all_parts(doc);
        let body = &parts["word/document.xml"];
        let styles = &parts["word/styles.xml"];
        if body.contains("FootnoteReference") {
            assert!(
                styles.contains(r#"w:styleId="FootnoteReference""#),
                "FootnoteReference referenced but not defined"
            );
        }
    }

    /// Word expects the separator and continuationSeparator notes.
    #[test]
    fn notes_part_carries_the_separator_notes() {
        let mut doc = DocxWriter::new();
        doc.add_footnote(1, &[crate::ir::Element::Paragraph(Default::default())]);
        let parts = all_parts(doc);
        let notes = &parts["word/footnotes.xml"];
        assert!(notes.contains(r#"w:type="separator""#), "missing separator note: {notes}");
        assert!(
            notes.contains(r#"w:type="continuationSeparator""#),
            "missing continuationSeparator note: {notes}"
        );
    }

    /// An ordered list nested in a cell must not silently become bullets.
    #[test]
    fn a_nested_ordered_list_uses_the_ordered_numbering_definition() {
        let mut doc = DocxWriter::new();
        let table = crate::ir::Table {
            rows: vec![crate::ir::TableRow {
                cells: vec![crate::ir::TableCell {
                    content: vec![ir_list(true, "first")],
                    ..Default::default()
                }],
                ..Default::default()
            }],
            ..Default::default()
        };
        doc.add_ir_table(&table);
        let xml = part_xml(doc, "word/document.xml");
        assert!(
            xml.contains(r#"<w:numId w:val="2"/>"#),
            "an ordered list must not use the bullet definition: {xml}"
        );
    }

    /// `w:type="first"` does nothing without `w:titlePg`; a `w:type="even"`
    /// header does nothing without `w:evenAndOddHeaders` in settings.xml.
    #[test]
    fn first_and_even_page_headers_carry_the_switches_that_enable_them() {
        let mut doc = DocxWriter::new();
        doc.add_paragraph("x");
        doc.add_section_header(HfType::FirstPageHeader, vec![]);
        doc.add_section_header(HfType::EvenPageHeader, vec![]);
        doc.set_section_props(
            Some(crate::ir::PageSetup::default()),
            None,
            crate::ir::SectionBreakType::NextPage,
        );
        let parts = all_parts(doc);
        assert!(
            parts["word/document.xml"].contains("<w:titlePg/>"),
            "a first-page header needs w:titlePg"
        );
        let settings = parts
            .get("word/settings.xml")
            .expect("an even-page header needs settings.xml");
        assert!(
            settings.contains("<w:evenAndOddHeaders/>"),
            "settings.xml must enable even/odd headers: {settings}"
        );
    }

    /// At most one header/footer reference of each type per section.
    #[test]
    fn section_properties_carry_one_reference_per_type() {
        let mut doc = DocxWriter::new();
        doc.add_paragraph("x");
        for _ in 0..3 {
            doc.add_section_header(HfType::DefaultHeader, vec![]);
        }
        doc.set_section_props(
            Some(crate::ir::PageSetup::default()),
            None,
            crate::ir::SectionBreakType::NextPage,
        );
        let xml = part_xml(doc, "word/document.xml");
        assert_eq!(
            xml.matches(r#"<w:headerReference w:type="default""#)
                .count(),
            1,
            "duplicate header references of one type: {xml}"
        );
    }

    #[test]
    fn rich_run_bold_italic() {
        let mut doc = DocxWriter::new();
        doc.add_rich_paragraph(&[
            Run::new("Hello ").bold(),
            Run::new("world").italic().color("FF0000"),
        ]);
        let parsed = roundtrip(doc);
        let text = parsed.plain_text();
        assert!(text.contains("Hello"));
        assert!(text.contains("world"));
    }

    #[test]
    fn alignment_center() {
        let mut doc = DocxWriter::new();
        doc.add_paragraph_aligned("Centred", Alignment::Center);
        let parsed = roundtrip(doc);
        assert!(parsed.plain_text().contains("Centred"));
    }

    #[test]
    fn page_break_roundtrip() {
        let mut doc = DocxWriter::new();
        doc.add_paragraph("Before");
        doc.add_page_break();
        doc.add_paragraph("After");
        let parsed = roundtrip(doc);
        let text = parsed.plain_text();
        assert!(text.contains("Before"));
        assert!(text.contains("After"));
    }

    #[test]
    fn font_size_and_name() {
        let mut doc = DocxWriter::new();
        doc.add_rich_paragraph(&[Run::new("Big text").font_size(24.0).font("Arial")]);
        let parsed = roundtrip(doc);
        assert!(parsed.plain_text().contains("Big text"));
    }

    #[test]
    fn underline_strikethrough() {
        let mut doc = DocxWriter::new();
        doc.add_rich_paragraph(&[
            Run::new("under").underline(),
            Run::new(" strike").strikethrough(),
        ]);
        let parsed = roundtrip(doc);
        let text = parsed.plain_text();
        assert!(text.contains("under"));
        assert!(text.contains("strike"));
    }

    #[test]
    fn column_break_roundtrip() {
        let mut doc = DocxWriter::new();
        doc.add_paragraph("Col 1");
        doc.add_column_break();
        doc.add_paragraph("Col 2");
        let parsed = roundtrip(doc);
        let text = parsed.plain_text();
        assert!(text.contains("Col 1"));
        assert!(text.contains("Col 2"));
    }

    #[test]
    fn ir_paragraph_with_props() {
        let mut doc = DocxWriter::new();
        let props = IrParaProps {
            alignment: Some(ParagraphAlignment::Center),
            space_before_twips: Some(240),
            ..Default::default()
        };
        doc.add_ir_paragraph(&[Run::new("Aligned")], Some(props));
        let parsed = roundtrip(doc);
        assert!(parsed.plain_text().contains("Aligned"));
    }

    #[test]
    fn code_block_roundtrip() {
        let mut doc = DocxWriter::new();
        doc.add_code_block("fn main() {\n    println!(\"hello\");\n}");
        let parsed = roundtrip(doc);
        assert!(parsed.plain_text().contains("fn main"));
    }
}
