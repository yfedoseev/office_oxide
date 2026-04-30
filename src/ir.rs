use crate::format::DocumentFormat;

fn default_true() -> bool {
    true
}

// ── Enums ────────────────────────────────────────────────────────────────────

/// Underline style for a text span.
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum UnderlineStyle {
    Single,
    Double,
    Thick,
    Dotted,
    Dash,
    DotDash,
    DotDotDash,
    Wave,
    Words,
    None,
}

/// Paragraph text alignment.
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum ParagraphAlignment {
    Left,
    Center,
    Right,
    Justify,
    Distribute,
}

/// Line spacing rule for a paragraph.
/// `Auto(240)` = single, `Auto(360)` = 1.5×, `Auto(480)` = double.
/// `Multiple` uses the same OOXML rule as `Auto`.
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum LineSpacing {
    Auto(u32),
    Multiple(u32),
    Exact(u32),
    AtLeast(u32),
}

/// Border style.
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum BorderStyle {
    None,
    Single,
    Thick,
    Double,
    Dotted,
    Dashed,
    Wave,
    DashSmallGap,
    Outset,
    Inset,
}

/// Vertical alignment within a table cell.
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum CellVerticalAlign {
    Top,
    Center,
    Bottom,
}

/// Horizontal alignment of a table on the page.
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum TableAlignment {
    Left,
    Center,
    Right,
}

/// Text direction within a cell or frame.
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum TextDirection {
    LrTb,
    TbRl,
    BtLr,
}

/// Raster / vector image format.
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum ImageFormat {
    Png,
    Jpeg,
    Gif,
    Tiff,
    Bmp,
    Emf,
    Wmf,
}

impl ImageFormat {
    pub fn content_type(&self) -> &'static str {
        match self {
            Self::Png => "image/png",
            Self::Jpeg => "image/jpeg",
            Self::Gif => "image/gif",
            Self::Tiff => "image/tiff",
            Self::Bmp => "image/bmp",
            Self::Emf => "image/x-emf",
            Self::Wmf => "image/x-wmf",
        }
    }

    pub fn extension(&self) -> &'static str {
        match self {
            Self::Png => "png",
            Self::Jpeg => "jpg",
            Self::Gif => "gif",
            Self::Tiff => "tiff",
            Self::Bmp => "bmp",
            Self::Emf => "emf",
            Self::Wmf => "wmf",
        }
    }
}

/// How an image is positioned relative to surrounding text.
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum ImagePositioning {
    Inline,
    Floating(FloatingImage),
}

impl Default for ImagePositioning {
    fn default() -> Self {
        Self::Inline
    }
}

/// Section break type.
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum SectionBreakType {
    Continuous,
    NextPage,
    EvenPage,
    OddPage,
}

impl Default for SectionBreakType {
    fn default() -> Self {
        Self::Continuous
    }
}

/// Vertical text alignment (superscript / subscript).
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum VerticalAlign {
    Superscript,
    Subscript,
    Baseline,
}

/// Anchor reference for a floating object.
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum FloatAnchor {
    Page,
    Margin,
    Column,
    Paragraph,
}

impl Default for FloatAnchor {
    fn default() -> Self {
        Self::Page
    }
}

/// Text wrap mode around a floating object.
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum TextWrap {
    Square,
    Tight,
    Through,
    TopAndBottom,
    Behind,
    InFront,
}

impl Default for TextWrap {
    fn default() -> Self {
        Self::Square
    }
}

/// List marker style.
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum ListStyle {
    Bullet,
    Decimal,
    LowerRoman,
    UpperRoman,
    LowerAlpha,
    UpperAlpha,
    Dash,
    Square,
    Circle,
}

// ── New structs ───────────────────────────────────────────────────────────────

/// A single border line definition.
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
pub struct BorderLine {
    pub style: BorderStyle,
    pub color: Option<[u8; 3]>,
    /// Line width in eighths of a point.
    pub size: Option<u32>,
    /// Spacing between border and content in points.
    pub space: Option<u32>,
}

/// Full border set for a table (all six edges).
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
pub struct TableBorder {
    pub top: Option<BorderLine>,
    pub bottom: Option<BorderLine>,
    pub left: Option<BorderLine>,
    pub right: Option<BorderLine>,
    pub inside_h: Option<BorderLine>,
    pub inside_v: Option<BorderLine>,
}

/// Page geometry and margins (all values in twips).
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
pub struct PageSetup {
    pub width_twips: u32,
    pub height_twips: u32,
    pub margin_top_twips: u32,
    pub margin_bottom_twips: u32,
    pub margin_left_twips: u32,
    pub margin_right_twips: u32,
    pub landscape: bool,
    /// Distance from top edge to header in twips (default 720 = 0.5").
    pub header_distance_twips: u32,
    /// Distance from bottom edge to footer in twips (default 720 = 0.5").
    pub footer_distance_twips: u32,
}

impl Default for PageSetup {
    fn default() -> Self {
        Self {
            width_twips: 12240,
            height_twips: 15840,
            margin_top_twips: 1440,
            margin_bottom_twips: 1440,
            margin_left_twips: 1800,
            margin_right_twips: 1800,
            landscape: false,
            header_distance_twips: 720,
            footer_distance_twips: 720,
        }
    }
}

/// Multi-column layout for a section.
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, Default, serde::Serialize, serde::Deserialize)]
pub struct ColumnLayout {
    pub count: u32,
    pub space_twips: Option<u32>,
    pub separator: bool,
    /// Per-column widths in twips (overrides uniform spacing when non-empty).
    #[serde(default)]
    pub column_widths_twips: Vec<u32>,
}

/// Paragraph border (four sides plus between-paragraph rule).
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
pub struct ParagraphBorder {
    pub top: Option<BorderLine>,
    pub bottom: Option<BorderLine>,
    pub left: Option<BorderLine>,
    pub right: Option<BorderLine>,
    pub between: Option<BorderLine>,
}

/// Per-edge cell padding (all values in twips).
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
pub struct CellPadding {
    pub top_twips: Option<u32>,
    pub bottom_twips: Option<u32>,
    pub left_twips: Option<u32>,
    pub right_twips: Option<u32>,
}

/// Positioning data for a floating (non-inline) image.
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
pub struct FloatingImage {
    pub x_emu: i64,
    pub y_emu: i64,
    pub width_emu: u64,
    pub height_emu: u64,
    pub h_anchor: FloatAnchor,
    pub v_anchor: FloatAnchor,
    pub text_wrap: TextWrap,
    #[serde(default)]
    pub allow_overlap: bool,
}

/// A header or footer containing block elements.
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, Default, serde::Serialize, serde::Deserialize)]
pub struct HeaderFooter {
    pub content: Vec<Element>,
}

/// A floating text box containing block elements.
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, Default, serde::Serialize, serde::Deserialize)]
pub struct TextBox {
    pub content: Vec<Element>,
    pub width_emu: Option<u64>,
    pub height_emu: Option<u64>,
    pub x_emu: Option<i64>,
    pub y_emu: Option<i64>,
    /// Horizontal anchor reference frame.
    #[serde(default)]
    pub h_anchor: FloatAnchor,
    /// Vertical anchor reference frame.
    #[serde(default)]
    pub v_anchor: FloatAnchor,
    /// Text wrap mode around this box.
    #[serde(default)]
    pub wrap: TextWrap,
}

/// A footnote or endnote body.
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, Default, serde::Serialize, serde::Deserialize)]
pub struct Note {
    pub id: u32,
    pub content: Vec<Element>,
    /// Optional custom marker text (when absent the auto-number is used).
    pub marker: Option<String>,
}

/// An inline reference mark pointing to a footnote or endnote.
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, Default, serde::Serialize, serde::Deserialize)]
pub struct FootnoteRef {
    pub note_id: u32,
    /// Optional custom marker text (when absent the auto-number is used).
    pub marker: Option<String>,
}

/// A preformatted code block with an optional language tag.
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, Default, serde::Serialize, serde::Deserialize)]
pub struct CodeBlock {
    pub language: Option<String>,
    pub content: String,
}

// ── Core document types ───────────────────────────────────────────────────────

/// A format-agnostic intermediate representation of a document.
#[derive(Debug, Clone, PartialEq, Default, serde::Serialize, serde::Deserialize)]
pub struct DocumentIR {
    /// Document-level metadata (format, title, etc.).
    pub metadata: Metadata,
    /// Ordered list of sections (pages, worksheets, slides, etc.).
    pub sections: Vec<Section>,
}

/// Document-level metadata extracted from the source file.
#[derive(Debug, Clone, PartialEq, Default, serde::Serialize, serde::Deserialize)]
pub struct Metadata {
    /// The source format this document was parsed from.
    pub format: DocumentFormat,
    /// Optional document title from core properties.
    pub title: Option<String>,
    /// Document author.
    pub author: Option<String>,
    /// Document subject.
    pub subject: Option<String>,
    /// Keywords / tags.
    #[serde(default)]
    pub keywords: Vec<String>,
    /// Creation date (ISO-8601 string).
    pub created: Option<String>,
    /// Last-modified date (ISO-8601 string).
    pub modified: Option<String>,
    /// Document description / comments.
    pub description: Option<String>,
}

/// A logical section (DOCX: section break, XLSX: worksheet, PPTX: slide).
#[derive(Debug, Clone, PartialEq, Default, serde::Serialize, serde::Deserialize)]
pub struct Section {
    /// Optional section title (e.g. slide title or worksheet name).
    pub title: Option<String>,
    /// Content elements within this section.
    pub elements: Vec<Element>,
    /// Page geometry for this section.
    pub page_setup: Option<PageSetup>,
    /// Multi-column layout, if any.
    pub columns: Option<ColumnLayout>,
    /// How this section break was introduced.
    #[serde(default)]
    pub break_type: SectionBreakType,
    pub header: Option<HeaderFooter>,
    pub footer: Option<HeaderFooter>,
    pub first_page_header: Option<HeaderFooter>,
    pub first_page_footer: Option<HeaderFooter>,
    pub even_page_header: Option<HeaderFooter>,
    pub even_page_footer: Option<HeaderFooter>,
}

/// A block-level content element.
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
#[serde(tag = "type", rename_all = "snake_case")]
pub enum Element {
    /// A heading with a numeric level (1–6).
    Heading(Heading),
    /// A paragraph of inline content.
    Paragraph(Paragraph),
    /// A table.
    Table(Table),
    /// An ordered or unordered list.
    List(List),
    /// An embedded image.
    Image(Image),
    /// A horizontal rule / thematic break.
    ThematicBreak,
    /// A floating or anchored text box.
    TextBox(TextBox),
    /// A hard page break.
    PageBreak,
    /// A column break.
    ColumnBreak,
    /// A footnote body (block-level, appears in footnote area).
    Footnote(Note),
    /// An endnote body (block-level, appears in endnote area).
    Endnote(Note),
    /// A preformatted code block.
    CodeBlock(CodeBlock),
}

/// A heading element with a nesting level.
#[derive(Debug, Clone, PartialEq, Default, serde::Serialize, serde::Deserialize)]
pub struct Heading {
    /// Heading level 1–6 (1 = largest).
    pub level: u8,
    /// Inline content of the heading.
    pub content: Vec<InlineContent>,
}

/// A paragraph of inline content.
#[derive(Debug, Clone, PartialEq, Default, serde::Serialize, serde::Deserialize)]
pub struct Paragraph {
    /// Inline runs making up this paragraph.
    pub content: Vec<InlineContent>,
    /// Horizontal alignment.
    pub alignment: Option<ParagraphAlignment>,
    /// Left indent in twips.
    pub indent_left_twips: Option<i32>,
    /// Right indent in twips.
    pub indent_right_twips: Option<i32>,
    /// First-line indent in twips (negative = hanging).
    pub first_line_indent_twips: Option<i32>,
    /// Space before the paragraph in twips.
    pub space_before_twips: Option<u32>,
    /// Space after the paragraph in twips.
    pub space_after_twips: Option<u32>,
    /// Line spacing rule.
    pub line_spacing: Option<LineSpacing>,
    /// Background / shading colour (RGB).
    pub background_color: Option<[u8; 3]>,
    /// Paragraph borders.
    pub border: Option<ParagraphBorder>,
    #[serde(default)]
    pub keep_with_next: bool,
    #[serde(default)]
    pub keep_together: bool,
    #[serde(default)]
    pub page_break_before: bool,
    /// Outline level (0 = body text, 1–9 = heading levels).
    pub outline_level: Option<u8>,
}

/// Inline content within a paragraph or heading.
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
#[serde(tag = "type", rename_all = "snake_case")]
pub enum InlineContent {
    /// A styled text span.
    Text(TextSpan),
    /// A line break within a paragraph.
    LineBreak,
    /// An inline footnote reference mark.
    FootnoteRef(FootnoteRef),
    /// An inline endnote reference mark.
    EndnoteRef(FootnoteRef),
}

/// A styled run of text.
#[derive(Debug, Clone, PartialEq, Default, serde::Serialize, serde::Deserialize)]
pub struct TextSpan {
    /// The text content.
    pub text: String,
    /// Whether the text is bold.
    pub bold: bool,
    /// Whether the text is italic.
    pub italic: bool,
    /// Whether the text has strikethrough.
    pub strikethrough: bool,
    /// Optional hyperlink URL.
    pub hyperlink: Option<String>,
    /// Font size in half-points (e.g. 24 = 12 pt).
    pub font_size_half_pt: Option<u32>,
    /// Foreground colour (RGB).
    pub color: Option<[u8; 3]>,
    /// Underline style, if any.
    pub underline: Option<UnderlineStyle>,
    /// Font family name.
    pub font_name: Option<String>,
    /// Highlight / background colour (RGB).
    pub highlight: Option<[u8; 3]>,
    /// Superscript / subscript alignment.
    pub vertical_align: Option<VerticalAlign>,
    #[serde(default)]
    pub all_caps: bool,
    #[serde(default)]
    pub small_caps: bool,
    /// Character spacing in half-points (negative = condensed).
    pub char_spacing_half_pt: Option<i32>,
}

impl TextSpan {
    /// Create a plain (unformatted) text span.
    pub fn plain(text: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            ..Default::default()
        }
    }
}

/// A table with rows and cells.
#[derive(Debug, Clone, PartialEq, Default, serde::Serialize, serde::Deserialize)]
pub struct Table {
    /// Rows in the table (first row is header when `is_header = true`).
    pub rows: Vec<TableRow>,
    /// Column widths in twips (may be shorter than the actual column count).
    #[serde(default)]
    pub column_widths_twips: Vec<u32>,
    /// Table-level borders.
    pub border: Option<TableBorder>,
    /// Horizontal alignment of the table on the page.
    pub alignment: Option<TableAlignment>,
    /// Default cell padding in twips (applied to all cells).
    pub cell_padding_twips: Option<u32>,
    /// Optional caption string.
    pub caption: Option<String>,
    /// Total table width in twips (`None` = auto).
    pub width_twips: Option<u32>,
    /// Left indent of the table from the margin in twips.
    pub indent_left_twips: Option<i32>,
}

/// A single row within a table.
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
pub struct TableRow {
    /// Cells within this row.
    pub cells: Vec<TableCell>,
    /// Whether this row is a header row.
    pub is_header: bool,
    /// Row height in twips, if set explicitly.
    pub height_twips: Option<u32>,
    /// Whether the row may break across pages.
    #[serde(default = "default_true")]
    pub allow_break: bool,
    /// Whether this row is repeated as a header on subsequent pages.
    #[serde(default)]
    pub repeat_as_header: bool,
}

impl Default for TableRow {
    fn default() -> Self {
        Self {
            cells: Vec::new(),
            is_header: false,
            height_twips: None,
            allow_break: true,
            repeat_as_header: false,
        }
    }
}

/// A single cell within a table row.
#[derive(Debug, Clone, PartialEq, Default, serde::Serialize, serde::Deserialize)]
pub struct TableCell {
    /// Block elements inside the cell.
    pub content: Vec<Element>,
    /// Number of columns this cell spans.
    pub col_span: u32,
    /// Number of rows this cell spans.
    pub row_span: u32,
    /// Cell background / shading colour (RGB).
    pub background_color: Option<[u8; 3]>,
    /// Cell-level borders.
    pub border: Option<TableBorder>,
    /// Vertical alignment within the cell.
    pub vertical_align: Option<CellVerticalAlign>,
    /// Horizontal alignment of text within the cell.
    pub text_align: Option<ParagraphAlignment>,
    /// Cell width in twips.
    pub width_twips: Option<u32>,
    /// Per-edge cell padding.
    pub padding: Option<CellPadding>,
    /// Text direction within the cell.
    pub text_direction: Option<TextDirection>,
}

/// An ordered or unordered list.
#[derive(Debug, Clone, PartialEq, Default, serde::Serialize, serde::Deserialize)]
pub struct List {
    /// `true` = numbered list, `false` = bullet list.
    pub ordered: bool,
    /// Items in the list.
    pub items: Vec<ListItem>,
    /// Starting number for ordered lists.
    pub start_number: Option<u32>,
    /// Marker / numbering style.
    pub style: Option<ListStyle>,
    /// Nesting depth (0 = top-level).
    #[serde(default)]
    pub level: u8,
}

/// A single item within a list.
#[derive(Debug, Clone, PartialEq, Default, serde::Serialize, serde::Deserialize)]
pub struct ListItem {
    /// Block-level content of this item (typically a single Paragraph).
    pub content: Vec<Element>,
    /// Optional nested sub-list.
    pub nested: Option<List>,
}

/// An embedded image reference.
#[derive(Debug, Clone, PartialEq, Default, serde::Serialize, serde::Deserialize)]
pub struct Image {
    /// Optional alt-text description of the image.
    pub alt_text: Option<String>,
    /// Raw image bytes, if extracted.
    pub data: Option<Vec<u8>>,
    /// Pixel format of the image data.
    pub format: Option<ImageFormat>,
    /// Display width in EMUs.
    pub display_width_emu: Option<u64>,
    /// Display height in EMUs.
    pub display_height_emu: Option<u64>,
    /// Source image pixel width.
    pub pixel_width: Option<u32>,
    /// Source image pixel height.
    pub pixel_height: Option<u32>,
    /// Whether the image is purely decorative (no semantic content).
    #[serde(default)]
    pub decorative: bool,
    /// Inline vs. floating positioning.
    #[serde(default)]
    pub positioning: ImagePositioning,
}
