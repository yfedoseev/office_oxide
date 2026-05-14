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
    /// Single underline.
    Single,
    /// Double underline.
    Double,
    /// Thick underline.
    Thick,
    /// Dotted underline.
    Dotted,
    /// Dashed underline.
    Dash,
    /// Dot-dash underline.
    DotDash,
    /// Dot-dot-dash underline.
    DotDotDash,
    /// Wavy underline.
    Wave,
    /// Underline applied only to words (not spaces).
    Words,
    /// No underline.
    None,
}

/// Paragraph text alignment.
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum ParagraphAlignment {
    /// Left-aligned.
    Left,
    /// Centered.
    Center,
    /// Right-aligned.
    Right,
    /// Justified (both edges).
    Justify,
    /// Distributed (space between characters).
    Distribute,
}

/// Line spacing rule for a paragraph.
/// `Auto(240)` = single, `Auto(360)` = 1.5×, `Auto(480)` = double.
/// `Multiple` uses the same OOXML rule as `Auto`.
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum LineSpacing {
    /// Automatic line height scaled by the given value (in twentieths of a point).
    Auto(u32),
    /// Multiple of normal line height (same units as `Auto`).
    Multiple(u32),
    /// Exact line height in twentieths of a point.
    Exact(u32),
    /// At-least line height in twentieths of a point.
    AtLeast(u32),
}

/// Border style.
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum BorderStyle {
    /// No border.
    None,
    /// Single-line border.
    Single,
    /// Thick single-line border.
    Thick,
    /// Double-line border.
    Double,
    /// Dotted border.
    Dotted,
    /// Dashed border.
    Dashed,
    /// Wavy border.
    Wave,
    /// Dashed border with small gaps.
    DashSmallGap,
    /// Outset (3-D) border.
    Outset,
    /// Inset (3-D) border.
    Inset,
}

/// Vertical alignment within a table cell.
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum CellVerticalAlign {
    /// Align content to the top of the cell.
    Top,
    /// Align content to the middle of the cell.
    Center,
    /// Align content to the bottom of the cell.
    Bottom,
}

/// Horizontal alignment of a table on the page.
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum TableAlignment {
    /// Table aligned to the left margin.
    Left,
    /// Table centered on the page.
    Center,
    /// Table aligned to the right margin.
    Right,
}

/// Text direction within a cell or frame.
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum TextDirection {
    /// Left-to-right, top-to-bottom (default).
    LrTb,
    /// Top-to-bottom, right-to-left (vertical CJK).
    TbRl,
    /// Bottom-to-top, left-to-right (rotated).
    BtLr,
}

/// Raster / vector image format.
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum ImageFormat {
    /// PNG raster image.
    Png,
    /// JPEG raster image.
    Jpeg,
    /// GIF raster image.
    Gif,
    /// TIFF raster image.
    Tiff,
    /// BMP raster image.
    Bmp,
    /// Enhanced Metafile vector image.
    Emf,
    /// Windows Metafile vector image.
    Wmf,
}

impl ImageFormat {
    /// Returns the MIME content-type string for this image format.
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

    /// Returns the file extension (without leading dot) for this image format.
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
#[derive(Debug, Clone, Default, PartialEq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum ImagePositioning {
    /// Image flows inline with surrounding text.
    #[default]
    Inline,
    /// Image is anchored at a fixed position with text wrap.
    Floating(FloatingImage),
}

/// Section break type.
#[allow(dead_code)]
#[derive(Debug, Clone, Default, PartialEq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum SectionBreakType {
    /// Continuous section break (no page break).
    #[default]
    Continuous,
    /// Section starts on the next page.
    NextPage,
    /// Section starts on the next even-numbered page.
    EvenPage,
    /// Section starts on the next odd-numbered page.
    OddPage,
}

/// Vertical text alignment (superscript / subscript).
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum VerticalAlign {
    /// Text raised above the baseline (superscript).
    Superscript,
    /// Text lowered below the baseline (subscript).
    Subscript,
    /// Normal baseline position.
    Baseline,
}

/// Anchor reference for a floating object.
#[allow(dead_code)]
#[derive(Debug, Clone, Default, PartialEq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum FloatAnchor {
    /// Anchored relative to the page.
    #[default]
    Page,
    /// Anchored relative to the page margin.
    Margin,
    /// Anchored relative to the column.
    Column,
    /// Anchored relative to the paragraph.
    Paragraph,
}

/// Text wrap mode around a floating object.
#[allow(dead_code)]
#[derive(Debug, Clone, Default, PartialEq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum TextWrap {
    /// Text wraps around a rectangular bounding box.
    #[default]
    Square,
    /// Text wraps tightly around the object contour.
    Tight,
    /// Text wraps through the object's contour.
    Through,
    /// Text appears only above and below the object.
    TopAndBottom,
    /// Object appears behind the text layer.
    Behind,
    /// Object appears in front of the text layer.
    InFront,
}

/// List marker style.
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum ListStyle {
    /// Filled circle bullet marker.
    Bullet,
    /// Decimal number marker (1, 2, 3, …).
    Decimal,
    /// Lowercase Roman numeral marker (i, ii, iii, …).
    LowerRoman,
    /// Uppercase Roman numeral marker (I, II, III, …).
    UpperRoman,
    /// Lowercase alphabetic marker (a, b, c, …).
    LowerAlpha,
    /// Uppercase alphabetic marker (A, B, C, …).
    UpperAlpha,
    /// Dash marker.
    Dash,
    /// Square bullet marker.
    Square,
    /// Open circle bullet marker.
    Circle,
}

// ── New structs ───────────────────────────────────────────────────────────────

/// A single border line definition.
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
pub struct BorderLine {
    /// Border line style.
    pub style: BorderStyle,
    /// Border colour (RGB), if specified.
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
    /// Top border of the table.
    pub top: Option<BorderLine>,
    /// Bottom border of the table.
    pub bottom: Option<BorderLine>,
    /// Left border of the table.
    pub left: Option<BorderLine>,
    /// Right border of the table.
    pub right: Option<BorderLine>,
    /// Horizontal interior borders between rows.
    pub inside_h: Option<BorderLine>,
    /// Vertical interior borders between columns.
    pub inside_v: Option<BorderLine>,
}

/// Page geometry and margins (all values in twips).
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
pub struct PageSetup {
    /// Page width in twips.
    pub width_twips: u32,
    /// Page height in twips.
    pub height_twips: u32,
    /// Top margin in twips.
    pub margin_top_twips: u32,
    /// Bottom margin in twips.
    pub margin_bottom_twips: u32,
    /// Left margin in twips.
    pub margin_left_twips: u32,
    /// Right margin in twips.
    pub margin_right_twips: u32,
    /// Whether the page is in landscape orientation.
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
    /// Number of columns.
    pub count: u32,
    /// Space between columns in twips.
    pub space_twips: Option<u32>,
    /// Whether a vertical separator line is drawn between columns.
    pub separator: bool,
    /// Per-column widths in twips (overrides uniform spacing when non-empty).
    #[serde(default)]
    pub column_widths_twips: Vec<u32>,
}

/// Paragraph border (four sides plus between-paragraph rule).
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
pub struct ParagraphBorder {
    /// Top border of the paragraph.
    pub top: Option<BorderLine>,
    /// Bottom border of the paragraph.
    pub bottom: Option<BorderLine>,
    /// Left border of the paragraph.
    pub left: Option<BorderLine>,
    /// Right border of the paragraph.
    pub right: Option<BorderLine>,
    /// Border drawn between consecutive bordered paragraphs.
    pub between: Option<BorderLine>,
}

/// Per-edge cell padding (all values in twips).
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
pub struct CellPadding {
    /// Top cell padding in twips.
    pub top_twips: Option<u32>,
    /// Bottom cell padding in twips.
    pub bottom_twips: Option<u32>,
    /// Left cell padding in twips.
    pub left_twips: Option<u32>,
    /// Right cell padding in twips.
    pub right_twips: Option<u32>,
}

/// Positioning data for a floating (non-inline) image.
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
pub struct FloatingImage {
    /// Horizontal offset from the anchor in EMUs.
    pub x_emu: i64,
    /// Vertical offset from the anchor in EMUs.
    pub y_emu: i64,
    /// Display width in EMUs.
    pub width_emu: u64,
    /// Display height in EMUs.
    pub height_emu: u64,
    /// Horizontal anchor reference frame.
    pub h_anchor: FloatAnchor,
    /// Vertical anchor reference frame.
    pub v_anchor: FloatAnchor,
    /// Text wrap mode around the image.
    pub text_wrap: TextWrap,
    /// Whether the image may overlap other floating objects.
    #[serde(default)]
    pub allow_overlap: bool,
}

/// A header or footer containing block elements.
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, Default, serde::Serialize, serde::Deserialize)]
pub struct HeaderFooter {
    /// Block elements that make up the header or footer.
    pub content: Vec<Element>,
}

/// A floating text box containing block elements.
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, Default, serde::Serialize, serde::Deserialize)]
pub struct TextBox {
    /// Block elements inside the text box.
    pub content: Vec<Element>,
    /// Width of the text box in EMUs.
    pub width_emu: Option<u64>,
    /// Height of the text box in EMUs.
    pub height_emu: Option<u64>,
    /// Horizontal position in EMUs from the anchor.
    pub x_emu: Option<i64>,
    /// Vertical position in EMUs from the anchor.
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
    /// Numeric identifier matching the inline reference mark.
    pub id: u32,
    /// Block elements comprising the note body.
    pub content: Vec<Element>,
    /// Optional custom marker text (when absent the auto-number is used).
    pub marker: Option<String>,
}

/// An inline reference mark pointing to a footnote or endnote.
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, Default, serde::Serialize, serde::Deserialize)]
pub struct FootnoteRef {
    /// Numeric identifier of the referenced note.
    pub note_id: u32,
    /// Optional custom marker text (when absent the auto-number is used).
    pub marker: Option<String>,
}

/// A preformatted code block with an optional language tag.
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq, Default, serde::Serialize, serde::Deserialize)]
pub struct CodeBlock {
    /// Optional language identifier for syntax highlighting.
    pub language: Option<String>,
    /// The preformatted code text.
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
    /// Default header for this section.
    pub header: Option<HeaderFooter>,
    /// Default footer for this section.
    pub footer: Option<HeaderFooter>,
    /// Header used on the first page of this section.
    pub first_page_header: Option<HeaderFooter>,
    /// Footer used on the first page of this section.
    pub first_page_footer: Option<HeaderFooter>,
    /// Header used on even-numbered pages of this section.
    pub even_page_header: Option<HeaderFooter>,
    /// Footer used on even-numbered pages of this section.
    pub even_page_footer: Option<HeaderFooter>,
    /// Solid background colour for this section (RGB).
    /// PPTX: parsed from `<p:cSld><p:bg><p:bgPr><a:solidFill>` on the slide.
    /// Image / gradient backgrounds are intentionally skipped — only the
    /// solid case round-trips through this minimal field.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub background_rgb: Option<[u8; 3]>,
}

/// A block-level content element.
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
#[serde(tag = "type", rename_all = "snake_case")]
#[non_exhaustive]
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
    /// A vector shape (line / rectangle) anchored on the page. Used by
    /// the layout-preserving DOCX path to round-trip rules and dividers.
    Shape(Shape),
}

/// A vector shape anchored at absolute page coordinates.
#[allow(dead_code)]
#[derive(Debug, Clone, Default, PartialEq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "snake_case")]
pub struct Shape {
    /// Geometry kind.
    pub kind: ShapeGeom,
    /// X offset from the anchor in EMUs.
    pub x_emu: i64,
    /// Y offset from the anchor in EMUs.
    pub y_emu: i64,
    /// Width in EMUs.
    pub width_emu: u64,
    /// Height in EMUs.
    pub height_emu: u64,
    /// Horizontal anchor reference frame.
    #[serde(default)]
    pub h_anchor: FloatAnchor,
    /// Vertical anchor reference frame.
    #[serde(default)]
    pub v_anchor: FloatAnchor,
    /// Stroke colour as RGB (0..255).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub stroke_rgb: Option<[u8; 3]>,
    /// Fill colour as RGB (0..255).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_rgb: Option<[u8; 3]>,
    /// Stroke width in EMUs.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub stroke_w_emu: Option<i64>,
}

/// Vector-shape geometry kinds we currently round-trip.
#[allow(dead_code)]
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum ShapeGeom {
    /// Straight line from `(x, y)` to `(x + width, y + height)`.
    #[default]
    Line,
    /// Axis-aligned rectangle.
    Rect,
}

/// A heading element with a nesting level.
#[derive(Debug, Clone, PartialEq, Default, serde::Serialize, serde::Deserialize)]
pub struct Heading {
    /// Heading level 1–6 (1 = largest).
    #[serde(default = "default_heading_level")]
    pub level: u8,
    /// Inline content of the heading.
    #[serde(default)]
    pub content: Vec<InlineContent>,
    /// Absolute frame position for layout-preserving DOCX
    /// (mirrors `Paragraph::frame_position`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub frame_position: Option<FramePosition>,
    /// Horizontal alignment (mirrors `Paragraph::alignment`). PDF
    /// title pages often centre their headings; without this the
    /// round-trip flattens them to left-aligned.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alignment: Option<ParagraphAlignment>,
}

fn default_heading_level() -> u8 {
    1
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
    /// Keep this paragraph on the same page as the next paragraph.
    #[serde(default)]
    pub keep_with_next: bool,
    /// Prevent a page break within this paragraph.
    #[serde(default)]
    pub keep_together: bool,
    /// Force a page break before this paragraph.
    #[serde(default)]
    pub page_break_before: bool,
    /// Outline level (0 = body text, 1–9 = heading levels).
    pub outline_level: Option<u8>,
    /// Absolute frame position (from `<w:framePr>`). Present when the
    /// DOCX uses page-anchored frames for layout-preserving content
    /// (see pdf_oxide's `to_docx_bytes_layout`). Twips relative to the
    /// page origin (top-left).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub frame_position: Option<FramePosition>,
}

/// Absolute frame position for a paragraph anchored to the page.
/// Mirrors the OOXML `<w:framePr>` attribute set we care about for
/// reproducing visual layout in downstream renderers.
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
pub struct FramePosition {
    /// X position in twips, anchored to the page origin (top-left).
    pub x_twips: i32,
    /// Y position in twips, anchored to the page origin (top-left).
    pub y_twips: i32,
    /// Frame width in twips.
    pub width_twips: i32,
    /// Frame height in twips.
    pub height_twips: i32,
}

/// Inline content within a paragraph or heading.
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
#[serde(tag = "type", rename_all = "snake_case")]
#[non_exhaustive]
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

/// Pick the dominant font size (in points) for a paragraph's worth of
/// inline content. Returns the *first* declared `font_size_half_pt`,
/// converted from half-points to points (e.g. 18 half-pt → 9 pt).
///
/// Used by both renderers and writers when one paragraph-level size is
/// needed: the IR groups runs into a paragraph by line clustering, so
/// the size on the first span is representative of the body text.
/// Mixed-size paragraphs (drop-caps, math marks mid-line) lose the
/// variation — that's the deliberate trade-off.
pub fn first_inline_font_size_pt(content: &[InlineContent]) -> Option<f32> {
    for ic in content {
        if let InlineContent::Text(span) = ic {
            if let Some(half_pt) = span.font_size_half_pt {
                return Some(half_pt as f32 / 2.0);
            }
        }
    }
    None
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
    /// Whether all characters are rendered as uppercase.
    #[serde(default)]
    pub all_caps: bool,
    /// Whether lowercase letters are rendered as smaller capitals.
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

/// Wrap a non-empty inline-content vector into a single-Paragraph
/// block, or return an empty Vec if the inline content is empty.
/// Used by list builders to turn each item's inline run into its
/// `Vec<Element>` content slot.
pub fn inline_to_element_block(content: Vec<InlineContent>) -> Vec<Element> {
    if content.is_empty() {
        Vec::new()
    } else {
        vec![Element::Paragraph(Paragraph {
            content,
            ..Default::default()
        })]
    }
}

/// Build a nested `List` from a flat `(level, inline)` sequence.
///
/// Items whose level matches `base_level` (or is shallower) become
/// `ListItem`s at the current depth. Items whose level is *deeper*
/// than `base_level` are recursively grouped into the most recent
/// item's `nested` sub-list. Levels are 0-indexed.
///
/// Used by both `convert_docx` and `convert_pptx` to translate flat
/// `<w:numPr w:ilvl=…>` / `<a:p lvl=…>` paragraph streams into the
/// IR's tree-shaped `List`.
pub fn build_nested_list(
    ordered: bool,
    items: &[(u8, Vec<InlineContent>)],
    base_level: u8,
) -> List {
    let mut list_items = Vec::new();
    let mut idx = 0;

    while idx < items.len() {
        let (level, content) = &items[idx];
        let nested_start = idx + 1;
        let mut nested_end = nested_start;
        while nested_end < items.len() && items[nested_end].0 > base_level {
            nested_end += 1;
        }
        let nested = if *level <= base_level && nested_end > nested_start {
            Some(build_nested_list(ordered, &items[nested_start..nested_end], base_level + 1))
        } else {
            None
        };
        list_items.push(ListItem {
            content: inline_to_element_block(content.clone()),
            nested,
        });
        idx = if nested_end > nested_start {
            nested_end
        } else {
            idx + 1
        };
    }

    List {
        ordered,
        items: list_items,
        ..Default::default()
    }
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

#[cfg(test)]
mod tests {
    use super::*;

    // ── first_inline_font_size_pt ────────────────────────────────────

    #[test]
    fn first_font_size_returns_half_pt_as_pt() {
        let content = vec![InlineContent::Text(TextSpan {
            text: "hi".into(),
            font_size_half_pt: Some(24), // 12pt
            ..Default::default()
        })];
        assert_eq!(first_inline_font_size_pt(&content), Some(12.0));
    }

    #[test]
    fn first_font_size_picks_first_declared() {
        // Second span's size is ignored — the first declared one wins.
        let content = vec![
            InlineContent::Text(TextSpan {
                text: "a".into(),
                font_size_half_pt: Some(20), // 10pt
                ..Default::default()
            }),
            InlineContent::Text(TextSpan {
                text: "b".into(),
                font_size_half_pt: Some(48), // 24pt — ignored
                ..Default::default()
            }),
        ];
        assert_eq!(first_inline_font_size_pt(&content), Some(10.0));
    }

    #[test]
    fn first_font_size_skips_unsized_runs() {
        // First run has no size; second does → returns the second's size.
        let content = vec![
            InlineContent::Text(TextSpan {
                text: "a".into(),
                ..Default::default()
            }),
            InlineContent::Text(TextSpan {
                text: "b".into(),
                font_size_half_pt: Some(16), // 8pt
                ..Default::default()
            }),
        ];
        assert_eq!(first_inline_font_size_pt(&content), Some(8.0));
    }

    #[test]
    fn first_font_size_empty_returns_none() {
        assert_eq!(first_inline_font_size_pt(&[]), None);
    }

    #[test]
    fn first_font_size_all_unsized_returns_none() {
        let content = vec![
            InlineContent::Text(TextSpan::plain("a")),
            InlineContent::Text(TextSpan::plain("b")),
        ];
        assert_eq!(first_inline_font_size_pt(&content), None);
    }

    // ── inline_to_element_block ──────────────────────────────────────

    #[test]
    fn inline_to_element_block_empty_returns_empty() {
        let result = inline_to_element_block(vec![]);
        assert!(result.is_empty());
    }

    #[test]
    fn inline_to_element_block_wraps_in_paragraph() {
        let inline = vec![InlineContent::Text(TextSpan::plain("hello"))];
        let result = inline_to_element_block(inline);
        assert_eq!(result.len(), 1);
        match &result[0] {
            Element::Paragraph(p) => {
                assert_eq!(p.content.len(), 1);
                assert!(matches!(
                    &p.content[0],
                    InlineContent::Text(s) if s.text == "hello"
                ));
            },
            _ => panic!("expected Paragraph"),
        }
    }

    // ── build_nested_list ────────────────────────────────────────────

    fn item(level: u8, text: &str) -> (u8, Vec<InlineContent>) {
        (level, vec![InlineContent::Text(TextSpan::plain(text))])
    }

    fn list_item_text(item: &ListItem) -> String {
        let mut out = String::new();
        for el in &item.content {
            if let Element::Paragraph(p) = el {
                for c in &p.content {
                    if let InlineContent::Text(s) = c {
                        out.push_str(&s.text);
                    }
                }
            }
        }
        out
    }

    #[test]
    fn build_nested_list_flat() {
        let items = vec![item(0, "A"), item(0, "B"), item(0, "C")];
        let list = build_nested_list(false, &items, 0);
        assert!(!list.ordered);
        assert_eq!(list.items.len(), 3);
        assert!(list.items.iter().all(|li| li.nested.is_none()));
        assert_eq!(list_item_text(&list.items[1]), "B");
    }

    #[test]
    fn build_nested_list_two_levels() {
        // Top:   A
        //   sub: A.1, A.2
        // Top:   B
        let items = vec![item(0, "A"), item(1, "A.1"), item(1, "A.2"), item(0, "B")];
        let list = build_nested_list(true, &items, 0);
        assert!(list.ordered);
        assert_eq!(list.items.len(), 2);
        let nested = list.items[0].nested.as_ref().expect("A has nested");
        assert_eq!(nested.items.len(), 2);
        assert_eq!(list_item_text(&nested.items[0]), "A.1");
        assert_eq!(list_item_text(&nested.items[1]), "A.2");
        // B has no nested children.
        assert!(list.items[1].nested.is_none());
    }

    #[test]
    fn build_nested_list_three_levels() {
        let items = vec![item(0, "A"), item(1, "A.1"), item(2, "A.1.x"), item(0, "B")];
        let list = build_nested_list(false, &items, 0);
        let l1 = list.items[0].nested.as_ref().unwrap();
        assert_eq!(l1.items.len(), 1);
        let l2 = l1.items[0].nested.as_ref().unwrap();
        assert_eq!(l2.items.len(), 1);
        assert_eq!(list_item_text(&l2.items[0]), "A.1.x");
    }

    #[test]
    fn build_nested_list_empty() {
        let list = build_nested_list(false, &[], 0);
        assert!(list.items.is_empty());
    }

    // ── TextSpan::plain ──────────────────────────────────────────────

    #[test]
    fn text_span_plain_has_default_styling() {
        let s = TextSpan::plain("hi");
        assert_eq!(s.text, "hi");
        assert!(!s.bold);
        assert!(!s.italic);
        assert!(s.font_size_half_pt.is_none());
        assert!(s.hyperlink.is_none());
    }

    // ── FramePosition / Shape defaults ───────────────────────────────

    #[test]
    fn shape_default_is_line_at_origin() {
        let s = Shape::default();
        assert!(matches!(s.kind, ShapeGeom::Line));
        assert_eq!(s.x_emu, 0);
        assert_eq!(s.width_emu, 0);
        assert!(s.stroke_rgb.is_none());
    }

    #[test]
    fn frame_position_round_trips_via_serde() {
        let fp = FramePosition {
            x_twips: 720,
            y_twips: 1080,
            width_twips: 5000,
            height_twips: 400,
        };
        let json = serde_json::to_string(&fp).unwrap();
        let back: FramePosition = serde_json::from_str(&json).unwrap();
        assert_eq!(fp, back);
    }
}
