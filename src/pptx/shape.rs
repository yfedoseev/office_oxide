/// A shape in a PPTX slide.
#[derive(Debug, Clone)]
pub enum Shape {
    /// Standard auto-shape (text box, rectangle, etc.).
    AutoShape(AutoShape),
    /// An image or picture shape.
    Picture(PictureShape),
    /// A group of child shapes.
    Group(GroupShape),
    /// A graphic frame (contains a table or chart).
    GraphicFrame(GraphicFrame),
    /// A connector line between shapes.
    Connector(ConnectorShape),
}

/// A standard auto-shape (text boxes, rectangles, callouts, etc.).
#[derive(Debug, Clone)]
pub struct AutoShape {
    /// Shape numeric identifier.
    pub id: u32,
    /// Shape name from `cNvPr`.
    pub name: String,
    /// Accessibility description text.
    pub alt_text: Option<String>,
    /// Bounding box position and size in EMU.
    pub position: Option<ShapePosition>,
    /// Text body content, if any.
    pub text_body: Option<TextBody>,
    /// Placeholder role, if this shape is a slide placeholder.
    pub placeholder: Option<PlaceholderInfo>,
}

/// An image or picture shape (`<p:pic>`).
#[derive(Debug, Clone)]
pub struct PictureShape {
    /// Shape numeric identifier.
    pub id: u32,
    /// Shape name.
    pub name: String,
    /// Accessibility description text.
    pub alt_text: Option<String>,
    /// Bounding box position and size in EMU.
    pub position: Option<ShapePosition>,
    /// Relationship ID (`r:embed`) of the underlying media part, if any.
    pub embed_rid: Option<String>,
    /// Raw image bytes resolved via `embed_rid`, if the slide carried a
    /// resolvable IMAGE relationship at parse time.
    pub data: Option<Vec<u8>>,
    /// Image format inferred from the relationship target extension or
    /// byte signature (e.g. `"png"`, `"jpeg"`, `"gif"`, `"emf"`).
    pub format: Option<String>,
}

/// A group of child shapes (`<p:grpSp>`).
#[derive(Debug, Clone)]
pub struct GroupShape {
    /// Shape numeric identifier.
    pub id: u32,
    /// Shape name.
    pub name: String,
    /// Bounding box position and size in EMU.
    pub position: Option<ShapePosition>,
    /// Member shapes of this group.
    pub children: Vec<Shape>,
}

/// A graphic frame that wraps a table, chart, or other graphic (`<p:graphicFrame>`).
#[derive(Debug, Clone)]
pub struct GraphicFrame {
    /// Shape numeric identifier.
    pub id: u32,
    /// Shape name.
    pub name: String,
    /// Bounding box position and size in EMU.
    pub position: Option<ShapePosition>,
    /// The graphic content inside this frame.
    pub content: GraphicContent,
}

/// Content held by a `GraphicFrame`.
#[derive(Debug, Clone)]
pub enum GraphicContent {
    /// A DrawingML table.
    Table(Table),
    /// Unsupported or unrecognised graphic type.
    Unknown,
}

/// A connector line (`<p:cxnSp>`).
#[derive(Debug, Clone)]
pub struct ConnectorShape {
    /// Shape numeric identifier.
    pub id: u32,
    /// Shape name.
    pub name: String,
    /// Bounding box position and size in EMU.
    pub position: Option<ShapePosition>,
}

/// Bounding box for a shape, all values in EMU (English Metric Units).
#[derive(Debug, Clone)]
pub struct ShapePosition {
    /// Left edge offset from slide origin.
    pub x: i64,
    /// Top edge offset from slide origin.
    pub y: i64,
    /// Width of the bounding box.
    pub cx: i64,
    /// Height of the bounding box.
    pub cy: i64,
}

/// Placeholder role information from `<p:ph>`.
#[derive(Debug, Clone)]
pub struct PlaceholderInfo {
    /// Placeholder type string, e.g. `"title"`, `"body"`, `"dt"`.
    pub ph_type: Option<String>,
    /// Placeholder index (for body/content placeholders).
    pub idx: Option<u32>,
}

// ---------------------------------------------------------------------------
// Text body types
// ---------------------------------------------------------------------------

/// The text body of a shape (`<p:txBody>`).
#[derive(Debug, Clone)]
pub struct TextBody {
    /// Ordered list of paragraphs in this text body.
    pub paragraphs: Vec<TextParagraph>,
}

/// A single paragraph within a text body (`<a:p>`).
#[derive(Debug, Clone, Default)]
pub struct TextParagraph {
    /// Outline level (0 = top level).
    pub level: u32,
    /// Paragraph alignment from `<a:pPr algn="…"/>`. None when the
    /// attribute is absent (renderer-default left alignment).
    pub alignment: Option<crate::ir::ParagraphAlignment>,
    /// Space before the paragraph, in 100ths of a point — read from
    /// `<a:pPr><a:spcBef><a:spcPts val="…"/></a:spcBef></a:pPr>`.
    pub space_before_hundredths_pt: Option<u32>,
    /// Inline content items in this paragraph.
    pub content: Vec<TextContent>,
}

/// An inline content item inside a paragraph.
#[derive(Debug, Clone)]
pub enum TextContent {
    /// A text run with optional formatting.
    Run(TextRun),
    /// An explicit line break (`<a:br>`).
    LineBreak,
    /// An auto-updated field (slide number, date, etc.).
    Field(TextField),
}

/// A text run with optional character formatting (`<a:r>`).
#[derive(Debug, Clone, Default)]
pub struct TextRun {
    /// The text content of this run.
    pub text: String,
    /// Bold toggle (`None` = inherit).
    pub bold: Option<bool>,
    /// Italic toggle (`None` = inherit).
    pub italic: Option<bool>,
    /// Whether strikethrough is applied.
    pub strikethrough: bool,
    /// Hyperlink attached to this run, if any.
    pub hyperlink: Option<HyperlinkInfo>,
    /// Font size in hundredths of a point (`<a:rPr sz="1800"/>` → `Some(1800)` = 18 pt).
    /// `None` when the run inherits its size from the placeholder/master.
    pub font_size_hundredths_pt: Option<u32>,
    /// Explicit run colour from `<a:rPr><a:solidFill><a:srgbClr val="…"/></a:solidFill></a:rPr>`.
    /// `None` when the run inherits its colour from the placeholder /
    /// theme, or when the fill is non-sRGB (gradient, scheme colour).
    pub color_rgb: Option<[u8; 3]>,
}

/// An auto-updated field inside a paragraph (`<a:fld>`).
#[derive(Debug, Clone)]
pub struct TextField {
    /// Field type string, e.g. `"slidenum"`, `"datetime"`.
    pub field_type: Option<String>,
    /// Cached display text of the field.
    pub text: String,
}

/// Hyperlink associated with a text run.
#[derive(Debug, Clone)]
pub struct HyperlinkInfo {
    /// The hyperlink destination.
    pub target: HyperlinkTarget,
    /// Optional tooltip displayed on hover.
    pub tooltip: Option<String>,
}

/// Hyperlink destination type.
#[derive(Debug, Clone)]
pub enum HyperlinkTarget {
    /// External URL (resolved from the slide relationship).
    External(String),
    /// Internal slide/anchor reference.
    Internal(String),
}

// ---------------------------------------------------------------------------
// Table types
// ---------------------------------------------------------------------------

/// A DrawingML table inside a graphic frame.
#[derive(Debug, Clone)]
pub struct Table {
    /// Ordered rows of the table.
    pub rows: Vec<TableRow>,
}

/// A single row within a `Table`.
#[derive(Debug, Clone)]
pub struct TableRow {
    /// Cells in this row.
    pub cells: Vec<TableCell>,
}

/// A single cell within a `TableRow`.
#[derive(Debug, Clone)]
pub struct TableCell {
    /// Text content of the cell, if any.
    pub text_body: Option<TextBody>,
    /// Number of columns this cell spans horizontally.
    pub grid_span: u32,
    /// Number of rows this cell spans vertically.
    pub row_span: u32,
    /// True when this cell is a horizontal merge continuation.
    pub h_merge: bool,
    /// True when this cell is a vertical merge continuation.
    pub v_merge: bool,
}
