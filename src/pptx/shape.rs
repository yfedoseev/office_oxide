/// A shape in a PPTX slide.
#[derive(Debug, Clone)]
pub enum Shape {
    AutoShape(AutoShape),
    Picture(PictureShape),
    Group(GroupShape),
    GraphicFrame(GraphicFrame),
    Connector(ConnectorShape),
}

#[derive(Debug, Clone)]
pub struct AutoShape {
    pub id: u32,
    pub name: String,
    pub alt_text: Option<String>,
    pub position: Option<ShapePosition>,
    pub text_body: Option<TextBody>,
    pub placeholder: Option<PlaceholderInfo>,
}

#[derive(Debug, Clone)]
pub struct PictureShape {
    pub id: u32,
    pub name: String,
    pub alt_text: Option<String>,
    pub position: Option<ShapePosition>,
}

#[derive(Debug, Clone)]
pub struct GroupShape {
    pub id: u32,
    pub name: String,
    pub position: Option<ShapePosition>,
    pub children: Vec<Shape>,
}

#[derive(Debug, Clone)]
pub struct GraphicFrame {
    pub id: u32,
    pub name: String,
    pub position: Option<ShapePosition>,
    pub content: GraphicContent,
}

#[derive(Debug, Clone)]
pub enum GraphicContent {
    Table(Table),
    Unknown,
}

#[derive(Debug, Clone)]
pub struct ConnectorShape {
    pub id: u32,
    pub name: String,
    pub position: Option<ShapePosition>,
}

#[derive(Debug, Clone)]
pub struct ShapePosition {
    pub x: i64,
    pub y: i64,
    pub cx: i64,
    pub cy: i64,
}

#[derive(Debug, Clone)]
pub struct PlaceholderInfo {
    pub ph_type: Option<String>,
    pub idx: Option<u32>,
}

// ---------------------------------------------------------------------------
// Text body types
// ---------------------------------------------------------------------------

#[derive(Debug, Clone)]
pub struct TextBody {
    pub paragraphs: Vec<TextParagraph>,
}

#[derive(Debug, Clone)]
pub struct TextParagraph {
    pub level: u32,
    pub content: Vec<TextContent>,
}

#[derive(Debug, Clone)]
pub enum TextContent {
    Run(TextRun),
    LineBreak,
    Field(TextField),
}

#[derive(Debug, Clone)]
pub struct TextRun {
    pub text: String,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub strikethrough: bool,
    pub hyperlink: Option<HyperlinkInfo>,
}

#[derive(Debug, Clone)]
pub struct TextField {
    pub field_type: Option<String>,
    pub text: String,
}

#[derive(Debug, Clone)]
pub struct HyperlinkInfo {
    pub target: HyperlinkTarget,
    pub tooltip: Option<String>,
}

#[derive(Debug, Clone)]
pub enum HyperlinkTarget {
    External(String),
    Internal(String),
}

// ---------------------------------------------------------------------------
// Table types
// ---------------------------------------------------------------------------

#[derive(Debug, Clone)]
pub struct Table {
    pub rows: Vec<TableRow>,
}

#[derive(Debug, Clone)]
pub struct TableRow {
    pub cells: Vec<TableCell>,
}

#[derive(Debug, Clone)]
pub struct TableCell {
    pub text_body: Option<TextBody>,
    pub grid_span: u32,
    pub row_span: u32,
    pub h_merge: bool,
    pub v_merge: bool,
}
