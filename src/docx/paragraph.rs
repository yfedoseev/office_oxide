use super::formatting::{ParagraphProperties, RunProperties};
use super::hyperlink::Hyperlink;
use super::image::DrawingInfo;

/// A single paragraph (`w:p`).
#[derive(Debug, Clone, Default)]
pub struct Paragraph {
    /// Formatting applied to this paragraph.
    pub properties: Option<ParagraphProperties>,
    /// Runs and hyperlinks in order.
    pub content: Vec<ParagraphContent>,
}

/// Content that can appear directly inside a paragraph.
#[derive(Debug, Clone)]
pub enum ParagraphContent {
    /// A plain text run (`w:r`).
    Run(Run),
    /// A hyperlink element (`w:hyperlink`).
    Hyperlink(Hyperlink),
}

/// A run of text with uniform formatting (`w:r`).
#[derive(Debug, Clone, Default)]
pub struct Run {
    /// Run-level formatting.
    pub properties: Option<RunProperties>,
    /// Text, breaks, tabs, and drawings in order.
    pub content: Vec<RunContent>,
}

/// Content within a run.
#[derive(Debug, Clone)]
pub enum RunContent {
    /// A `w:t` text node.
    Text(String),
    /// A `w:br` break element.
    Break(BreakType),
    /// A `w:tab` tab character.
    Tab,
    /// A `w:drawing` inline or anchored image.
    Drawing(DrawingInfo),
    /// A `w:txbxContent` text-box body found inside a `w:pict` /
    /// `w:drawing` / `mc:AlternateContent` shape. Text boxes are ordinary
    /// block content that happens to be drawn in a frame; leaving them
    /// unread dropped whole documents' worth of prose (issue #102: 90% of
    /// the reporter's text lived here).
    TextBox(Vec<super::document::BlockElement>),
}

/// Types of breaks.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BreakType {
    /// Line break within a paragraph.
    Line,
    /// Hard page break.
    Page,
    /// Column break.
    Column,
}
