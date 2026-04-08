use super::formatting::{ParagraphProperties, RunProperties};
use super::hyperlink::Hyperlink;
use super::image::DrawingInfo;

/// A single paragraph (`w:p`).
#[derive(Debug, Clone, Default)]
pub struct Paragraph {
    pub properties: Option<ParagraphProperties>,
    pub content: Vec<ParagraphContent>,
}

/// Content that can appear directly inside a paragraph.
#[derive(Debug, Clone)]
pub enum ParagraphContent {
    Run(Run),
    Hyperlink(Hyperlink),
}

/// A run of text with uniform formatting (`w:r`).
#[derive(Debug, Clone, Default)]
pub struct Run {
    pub properties: Option<RunProperties>,
    pub content: Vec<RunContent>,
}

/// Content within a run.
#[derive(Debug, Clone)]
pub enum RunContent {
    Text(String),
    Break(BreakType),
    Tab,
    Drawing(DrawingInfo),
}

/// Types of breaks.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BreakType {
    Line,
    Page,
    Column,
}
