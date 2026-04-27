use super::paragraph::Run;

/// A hyperlink element within a paragraph (`w:hyperlink`).
#[derive(Debug, Clone)]
pub struct Hyperlink {
    /// The link destination.
    pub target: HyperlinkTarget,
    /// Optional screen-tip tooltip text.
    pub tooltip: Option<String>,
    /// Text runs that form the visible link text.
    pub runs: Vec<Run>,
}

/// The destination of a hyperlink.
#[derive(Debug, Clone)]
pub enum HyperlinkTarget {
    /// External URL, resolved from relationship with TargetMode=External.
    External(String),
    /// Internal bookmark name (from `w:anchor` attribute).
    Internal(String),
}
