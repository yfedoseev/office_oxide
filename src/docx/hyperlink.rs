use super::paragraph::Run;

/// A hyperlink element within a paragraph (`w:hyperlink`).
#[derive(Debug, Clone)]
pub struct Hyperlink {
    pub target: HyperlinkTarget,
    pub tooltip: Option<String>,
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
