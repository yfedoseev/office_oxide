use super::paragraph::Run;

/// A hyperlink element within a paragraph (`w:hyperlink`).
#[derive(Debug, Clone)]
pub struct Hyperlink {
    /// The link destination.
    pub target: HyperlinkTarget,
    /// URI fragment from `w:anchor` when the element *also* carries an
    /// `r:id`. Per ECMA-376 the anchor is then a fragment appended to the
    /// relationship's target (`externalURL#fragment`), not a
    /// same-document bookmark — reading only the anchor dropped the real
    /// URL and left a dead fragment behind (issues #242, #292).
    pub fragment: Option<String>,
    /// Optional screen-tip tooltip text.
    pub tooltip: Option<String>,
    /// Text runs that form the visible link text.
    pub runs: Vec<Run>,
}

/// The destination of a hyperlink.
#[derive(Debug, Clone, PartialEq)]
pub enum HyperlinkTarget {
    /// External URL, resolved from relationship with TargetMode=External.
    External(String),
    /// Internal bookmark name (from `w:anchor` attribute).
    Internal(String),
}
