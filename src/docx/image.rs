use crate::core::units::Emu;

/// Information about a drawing/image reference within a run.
#[derive(Debug, Clone)]
pub struct DrawingInfo {
    /// Relationship ID pointing to the image part.
    pub relationship_id: String,
    /// Alt-text description from `wp:docPr/@descr`.
    pub description: Option<String>,
    /// Image width in EMUs.
    pub width: Emu,
    /// Image height in EMUs.
    pub height: Emu,
    /// `true` = inline, `false` = anchor (floating).
    pub inline: bool,
}
