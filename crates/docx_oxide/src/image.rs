use office_core::units::Emu;

/// Information about a drawing/image reference within a run.
#[derive(Debug, Clone)]
pub struct DrawingInfo {
    pub relationship_id: String,
    pub description: Option<String>,
    pub width: Emu,
    pub height: Emu,
    /// `true` = inline, `false` = anchor (floating).
    pub inline: bool,
}
