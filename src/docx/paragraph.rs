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
    /// unread dropped whole documents' worth of prose (90% of
    /// the reporter's text lived here).
    TextBox(Vec<super::document::BlockElement>),
    /// A `w:footnoteReference` mark: the citation point in the body text.
    /// Carries the referenced note's `w:id`. The note *body* was already
    /// read from `footnotes.xml`; without this the IR could not say where
    /// it was cited. The second field mirrors
    /// `w:customMarkFollows`: `true` means the note body supplies its own
    /// mark glyph as a leading run instead of Word's auto-number.
    FootnoteRef(u32, bool),
    /// A `w:endnoteReference` mark. See [`RunContent::FootnoteRef`].
    EndnoteRef(u32, bool),
    /// A `w:commentReference` mark. See [`RunContent::FootnoteRef`].
    CommentRef(u32),
    /// Legacy form-field state parsed out of `<w:fldChar><w:ffData>`.
    /// A checkbox's checked state exists nowhere else in the document, so
    /// skipping `w:ffData` lost it unrecoverably.
    FormField(FormField),
    /// A reference to a separate OPC part whose content is folded in once
    /// the package is readable: a SmartArt `word/diagrams/dataN.xml`
    /// or an embedded OOXML package from `<o:OLEObject>`
    ///. Replaced with [`RunContent::TextBox`] during
    /// `DocxDocument::from_opc`; an unresolvable reference is dropped.
    DeferredPart(String),
}

/// Legacy form-field state (`<w:fldChar><w:ffData>`), the data behind a
/// `FORMCHECKBOX` / `FORMDROPDOWN` / `FORMTEXT` field.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct FormField {
    /// `<w:name w:val="…"/>`, the field's bookmark name.
    pub name: Option<String>,
    /// Kind-specific state.
    pub kind: FormFieldKind,
    /// Text to surface in extraction, set only when the field carries no
    /// cached result runs of its own. A `FORMCHECKBOX` never has any, so
    /// its glyph always comes from here; a dropdown with a cached display
    /// run keeps `None` so the value is not emitted twice.
    pub display_text: Option<String>,
}

/// The three legacy form-field kinds `w:ffData` can describe.
#[derive(Debug, Clone, Default, PartialEq)]
pub enum FormFieldKind {
    /// `<w:checkBox>`: the checked state (`w:checked`, else `w:default`).
    CheckBox {
        /// Whether the box is ticked.
        checked: bool,
    },
    /// `<w:ddList>`: the full option list and the selected index.
    DropDown {
        /// Every `<w:listEntry w:val="…"/>`, in order.
        entries: Vec<String>,
        /// Zero-based `<w:result w:val="N"/>` (defaults to 0).
        selected: usize,
    },
    /// `<w:textInput>`: the default value, when one is given.
    TextInput {
        /// `<w:default w:val="…"/>`.
        default: Option<String>,
    },
    /// `w:ffData` present but of no recognised kind.
    #[default]
    Unknown,
}

impl FormField {
    /// The value a renderer should show for this field: a checkbox glyph,
    /// the selected dropdown entry, or the text field's default.
    pub fn value_text(&self) -> Option<String> {
        match &self.kind {
            FormFieldKind::CheckBox { checked } => {
                Some(if *checked { "\u{2612}" } else { "\u{2610}" }.to_string())
            },
            FormFieldKind::DropDown { entries, selected } => entries
                .get(*selected)
                .or_else(|| entries.first())
                .filter(|s| !s.is_empty())
                .cloned(),
            FormFieldKind::TextInput { default } => default.clone().filter(|s| !s.is_empty()),
            FormFieldKind::Unknown => None,
        }
    }
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
