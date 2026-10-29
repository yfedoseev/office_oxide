//! Choices about *which* text a reader returns, as opposed to the
//! limits (the `limits` module) that bound how much.
//!
//! By default a document is read as Word shows it with revisions accepted:
//! text a run hides (`w:vanish`, `sprmCFVanish`) and text a tracked change
//! deleted (`w:del`, `sprmCFRMarkDel`) are left out. Search and archival
//! callers want that text too; both settings bring it back, for the `.docx`
//! and the `.doc` readers alike, and for every output built from them.
//!
//! These functions set the process-wide defaults: set them before reading, not
//! while documents are being read on other threads. To choose per document,
//! use [`ParseOptions`](crate::options::ParseOptions) with
//! [`Document::open_with`](crate::Document::open_with).

use std::sync::atomic::{AtomicBool, Ordering};

static INCLUDE_HIDDEN_TEXT: AtomicBool = AtomicBool::new(false);
static INCLUDE_DELETED_TEXT: AtomicBool = AtomicBool::new(false);

/// Whether text formatted as hidden is part of what readers return.
/// `false` by default.
pub fn include_hidden_text() -> bool {
    scoped().map_or_else(|| INCLUDE_HIDDEN_TEXT.load(Ordering::Relaxed), |(hidden, _)| hidden)
}

/// Keep (`true`) or leave out (`false`, the default) hidden text in every
/// document read afterwards.
pub fn set_include_hidden_text(include: bool) {
    INCLUDE_HIDDEN_TEXT.store(include, Ordering::Relaxed);
}

/// Whether text a tracked change deleted is part of what readers return.
/// `false` by default.
///
/// In a `.docx` this covers `w:del` runs and `w:cellDel` cells. Text under
/// `w:moveFrom` stays out: the move's destination (`w:moveTo`) already
/// carries it, and keeping both would say it twice.
pub fn include_deleted_text() -> bool {
    scoped().map_or_else(|| INCLUDE_DELETED_TEXT.load(Ordering::Relaxed), |(_, deleted)| deleted)
}

/// Keep (`true`) or leave out (`false`, the default) text deleted by a
/// tracked change in every document read afterwards.
pub fn set_include_deleted_text(include: bool) {
    INCLUDE_DELETED_TEXT.store(include, Ordering::Relaxed);
}

thread_local! {
    /// `(hidden, deleted)` of the call running on this thread, when it has
    /// its own (see [`ReadScope`]).
    static SCOPED: std::cell::Cell<Option<(bool, bool)>> = const { std::cell::Cell::new(None) };
}

/// The settings of the call running on this thread, if it has its own, for
/// work that continues on another thread to carry them over.
pub(crate) fn scope_snapshot() -> Option<(bool, bool)> {
    scoped()
}

fn scoped() -> Option<(bool, bool)> {
    SCOPED.try_with(|c| c.get()).ok().flatten()
}

/// While alive, the two settings on this thread are the call's own instead of
/// the process-wide ones. Restores what it replaced when dropped.
#[must_use = "the settings apply only while the scope is alive; bind it for the length of the call"]
pub(crate) struct ReadScope(Option<(bool, bool)>);

impl ReadScope {
    pub(crate) fn new(own: Option<(bool, bool)>) -> Self {
        let before = SCOPED.with(|c| c.get());
        if own.is_some() {
            SCOPED.with(|c| c.set(own));
        }
        Self(before)
    }
}

impl Drop for ReadScope {
    fn drop(&mut self) {
        let _ = SCOPED.try_with(|c| c.set(self.0));
    }
}
