//! Choices about *which* text a reader returns, as opposed to the
//! limits (the `limits` module) that bound how much.
//!
//! By default a document is read as Word shows it with revisions accepted:
//! text a run hides (`w:vanish`, `sprmCFVanish`) and text a tracked change
//! deleted (`w:del`, `sprmCFRMarkDel`) are left out. Search and archival
//! callers want that text too; both settings bring it back, for the `.docx`
//! and the `.doc` readers alike, and for every output built from them.
//!
//! The settings are process-wide, like the limits: set them before reading,
//! not while documents are being read on other threads.

use std::sync::atomic::{AtomicBool, Ordering};

static INCLUDE_HIDDEN_TEXT: AtomicBool = AtomicBool::new(false);
static INCLUDE_DELETED_TEXT: AtomicBool = AtomicBool::new(false);

/// Whether text formatted as hidden is part of what readers return.
/// `false` by default.
pub fn include_hidden_text() -> bool {
    INCLUDE_HIDDEN_TEXT.load(Ordering::Relaxed)
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
    INCLUDE_DELETED_TEXT.load(Ordering::Relaxed)
}

/// Keep (`true`) or leave out (`false`, the default) text deleted by a
/// tracked change in every document read afterwards.
pub fn set_include_deleted_text(include: bool) {
    INCLUDE_DELETED_TEXT.store(include, Ordering::Relaxed);
}
