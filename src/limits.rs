//! Resource limits that bound what an untrusted document can make the
//! crate allocate.
//!
//! A spreadsheet can reference one shared string from every cell — the
//! string is stored once, the rendered text is the string times the cell
//! count. A 110 KB `.xlsx` holding a 1 MB string referenced from 12,000
//! cells rendered to 393 MB; the same string from a million rows would be
//! 32 GB from a 300 KB file. The spreadsheet paths — XLSX text rendering
//! (`plain_text`/CSV/Markdown/HTML over cells), the XLSX → IR converter,
//! and the legacy `.xls` reader (which materialises each referencing
//! cell's copy at open) — charge the characters they emit against one
//! per-document budget and stop, loudly, when it is spent. The word-
//! processing and presentation paths do not: their text is stored once
//! per occurrence, so its size is already bounded by the package's
//! decompressed size. The one exception is a paragraph style's name,
//! which is stored once in the style sheet and copied into every
//! paragraph that uses the style; `classify_style_name` bounds it. Apache
//! POI does the same with `ZipSecureFile.setMaxTextSize`.
//!
//! The default keeps every genuine workbook this crate has been tested
//! against (the largest, 181 MB of cell text, with room to spare) and
//! refuses the fan-out shapes. Processes that read larger workbooks raise
//! it once at startup with [the text-budget setter](crate::limits::set_max_text_chars).

//!
//! OOXML packages are zip archives, and the per-part decompression cap
//! ([`MAX_PART_SIZE`](crate::core::opc::MAX_PART_SIZE)) bounds one part
//! only. Two package-wide bounds sit on top of it, mirroring Apache POI's
//! `ZipSecureFile` limits: the number of entries an archive may hold
//! (set with [the entry-count setter](crate::limits::set_max_package_entries))
//! and the total decompressed bytes that may be read from one package (set
//! with [the byte-budget setter](crate::limits::set_max_package_bytes)). A
//! package that
//! exceeds either is refused with
//! [`Error::PackageLimit`](crate::core::Error::PackageLimit).

use std::sync::atomic::{AtomicU64, AtomicUsize, Ordering};

/// The default maximum number of entries in one OOXML package: 100,000.
///
/// A zip without ZIP64 extensions cannot hold more than 65,535; the
/// largest genuine packages this crate has been tested against (thousands
/// of sheets or slides, each with its own relationships part) stay well
/// below that.
pub const DEFAULT_MAX_PACKAGE_ENTRIES: usize = 100_000;

/// The default maximum total of decompressed bytes read from one OOXML
/// package: 2 GiB, four times the per-part cap.
pub const DEFAULT_MAX_PACKAGE_BYTES: u64 = 2 << 30;

static MAX_PACKAGE_ENTRIES: AtomicUsize = AtomicUsize::new(DEFAULT_MAX_PACKAGE_ENTRIES);
static MAX_PACKAGE_BYTES: AtomicU64 = AtomicU64::new(DEFAULT_MAX_PACKAGE_BYTES);

/// The maximum number of entries an OOXML package may hold.
pub fn max_package_entries() -> usize {
    MAX_PACKAGE_ENTRIES.load(Ordering::Relaxed)
}

/// Set the maximum number of entries for every package opened afterwards.
/// `usize::MAX` disables the bound.
pub fn set_max_package_entries(entries: usize) {
    MAX_PACKAGE_ENTRIES.store(entries, Ordering::Relaxed);
}

/// The maximum total of decompressed bytes read from one OOXML package.
/// Each distinct part is charged once, however often it is read.
pub fn max_package_bytes() -> u64 {
    MAX_PACKAGE_BYTES.load(Ordering::Relaxed)
}

/// Set the package-wide decompressed-byte limit for every package opened
/// afterwards. `u64::MAX` disables the bound.
pub fn set_max_package_bytes(bytes: u64) {
    MAX_PACKAGE_BYTES.store(bytes, Ordering::Relaxed);
}

/// The default per-document text budget: 256 Mi characters.
pub const DEFAULT_MAX_TEXT_CHARS: usize = 256 << 20;

static MAX_TEXT_CHARS: AtomicUsize = AtomicUsize::new(DEFAULT_MAX_TEXT_CHARS);

/// The per-document text budget in characters, the amount of cell text a
/// spreadsheet may render before output is truncated with a notice.
pub fn max_text_chars() -> usize {
    MAX_TEXT_CHARS.load(Ordering::Relaxed)
}

/// Set the per-document text budget for every document read afterwards.
/// `usize::MAX` disables the bound.
pub fn set_max_text_chars(chars: usize) {
    MAX_TEXT_CHARS.store(chars, Ordering::Relaxed);
}

/// The default per-document budget for image bytes copied into an IR: 1 GiB.
///
/// One image part can be used by any number of pictures; the reader keeps
/// the bytes once, but each `Image` in the IR owns its own copy.
pub const DEFAULT_MAX_IMAGE_BYTES: u64 = 1 << 30;

static MAX_IMAGE_BYTES: AtomicU64 = AtomicU64::new(DEFAULT_MAX_IMAGE_BYTES);

/// The maximum total of image bytes one converted document's IR may hold.
pub fn max_image_bytes() -> u64 {
    MAX_IMAGE_BYTES.load(Ordering::Relaxed)
}

/// Set the image-byte budget for every conversion afterwards. `u64::MAX`
/// disables the bound.
pub fn set_max_image_bytes(bytes: u64) {
    MAX_IMAGE_BYTES.store(bytes, Ordering::Relaxed);
}

thread_local! {
    /// Image bytes still available to the conversion running on this
    /// thread, and whether it ran out.
    static IMAGE_BUDGET: std::cell::Cell<(u64, bool)> =
        const { std::cell::Cell::new((u64::MAX, false)) };
}

/// Start a conversion's image budget at the configured limit.
pub(crate) fn begin_image_budget() {
    IMAGE_BUDGET.with(|b| b.set((max_image_bytes(), false)));
}

/// Charge `bytes` of image data against the running conversion's budget.
/// `false` (from the first refusal on) means the caller leaves the bytes out.
pub(crate) fn charge_image_bytes(bytes: usize) -> bool {
    IMAGE_BUDGET.with(|b| {
        let (remaining, exhausted) = b.get();
        if exhausted {
            return false;
        }
        match remaining.checked_sub(bytes as u64) {
            Some(rest) => {
                b.set((rest, false));
                true
            },
            None => {
                b.set((0, true));
                false
            },
        }
    })
}

/// The warning a conversion records when image bytes were left out, or
/// `None` when the budget held.
pub(crate) fn image_budget_warning() -> Option<String> {
    IMAGE_BUDGET.with(|b| b.get().1).then(|| {
        format!(
            "some image bytes were left out: the document's images exceed the {} byte budget \
             (office_oxide::limits::set_max_image_bytes)",
            max_image_bytes()
        )
    })
}

/// A document's remaining text budget.
#[derive(Debug, Clone, Copy)]
pub(crate) struct TextBudget {
    total: usize,
    remaining: usize,
    exhausted: bool,
}

impl TextBudget {
    /// A fresh budget at the configured limit.
    pub(crate) fn new() -> Self {
        Self::with_limit(max_text_chars())
    }

    pub(crate) fn with_limit(total: usize) -> Self {
        Self {
            total,
            remaining: total,
            exhausted: false,
        }
    }

    /// Charge `chars` against the budget. Returns `false` — once, and
    /// thereafter — when it is spent; the caller stops emitting.
    pub(crate) fn charge(&mut self, chars: usize) -> bool {
        if self.exhausted {
            return false;
        }
        match self.remaining.checked_sub(chars) {
            Some(rest) => {
                self.remaining = rest;
                true
            },
            None => {
                self.remaining = 0;
                self.exhausted = true;
                false
            },
        }
    }

    pub(crate) fn exhausted(&self) -> bool {
        self.exhausted
    }

    /// The line every renderer appends when the budget ran out, so a
    /// truncated output never passes for a complete one.
    pub(crate) fn notice(&self) -> String {
        format!(
            "[output truncated: the document's text exceeds the {} character budget \
             (office_oxide::limits::set_max_text_chars)]",
            self.total
        )
    }
}

/// The longest paragraph-style name exposed, in characters: Word's own limit
/// for a style name.
///
/// A style's name is stored once in the style sheet and copied into every
/// paragraph that uses the style, so without a bound a few kilobytes of
/// document multiply into gigabytes of output (one megabyte of name times the
/// number of paragraphs).
pub(crate) const MAX_STYLE_NAME_CHARS: usize = 255;

#[cfg(test)]
thread_local! {
    /// Bytes of style name examined on this thread, so a test can assert that a
    /// name is examined once per style rather than once per paragraph.
    pub(crate) static STYLE_NAME_BYTES_SCANNED: std::cell::Cell<usize> =
        const { std::cell::Cell::new(0) };
}

/// The name without surrounding whitespace and without a NUL terminator.
///
/// `xstzName` is a null-terminated string, and some producers count the
/// terminator in its length, so `body\0` is the name `body`.
fn trim_style_name(raw: &str) -> &str {
    #[cfg(test)]
    STYLE_NAME_BYTES_SCANNED.with(|c| c.set(c.get() + raw.len()));
    raw.trim_matches(|c: char| c.is_whitespace() || c == '\0')
}

/// What a stored style name is, as far as `Paragraph::style` is concerned.
pub(crate) enum StyleName {
    /// Blank: simply no name, not worth a warning.
    Absent,
    /// Longer than [`MAX_STYLE_NAME_CHARS`] or containing a control
    /// character: not exposed, and reported.
    Unusable,
    /// The trimmed name.
    Usable(String),
}

/// Classify a stored style name in one pass over it.
///
/// Callers do this once per style when the style sheet is parsed and keep the
/// result: a name can be a megabyte, and it is looked up for every paragraph
/// that uses the style.
pub(crate) fn classify_style_name(raw: &str) -> StyleName {
    let name = trim_style_name(raw);
    if name.is_empty() {
        return StyleName::Absent;
    }
    // Stops at the first control character or at the 256th, so the cost is
    // bounded by the limit rather than by the name's length once it is over.
    let mut chars = 0usize;
    for c in name.chars() {
        chars += 1;
        if chars > MAX_STYLE_NAME_CHARS || c.is_control() {
            return StyleName::Unusable;
        }
    }
    StyleName::Usable(name.to_owned())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_budget_is_spent_once_and_stays_spent() {
        let mut b = TextBudget::with_limit(10);
        assert!(b.charge(4));
        assert!(b.charge(6));
        assert!(!b.charge(1));
        assert!(b.exhausted());
        assert!(!b.charge(0), "spent stays spent");
        assert!(b.notice().contains("10 character budget"));
    }

    /// Raising the limit is safe under the parallel test runner; lowering
    /// it would truncate every other test's output.
    #[test]
    fn test_setter_round_trips() {
        let before = max_text_chars();
        set_max_text_chars(before.saturating_add(1));
        assert_eq!(max_text_chars(), before.saturating_add(1));
        set_max_text_chars(before);
        assert_eq!(max_text_chars(), before);
    }
}
