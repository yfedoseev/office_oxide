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

/// The default per-document allowance of empty grid slots: 8 Mi.
///
/// A legacy `.xls` sheet is stored as rows that run to their last used
/// column, so one number in column 255 of each of 65,535 rows costs 256
/// slots per row: 16.7M slots from 65,535 records, and six such sheets
/// made a 7 MB file allocate 2.6 GB. Each slot is a `CellValue` and an
/// `XF` index in the sheet, and converting the sheet to the IR adds a table
/// cell for it, about 70 bytes in all: the earlier 32 Mi allowance let that
/// file reach 0.9 GB on open and 2.1 GB in `to_ir`; 8 Mi bounds both near a
/// quarter of that. Only the *empty* slots that pad a row out to its last
/// cell are charged, never the cells themselves, so a dense 65,536 x 256
/// sheet (which carries 16.7M cells and a file to match) costs nothing, and
/// a sheet whose rows reach their last column with up to about 32,000 rows
/// of nothing in between still fits. A workbook that needs more is cut at
/// the row that does not fit, loudly; processes that read such workbooks
/// raise the allowance with
/// [the setter](crate::limits::set_max_grid_padding_cells).
pub const DEFAULT_MAX_GRID_PADDING_CELLS: usize = 8 << 20;

static MAX_GRID_PADDING_CELLS: AtomicUsize = AtomicUsize::new(DEFAULT_MAX_GRID_PADDING_CELLS);

/// The per-document allowance of empty slots padding spreadsheet rows.
pub fn max_grid_padding_cells() -> usize {
    MAX_GRID_PADDING_CELLS.load(Ordering::Relaxed)
}

/// Set the per-document empty-slot allowance for every document read
/// afterwards. `usize::MAX` disables the bound.
pub fn set_max_grid_padding_cells(cells: usize) {
    MAX_GRID_PADDING_CELLS.store(cells, Ordering::Relaxed);
}

/// The default per-document allowance of inflated image bytes: 256 Mi.
///
/// A compressed metafile picture (EMF/WMF/PICT) declares its uncompressed
/// size and is stored as a zlib stream, which can shrink a run of zeros a
/// thousandfold: four such pictures in a 1 MB `.ppt` inflated to 3.3 GB.
/// Every compressed picture of a document is charged its declared size
/// against this allowance before it is inflated; one that does not fit is
/// left out. Pictures stored uncompressed are bounded by the file itself and
/// cost nothing.
pub const DEFAULT_MAX_INFLATED_IMAGE_BYTES: u64 = 256 << 20;

static MAX_INFLATED_IMAGE_BYTES: AtomicU64 = AtomicU64::new(DEFAULT_MAX_INFLATED_IMAGE_BYTES);

/// The per-document allowance of bytes a legacy file's compressed metafile
/// pictures may inflate to.
pub fn max_inflated_image_bytes() -> u64 {
    MAX_INFLATED_IMAGE_BYTES.load(Ordering::Relaxed)
}

/// Set the per-document inflated-image allowance for every document read
/// afterwards. `u64::MAX` disables the bound.
pub fn set_max_inflated_image_bytes(bytes: u64) {
    MAX_INFLATED_IMAGE_BYTES.store(bytes, Ordering::Relaxed);
}

/// A document's remaining allowance of inflated image bytes.
#[derive(Debug, Clone, Copy)]
pub(crate) struct InflateBudget {
    remaining: u64,
}

impl InflateBudget {
    pub(crate) fn new() -> Self {
        Self::with_limit(max_inflated_image_bytes())
    }

    pub(crate) fn with_limit(total: u64) -> Self {
        Self { remaining: total }
    }

    /// Reserve `bytes` before inflating them. `false` when they do not fit;
    /// nothing is taken then, so a smaller picture later can still fit.
    pub(crate) fn reserve(&mut self, bytes: u64) -> bool {
        match self.remaining.checked_sub(bytes) {
            Some(rest) => {
                self.remaining = rest;
                true
            },
            None => false,
        }
    }

    /// Give back the part of a reservation the picture did not use.
    pub(crate) fn release(&mut self, bytes: u64) {
        self.remaining = self.remaining.saturating_add(bytes);
    }
}

/// A document's remaining allowance of empty grid slots.
#[derive(Debug, Clone, Copy)]
pub(crate) struct GridBudget {
    total: usize,
    remaining: usize,
    exhausted: bool,
}

impl GridBudget {
    pub(crate) fn new() -> Self {
        Self::with_limit(max_grid_padding_cells())
    }

    pub(crate) fn with_limit(total: usize) -> Self {
        Self {
            total,
            remaining: total,
            exhausted: false,
        }
    }

    /// Charge `slots` empty slots. Returns `false` — once, and thereafter —
    /// when the allowance is spent; the caller stops building rows.
    pub(crate) fn charge(&mut self, slots: usize) -> bool {
        if self.exhausted {
            return false;
        }
        match self.remaining.checked_sub(slots) {
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

    pub(crate) fn notice(&self) -> String {
        format!(
            "[workbook truncated: its sparse rows need more than {} empty cells of padding \
             (office_oxide::limits::set_max_grid_padding_cells) — later rows are missing]",
            self.total
        )
    }
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

    /// Take up to `want` characters from what is left: all of them while the
    /// budget lasts, otherwise whatever remains, after which the budget is
    /// spent. For a caller that can use a part of a run (cut a piece short)
    /// rather than all of it or none.
    pub(crate) fn take(&mut self, want: usize) -> usize {
        let granted = want.min(self.remaining);
        if !self.charge(want) {
            return granted;
        }
        want
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

    #[test]
    fn test_take_grants_what_is_left_then_nothing() {
        let mut b = TextBudget::with_limit(10);
        assert_eq!(b.take(4), 4);
        assert_eq!(b.take(10), 6, "the rest is granted, and the budget is spent");
        assert!(b.exhausted());
        assert_eq!(b.take(1), 0);
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
