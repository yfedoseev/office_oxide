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

    /// Charge `slots` empty slots. Returns `false` when they do not fit, and
    /// for every nonzero charge once the allowance is spent; the caller stops
    /// building rows. A charge of zero always succeeds: a row that needs no
    /// padding costs nothing, however much of the allowance is gone.
    pub(crate) fn charge(&mut self, slots: usize) -> bool {
        if slots == 0 {
            return true;
        }
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

/// The default allowance of memory for the empty table cells that pad the
/// rows of an `.xlsx` sheet in the IR: 512 MiB.
///
/// A spreadsheet's IR table is as wide as the sheet, so a row with cells at
/// `A` and `CV` is laid out with 98 empty table cells between them, each a
/// `TableCell` and its paragraph element, about 600 bytes. They are charged
/// at that size, apart from the text budget (which is also charged 3
/// characters a cell, as Markdown charges an absent column): at the default a
/// sheet may hold about 900,000 of them, enough for a multi-sheet export of
/// a hundred sparse columns and several thousand rows.
///
/// This is the allowance of a conversion that holds the whole workbook at once
/// (`Document::to_ir`, which every sheet shares). A conversion that holds one
/// sheet at a time has each sheet's own,
/// [`DEFAULT_MAX_SHEET_IR_PADDING_BYTES`].
pub const DEFAULT_MAX_IR_PADDING_BYTES: usize = 512 << 20;

static MAX_IR_PADDING_BYTES: AtomicUsize = AtomicUsize::new(DEFAULT_MAX_IR_PADDING_BYTES);

/// The per-document allowance of bytes for the empty cells padding the rows
/// of a spreadsheet's IR.
pub fn max_ir_padding_bytes() -> usize {
    MAX_IR_PADDING_BYTES.load(Ordering::Relaxed)
}

/// Set the empty-cell padding allowance for every document converted
/// afterwards. `usize::MAX` disables the bound.
pub fn set_max_ir_padding_bytes(bytes: usize) {
    MAX_IR_PADDING_BYTES.store(bytes, Ordering::Relaxed);
}

/// The default allowance of memory for the empty table cells that pad the
/// rows of one `.xlsx` sheet while the sheet is held in the IR: 1 GiB.
///
/// When a workbook is converted a sheet at a time (the sheet-by-sheet IR of
/// [`XlsxDocument::sheets_ir`](crate::xlsx::XlsxDocument::sheets_ir), and
/// `to_html`, `to_markdown_with` over a spreadsheet) only one sheet's cells
/// exist at once, so the allowance is the sheet's, not the document's, and it
/// is given back when the sheet is released. At about 600 bytes a cell the
/// default holds about 1.8 million of them in one sheet: 18,000 rows of a
/// hundred sparse columns. The characters the padding adds to the output are
/// still charged, for the whole run, to the text budget's limit.
pub const DEFAULT_MAX_SHEET_IR_PADDING_BYTES: usize = 1 << 30;

static MAX_SHEET_IR_PADDING_BYTES: AtomicUsize =
    AtomicUsize::new(DEFAULT_MAX_SHEET_IR_PADDING_BYTES);

/// The per-sheet allowance of bytes for the empty cells padding the rows of
/// a spreadsheet sheet in the IR, when sheets are converted one at a time.
pub fn max_sheet_ir_padding_bytes() -> usize {
    MAX_SHEET_IR_PADDING_BYTES.load(Ordering::Relaxed)
}

/// Set the per-sheet empty-cell padding allowance for every sheet converted
/// afterwards. `usize::MAX` disables the bound.
pub fn set_max_sheet_ir_padding_bytes(bytes: usize) {
    MAX_SHEET_IR_PADDING_BYTES.store(bytes, Ordering::Relaxed);
}

/// A remaining allowance of bytes for IR padding cells: the document's, or
/// one sheet's.
#[derive(Debug, Clone, Copy)]
pub(crate) struct IrPaddingBudget {
    total: usize,
    remaining: usize,
    exhausted: bool,
    per_sheet: bool,
}

impl IrPaddingBudget {
    pub(crate) fn new() -> Self {
        Self::with_limit(max_ir_padding_bytes())
    }

    /// A fresh allowance for one sheet.
    pub(crate) fn new_for_sheet() -> Self {
        Self {
            per_sheet: true,
            ..Self::with_limit(max_sheet_ir_padding_bytes())
        }
    }

    pub(crate) fn with_limit(total: usize) -> Self {
        Self {
            total,
            remaining: total,
            exhausted: false,
            per_sheet: false,
        }
    }

    /// What one empty table cell holds in the IR: the cell and the element of
    /// its paragraph.
    pub(crate) fn cell_bytes() -> usize {
        std::mem::size_of::<crate::ir::TableCell>() + std::mem::size_of::<crate::ir::Element>()
    }

    /// Charge `cells` empty cells. `false` — from the first refusal on —
    /// when the allowance is spent; the caller stops building rows.
    pub(crate) fn charge_cells(&mut self, cells: usize) -> bool {
        if self.exhausted {
            return false;
        }
        match cells
            .checked_mul(Self::cell_bytes())
            .and_then(|bytes| self.remaining.checked_sub(bytes))
        {
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
        if self.per_sheet {
            return format!(
                "[sheet truncated: its rows are padded to the sheet's width with more than {} \
                 bytes of empty table cells (office_oxide::limits::set_max_sheet_ir_padding_bytes) \
                 — later rows of this sheet are missing]",
                self.total
            );
        }
        format!(
            "[workbook truncated: its rows are padded to the sheet's width with more than {} \
             bytes of empty table cells (office_oxide::limits::set_max_ir_padding_bytes) — \
             later rows are missing]",
            self.total
        )
    }
}

/// The default budget for picture bytes in the IR: 128 MiB.
///
/// The IR gives every picture its own copy of the bytes, and its JSON form
/// writes them again as base64 (a third larger): at the earlier 1 GiB a 5 KB
/// `.docx` drawing one 1 MB image 1,100 times converted to 1 GB of IR and 1.4 GB
/// of JSON. Genuine documents are far smaller: a photo-heavy report or a deck of
/// a hundred full-resolution photographs is tens of megabytes of distinct
/// pictures, each used once. A picture past the budget keeps its place without
/// its bytes, and the conversion says so; processes that convert larger
/// documents raise it with [the setter](crate::limits::set_max_image_bytes).
pub const DEFAULT_MAX_IMAGE_BYTES: usize = 128 << 20;

static MAX_IMAGE_BYTES: AtomicUsize = AtomicUsize::new(DEFAULT_MAX_IMAGE_BYTES);

/// The bytes of picture data a spreadsheet's IR may hold, counting a picture
/// once for every anchor that draws it: the workbook stores a media part once,
/// and the IR gives each image its own copy.
pub fn max_image_bytes() -> usize {
    MAX_IMAGE_BYTES.load(Ordering::Relaxed)
}

/// Set the picture-byte budget for every document converted afterwards.
/// `usize::MAX` disables the bound.
pub fn set_max_image_bytes(bytes: usize) {
    MAX_IMAGE_BYTES.store(bytes, Ordering::Relaxed);
}

/// The default total of picture bytes a run may copy into the IR when a
/// workbook is converted a sheet at a time: 4 GiB, four times a sheet's.
///
/// A sheet's own allowance ([`max_image_bytes`]) is given back with the sheet,
/// so without a total, a workbook of many sheets each drawing the same picture
/// would copy it once per sheet without bound.
pub const DEFAULT_MAX_TOTAL_IMAGE_BYTES: u64 = 4 << 30;

static MAX_TOTAL_IMAGE_BYTES: AtomicU64 = AtomicU64::new(DEFAULT_MAX_TOTAL_IMAGE_BYTES);

/// The most picture bytes all the sheets of one workbook may copy into the IR
/// in a sheet-by-sheet conversion.
pub fn max_total_image_bytes() -> u64 {
    MAX_TOTAL_IMAGE_BYTES.load(Ordering::Relaxed)
}

/// Set the total picture-byte limit for every sheet-by-sheet conversion
/// afterwards. `u64::MAX` disables the bound.
pub fn set_max_total_image_bytes(bytes: u64) {
    MAX_TOTAL_IMAGE_BYTES.store(bytes, Ordering::Relaxed);
}

thread_local! {
    /// Image bytes still available to the conversion running on this
    /// thread, and whether it ran out.
    static IMAGE_BUDGET: std::cell::Cell<(u64, bool)> =
        const { std::cell::Cell::new((u64::MAX, false)) };
}

/// Start a conversion's image budget at the configured limit.
pub(crate) fn begin_image_budget() {
    IMAGE_BUDGET.with(|b| b.set((max_image_bytes() as u64, false)));
}

thread_local! {
    /// Text the running conversion has copied out of the document's shared
    /// storage (a hyperlink target, a shared header, a speaker-notes body).
    static COPIED_TEXT: std::cell::RefCell<TextBudget> =
        std::cell::RefCell::new(TextBudget::new());
}

/// Start a conversion's budget for copied text at the configured limit
/// ([`max_text_chars`]).
pub(crate) fn begin_copied_text_budget() {
    COPIED_TEXT.with(|b| *b.borrow_mut() = TextBudget::new());
}

/// Charge `chars` of text about to be copied into the output against the
/// running conversion's budget, before the copy is made. `false` (from the
/// first refusal on) means the caller leaves the text out.
pub(crate) fn charge_copied_text(chars: usize) -> bool {
    COPIED_TEXT.with(|b| b.borrow_mut().charge(chars))
}

/// The line to report when copied text was left out, or `None` when the
/// budget held.
pub(crate) fn copied_text_notice() -> Option<String> {
    COPIED_TEXT.with(|b| {
        let b = b.borrow();
        b.exhausted().then(|| b.notice())
    })
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

/// The default most table cells one IR may hold when it is written to
/// `.docx` or `.pptx`: 250,000.
///
/// The writers build every cell in memory; a million-cell table took
/// 2.2 GB to write. Genuine tables in these formats stay far below this.
pub const DEFAULT_MAX_WRITE_TABLE_CELLS: usize = 250_000;

static MAX_WRITE_TABLE_CELLS: AtomicUsize = AtomicUsize::new(DEFAULT_MAX_WRITE_TABLE_CELLS);

/// The most table cells (nested tables included) an IR written to `.docx` or
/// `.pptx` may hold before the write is refused.
pub fn max_write_table_cells() -> usize {
    MAX_WRITE_TABLE_CELLS.load(Ordering::Relaxed)
}

/// Set the table-cell limit for every write afterwards. `usize::MAX`
/// disables the bound.
pub fn set_max_write_table_cells(cells: usize) {
    MAX_WRITE_TABLE_CELLS.store(cells, Ordering::Relaxed);
}

/// The default most list items one Markdown document may parse into:
/// 100,000. Each item becomes about 250 bytes of IR.
pub const DEFAULT_MAX_MARKDOWN_LIST_ITEMS: usize = 100_000;

static MAX_MARKDOWN_LIST_ITEMS: AtomicUsize = AtomicUsize::new(DEFAULT_MAX_MARKDOWN_LIST_ITEMS);

/// The most list items (nested ones included) `DocumentIR::from_markdown`
/// keeps; the rest are dropped with a notice.
pub fn max_markdown_list_items() -> usize {
    MAX_MARKDOWN_LIST_ITEMS.load(Ordering::Relaxed)
}

/// Set the Markdown list-item limit for every parse afterwards.
/// `usize::MAX` disables the bound.
pub fn set_max_markdown_list_items(items: usize) {
    MAX_MARKDOWN_LIST_ITEMS.store(items, Ordering::Relaxed);
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

    /// The limit this budget was created with.
    pub(crate) fn limit(&self) -> usize {
        self.total
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

/// The longest formula, in characters: Excel's own limit (ECMA-376 Part 1
/// §18.17, 8,192 characters). A shared formula's master is copied into every
/// cell that follows it, so a longer one is refused.
pub(crate) const MAX_FORMULA_CHARS: usize = 8192;

/// The longest hyperlink target kept, in characters: Excel's own limit for a
/// hyperlink address (2,083, the length Internet Explorer accepts for a URL).
///
/// A relationship's target is stored once and named by every `<hyperlink>`
/// that uses it, so a megabyte target multiplies by the link count into the
/// reader's maps, the IR and every rendering. A longer one cannot come from a
/// spreadsheet application: the links that name it are left out, reported.
pub(crate) const MAX_HYPERLINK_TARGET_CHARS: usize = 2083;

/// The longest number-format code listed on an IR cell, in characters:
/// Excel's own limit for a custom format (255).
///
/// A format code is stored once in the stylesheet and listed on every numeric
/// cell that uses the style, so a megabyte code would be a megabyte per cell.
/// A longer code cannot come from a spreadsheet application; the cells keep
/// their format id and their values, and the code is left off, reported.
pub(crate) const MAX_NUMBER_FORMAT_CHARS: usize = 255;

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
    fn test_ir_padding_is_charged_at_the_size_of_the_cells_it_makes() {
        let mut b = IrPaddingBudget::with_limit(IrPaddingBudget::cell_bytes() * 10);
        assert!(b.charge_cells(10));
        assert!(!b.charge_cells(1));
        assert!(b.exhausted());
        assert!(b.notice().contains("empty table cells"));
        assert!(!IrPaddingBudget::with_limit(usize::MAX).charge_cells(usize::MAX));
    }

    #[test]
    fn test_a_sheets_padding_allowance_names_its_own_limit_and_the_sheet() {
        let b = IrPaddingBudget::new_for_sheet();
        assert!(b.notice().contains("set_max_sheet_ir_padding_bytes"));
        assert!(b.notice().starts_with("[sheet truncated"));
        let d = IrPaddingBudget::new();
        assert!(d.notice().contains("set_max_ir_padding_bytes"));
        assert!(d.notice().starts_with("[workbook truncated"));
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
