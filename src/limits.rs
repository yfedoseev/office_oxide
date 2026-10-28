//! Resource limits that bound what an untrusted document can make the
//! crate allocate.
//!
//! A spreadsheet can reference one shared string from every cell — the
//! string is stored once, the rendered text is the string times the cell
//! count. A 110 KB `.xlsx` holding a 1 MB string referenced from 12,000
//! cells rendered to 393 MB; the same string from a million rows would be
//! 32 GB from a 300 KB file. Every reader and converter that can copy
//! stored text charges the characters it emits against one per-document
//! budget ([`max_text_chars`]) and stops, loudly, when it is spent: the
//! XLSX text rendering and IR converter, the legacy `.xls` reader (which
//! materialises each referencing cell's copy at open), the legacy `.doc`
//! piece table and its header, footer and comment stories, the legacy `.ppt`
//! slide, notes and master text, and the DOCX and PPTX Markdown and IR
//! converters for what they copy out of shared storage (hyperlink targets,
//! shared headers, notes, a paragraph style's name — `classify_style_name`
//! bounds that one). Apache POI does the same with
//! `ZipSecureFile.setMaxTextSize`.
//!
//! The text budget is one of several, each for the thing that multiplies:
//! pictures ([`max_image_bytes`], [`max_total_image_bytes`] and, for the
//! compressed metafiles of the legacy formats,
//! [`max_inflated_image_bytes`]); the empty slots and table cells that pad
//! the rows of a sparse sheet ([`max_grid_padding_cells`] and
//! [`max_sheet_grid_padding_cells`] for `.xls`, [`max_ir_padding_bytes`] and
//! [`max_sheet_ir_padding_bytes`] for the IR); the blocks, list items and
//! table cells of parsed Markdown ([`max_markdown_elements`],
//! [`max_markdown_list_items`], [`max_markdown_table_cells`]); and the
//! sheets and table cells a writer will write ([`max_write_sheets`],
//! [`max_write_table_cells`]).
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

/// The default per-document allowance of empty grid slots: 16 Mi.
///
/// A legacy `.xls` sheet is stored as rows that run to their last used
/// column, so one number in column 255 of each of 65,535 rows costs 256
/// slots per row: 16.7M slots from 65,535 records, and six such sheets
/// made a 7 MB file allocate 2.6 GB. Each slot is a `CellValue` and an
/// `XF` index in the sheet, and converting the sheet to the IR adds a table
/// cell for it, about 70 bytes in all: a 32 Mi allowance let that file
/// reach 0.9 GB on open and 2.1 GB in `to_ir`; 16 Mi bounds both near half of
/// that. Only the *empty* slots that pad a row out to its last cell are
/// charged, never the cells themselves, so a dense 65,536 x 256 sheet (which
/// carries 16.7M cells and a file to match) costs nothing, and a sheet whose
/// rows reach their last column with up to about 64,000 rows of nothing in
/// between still fits.
///
/// The document's allowance is shared by its sheets, but no sheet may spend
/// more than [`DEFAULT_MAX_SHEET_GRID_PADDING_CELLS`] of it, so one sparse
/// sheet cannot leave the later sheets of a workbook without rows. A
/// workbook that needs more is cut at the row that does not fit, loudly, and
/// the notice names the sheet; processes that read such workbooks raise the
/// allowance with [the setter](crate::limits::set_max_grid_padding_cells).
pub const DEFAULT_MAX_GRID_PADDING_CELLS: usize = 16 << 20;

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

/// The default allowance of empty grid slots for one `.xls` sheet: 8 Mi.
///
/// This is what the whole document was allowed before sheets had their own:
/// a sheet that fitted then fits now. A realistic wide sheet (65,000 rows of
/// 60 columns, 60% filled) pads about 2.4M slots, so a workbook of several
/// such sheets fits within the document's
/// [`DEFAULT_MAX_GRID_PADDING_CELLS`], while a hostile sparse sheet is cut at
/// this and leaves the rest of the document its allowance.
pub const DEFAULT_MAX_SHEET_GRID_PADDING_CELLS: usize = 8 << 20;

static MAX_SHEET_GRID_PADDING_CELLS: AtomicUsize =
    AtomicUsize::new(DEFAULT_MAX_SHEET_GRID_PADDING_CELLS);

/// The per-sheet allowance of empty slots padding spreadsheet rows.
pub fn max_sheet_grid_padding_cells() -> usize {
    MAX_SHEET_GRID_PADDING_CELLS.load(Ordering::Relaxed)
}

/// Set the per-sheet empty-slot allowance for every document read
/// afterwards. `usize::MAX` disables the bound.
pub fn set_max_sheet_grid_padding_cells(cells: usize) {
    MAX_SHEET_GRID_PADDING_CELLS.store(cells, Ordering::Relaxed);
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

/// How many cut sheets a notice names before it counts the rest.
const GRID_CUTS_NAMED: usize = 5;

/// A document's remaining allowance of empty grid slots: a total for the
/// document, and one for the sheet being read.
#[derive(Debug, Clone)]
pub(crate) struct GridBudget {
    total: usize,
    remaining: usize,
    sheet_total: usize,
    sheet_remaining: usize,
    /// The sheet being read has spent its own allowance.
    sheet_exhausted: bool,
    /// The document's allowance is spent.
    exhausted: bool,
    /// Name of the sheet being read.
    sheet: String,
    /// Sheets cut so far: `(name, cut by the document's allowance)`.
    cuts: Vec<(String, bool)>,
}

impl GridBudget {
    pub(crate) fn new() -> Self {
        Self::with_limits(max_grid_padding_cells(), max_sheet_grid_padding_cells())
    }

    /// An allowance of `total` for the document that no sheet may exceed on
    /// its own: the first sheet to need more than `total` takes it all.
    #[cfg(test)]
    pub(crate) fn with_limit(total: usize) -> Self {
        Self::with_limits(total, total)
    }

    pub(crate) fn with_limits(total: usize, per_sheet: usize) -> Self {
        Self {
            total,
            remaining: total,
            sheet_total: per_sheet,
            sheet_remaining: per_sheet,
            sheet_exhausted: false,
            exhausted: false,
            sheet: String::new(),
            cuts: Vec::new(),
        }
    }

    /// Start a sheet: it gets its own allowance, within what is left of the
    /// document's.
    pub(crate) fn begin_sheet(&mut self, name: &str) {
        self.sheet_remaining = self.sheet_total;
        self.sheet_exhausted = false;
        self.sheet.clear();
        self.sheet.push_str(name);
    }

    /// Charge `slots` empty slots. Returns `false` when they do not fit the
    /// sheet's or the document's allowance, and for every nonzero charge once
    /// either is spent; the caller stops building rows. A charge of zero
    /// always succeeds: a row that needs no padding costs nothing, however
    /// much of the allowance is gone.
    pub(crate) fn charge(&mut self, slots: usize) -> bool {
        if slots == 0 {
            return true;
        }
        if self.exhausted || self.sheet_exhausted {
            return false;
        }
        match (self.remaining.checked_sub(slots), self.sheet_remaining.checked_sub(slots)) {
            (Some(rest), Some(sheet_rest)) => {
                self.remaining = rest;
                self.sheet_remaining = sheet_rest;
                true
            },
            (rest, _) => {
                // The document's allowance is what ran out when it cannot
                // pay; otherwise the sheet's own.
                let by_document = rest.is_none();
                if by_document {
                    self.remaining = 0;
                    self.exhausted = true;
                } else {
                    self.sheet_remaining = 0;
                    self.sheet_exhausted = true;
                }
                self.cuts.push((self.sheet.clone(), by_document));
                false
            },
        }
    }

    /// One notice for the sheets cut by their own allowance and one for the
    /// sheet the document's allowance ran out on, each naming the sheets.
    pub(crate) fn notices(&self) -> Vec<String> {
        let mut out = Vec::new();
        let by_sheet: Vec<&str> = self
            .cuts
            .iter()
            .filter(|(_, by_document)| !by_document)
            .map(|(name, _)| name.as_str())
            .collect();
        if !by_sheet.is_empty() {
            out.push(format!(
                "[sheet truncated: the sparse rows of {} need more than {} empty cells of padding \
                 (office_oxide::limits::set_max_sheet_grid_padding_cells) — later rows of \
                 {} are missing]",
                Self::name_list(&by_sheet),
                self.sheet_total,
                if by_sheet.len() == 1 {
                    "this sheet"
                } else {
                    "these sheets"
                }
            ));
        }
        if let Some((name, _)) = self.cuts.iter().find(|(_, by_document)| *by_document) {
            out.push(format!(
                "[workbook truncated: its sparse rows need more than {} empty cells of padding \
                 (office_oxide::limits::set_max_grid_padding_cells) — rows from sheet '{name}' \
                 on are missing]",
                self.total
            ));
        }
        out
    }

    fn name_list(names: &[&str]) -> String {
        let shown: Vec<String> = names
            .iter()
            .take(GRID_CUTS_NAMED)
            .map(|n| format!("'{n}'"))
            .collect();
        let mut list = shown.join(", ");
        if names.len() > GRID_CUTS_NAMED {
            list.push_str(&format!(" and {} more", names.len() - GRID_CUTS_NAMED));
        }
        format!("sheet{} {list}", if names.len() == 1 { "" } else { "s" })
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
    ///
    /// A charge of no cells is always allowed: a row that needs no padding
    /// costs nothing, whatever an earlier sheet spent.
    pub(crate) fn charge_cells(&mut self, cells: usize) -> bool {
        if cells == 0 {
            return true;
        }
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

    #[cfg(test)]
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
             this sheet's later rows, and the rows of each later sheet that needs padding, are \
             missing]",
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
    /// Image bytes still available to the conversion running on this thread,
    /// and whether it ran out; `None` outside a conversion's scope.
    static IMAGE_BUDGET: std::cell::Cell<Option<(u64, bool)>> =
        const { std::cell::Cell::new(None) };
}

/// A conversion's image budget, alive while the guard is. Dropping it —
/// at the end of the conversion, on an early return or while unwinding —
/// puts back the budget that was running before, so a conversion's spent
/// allowance can never reach the next one on the same thread.
#[must_use = "the budget ends when the guard is dropped"]
pub(crate) struct ImageBudgetScope {
    previous: Option<(u64, bool)>,
}

impl Drop for ImageBudgetScope {
    fn drop(&mut self) {
        let _ = IMAGE_BUDGET.try_with(|b| b.set(self.previous));
    }
}

/// Start a conversion's image budget at the configured limit.
pub(crate) fn begin_image_budget() -> ImageBudgetScope {
    let previous = IMAGE_BUDGET.with(|b| b.replace(Some((max_image_bytes() as u64, false))));
    ImageBudgetScope { previous }
}

thread_local! {
    /// Text the running conversion has copied out of the document's shared
    /// storage (a hyperlink target, a shared header, a speaker-notes body);
    /// `None` outside a conversion's scope.
    static COPIED_TEXT: std::cell::Cell<Option<TextBudget>> =
        const { std::cell::Cell::new(None) };
}

/// A conversion's budget for copied text, alive while the guard is; see
/// [`ImageBudgetScope`].
#[must_use = "the budget ends when the guard is dropped"]
pub(crate) struct CopiedTextScope {
    previous: Option<TextBudget>,
}

impl Drop for CopiedTextScope {
    fn drop(&mut self) {
        let _ = COPIED_TEXT.try_with(|b| b.set(self.previous));
    }
}

/// Start a conversion's budget for copied text at the configured limit
/// ([`max_text_chars`]).
pub(crate) fn begin_copied_text_budget() -> CopiedTextScope {
    let previous = COPIED_TEXT.with(|b| b.replace(Some(TextBudget::new())));
    CopiedTextScope { previous }
}

/// Charge `chars` of text about to be copied into the output against the
/// running conversion's budget, before the copy is made. `false` (from the
/// first refusal on) means the caller leaves the text out. Outside a
/// conversion's scope each copy is held to the configured limit on its own,
/// as no running total exists to charge.
pub(crate) fn charge_copied_text(chars: usize) -> bool {
    COPIED_TEXT.with(|b| match b.get() {
        Some(mut budget) => {
            let ok = budget.charge(chars);
            b.set(Some(budget));
            ok
        },
        None => chars <= max_text_chars(),
    })
}

/// The line to report when copied text was left out, or `None` when the
/// budget held.
pub(crate) fn copied_text_notice() -> Option<String> {
    COPIED_TEXT.with(|b| {
        let budget = b.get()?;
        budget.exhausted().then(|| budget.notice())
    })
}

/// Charge `bytes` of image data against the running conversion's budget.
/// `false` (from the first refusal on) means the caller leaves the bytes out.
/// Outside a conversion's scope each picture is held to the configured
/// limit on its own, as no running total exists to charge.
pub(crate) fn charge_image_bytes(bytes: usize) -> bool {
    IMAGE_BUDGET.with(|b| {
        let Some((remaining, exhausted)) = b.get() else {
            return bytes as u64 <= max_image_bytes() as u64;
        };
        if exhausted {
            return false;
        }
        match remaining.checked_sub(bytes as u64) {
            Some(rest) => {
                b.set(Some((rest, false)));
                true
            },
            None => {
                b.set(Some((0, true)));
                false
            },
        }
    })
}

/// The warning a conversion records when image bytes were left out, or
/// `None` when the budget held.
pub(crate) fn image_budget_warning() -> Option<String> {
    IMAGE_BUDGET
        .with(|b| b.get())
        .is_some_and(|(_, exhausted)| exhausted)
        .then(|| {
            format!(
                "some image bytes were left out: the document's images exceed the {} byte \
                 budget (office_oxide::limits::set_max_image_bytes)",
                max_image_bytes()
            )
        })
}

/// The line a rendering of the IR shows when image bytes were left out, or
/// `None` when the budget held. Said in the content itself because a
/// Markdown or HTML rendering carries no metadata.
pub(crate) fn image_budget_notice() -> Option<String> {
    IMAGE_BUDGET
        .with(|b| b.get().is_some_and(|state| state.1))
        .then(|| {
            format!(
                "[some pictures not shown: together they exceed the {} byte picture budget \
             (office_oxide::limits::set_max_image_bytes)]",
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

/// The default most worksheets an IR written to `.xlsx` may hold: 32,767.
///
/// Every sheet is a part of its own with a relationships part beside it, so
/// this keeps the package under the 65,535 entries a zip without ZIP64
/// extensions can hold. Excel itself has no fixed sheet limit (memory
/// decides) and openpyxl none at all, but a genuine workbook has tens or
/// hundreds of sheets, not tens of thousands.
pub const DEFAULT_MAX_WRITE_SHEETS: usize = 32_767;

static MAX_WRITE_SHEETS: AtomicUsize = AtomicUsize::new(DEFAULT_MAX_WRITE_SHEETS);

/// The most sections (worksheets) an IR written to `.xlsx` may hold before
/// the write is refused.
pub fn max_write_sheets() -> usize {
    MAX_WRITE_SHEETS.load(Ordering::Relaxed)
}

/// Set the worksheet limit for every write afterwards. `usize::MAX`
/// disables the bound.
pub fn set_max_write_sheets(sheets: usize) {
    MAX_WRITE_SHEETS.store(sheets, Ordering::Relaxed);
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

/// The default most blocks one Markdown document may parse into: 100,000.
///
/// A block costs a section, a paragraph, a heading, a code block, a break,
/// a list or a table as 1. Each is 0.5-1.4 KB of IR from a few bytes of
/// Markdown: 100,000 sections were 214 MB and 200,000 paragraphs 285 MB.
/// List items and table cells have limits of their own,
/// [`max_markdown_list_items`] and [`max_markdown_table_cells`].
pub const DEFAULT_MAX_MARKDOWN_ELEMENTS: usize = 100_000;

/// The default most table cells one Markdown document may parse into: 150,000.
///
/// A cell is about 1.2 KB of IR whether it is in a header row or a body row
/// (measured: peak memory of 118 MB for a 100,000-column header row, 292 MB
/// for 250,000), so the default keeps the cells of one document near 180 MB;
/// with the block limit's own bound the IR of a parse stays under about
/// 350 MB. A table of 1,000 rows by 25 columns is a sixth of it. Past it the rows that do not fit are dropped, with a
/// notice, and the rest of the document is read.
pub const DEFAULT_MAX_MARKDOWN_TABLE_CELLS: usize = 150_000;

static MAX_MARKDOWN_TABLE_CELLS: AtomicUsize = AtomicUsize::new(DEFAULT_MAX_MARKDOWN_TABLE_CELLS);

/// The most table cells `DocumentIR::from_markdown` keeps, over all of a
/// document's tables.
pub fn max_markdown_table_cells() -> usize {
    MAX_MARKDOWN_TABLE_CELLS.load(Ordering::Relaxed)
}

/// Set the Markdown table-cell limit for every parse afterwards.
/// `usize::MAX` disables the bound.
pub fn set_max_markdown_table_cells(cells: usize) {
    MAX_MARKDOWN_TABLE_CELLS.store(cells, Ordering::Relaxed);
}

static MAX_MARKDOWN_ELEMENTS: AtomicUsize = AtomicUsize::new(DEFAULT_MAX_MARKDOWN_ELEMENTS);

/// The most blocks (sections, paragraphs, headings, code blocks, breaks,
/// lists and tables) `DocumentIR::from_markdown`
/// keeps; the rest of the document is dropped with a notice.
pub fn max_markdown_elements() -> usize {
    MAX_MARKDOWN_ELEMENTS.load(Ordering::Relaxed)
}

/// Set the Markdown block limit for every parse afterwards. `usize::MAX`
/// disables the bound.
pub fn set_max_markdown_elements(elements: usize) {
    MAX_MARKDOWN_ELEMENTS.store(elements, Ordering::Relaxed);
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
        assert!(b.charge_cells(0), "a row that needs no padding costs nothing, spent or not");
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
    fn test_grid_notices_name_the_sheets_and_the_limit_that_cut_them() {
        let mut b = GridBudget::with_limits(100, 10);
        b.begin_sheet("A");
        assert!(b.charge(10));
        assert!(!b.charge(1), "the sheet's allowance is spent");
        assert!(b.charge(0));
        b.begin_sheet("B");
        assert!(b.charge(10), "a new sheet has its own allowance");
        for name in ["C", "D", "E", "F", "G", "H"] {
            b.begin_sheet(name);
            assert!(!b.charge(11));
        }
        let notices = b.notices();
        assert_eq!(notices.len(), 1, "{notices:?}");
        assert!(notices[0].contains("'A'") && notices[0].contains("and 2 more"), "{notices:?}");
        assert!(notices[0].contains("set_max_sheet_grid_padding_cells"));
        let mut b = GridBudget::with_limits(15, 10);
        b.begin_sheet("A");
        assert!(b.charge(10));
        b.begin_sheet("B");
        assert!(!b.charge(6), "only 5 of the document's allowance is left");
        let notices = b.notices();
        assert!(
            notices[0].contains("sheet 'B'") && notices[0].contains("set_max_grid_padding_cells")
        );
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
        // One more than the default of a limit that only sheet-by-sheet
        // conversions read, so no other test in this binary can notice it.
        let mut limits = test_support::LimitScope::new();
        limits.set(
            max_total_image_bytes,
            set_max_total_image_bytes,
            DEFAULT_MAX_TOTAL_IMAGE_BYTES + 1,
        );
        assert_eq!(max_total_image_bytes(), DEFAULT_MAX_TOTAL_IMAGE_BYTES + 1);
        drop(limits);
        assert_eq!(max_total_image_bytes(), DEFAULT_MAX_TOTAL_IMAGE_BYTES);
    }

    /// A conversion that returns early (or fails) leaves nothing behind for
    /// the next one on the thread: the scopes end with the function, so the
    /// budgets of one conversion are never the next one's.
    #[test]
    fn test_an_early_return_leaves_no_budget_for_the_next_conversion() {
        fn conversion(fail: bool) -> bool {
            let _image = begin_image_budget();
            let _text = begin_copied_text_budget();
            assert!(!charge_copied_text(usize::MAX));
            assert!(!charge_image_bytes(usize::MAX));
            assert!(copied_text_notice().is_some());
            assert!(image_budget_warning().is_some());
            if fail {
                return false;
            }
            true
        }
        assert!(!conversion(true));
        assert!(copied_text_notice().is_none(), "the spent text budget leaked");
        assert!(image_budget_warning().is_none(), "the spent image budget leaked");

        // A panic unwinding through a conversion ends its scope as well.
        let unwound = std::panic::catch_unwind(|| {
            let _text = begin_copied_text_budget();
            let _image = begin_image_budget();
            charge_copied_text(usize::MAX);
            charge_image_bytes(usize::MAX);
            panic!("conversion failed");
        });
        assert!(unwound.is_err());
        assert!(copied_text_notice().is_none());
        assert!(image_budget_warning().is_none());

        // The next conversion starts fresh.
        let _text = begin_copied_text_budget();
        assert!(charge_copied_text(10));
        assert!(copied_text_notice().is_none());
    }

    /// A scope inside another gives the outer budget back when it ends.
    #[test]
    fn test_a_nested_scope_restores_the_budget_it_replaced() {
        let _outer = begin_copied_text_budget();
        assert!(charge_copied_text(5));
        {
            let _inner = begin_copied_text_budget();
            assert!(!charge_copied_text(usize::MAX));
            assert!(copied_text_notice().is_some());
        }
        assert!(
            copied_text_notice().is_none(),
            "the inner budget replaced the outer one for good"
        );
        assert!(charge_copied_text(5));
    }

    /// Outside any conversion there is no running total, but a charge is not
    /// unlimited either: it is held to the configured limit on its own.
    #[test]
    fn test_charges_outside_a_conversion_are_held_to_the_configured_limit() {
        assert!(charge_image_bytes(1));
        assert!(!charge_image_bytes(usize::MAX));
        assert!(charge_copied_text(1));
        assert!(!charge_copied_text(usize::MAX));
        assert!(copied_text_notice().is_none());
        assert!(image_budget_warning().is_none());
    }
}

/// Serialises tests that change a process-wide limit, and puts the limit
/// back when the test ends, however it ends.
#[cfg(test)]
pub(crate) mod test_support {
    use std::sync::{Mutex, MutexGuard};

    static LIMITS: Mutex<()> = Mutex::new(());

    /// Holds the lock on the limits for as long as it lives and restores
    /// every limit it set, in reverse order, when dropped.
    pub(crate) struct LimitScope {
        restore: Vec<Box<dyn FnOnce()>>,
        _lock: MutexGuard<'static, ()>,
    }

    impl LimitScope {
        pub(crate) fn new() -> Self {
            Self {
                restore: Vec::new(),
                // A test that failed while holding the lock leaves it
                // poisoned, and its guard has already restored the limits.
                _lock: LIMITS.lock().unwrap_or_else(|e| e.into_inner()),
            }
        }

        /// Set a limit for the rest of the scope.
        pub(crate) fn set<T: Copy + 'static>(&mut self, get: fn() -> T, set: fn(T), value: T) {
            let before = get();
            set(value);
            self.restore.push(Box::new(move || set(before)));
        }
    }

    impl Drop for LimitScope {
        fn drop(&mut self) {
            while let Some(restore) = self.restore.pop() {
                restore();
            }
        }
    }
}
