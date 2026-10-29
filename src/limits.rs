//! Resource limits that bound what an untrusted document can make the
//! crate allocate.
//!
//! A spreadsheet can reference one shared string from every cell — the
//! string is stored once, the rendered text is the string times the cell
//! count. A 110 KB `.xlsx` holding a 1 MB string referenced from 12,000
//! cells rendered to 393 MB; the same string from a million rows would be
//! 32 GB from a 300 KB file. Every reader and converter that can copy
//! stored text charges the characters it emits against one per-document
//! budget ([`crate::limits::max_text_chars`]) and stops, loudly, when it is spent: the
//! XLSX text rendering and IR converter, the legacy `.xls` reader (which
//! materialises each referencing cell's copy at open), the legacy `.doc`
//! piece table and its header, footer and comment stories, the legacy `.ppt`
//! slide, notes and master text, and the DOCX and PPTX Markdown and IR
//! converters for what they copy out of shared storage (hyperlink targets,
//! shared headers, notes, a paragraph style's name — `classify_style_name`
//! bounds that one).
//!
//! **Two ways to set them.** The `set_max_*` functions change the
//! process-wide defaults: every document read or written afterwards, on every
//! thread, so set them once at start-up. For limits that belong to one
//! document (a server with different tenants, a test beside others) use
//! `Limits` in [`ParseOptions`](crate::options::ParseOptions) with
//! [`Document::open_with`](crate::Document::open_with): they apply to that
//! document's reading and conversions and to nothing else. (The tuning
//! constants behind the scaled text budget and the write-memory estimate are
//! `doc(hidden)`: they may change.)
//!
//! **Eight settings, grouped by what they guard.** The zip layer has the four
//! POI's `ZipSecureFile` has: a part's decompressed size
//! ([`max_part_bytes()`](crate::limits::max_part_bytes)), a package's
//! ([`max_package_bytes()`](crate::limits::max_package_bytes)), its entries
//! ([`max_package_entries()`](crate::limits::max_package_entries)) and a
//! part's inflate ratio ([`max_inflate_ratio()`](crate::limits::max_inflate_ratio)).
//! The output has the text budget (below). Embedded packages have a depth and
//! a count ([`max_embedded_depth()`](crate::limits::max_embedded_depth),
//! [`max_embedded_packages()`](crate::limits::max_embedded_packages)). And one
//! memory budget ([`max_memory_bytes()`](crate::limits::max_memory_bytes), as
//! `image::Limits::max_alloc` is) bounds everything a small file can multiply
//! in memory: pictures copied or inflated, the empty cells that pad a sparse
//! sheet, and what a writer builds. Each of those bounds is a fixed share of
//! it, so lowering the one number lowers them together. A bound finer than
//! that (Markdown blocks, list items and table cells, Excel's 32,767 sheets,
//! formula and style-name lengths) is a constant, not a setting.
//!
//! A package embedded in a `.docx` is opened to read its text, and may embed
//! packages of its own. What they cost is charged to the one document that
//! holds them: how deep they nest, how many are opened, and the bytes they take
//! (a quarter of the package budget), with their text counted against the text
//! budget.
//!
//! The default keeps every genuine workbook this crate has been tested
//! against (the largest, 181 MB of cell text, with room to spare) and
//! refuses the fan-out shapes. Processes that read larger workbooks raise
//! it once at startup with [the text-budget setter](crate::limits::set_max_text_chars).

//!
//! OOXML packages are zip archives. A part is bounded by its decompressed size
//! ([`max_part_bytes()`](crate::limits::max_part_bytes), 512 MiB) and, once it is larger than
//! `INFLATE_RATIO_FLOOR_BYTES`, by its ratio to the bytes that store it
//! ([`max_inflate_ratio()`](crate::limits::max_inflate_ratio), 1,000; Apache POI's `ZipSecureFile` does the same
//! with 100). Two package-wide bounds sit on top, as POI's do: the number of
//! entries an archive may hold ([`max_package_entries()`](crate::limits::max_package_entries)) and the total
//! decompressed bytes read from one package ([`max_package_bytes()`](crate::limits::max_package_bytes)). A
//! package or part that exceeds any is refused with
//! [`Error::PackageLimit`](crate::core::Error::PackageLimit) or
//! [`Error::DecompressionLimit`](crate::core::Error::DecompressionLimit).

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
    scoped(|l| l.package_entries).unwrap_or_else(|| MAX_PACKAGE_ENTRIES.load(Ordering::Relaxed))
}

/// Set the maximum number of entries for every package opened afterwards.
/// `usize::MAX` disables the bound.
pub fn set_max_package_entries(entries: usize) {
    MAX_PACKAGE_ENTRIES.store(entries, Ordering::Relaxed);
}

/// The maximum total of decompressed bytes read from one OOXML package.
/// Each distinct part is charged once, however often it is read; reading a
/// part again, and reads that fail after inflating, are totalled separately
/// and each held to this limit too, so the work a package can cause is bounded
/// by a small multiple of it.
///
/// For a package embedded in a document this is what the document's
/// embedded-package allowance (a quarter of this limit) has left, if that is
/// less.
pub fn max_package_bytes() -> u64 {
    let configured =
        scoped(|l| l.package_bytes).unwrap_or_else(|| MAX_PACKAGE_BYTES.load(Ordering::Relaxed));
    crate::core::embedded::package_bytes_cap().map_or(configured, |cap| configured.min(cap))
}

/// Set the package-wide decompressed-byte limit for every package opened
/// afterwards. `u64::MAX` disables the bound.
pub fn set_max_package_bytes(bytes: u64) {
    MAX_PACKAGE_BYTES.store(bytes, Ordering::Relaxed);
}

/// The default largest decompressed size of one part of a package: 512 MiB.
///
/// Real documents do not have half-gigabyte parts; a part that claims to is
/// corrupt or a decompression bomb. Raise it for a collection known to hold
/// very large worksheets; lower it to bound the memory one part can take.
pub const DEFAULT_MAX_PART_BYTES: u64 = 512 << 20;

/// The default largest ratio of a part's decompressed to its compressed size,
/// once the part is larger than [`INFLATE_RATIO_FLOOR_BYTES`]: 1,000.
///
/// Ordinary XML deflates 5-20 times; a sparse worksheet reaches a few
/// hundred. A part 1,000 times larger than the bytes that encode it, and
/// large besides, is a bomb. Apache POI refuses a ratio above 100 for any
/// size (and added a size floor after reporting small, legitimate parts).
pub const DEFAULT_MAX_INFLATE_RATIO: u64 = 1000;

/// A part no larger than this (decompressed) is never refused for its ratio:
/// 64 MiB. A ratio only says something about a part big enough to hurt.
pub const INFLATE_RATIO_FLOOR_BYTES: u64 = 64 << 20;

static MAX_PART_BYTES: AtomicU64 = AtomicU64::new(DEFAULT_MAX_PART_BYTES);
static MAX_INFLATE_RATIO: AtomicU64 = AtomicU64::new(DEFAULT_MAX_INFLATE_RATIO);

/// The largest decompressed size accepted for one part of a package.
pub fn max_part_bytes() -> u64 {
    scoped(|l| l.part_bytes).unwrap_or_else(|| MAX_PART_BYTES.load(Ordering::Relaxed))
}

/// Set the per-part decompressed-size limit for every package opened
/// afterwards. `u64::MAX` disables the bound.
pub fn set_max_part_bytes(bytes: u64) {
    MAX_PART_BYTES.store(bytes, Ordering::Relaxed);
}

/// The largest decompressed-to-compressed ratio accepted for a part larger
/// than [`INFLATE_RATIO_FLOOR_BYTES`].
pub fn max_inflate_ratio() -> u64 {
    scoped(|l| l.inflate_ratio).unwrap_or_else(|| MAX_INFLATE_RATIO.load(Ordering::Relaxed))
}

/// Set the inflate-ratio limit for every package opened afterwards.
/// `u64::MAX` disables the check.
pub fn set_max_inflate_ratio(ratio: u64) {
    MAX_INFLATE_RATIO.store(ratio, Ordering::Relaxed);
}

/// The most a part with `compressed` stored bytes may inflate to: the
/// per-part size limit, and (past the floor) the ratio limit.
pub(crate) fn part_inflate_cap(compressed: u64) -> u64 {
    let size = max_part_bytes();
    let by_ratio = compressed
        .saturating_mul(max_inflate_ratio())
        .max(INFLATE_RATIO_FLOOR_BYTES);
    size.min(by_ratio)
}

/// The default depth to which embedded packages are opened: 3.
///
/// A `.docx` can embed an OOXML package (`<o:OLEObject Type="Embed">`) whose
/// text is folded into the document's, and that package can embed another.
/// Documents in the wild nest one level (a worksheet pasted into a letter);
/// a package that embeds a package that embeds a package is already unusual.
/// A package embedded deeper than this is not opened: it is reported in
/// `unreadable_parts` and `Metadata::warnings`, and the document is flagged
/// `text_truncated`.
pub const DEFAULT_MAX_EMBEDDED_DEPTH: usize = 3;

/// The default number of embedded packages opened for one document: 256.
///
/// A package reached again through another entry name is recognised by its
/// content and opened once, so this counts distinct embedded packages,
/// at every depth together.
pub const DEFAULT_MAX_EMBEDDED_PACKAGES: usize = 256;

static MAX_EMBEDDED_DEPTH: AtomicUsize = AtomicUsize::new(DEFAULT_MAX_EMBEDDED_DEPTH);
static MAX_EMBEDDED_PACKAGES: AtomicUsize = AtomicUsize::new(DEFAULT_MAX_EMBEDDED_PACKAGES);
/// How many levels deep an embedded package is opened (the document itself is
/// level 0, a package it embeds level 1).
pub fn max_embedded_depth() -> usize {
    scoped(|l| l.embedded_depth).unwrap_or_else(|| MAX_EMBEDDED_DEPTH.load(Ordering::Relaxed))
}

/// Set the embedding depth limit for every document read afterwards. `0`
/// opens no embedded package; `usize::MAX` disables the bound.
pub fn set_max_embedded_depth(depth: usize) {
    MAX_EMBEDDED_DEPTH.store(depth, Ordering::Relaxed);
}

/// The most distinct embedded packages opened for one document.
pub fn max_embedded_packages() -> usize {
    scoped(|l| l.embedded_packages).unwrap_or_else(|| MAX_EMBEDDED_PACKAGES.load(Ordering::Relaxed))
}

/// Set the limit on embedded packages opened per document, for every
/// document read afterwards. `usize::MAX` disables the bound.
pub fn set_max_embedded_packages(packages: usize) {
    MAX_EMBEDDED_PACKAGES.store(packages, Ordering::Relaxed);
}

/// The total bytes the embedded packages of one document may take, counting
/// their sizes and what reading them inflates.
pub(crate) fn max_embedded_bytes() -> u64 {
    configured_package_bytes() / 4
}

/// The default per-document text budget: 256 Mi characters.
///
/// The budget scales *down* with a small input. A document opened through
/// [`crate::Document`] knows the size of the file it came from, and its text is
/// bounded by the larger of [`TEXT_BUDGET_FLOOR_CHARS`] (64 Mi characters, which
/// no genuine small document comes near) and [`TEXT_BUDGET_CHARS_PER_INPUT_BYTE`]
/// (4,096) characters per byte of input, and never by more than this limit. So
/// a file under 64 KB cannot produce more than the floor, where before it could
/// produce all 256 Mi characters; a larger file is bounded by this limit as
/// before. Raising or lowering the limit with [`set_max_text_chars`] still
/// sets the ceiling. The readers of a single format used directly (for instance
/// `XlsxDocument::plain_text`) do not know the file's size and use the limit
/// as it is.
pub const DEFAULT_MAX_TEXT_CHARS: usize = 256 << 20;

/// The least a small input is allowed to produce: 64 Mi characters.
#[doc(hidden)]
pub const TEXT_BUDGET_FLOOR_CHARS: usize = 64 << 20;

/// The characters an input may produce for each of its bytes before the limit
/// ([`DEFAULT_MAX_TEXT_CHARS`]) applies: 4,096.
#[doc(hidden)]
pub const TEXT_BUDGET_CHARS_PER_INPUT_BYTE: u64 = 4096;

thread_local! {
    /// The size of the input file of the document being opened or converted
    /// on this thread (see [`InputSizeScope`]).
    static INPUT_BYTES: std::cell::Cell<Option<u64>> = const { std::cell::Cell::new(None) };
}

/// While alive, text budgets created on this thread scale with an input of
/// `bytes` bytes. Restores what it replaced when dropped.
#[must_use = "the scale ends as soon as the scope is dropped; bind it for the length of the conversion"]
pub(crate) struct InputSizeScope(Option<u64>);

impl InputSizeScope {
    /// An input of `bytes` bytes; `0` is an unknown size and changes nothing.
    pub(crate) fn new(bytes: u64) -> Self {
        let previous = INPUT_BYTES.with(|c| c.get());
        if bytes > 0 {
            INPUT_BYTES.with(|c| c.set(Some(bytes)));
        }
        Self(previous)
    }
}

impl Drop for InputSizeScope {
    fn drop(&mut self) {
        INPUT_BYTES.with(|c| c.set(self.0));
    }
}

/// The length of `reader`'s data (0 when it cannot be told), leaving the
/// position where it was.
pub(crate) fn stream_len<R: std::io::Seek>(reader: &mut R) -> u64 {
    let here = reader.stream_position().unwrap_or(0);
    let end = reader.seek(std::io::SeekFrom::End(0)).unwrap_or(0);
    let _ = reader.seek(std::io::SeekFrom::Start(here));
    end
}

/// The text budget for an input of `input` bytes (when known) under the
/// ceiling `limit`.
fn text_limit_for_input(limit: usize, input: Option<u64>) -> usize {
    // No bound is no bound, whatever the size of the input.
    let Some(bytes) = input.filter(|_| limit != usize::MAX) else {
        return limit;
    };
    let scaled = bytes.saturating_mul(TEXT_BUDGET_CHARS_PER_INPUT_BYTE);
    let scaled = usize::try_from(scaled).unwrap_or(usize::MAX);
    limit.min(scaled.max(TEXT_BUDGET_FLOOR_CHARS))
}

/// The default memory budget for what a small file can multiply: 1 GiB.
///
/// The one knob behind every bound on copies and padding, in the manner of
/// `image::Limits::max_alloc`. Each internal bound is a fixed fraction of it
/// (pictures copied: an eighth; their total: four times; a compressed metafile
/// inflated: a quarter; the empty cells padding a sparse sheet: a 64th in cells
/// of 64 bytes; the IR's padding: half, a whole sheet's: all; what a writer
/// builds: half), so that the defaults are the values this crate was
/// tested with and lowering the one number lowers them together.
pub const DEFAULT_MAX_MEMORY_BYTES: u64 = 1 << 30;

static MAX_MEMORY_BYTES: AtomicU64 = AtomicU64::new(DEFAULT_MAX_MEMORY_BYTES);

/// The memory budget for what a small file can multiply (see
/// [`DEFAULT_MAX_MEMORY_BYTES`]).
pub fn max_memory_bytes() -> u64 {
    scoped(|l| l.memory_bytes).unwrap_or_else(|| MAX_MEMORY_BYTES.load(Ordering::Relaxed))
}

/// Set the memory budget for every document read or written afterwards.
/// `u64::MAX` disables the bounds derived from it.
pub fn set_max_memory_bytes(bytes: u64) {
    MAX_MEMORY_BYTES.store(bytes, Ordering::Relaxed);
}

/// The package-wide byte limit as configured, before an embedded package's
/// share of its document's allowance lowers it.
fn configured_package_bytes() -> u64 {
    scoped(|l| l.package_bytes).unwrap_or_else(|| MAX_PACKAGE_BYTES.load(Ordering::Relaxed))
}

static MAX_TEXT_CHARS: AtomicUsize = AtomicUsize::new(DEFAULT_MAX_TEXT_CHARS);

/// The per-document text budget in characters, the amount of cell text a
/// spreadsheet may render before output is truncated with a notice.
///
/// For a package embedded in a document this is what the document's
/// other embedded packages leave of it, if that is less.
pub fn max_text_chars() -> usize {
    let configured = configured_max_text_chars();
    crate::core::embedded::text_cap().map_or(configured, |cap| configured.min(cap))
}

/// The text budget as configured, whatever has been spent of it by the
/// packages embedded in the document being read.
pub(crate) fn configured_max_text_chars() -> usize {
    scoped(|l| l.text_chars).unwrap_or_else(|| MAX_TEXT_CHARS.load(Ordering::Relaxed))
}

/// Set the per-document text budget for every document read afterwards.
/// `usize::MAX` disables the bound.
pub fn set_max_text_chars(chars: usize) {
    MAX_TEXT_CHARS.store(chars, Ordering::Relaxed);
}

/// The per-document allowance of empty slots padding spreadsheet rows.
pub(crate) fn max_grid_padding_cells() -> usize {
    usize::try_from(max_memory_bytes() / 64).unwrap_or(usize::MAX)
}

/// The per-sheet allowance of empty slots padding spreadsheet rows.
pub(crate) fn max_sheet_grid_padding_cells() -> usize {
    usize::try_from(max_memory_bytes() / 128).unwrap_or(usize::MAX)
}

/// The per-document allowance of bytes a legacy file's compressed metafile
/// pictures may inflate to.
pub(crate) fn max_inflated_image_bytes() -> u64 {
    max_memory_bytes() / 4
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

/// A remaining allowance of some unit (empty grid slots, bytes of IR padding):
/// the one concept behind the document-wide and the per-sheet allowances,
/// which are two instances of it. A refusal spends what was left, so every
/// later nonzero charge is refused too.
#[derive(Debug, Clone, Copy)]
pub(crate) struct Allowance {
    total: usize,
    remaining: usize,
    exhausted: bool,
}

impl Allowance {
    pub(crate) fn new(total: usize) -> Self {
        Self {
            total,
            remaining: total,
            exhausted: false,
        }
    }

    /// The amount this allowance started with.
    pub(crate) fn total(&self) -> usize {
        self.total
    }

    /// Give a new owner (the next sheet) the whole allowance again.
    pub(crate) fn refill(&mut self) {
        *self = Self::new(self.total);
    }

    /// What is left after charging `amount`, or `None` when it does not fit
    /// or the allowance is already spent. Charges nothing.
    fn left_after(&self, amount: usize) -> Option<usize> {
        if self.exhausted {
            None
        } else {
            self.remaining.checked_sub(amount)
        }
    }

    /// Charge `amount`; `false` (and the allowance is spent) when it does not
    /// fit. A charge of nothing always succeeds.
    pub(crate) fn charge(&mut self, amount: usize) -> bool {
        if amount == 0 {
            return true;
        }
        match self.left_after(amount) {
            Some(rest) => {
                self.remaining = rest;
                true
            },
            None => {
                self.spend();
                false
            },
        }
    }

    fn spend(&mut self) {
        self.remaining = 0;
        self.exhausted = true;
    }

    pub(crate) fn is_spent(&self) -> bool {
        self.exhausted
    }
}

/// How many cut sheets a notice names before it counts the rest.
const GRID_CUTS_NAMED: usize = 5;

/// A document's remaining allowance of empty grid slots: a total for the
/// document, and one for the sheet being read.
#[derive(Debug, Clone)]
pub(crate) struct GridBudget {
    document: Allowance,
    sheet: Allowance,
    /// Name of the sheet being read.
    sheet_name: String,
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
            document: Allowance::new(total),
            sheet: Allowance::new(per_sheet),
            sheet_name: String::new(),
            cuts: Vec::new(),
        }
    }

    /// Start a sheet: it gets its own allowance, within what is left of the
    /// document's.
    pub(crate) fn begin_sheet(&mut self, name: &str) {
        self.sheet.refill();
        self.sheet_name.clear();
        self.sheet_name.push_str(name);
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
        if self.document.is_spent() || self.sheet.is_spent() {
            return false;
        }
        match (self.document.left_after(slots), self.sheet.left_after(slots)) {
            (Some(_), Some(_)) => {
                self.document.charge(slots);
                self.sheet.charge(slots);
                true
            },
            (document_rest, _) => {
                // The document's allowance is what ran out when it cannot
                // pay; otherwise the sheet's own.
                let by_document = document_rest.is_none();
                if by_document {
                    self.document.spend();
                } else {
                    self.sheet.spend();
                }
                self.cuts.push((self.sheet_name.clone(), by_document));
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
                 (office_oxide::limits::set_max_memory_bytes) — later rows of \
                 {} are missing]",
                Self::name_list(&by_sheet),
                self.sheet.total(),
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
                 (office_oxide::limits::set_max_memory_bytes) — rows from sheet '{name}' \
                 on are missing]",
                self.document.total()
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

/// The per-document allowance of bytes for the empty cells padding the rows
/// of a spreadsheet's IR.
pub(crate) fn max_ir_padding_bytes() -> usize {
    usize::try_from(max_memory_bytes() / 2).unwrap_or(usize::MAX)
}

/// The per-sheet allowance of bytes for the empty cells padding the rows of
/// a spreadsheet sheet in the IR, when sheets are converted one at a time.
pub(crate) fn max_sheet_ir_padding_bytes() -> usize {
    usize::try_from(max_memory_bytes()).unwrap_or(usize::MAX)
}

/// A remaining allowance of bytes for IR padding cells: the document's, or
/// one sheet's.
#[derive(Debug, Clone, Copy)]
pub(crate) struct IrPaddingBudget {
    allowance: Allowance,
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
            allowance: Allowance::new(total),
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
        match cells.checked_mul(Self::cell_bytes()) {
            Some(bytes) => self.allowance.charge(bytes),
            None => {
                self.allowance.spend();
                false
            },
        }
    }

    #[cfg(test)]
    pub(crate) fn exhausted(&self) -> bool {
        self.allowance.is_spent()
    }

    pub(crate) fn notice(&self) -> String {
        let total = self.allowance.total();
        if self.per_sheet {
            return format!(
                "[sheet truncated: its rows are padded to the sheet's width with more than {total} \
                 bytes of empty table cells (office_oxide::limits::set_max_memory_bytes) \
                 — later rows of this sheet are missing]"
            );
        }
        format!(
            "[workbook truncated: its rows are padded to the sheet's width with more than {total} \
             bytes of empty table cells (office_oxide::limits::set_max_memory_bytes) — \
             this sheet's later rows, and the rows of each later sheet that needs padding, are \
             missing]"
        )
    }
}

/// The bytes of picture data a spreadsheet's IR may hold, counting a picture
/// once for every anchor that draws it: the workbook stores a media part once,
/// and the IR gives each image its own copy.
pub(crate) fn max_image_bytes() -> usize {
    usize::try_from(max_memory_bytes() / 8).unwrap_or(usize::MAX)
}

/// The most picture bytes all the sheets of one workbook may copy into the IR
/// in a sheet-by-sheet conversion.
pub(crate) fn max_total_image_bytes() -> u64 {
    max_memory_bytes().saturating_mul(4)
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
                 budget (office_oxide::limits::set_max_memory_bytes)",
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
             (office_oxide::limits::set_max_memory_bytes)]",
                max_image_bytes()
            )
        })
}

/// What a table cell costs the `.docx` writer, in bytes (measured 1,832).
#[doc(hidden)]
pub const WRITE_BYTES_PER_DOCX_CELL: u64 = 1_900;

/// What a table cell costs the `.pptx` writer, in bytes (measured 632).
#[doc(hidden)]
pub const WRITE_BYTES_PER_PPTX_CELL: u64 = 650;

/// What a worksheet cell costs the `.xlsx` writer, in bytes (measured 161).
#[doc(hidden)]
pub const WRITE_BYTES_PER_XLSX_CELL: u64 = 170;

/// The most memory, in bytes, the cells (nested tables included, each table
/// padded to its widest row) of an IR written to `.docx`, `.pptx` or `.xlsx`
/// may be estimated to need before the write is refused.
pub(crate) fn max_write_bytes() -> u64 {
    max_memory_bytes() / 2
}

/// The default most worksheets an IR written to `.xlsx` may hold: 32,767.
///
/// Every sheet is a part of its own with a relationships part beside it, so
/// this keeps the package under the 65,535 entries a zip without ZIP64
/// extensions can hold. Excel itself has no fixed sheet limit (memory
/// decides), but a genuine workbook has tens or hundreds of sheets, not tens
/// of thousands.
pub(crate) const DEFAULT_MAX_WRITE_SHEETS: usize = 32_767;

/// The most sections (worksheets) an IR written to `.xlsx` may hold before
/// the write is refused.
pub(crate) fn max_write_sheets() -> usize {
    DEFAULT_MAX_WRITE_SHEETS
}

/// The most list items (nested ones included) `DocumentIR::from_markdown`
/// keeps; the rest are dropped with a notice.
pub(crate) fn max_markdown_list_items() -> usize {
    100_000
}

/// The most table cells `DocumentIR::from_markdown` keeps, over all of a
/// document's tables.
pub(crate) fn max_markdown_table_cells() -> usize {
    150_000
}

/// The most blocks (sections, paragraphs, headings, code blocks, breaks,
/// lists and tables) `DocumentIR::from_markdown`
/// keeps; the rest of the document is dropped with a notice.
pub(crate) fn max_markdown_elements() -> usize {
    100_000
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
        Self::with_limit(text_limit_for_input(max_text_chars(), INPUT_BYTES.with(|c| c.get())))
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

    /// Whether `chars` would be paid for now, charging nothing.
    pub(crate) fn fits(&self, chars: usize) -> bool {
        !self.exhausted && chars <= self.remaining
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

    #[test]
    fn test_a_small_input_gets_the_floor_and_a_large_one_the_limit() {
        let limit = DEFAULT_MAX_TEXT_CHARS;
        // Unknown size: the limit as it is.
        assert_eq!(text_limit_for_input(limit, None), limit);
        // A tiny input gets the floor, not the whole limit.
        assert_eq!(text_limit_for_input(limit, Some(1)), TEXT_BUDGET_FLOOR_CHARS);
        assert_eq!(text_limit_for_input(limit, Some(10_000)), TEXT_BUDGET_FLOOR_CHARS);
        // 64 KB is where 4,096 characters a byte reaches the limit.
        let at = limit as u64 / TEXT_BUDGET_CHARS_PER_INPUT_BYTE;
        assert_eq!(text_limit_for_input(limit, Some(at)), limit);
        assert_eq!(text_limit_for_input(limit, Some(at * 1000)), limit);
        // Between the two it scales.
        let mid = TEXT_BUDGET_FLOOR_CHARS as u64 / TEXT_BUDGET_CHARS_PER_INPUT_BYTE * 2;
        assert_eq!(
            text_limit_for_input(limit, Some(mid)),
            mid as usize * TEXT_BUDGET_CHARS_PER_INPUT_BYTE as usize
        );
        // A configured ceiling below the floor is the ceiling.
        assert_eq!(text_limit_for_input(1000, Some(1)), 1000);
        // No overflow with an absurd size or limit.
        assert_eq!(text_limit_for_input(usize::MAX, Some(u64::MAX)), usize::MAX);
        // No bound is no bound, even for a tiny input.
        assert_eq!(text_limit_for_input(usize::MAX, Some(1)), usize::MAX);
    }

    #[test]
    fn test_the_input_scope_applies_to_budgets_and_restores() {
        let _lock = test_support::LimitScope::new();
        assert_eq!(TextBudget::new().total, max_text_chars());
        {
            let _scope = InputSizeScope::new(100);
            assert_eq!(TextBudget::new().total, max_text_chars().min(TEXT_BUDGET_FLOOR_CHARS));
            {
                let _unknown = InputSizeScope::new(0);
                assert_eq!(TextBudget::new().total, max_text_chars().min(TEXT_BUDGET_FLOOR_CHARS));
            }
        }
        assert_eq!(TextBudget::new().total, max_text_chars());
    }

    #[test]
    fn test_an_allowance_spends_what_is_left_on_a_refusal_and_refills_for_the_next_owner() {
        let mut a = Allowance::new(10);
        assert!(a.charge(4) && a.charge(0) && a.charge(6));
        assert!(!a.charge(1), "nothing is left");
        assert!(a.charge(0), "a charge of nothing is always allowed");
        let mut b = Allowance::new(10);
        assert!(!b.charge(11), "a charge that does not fit is refused");
        assert!(!b.charge(1), "and spends the rest, so a smaller one is refused too");
        assert!(b.is_spent());
        b.refill();
        assert!(!b.is_spent() && b.charge(10));
    }

    #[test]
    fn test_both_padding_budgets_are_the_same_allowance_at_two_scales() {
        // Grid slots: the sheet's own allowance ends before the document's.
        let mut grid = GridBudget::with_limits(100, 30);
        grid.begin_sheet("a");
        assert!(grid.charge(30) && !grid.charge(1));
        grid.begin_sheet("b");
        assert!(grid.charge(30), "the next sheet has its own allowance");
        // IR padding cells: a charge past the allowance spends it.
        let cells = 3;
        let mut ir = IrPaddingBudget::with_limit(cells * IrPaddingBudget::cell_bytes());
        assert!(ir.charge_cells(cells) && !ir.charge_cells(1) && ir.exhausted());
    }
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
        assert!(b.notice().contains("set_max_memory_bytes"));
        assert!(b.notice().starts_with("[sheet truncated"));
        let d = IrPaddingBudget::new();
        assert!(d.notice().contains("set_max_memory_bytes"));
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
        assert!(notices[0].contains("set_max_memory_bytes"));
        let mut b = GridBudget::with_limits(15, 10);
        b.begin_sheet("A");
        assert!(b.charge(10));
        b.begin_sheet("B");
        assert!(!b.charge(6), "only 5 of the document's allowance is left");
        let notices = b.notices();
        assert!(notices[0].contains("sheet 'B'") && notices[0].contains("set_max_memory_bytes"));
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
        limits.set(max_memory_bytes, set_max_memory_bytes, DEFAULT_MAX_MEMORY_BYTES + 1);
        assert_eq!(max_memory_bytes(), DEFAULT_MAX_MEMORY_BYTES + 1);
        drop(limits);
        assert_eq!(max_memory_bytes(), DEFAULT_MAX_MEMORY_BYTES);
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

    #[test]
    fn test_part_cap_is_the_size_limit_and_past_the_floor_the_ratio_limit() {
        let mut limits = test_support::LimitScope::new();
        // Stored bytes few: the floor governs, so a modest part is allowed.
        assert_eq!(part_inflate_cap(10), INFLATE_RATIO_FLOOR_BYTES);
        // Many stored bytes: the ratio allows more than the floor, up to the size limit.
        let stored = 100 << 10;
        assert_eq!(part_inflate_cap(stored), stored * DEFAULT_MAX_INFLATE_RATIO);
        assert_eq!(part_inflate_cap(u64::MAX), DEFAULT_MAX_PART_BYTES);
        limits.set(max_inflate_ratio, set_max_inflate_ratio, u64::MAX);
        assert_eq!(part_inflate_cap(10), DEFAULT_MAX_PART_BYTES);
        limits.set(max_part_bytes, set_max_part_bytes, 1000);
        assert_eq!(part_inflate_cap(10), 1000);
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

/// Every limit the crate lets you set, as one value, for use per call.
///
/// The `set_max_*` functions change the process-wide defaults, which suits a
/// program that configures the library once. A server reading documents for
/// different tenants, or a test running beside others, wants the limits to
/// belong to the *call*: build a `Limits`, put it in
/// [`ParseOptions`](crate::options::ParseOptions) and read with
/// [`Document::open_with`](crate::Document::open_with). Nothing global
/// changes; other threads are unaffected.
///
/// Eight settings, grouped by what they guard: the zip layer (`part_bytes`,
/// `package_bytes`, `package_entries`, `inflate_ratio`, the same four Apache
/// POI's `ZipSecureFile` has), the output (`text_chars`, as POI's
/// `setMaxTextSize`), embedded packages (`embedded_depth`,
/// `embedded_packages`) and one memory budget (`memory_bytes`, as
/// `image::Limits::max_alloc`) from which every bound on copies, padding and
/// writing is derived. Anything finer is an internal constant.
///
/// [`Limits::default`] holds the documented defaults, [`Limits::current`] the
/// process-wide values now in force, [`Limits::strict`] a tighter set for
/// untrusted input, and [`Limits::unlimited`] switches every bound off.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub struct Limits {
    /// Entries one OOXML package may hold. See [`max_package_entries`].
    pub package_entries: usize,
    /// Decompressed bytes read from one OOXML package (embedded packages read within a quarter of it). See [`max_package_bytes`].
    pub package_bytes: u64,
    /// Decompressed size of one part. See [`max_part_bytes`].
    pub part_bytes: u64,
    /// Decompressed-to-stored ratio of a part larger than 64 MiB. See [`max_inflate_ratio`].
    pub inflate_ratio: u64,
    /// Depth to which embedded packages are opened. See [`max_embedded_depth`].
    pub embedded_depth: usize,
    /// Embedded packages opened for one document. See [`max_embedded_packages`].
    pub embedded_packages: usize,
    /// Characters of text one document may produce. See [`max_text_chars`].
    pub text_chars: usize,
    /// Memory the conversion of one document may spend on what a small file can multiply: pictures copied or inflated, the empty cells that pad a sparse grid, and what a writer builds. Every such bound is a fixed fraction of it. See [`max_memory_bytes`].
    pub memory_bytes: u64,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            package_entries: DEFAULT_MAX_PACKAGE_ENTRIES,
            package_bytes: DEFAULT_MAX_PACKAGE_BYTES,
            part_bytes: DEFAULT_MAX_PART_BYTES,
            inflate_ratio: DEFAULT_MAX_INFLATE_RATIO,
            embedded_depth: DEFAULT_MAX_EMBEDDED_DEPTH,
            embedded_packages: DEFAULT_MAX_EMBEDDED_PACKAGES,
            text_chars: DEFAULT_MAX_TEXT_CHARS,
            memory_bytes: DEFAULT_MAX_MEMORY_BYTES,
        }
    }
}

impl Limits {
    /// The process-wide limits now in force (the defaults, or what the
    /// `set_max_*` functions changed them to).
    pub fn current() -> Self {
        Self {
            package_entries: MAX_PACKAGE_ENTRIES.load(Ordering::Relaxed),
            package_bytes: MAX_PACKAGE_BYTES.load(Ordering::Relaxed),
            part_bytes: MAX_PART_BYTES.load(Ordering::Relaxed),
            inflate_ratio: MAX_INFLATE_RATIO.load(Ordering::Relaxed),
            embedded_depth: MAX_EMBEDDED_DEPTH.load(Ordering::Relaxed),
            embedded_packages: MAX_EMBEDDED_PACKAGES.load(Ordering::Relaxed),
            text_chars: MAX_TEXT_CHARS.load(Ordering::Relaxed),
            memory_bytes: MAX_MEMORY_BYTES.load(Ordering::Relaxed),
        }
    }

    /// A tighter set for input from a source nobody vouches for: an eighth of
    /// each size default, at most one level of embedded packages, and a
    /// part may not inflate more than 100 times its stored size. Documents
    /// of ordinary size read as usual.
    pub fn strict() -> Self {
        let d = Self::default();
        Self {
            package_entries: d.package_entries / 8,
            package_bytes: d.package_bytes / 8,
            part_bytes: d.part_bytes / 8,
            inflate_ratio: 100,
            embedded_depth: 1,
            embedded_packages: d.embedded_packages / 8,
            text_chars: d.text_chars / 8,
            memory_bytes: d.memory_bytes / 8,
        }
    }

    /// No bound at all. Only for input you produced yourself.
    pub fn unlimited() -> Self {
        Self {
            package_entries: usize::MAX,
            package_bytes: u64::MAX,
            part_bytes: u64::MAX,
            inflate_ratio: u64::MAX,
            embedded_depth: usize::MAX,
            embedded_packages: usize::MAX,
            text_chars: usize::MAX,
            memory_bytes: u64::MAX,
        }
    }

    /// Make these the process-wide limits, as the `set_max_*` functions do.
    pub fn apply(&self) {
        MAX_PACKAGE_ENTRIES.store(self.package_entries, Ordering::Relaxed);
        MAX_PACKAGE_BYTES.store(self.package_bytes, Ordering::Relaxed);
        MAX_PART_BYTES.store(self.part_bytes, Ordering::Relaxed);
        MAX_INFLATE_RATIO.store(self.inflate_ratio, Ordering::Relaxed);
        MAX_EMBEDDED_DEPTH.store(self.embedded_depth, Ordering::Relaxed);
        MAX_EMBEDDED_PACKAGES.store(self.embedded_packages, Ordering::Relaxed);
        MAX_TEXT_CHARS.store(self.text_chars, Ordering::Relaxed);
        MAX_MEMORY_BYTES.store(self.memory_bytes, Ordering::Relaxed);
    }
}

impl Limits {
    /// The name of every limit, as [`Limits::set`] takes it: the field names.
    pub const NAMES: &'static [&'static str] = &[
        "package_entries",
        "package_bytes",
        "part_bytes",
        "inflate_ratio",
        "embedded_depth",
        "embedded_packages",
        "text_chars",
        "memory_bytes",
    ];

    /// Set the limit called `name` (a field name; see [`Limits::NAMES`]) to
    /// `value`. For callers that cannot name a Rust field: the command line,
    /// the C interface and the language bindings take limits by name. On a
    /// target where `usize` is narrower than 64 bits a count too large for it
    /// is an error; `u64::MAX` alone means no bound.
    pub fn set(&mut self, name: &str, value: u64) -> Result<(), LimitError> {
        let too_big = usize::try_from(value).is_err() && value != u64::MAX;
        match name {
            "package_entries" => {
                if too_big {
                    return Err(LimitError::TooLarge(name.to_string(), value));
                }
                self.package_entries = usize::try_from(value).unwrap_or(usize::MAX);
            },
            "package_bytes" => self.package_bytes = value,
            "part_bytes" => self.part_bytes = value,
            "inflate_ratio" => self.inflate_ratio = value,
            "embedded_depth" => {
                if too_big {
                    return Err(LimitError::TooLarge(name.to_string(), value));
                }
                self.embedded_depth = usize::try_from(value).unwrap_or(usize::MAX);
            },
            "embedded_packages" => {
                if too_big {
                    return Err(LimitError::TooLarge(name.to_string(), value));
                }
                self.embedded_packages = usize::try_from(value).unwrap_or(usize::MAX);
            },
            "text_chars" => {
                if too_big {
                    return Err(LimitError::TooLarge(name.to_string(), value));
                }
                self.text_chars = usize::try_from(value).unwrap_or(usize::MAX);
            },
            "memory_bytes" => self.memory_bytes = value,
            _ => return Err(LimitError::Unknown(name.to_string())),
        }
        Ok(())
    }
}

/// Why [`Limits::set`] refused a limit.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum LimitError {
    /// The name is not a limit.
    Unknown(String),
    /// The count does not fit this target's `usize` (only `u64::MAX` means no
    /// bound).
    TooLarge(String, u64),
}

impl std::fmt::Display for LimitError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Unknown(name) => {
                write!(f, "unknown limit '{name}'; the limits are: {}", Limits::NAMES.join(", "))
            },
            Self::TooLarge(name, value) => write!(
                f,
                "limit '{name}': {value} is more than this target can count; use the maximum \
                 (u64::MAX) for no bound"
            ),
        }
    }
}

impl std::error::Error for LimitError {}

thread_local! {
    /// The limits of the call running on this thread (see [`LimitsScope`]).
    static SCOPED: std::cell::RefCell<Option<std::rc::Rc<Limits>>> =
        const { std::cell::RefCell::new(None) };
}

/// Read one field of the limits set for this call, if one is set.
fn scoped<T>(field: impl FnOnce(&Limits) -> T) -> Option<T> {
    SCOPED
        .try_with(|c| c.borrow().as_ref().map(|l| field(l)))
        .ok()
        .flatten()
}

/// The limits set for the call running on this thread, if it has its own, for
/// work that continues on another thread to carry them over.
pub(crate) fn scope_snapshot() -> Option<Limits> {
    scoped(Limits::clone)
}

/// While alive, the limits getters on this thread return `limits`' values
/// instead of the process-wide ones. Restores what it replaced when dropped.
#[must_use = "the limits apply only while the scope is alive; bind it for the length of the call"]
pub(crate) struct LimitsScope(Option<std::rc::Rc<Limits>>);

impl LimitsScope {
    pub(crate) fn new(limits: Option<&Limits>) -> Self {
        let Some(limits) = limits else {
            return Self(SCOPED.with(|c| c.borrow().clone()));
        };
        let before = SCOPED.with(|c| c.borrow_mut().replace(std::rc::Rc::new(limits.clone())));
        Self(before)
    }
}

impl Drop for LimitsScope {
    fn drop(&mut self) {
        let _ = SCOPED.try_with(|c| *c.borrow_mut() = self.0.take());
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
    #[must_use = "the limits are restored as soon as the scope is dropped; bind it for the length of the test"]
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
