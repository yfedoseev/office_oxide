//! Peak heap of reading and converting a fixed set of documents, each built
//! in code, held under a ceiling.
//!
//! The memory work of a release (pictures kept as stored streams,
//! parts read into buffers of their own size, the `.xls` grid built without
//! a second copy of its cells) is only kept if a regression fails a test;
//! sheets parsed in bounded pieces, which need parts of 32 and 64 MiB, are
//! held in `test_allocation_budget.rs`. Each document here
//! is dominated by one thing a real file has many of: cells, sheets,
//! paragraphs, slides or picture bytes. Each is opened and then converted
//! every way the library converts, and two peaks are checked against a
//! ceiling in bytes per unit of that thing:
//!
//! - the open: the most live at once while opening, above what was live
//!   before;
//! - the conversion: the most live at once above the open document, over
//!   the IR, text, Markdown, HTML, IR JSON and, for pictures, Markdown and
//!   HTML with the images embedded.
//!
//! A counted heap peak is deterministic where resident memory is not, and the
//! ceilings sit about a third above what this source measures, so growth of
//! a third fails while allocator and layout noise does not; tighter where a
//! fix's whole saving is less than a third. Run with
//! `--nocapture` for the table of every measured peak.
//!
//! The counter is a global allocator, so this file is its own test binary,
//! and the measurements run one at a time.

mod common;

use std::alloc::{GlobalAlloc, Layout, System};
use std::io::Cursor;
use std::sync::atomic::{AtomicUsize, Ordering};

use common::ooxml::{PICTURE_LEN, docx, docx_with_body, pptx, xlsx, xlsx_with_sheets};
use office_oxide::ir_render::{HtmlOptions, ImageEmbed, MarkdownOptions};
use office_oxide::options::ParseOptions;
use office_oxide::{Document, DocumentFormat};

struct Counting;
static LIVE: AtomicUsize = AtomicUsize::new(0);
static PEAK: AtomicUsize = AtomicUsize::new(0);

fn grow(by: usize) {
    let now = LIVE.fetch_add(by, Ordering::Relaxed) + by;
    PEAK.fetch_max(now, Ordering::Relaxed);
}

// SAFETY: forwards every call to `System` unchanged and only counts.
unsafe impl GlobalAlloc for Counting {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        grow(layout.size());
        // SAFETY: same layout contract as the caller's.
        unsafe { System.alloc(layout) }
    }
    unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
        LIVE.fetch_sub(layout.size(), Ordering::Relaxed);
        // SAFETY: `ptr` came from `System.alloc` with this layout.
        unsafe { System.dealloc(ptr, layout) }
    }
    unsafe fn realloc(&self, ptr: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        LIVE.fetch_sub(layout.size(), Ordering::Relaxed);
        grow(new_size);
        // SAFETY: forwarded unchanged.
        unsafe { System.realloc(ptr, layout, new_size) }
    }
}

#[global_allocator]
static GLOBAL: Counting = Counting;

/// The most live at once while running `f`, above what was live when it
/// began.
fn peak_during<T>(f: impl FnOnce() -> T) -> (T, usize) {
    let start = LIVE.load(Ordering::SeqCst);
    PEAK.store(start, Ordering::SeqCst);
    let out = f();
    (out, PEAK.load(Ordering::SeqCst).saturating_sub(start))
}

/// One document of the suite.
struct Case {
    name: &'static str,
    format: DocumentFormat,
    bytes: Vec<u8>,
    lazy_pictures: bool,
    /// How many of the thing that dominates the document it holds, and
    /// what that thing is.
    units: usize,
    unit: &'static str,
    /// What this source measures, in bytes per unit: the open's peak (the
    /// higher of the default build and the `parallel` one, whose workers
    /// add to an `.xlsx` open), and the costliest conversion's.
    open_now: f64,
    convert_now: f64,
    /// How far above `*_now` a peak may go before the test fails.
    margin: f64,
}

/// Pictures in the picture documents, and how often each is drawn.
const PICTURES: usize = 2;
const DRAWS: usize = 6;

fn cases() -> Vec<Case> {
    let picture_bytes = PICTURES * PICTURE_LEN;
    let mut cases = Vec::new();

    // A sheet whose cells lie far apart. The IR lays a sheet out as the
    // grid its cells span (bounded by `sheet_grid_padding_cells`), so the
    // unit is a cell of that grid, not a cell the sheet holds.
    let (span_rows, span_cols) = (2_000, 200);
    let sparse: String = (0..200)
        .map(|i| {
            let r = 1 + i * 10;
            format!(r#"<row r="{r}"><c r="{}{r}"><v>{i}</v></c></row>"#, column(i * 7 % span_cols))
        })
        .collect();
    cases.push(Case {
        name: "sparse sheet",
        format: DocumentFormat::Xlsx,
        bytes: xlsx_with_sheets(&[sparse]),
        lazy_pictures: false,
        units: span_rows * span_cols,
        unit: "grid cell",
        open_now: 0.33,
        convert_now: 75.82,
        margin: LOOSE,
    });

    // Many small sheets: per-sheet costs, not per-cell ones.
    let sheets = 100;
    let small: Vec<String> = (0..sheets)
        .map(|s| {
            (1..=10)
                .map(|r| {
                    format!(
                        r#"<row r="{r}"><c r="A{r}"><v>{s}</v></c><c r="B{r}" t="inlineStr"><is><t>sheet {s} row {r}</t></is></c></row>"#
                    )
                })
                .collect()
        })
        .collect();
    cases.push(Case {
        name: "many-sheet workbook",
        format: DocumentFormat::Xlsx,
        bytes: xlsx_with_sheets(&small),
        lazy_pictures: false,
        units: sheets,
        unit: "sheet",
        open_now: 8337.7,
        convert_now: 22255.34,
        margin: LOOSE,
    });

    // A long document of plain paragraphs.
    let paragraphs = 5_000;
    let body: String = (0..paragraphs)
        .map(|i| {
            format!(
                r#"<w:p><w:r><w:t xml:space="preserve">Paragraph {i}, </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>with a bold run.</w:t></w:r></w:p>"#
            )
        })
        .collect();
    cases.push(Case {
        name: "long docx",
        format: DocumentFormat::Docx,
        bytes: docx_with_body(&body),
        lazy_pictures: false,
        units: paragraphs,
        unit: "paragraph",
        open_now: 1100.18,
        convert_now: 982.57,
        margin: LOOSE,
    });

    // Pictures, each drawn many times, read eagerly and lazily.
    for lazy in [false, true] {
        for (name, format, bytes) in [
            ("pictures docx", DocumentFormat::Docx, docx(PICTURES, DRAWS)),
            ("pictures pptx", DocumentFormat::Pptx, pptx(PICTURES, DRAWS)),
            ("pictures xlsx", DocumentFormat::Xlsx, xlsx(PICTURES, DRAWS)),
        ] {
            cases.push(Case {
                name,
                format,
                bytes,
                lazy_pictures: lazy,
                units: picture_bytes,
                unit: "picture byte",
                open_now: if lazy { 0.52 } else { 1.51 },
                // The IR copies each picture in both modes.
                convert_now: match format {
                    DocumentFormat::Docx => 13.19,
                    _ => 11.0,
                },
                margin: LOOSE,
            });
        }
    }

    // A dense legacy sheet, and a sparse one.
    let (rows, cols) = (1_000u16, 20u16);
    cases.push(Case {
        name: "dense xls",
        format: DocumentFormat::Xls,
        bytes: common::xls_with_cells(rows, cols, &["alpha", "beta", "gamma"]),
        lazy_pictures: false,
        units: rows as usize * cols as usize,
        unit: "cell",
        open_now: 87.23,
        convert_now: 998.65,
        // The open of a sheet that held its cell list and its grid at
        // once peaked a third higher, inside the usual margin.
        margin: 1.15,
    });
    // As for the sparse sheet, the unit is a cell of the spanned grid.
    let far: Vec<(u16, u16)> = (0..200u16).map(|i| (i * 10, (i * 7) % 200)).collect();
    cases.push(Case {
        name: "sparse xls",
        format: DocumentFormat::Xls,
        bytes: common::xls_with_cells_at(&far, &["alpha", "beta"]),
        lazy_pictures: false,
        units: 2_000 * 200,
        unit: "grid cell",
        open_now: 1.6,
        convert_now: 38.29,
        margin: LOOSE,
    });

    // A long legacy document, and a deck of many slides.
    let doc_paragraphs = 2_000;
    let paras: Vec<common::Para> = (0..doc_paragraphs)
        .map(|_| common::Para {
            text: "A paragraph of a long legacy document.",
            terminator: '\r',
            grpprl: common::prose_grpprl(),
        })
        .collect();
    cases.push(Case {
        name: "long doc",
        format: DocumentFormat::Doc,
        bytes: {
            let (word, table) = common::build_doc_streams(
                &paras,
                &common::Subdocs::default(),
                common::FibTweaks::default(),
            );
            common::cfb_with_streams(&[("WordDocument", &word), ("0Table", &table)])
        },
        lazy_pictures: false,
        units: doc_paragraphs,
        unit: "paragraph",
        open_now: 2437.04,
        convert_now: 1006.09,
        margin: LOOSE,
    });
    let slides = 1_000;
    let deck = common::ppt::PptBuilder {
        slides: (0..slides)
            .map(|i| common::ppt::PptSlide {
                shapes: common::ppt::text_shape(1, &format!("Slide {i} body text"), &[]),
                ..Default::default()
            })
            .collect(),
        ..Default::default()
    };
    cases.push(Case {
        name: "many-slide ppt",
        format: DocumentFormat::Ppt,
        bytes: deck.build(),
        lazy_pictures: false,
        units: slides,
        unit: "slide",
        open_now: 1203.39,
        convert_now: 2162.1,
        margin: LOOSE,
    });
    cases
}

/// The spreadsheet column name of the 0-based column `i`.
fn column(mut i: usize) -> String {
    let mut name = Vec::new();
    loop {
        name.push(b'A' + (i % 26) as u8);
        if i < 26 {
            break;
        }
        i = i / 26 - 1;
    }
    name.reverse();
    String::from_utf8(name).unwrap()
}

/// The open peak, what the open document keeps, and each conversion's peak
/// above the open document.
fn measure(case: &Case) -> (usize, usize, Vec<(&'static str, usize)>) {
    let mut options = ParseOptions::default();
    options.lazy_pictures = case.lazy_pictures;
    let before = LIVE.load(Ordering::SeqCst);
    let (doc, open) = peak_during(|| {
        Document::from_reader_with(Cursor::new(case.bytes.clone()), case.format, &options)
            .unwrap_or_else(|e| panic!("{}: {e}", case.name))
    });
    let kept = LIVE.load(Ordering::SeqCst).saturating_sub(before);
    let embedded = |image_embed| (MarkdownOptions { image_embed }, HtmlOptions { image_embed });
    let (md_images, html_images) = embedded(ImageEmbed::Base64);
    let mut stages: Vec<(&'static str, usize)> = vec![
        ("ir", peak_during(|| doc.to_ir()).1),
        ("text", peak_during(|| doc.plain_text()).1),
        ("markdown", peak_during(|| doc.to_markdown()).1),
        ("html", peak_during(|| doc.to_html()).1),
        ("ir json", peak_during(|| doc.to_ir().to_json_string().unwrap()).1),
    ];
    if case.unit == "picture byte" {
        stages.push(("markdown+images", peak_during(|| doc.to_markdown_with(md_images)).1));
        stages.push(("html+images", peak_during(|| doc.to_html_with(html_images)).1));
    }
    (open, kept, stages)
}

/// The usual margin: a third above what this source measures.
const LOOSE: f64 = 4.0 / 3.0;

#[test]
fn test_reading_and_converting_each_document_stays_under_its_peak_heap_ceiling() {
    let _serial = common::LimitScope::new();
    let mut failures = Vec::new();
    println!(
        "{:<22} {:>5} {:>10} {:>12} {:>9} {:>12} {:>12} {:>9}  costliest",
        "document", "lazy", "input", "open peak", "/unit", "kept", "convert peak", "/unit"
    );
    for case in cases() {
        let (open, kept, stages) = measure(&case);
        let (stage, convert) = stages.iter().copied().max_by_key(|&(_, b)| b).unwrap();
        let per = |b: usize| b as f64 / case.units.max(1) as f64;
        println!(
            "{:<22} {:>5} {:>10} {:>12} {:>9.2} {:>12} {:>12} {:>9.2}  {stage} (per {})",
            case.name,
            case.lazy_pictures,
            case.bytes.len(),
            open,
            per(open),
            kept,
            convert,
            per(convert),
            case.unit
        );
        if per(open) > case.open_now * case.margin {
            failures.push(format!(
                "{} (lazy {}): open peaked at {:.2} bytes per {}, ceiling {:.2}",
                case.name,
                case.lazy_pictures,
                per(open),
                case.unit,
                case.open_now * case.margin
            ));
        }
        if per(convert) > case.convert_now * case.margin {
            failures.push(format!(
                "{} (lazy {}): {stage} peaked at {:.2} bytes per {}, ceiling {:.2} ({stages:?})",
                case.name,
                case.lazy_pictures,
                per(convert),
                case.unit,
                case.convert_now * case.margin
            ));
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}
