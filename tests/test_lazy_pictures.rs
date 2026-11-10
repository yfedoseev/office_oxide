//! Pictures held as their stored streams (`lazy_pictures`).
//!
//! Opening a document inflated every picture it holds and kept them for its
//! life, although text, Markdown and HTML without images never read them,
//! and the IR copied each picture into every image that drew it. With
//! `ParseOptions::lazy_pictures` a deflated picture is kept as its stored
//! stream and inflated only when used, and the IR shares it. Every output
//! must be the same either way, and a picture that cannot be read must be
//! reported the same way.
//!
//! The memory checks count the heap with a global allocator, so this file is
//! its own test binary.

mod common;

use std::alloc::{GlobalAlloc, Layout, System};
use std::io::Cursor;
use std::sync::atomic::{AtomicUsize, Ordering};

use office_oxide::ir::Element;
use office_oxide::ir_render::{HtmlOptions, ImageEmbed, MarkdownOptions};
use office_oxide::options::ParseOptions;
use office_oxide::{Document, DocumentFormat};

struct Counting;
static LIVE: AtomicUsize = AtomicUsize::new(0);
static PEAK: AtomicUsize = AtomicUsize::new(0);
/// Every byte requested, freed or not.
static TOTAL: AtomicUsize = AtomicUsize::new(0);

fn grow(by: usize) {
    TOTAL.fetch_add(by, Ordering::Relaxed);
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

/// The counter is process-wide: every test holds this for its whole body.
fn serial() -> common::LimitScope {
    common::LimitScope::new()
}

/// What `f` left live, and the most that was live at once, above what was
/// live when it began.
fn measure<T>(f: impl FnOnce() -> T) -> (T, usize, usize) {
    let start = LIVE.load(Ordering::SeqCst);
    PEAK.store(start, Ordering::SeqCst);
    let out = f();
    let live = LIVE.load(Ordering::SeqCst).saturating_sub(start);
    (out, live, PEAK.load(Ordering::SeqCst).saturating_sub(start))
}

use common::ooxml::{
    PICTURE_LEN, docx, docx_with_media_names, picture, pptx, pptx_with_ole_previews, xlsx,
};

/// Pictures per document.
const PICTURES: usize = 4;

fn options(lazy: bool) -> ParseOptions {
    let mut o = ParseOptions::default();
    o.lazy_pictures = lazy;
    o
}

fn open(bytes: &[u8], format: DocumentFormat, lazy: bool) -> Document {
    Document::from_reader_with(Cursor::new(bytes.to_vec()), format, &options(lazy)).unwrap()
}

fn formats() -> [(DocumentFormat, Vec<u8>); 3] {
    [
        (DocumentFormat::Docx, docx(PICTURES, PICTURES)),
        (DocumentFormat::Pptx, pptx(PICTURES, PICTURES)),
        (DocumentFormat::Xlsx, xlsx(PICTURES, PICTURES)),
    ]
}

/// Opening holds the pictures' stored size, not their inflated size, and
/// inflates at most about one picture at a time.
#[test]
fn test_a_lazy_open_holds_the_stored_pictures_not_the_inflated_ones() {
    let _serial = serial();
    let inflated = PICTURES * PICTURE_LEN;
    for (format, bytes) in formats() {
        let (eager, eager_live, _) = measure(|| open(&bytes, format, false));
        assert!(eager_live >= inflated, "{format:?}: an eager open held only {eager_live} bytes");
        drop(eager);
        let (lazy, lazy_live, lazy_peak) = measure(|| open(&bytes, format, true));
        assert!(
            lazy_live < inflated / 8,
            "{format:?}: a lazy open held {lazy_live} bytes of {inflated} inflated"
        );
        // One picture inflated at a time, into a buffer of its own size: a
        // buffer grown by doubling past the declared size held twice that.
        assert!(
            lazy_peak < PICTURE_LEN + PICTURE_LEN / 4,
            "{format:?}: a lazy open peaked at {lazy_peak} bytes (one picture is {PICTURE_LEN})"
        );
        // Text never inflates a picture.
        let (_, _, text_peak) = measure(|| lazy.plain_text());
        assert!(text_peak < PICTURE_LEN, "{format:?}: plain_text peaked at {text_peak}");
    }
}

/// Every output is byte for byte the same whether pictures are lazy or not.
#[test]
fn test_every_output_is_the_same_with_lazy_pictures() {
    let _serial = serial();
    for (format, bytes) in formats() {
        assert_same_outputs(format, &bytes, &ParseOptions::default(), "plain");
        let md = open(&bytes, format, true).to_markdown_with(MarkdownOptions {
            image_embed: ImageEmbed::Base64,
        });
        assert!(md.contains("[image-base64:iVBORw0KGgo"), "{format:?}: images embedded");
    }
    use common::ooxml::{pptx_with_slides, xlsx_with_picture_sheets};
    assert_same_outputs(
        DocumentFormat::Pptx,
        &pptx_with_slides(5, 2),
        &ParseOptions::default(),
        "slides",
    );
    assert_same_outputs(
        DocumentFormat::Xlsx,
        &xlsx_with_picture_sheets(5, 2),
        &ParseOptions::default(),
        "sheets",
    );
    assert_same_outputs(
        DocumentFormat::Pptx,
        &pptx_with_ole_previews(3),
        &ParseOptions::default(),
        "ole",
    );
}

/// Every output of `bytes` read under `base` with lazy pictures off and on.
fn assert_same_outputs(format: DocumentFormat, bytes: &[u8], base: &ParseOptions, what: &str) {
    let read = |lazy: bool| {
        let mut o = base.clone();
        o.lazy_pictures = lazy;
        Document::from_reader_with(Cursor::new(bytes.to_vec()), format, &o).unwrap()
    };
    let (eager, lazy) = (read(false), read(true));
    let md = |d: &Document| {
        d.to_markdown_with(MarkdownOptions {
            image_embed: ImageEmbed::Base64,
        })
    };
    let html = |d: &Document| {
        d.to_html_with(HtmlOptions {
            image_embed: ImageEmbed::Base64,
        })
    };
    assert_eq!(eager.plain_text(), lazy.plain_text(), "{what} {format:?} text");
    assert_eq!(eager.to_markdown(), lazy.to_markdown(), "{what} {format:?} markdown");
    assert_eq!(md(&eager), md(&lazy), "{what} {format:?} markdown with images");
    assert_eq!(eager.to_html(), lazy.to_html(), "{what} {format:?} html");
    assert_eq!(html(&eager), html(&lazy), "{what} {format:?} html with images");
    let (e, l) = (eager.to_ir(), lazy.to_ir());
    assert_eq!(
        e.to_json_string().unwrap(),
        l.to_json_string().unwrap(),
        "{what} {format:?} IR JSON"
    );
    assert_eq!(eager.damage(), lazy.damage(), "{what} {format:?} damage");
}

/// The budget for picture bytes copied into one IR is charged the same in
/// both modes, so a budget too small for every picture keeps the same ones
/// and reports the same cut.
#[test]
fn test_an_image_budget_too_small_for_every_picture_cuts_the_same_lazy_or_not() {
    let _serial = serial();
    let mut small = ParseOptions::default();
    // Half of `memory_bytes` is the IR's picture budget: room for two of
    // the four 2 MiB pictures.
    small.limits.memory_bytes = (5 * PICTURE_LEN) as u64;
    for (format, bytes) in formats() {
        assert_same_outputs(format, &bytes, &small, "budget");
        let mut o = small.clone();
        o.lazy_pictures = true;
        let doc = Document::from_reader_with(Cursor::new(bytes.clone()), format, &o).unwrap();
        let mut found = Vec::new();
        for section in &doc.to_ir().sections {
            images(&section.elements, &mut found);
        }
        let kept = found.iter().filter(|i| i.has_bytes()).count();
        assert!(kept < PICTURES, "{format:?}: the budget kept all {kept} pictures");
    }
}

/// A picture whose zip entry fails its checksum is read as stored and
/// reported, the same lazy or not, and its bytes are the bytes stored.
#[test]
fn test_a_picture_failing_its_checksum_reads_the_same_lazy_or_not() {
    let _serial = serial();
    let mut bytes = docx(1, 1);
    // Flip the CRC-32 of the picture's entry in its local header and in the
    // central directory.
    let name = b"word/media/img0.png";
    let mut at = 0;
    while let Some(i) = bytes[at..].windows(name.len()).position(|w| w == name) {
        let i = at + i;
        if i >= 30 && bytes[i - 30..i - 26] == [0x50, 0x4b, 0x03, 0x04] {
            bytes[i - 30 + 14] ^= 0xFF;
        } else if i >= 46 && bytes[i - 46..i - 42] == [0x50, 0x4b, 0x01, 0x02] {
            bytes[i - 46 + 16] ^= 0xFF;
        }
        at = i + 1;
    }
    assert_same_outputs(DocumentFormat::Docx, &bytes, &ParseOptions::default(), "crc");
    let lazy = open(&bytes, DocumentFormat::Docx, true);
    assert!(!lazy.damage().is_empty(), "the checksum is reported");
    let (pic, _) = lazy.as_docx().unwrap().picture("rIdP0").unwrap();
    assert_eq!(&pic.bytes().unwrap()[..], &picture(0)[..]);
}

fn images(elements: &[Element], out: &mut Vec<office_oxide::ir::Image>) {
    for e in elements {
        match e {
            Element::Image(i) => out.push(i.clone()),
            Element::TextBox(t) => images(&t.content, out),
            _ => {},
        }
    }
}

/// One picture drawn many times is one allocation in a lazy IR, where an
/// eager IR holds a copy per drawing.
#[test]
fn test_a_lazy_ir_shares_a_picture_drawn_many_times() {
    let _serial = serial();
    for (format, bytes) in [
        (DocumentFormat::Docx, docx(1, 2)),
        (DocumentFormat::Pptx, pptx(1, 2)),
        (DocumentFormat::Xlsx, xlsx(1, 2)),
    ] {
        let lazy = open(&bytes, format, true);
        let mut found = Vec::new();
        for s in &lazy.to_ir().sections {
            images(&s.elements, &mut found);
        }
        assert_eq!(found.len(), 2, "{format:?}");
        let (a, b) = (found[0].shared_data.as_ref(), found[1].shared_data.as_ref());
        assert!(found.iter().all(|i| i.data.is_none()), "{format:?}: no copies");
        assert!(a.unwrap().ptr_eq(b.unwrap()), "{format:?}: one allocation");
        assert_eq!(found[0].byte_len(), Some(PICTURE_LEN));
        assert_eq!(&found[0].bytes().unwrap()[..], &picture(0)[..]);
        let eager = open(&bytes, format, false);
        let mut copies = Vec::new();
        for s in &eager.to_ir().sections {
            images(&s.elements, &mut copies);
        }
        assert!(
            copies
                .iter()
                .all(|i| i.shared_data.is_none() && i.data.is_some())
        );
    }
}

/// The readers' own accessors give the pictures in both modes.
#[test]
fn test_the_readers_give_their_pictures_lazy_or_not() {
    let _serial = serial();
    let want = picture(0);
    for lazy in [false, true] {
        let d = open(&docx(1, 1), DocumentFormat::Docx, lazy);
        let (pic, ext) = d.as_docx().unwrap().picture("rIdP0").unwrap();
        assert_eq!(&pic.bytes().unwrap()[..], &want[..]);
        assert_eq!(ext, Some("png"));
        assert_eq!(pic.is_deflated(), lazy);

        let p = open(&pptx(1, 1), DocumentFormat::Pptx, lazy);
        let pic = p.as_pptx().unwrap().slide_picture(0, "rIdP0").unwrap();
        assert_eq!(&pic.bytes().unwrap()[..], &want[..]);

        let x = open(&xlsx(1, 1), DocumentFormat::Xlsx, lazy);
        let book = x.as_xlsx().unwrap();
        let ws = &book.worksheets[0];
        assert_eq!(ws.images.len(), 1);
        assert_eq!(&book.picture(&ws.images[0]).bytes().unwrap()[..], &want[..]);
    }
}

/// A picture over the part limit is refused and reported the same way.
#[test]
fn test_a_picture_over_a_limit_is_reported_the_same_lazy_or_not() {
    let _serial = serial();
    for (format, bytes) in formats() {
        let damage = |lazy: bool| {
            let mut o = options(lazy);
            o.limits.part_bytes = (PICTURE_LEN / 2) as u64;
            let d = Document::from_reader_with(Cursor::new(bytes.clone()), format, &o).unwrap();
            (d.damage(), d.plain_text())
        };
        let (eager, lazy) = (damage(false), damage(true));
        assert!(!eager.0.is_empty(), "{format:?}: the oversized pictures are reported");
        assert_eq!(eager, lazy, "{format:?}");
    }
}

/// An OLE object's preview is held inflated, but a picture that previews
/// many objects is inflated once, as an eager open shares it: a lazy picture
/// inflates afresh on every read, and a frame each held a copy each.
#[test]
fn test_a_picture_previewing_many_ole_objects_is_inflated_once() {
    let _serial = serial();
    const FRAMES: usize = 50;
    let bytes = pptx_with_ole_previews(FRAMES);
    let (lazy, live, _) = measure(|| open(&bytes, DocumentFormat::Pptx, true));
    assert!(
        live < 2 * PICTURE_LEN,
        "a lazy open of {FRAMES} previews of one picture held {live} bytes (one is {PICTURE_LEN})"
    );
    let eager = open(&bytes, DocumentFormat::Pptx, false);
    assert_eq!(lazy.plain_text(), eager.plain_text());
    assert_eq!(lazy.to_html(), eager.to_html());
    assert_eq!(lazy.to_ir().to_json_string().unwrap(), eager.to_ir().to_json_string().unwrap());
}

/// HTML with images inline tells an image's type from its bytes when its
/// part has no extension, and reads the bytes once for both: a lazy picture
/// inflates on every read, so sniffing the type inflated it a second time.
#[test]
fn test_html_with_images_reads_a_picture_of_unknown_type_once() {
    let _serial = serial();
    let bytes = docx_with_media_names(1, 1, "");
    let options = HtmlOptions {
        image_embed: ImageEmbed::Base64,
    };
    let allocated = |doc: &Document| {
        let before = TOTAL.load(Ordering::SeqCst);
        let html = doc.to_html_with(options);
        (html, TOTAL.load(Ordering::SeqCst) - before)
    };
    let (lazy_html, lazy) = allocated(&open(&bytes, DocumentFormat::Docx, true));
    let (eager_html, eager) = allocated(&open(&bytes, DocumentFormat::Docx, false));
    assert!(lazy_html.contains("data:image/png;base64,"), "the type is told from the bytes");
    assert_eq!(lazy_html, eager_html);
    // A lazy render inflates the picture once more than an eager one.
    let extra = lazy.saturating_sub(eager);
    assert!(
        extra < PICTURE_LEN + PICTURE_LEN / 2,
        "a lazy render allocated {extra} bytes more than an eager one (one picture is {PICTURE_LEN})"
    );
}

/// `slide_picture` answers the same whether pictures are read lazily or
/// not, for any image relationship of the slide: eagerly it found only a
/// picture shape's, so a picture only an OLE object previews was `None`.
#[test]
fn test_slide_picture_answers_the_same_lazy_or_not() {
    let _serial = serial();
    let bytes = pptx_with_ole_previews(1);
    let want = picture(0);
    for lazy in [false, true] {
        let doc = open(&bytes, DocumentFormat::Pptx, lazy);
        let deck = doc.as_pptx().unwrap();
        let pic = deck
            .slide_picture(0, "rIdP0")
            .expect("the preview's picture");
        assert_eq!(&pic.bytes().unwrap()[..], &want[..], "lazy {lazy}");
        assert!(deck.slide_picture(0, "rIdNone").is_none());
        assert!(deck.slide_picture(1, "rIdP0").is_none());
    }
}

/// `XlsxDocument::picture` finds a lazy picture by its anchor's token, and
/// only that token: a caller who drops a sheet's pictures frees their tokens,
/// and an allocation reusing a freed token's address used to be handed that
/// picture's bytes.
#[test]
#[allow(deprecated)]
fn test_a_workbook_picture_is_found_by_its_own_token_only() {
    let mut scope = serial();
    scope.set(
        office_oxide::read_options::lazy_pictures,
        office_oxide::read_options::set_lazy_pictures,
        true,
    );
    let mut book = office_oxide::xlsx::XlsxDocument::from_reader(Cursor::new(xlsx(1, 4))).unwrap();
    let freed: Vec<usize> = book.worksheets[0]
        .images
        .iter()
        .map(|p| p.data.as_ptr() as usize)
        .collect();
    assert_eq!(book.picture(&book.worksheets[0].images[0]).len(), PICTURE_LEN);
    // A picture of the caller's own, holding none of the workbook's tokens.
    let mut stranger = book.worksheets[0].images[0].clone();
    stranger.data = std::sync::Arc::from(vec![7u8]);
    book.worksheets[0].images = Vec::new().into();
    // Allocate empty slices until one lands where a freed token was, or
    // give up: either way it must not be taken for the picture.
    let mut kept = Vec::new();
    for _ in 0..10_000 {
        let token: std::sync::Arc<[u8]> = std::sync::Arc::from(Vec::new());
        let reused = freed.contains(&(token.as_ptr() as usize));
        kept.push(token);
        if reused {
            break;
        }
    }
    for token in &kept {
        stranger.data = token.clone();
        assert_eq!(book.picture(&stranger).len(), 0, "a stranger's bytes are its own");
    }
}
