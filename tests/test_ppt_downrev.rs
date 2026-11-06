//! Shape text in `metroBlob` packages in legacy `.ppt` files.
//!
//! Defect class: a shape can keep its text in a self-contained package (a
//! `drs/shapexml.xml` entry holding a DrawingML `<p:sp>` fragment) hung off its
//! tertiary property table, with the legacy `TextCharsAtom` / `TextBytesAtom`
//! atoms nearly empty, as WPS writes an OOXML-origin deck. A reader that walks
//! only the atoms returns an almost-empty deck with `Ok`. PowerPoint writes
//! the same packages beside full legacy text, which must not come out twice.
//!
//! Fixtures are synthesised in code (see `tests/common`): the ZIP packages
//! are real deflate archives, the record stream is a real MS-PPT structure,
//! and both extraction paths (persist-directory resolution and the
//! no-directory fallback) are exercised.

mod common;

use common::{
    RT_DG_CONTAINER, RT_SPGR_CONTAINER, RT_TEXT_CHARS, broken_downrev_carrier, cfb_with_stream,
    downrev_carrier, downrev_carrier_short_length, downrev_carrier_with_decoy,
    downrev_group_from_carrier, downrev_shape_group, downrev_shape_xml, downrev_zip,
    non_shape_downrev_carrier, ppt_container, ppt_record, ppt_stream_with_persist,
    ppt_stream_without_persist,
};

use office_oxide::ppt::TextType;
use office_oxide::{Document, DocumentFormat};
use std::io::Cursor;

/// UTF-16LE bytes for a `TextCharsAtom` payload ([MS-PPT] 2.9.19).
fn text_chars_payload(text: &str) -> Vec<u8> {
    text.encode_utf16()
        .flat_map(u16::to_le_bytes)
        .collect::<Vec<u8>>()
}

/// Slide 1: a title and a body paragraph, one shape, one carrier (with the
/// 6-byte prefix real WPS files write before the ZIP).
fn slide_one() -> Vec<Vec<u8>> {
    let xml = downrev_shape_xml(&["Quarterly results", "Revenue up 12% year over year"]);
    vec![downrev_shape_group(&downrev_zip(&xml))]
}

/// Slide 2: two shapes, the second a duplicate of the first's text (real
/// files repeat a shape's package across several records), plus a carrier
/// with a malformed ZIP and one with a non-shape package — both must be
/// skipped without disturbing the good text.
fn slide_two() -> Vec<Vec<u8>> {
    let xml = downrev_shape_xml(&["Outlook", "Steady growth expected"]);
    vec![
        downrev_shape_group(&downrev_zip(&xml)),
        downrev_shape_group(&downrev_zip(&xml)), // duplicate package
        downrev_shape_group(&Vec::new()),        // carrier with no ZIP at all
        ppt_container(
            common::RT_SP_CONTAINER,
            0,
            &[broken_downrev_carrier(), non_shape_downrev_carrier()],
        ),
    ]
}

/// The master's shape group — its text must never reach a slide.
fn master_group() -> Vec<u8> {
    let xml = downrev_shape_xml(&["Click to edit Master title style", "Second level"]);
    downrev_shape_group(&downrev_zip(&xml))
}

fn open_ppt(stream: Vec<u8>) -> Document {
    let cfb = cfb_with_stream("PowerPoint Document", &stream);
    Document::from_reader(Cursor::new(cfb), DocumentFormat::Ppt).expect("synthetic ppt opens")
}

/// The persist-directory path: every slide's downrev shape text is
/// extracted, per slide, with duplicates collapsed and master packages
/// excluded.
#[test]
fn test_wps_downrev_shape_text_reaches_the_reader_via_persist_path() {
    let stream = ppt_stream_with_persist(&[slide_one(), slide_two()], &[master_group()]);
    let doc = open_ppt(stream);
    let ppt = doc.as_ppt().expect("parsed as ppt");

    assert_eq!(ppt.slides.len(), 2, "one slide per SlideContainer");

    let first = &ppt.slides[0];
    assert!(
        first
            .text_runs
            .iter()
            .any(|r| r.text.contains("Quarterly results")),
        "slide 1 keeps its downrev title: {:?}",
        first.text_runs
    );
    assert!(
        first
            .text_runs
            .iter()
            .any(|r| r.text.contains("Revenue up 12% year over year")),
        "slide 1 keeps its downrev body: {:?}",
        first.text_runs
    );

    let second = &ppt.slides[1];
    assert!(
        second.text_runs.iter().any(|r| r.text.contains("Outlook")),
        "slide 2 keeps its downrev text: {:?}",
        second.text_runs
    );

    let text = ppt.plain_text();
    assert!(
        !text.contains("Master title style"),
        "master downrev text must not leak into slides: {text:?}"
    );
    assert_eq!(
        text.matches("Steady growth expected").count(),
        1,
        "duplicate downrev packages collapse to one run: {text:?}"
    );
    assert!(
        text.contains("Quarterly results") && text.contains("Outlook"),
        "plain_text surfaces both slides: {text:?}"
    );

    // The IR and markdown views ride on the same runs.
    let ir = doc.to_ir();
    assert_eq!(ir.sections.len(), 2, "one IR section per slide");
    let ir_text = ir.plain_text();
    assert!(
        ir_text.contains("Revenue up 12% year over year"),
        "IR sections carry the downrev text: {ir_text:?}"
    );
    assert!(doc.to_markdown().contains("Outlook"), "markdown surfaces the downrev text");
}

/// The persist path really is the one that ran: a `Slide` container the
/// directory doesn't name, as an earlier save leaves behind, isn't read. The
/// fallback, which walks every container, would read it.
#[test]
fn test_persist_path_reads_only_the_slides_the_directory_names() {
    let mut stream = ppt_stream_with_persist(&[slide_one()], &[master_group()]);
    let stale = downrev_shape_xml(&["Superseded slide"]);
    stream.extend(common::ppt_slide(&[downrev_shape_group(&downrev_zip(&stale))]));
    let text = open_ppt(stream)
        .as_ppt()
        .expect("parsed as ppt")
        .plain_text();
    assert!(text.contains("Quarterly results"), "{text:?}");
    assert!(!text.contains("Superseded slide"), "{text:?}");

    let stream = ppt_stream_without_persist(
        &[slide_one(), vec![downrev_shape_group(&downrev_zip(&stale))]],
        &[master_group()],
    );
    let text = open_ppt(stream)
        .as_ppt()
        .expect("parsed as ppt")
        .plain_text();
    assert!(
        text.contains("Superseded slide"),
        "the fallback reads every container: {text:?}"
    );
}

/// The no-directory fallback path: the same containers, no persist records.
#[test]
fn test_wps_downrev_shape_text_reaches_the_reader_without_persist_path() {
    let stream = ppt_stream_without_persist(&[slide_one(), slide_two()], &[master_group()]);
    let doc = open_ppt(stream);
    let ppt = doc.as_ppt().expect("parsed as ppt");

    assert_eq!(ppt.slides.len(), 2, "one slide per SlideContainer");
    let text = ppt.plain_text();
    assert!(
        text.contains("Quarterly results") && text.contains("Outlook"),
        "fallback path surfaces the downrev text: {text:?}"
    );
    assert!(
        !text.contains("Master title style"),
        "master downrev text must not leak: {text:?}"
    );
    assert_eq!(
        text.matches("Steady growth expected").count(),
        1,
        "duplicate downrev packages collapse to one run: {text:?}"
    );
}

/// A downrev package whose shape XML carries non-ASCII text round-trips
/// byte-for-byte (the package is UTF-8 XML, not a legacy code page).
#[test]
fn test_wps_downrev_shape_text_preserves_non_ascii() {
    let xml = downrev_shape_xml(&["国家电建案例", "投研与风控智能体"]);
    let stream = ppt_stream_with_persist(
        &[vec![downrev_shape_group(&downrev_zip(&xml))]],
        &[master_group()],
    );
    let doc = open_ppt(stream);
    let text = doc.as_ppt().expect("parsed as ppt").plain_text();
    assert!(
        text.contains("国家电建案例") && text.contains("投研与风控智能体"),
        "UTF-8 shape text survives extraction: {text:?}"
    );
}

/// A carrier whose payload holds no ZIP at all is simply skipped.
#[test]
fn test_wps_downrev_carrier_without_zip_is_skipped() {
    let xml = downrev_shape_xml(&["Only real text"]);
    let stream = ppt_stream_with_persist(
        &[vec![
            downrev_shape_group(&Vec::new()),
            downrev_shape_group(&downrev_zip(&xml)),
        ]],
        &[master_group()],
    );
    let doc = open_ppt(stream);
    let text = doc.as_ppt().expect("parsed as ppt").plain_text();
    assert!(
        text.contains("Only real text"),
        "empty carrier is skipped, good text kept: {text:?}"
    );
}

/// A shape with legacy text beside a shape that has only a package: both
/// reach the reader, once each.
#[test]
fn test_shape_with_legacy_text_and_shape_with_package_both_reach_the_reader() {
    let xml = downrev_shape_xml(&["Downrev body"]);
    let slide = vec![
        normal_legacy_text_shape("Legacy title"),
        downrev_shape_group(&downrev_zip(&xml)),
    ];
    let stream = ppt_stream_with_persist(&[slide], &[master_group()]);
    let doc = open_ppt(stream);
    let text = doc.as_ppt().expect("parsed as ppt").plain_text();
    assert_eq!(text.matches("Legacy title").count(), 1, "{text:?}");
    assert_eq!(text.matches("Downrev body").count(), 1, "{text:?}");
}

/// What PowerPoint writes: full legacy text with a package carrying the same
/// shape again, in a form that differs in spacing and line breaks. The deck
/// must read exactly as it does with no packages at all.
#[test]
fn test_packages_beside_full_legacy_text_change_nothing() {
    let legacy = "Second\u{b}line\rHello";
    let package_xml = "<p:sp><a:p><a:r><a:t>Second</a:t></a:r><a:br/><a:r><a:t>line</a:t></a:r>\
                       </a:p><a:p><a:r><a:t>Hello </a:t></a:r></a:p></p:sp>";
    let shape = |carrier: Option<Vec<u8>>| {
        let mut children = vec![ppt_container(
            RT_CLIENT_TEXTBOX,
            0,
            &[
                ppt_record(RT_TEXT_HEADER, 0, 0, &1u32.to_le_bytes()),
                ppt_record(RT_TEXT_CHARS, 0, 0, &text_chars_payload(legacy)),
            ],
        )];
        children.extend(carrier);
        ppt_container(common::RT_SP_CONTAINER, 0, &children)
    };
    let read = |carrier| {
        let doc = open_ppt(ppt_stream_with_persist(&[vec![shape(carrier)]], &[]));
        (doc.plain_text(), doc.to_markdown(), doc.to_html())
    };
    let with = read(Some(downrev_carrier(&downrev_zip(package_xml))));
    let without = read(None);
    assert_eq!(with, without);
}

/// Real tables put a simple property after `metroBlob`: its data follows
/// every entry, not its own.
#[test]
fn test_package_with_a_property_after_it_reaches_the_reader() {
    let zip = downrev_zip(&downrev_shape_xml(&["Past the next entry"]));
    let mut payload = Vec::new();
    payload.extend_from_slice(&0xC3A9u16.to_le_bytes());
    payload.extend_from_slice(&(zip.len() as u32).to_le_bytes());
    payload.extend_from_slice(&0x03BFu16.to_le_bytes());
    payload.extend_from_slice(&0u32.to_le_bytes());
    payload.extend_from_slice(&zip);
    let carrier = ppt_record(common::RT_DOWNREV_CARRIER, 2, 3, &payload);
    let doc = open_ppt(ppt_stream_with_persist(&[vec![downrev_group_from_carrier(carrier)]], &[]));
    let ppt = doc.as_ppt().expect("parsed as ppt");
    assert!(ppt.plain_text().contains("Past the next entry"));
    assert!(ppt.warnings().is_empty(), "{:?}", ppt.warnings());
}

/// Past a limit, the deck says it isn't complete. Within one, it doesn't.
#[test]
fn test_text_past_a_package_limit_is_reported() {
    let slide_of = |s: usize, shapes: usize| -> Vec<Vec<u8>> {
        (0..shapes)
            .map(|i| {
                let text = format!("S{s}T{i}");
                downrev_shape_group(&downrev_zip(&downrev_shape_xml(&[text.as_str()])))
            })
            .collect()
    };
    let doc = open_ppt(ppt_stream_with_persist(&[slide_of(0, 256)], &[]));
    let ppt = doc.as_ppt().expect("parsed as ppt");
    assert!(ppt.text_complete(), "{:?}", ppt.warnings());
    assert!(ppt.plain_text().contains("S0T255"));

    let doc = open_ppt(ppt_stream_with_persist(&[slide_of(0, 257)], &[]));
    let ppt = doc.as_ppt().expect("parsed as ppt");
    assert!(!ppt.text_complete());
    assert!(
        ppt.warnings().iter().any(|w| w.contains("package limit")),
        "{:?}",
        ppt.warnings()
    );
    assert!(!ppt.plain_text().contains("S0T256"));
}

/// A damaged package is reported, and the rest of the deck still reads.
#[test]
fn test_damaged_package_is_reported() {
    let bad = zip_with_invalid_utf8();
    let good = downrev_zip(&downrev_shape_xml(&["Still here"]));
    let slide = vec![downrev_shape_group(&bad), downrev_shape_group(&good)];
    let doc = open_ppt(ppt_stream_with_persist(&[slide], &[]));
    let ppt = doc.as_ppt().expect("parsed as ppt");
    assert!(ppt.plain_text().contains("Still here"));
    assert!(!ppt.text_complete());
    assert!(ppt.warnings().iter().any(|w| w.contains("damaged")), "{:?}", ppt.warnings());
}

fn zip_with_invalid_utf8() -> Vec<u8> {
    use std::io::Write;
    let mut buf = Vec::new();
    {
        let mut w = zip::ZipWriter::new(Cursor::new(&mut buf));
        let opts = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated);
        w.start_file("drs/shapexml.xml", opts).expect("zip entry");
        w.write_all(&[0xFF; 64]).expect("zip payload");
        w.finish().expect("zip finish");
    }
    buf
}

/// With no usable persist directory, a slide list holding only a stub must not
/// hide the words the packages hold.
#[test]
fn test_package_text_beats_a_stub_in_the_slide_list_cache() {
    let xml = downrev_shape_xml(&["The real words of the slide"]);
    let slide_list = ppt_container(
        common::RT_SLIDE_LIST_WITH_TEXT,
        0,
        &[
            ppt_record(common::RT_SLIDE_PERSIST_ATOM, 0, 0, &[0u8; 20]),
            ppt_record(RT_TEXT_HEADER, 0, 0, &0u32.to_le_bytes()),
            ppt_record(RT_TEXT_CHARS, 0, 0, &text_chars_payload("x")),
        ],
    );
    let mut stream = ppt_container(common::RT_DOCUMENT, 0, &[slide_list]);
    stream.extend(common::ppt_slide(&[downrev_shape_group(&downrev_zip(&xml))]));
    let text = open_ppt(stream)
        .as_ppt()
        .expect("parsed as ppt")
        .plain_text();
    assert!(text.contains("The real words of the slide"), "{text:?}");
}

/// The cache's fuller text stays when the packages hold less.
#[test]
fn test_slide_list_cache_with_more_text_than_the_packages_is_kept() {
    let xml = downrev_shape_xml(&["Short"]);
    let slide_list = ppt_container(
        common::RT_SLIDE_LIST_WITH_TEXT,
        0,
        &[
            ppt_record(common::RT_SLIDE_PERSIST_ATOM, 0, 0, &[0u8; 20]),
            ppt_record(RT_TEXT_HEADER, 0, 0, &0u32.to_le_bytes()),
            ppt_record(RT_TEXT_CHARS, 0, 0, &text_chars_payload("The fuller cached words")),
        ],
    );
    let mut stream = ppt_container(common::RT_DOCUMENT, 0, &[slide_list]);
    stream.extend(common::ppt_slide(&[downrev_shape_group(&downrev_zip(&xml))]));
    let text = open_ppt(stream)
        .as_ppt()
        .expect("parsed as ppt")
        .plain_text();
    assert!(text.contains("The fuller cached words"), "{text:?}");
}

/// A title that exists only in a package renders as the slide's title: first,
/// not after the body.
#[test]
fn test_title_in_a_package_renders_as_the_title() {
    let xml = r#"<p:sp><p:nvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><a:p><a:r><a:t>The Title</a:t></a:r></a:p></p:sp>"#;
    let slide = vec![
        normal_legacy_text_body("Body text"),
        downrev_shape_group(&downrev_zip(xml)),
    ];
    let doc = open_ppt(ppt_stream_with_persist(&[slide], &[]));
    let ppt = doc.as_ppt().expect("parsed as ppt");
    assert_eq!(ppt.slides[0].text_runs[0].text, "The Title");
    assert_eq!(ppt.slides[0].text_runs[0].text_type, TextType::Title);
    let markdown = doc.to_markdown();
    assert!(
        markdown.find("The Title").expect("title") < markdown.find("Body text").expect("body"),
        "{markdown}"
    );
}

/// A decoy package (a complete, valid ZIP with the wrong text) in a
/// non-metroBlob complex property must not swallow the real package: the
/// property-table walk skips the decoy and the real text reaches the reader.
#[test]
fn test_decoy_package_before_the_real_one_does_not_lose_the_real_text() {
    let xml = downrev_shape_xml(&["Real slide text"]);
    let carrier = downrev_carrier_with_decoy(&downrev_zip(&xml));
    let stream =
        ppt_stream_with_persist(&[vec![downrev_group_from_carrier(carrier)]], &[master_group()]);
    let doc = open_ppt(stream);
    let text = doc.as_ppt().expect("parsed as ppt").plain_text();
    assert!(
        text.contains("Real slide text"),
        "the real package must be read past the decoy: {text:?}"
    );
    assert!(
        !text.contains("Decoy package text"),
        "the decoy package must not be read: {text:?}"
    );
}

/// A metroBlob whose declared length is shorter than the package holds must
/// bound the read: the declared bytes are all that may be consumed, and the
/// package is not carved around by scanning past the declaration.
#[test]
fn test_corrupt_declared_length_does_not_carve_past_the_declared_bytes() {
    let xml = downrev_shape_xml(&["Corruptly declared text"]);
    let carrier = downrev_carrier_short_length(&downrev_zip(&xml), 8);
    let stream =
        ppt_stream_with_persist(&[vec![downrev_group_from_carrier(carrier)]], &[master_group()]);
    let doc = open_ppt(stream);
    let text = doc.as_ppt().expect("parsed as ppt").plain_text();
    assert!(
        !text.contains("Corruptly declared text"),
        "a corrupt declared length must not be carved around: {text:?}"
    );
}

/// The fuzz target's seed corpus must carry a deck that parses and reaches
/// the downrev path, so mutations start from a valid structure.
#[test]
fn test_fuzz_seed_corpus_carries_a_valid_downrev_deck() {
    let path = concat!(env!("CARGO_MANIFEST_DIR"), "/fuzz/corpus/fuzz_parse/ppt_wps_downrev");
    let seed = std::fs::read(path).expect("fuzz seed corpus file must exist");
    let doc = Document::from_reader(Cursor::new(seed), DocumentFormat::Ppt)
        .expect("the seed deck must open");
    let ppt = doc.as_ppt().expect("parsed as ppt");
    assert_eq!(ppt.slides.len(), 14, "the seed deck has 14 slides");
    let text = ppt.plain_text();
    assert!(
        text.contains("Slide 1 heading"),
        "the seed deck must reach the downrev path: {text:?}"
    );
}

/// The malformed seed must stay parseable: its robustness carriers (a decoy
/// package, a corrupt declared length, an empty package) are skipped and the
/// one good package on the slide still reaches the reader.
#[test]
fn test_fuzz_seed_corpus_malformed_deck_stays_parseable() {
    let path =
        concat!(env!("CARGO_MANIFEST_DIR"), "/fuzz/corpus/fuzz_parse/ppt_wps_downrev_malformed");
    let seed = std::fs::read(path).expect("malformed fuzz seed corpus file must exist");
    let doc = Document::from_reader(Cursor::new(seed), DocumentFormat::Ppt)
        .expect("the malformed seed deck must open");
    let text = doc.as_ppt().expect("parsed as ppt").plain_text();
    assert!(
        text.contains("Slide 3 heading"),
        "the good package on the mixed slide must survive: {text:?}"
    );
    assert!(
        !text.contains("Decoy package text"),
        "the decoy package must not be read: {text:?}"
    );
}

/// A carrier nested as deep as real files put it — a slide's drawing
/// container, two shape-group levels and the shape tree — is still
/// reached: the walk recurses through containers instead of probing their
/// data. Real WPS files sit the carrier four to seven container levels
/// deep; every earlier test in this file nests it exactly one.
#[test]
fn test_deeply_nested_carrier_is_reached_through_the_container_chain() {
    let xml = downrev_shape_xml(&["Deeply nested text"]);
    let shape_group = downrev_group_from_carrier(downrev_carrier(&downrev_zip(&xml)));
    // slide → DgContainer → SpgrContainer → SpgrContainer → SpContainer → carrier
    let nested = ppt_container(
        RT_DG_CONTAINER,
        0,
        &[ppt_container(
            RT_SPGR_CONTAINER,
            0,
            &[ppt_container(RT_SPGR_CONTAINER, 0, &[shape_group])],
        )],
    );
    let stream = ppt_stream_with_persist(&[vec![nested]], &[master_group()]);
    let doc = open_ppt(stream);
    let text = doc.as_ppt().expect("parsed as ppt").plain_text();
    assert!(
        text.contains("Deeply nested text"),
        "a carrier four container levels deep must still be reached: {text:?}"
    );
}

/// One budget covers both passes over a deck: when the persist path finds no
/// text and the fallback walks the same packages again, what the first spent
/// isn't available to the second. 20 slides of 256 masks is 5,120 packages:
/// within the file's 8,192 once, past it twice.
#[test]
fn test_package_budget_is_shared_between_persist_and_fallback() {
    let groups: Vec<Vec<u8>> = (0..256)
        .map(|i| {
            let mask = "_".repeat(i + 1);
            downrev_shape_group(&downrev_zip(&downrev_shape_xml(&[mask.as_str()])))
        })
        .collect();
    let slides = vec![groups; 20];
    let doc = open_ppt(ppt_stream_with_persist(&slides, &[]));
    let ppt = doc.as_ppt().expect("parsed as ppt");
    assert!(ppt.plain_text().trim().is_empty());
    assert!(
        ppt.warnings().iter().any(|w| w.contains("package limit")),
        "{:?}",
        ppt.warnings()
    );
}

/// Each slide's packages reach that slide, through the public API.
#[test]
fn test_downrev_text_per_slide_through_public_api() {
    let xml3 = downrev_shape_xml(&["Appendix", "See attached notes"]);
    let slide_three = vec![downrev_shape_group(&downrev_zip(&xml3))];
    let stream =
        ppt_stream_with_persist(&[slide_one(), slide_two(), slide_three], &[master_group()]);
    let doc = open_ppt(stream);
    let ppt = doc.as_ppt().expect("parsed as ppt");
    assert_eq!(ppt.slides.len(), 3, "three slides");
    assert!(
        ppt.slides[0]
            .text_runs
            .iter()
            .any(|r| r.text.contains("Quarterly results")),
        "slide 1 keeps its downrev text"
    );
    assert!(
        ppt.slides[1]
            .text_runs
            .iter()
            .any(|r| r.text.contains("Outlook")),
        "slide 2 keeps its downrev text"
    );
    assert!(
        ppt.slides[2]
            .text_runs
            .iter()
            .any(|r| r.text.contains("Appendix")),
        "slide 3 keeps its downrev text"
    );
}

/// The same package on two slides contributes its text to both: each slide
/// keeps its own copy.
#[test]
fn test_downrev_identical_package_kept_on_each_slide() {
    let xml = downrev_shape_xml(&["Shared heading"]);
    let grp = downrev_shape_group(&downrev_zip(&xml));
    let stream = ppt_stream_with_persist(&[vec![grp.clone()], vec![grp]], &[]);
    let doc = open_ppt(stream);
    let ppt = doc.as_ppt().expect("parsed as ppt");
    assert!(
        ppt.slides[0]
            .text_runs
            .iter()
            .any(|r| r.text.contains("Shared heading")),
        "slide 1 extracts the shared package"
    );
    assert!(
        ppt.slides[1]
            .text_runs
            .iter()
            .any(|r| r.text.contains("Shared heading")),
        "slide 2 must also extract the shared package"
    );
}

/// A legacy `.ppt` carrying packages on two slides, synthesised by the same
/// builders as the other tests and committed as
/// `tests/fixtures/wps_downrev.ppt`. It isn't a capture from WPS.
fn build_downrev_ppt_bytes() -> Vec<u8> {
    let stream = ppt_stream_with_persist(&[slide_one(), slide_two()], &[master_group()]);
    cfb_with_stream("PowerPoint Document", &stream)
}

/// The committed fixture reads from disk, per slide, with no master text. It
/// must read the same as its builder's deck, so the two can't drift apart
/// unnoticed. The bytes can differ: a ZIP records the system that wrote it.
#[test]
fn test_committed_fixture_reads_from_a_file() {
    let path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/wps_downrev.ppt");
    let file = std::fs::File::open(path).expect("wps_downrev.ppt fixture must open");
    let doc = Document::from_reader(file, DocumentFormat::Ppt).expect("fixture opens");
    let built = Document::from_reader(Cursor::new(build_downrev_ppt_bytes()), DocumentFormat::Ppt)
        .expect("built deck opens");
    assert_eq!(
        doc.to_markdown(),
        built.to_markdown(),
        "the fixture no longer matches its builder"
    );
    let ppt = doc.as_ppt().expect("parsed as ppt");
    assert_eq!(ppt.slides.len(), 2, "one slide per SlideContainer");
    let text = ppt.plain_text();
    assert!(text.contains("Quarterly results"), "slide 1 downrev text");
    assert!(text.contains("Outlook"), "slide 2 downrev text");
    assert!(
        !text.contains("Master title style"),
        "master downrev text must not leak into slides"
    );
    assert_eq!(
        text.matches("Steady growth expected").count(),
        1,
        "within-slide duplicate packages collapse to one run"
    );
}

/// The fuzz target's seeds must reach the package reader, by the target's own
/// reading of them: byte 0 packs the carrier count and the property count.
#[test]
fn test_fuzz_seeds_reach_the_package_reader() {
    for (seed, expected) in [
        ("ppt_downrev_valid", "Seed text"),
        ("ppt_downrev_decoy", "Seed text"),
        ("ppt_downrev_degenerate", "Seed text"),
        ("ppt_downrev_short_length", ""),
        ("ppt_downrev_trailing_property", "Seed text"),
        ("ppt_downrev_two_carriers", "Second seed text"),
    ] {
        let path = format!("{}/fuzz/corpus/fuzz_ppt_downrev/{seed}", env!("CARGO_MANIFEST_DIR"));
        let data = std::fs::read(path).expect("seed must exist");
        let (&control, rest) = data.split_first().expect("seed must not be empty");
        let carriers = usize::from(control & 3) + 1;
        let properties = u16::from(control >> 2);
        let chunk = rest.len() / carriers;
        let groups: Vec<Vec<u8>> = (0..carriers)
            .map(|i| {
                let payload = &rest[i * chunk..((i + 1) * chunk).min(rest.len())];
                let carrier = ppt_record(common::RT_DOWNREV_CARRIER, properties, 3, payload);
                downrev_group_from_carrier(carrier)
            })
            .collect();
        let text = open_ppt(ppt_stream_with_persist(&[groups], &[]))
            .as_ppt()
            .expect("parsed as ppt")
            .plain_text();
        assert!(text.contains(expected), "{seed}: {text:?}");
        assert!(!text.contains("Decoy"), "{seed}: {text:?}");
        if expected.is_empty() {
            assert!(text.trim().is_empty(), "{seed}: {text:?}");
        }
    }
}

/// `RT_CLIENT_TEXTBOX` ([MS-PPT] 2.5.28) and `RT_TEXT_HEADER` ([MS-PPT]
/// 2.9.44) record types, used to build a normal legacy-text shape.
const RT_CLIENT_TEXTBOX: u16 = 0xF00D;
const RT_TEXT_HEADER: u16 = 0x0F9F;

/// A normal legacy-text slide: one shape whose text lives in a
/// `TextCharsAtom`, with no WPS downrev carrier anywhere in the stream.
fn normal_legacy_text_shape(text: &str) -> Vec<u8> {
    ppt_container(
        common::RT_SP_CONTAINER,
        0,
        &[ppt_container(
            RT_CLIENT_TEXTBOX,
            0,
            &[
                ppt_record(RT_TEXT_HEADER, 0, 0, &0u32.to_le_bytes()), // Title
                ppt_record(RT_TEXT_CHARS, 0, 0, &text_chars_payload(text)),
            ],
        )],
    )
}

/// As [`normal_legacy_text_shape`], typed as body text.
fn normal_legacy_text_body(text: &str) -> Vec<u8> {
    ppt_container(
        common::RT_SP_CONTAINER,
        0,
        &[ppt_container(
            RT_CLIENT_TEXTBOX,
            0,
            &[
                ppt_record(RT_TEXT_HEADER, 0, 0, &1u32.to_le_bytes()),
                ppt_record(RT_TEXT_CHARS, 0, 0, &text_chars_payload(text)),
            ],
        )],
    )
}

/// Regression: enabling the downrev scan must not change a normal legacy
/// `.ppt` deck. A deck made only of `TextCharsAtom` text (no `0xF122`
/// carriers) must extract exactly that text and add no `TextType::Other`
/// runs — `Other` is the type downrev packages surface as, so its absence
/// proves the downrev walk contributed nothing.
#[test]
fn test_downrev_scan_leaves_normal_deck_unchanged() {
    let stream = ppt_stream_with_persist(
        &[
            vec![normal_legacy_text_shape("Quarterly results")],
            vec![normal_legacy_text_shape("Outlook")],
        ],
        &[],
    );
    let doc = open_ppt(stream);
    let ppt = doc.as_ppt().expect("parsed as ppt");
    assert_eq!(ppt.slides.len(), 2, "two slides");
    assert!(
        ppt.slides[0]
            .text_runs
            .iter()
            .any(|r| r.text.contains("Quarterly results")),
        "slide 1 keeps its legacy text"
    );
    assert!(
        ppt.slides[1]
            .text_runs
            .iter()
            .any(|r| r.text.contains("Outlook")),
        "slide 2 keeps its legacy text"
    );
    // No carrier means the downrev walk must have added nothing: there must
    // be zero `TextType::Other` runs anywhere in the deck.
    let other_runs = ppt
        .slides
        .iter()
        .flat_map(|s| &s.text_runs)
        .filter(|r| matches!(r.text_type, TextType::Other))
        .count();
    assert_eq!(other_runs, 0, "no downrev (TextType::Other) runs in a normal deck");
}
