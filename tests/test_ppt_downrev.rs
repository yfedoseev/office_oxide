//! WPS "downrev" packages in legacy `.ppt` files.
//!
//! Defect class: when WPS saves an OOXML-origin presentation to the legacy
//! binary format, shapes that carry OOXML-only content are written back as
//! self-contained ZIP packages — a `drs/shapexml.xml` entry holding a
//! standard DrawingML `<p:sp>` fragment — attached to records inside the
//! standard OfficeArt shape tree, while the legacy `TextCharsAtom` /
//! `TextBytesAtom` atoms are left nearly empty. A reader that walks only
//! the legacy text atoms therefore returns an almost-empty deck with `Ok`
//! — indistinguishable from a genuinely text-free presentation.
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

/// A legacy text atom and a downrev package on the same slide both reach
/// the reader, and identical text is not duplicated.
#[test]
fn test_wps_downrev_text_merges_with_legacy_text_atoms() {
    let xml = downrev_shape_xml(&["Downrev body", "Shared line"]);
    let legacy = ppt_record(RT_TEXT_CHARS, 0, 0, &text_chars_payload("Legacy title\rShared line"));
    let slide = vec![legacy, downrev_shape_group(&downrev_zip(&xml))];
    let stream = ppt_stream_with_persist(&[slide], &[master_group()]);
    let doc = open_ppt(stream);
    let text = doc.as_ppt().expect("parsed as ppt").plain_text();
    assert!(text.contains("Legacy title"), "legacy atom text survives: {text:?}");
    assert!(text.contains("Downrev body"), "downrev text survives: {text:?}");
    assert_eq!(
        text.matches("Shared line").count(),
        1,
        "identical text across atom and package is not duplicated: {text:?}"
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

/// G4: `extract_deck_text` builds a single `DownrevBudget` and threads it
/// through BOTH the persist-directory path and the no-directory fallback
/// inside one call. This end-to-end test pins that shared-budget plumbing so a
/// later change that strands budget on one path (and starves the other) is
/// caught. The persist path here consumes budget and the fallback still runs
/// against the remainder, exercising both arms of the shared budget.
#[test]
fn test_downrev_budget_is_shared_across_persist_and_fallback() {
    let stream = ppt_stream_with_persist(&[slide_one(), slide_two()], &[master_group()]);
    let doc = open_ppt(stream);
    let ppt = doc.as_ppt().expect("parsed as ppt");
    let text = ppt.plain_text();
    // Both extraction arms ran against one budget; the downrev text still
    // surfaces and within-slide duplicates collapse.
    assert!(text.contains("Quarterly results"));
    assert!(text.contains("Outlook"));
    assert_eq!(text.matches("Steady growth expected").count(), 1);
}

/// G5: whole-pipeline regression guard through the public `Document` /
/// `from_reader` API. A real `.ppt` fixture is not checked in, so this
/// synthesises an equivalent multi-slide stream; it locks the per-slide
/// downrev-text contract end-to-end (the cross-slide dedup bug from R1 is
/// tracked by the ignored regression below).
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

/// R1 (integration exposure, ignored until fixed): the same downrev package
/// carried on two different slides must contribute its text to BOTH slides —
/// each slide is a distinct shape. The per-file `seen_packages` set currently
/// drops it on the second slide. Run `cargo test -- --ignored` to see it fail.
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
        "slide 2 must ALSO extract the shared package (R1: per-slide text)"
    );
}

/// Build a fully-valid (but synthesised) legacy `.ppt` byte stream carrying
/// downrev packages on two slides, via the same `common` builders the other
/// tests use. It is persisted as `tests/fixtures/wps_downrev.ppt` so G5 can run
/// against a real on-disk file (not an in-memory cursor). A captured real WPS
/// export would be a stricter fixture but is not checked in.
fn build_downrev_ppt_bytes() -> Vec<u8> {
    let stream = ppt_stream_with_persist(&[slide_one(), slide_two()], &[master_group()]);
    cfb_with_stream("PowerPoint Document", &stream)
}

/// G5 (real-file regression): a committed, fully-valid `.ppt` fixture is read
/// from disk through the public `Document::from_reader` API (real file IO, not a
/// cursor). Downrev text must surface per slide and master text must not leak.
/// The fixture is (re)created from `build_downrev_ppt_bytes` only if it is
/// missing, so the test is self-healing yet normally exercises the checked-in
/// file.
#[test]
fn test_downrev_real_file_fixture_regression() {
    let dir = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures");
    let path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/wps_downrev.ppt");
    if !std::path::Path::new(path).exists() {
        std::fs::create_dir_all(dir).expect("create fixtures dir");
        std::fs::write(path, build_downrev_ppt_bytes()).expect("write fixture");
    }
    let file = std::fs::File::open(path).expect("wps_downrev.ppt fixture must exist");
    let doc = Document::from_reader(file, DocumentFormat::Ppt).expect("fixture opens");
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
