//! Style-sheet heading levels, reached through the public API.
//!
//! These live in `tests/` rather than in a `#[cfg(test)]` module under `src/`
//! on purpose. `scripts/revert-check.sh` reverts every changed file under
//! `src/` before running the suite, so a unit test that sits in the same file
//! as the code it covers is reverted along with it — the suite then passes with
//! the production change removed, and the gate correctly reports that nothing
//! was proven. A test out here survives the revert, so reverting the
//! style-sheet parser must make *this* fail.
//!
//! The document is built in code (AGENTS.md rule #4 — no committed third-party
//! fixture).

mod common;

use common::{Para, build_doc_with_styles, heading_style_sheet, open_doc, prose_grpprl};
use office_oxide::ir::Element;

/// Pull the `(level, text)` of every `Heading` out of a parsed document.
fn headings(doc: &office_oxide::Document) -> Vec<(u8, String)> {
    let ir = doc.to_ir();
    let mut out = Vec::new();
    for section in &ir.sections {
        for element in &section.elements {
            if let Element::Heading(h) = element {
                let text: String = h
                    .content
                    .iter()
                    .filter_map(|c| match c {
                        office_oxide::ir::InlineContent::Text(t) => Some(t.text.as_str()),
                        _ => None,
                    })
                    .collect();
                out.push((h.level, text));
            }
        }
    }
    out
}

/// A paragraph whose PAPX `istd` points at the built-in `Heading 3` style must
/// come out as a level-3 heading.
///
/// Level 3 cannot come from the line-shape guess, which caps at 2, and
/// "Subsection Three" is neither ALL-CAPS nor a short opening line, so under a
/// build with no style-sheet parsing this paragraph is ordinary prose. That
/// asymmetry is what makes this test evidence rather than a tautology.
#[test]
fn test_heading_style_from_the_style_sheet_sets_the_real_level() {
    let paras = [
        Para {
            text: "Introduction.",
            terminator: '\r',
            grpprl: prose_grpprl(),
        },
        Para {
            text: "Subsection Three",
            terminator: '\r',
            grpprl: prose_grpprl(),
        },
    ];
    // istd 0 = Normal, istd 3 = Heading 3 per the style sheet below.
    let bytes = build_doc_with_styles(&paras, &[0, 3], &heading_style_sheet());
    let doc = open_doc(&bytes);

    let found = headings(&doc);
    assert_eq!(found.len(), 1, "only the styled paragraph is a heading; got {found:?}");
    assert_eq!(found[0].1, "Subsection Three");
    assert_eq!(
        found[0].0, 3,
        "a `Heading 3` style must produce level 3, not the heuristic's 1 or 2"
    );
}

/// The same document without a style sheet must not invent the level: nothing
/// marks the paragraph as a heading, so it stays prose. This is the contrast
/// that keeps the test above honest.
#[test]
fn test_without_a_style_sheet_the_same_paragraph_is_not_a_heading() {
    let paras = [
        Para {
            text: "Introduction.",
            terminator: '\r',
            grpprl: prose_grpprl(),
        },
        Para {
            text: "Subsection Three",
            terminator: '\r',
            grpprl: prose_grpprl(),
        },
    ];
    let bytes = build_doc_with_styles(&paras, &[0, 3], &[]);
    let doc = open_doc(&bytes);

    assert!(
        headings(&doc).is_empty(),
        "no style sheet means no style-derived level; got {:?}",
        headings(&doc)
    );
}

/// A plain `Heading 3` paragraph carries no grpprl: Word writes it as an
/// *istd-only* PAPX (`cw = 0`, then a Word8 re-read byte `cb' = 1` meaning a
/// 2-byte `GrpPrlAndIstd` holding only the `istd`). The parser must recover the
/// style index from that form.
///
/// This is the regression for the bug where `extract_grpprl` returned `istd: 0`
/// from its `cb < 3` early return *before* reading `istd` — so every
/// style-only paragraph (≈33% of paragraphs in real corpora) silently lost its
/// level and fell back to prose. Reverting the production fix makes this fail.
#[test]
fn test_istd_only_papx_resolves_its_heading_style() {
    // `grpprl` empty → `build_fkp_page` emits the istd-only form.
    let paras = [Para {
        text: "Subsection Three",
        terminator: '\r',
        grpprl: Vec::new(),
    }];
    let bytes = build_doc_with_styles(&paras, &[3], &heading_style_sheet());
    let doc = open_doc(&bytes);

    let found = headings(&doc);
    assert_eq!(
        found.len(),
        1,
        "the istd-only paragraph must resolve to a heading; got {found:?}"
    );
    assert_eq!(found[0].1, "Subsection Three");
    assert_eq!(
        found[0].0, 3,
        "a `Heading 3` style must produce level 3 even from the istd-only form"
    );
}

/// Word also writes style-bearing paragraphs with a non-empty grpprl (the
/// `cb != 0` form). This keeps one test on that form so both PAPX shapes the
/// parser sees in real files are covered. `sprmPJc` (0x2461) only sets
/// justification and must not affect the style-derived level.
#[test]
fn test_papx_cb_nonzero_form_resolves_heading() {
    let jc_grpprl = vec![0x61, 0x24, 0x00, 0x00]; // sprmPJc, 2-byte operand = 0
    let paras = [Para {
        text: "Subsection Three",
        terminator: '\r',
        grpprl: jc_grpprl,
    }];
    let bytes = build_doc_with_styles(&paras, &[3], &heading_style_sheet());
    let doc = open_doc(&bytes);

    let found = headings(&doc);
    assert_eq!(
        found.len(),
        1,
        "the cb!=0 styled paragraph must resolve to a heading; got {found:?}"
    );
    assert_eq!(found[0].0, 3);
}

/// A level resolved from a *style* does **not** switch the line-shape guess
/// off: `has_structured_headings` is keyed on `sprmPOutLvl` only.
///
/// The body line here is short and ALL-CAPS with no trailing '.', i.e. exactly
/// the shape the guess turns into a heading — so it still being recognised is
/// what shows the gate did not fire for its neighbour. This is the integration
/// twin of the unit test that used to live under `src/doc/document.rs`; it
/// survives `scripts/revert-check.sh` (which reverts `src/`), so reverting the
/// style-sheet parser must make *this* fail.
#[test]
fn test_style_derived_heading_does_not_switch_off_line_shape_guess() {
    let paras = [
        Para {
            text: "SHORT ALL CAPS BODY LINE",
            terminator: '\r',
            grpprl: prose_grpprl(),
        },
        Para {
            text: "Subsection Three",
            terminator: '\r',
            grpprl: prose_grpprl(),
        },
    ];
    // istd 0 = Normal (body), istd 3 = Heading 3 per the style sheet below.
    let bytes = build_doc_with_styles(&paras, &[0, 3], &heading_style_sheet());
    let doc = open_doc(&bytes);

    let found = headings(&doc);
    assert_eq!(
        found.len(),
        2,
        "the styled paragraph gets its real level and the ALL-CAPS line is \
         still guessed; got {found:?}"
    );
    assert!(
        found.iter().any(|(lvl, _)| *lvl == 3),
        "the style-derived level must still be used; got {found:?}"
    );
}

/// The other half of that distinction: a level the paragraph states itself
/// (`sprmPOutLvl`, `.source == Sprm`) **does** switch the line-shape guess
/// off. `sprmPOutLvl` (0x2640) carries a 1-byte operand; `0x04` is Heading 5.
///
/// Both paragraphs are Normal (istd 0), so the Heading level comes solely from
/// the SPRM, not the style sheet — and the line heuristic (which tops out at
/// level 2) would never produce level 5. This is the integration twin of the
/// unit test that used to live under `src/doc/document.rs`; it survives
/// `scripts/revert-check.sh`, so reverting the outline-SPRM parser must make
/// *this* fail.
#[test]
fn test_sprm_outline_papx_sets_the_real_level() {
    let outline_grpprl = vec![0x40, 0x26, 0x04]; // sprmPOutLvl, operand 0x04 -> Heading 5
    let paras = [
        Para {
            text: "Introduction.",
            terminator: '\r',
            grpprl: prose_grpprl(),
        },
        Para {
            text: "Subsection Three",
            terminator: '\r',
            grpprl: outline_grpprl,
        },
    ];
    // Both paragraphs Normal (istd 0); the Heading level comes solely from the
    // SPRM, not the style sheet.
    let bytes = build_doc_with_styles(&paras, &[0, 0], &[]);
    let doc = open_doc(&bytes);

    let found = headings(&doc);
    assert_eq!(found.len(), 1, "only the SPRM-marked paragraph is a heading; got {found:?}");
    assert_eq!(found[0].1, "Subsection Three");
    assert_eq!(
        found[0].0, 5,
        "the outline SPRM must set the real level 5, not the heuristic's <=2"
    );
}

/// A level stated by the paragraph itself (`sprmPOutLvl`, `.source == Sprm`)
/// **does** switch the line-shape guess off. `sprmPOutLvl` (0x2640) with a
/// 1-byte operand `0x01` is Heading level 2 (the operand is zero-based).
///
/// The first paragraph is short and ALL-CAPS — exactly the shape the guess
/// would otherwise turn into a heading. With a stated level present it must
/// stay prose, so exactly one heading (the stated one) reaches the IR. This is
/// the integration twin of the unit test that used to live under
/// `src/doc/document.rs`; it survives `scripts/revert-check.sh`, so reverting
/// the outline-SPRM parse must make *this* fail.
#[test]
fn test_sprm_derived_heading_switches_off_line_shape_guess() {
    let outline_grpprl = vec![0x40, 0x26, 0x01]; // sprmPOutLvl, operand 0x01 -> Heading 2
    let paras = [
        Para {
            text: "ALL CAPS PROSE",
            terminator: '\r',
            grpprl: prose_grpprl(),
        },
        Para {
            text: "Real Heading",
            terminator: '\r',
            grpprl: outline_grpprl,
        },
    ];
    // Both paragraphs Normal (istd 0); the Heading level comes solely from the
    // SPRM, which also switches the line-shape guess off.
    let bytes = build_doc_with_styles(&paras, &[0, 0], &[]);
    let doc = open_doc(&bytes);

    let found = headings(&doc);
    assert_eq!(
        found.len(),
        1,
        "the ALL-CAPS line must not be guessed once a paragraph states its own \
         level; got {found:?}"
    );
    assert_eq!(found[0].1, "Real Heading");
    assert_eq!(found[0].0, 2, "sprmPOutLvl operand 0x01 must become Heading level 2");
}

/// A direct `sprmPOutLvl` of 9 is [MS-DOC]'s explicit "body text" level, and
/// direct formatting beats the style: a paragraph styled `Heading 3` that
/// carries it must come out as prose, not fall through to its style.
///
/// The first paragraph is the contrast. It has the same style and no override,
/// so it must be a level-3 heading; that is only true when the style sheet is
/// parsed, which keeps this test from passing on a build without the feature.
#[test]
fn test_direct_body_text_level_suppresses_a_heading_style() {
    let body_text_grpprl = vec![0x40, 0x26, 0x09]; // sprmPOutLvl, operand 9 = body text
    let paras = [
        Para {
            text: "Subsection Three",
            terminator: '\r',
            grpprl: prose_grpprl(),
        },
        Para {
            text: "Styled but body text",
            terminator: '\r',
            grpprl: body_text_grpprl,
        },
    ];
    // Both paragraphs carry istd 3 = Heading 3.
    let bytes = build_doc_with_styles(&paras, &[3, 3], &heading_style_sheet());
    let doc = open_doc(&bytes);

    let found = headings(&doc);
    assert_eq!(
        found,
        vec![(3, "Subsection Three".to_string())],
        "only the paragraph without the body-text override is a heading"
    );
}

/// `sprmPIstd` (0x4600, 2-byte operand) overrides the `istd` in the PAPX
/// header ([MS-DOC] §2.6.2). Checked in both directions: a Normal header
/// pointed at `Heading 3` becomes a heading, and a `Heading 3` header pointed
/// at Normal stops being one.
#[test]
fn test_sprm_p_istd_overrides_the_papx_header_istd() {
    let to_heading_3 = vec![0x00, 0x46, 0x03, 0x00]; // sprmPIstd = 3
    let to_normal = vec![0x00, 0x46, 0x00, 0x00]; // sprmPIstd = 0
    let paras = [
        Para {
            text: "Promoted By Override",
            terminator: '\r',
            grpprl: to_heading_3,
        },
        Para {
            text: "Demoted By Override",
            terminator: '\r',
            grpprl: to_normal,
        },
    ];
    // Header istds are the opposite of the overrides: Normal, then Heading 3.
    let bytes = build_doc_with_styles(&paras, &[0, 3], &heading_style_sheet());
    let doc = open_doc(&bytes);

    let found = headings(&doc);
    assert_eq!(
        found,
        vec![(3, "Promoted By Override".to_string())],
        "the override, not the header istd, decides which style applies"
    );
}

/// The warnings a document reports through its IR metadata.
fn warnings(doc: &office_oxide::Document) -> Vec<String> {
    doc.to_ir().metadata.warnings
}

/// An opening prose line, then a paragraph whose style (istd 3) is `Heading 3`.
/// The opener keeps the old line-shape guess from claiming the second line.
fn prose_then_styled_paragraph() -> [Para; 2] {
    [
        Para {
            text: "Introduction.",
            terminator: '\r',
            grpprl: Vec::new(),
        },
        Para {
            text: "Subsection Three",
            terminator: '\r',
            grpprl: Vec::new(),
        },
    ]
}

/// A style sheet that is rejected must not cost the document its style-derived
/// headings in silence: the reason reaches `Metadata::warnings`. The
/// well-formed sheet is the contrast, so the assertion cannot be satisfied by
/// a build that warns about everything.
#[test]
fn test_a_rejected_style_sheet_is_reported_in_the_metadata_warnings() {
    let mut bad = heading_style_sheet();
    bad[4..6].copy_from_slice(&0x0020u16.to_le_bytes()); // cbSTDBaseInFile

    let good_doc = open_doc(&build_doc_with_styles(
        &prose_then_styled_paragraph(),
        &[0, 3],
        &heading_style_sheet(),
    ));
    assert!(
        warnings(&good_doc)
            .iter()
            .all(|w| !w.contains("style sheet")),
        "a well-formed sheet must not warn; got {:?}",
        warnings(&good_doc)
    );

    let bad_doc = open_doc(&build_doc_with_styles(&prose_then_styled_paragraph(), &[0, 3], &bad));
    assert!(headings(&bad_doc).is_empty(), "a rejected sheet yields no style heading");
    assert!(
        warnings(&bad_doc)
            .iter()
            .any(|w| w.contains("style sheet") && w.contains("cbSTDBaseInFile")),
        "the rejection must be reported; got {:?}",
        warnings(&bad_doc)
    );
}

/// A sheet cut off mid-entry keeps the headings it did read, and reports that
/// it is incomplete.
#[test]
fn test_a_truncated_style_sheet_is_reported_and_keeps_the_styles_it_read() {
    let mut cut = heading_style_sheet();
    cut.truncate(cut.len() - 20);

    let doc = open_doc(&build_doc_with_styles(&prose_then_styled_paragraph(), &[0, 3], &cut));
    assert_eq!(headings(&doc), vec![(3, "Subsection Three".to_string())]);
    assert!(
        warnings(&doc)
            .iter()
            .any(|w| w.contains("style sheet") && w.contains("declares")),
        "the truncation must be reported; got {:?}",
        warnings(&doc)
    );
}
