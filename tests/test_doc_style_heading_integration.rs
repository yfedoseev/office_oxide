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
fn heading_style_from_the_style_sheet_sets_the_real_level() {
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
fn without_a_style_sheet_the_same_paragraph_is_not_a_heading() {
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
