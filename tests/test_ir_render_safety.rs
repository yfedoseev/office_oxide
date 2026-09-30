//! Renderer agreement and output-safety guards.
//!
//! `plain_text`, `to_markdown` and `to_html` are three views of the same
//! document; a consumer that switches between them must not see the text
//! change. These tests also cover the cases where document content — which
//! is untrusted input — could inject structure into the rendered output.

use office_oxide::ir::*;
use office_oxide::{DocumentFormat, DocumentIR};

fn span(text: &str) -> InlineContent {
    InlineContent::Text(TextSpan::plain(text))
}

fn linked(text: &str, url: &str) -> InlineContent {
    InlineContent::Text(TextSpan {
        text: text.to_string(),
        hyperlink: Some(url.to_string()),
        ..Default::default()
    })
}

fn ir_with(elements: Vec<Element>) -> DocumentIR {
    DocumentIR {
        metadata: Metadata {
            format: DocumentFormat::Docx,
            ..Default::default()
        },
        sections: vec![Section {
            elements,
            ..Default::default()
        }],
        defined_names: Vec::new(),
    }
}

// ---------------------------------------------------------------------------
// URL scheme filtering
// ---------------------------------------------------------------------------

#[test]
fn test_dangerous_url_schemes_are_dropped_from_links() {
    for url in [
        "javascript:alert(1)",
        "JavaScript:alert(1)",
        "vbscript:msgbox",
        "data:text/html;base64,PHNjcmlwdD4=",
    ] {
        let ir = ir_with(vec![Element::Paragraph(Paragraph {
            content: vec![linked("click me", url)],
            ..Default::default()
        })]);
        let md = ir.to_markdown();
        let html = ir.to_html();
        assert!(
            !md.contains("javascript") && !md.contains("JavaScript") && !md.contains("data:"),
            "markdown kept a dangerous scheme for {url:?}: {md}"
        );
        assert!(!html.contains("href"), "html emitted an href for {url:?}: {html}");
        // The link *text* must still be there — filtering the target is
        // not a licence to drop the words.
        assert!(md.contains("click me") && html.contains("click me"));
    }
}

#[test]
fn test_a_control_character_cannot_smuggle_a_scheme_past_the_filter() {
    let ir = ir_with(vec![Element::Paragraph(Paragraph {
        content: vec![linked("x", "java\u{0}script:alert(1)")],
        ..Default::default()
    })]);
    assert!(!ir.to_html().contains("href"));
}

#[test]
fn test_ordinary_urls_still_render_as_links() {
    for url in [
        "https://example.com/a?b=1",
        "mailto:someone@example.com",
        "#internal-anchor",
        "relative/path.html",
    ] {
        let ir = ir_with(vec![Element::Paragraph(Paragraph {
            content: vec![linked("go", url)],
            ..Default::default()
        })]);
        assert!(ir.to_html().contains("href="), "expected a link for {url:?}: {}", ir.to_html());
        assert!(ir.to_markdown().contains("[go]("));
    }
}

// ---------------------------------------------------------------------------
// Markdown metacharacters in document text
// ---------------------------------------------------------------------------

#[test]
fn test_document_text_cannot_inject_markdown_structure() {
    let ir = ir_with(vec![Element::Paragraph(Paragraph {
        content: vec![span("a | b [x](javascript:alert(1)) *not emphasis*")],
        ..Default::default()
    })]);
    let md = ir.to_markdown();
    // The brackets are escaped, so the `[x](…)` in the document text stays
    // literal text instead of becoming a javascript: link.
    assert!(md.contains(r"\[x\]"), "link brackets not escaped: {md}");
    assert!(md.contains(r"\|"), "table pipe not escaped: {md}");
    assert!(md.contains(r"\*not emphasis\*"), "emphasis not escaped: {md}");
}

// ---------------------------------------------------------------------------
// The three renderers must agree
// ---------------------------------------------------------------------------

fn hf(text: &str) -> HeaderFooter {
    HeaderFooter {
        content: vec![Element::Paragraph(Paragraph {
            content: vec![span(text)],
            ..Default::default()
        })],
    }
}

#[test]
fn test_all_three_renderers_include_headers_and_footers() {
    let ir = DocumentIR {
        metadata: Metadata {
            format: DocumentFormat::Docx,
            ..Default::default()
        },
        sections: vec![Section {
            elements: vec![Element::Paragraph(Paragraph {
                content: vec![span("BODY_TEXT")],
                ..Default::default()
            })],
            header: Some(hf("HEADER_TEXT")),
            footer: Some(hf("FOOTER_TEXT")),
            ..Default::default()
        }],
        defined_names: Vec::new(),
    };
    for (name, out) in [
        ("plain", ir.plain_text()),
        ("markdown", ir.to_markdown()),
        ("html", ir.to_html()),
    ] {
        for token in ["HEADER_TEXT", "BODY_TEXT", "FOOTER_TEXT"] {
            assert!(out.contains(token), "{name} is missing {token}: {out}");
        }
    }
}

#[test]
fn test_plain_text_does_not_emit_markdown_syntax_as_a_separator() {
    let ir = DocumentIR {
        metadata: Metadata {
            format: DocumentFormat::Pptx,
            ..Default::default()
        },
        sections: (1..=2)
            .map(|i| Section {
                elements: vec![Element::Paragraph(Paragraph {
                    content: vec![span(&format!("SLIDE{i}"))],
                    ..Default::default()
                })],
                ..Default::default()
            })
            .collect(),
        defined_names: Vec::new(),
    };
    let plain = ir.plain_text();
    assert!(plain.contains("SLIDE1") && plain.contains("SLIDE2"));
    assert!(!plain.contains("---"), "plain text leaked markdown syntax: {plain:?}");
    assert!(plain.contains('\u{000C}'), "expected a form-feed separator: {plain:?}");
}

#[test]
fn test_an_image_with_no_source_does_not_render_as_a_broken_reference() {
    let ir = ir_with(vec![Element::Image(Image {
        alt_text: Some("A chart".to_string()),
        ..Default::default()
    })]);
    let md = ir.to_markdown();
    let html = ir.to_html();
    assert!(!md.contains("]()"), "empty markdown link target: {md}");
    assert!(!html.contains("<img"), "img without a src: {html}");
    // The description must survive in some form.
    assert!(md.contains("A chart") && html.contains("A chart"));
}

#[test]
fn test_an_image_with_neither_source_nor_alt_text_renders_nothing() {
    let ir = ir_with(vec![Element::Image(Image::default())]);
    assert_eq!(ir.to_markdown().trim(), "");
    assert_eq!(ir.to_html().trim(), "");
}

// ---------------------------------------------------------------------------
// Regression: an all-whitespace span must not panic
// ---------------------------------------------------------------------------

#[test]
fn test_an_all_whitespace_span_does_not_panic() {
    // Leading and trailing whitespace are split out of the emphasis
    // delimiters. Computing the two spans independently makes them overlap
    // on an all-whitespace run, which inverts the slice range and aborts the
    // process. Found on 14 real corpus files in the 0.1.9 -> 0.1.10 sweep.
    for text in ["   ", "\t", "\n", " \t \n ", ""] {
        for (bold, italic, strike) in [(true, false, false), (false, true, true)] {
            let ir = ir_with(vec![Element::Paragraph(Paragraph {
                content: vec![InlineContent::Text(TextSpan {
                    text: text.to_string(),
                    bold,
                    italic,
                    strikethrough: strike,
                    ..Default::default()
                })],
                ..Default::default()
            })]);
            // Must not panic, and must not invent emphasis around nothing.
            let md = ir.to_markdown();
            assert!(!md.contains("**") && !md.contains("~~"), "emphasised empty text: {md:?}");
        }
    }
}

#[test]
fn test_whitespace_between_two_emphasised_runs_survives() {
    // The whitespace-only run sits between two bold runs; it must neither
    // panic nor swallow the space that separates the words.
    let bold = |t: &str| {
        InlineContent::Text(TextSpan {
            text: t.to_string(),
            bold: true,
            ..Default::default()
        })
    };
    let ir = ir_with(vec![Element::Paragraph(Paragraph {
        content: vec![
            bold("ONE"),
            InlineContent::Text(TextSpan::plain("   ")),
            bold("TWO"),
        ],
        ..Default::default()
    })]);
    let md = ir.to_markdown();
    assert_eq!(md, "**ONE**   **TWO**", "got {md}");
}

// ---------------------------------------------------------------------------
// Untrusted spans in a deserialized IR
// ---------------------------------------------------------------------------

/// `DocumentIR` is `Deserialize`, so its spans can be anything. Laying a
/// table out looped over `row_span x col_span` before clipping either to
/// the grid, so one cell claiming the maximum of both was ~10^19 iterations.
#[test]
fn test_maximal_row_and_column_spans_render_promptly() {
    let cell = |text: &str, col_span: u32, row_span: u32| TableCell {
        content: vec![Element::Paragraph(Paragraph {
            content: vec![span(text)],
            ..Default::default()
        })],
        col_span,
        row_span,
        ..Default::default()
    };
    let ir = ir_with(vec![Element::Table(Table {
        rows: vec![
            TableRow {
                cells: vec![cell("A", u32::MAX, u32::MAX), cell("B", 1, 1)],
                ..Default::default()
            },
            TableRow {
                cells: vec![cell("C", 1, 1)],
                ..Default::default()
            },
        ],
        ..Default::default()
    })]);
    let md = ir.to_markdown();
    let text = ir.plain_text();
    for t in [&md, &text] {
        assert!(t.contains('A') && t.contains('B') && t.contains('C'), "content lost: {t:?}");
    }
}

// ---------------------------------------------------------------------------
// Recursion bound
// ---------------------------------------------------------------------------

/// Nesting far past anything a bounded reader produces. `DocumentIR` is
/// `Deserialize`, so a value this deep can be built by any caller.
const DEEP: usize = 100_000;

fn deep_text_box() -> Element {
    let mut inner = Element::Paragraph(Paragraph {
        content: vec![span("leaf")],
        ..Default::default()
    });
    for _ in 0..DEEP {
        inner = Element::TextBox(TextBox {
            content: vec![inner],
            ..Default::default()
        });
    }
    inner
}

fn deep_list() -> Element {
    let mut list = List {
        items: vec![ListItem {
            content: vec![Element::Paragraph(Paragraph {
                content: vec![span("leaf")],
                ..Default::default()
            })],
            nested: None,
        }],
        ..Default::default()
    };
    for _ in 0..DEEP {
        list = List {
            items: vec![ListItem {
                content: vec![Element::Paragraph(Paragraph {
                    content: vec![span("item")],
                    ..Default::default()
                })],
                nested: Some(list),
            }],
            ..Default::default()
        };
    }
    Element::List(list)
}

fn deep_table() -> Element {
    let mut inner = Element::Paragraph(Paragraph {
        content: vec![span("leaf")],
        ..Default::default()
    });
    for _ in 0..DEEP {
        inner = Element::Table(Table {
            rows: vec![TableRow {
                cells: vec![TableCell {
                    content: vec![inner],
                    col_span: 1,
                    row_span: 1,
                    ..Default::default()
                }],
                ..Default::default()
            }],
            ..Default::default()
        });
    }
    inner
}

/// Every renderer walks the element tree recursively. Without a depth
/// bound, a deeply nested IR overflowed the stack — an abort no caller can
/// catch. The walk must stop at the shared nesting limit, still render the
/// content above it, and record the truncation where a caller can see it.
#[test]
fn test_renderers_bound_recursion_on_deeply_nested_ir() {
    // A default-sized (2 MiB) thread: the renderers run on the caller's
    // stack, which this crate does not control.
    let handle = std::thread::Builder::new()
        .stack_size(2 * 1024 * 1024)
        .spawn(|| {
            for (name, element) in [
                ("text box", deep_text_box()),
                ("list", deep_list()),
                ("table", deep_table()),
            ] {
                let ir = ir_with(vec![element]);
                office_oxide::core::xml::reset_truncated_subtrees();
                let plain = ir.plain_text();
                let md = ir.to_markdown();
                let html = ir.to_html();
                assert!(
                    office_oxide::core::xml::truncated_subtrees() > 0,
                    "{name}: truncation was not recorded"
                );
                for (surface, out) in [("plain", &plain), ("markdown", &md), ("html", &html)] {
                    // The leaf sits past the bound, so it must not appear.
                    assert!(!out.contains("leaf"), "{name}/{surface}: rendered past the bound");
                    if name == "list" {
                        assert!(out.contains("item"), "{name}/{surface}: lost shallow content");
                    }
                }
            }
        })
        .expect("spawn");
    handle.join().expect("renderer thread panicked");
}

/// The IR→PPTX/XLSX/DOCX writers walk the same trees and run on the
/// caller's stack too; they must stop at the same bound.
#[test]
fn test_ir_writers_bound_recursion_on_deeply_nested_ir() {
    let handle = std::thread::Builder::new()
        .stack_size(2 * 1024 * 1024)
        .spawn(|| {
            for (name, element) in [
                ("text box", deep_text_box()),
                ("list", deep_list()),
                ("table", deep_table()),
            ] {
                let mut ir = ir_with(vec![element]);
                // Speaker notes are walked by the PPTX notes writer.
                ir.sections[0].speaker_notes = Some(vec![deep_list()]);
                for format in [
                    DocumentFormat::Pptx,
                    DocumentFormat::Xlsx,
                    DocumentFormat::Docx,
                ] {
                    let mut buf = std::io::Cursor::new(Vec::new());
                    office_oxide::create::create_from_ir_to_writer(&ir, format, &mut buf)
                        .unwrap_or_else(|e| panic!("{name}/{format:?}: {e}"));
                    assert!(!buf.get_ref().is_empty(), "{name}/{format:?}: empty package");
                }
            }
        })
        .expect("spawn");
    handle.join().expect("writer thread panicked");
}
