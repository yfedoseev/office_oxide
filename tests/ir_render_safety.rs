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
    }
}

// ---------------------------------------------------------------------------
// #157 — URL scheme filtering
// ---------------------------------------------------------------------------

#[test]
fn dangerous_url_schemes_are_dropped_from_links() {
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
fn a_control_character_cannot_smuggle_a_scheme_past_the_filter() {
    let ir = ir_with(vec![Element::Paragraph(Paragraph {
        content: vec![linked("x", "java\u{0}script:alert(1)")],
        ..Default::default()
    })]);
    assert!(!ir.to_html().contains("href"));
}

#[test]
fn ordinary_urls_still_render_as_links() {
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
// #157 — markdown metacharacters in document text
// ---------------------------------------------------------------------------

#[test]
fn document_text_cannot_inject_markdown_structure() {
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
// #166 — the three renderers must agree
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
fn all_three_renderers_include_headers_and_footers() {
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
fn plain_text_does_not_emit_markdown_syntax_as_a_separator() {
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
    };
    let plain = ir.plain_text();
    assert!(plain.contains("SLIDE1") && plain.contains("SLIDE2"));
    assert!(!plain.contains("---"), "plain text leaked markdown syntax: {plain:?}");
    assert!(plain.contains('\u{000C}'), "expected a form-feed separator: {plain:?}");
}

#[test]
fn an_image_with_no_source_does_not_render_as_a_broken_reference() {
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
fn an_image_with_neither_source_nor_alt_text_renders_nothing() {
    let ir = ir_with(vec![Element::Image(Image::default())]);
    assert_eq!(ir.to_markdown().trim(), "");
    assert_eq!(ir.to_html().trim(), "");
}

// ---------------------------------------------------------------------------
// Regression: an all-whitespace span must not panic
// ---------------------------------------------------------------------------

#[test]
fn an_all_whitespace_span_does_not_panic() {
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
fn whitespace_between_two_emphasised_runs_survives() {
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
