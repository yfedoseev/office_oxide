//! Generate a corpus of documents for the OOXML validation gate.
//!
//! Usage: `cargo run --example gen_validation_corpus -- <out-dir>`
//!
//! Deliberately exercises the builder APIs as well as `create_from_markdown`:
//! several defects found in the 0.1.11 sweep were unreachable from markdown,
//! and one (`w:pPr` element order) needed two properties set on the *same*
//! paragraph, which a one-property-at-a-time matrix never produced.

use office_oxide::create::{create_from_ir, create_from_markdown};
use office_oxide::docx::write::{DocxWriter, HfType, Run as DRun};
use office_oxide::format::DocumentFormat;
use office_oxide::ir::*;
use office_oxide::pptx::write::PptxWriter;
use office_oxide::xlsx::write::{CellData, CellStyle, XlsxWriter};

const PNG: &[u8] = &[
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
    0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, 0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
    0x00, 0x03, 0x01, 0x01, 0x00, 0x18, 0xDD, 0x8D, 0xB0, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E,
    0x44, 0xAE, 0x42, 0x60, 0x82,
];

/// Content that has broken XML escaping before: entities, CDATA close,
/// a processing instruction, RTL, CJK, emoji, ZWJ, BOM and NBSP.
const HOSTILE: &str = "amp & lt < gt > quot \" apos ' cdata ]]> pi <?xml ?> \
comment <!-- --> cjk \u{65E5}\u{672C}\u{8A9E} rtl \u{5E9}\u{5DC}\u{5D5}\u{5DD} \
emoji \u{1F600} zwj \u{200D} bom \u{FEFF} nbsp \u{00A0}";

const RICH_MD: &str = "\
# Title

Intro with **bold**, *italic*, `code` and a [link](https://example.com/a?b=1&c=2).

## Second level

- one
  - one-a
    - one-a-i
- two

1. first
2. second

| Col A | Col B |
|---|---|
| a1 | b1 |

```rust
fn main() { println!(\"hi\"); }
```

---

Tail paragraph.
";

fn span(text: &str) -> InlineContent {
    InlineContent::Text(TextSpan {
        text: text.into(),
        ..Default::default()
    })
}

fn cell(text: &str) -> TableCell {
    TableCell {
        content: vec![Element::Paragraph(Paragraph {
            content: vec![span(text)],
            ..Default::default()
        })],
        ..Default::default()
    }
}

/// A document that sets *every* property of the three order-sensitive
/// property groups at once.
///
/// The schema-invalid element orders in 0.1.11 (`w:pPr`, `w:tblPr`, `w:tcPr`)
/// were all found by a 6,000-file real corpus, not by this gate — because a
/// generator that sets one property at a time can never produce the pair that
/// trips an ordering rule. Only `CT_PPrBase`, `CT_TblPrBase`, `CT_TcPrBase`
/// and `CT_SectPr` are sequences in WML; `EG_RPrBase` and `CT_TrPrBase` are
/// choices and order-free. So this covers the whole order-sensitive surface.
fn maximal_properties_corpus(out: &str) {
    let border = || {
        Some(TableBorder {
            top: Some(BorderLine {
                style: BorderStyle::Single,
                color: Some([0, 0, 0]),
                size: Some(4),
                space: Some(0),
            }),
            bottom: None,
            left: None,
            right: None,
            inside_h: None,
            inside_v: None,
        })
    };

    // Every CT_TcPrBase child we can emit, on one cell.
    let cell = TableCell {
        content: vec![Element::Paragraph(Paragraph {
            content: vec![span("cell")],
            ..Default::default()
        })],
        col_span: 2,
        row_span: 2,
        width_twips: Some(1000),
        background_color: Some([0xEE, 0xEE, 0xEE]),
        border: border(),
        padding: Some(CellPadding {
            top_twips: Some(10),
            left_twips: Some(10),
            bottom_twips: Some(10),
            right_twips: Some(10),
        }),
        vertical_align: Some(CellVerticalAlign::Center),
        text_direction: Some(TextDirection::TbRl),
        text_align: Some(ParagraphAlignment::Center),
        ..Default::default()
    };

    // Every CT_TblPrBase child we can emit, on one table.
    let table = Table {
        caption: Some("Caption".into()),
        alignment: Some(TableAlignment::Center),
        indent_left_twips: Some(120),
        cell_padding_twips: Some(60),
        width_twips: Some(8000),
        border: border(),
        column_widths_twips: vec![4000, 4000],
        rows: vec![
            TableRow {
                cells: vec![cell.clone(), cell.clone()],
                is_header: true,
                height_twips: Some(400),
                repeat_as_header: true,
                allow_break: false,
            },
            TableRow {
                cells: vec![cell.clone(), cell],
                ..Default::default()
            },
        ],
    };

    // Every CT_PPrBase child we can emit, on one paragraph.
    let para = Paragraph {
        content: vec![span("para")],
        alignment: Some(ParagraphAlignment::Justify),
        indent_left_twips: Some(720),
        indent_right_twips: Some(360),
        first_line_indent_twips: Some(-360),
        space_before_twips: Some(120),
        space_after_twips: Some(240),
        line_spacing: Some(LineSpacing::Exact(360)),
        background_color: Some([0xFF, 0xFF, 0xCC]),
        border: Some(ParagraphBorder {
            top: Some(BorderLine {
                style: BorderStyle::Single,
                color: Some([0, 0, 0]),
                size: Some(4),
                space: Some(1),
            }),
            bottom: None,
            left: None,
            right: None,
            between: None,
        }),
        tabs: vec![TabStop {
            position_twips: 8640,
            alignment: TabAlignment::Right,
            leader: TabLeader::Dot,
        }],
        keep_with_next: true,
        keep_together: true,
        page_break_before: true,
        outline_level: Some(1),
        frame_position: Some(FramePosition {
            x_twips: 100,
            y_twips: 200,
            width_twips: 3000,
            height_twips: 400,
        }),
    };

    let ir = DocumentIR {
        sections: vec![Section {
            title: Some("Maximal".into()),
            elements: vec![Element::Paragraph(para), Element::Table(table)],
            page_setup: Some(PageSetup::default()),
            background_rgb: Some([0x11, 0x22, 0x33]),
            speaker_notes: Some("notes".into()),
            ..Default::default()
        }],
        ..Default::default()
    };

    for (fmt, ext) in [
        (DocumentFormat::Docx, "docx"),
        (DocumentFormat::Xlsx, "xlsx"),
        (DocumentFormat::Pptx, "pptx"),
    ] {
        create_from_ir(&ir, fmt, format!("{out}/maximal.{ext}")).unwrap();
    }
}

fn markdown_corpus(out: &str) {
    for (name, md) in [
        ("minimal", "# Slide one\n\nHello world.\n"),
        ("rich", RICH_MD),
        ("empty", ""),
        ("textonly", "Just a bare paragraph.\n"),
        ("hostile", HOSTILE),
    ] {
        for (fmt, ext) in [
            (DocumentFormat::Docx, "docx"),
            (DocumentFormat::Xlsx, "xlsx"),
            (DocumentFormat::Pptx, "pptx"),
        ] {
            create_from_markdown(md, fmt, format!("{out}/md_{name}.{ext}")).unwrap();
        }
    }
}

fn docx_builder_corpus(out: &str) {
    let mut w = DocxWriter::new();
    w.set_metadata(&Metadata {
        title: Some(HOSTILE.into()),
        author: Some("A".into()),
        ..Default::default()
    });
    w.add_heading("Heading", 1);
    w.add_rich_paragraph(&[DRun::new("bold").bold(), DRun::new(" red").color("FF0000")]);
    // Both an indent and spacing on ONE paragraph: the pPr ordering defect
    // needed the pair to co-occur.
    w.add_ir_paragraph(
        &[DRun::new("indented and spaced")],
        Some(office_oxide::docx::write::IrParaProps {
            indent_left_twips: Some(720),
            space_after_twips: Some(240),
            ..Default::default()
        }),
    );
    w.add_table(&[vec!["a", "b"], vec!["c", "d"]]);
    w.add_ir_table(&Table {
        rows: vec![TableRow {
            cells: vec![cell("x"), cell("y")],
            ..Default::default()
        }],
        ..Default::default()
    });
    w.add_list(&["one", "two"], false);
    w.add_ir_list(&List {
        ordered: true,
        items: vec![ListItem {
            content: inline_to_element_block(vec![span("ordered")]),
            nested: Some(List {
                items: vec![ListItem {
                    content: inline_to_element_block(vec![span("nested")]),
                    nested: None,
                }],
                ..Default::default()
            }),
        }],
        ..Default::default()
    });
    w.add_page_break();
    w.add_code_block("fn main() {}");
    w.add_ir_image(&Image {
        data: Some(PNG.to_vec()),
        format: Some(ImageFormat::Png),
        display_width_emu: Some(500_000),
        display_height_emu: Some(500_000),
        alt_text: Some("alt".into()),
        ..Default::default()
    });
    w.add_text_box(&TextBox::default());
    w.add_footnote(1, &[Element::Paragraph(Default::default())]);
    w.add_endnote(1, &[Element::Paragraph(Default::default())]);
    for t in [
        HfType::DefaultHeader,
        HfType::DefaultFooter,
        HfType::FirstPageHeader,
        HfType::EvenPageHeader,
    ] {
        w.add_section_header(t, vec![Element::Paragraph(Default::default())]);
    }
    w.set_section_props(Some(PageSetup::default()), None, SectionBreakType::NextPage);
    w.save(format!("{out}/builder.docx")).unwrap();
}

fn pptx_builder_corpus(out: &str) {
    let mut p = PptxWriter::new();
    p.set_presentation_size(9_144_000, 6_858_000);
    {
        let s = p.add_slide();
        s.set_title(HOSTILE);
        s.add_text("body");
        s.add_bullet_list(&["x", "y"]);
        s.add_text_box("boxed", 100_000, 100_000, 2_000_000, 500_000);
        s.add_image(PNG.to_vec(), ImageFormat::Png, 0, 900_000, 500_000, 500_000);
        s.set_notes("presenter only");
        s.add_table(vec![vec!["A".into(), "B".into()], vec!["1".into(), "2".into()]]);
    }
    p.add_slide();
    p.save(format!("{out}/builder.pptx")).unwrap();
}

fn xlsx_builder_corpus(out: &str) {
    let mut x = XlsxWriter::new();
    {
        let mut s = x.add_sheet(HOSTILE);
        s.add_row(vec![
            CellData::String(HOSTILE.into()),
            CellData::Number(1.5),
            CellData::Boolean(true),
            CellData::Formula("SUM(B1:B2)".into()),
        ]);
        s.set_cell_styled(
            1,
            0,
            CellData::Number(2.0),
            CellStyle::new().bold().background("FFFF00").font_size(10.5),
        );
        s.merge_cells(2, 0, 2, 2);
        s.set_column_width(0, 20.0);
        s.add_image(PNG.to_vec(), "png", 0, 0, 500_000, 500_000);
    }
    x.add_sheet("Second");
    x.save(format!("{out}/builder.xlsx")).unwrap();
}

fn conversion_corpus(out: &str) {
    // save_as across all nine pairs: the w:pgMar gutter defect was only
    // reachable through a section that carries a page setup.
    for src in ["docx", "xlsx", "pptx"] {
        let path = format!("{out}/md_rich.{src}");
        let Ok(doc) = office_oxide::Document::open(&path) else {
            continue;
        };
        for dst in ["docx", "xlsx", "pptx"] {
            let _ = doc.save_as(format!("{out}/conv_{src}_to_{dst}.{dst}"));
        }
    }
}

fn extremes_corpus(out: &str) {
    // Values that are legal in Rust and out of range in OOXML.
    let mut p = PptxWriter::new();
    p.set_presentation_size(1000, 99_000_000);
    {
        let s = p.add_slide();
        s.add_rich_text(&[office_oxide::pptx::write::Run::new("x").font_size(f64::NAN)]);
        s.add_rich_text(&[office_oxide::pptx::write::Run::new("y").color("#FF0000")]);
        s.add_text_box("neg", 0, 0, -100, -100);
    }
    p.save(format!("{out}/extremes.pptx")).unwrap();

    let mut x = XlsxWriter::new();
    {
        let mut s = x.add_sheet("ctl\u{1}chr");
        s.add_row(vec![CellData::String("x".repeat(40_000))]);
        s.set_cell_styled(0, 1, CellData::Empty, CellStyle::new().bold());
        s.set_column_width(1, f64::INFINITY);
        s.merge_cells(0, 0, usize::MAX, 2);
    }
    x.save(format!("{out}/extremes.xlsx")).unwrap();

    let ir = DocumentIR {
        sections: vec![Section {
            elements: vec![Element::Paragraph(Paragraph {
                content: vec![span("tabbed")],
                tabs: vec![TabStop {
                    position_twips: 8640,
                    alignment: TabAlignment::Right,
                    leader: TabLeader::Dot,
                }],
                ..Default::default()
            })],
            ..Default::default()
        }],
        ..Default::default()
    };
    create_from_ir(&ir, DocumentFormat::Docx, format!("{out}/extremes.docx")).unwrap();
}

fn main() {
    let out = std::env::args()
        .nth(1)
        .expect("usage: gen_validation_corpus <out-dir>");
    std::fs::create_dir_all(&out).expect("create out dir");
    markdown_corpus(&out);
    docx_builder_corpus(&out);
    pptx_builder_corpus(&out);
    xlsx_builder_corpus(&out);
    conversion_corpus(&out);
    extremes_corpus(&out);
    maximal_properties_corpus(&out);
    let n = std::fs::read_dir(&out).unwrap().count();
    println!("wrote {n} packages to {out}");
}
