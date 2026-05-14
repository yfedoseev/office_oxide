use std::io::Cursor;

// ===========================================================================
// DOCX Writer round-trip tests
// ===========================================================================

#[test]
fn docx_write_paragraph_round_trip() {
    let mut writer = office_oxide::docx::write::DocxWriter::new();
    writer.add_paragraph("Hello world");
    writer.add_paragraph("Second paragraph");

    let mut buf = Cursor::new(Vec::new());
    writer.write_to(&mut buf).unwrap();
    buf.set_position(0);

    let doc = office_oxide::docx::DocxDocument::from_reader(buf).unwrap();
    let text = doc.plain_text();
    assert!(text.contains("Hello world"), "text: {text}");
    assert!(text.contains("Second paragraph"), "text: {text}");
}

#[test]
fn docx_write_heading_round_trip() {
    let mut writer = office_oxide::docx::write::DocxWriter::new();
    writer.add_heading("Chapter One", 1);
    writer.add_paragraph("Content here");

    let mut buf = Cursor::new(Vec::new());
    writer.write_to(&mut buf).unwrap();
    buf.set_position(0);

    let doc = office_oxide::docx::DocxDocument::from_reader(buf).unwrap();
    let text = doc.plain_text();
    assert!(text.contains("Chapter One"), "text: {text}");
    assert!(text.contains("Content here"), "text: {text}");
}

#[test]
fn docx_write_table_round_trip() {
    let mut writer = office_oxide::docx::write::DocxWriter::new();
    writer.add_table(&[vec!["Name", "Age"], vec!["Alice", "30"]]);

    let mut buf = Cursor::new(Vec::new());
    writer.write_to(&mut buf).unwrap();
    buf.set_position(0);

    let doc = office_oxide::docx::DocxDocument::from_reader(buf).unwrap();
    let text = doc.plain_text();
    assert!(text.contains("Name"), "text: {text}");
    assert!(text.contains("Alice"), "text: {text}");
    assert!(text.contains("30"), "text: {text}");
}

#[test]
fn docx_write_list_round_trip() {
    let mut writer = office_oxide::docx::write::DocxWriter::new();
    writer.add_list(&["First", "Second", "Third"], false);

    let mut buf = Cursor::new(Vec::new());
    writer.write_to(&mut buf).unwrap();
    buf.set_position(0);

    let doc = office_oxide::docx::DocxDocument::from_reader(buf).unwrap();
    let text = doc.plain_text();
    assert!(text.contains("First"), "text: {text}");
    assert!(text.contains("Second"), "text: {text}");
    assert!(text.contains("Third"), "text: {text}");
}

// ===========================================================================
// XLSX Writer round-trip tests
// ===========================================================================

#[test]
fn xlsx_write_cells_round_trip() {
    let mut writer = office_oxide::xlsx::write::XlsxWriter::new();
    {
        let mut sheet = writer.add_sheet("Data");
        sheet.add_row(vec![
            office_oxide::xlsx::write::CellData::String("Name".into()),
            office_oxide::xlsx::write::CellData::String("Score".into()),
        ]);
        sheet.add_row(vec![
            office_oxide::xlsx::write::CellData::String("Alice".into()),
            office_oxide::xlsx::write::CellData::Number(95.5),
        ]);
        sheet.add_row(vec![
            office_oxide::xlsx::write::CellData::String("Bob".into()),
            office_oxide::xlsx::write::CellData::Boolean(true),
        ]);
    }

    let mut buf = Cursor::new(Vec::new());
    writer.write_to(&mut buf).unwrap();
    buf.set_position(0);

    let doc = office_oxide::xlsx::XlsxDocument::from_reader(buf).unwrap();
    let text = doc.plain_text();
    assert!(text.contains("Name"), "text: {text}");
    assert!(text.contains("Alice"), "text: {text}");
    assert!(text.contains("95.5"), "text: {text}");
    assert!(text.contains("TRUE"), "text: {text}");
}

#[test]
fn xlsx_write_multiple_sheets_round_trip() {
    let mut writer = office_oxide::xlsx::write::XlsxWriter::new();
    {
        let mut s1 = writer.add_sheet("Sheet1");
        s1.add_row(vec![office_oxide::xlsx::write::CellData::String("A1".into())]);
    }
    {
        let mut s2 = writer.add_sheet("Sheet2");
        s2.add_row(vec![office_oxide::xlsx::write::CellData::String("B1".into())]);
    }

    let mut buf = Cursor::new(Vec::new());
    writer.write_to(&mut buf).unwrap();
    buf.set_position(0);

    let doc = office_oxide::xlsx::XlsxDocument::from_reader(buf).unwrap();
    assert_eq!(doc.worksheets.len(), 2);
    let text = doc.plain_text();
    assert!(text.contains("A1"), "text: {text}");
    assert!(text.contains("B1"), "text: {text}");
}

#[test]
fn xlsx_write_empty_cells_round_trip() {
    let mut writer = office_oxide::xlsx::write::XlsxWriter::new();
    {
        let mut sheet = writer.add_sheet("Sparse");
        sheet.add_row(vec![
            office_oxide::xlsx::write::CellData::String("Start".into()),
            office_oxide::xlsx::write::CellData::Empty,
            office_oxide::xlsx::write::CellData::String("End".into()),
        ]);
    }

    let mut buf = Cursor::new(Vec::new());
    writer.write_to(&mut buf).unwrap();
    buf.set_position(0);

    let doc = office_oxide::xlsx::XlsxDocument::from_reader(buf).unwrap();
    let text = doc.plain_text();
    assert!(text.contains("Start"), "text: {text}");
    assert!(text.contains("End"), "text: {text}");
}

// ===========================================================================
// PPTX Writer round-trip tests
// ===========================================================================

#[test]
fn pptx_write_slide_round_trip() {
    let mut writer = office_oxide::pptx::write::PptxWriter::new();
    {
        let slide = writer.add_slide();
        slide.set_title("My Title");
        slide.add_text("Some content");
    }

    let mut buf = Cursor::new(Vec::new());
    writer.write_to(&mut buf).unwrap();
    buf.set_position(0);

    let doc = office_oxide::pptx::PptxDocument::from_reader(buf).unwrap();
    let text = doc.plain_text();
    assert!(text.contains("My Title"), "text: {text}");
    assert!(text.contains("Some content"), "text: {text}");
}

#[test]
fn pptx_write_multiple_slides_round_trip() {
    let mut writer = office_oxide::pptx::write::PptxWriter::new();
    {
        let s1 = writer.add_slide();
        s1.set_title("Slide 1");
        s1.add_text("First slide content");
    }
    {
        let s2 = writer.add_slide();
        s2.set_title("Slide 2");
        s2.add_text("Second slide content");
    }

    let mut buf = Cursor::new(Vec::new());
    writer.write_to(&mut buf).unwrap();
    buf.set_position(0);

    let doc = office_oxide::pptx::PptxDocument::from_reader(buf).unwrap();
    assert_eq!(doc.slides.len(), 2);
    let text = doc.plain_text();
    assert!(text.contains("Slide 1"), "text: {text}");
    assert!(text.contains("Second slide content"), "text: {text}");
}

#[test]
fn pptx_write_bullet_list_round_trip() {
    let mut writer = office_oxide::pptx::write::PptxWriter::new();
    {
        let slide = writer.add_slide();
        slide.set_title("Bullets");
        slide.add_bullet_list(&["Item A", "Item B", "Item C"]);
    }

    let mut buf = Cursor::new(Vec::new());
    writer.write_to(&mut buf).unwrap();
    buf.set_position(0);

    let doc = office_oxide::pptx::PptxDocument::from_reader(buf).unwrap();
    let text = doc.plain_text();
    assert!(text.contains("Item A"), "text: {text}");
    assert!(text.contains("Item B"), "text: {text}");
    assert!(text.contains("Item C"), "text: {text}");
}

// ===========================================================================
// create_from_ir round-trip tests
// ===========================================================================

fn sample_ir(format: office_oxide::DocumentFormat) -> office_oxide::DocumentIR {
    use office_oxide::ir::*;

    DocumentIR {
        metadata: Metadata {
            format,
            title: Some("Test Doc".to_string()),
            ..Default::default()
        },
        sections: vec![Section {
            title: Some("Section 1".to_string()),
            elements: vec![
                Element::Heading(Heading {
                    level: 1,
                    content: vec![InlineContent::Text(TextSpan::plain("Main Heading"))],
                    ..Default::default()
                }),
                Element::Paragraph(Paragraph {
                    content: vec![InlineContent::Text(TextSpan::plain("Body text here"))],
                    ..Default::default()
                }),
                Element::Table(Table {
                    rows: vec![
                        TableRow {
                            cells: vec![
                                TableCell {
                                    content: vec![Element::Paragraph(Paragraph {
                                        content: vec![InlineContent::Text(TextSpan::plain("H1"))],
                                        ..Default::default()
                                    })],
                                    col_span: 1,
                                    row_span: 1,
                                    ..Default::default()
                                },
                                TableCell {
                                    content: vec![Element::Paragraph(Paragraph {
                                        content: vec![InlineContent::Text(TextSpan::plain("H2"))],
                                        ..Default::default()
                                    })],
                                    col_span: 1,
                                    row_span: 1,
                                    ..Default::default()
                                },
                            ],
                            is_header: true,
                            ..Default::default()
                        },
                        TableRow {
                            cells: vec![
                                TableCell {
                                    content: vec![Element::Paragraph(Paragraph {
                                        content: vec![InlineContent::Text(TextSpan::plain("A"))],
                                        ..Default::default()
                                    })],
                                    col_span: 1,
                                    row_span: 1,
                                    ..Default::default()
                                },
                                TableCell {
                                    content: vec![Element::Paragraph(Paragraph {
                                        content: vec![InlineContent::Text(TextSpan::plain("B"))],
                                        ..Default::default()
                                    })],
                                    col_span: 1,
                                    row_span: 1,
                                    ..Default::default()
                                },
                            ],
                            is_header: false,
                            ..Default::default()
                        },
                    ],
                    ..Default::default()
                }),
            ],
            ..Default::default()
        }],
    }
}

#[test]
fn create_from_ir_docx_round_trip() {
    let ir = sample_ir(office_oxide::DocumentFormat::Docx);

    let mut buf = Cursor::new(Vec::new());
    office_oxide::create::create_from_ir_to_writer(
        &ir,
        office_oxide::DocumentFormat::Docx,
        &mut buf,
    )
    .unwrap();
    buf.set_position(0);

    let doc = office_oxide::Document::from_reader(buf, office_oxide::DocumentFormat::Docx).unwrap();
    let text = doc.plain_text();
    assert!(text.contains("Main Heading"), "text: {text}");
    assert!(text.contains("Body text here"), "text: {text}");
}

#[test]
fn create_from_ir_xlsx_round_trip() {
    let ir = sample_ir(office_oxide::DocumentFormat::Xlsx);

    let mut buf = Cursor::new(Vec::new());
    office_oxide::create::create_from_ir_to_writer(
        &ir,
        office_oxide::DocumentFormat::Xlsx,
        &mut buf,
    )
    .unwrap();
    buf.set_position(0);

    let doc = office_oxide::Document::from_reader(buf, office_oxide::DocumentFormat::Xlsx).unwrap();
    let text = doc.plain_text();
    assert!(text.contains("H1"), "text: {text}");
    assert!(text.contains("A"), "text: {text}");
}

#[test]
fn create_from_ir_pptx_round_trip() {
    let ir = sample_ir(office_oxide::DocumentFormat::Pptx);

    let mut buf = Cursor::new(Vec::new());
    office_oxide::create::create_from_ir_to_writer(
        &ir,
        office_oxide::DocumentFormat::Pptx,
        &mut buf,
    )
    .unwrap();
    buf.set_position(0);

    let doc = office_oxide::Document::from_reader(buf, office_oxide::DocumentFormat::Pptx).unwrap();
    let text = doc.plain_text();
    assert!(text.contains("Main Heading"), "text: {text}");
    assert!(text.contains("Body text here"), "text: {text}");
}

// ===========================================================================
// Edit round-trip tests
// ===========================================================================

#[test]
fn docx_edit_replace_text_round_trip() {
    // Create a docx, then edit it
    let mut writer = office_oxide::docx::write::DocxWriter::new();
    writer.add_paragraph("Hello PLACEHOLDER world");

    let mut buf = Cursor::new(Vec::new());
    writer.write_to(&mut buf).unwrap();
    let bytes = buf.into_inner();

    let mut editable =
        office_oxide::docx::edit::EditableDocx::from_reader(Cursor::new(bytes.clone())).unwrap();
    let count = editable.replace_text("PLACEHOLDER", "beautiful");
    assert!(count > 0, "should replace at least once");

    let mut out = Cursor::new(Vec::new());
    editable.write_to(&mut out).unwrap();
    out.set_position(0);

    let doc = office_oxide::docx::DocxDocument::from_reader(out).unwrap();
    let text = doc.plain_text();
    assert!(text.contains("beautiful"), "text: {text}");
    assert!(!text.contains("PLACEHOLDER"), "text: {text}");
}

#[test]
fn xlsx_edit_set_cell_round_trip() {
    let mut writer = office_oxide::xlsx::write::XlsxWriter::new();
    {
        let mut sheet = writer.add_sheet("Sheet1");
        sheet.add_row(vec![office_oxide::xlsx::write::CellData::String(
            "Original".into(),
        )]);
    }

    let mut buf = Cursor::new(Vec::new());
    writer.write_to(&mut buf).unwrap();
    let bytes = buf.into_inner();

    let mut editable =
        office_oxide::xlsx::edit::EditableXlsx::from_reader(Cursor::new(bytes)).unwrap();
    editable
        .set_cell(0, "A1", office_oxide::xlsx::edit::CellValue::String("Modified".into()))
        .unwrap();

    let mut out = Cursor::new(Vec::new());
    editable.write_to(&mut out).unwrap();
    out.set_position(0);

    let doc = office_oxide::xlsx::XlsxDocument::from_reader(out).unwrap();
    let text = doc.plain_text();
    assert!(text.contains("Modified"), "text: {text}");
}

#[test]
fn pptx_edit_replace_text_round_trip() {
    let mut writer = office_oxide::pptx::write::PptxWriter::new();
    {
        let slide = writer.add_slide();
        slide.set_title("Hello MARKER");
        slide.add_text("Some MARKER content");
    }

    let mut buf = Cursor::new(Vec::new());
    writer.write_to(&mut buf).unwrap();
    let bytes = buf.into_inner();

    let mut editable =
        office_oxide::pptx::edit::EditablePptx::from_reader(Cursor::new(bytes)).unwrap();
    let count = editable.replace_text("MARKER", "REPLACED");
    assert!(count > 0, "should replace at least once");

    let mut out = Cursor::new(Vec::new());
    editable.write_to(&mut out).unwrap();
    out.set_position(0);

    let doc = office_oxide::pptx::PptxDocument::from_reader(out).unwrap();
    let text = doc.plain_text();
    assert!(text.contains("REPLACED"), "text: {text}");
}

// ===========================================================================
// P0 rich IR → DOCX round-trip tests
// ===========================================================================

#[test]
fn ir_paragraph_alignment_round_trip() {
    use office_oxide::ir::*;

    let ir = DocumentIR {
        metadata: Metadata {
            format: office_oxide::DocumentFormat::Docx,
            title: None,
            ..Default::default()
        },
        sections: vec![Section {
            elements: vec![
                Element::Paragraph(Paragraph {
                    content: vec![InlineContent::Text(TextSpan::plain("Left aligned"))],
                    alignment: Some(ParagraphAlignment::Left),
                    ..Default::default()
                }),
                Element::Paragraph(Paragraph {
                    content: vec![InlineContent::Text(TextSpan::plain("Center aligned"))],
                    alignment: Some(ParagraphAlignment::Center),
                    ..Default::default()
                }),
                Element::Paragraph(Paragraph {
                    content: vec![InlineContent::Text(TextSpan::plain("Justified text"))],
                    alignment: Some(ParagraphAlignment::Justify),
                    ..Default::default()
                }),
            ],
            ..Default::default()
        }],
    };

    let mut buf = Cursor::new(Vec::new());
    office_oxide::create::create_from_ir_to_writer(
        &ir,
        office_oxide::DocumentFormat::Docx,
        &mut buf,
    )
    .unwrap();
    buf.set_position(0);

    let doc = office_oxide::docx::DocxDocument::from_reader(buf).unwrap();
    let text = doc.plain_text();
    assert!(text.contains("Left aligned"), "text: {text}");
    assert!(text.contains("Center aligned"), "text: {text}");
    assert!(text.contains("Justified text"), "text: {text}");
}

#[test]
fn ir_paragraph_indentation_round_trip() {
    use office_oxide::ir::*;

    let ir = DocumentIR {
        metadata: Metadata {
            format: office_oxide::DocumentFormat::Docx,
            ..Default::default()
        },
        sections: vec![Section {
            elements: vec![Element::Paragraph(Paragraph {
                content: vec![InlineContent::Text(TextSpan::plain("Indented paragraph"))],
                indent_left_twips: Some(720),
                first_line_indent_twips: Some(360),
                ..Default::default()
            })],
            ..Default::default()
        }],
    };

    let mut buf = Cursor::new(Vec::new());
    office_oxide::create::create_from_ir_to_writer(
        &ir,
        office_oxide::DocumentFormat::Docx,
        &mut buf,
    )
    .unwrap();
    buf.set_position(0);

    let doc = office_oxide::docx::DocxDocument::from_reader(buf).unwrap();
    assert!(doc.plain_text().contains("Indented paragraph"));
}

#[test]
fn ir_paragraph_line_spacing_round_trip() {
    use office_oxide::ir::*;

    let ir = DocumentIR {
        metadata: Metadata {
            format: office_oxide::DocumentFormat::Docx,
            ..Default::default()
        },
        sections: vec![Section {
            elements: vec![Element::Paragraph(Paragraph {
                content: vec![InlineContent::Text(TextSpan::plain("1.5 line spacing"))],
                line_spacing: Some(LineSpacing::Auto(360)), // 360 = 1.5x
                ..Default::default()
            })],
            ..Default::default()
        }],
    };

    let mut buf = Cursor::new(Vec::new());
    office_oxide::create::create_from_ir_to_writer(
        &ir,
        office_oxide::DocumentFormat::Docx,
        &mut buf,
    )
    .unwrap();
    buf.set_position(0);

    let doc = office_oxide::docx::DocxDocument::from_reader(buf).unwrap();
    assert!(doc.plain_text().contains("1.5 line spacing"));
}

#[test]
fn ir_table_with_borders_round_trip() {
    use office_oxide::ir::*;

    let border_line = BorderLine {
        style: BorderStyle::Single,
        color: Some([0, 0, 0]),
        size: Some(4),
        space: Some(0),
    };
    let table_border = TableBorder {
        top: Some(border_line.clone()),
        bottom: Some(border_line.clone()),
        left: Some(border_line.clone()),
        right: Some(border_line.clone()),
        inside_h: Some(border_line.clone()),
        inside_v: Some(border_line.clone()),
    };

    let ir = DocumentIR {
        metadata: Metadata {
            format: office_oxide::DocumentFormat::Docx,
            ..Default::default()
        },
        sections: vec![Section {
            elements: vec![Element::Table(Table {
                rows: vec![
                    TableRow {
                        cells: vec![
                            TableCell {
                                content: vec![Element::Paragraph(Paragraph {
                                    content: vec![InlineContent::Text(TextSpan::plain("Name"))],
                                    ..Default::default()
                                })],
                                col_span: 1,
                                row_span: 1,
                                ..Default::default()
                            },
                            TableCell {
                                content: vec![Element::Paragraph(Paragraph {
                                    content: vec![InlineContent::Text(TextSpan::plain("Value"))],
                                    ..Default::default()
                                })],
                                col_span: 1,
                                row_span: 1,
                                ..Default::default()
                            },
                        ],
                        is_header: true,
                        ..Default::default()
                    },
                    TableRow {
                        cells: vec![
                            TableCell {
                                content: vec![Element::Paragraph(Paragraph {
                                    content: vec![InlineContent::Text(TextSpan::plain("Alice"))],
                                    ..Default::default()
                                })],
                                col_span: 1,
                                row_span: 1,
                                ..Default::default()
                            },
                            TableCell {
                                content: vec![Element::Paragraph(Paragraph {
                                    content: vec![InlineContent::Text(TextSpan::plain("42"))],
                                    ..Default::default()
                                })],
                                col_span: 1,
                                row_span: 1,
                                ..Default::default()
                            },
                        ],
                        is_header: false,
                        ..Default::default()
                    },
                ],
                column_widths_twips: vec![2880, 2880],
                border: Some(table_border),
                ..Default::default()
            })],
            ..Default::default()
        }],
    };

    let mut buf = Cursor::new(Vec::new());
    office_oxide::create::create_from_ir_to_writer(
        &ir,
        office_oxide::DocumentFormat::Docx,
        &mut buf,
    )
    .unwrap();
    buf.set_position(0);

    let doc = office_oxide::docx::DocxDocument::from_reader(buf).unwrap();
    let text = doc.plain_text();
    assert!(text.contains("Name"), "text: {text}");
    assert!(text.contains("Alice"), "text: {text}");
    assert!(text.contains("42"), "text: {text}");
}

#[test]
fn ir_table_with_cell_shading_round_trip() {
    use office_oxide::ir::*;

    let ir = DocumentIR {
        metadata: Metadata {
            format: office_oxide::DocumentFormat::Docx,
            ..Default::default()
        },
        sections: vec![Section {
            elements: vec![Element::Table(Table {
                rows: vec![TableRow {
                    cells: vec![
                        TableCell {
                            content: vec![Element::Paragraph(Paragraph {
                                content: vec![InlineContent::Text(TextSpan::plain("Red cell"))],
                                ..Default::default()
                            })],
                            col_span: 1,
                            row_span: 1,
                            background_color: Some([255, 0, 0]),
                            ..Default::default()
                        },
                        TableCell {
                            content: vec![Element::Paragraph(Paragraph {
                                content: vec![InlineContent::Text(TextSpan::plain("Blue cell"))],
                                ..Default::default()
                            })],
                            col_span: 1,
                            row_span: 1,
                            background_color: Some([0, 0, 255]),
                            ..Default::default()
                        },
                    ],
                    is_header: false,
                    ..Default::default()
                }],
                ..Default::default()
            })],
            ..Default::default()
        }],
    };

    let mut buf = Cursor::new(Vec::new());
    office_oxide::create::create_from_ir_to_writer(
        &ir,
        office_oxide::DocumentFormat::Docx,
        &mut buf,
    )
    .unwrap();
    buf.set_position(0);

    let doc = office_oxide::docx::DocxDocument::from_reader(buf).unwrap();
    let text = doc.plain_text();
    assert!(text.contains("Red cell"), "text: {text}");
    assert!(text.contains("Blue cell"), "text: {text}");
}

#[test]
fn ir_inline_image_round_trip() {
    use office_oxide::ir::*;

    // Minimal 1×1 white PNG (67 bytes)
    let png_bytes: Vec<u8> = vec![
        0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, // PNG signature
        0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52, // IHDR length + type
        0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, // 1x1
        0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53, // 8-bit RGB
        0xde, 0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, // IDAT
        0x54, 0x08, 0xd7, 0x63, 0xf8, 0xcf, 0xc0, 0x00, // compressed pixel
        0x00, 0x00, 0x02, 0x00, 0x01, 0xe2, 0x21, 0xbc, // CRC
        0x33, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, // IEND
        0x44, 0xae, 0x42, 0x60, 0x82, // IEND CRC
    ];

    let ir = DocumentIR {
        metadata: Metadata {
            format: office_oxide::DocumentFormat::Docx,
            ..Default::default()
        },
        sections: vec![Section {
            elements: vec![
                Element::Paragraph(Paragraph {
                    content: vec![InlineContent::Text(TextSpan::plain("Before image"))],
                    ..Default::default()
                }),
                Element::Image(Image {
                    alt_text: Some("test image".to_string()),
                    data: Some(png_bytes),
                    format: Some(ImageFormat::Png),
                    display_width_emu: Some(914400),
                    display_height_emu: Some(914400),
                    ..Default::default()
                }),
                Element::Paragraph(Paragraph {
                    content: vec![InlineContent::Text(TextSpan::plain("After image"))],
                    ..Default::default()
                }),
            ],
            ..Default::default()
        }],
    };

    let mut buf = Cursor::new(Vec::new());
    office_oxide::create::create_from_ir_to_writer(
        &ir,
        office_oxide::DocumentFormat::Docx,
        &mut buf,
    )
    .unwrap();

    let bytes = buf.into_inner();

    // Verify the media part exists in the ZIP
    let cursor = Cursor::new(bytes.clone());
    let mut zip = zip::ZipArchive::new(cursor).unwrap();
    assert!(
        zip.by_name("word/media/image1.png").is_ok(),
        "image1.png should be in the DOCX package"
    );

    // Verify text content
    let doc = office_oxide::docx::DocxDocument::from_reader(Cursor::new(bytes)).unwrap();
    let text = doc.plain_text();
    assert!(text.contains("Before image"), "text: {text}");
    assert!(text.contains("After image"), "text: {text}");
}

#[test]
fn ir_section_page_setup_round_trip() {
    use office_oxide::ir::*;

    let ir = DocumentIR {
        metadata: Metadata {
            format: office_oxide::DocumentFormat::Docx,
            ..Default::default()
        },
        sections: vec![Section {
            elements: vec![Element::Paragraph(Paragraph {
                content: vec![InlineContent::Text(TextSpan::plain("A4 page content"))],
                ..Default::default()
            })],
            page_setup: Some(PageSetup {
                width_twips: 11906,  // A4 width
                height_twips: 16838, // A4 height
                margin_top_twips: 1440,
                margin_bottom_twips: 1440,
                margin_left_twips: 1800,
                margin_right_twips: 1800,
                landscape: false,
                ..Default::default()
            }),
            break_type: SectionBreakType::NextPage,
            ..Default::default()
        }],
    };

    let mut buf = Cursor::new(Vec::new());
    office_oxide::create::create_from_ir_to_writer(
        &ir,
        office_oxide::DocumentFormat::Docx,
        &mut buf,
    )
    .unwrap();
    buf.set_position(0);

    let doc = office_oxide::docx::DocxDocument::from_reader(buf).unwrap();
    assert!(doc.plain_text().contains("A4 page content"));
}

#[test]
fn ir_two_column_section_round_trip() {
    use office_oxide::ir::*;

    let ir = DocumentIR {
        metadata: Metadata {
            format: office_oxide::DocumentFormat::Docx,
            ..Default::default()
        },
        sections: vec![Section {
            elements: vec![Element::Paragraph(Paragraph {
                content: vec![InlineContent::Text(TextSpan::plain("Two column text"))],
                ..Default::default()
            })],
            columns: Some(ColumnLayout {
                count: 2,
                space_twips: Some(720),
                separator: true,
                ..Default::default()
            }),
            ..Default::default()
        }],
    };

    let mut buf = Cursor::new(Vec::new());
    office_oxide::create::create_from_ir_to_writer(
        &ir,
        office_oxide::DocumentFormat::Docx,
        &mut buf,
    )
    .unwrap();
    buf.set_position(0);

    let doc = office_oxide::docx::DocxDocument::from_reader(buf).unwrap();
    assert!(doc.plain_text().contains("Two column text"));
}

#[test]
fn ir_run_typography_round_trip() {
    use office_oxide::ir::*;

    let ir = DocumentIR {
        metadata: Metadata {
            format: office_oxide::DocumentFormat::Docx,
            ..Default::default()
        },
        sections: vec![Section {
            elements: vec![Element::Paragraph(Paragraph {
                content: vec![
                    InlineContent::Text(TextSpan {
                        text: "Big red".to_string(),
                        font_size_half_pt: Some(48), // 24pt
                        color: Some([255, 0, 0]),
                        bold: true,
                        ..Default::default()
                    }),
                    InlineContent::Text(TextSpan {
                        text: " small blue".to_string(),
                        font_size_half_pt: Some(16), // 8pt
                        color: Some([0, 0, 255]),
                        ..Default::default()
                    }),
                ],
                ..Default::default()
            })],
            ..Default::default()
        }],
    };

    let mut buf = Cursor::new(Vec::new());
    office_oxide::create::create_from_ir_to_writer(
        &ir,
        office_oxide::DocumentFormat::Docx,
        &mut buf,
    )
    .unwrap();
    buf.set_position(0);

    let doc = office_oxide::docx::DocxDocument::from_reader(buf).unwrap();
    let text = doc.plain_text();
    assert!(text.contains("Big red"), "text: {text}");
    assert!(text.contains("small blue"), "text: {text}");
}

#[test]
fn ir_code_block_round_trip() {
    use office_oxide::ir::*;

    let ir = DocumentIR {
        metadata: Metadata {
            format: office_oxide::DocumentFormat::Docx,
            ..Default::default()
        },
        sections: vec![Section {
            elements: vec![Element::CodeBlock(CodeBlock {
                language: Some("rust".to_string()),
                content: "fn main() {\n    println!(\"hello\");\n}".to_string(),
            })],
            ..Default::default()
        }],
    };

    let mut buf = Cursor::new(Vec::new());
    office_oxide::create::create_from_ir_to_writer(
        &ir,
        office_oxide::DocumentFormat::Docx,
        &mut buf,
    )
    .unwrap();
    buf.set_position(0);

    let doc = office_oxide::docx::DocxDocument::from_reader(buf).unwrap();
    let text = doc.plain_text();
    assert!(text.contains("fn main"), "text: {text}");
    assert!(text.contains("println"), "text: {text}");
}

#[test]
fn ir_table_cell_padding_round_trip() {
    use office_oxide::ir::*;

    let ir = DocumentIR {
        metadata: Metadata {
            format: office_oxide::DocumentFormat::Docx,
            ..Default::default()
        },
        sections: vec![Section {
            elements: vec![Element::Table(Table {
                rows: vec![TableRow {
                    cells: vec![
                        TableCell {
                            content: vec![Element::Paragraph(Paragraph {
                                content: vec![InlineContent::Text(TextSpan::plain("Padded"))],
                                ..Default::default()
                            })],
                            col_span: 1,
                            row_span: 1,
                            padding: Some(CellPadding {
                                top_twips: Some(144),
                                bottom_twips: Some(144),
                                left_twips: Some(288),
                                right_twips: Some(288),
                            }),
                            ..Default::default()
                        },
                        TableCell {
                            content: vec![Element::Paragraph(Paragraph {
                                content: vec![InlineContent::Text(TextSpan::plain("Normal"))],
                                ..Default::default()
                            })],
                            col_span: 1,
                            row_span: 1,
                            ..Default::default()
                        },
                    ],
                    is_header: false,
                    ..Default::default()
                }],
                ..Default::default()
            })],
            ..Default::default()
        }],
    };

    let mut buf = Cursor::new(Vec::new());
    office_oxide::create::create_from_ir_to_writer(
        &ir,
        office_oxide::DocumentFormat::Docx,
        &mut buf,
    )
    .unwrap();
    buf.set_position(0);

    // Verify the DOCX is valid and content is preserved
    let doc = office_oxide::docx::DocxDocument::from_reader(buf.clone()).unwrap();
    let text = doc.plain_text();
    assert!(text.contains("Padded"), "text: {text}");
    assert!(text.contains("Normal"), "text: {text}");

    // Verify <w:tcMar> appears in the raw XML
    buf.set_position(0);
    let zip_bytes = buf.into_inner();
    let mut zip = zip::ZipArchive::new(Cursor::new(zip_bytes)).unwrap();
    let mut doc_xml = String::new();
    {
        use std::io::Read;
        zip.by_name("word/document.xml")
            .unwrap()
            .read_to_string(&mut doc_xml)
            .unwrap();
    }
    assert!(doc_xml.contains("w:tcMar"), "expected w:tcMar in document.xml");
    assert!(doc_xml.contains(r#"w:w="288""#), "expected left/right padding value");
}

#[test]
fn ir_table_cell_text_align_round_trip() {
    use office_oxide::ir::*;

    let ir = DocumentIR {
        metadata: Metadata {
            format: office_oxide::DocumentFormat::Docx,
            ..Default::default()
        },
        sections: vec![Section {
            elements: vec![Element::Table(Table {
                rows: vec![TableRow {
                    cells: vec![
                        TableCell {
                            content: vec![Element::Paragraph(Paragraph {
                                content: vec![InlineContent::Text(TextSpan::plain("Centered"))],
                                ..Default::default()
                            })],
                            col_span: 1,
                            row_span: 1,
                            text_align: Some(ParagraphAlignment::Center),
                            ..Default::default()
                        },
                        TableCell {
                            content: vec![Element::Paragraph(Paragraph {
                                content: vec![InlineContent::Text(TextSpan::plain("Right"))],
                                // Explicit alignment on the paragraph takes priority
                                alignment: Some(ParagraphAlignment::Right),
                                ..Default::default()
                            })],
                            col_span: 1,
                            row_span: 1,
                            text_align: Some(ParagraphAlignment::Center),
                            ..Default::default()
                        },
                    ],
                    is_header: false,
                    ..Default::default()
                }],
                ..Default::default()
            })],
            ..Default::default()
        }],
    };

    let mut buf = Cursor::new(Vec::new());
    office_oxide::create::create_from_ir_to_writer(
        &ir,
        office_oxide::DocumentFormat::Docx,
        &mut buf,
    )
    .unwrap();
    buf.set_position(0);

    let doc = office_oxide::docx::DocxDocument::from_reader(buf.clone()).unwrap();
    let text = doc.plain_text();
    assert!(text.contains("Centered"), "text: {text}");
    assert!(text.contains("Right"), "text: {text}");

    // Verify <w:jc w:val="center"> appears in the raw XML (from cell-level alignment)
    buf.set_position(0);
    let zip_bytes = buf.into_inner();
    let mut zip = zip::ZipArchive::new(Cursor::new(zip_bytes)).unwrap();
    let mut doc_xml = String::new();
    {
        use std::io::Read;
        zip.by_name("word/document.xml")
            .unwrap()
            .read_to_string(&mut doc_xml)
            .unwrap();
    }
    assert!(
        doc_xml.contains(r#"w:val="center""#),
        "expected center alignment in document.xml"
    );
    // The right-aligned paragraph must not be overwritten to center
    assert!(
        doc_xml.contains(r#"w:val="right""#),
        "expected right alignment preserved in document.xml"
    );
}

#[test]
fn ir_table_caption_round_trip() {
    use office_oxide::ir::*;

    let ir = DocumentIR {
        metadata: Metadata {
            format: office_oxide::DocumentFormat::Docx,
            ..Default::default()
        },
        sections: vec![Section {
            elements: vec![Element::Table(Table {
                rows: vec![TableRow {
                    cells: vec![TableCell {
                        content: vec![Element::Paragraph(Paragraph {
                            content: vec![InlineContent::Text(TextSpan::plain("Data"))],
                            ..Default::default()
                        })],
                        col_span: 1,
                        row_span: 1,
                        ..Default::default()
                    }],
                    is_header: false,
                    ..Default::default()
                }],
                caption: Some("Table 1: Summary of results".to_string()),
                ..Default::default()
            })],
            ..Default::default()
        }],
    };

    let mut buf = Cursor::new(Vec::new());
    office_oxide::create::create_from_ir_to_writer(
        &ir,
        office_oxide::DocumentFormat::Docx,
        &mut buf,
    )
    .unwrap();
    buf.set_position(0);

    let doc = office_oxide::docx::DocxDocument::from_reader(buf.clone()).unwrap();
    let text = doc.plain_text();
    assert!(text.contains("Table 1: Summary of results"), "text: {text}");
    assert!(text.contains("Data"), "text: {text}");

    // Verify the Caption style appears in the raw XML
    buf.set_position(0);
    let zip_bytes = buf.into_inner();
    let mut zip = zip::ZipArchive::new(Cursor::new(zip_bytes)).unwrap();
    let mut doc_xml = String::new();
    {
        use std::io::Read;
        zip.by_name("word/document.xml")
            .unwrap()
            .read_to_string(&mut doc_xml)
            .unwrap();
    }
    assert!(doc_xml.contains("Caption"), "expected Caption style in document.xml");
    assert!(
        doc_xml.contains("Table 1: Summary of results"),
        "expected caption text in document.xml"
    );
}
