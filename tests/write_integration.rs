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
        },
        sections: vec![Section {
            title: Some("Section 1".to_string()),
            elements: vec![
                Element::Heading(Heading {
                    level: 1,
                    content: vec![InlineContent::Text(TextSpan::plain("Main Heading"))],
                }),
                Element::Paragraph(Paragraph {
                    content: vec![InlineContent::Text(TextSpan::plain("Body text here"))],
                }),
                Element::Table(Table {
                    rows: vec![
                        TableRow {
                            cells: vec![
                                TableCell {
                                    content: vec![Element::Paragraph(Paragraph {
                                        content: vec![InlineContent::Text(TextSpan::plain("H1"))],
                                    })],
                                    col_span: 1,
                                    row_span: 1,
                                },
                                TableCell {
                                    content: vec![Element::Paragraph(Paragraph {
                                        content: vec![InlineContent::Text(TextSpan::plain("H2"))],
                                    })],
                                    col_span: 1,
                                    row_span: 1,
                                },
                            ],
                            is_header: true,
                        },
                        TableRow {
                            cells: vec![
                                TableCell {
                                    content: vec![Element::Paragraph(Paragraph {
                                        content: vec![InlineContent::Text(TextSpan::plain("A"))],
                                    })],
                                    col_span: 1,
                                    row_span: 1,
                                },
                                TableCell {
                                    content: vec![Element::Paragraph(Paragraph {
                                        content: vec![InlineContent::Text(TextSpan::plain("B"))],
                                    })],
                                    col_span: 1,
                                    row_span: 1,
                                },
                            ],
                            is_header: false,
                        },
                    ],
                }),
            ],
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
