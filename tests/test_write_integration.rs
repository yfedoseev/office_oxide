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

/// Regression: `set_cell` used to rebuild the `<c>` element from scratch, dropping
/// every attribute of the cell it wrote — including `s`, the style index. A cell that
/// carried a number format came back naked.
///
/// Reproduce on `main` by dropping the `s="5"` assertion's negation: the written cell
/// is emitted as `<c r="A1"><v>99</v></c>`.
#[test]
fn xlsx_edit_set_cell_preserves_cell_style() {
    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData><row r="1"><c r="A1" s="5" t="n"><v>42</v></c></row></sheetData></worksheet>"#;

    let bytes = workbook_with_sheet(sheet_xml);
    let mut editable = office_oxide::xlsx::edit::EditableXlsx::from_reader(Cursor::new(bytes))
        .expect("open workbook");
    editable
        .set_cell(0, "A1", office_oxide::xlsx::edit::CellValue::Number(99.0))
        .expect("set cell");

    let mut out = Cursor::new(Vec::new());
    editable.write_to(&mut out).expect("write");
    let sheet = read_sheet_xml(out.into_inner());

    assert!(sheet.contains("<v>99</v>"), "sheet: {sheet}");
    assert!(
        sheet.contains(r#"s="5""#),
        "the cell's style index must survive the write; sheet: {sheet}"
    );
}

/// Regression: an empty but *styled* cell is self-closing — `<c r="A1" s="5"/>`, with
/// no `</c>`. `set_cell` searched for `</c>` first, found the **next** cell's closing
/// tag, and the replacement swallowed the cell in between.
///
/// A pre-formatted spreadsheet template is made entirely of such cells, so the nominal
/// "inject data into a template" case silently destroyed data.
///
/// Reproduce on `main`: `B1` is absent from the output.
#[test]
fn xlsx_edit_set_self_closing_cell_does_not_delete_the_next_cell() {
    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData><row r="1"><c r="A1" s="5"/><c r="B1" s="6"><v>7</v></c></row></sheetData></worksheet>"#;

    let bytes = workbook_with_sheet(sheet_xml);
    let mut editable = office_oxide::xlsx::edit::EditableXlsx::from_reader(Cursor::new(bytes))
        .expect("open workbook");
    editable
        .set_cell(0, "A1", office_oxide::xlsx::edit::CellValue::Number(1.0))
        .expect("set cell");

    let mut out = Cursor::new(Vec::new());
    editable.write_to(&mut out).expect("write");
    let sheet = read_sheet_xml(out.into_inner());

    assert!(sheet.contains(r#"<c r="A1" s="5"><v>1</v></c>"#), "sheet: {sheet}");
    assert!(
        sheet.contains(r#"<c r="B1" s="6"><v>7</v></c>"#),
        "the neighbouring cell must survive; sheet: {sheet}"
    );
}

/// Minimal OPC package carrying a single worksheet, so the tests above can control the
/// exact `<c>` shapes they exercise.
fn workbook_with_sheet(sheet_xml: &str) -> Vec<u8> {
    use std::io::Write as _;
    use zip::write::SimpleFileOptions;

    let mut buf = Cursor::new(Vec::new());
    {
        let mut zip = zip::ZipWriter::new(&mut buf);
        let opts = SimpleFileOptions::default();

        let parts = [
            (
                "[Content_Types].xml",
                r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"#,
            ),
            (
                "_rels/.rels",
                r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#,
            ),
            (
                "xl/workbook.xml",
                r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#,
            ),
            ("xl/worksheets/sheet1.xml", sheet_xml),
        ];

        for (name, body) in parts {
            zip.start_file(name, opts).expect("start part");
            zip.write_all(body.as_bytes()).expect("write part");
        }
        zip.finish().expect("finish zip");
    }
    buf.into_inner()
}

fn read_sheet_xml(bytes: Vec<u8>) -> String {
    use std::io::Read as _;

    let mut zip = zip::ZipArchive::new(Cursor::new(bytes)).expect("open zip");
    let mut sheet = String::new();
    zip.by_name("xl/worksheets/sheet1.xml")
        .expect("sheet part")
        .read_to_string(&mut sheet)
        .expect("read sheet");
    sheet
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

// ---------------------------------------------------------------------------
// Tests for Element variants that were previously uncovered: ThematicBreak,
// PageBreak, ColumnBreak, TextBox, Footnote, Endnote. Each test goes
// through the full IR → DOCX → re-parse round-trip so the create.rs
// dispatch arms, the corresponding `DocxWriter` methods, and the
// downstream re-parser all get exercised.
// ---------------------------------------------------------------------------

#[test]
fn ir_thematic_break_emits_bordered_paragraph() {
    use office_oxide::ir::*;

    let ir = DocumentIR {
        metadata: Metadata {
            format: office_oxide::DocumentFormat::Docx,
            ..Default::default()
        },
        sections: vec![Section {
            elements: vec![
                Element::Paragraph(Paragraph {
                    content: vec![InlineContent::Text(TextSpan::plain("Before"))],
                    ..Default::default()
                }),
                Element::ThematicBreak,
                Element::Paragraph(Paragraph {
                    content: vec![InlineContent::Text(TextSpan::plain("After"))],
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
    // Thematic break = empty paragraph with a bottom border. The raw
    // XML should include a w:pBdr/w:bottom element somewhere between
    // "Before" and "After".
    assert!(
        doc_xml.contains("w:pBdr"),
        "expected pBdr (paragraph border) for thematic break"
    );
    assert!(doc_xml.contains("Before") && doc_xml.contains("After"));
}

#[test]
fn ir_page_and_column_breaks_round_trip() {
    use office_oxide::ir::*;

    let ir = DocumentIR {
        metadata: Metadata {
            format: office_oxide::DocumentFormat::Docx,
            ..Default::default()
        },
        sections: vec![Section {
            elements: vec![
                Element::Paragraph(Paragraph {
                    content: vec![InlineContent::Text(TextSpan::plain("Page 1"))],
                    ..Default::default()
                }),
                Element::PageBreak,
                Element::Paragraph(Paragraph {
                    content: vec![InlineContent::Text(TextSpan::plain("Page 2 col 1"))],
                    ..Default::default()
                }),
                Element::ColumnBreak,
                Element::Paragraph(Paragraph {
                    content: vec![InlineContent::Text(TextSpan::plain("Page 2 col 2"))],
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
    // Page break = <w:br w:type="page"/>; column break = <w:br w:type="column"/>.
    assert!(doc_xml.contains("w:type=\"page\""), "expected page break w:br: {doc_xml:.500}",);
    assert!(doc_xml.contains("w:type=\"column\""), "expected column break w:br",);
}

#[test]
fn ir_footnote_endnote_round_trip() {
    use office_oxide::ir::*;

    let footnote_content = vec![Element::Paragraph(Paragraph {
        content: vec![InlineContent::Text(TextSpan::plain("This is a footnote."))],
        ..Default::default()
    })];
    let endnote_content = vec![Element::Paragraph(Paragraph {
        content: vec![InlineContent::Text(TextSpan::plain("This is an endnote."))],
        ..Default::default()
    })];

    let ir = DocumentIR {
        metadata: Metadata {
            format: office_oxide::DocumentFormat::Docx,
            ..Default::default()
        },
        sections: vec![Section {
            elements: vec![
                Element::Paragraph(Paragraph {
                    content: vec![
                        InlineContent::Text(TextSpan::plain("Main body")),
                        InlineContent::FootnoteRef(FootnoteRef {
                            note_id: 1,
                            marker: None,
                        }),
                        InlineContent::EndnoteRef(FootnoteRef {
                            note_id: 2,
                            marker: None,
                        }),
                    ],
                    ..Default::default()
                }),
                Element::Footnote(Note {
                    id: 1,
                    content: footnote_content,
                    marker: None,
                }),
                Element::Endnote(Note {
                    id: 2,
                    content: endnote_content,
                    marker: None,
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

    // The package should now contain a footnotes part and an endnotes part.
    let zip_bytes = buf.into_inner();
    let zip = zip::ZipArchive::new(Cursor::new(zip_bytes)).unwrap();
    let names: Vec<String> = zip.file_names().map(String::from).collect();
    assert!(
        names.iter().any(|n| n == "word/footnotes.xml"),
        "expected word/footnotes.xml in: {names:?}"
    );
    assert!(
        names.iter().any(|n| n == "word/endnotes.xml"),
        "expected word/endnotes.xml in: {names:?}"
    );
}

#[test]
fn ir_text_box_round_trip() {
    use office_oxide::ir::*;

    let ir = DocumentIR {
        metadata: Metadata {
            format: office_oxide::DocumentFormat::Docx,
            ..Default::default()
        },
        sections: vec![Section {
            elements: vec![Element::TextBox(TextBox {
                content: vec![Element::Paragraph(Paragraph {
                    content: vec![InlineContent::Text(TextSpan::plain("Floating callout"))],
                    ..Default::default()
                })],
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
    // Text-box content lives inside a w:txbxContent element rather
    // than as a top-level paragraph, so look for the raw text in
    // the XML. Round-trip plain_text extraction of floating shapes
    // is intentionally not exposed today.
    assert!(
        doc_xml.contains("Floating callout"),
        "expected text-box content in document.xml"
    );
}

#[test]
fn ir_numbered_list_round_trip() {
    use office_oxide::ir::*;

    let item = |text: &str| ListItem {
        content: vec![Element::Paragraph(Paragraph {
            content: vec![InlineContent::Text(TextSpan::plain(text))],
            ..Default::default()
        })],
        ..Default::default()
    };
    let ir = DocumentIR {
        metadata: Metadata {
            format: office_oxide::DocumentFormat::Docx,
            ..Default::default()
        },
        sections: vec![Section {
            elements: vec![Element::List(List {
                ordered: true,
                start_number: Some(5),
                items: vec![item("Five"), item("Six"), item("Seven")],
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
    let md = doc.to_markdown();
    assert!(md.contains("Five") && md.contains("Six") && md.contains("Seven"), "md: {md}");
}

#[test]
fn ir_multi_section_round_trip() {
    use office_oxide::ir::*;

    let make_section = |label: &str, break_type: SectionBreakType| Section {
        elements: vec![Element::Paragraph(Paragraph {
            content: vec![InlineContent::Text(TextSpan::plain(label))],
            ..Default::default()
        })],
        break_type,
        ..Default::default()
    };

    let ir = DocumentIR {
        metadata: Metadata {
            format: office_oxide::DocumentFormat::Docx,
            ..Default::default()
        },
        sections: vec![
            make_section("Section A", SectionBreakType::Continuous),
            make_section("Section B", SectionBreakType::NextPage),
            make_section("Section C", SectionBreakType::OddPage),
        ],
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
    assert!(
        text.contains("Section A") && text.contains("Section B") && text.contains("Section C"),
        "text: {text}"
    );
}

#[test]
fn convenience_functions_round_trip() {
    use office_oxide::ir::*;

    let ir = DocumentIR {
        metadata: Metadata {
            format: office_oxide::DocumentFormat::Docx,
            ..Default::default()
        },
        sections: vec![Section {
            elements: vec![
                Element::Heading(Heading {
                    level: 1,
                    content: vec![InlineContent::Text(TextSpan::plain("Title"))],
                    ..Default::default()
                }),
                Element::Paragraph(Paragraph {
                    content: vec![InlineContent::Text(TextSpan::plain("Hello"))],
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

    // Exercise the crate-level extract_text / to_markdown / Document::open paths.
    let bytes = buf.into_inner();
    let doc = office_oxide::Document::from_reader(
        Cursor::new(bytes.clone()),
        office_oxide::DocumentFormat::Docx,
    )
    .unwrap();
    assert!(doc.plain_text().contains("Hello"));
    assert!(doc.to_markdown().contains("Hello"));
    let ir2 = doc.to_ir();
    assert!(!ir2.sections.is_empty());
}

// ---------------------------------------------------------------------------
// Nested content must reach the file
// ---------------------------------------------------------------------------

fn nested_list_ir() -> office_oxide::ir::DocumentIR {
    use office_oxide::ir::*;
    fn item(text: &str, nested: Option<List>) -> ListItem {
        ListItem {
            content: inline_to_element_block(vec![InlineContent::Text(TextSpan {
                text: text.into(),
                ..Default::default()
            })]),
            nested,
        }
    }
    let deep = List {
        ordered: false,
        items: vec![item("DeepItem", None)],
        ..Default::default()
    };
    let inner = List {
        ordered: true,
        items: vec![item("NestedItemA", Some(deep))],
        ..Default::default()
    };
    let outer = List {
        ordered: true,
        items: vec![item("ItemOne", Some(inner)), item("ItemTwo", None)],
        ..Default::default()
    };
    DocumentIR {
        sections: vec![Section {
            elements: vec![Element::List(outer)],
            ..Default::default()
        }],
        ..Default::default()
    }
}

/// `ListItem::nested` was consumed by the renderers but by no writer, so every
/// item below level 0 vanished on write while the API reported success.
#[test]
fn nested_list_items_reach_every_format() {
    use office_oxide::format::DocumentFormat;

    let ir = nested_list_ir();
    for (fmt, label) in [
        (DocumentFormat::Docx, "docx"),
        (DocumentFormat::Pptx, "pptx"),
    ] {
        let mut buf = std::io::Cursor::new(Vec::new());
        office_oxide::create::create_from_ir_to_writer(&ir, fmt, &mut buf).unwrap();
        buf.set_position(0);
        let mut zip = zip::ZipArchive::new(buf).unwrap();
        let mut all = String::new();
        for i in 0..zip.len() {
            let mut e = zip.by_index(i).unwrap();
            let mut s = String::new();
            if std::io::Read::read_to_string(&mut e, &mut s).is_ok() {
                all.push_str(&s);
            }
        }
        for expected in ["ItemOne", "ItemTwo", "NestedItemA", "DeepItem"] {
            assert!(all.contains(expected), "{label}: {expected} never reached the package");
        }
    }
}

/// The XLSX bridge re-parsed the *rendered* cell string instead of using the
/// type the reader recorded, so "007" became 7, a currency cell became text,
/// and a cell reading "inf" became an Excel error cell.
#[test]
fn xlsx_cells_keep_the_type_and_format_the_reader_recorded() {
    use office_oxide::format::DocumentFormat;
    use office_oxide::ir::*;

    fn cell(
        text: &str,
        dt: Option<CellDataType>,
        raw: Option<f64>,
        fmt: Option<&str>,
    ) -> TableCell {
        TableCell {
            content: vec![Element::Paragraph(Paragraph {
                content: vec![InlineContent::Text(TextSpan {
                    text: text.into(),
                    ..Default::default()
                })],
                ..Default::default()
            })],
            data_type: dt,
            raw_number: raw,
            number_format: fmt.map(str::to_string),
            ..Default::default()
        }
    }

    let ir = DocumentIR {
        sections: vec![Section {
            elements: vec![Element::Table(Table {
                rows: vec![TableRow {
                    cells: vec![
                        cell("007", Some(CellDataType::Text), None, None),
                        cell("inf", Some(CellDataType::Text), None, None),
                        cell(
                            "$1,234.50",
                            Some(CellDataType::Number),
                            Some(1234.5),
                            Some("$#,##0.00"),
                        ),
                    ],
                    ..Default::default()
                }],
                ..Default::default()
            })],
            ..Default::default()
        }],
        ..Default::default()
    };

    let mut buf = std::io::Cursor::new(Vec::new());
    office_oxide::create::create_from_ir_to_writer(&ir, DocumentFormat::Xlsx, &mut buf).unwrap();
    buf.set_position(0);
    let mut zip = zip::ZipArchive::new(buf).unwrap();

    let mut sheet = String::new();
    {
        let mut e = zip.by_name("xl/worksheets/sheet1.xml").unwrap();
        std::io::Read::read_to_string(&mut e, &mut sheet).unwrap();
    }
    assert!(sheet.contains("007"), "leading zeros destroyed: {sheet}");
    assert!(!sheet.contains("#NUM!"), "a text cell became an error cell: {sheet}");
    assert!(sheet.contains("1234.5"), "the number must be written as a number: {sheet}");

    let mut styles = String::new();
    {
        let mut e = zip.by_name("xl/styles.xml").unwrap();
        std::io::Read::read_to_string(&mut e, &mut styles).unwrap();
    }
    // "$#,##0.00" is built-in id 7, so it is referenced rather than redeclared.
    // What matters is that the cell resolves to that format, not to General.
    assert!(
        styles.contains(r#"numFmtId="7""#),
        "the cell's number format was dropped: {styles}"
    );
}

/// `ir_to_xlsx` ended in `_ => {}`, so lists, code blocks, notes and the
/// contents of a text box were discarded. PPTX wraps slide bodies in a text
/// box, which made PPTX → XLSX near-total text loss.
#[test]
fn xlsx_conversion_keeps_list_and_text_box_content() {
    use office_oxide::format::DocumentFormat;
    use office_oxide::ir::*;

    let para = |t: &str| {
        Element::Paragraph(Paragraph {
            content: vec![InlineContent::Text(TextSpan {
                text: t.into(),
                ..Default::default()
            })],
            ..Default::default()
        })
    };
    let ir = DocumentIR {
        sections: vec![Section {
            elements: vec![
                Element::List(List {
                    items: vec![ListItem {
                        content: vec![para("ListWord")],
                        nested: None,
                    }],
                    ..Default::default()
                }),
                Element::CodeBlock(CodeBlock {
                    content: "CodeWord".into(),
                    ..Default::default()
                }),
                Element::TextBox(TextBox {
                    content: vec![para("BoxedWord")],
                    ..Default::default()
                }),
            ],
            ..Default::default()
        }],
        ..Default::default()
    };

    let mut buf = std::io::Cursor::new(Vec::new());
    office_oxide::create::create_from_ir_to_writer(&ir, DocumentFormat::Xlsx, &mut buf).unwrap();
    buf.set_position(0);
    let mut zip = zip::ZipArchive::new(buf).unwrap();
    let mut sheet = String::new();
    let mut e = zip.by_name("xl/worksheets/sheet1.xml").unwrap();
    std::io::Read::read_to_string(&mut e, &mut sheet).unwrap();

    for word in ["ListWord", "CodeWord", "BoxedWord"] {
        assert!(sheet.contains(word), "{word} was dropped: {sheet}");
    }
}

/// `TextSpan::hyperlink` was read and explicitly discarded, so the URL was not
/// recoverable from the output at all.
#[test]
fn hyperlinks_survive_the_write_path() {
    use office_oxide::format::DocumentFormat;

    let md = "See [the docs](https://example.com/a?b=1&c=2) for details.\n";
    let mut buf = std::io::Cursor::new(Vec::new());
    office_oxide::create::create_from_markdown_to_writer(md, DocumentFormat::Docx, &mut buf)
        .unwrap();
    buf.set_position(0);
    let mut zip = zip::ZipArchive::new(buf).unwrap();

    let mut body = String::new();
    {
        let mut e = zip.by_name("word/document.xml").unwrap();
        std::io::Read::read_to_string(&mut e, &mut body).unwrap();
    }
    assert!(body.contains("<w:hyperlink"), "no w:hyperlink emitted: {body}");

    let mut rels = String::new();
    {
        let mut e = zip.by_name("word/_rels/document.xml.rels").unwrap();
        std::io::Read::read_to_string(&mut e, &mut rels).unwrap();
    }
    assert!(
        rels.contains("https://example.com/a?b=1&amp;c=2"),
        "the URL never reached a relationship: {rels}"
    );
    assert!(
        rels.contains(r#"TargetMode="External""#),
        "a hyperlink relationship must be external: {rels}"
    );
}

/// DrawingML has no in-text newline: a dropped break joins the words on
/// either side of it.
#[test]
fn pptx_line_breaks_are_emitted_as_br_elements() {
    use office_oxide::format::DocumentFormat;
    use office_oxide::ir::*;

    let ir = DocumentIR {
        sections: vec![Section {
            elements: vec![Element::Paragraph(Paragraph {
                content: vec![
                    InlineContent::Text(TextSpan {
                        text: "LINEA".into(),
                        ..Default::default()
                    }),
                    InlineContent::LineBreak,
                    InlineContent::Text(TextSpan {
                        text: "LINEB".into(),
                        underline: Some(UnderlineStyle::Single),
                        ..Default::default()
                    }),
                ],
                ..Default::default()
            })],
            ..Default::default()
        }],
        ..Default::default()
    };

    let mut buf = std::io::Cursor::new(Vec::new());
    office_oxide::create::create_from_ir_to_writer(&ir, DocumentFormat::Pptx, &mut buf).unwrap();
    buf.set_position(0);
    let mut zip = zip::ZipArchive::new(buf).unwrap();
    let mut slide = String::new();
    let mut e = zip.by_name("ppt/slides/slide1.xml").unwrap();
    std::io::Read::read_to_string(&mut e, &mut slide).unwrap();

    assert!(slide.contains("<a:br/>"), "no <a:br/> emitted: {slide}");
    assert!(slide.contains(r#"u="sng""#), "underline was dropped: {slide}");
}

/// Tab stops were accepted by the API and emitted nowhere, so a dot-leader
/// table of contents lost both its leaders and its alignment.
#[test]
fn paragraph_tab_stops_are_emitted() {
    use office_oxide::format::DocumentFormat;
    use office_oxide::ir::*;

    let ir = DocumentIR {
        sections: vec![Section {
            elements: vec![Element::Paragraph(Paragraph {
                content: vec![InlineContent::Text(TextSpan {
                    text: "Chapter 1".into(),
                    ..Default::default()
                })],
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

    let mut buf = std::io::Cursor::new(Vec::new());
    office_oxide::create::create_from_ir_to_writer(&ir, DocumentFormat::Docx, &mut buf).unwrap();
    buf.set_position(0);
    let mut zip = zip::ZipArchive::new(buf).unwrap();
    let mut body = String::new();
    let mut e = zip.by_name("word/document.xml").unwrap();
    std::io::Read::read_to_string(&mut e, &mut body).unwrap();

    assert!(body.contains("<w:tabs>"), "no w:tabs emitted: {body}");
    assert!(body.contains(r#"w:pos="8640""#), "tab position lost: {body}");
    assert!(body.contains(r#"w:leader="dot""#), "dot leader lost: {body}");
}

/// The markdown front end produced neither `Element::CodeBlock` nor nested
/// lists, so the fence language leaked into the text and every bullet
/// flattened to level 0.
#[test]
fn markdown_produces_code_blocks_and_nested_lists() {
    use office_oxide::{DocumentIR, format::DocumentFormat, ir::Element};

    let md = "```rust\nlet x = 1;\n```\n\n- one\n  - one-a\n    - one-a-i\n- two\n";
    let ir = DocumentIR::from_markdown(md, DocumentFormat::Docx);
    let elements = &ir.sections[0].elements;

    let code = elements
        .iter()
        .find_map(|e| match e {
            Element::CodeBlock(c) => Some(c),
            _ => None,
        })
        .expect("a fenced block must become Element::CodeBlock");
    assert_eq!(code.language.as_deref(), Some("rust"), "fence language lost");
    assert_eq!(code.content, "let x = 1;", "fence body wrong: {:?}", code.content);

    let list = elements
        .iter()
        .find_map(|e| match e {
            Element::List(l) => Some(l),
            _ => None,
        })
        .expect("a list");
    assert_eq!(list.items.len(), 2, "top level should hold two items");
    let level1 = list.items[0]
        .nested
        .as_ref()
        .expect("one-a must nest under one");
    assert_eq!(level1.items.len(), 1);
    let level2 = level1.items[0]
        .nested
        .as_ref()
        .expect("one-a-i must nest under one-a");
    assert_eq!(level2.items.len(), 1);
}

/// A table flattened into tab-joined text loses the grid entirely, and every
/// bullet emitted at level 0 loses the nesting the IR carries.
#[test]
fn pptx_writes_real_tables_and_nested_bullets() {
    use office_oxide::format::DocumentFormat;

    let md = "# Deck\n\n| A | B |\n|---|---|\n| 1 | 2 |\n\n- one\n  - one-a\n    - one-a-i\n";
    let mut buf = std::io::Cursor::new(Vec::new());
    office_oxide::create::create_from_markdown_to_writer(md, DocumentFormat::Pptx, &mut buf)
        .unwrap();
    buf.set_position(0);
    let mut zip = zip::ZipArchive::new(buf).unwrap();
    let mut slide = String::new();
    let mut e = zip.by_name("ppt/slides/slide1.xml").unwrap();
    std::io::Read::read_to_string(&mut e, &mut slide).unwrap();

    assert!(slide.contains("<a:tbl>"), "no real table emitted: {slide}");
    assert!(slide.contains("<a:gridCol"), "table has no grid: {slide}");
    assert_eq!(slide.matches("<a:tr ").count(), 2, "expected two table rows");
    assert!(!slide.contains('\t'), "a raw tab means the table was flattened");

    assert!(slide.contains(r#"lvl="1""#), "second-level bullet lost: {slide}");
    assert!(slide.contains(r#"lvl="2""#), "third-level bullet lost: {slide}");
    assert!(slide.contains("marL="), "bullets need a hanging indent: {slide}");
}

/// Frame position, page background and the table caption were all IR fields
/// with no writer behind them.
#[test]
fn frame_position_background_and_table_caption_are_emitted() {
    use office_oxide::format::DocumentFormat;
    use office_oxide::ir::*;

    let ir = DocumentIR {
        sections: vec![Section {
            background_rgb: Some([0x11, 0x22, 0x33]),
            elements: vec![
                Element::Paragraph(Paragraph {
                    content: vec![InlineContent::Text(TextSpan {
                        text: "framed".into(),
                        ..Default::default()
                    })],
                    frame_position: Some(FramePosition {
                        x_twips: 100,
                        y_twips: 200,
                        width_twips: 3000,
                        height_twips: 400,
                    }),
                    ..Default::default()
                }),
                Element::Table(Table {
                    caption: Some("TableCaptionText".into()),
                    rows: vec![TableRow {
                        cells: vec![TableCell::default()],
                        ..Default::default()
                    }],
                    ..Default::default()
                }),
            ],
            ..Default::default()
        }],
        ..Default::default()
    };

    let mut buf = std::io::Cursor::new(Vec::new());
    office_oxide::create::create_from_ir_to_writer(&ir, DocumentFormat::Docx, &mut buf).unwrap();
    buf.set_position(0);
    let mut zip = zip::ZipArchive::new(buf).unwrap();
    let mut body = String::new();
    let mut e = zip.by_name("word/document.xml").unwrap();
    std::io::Read::read_to_string(&mut e, &mut body).unwrap();

    assert!(body.contains("<w:framePr"), "frame position lost: {body}");
    assert!(body.contains(r#"w:color="112233""#), "page background lost: {body}");
    assert!(
        body.contains(r#"<w:tblCaption w:val="TableCaptionText"/>"#),
        "table caption not written as w:tblCaption: {body}"
    );
}

/// A hyperlink inside a footnote emits `r:id` into `footnotes.xml`. Without
/// `xmlns:r` on that part's root it is not well-formed XML at all — found by
/// converting a real document, not by any unit test.
#[test]
fn a_hyperlink_in_a_footnote_keeps_the_part_well_formed() {
    use office_oxide::docx::write::{DocxWriter, Run};
    use office_oxide::ir::*;

    let mut w = DocxWriter::new();
    w.add_paragraph("body");
    w.add_footnote(
        1,
        &[Element::Paragraph(Paragraph {
            content: vec![InlineContent::Text(TextSpan {
                text: "see here".into(),
                hyperlink: Some("https://example.com/x".into()),
                ..Default::default()
            })],
            ..Default::default()
        })],
    );
    let _ = Run::new("");

    let mut buf = std::io::Cursor::new(Vec::new());
    w.write_to(&mut buf).unwrap();
    buf.set_position(0);
    let mut zip = zip::ZipArchive::new(buf).unwrap();
    let mut xml = String::new();
    let mut e = zip.by_name("word/footnotes.xml").unwrap();
    std::io::Read::read_to_string(&mut e, &mut xml).unwrap();

    if xml.contains("r:id") {
        assert!(
            xml.contains("xmlns:r="),
            "footnotes.xml uses the r: prefix without declaring it:\n{xml}"
        );
    }
    // Well-formedness: every prefix used must be declared on the root.
    let mut reader = quick_xml::Reader::from_str(&xml);
    let mut buf2 = Vec::new();
    loop {
        match reader.read_event_into(&mut buf2) {
            Ok(quick_xml::events::Event::Eof) => break,
            Ok(_) => {},
            Err(err) => panic!("footnotes.xml is not well-formed: {err}\n{xml}"),
        }
        buf2.clear();
    }
}

/// A drawing inside a header emits the `wp:`/`a:`/`pic:`/`wps:` prefixes. The
/// header root declared only `w:` and `r:`, so the part was not well-formed
/// XML. Pre-existing in v0.1.10; found by converting real documents.
#[test]
fn a_drawing_in_a_header_keeps_the_part_well_formed() {
    use office_oxide::docx::write::{DocxWriter, HfType};
    use office_oxide::ir::*;

    let mut w = DocxWriter::new();
    w.add_paragraph("body");
    w.add_section_header(
        HfType::DefaultHeader,
        vec![Element::TextBox(TextBox {
            content: vec![Element::Paragraph(Paragraph {
                content: vec![InlineContent::Text(TextSpan {
                    text: "in header".into(),
                    ..Default::default()
                })],
                ..Default::default()
            })],
            ..Default::default()
        })],
    );

    let mut buf = std::io::Cursor::new(Vec::new());
    w.write_to(&mut buf).unwrap();
    buf.set_position(0);
    let mut zip = zip::ZipArchive::new(buf).unwrap();
    let name = (0..zip.len())
        .map(|i| zip.by_index(i).unwrap().name().to_string())
        .find(|n| n.starts_with("word/header"))
        .expect("a header part");
    let mut xml = String::new();
    let mut e = zip.by_name(&name).unwrap();
    std::io::Read::read_to_string(&mut e, &mut xml).unwrap();

    let mut reader = quick_xml::Reader::from_str(&xml);
    let mut b = Vec::new();
    loop {
        match reader.read_event_into(&mut b) {
            Ok(quick_xml::events::Event::Eof) => break,
            Ok(_) => {},
            Err(err) => panic!("{name} is not well-formed: {err}\n{xml}"),
        }
        b.clear();
    }
    if xml.contains("<wp:") {
        assert!(xml.contains("xmlns:wp="), "wp: used without a declaration:\n{xml}");
    }
}

/// `CT_TblPrBase` is a strict sequence: tblW, jc, tblInd, tblBorders, shd,
/// tblCellMar, tblCaption. A table carrying an alignment or a caption was
/// invalid — 183 corpus conversions failed on this alone.
#[test]
fn table_properties_follow_the_schema_sequence() {
    use office_oxide::format::DocumentFormat;
    use office_oxide::ir::*;

    let ir = DocumentIR {
        sections: vec![Section {
            elements: vec![Element::Table(Table {
                caption: Some("Cap".into()),
                alignment: Some(TableAlignment::Center),
                indent_left_twips: Some(100),
                cell_padding_twips: Some(50),
                rows: vec![TableRow {
                    cells: vec![TableCell::default()],
                    ..Default::default()
                }],
                ..Default::default()
            })],
            ..Default::default()
        }],
        ..Default::default()
    };
    let mut buf = std::io::Cursor::new(Vec::new());
    office_oxide::create::create_from_ir_to_writer(&ir, DocumentFormat::Docx, &mut buf).unwrap();
    buf.set_position(0);
    let mut zip = zip::ZipArchive::new(buf).unwrap();
    let mut xml = String::new();
    let mut e = zip.by_name("word/document.xml").unwrap();
    std::io::Read::read_to_string(&mut e, &mut xml).unwrap();

    let pos = |needle: &str| {
        xml.find(needle)
            .unwrap_or_else(|| panic!("missing {needle}: {xml}"))
    };
    let (w, jc, ind, mar, cap) = (
        pos("<w:tblW"),
        pos("<w:jc"),
        pos("<w:tblInd"),
        pos("<w:tblCellMar"),
        pos("<w:tblCaption"),
    );
    assert!(w < jc, "tblW must precede jc");
    assert!(jc < ind, "jc must precede tblInd");
    assert!(ind < mar, "tblInd must precede tblCellMar");
    assert!(mar < cap, "tblCellMar must precede tblCaption");
}

/// A drawing nested inside a table cell is not reached by the top-level
/// content scan, so gating the drawing namespaces on that scan left
/// `document.xml` using undeclared prefixes.
#[test]
fn a_drawing_nested_in_a_table_keeps_document_xml_well_formed() {
    use office_oxide::docx::write::DocxWriter;
    use office_oxide::ir::*;

    let mut w = DocxWriter::new();
    w.add_ir_table(&Table {
        rows: vec![TableRow {
            cells: vec![TableCell {
                content: vec![Element::TextBox(TextBox {
                    content: vec![Element::Paragraph(Paragraph {
                        content: vec![InlineContent::Text(TextSpan {
                            text: "in a cell".into(),
                            ..Default::default()
                        })],
                        ..Default::default()
                    })],
                    ..Default::default()
                })],
                ..Default::default()
            }],
            ..Default::default()
        }],
        ..Default::default()
    });

    let mut buf = std::io::Cursor::new(Vec::new());
    w.write_to(&mut buf).unwrap();
    buf.set_position(0);
    let mut zip = zip::ZipArchive::new(buf).unwrap();
    let mut xml = String::new();
    let mut e = zip.by_name("word/document.xml").unwrap();
    std::io::Read::read_to_string(&mut e, &mut xml).unwrap();

    let mut reader = quick_xml::Reader::from_str(&xml);
    let mut b = Vec::new();
    loop {
        match reader.read_event_into(&mut b) {
            Ok(quick_xml::events::Event::Eof) => break,
            Ok(_) => {},
            Err(err) => panic!("document.xml is not well-formed: {err}\n{xml}"),
        }
        b.clear();
    }
}

/// Every anchor attribute must appear exactly once. A duplicate makes the
/// part not well-formed, which no schema check reaches — the parse fails
/// first.
#[test]
fn a_floating_image_anchor_has_no_duplicate_attributes() {
    use office_oxide::docx::write::DocxWriter;
    use office_oxide::ir::*;

    const PNG: &[u8] = &[0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];
    let mut w = DocxWriter::new();
    w.add_ir_image(&Image {
        data: Some(PNG.to_vec()),
        format: Some(ImageFormat::Png),
        display_width_emu: Some(500_000),
        display_height_emu: Some(500_000),
        positioning: ImagePositioning::Floating(FloatingImage {
            x_emu: 0,
            y_emu: 0,
            width_emu: 500_000,
            height_emu: 500_000,
            h_anchor: FloatAnchor::Page,
            v_anchor: FloatAnchor::Page,
            text_wrap: TextWrap::Square,
            allow_overlap: true,
        }),
        ..Default::default()
    });

    let mut buf = std::io::Cursor::new(Vec::new());
    w.write_to(&mut buf).unwrap();
    buf.set_position(0);
    let mut zip = zip::ZipArchive::new(buf).unwrap();
    let mut xml = String::new();
    let mut e = zip.by_name("word/document.xml").unwrap();
    std::io::Read::read_to_string(&mut e, &mut xml).unwrap();

    if let Some(start) = xml.find("<wp:anchor") {
        let tag = &xml[start..start + xml[start..].find('>').unwrap()];
        for attr in ["behindDoc", "locked", "layoutInCell", "simplePos", "distT"] {
            assert_eq!(
                tag.matches(&format!("{attr}=")).count(),
                1,
                "wp:anchor repeats {attr}: {tag}"
            );
        }
    }

    let mut reader = quick_xml::Reader::from_str(&xml);
    let mut b = Vec::new();
    loop {
        match reader.read_event_into(&mut b) {
            Ok(quick_xml::events::Event::Eof) => break,
            Ok(_) => {},
            Err(err) => panic!("document.xml is not well-formed: {err}"),
        }
        b.clear();
    }
}
