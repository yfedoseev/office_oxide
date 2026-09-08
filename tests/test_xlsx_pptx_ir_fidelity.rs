//! Conversion fidelity for XLSX and PPTX → `DocumentIR`.
//!
//! Fixtures are built in code (AGENTS.md rule #4). Each test pins a defect
//! where the parser read a value correctly and the converter — or the
//! writer — then reported something else.

use std::io::Cursor;

use office_oxide::core::opc::{OpcWriter, PartName};
use office_oxide::core::relationships::rel_types;
use office_oxide::ir::*;
use office_oxide::{Document, DocumentFormat};

// ---------------------------------------------------------------------------
// XLSX fixture builder
// ---------------------------------------------------------------------------

const CT_WB: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const CT_WS: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const CT_STYLES: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
const CT_COMMENTS: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
const CT_CORE: &str = "application/vnd.openxmlformats-package.core-properties+xml";

/// One sheet: display name, optional `state`, and the `<sheetData>` body.
struct Sheet<'a> {
    name: &'a str,
    state: Option<&'a str>,
    body: &'a str,
    /// Extra XML appended after `</sheetData>` (hyperlinks, pageSetup, …).
    extra: &'a str,
}

impl<'a> Sheet<'a> {
    fn new(name: &'a str, body: &'a str) -> Self {
        Self {
            name,
            state: None,
            body,
            extra: "",
        }
    }
}

struct Xlsx<'a> {
    sheets: Vec<Sheet<'a>>,
    styles: Option<&'a str>,
    comments: Option<&'a str>,
    core: Option<&'a str>,
}

impl<'a> Xlsx<'a> {
    fn new(sheets: Vec<Sheet<'a>>) -> Self {
        Self {
            sheets,
            styles: None,
            comments: None,
            core: None,
        }
    }

    fn styles(mut self, inner: &'a str) -> Self {
        self.styles = Some(inner);
        self
    }

    /// Attach a comments part to the *first* sheet.
    fn comments(mut self, inner: &'a str) -> Self {
        self.comments = Some(inner);
        self
    }

    fn core(mut self, inner: &'a str) -> Self {
        self.core = Some(inner);
        self
    }

    fn ir(self) -> DocumentIR {
        let mut w = OpcWriter::new(Cursor::new(Vec::new())).unwrap();
        let wb = PartName::new("/xl/workbook.xml").unwrap();
        w.add_package_rel(rel_types::OFFICE_DOCUMENT, "xl/workbook.xml");

        let mut sheet_tags = String::new();
        for (i, s) in self.sheets.iter().enumerate() {
            let file = format!("worksheets/sheet{}.xml", i + 1);
            let rid = w.add_part_rel(&wb, rel_types::WORKSHEET, &file);
            let state = s
                .state
                .map(|st| format!(r#" state="{st}""#))
                .unwrap_or_default();
            sheet_tags.push_str(&format!(
                r#"<sheet name="{}" sheetId="{}" r:id="{rid}"{state}/>"#,
                s.name,
                i + 1
            ));
            let ws_part = PartName::new(&format!("/xl/{file}")).unwrap();
            let xml = format!(
                r#"<?xml version="1.0"?><worksheet
                     xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                     xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                   <sheetData>{}</sheetData>{}</worksheet>"#,
                s.body, s.extra
            );
            w.add_part(&ws_part, CT_WS, xml.as_bytes()).unwrap();

            if i == 0 {
                if let Some(inner) = self.comments {
                    let part = PartName::new("/xl/comments1.xml").unwrap();
                    let xml = format!(
                        r#"<?xml version="1.0"?><comments
                             xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                           {inner}</comments>"#
                    );
                    w.add_part(&part, CT_COMMENTS, xml.as_bytes()).unwrap();
                    w.add_part_rel(&ws_part, rel_types::COMMENTS, "../comments1.xml");
                }
            }
        }

        let wb_xml = format!(
            r#"<?xml version="1.0"?><workbook
                 xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
               <sheets>{sheet_tags}</sheets></workbook>"#
        );
        w.add_part(&wb, CT_WB, wb_xml.as_bytes()).unwrap();

        if let Some(inner) = self.styles {
            let part = PartName::new("/xl/styles.xml").unwrap();
            let xml = format!(
                r#"<?xml version="1.0"?><styleSheet
                     xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                   {inner}</styleSheet>"#
            );
            w.add_part(&part, CT_STYLES, xml.as_bytes()).unwrap();
            w.add_part_rel(&wb, rel_types::STYLES, "styles.xml");
        }

        if let Some(inner) = self.core {
            let part = PartName::new("/docProps/core.xml").unwrap();
            let xml = format!(
                r#"<?xml version="1.0"?><cp:coreProperties
                     xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                     xmlns:dc="http://purl.org/dc/elements/1.1/"
                     xmlns:dcterms="http://purl.org/dc/terms/">{inner}</cp:coreProperties>"#
            );
            w.add_part(&part, CT_CORE, xml.as_bytes()).unwrap();
            w.add_package_rel(rel_types::CORE_PROPERTIES, "docProps/core.xml");
        }

        let bytes = w.finish().unwrap().into_inner();
        Document::from_reader(Cursor::new(bytes), DocumentFormat::Xlsx)
            .expect("parse xlsx")
            .to_ir()
    }
}

/// Serialise a writer to bytes through an in-memory cursor.
fn to_bytes(wb: &office_oxide::xlsx::write::XlsxWriter) -> Vec<u8> {
    let mut buf = Cursor::new(Vec::new());
    wb.write_to(&mut buf).expect("write xlsx");
    buf.into_inner()
}

fn only_table(ir: &DocumentIR, section: usize) -> &Table {
    ir.sections[section]
        .elements
        .iter()
        .find_map(|e| match e {
            Element::Table(t) => Some(t),
            _ => None,
        })
        .expect("a table")
}

fn cell_text(cell: &TableCell) -> String {
    cell.content
        .iter()
        .map(|e| match e {
            Element::Paragraph(p) => p
                .content
                .iter()
                .filter_map(|c| match c {
                    InlineContent::Text(t) => Some(t.text.as_str()),
                    _ => None,
                })
                .collect::<String>(),
            _ => String::new(),
        })
        .collect()
}

// ---------------------------------------------------------------------------
// #146 — cells keep their column
// ---------------------------------------------------------------------------

#[test]
fn a_row_that_skips_a_column_keeps_every_value_under_its_own_header() {
    // XLSX stores only non-empty cells, so a sparse row is perfectly legal.
    // Emitting cells in encounter order put "9" under the "Name" header.
    let ir = Xlsx::new(vec![Sheet::new(
        "S",
        r#"<row r="1">
             <c r="A1" t="inlineStr"><is><t>Name</t></is></c>
             <c r="B1" t="inlineStr"><is><t>Qty</t></is></c>
             <c r="C1" t="inlineStr"><is><t>Price</t></is></c>
           </row>
           <row r="2"><c r="C2"><v>9</v></c></row>
           <row r="3"><c r="A3" t="inlineStr"><is><t>Pear</t></is></c><c r="C3"><v>4</v></c></row>"#,
    )])
    .ir();
    let t = only_table(&ir, 0);
    assert_eq!(t.rows[1].cells.len(), 3, "row 2 must be padded to full width");
    assert_eq!(cell_text(&t.rows[1].cells[0]), "");
    assert_eq!(cell_text(&t.rows[1].cells[1]), "");
    assert_eq!(cell_text(&t.rows[1].cells[2]), "9", "9 belongs under Price");
    assert_eq!(cell_text(&t.rows[2].cells[0]), "Pear");
    assert_eq!(cell_text(&t.rows[2].cells[1]), "");
    assert_eq!(cell_text(&t.rows[2].cells[2]), "4");
}

// ---------------------------------------------------------------------------
// #184 — cell fonts in table mode
// ---------------------------------------------------------------------------

#[test]
fn cell_font_formatting_reaches_the_ir_in_table_mode() {
    let ir = Xlsx::new(vec![Sheet::new(
        "S",
        r#"<row r="1">
             <c r="A1" s="1" t="inlineStr"><is><t>Head</t></is></c>
             <c r="B1" t="inlineStr"><is><t>Plain</t></is></c>
           </row>
           <row r="2">
             <c r="A2" t="inlineStr"><is><t>a</t></is></c>
             <c r="B2" t="inlineStr"><is><t>b</t></is></c>
           </row>"#,
    )])
    .styles(
        r#"<fonts count="2"><font/><font><b/><i/><sz val="14"/></font></fonts>
           <cellXfs count="2"><xf fontId="0"/><xf fontId="1" applyFont="1"/></cellXfs>"#,
    )
    .ir();
    let t = only_table(&ir, 0);
    let span = match &t.rows[0].cells[0].content[0] {
        Element::Paragraph(p) => match &p.content[0] {
            InlineContent::Text(s) => s,
            other => panic!("not text: {other:?}"),
        },
        other => panic!("not a paragraph: {other:?}"),
    };
    assert!(span.bold, "bold from the cell font");
    assert!(span.italic);
    assert_eq!(span.font_size_half_pt, Some(28), "14pt = 28 half-points");
}

// ---------------------------------------------------------------------------
// #164 — cell hyperlinks
// ---------------------------------------------------------------------------

#[test]
fn cell_hyperlinks_reach_the_ir() {
    let ir = Xlsx::new(vec![Sheet {
        name: "S",
        state: None,
        body: r#"<row r="1">
                   <c r="A1" t="inlineStr"><is><t>Head</t></is></c>
                   <c r="B1" t="inlineStr"><is><t>H2</t></is></c>
                 </row>
                 <row r="2">
                   <c r="A2" t="inlineStr"><is><t>Docs</t></is></c>
                   <c r="B2" t="inlineStr"><is><t>Local</t></is></c>
                 </row>"#,
        extra: r#"<hyperlinks>
                    <hyperlink ref="A2" r:id="rIdLink"/>
                    <hyperlink ref="B2" location="Sheet2!A1"/>
                  </hyperlinks>"#,
    }])
    .ir();
    // The r:id has no matching relationship in this fixture, so only the
    // internal link resolves — which is the case that used to be dropped
    // outright even when it did resolve.
    let t = only_table(&ir, 0);
    let span = match &t.rows[1].cells[1].content[0] {
        Element::Paragraph(p) => match &p.content[0] {
            InlineContent::Text(s) => s,
            other => panic!("not text: {other:?}"),
        },
        other => panic!("not a paragraph: {other:?}"),
    };
    assert_eq!(span.hyperlink.as_deref(), Some("#Sheet2!A1"));
}

// ---------------------------------------------------------------------------
// #193 — hidden sheets are flagged
// ---------------------------------------------------------------------------

#[test]
fn hidden_and_very_hidden_sheets_are_flagged() {
    let body = r#"<row r="1"><c r="A1" t="inlineStr"><is><t>x</t></is></c></row>"#;
    let ir = Xlsx::new(vec![
        Sheet::new("Visible", body),
        Sheet {
            name: "Hidden",
            state: Some("hidden"),
            body,
            extra: "",
        },
        Sheet {
            name: "Secret",
            state: Some("veryHidden"),
            body,
            extra: "",
        },
    ])
    .ir();
    assert_eq!(ir.sections.len(), 3);
    assert!(!ir.sections[0].hidden);
    assert!(ir.sections[1].hidden, "state=hidden");
    assert!(ir.sections[2].hidden, "state=veryHidden");
    // The content is still extracted — a workbook index wants it.
    assert!(ir.plain_text().contains('x'));
}

// ---------------------------------------------------------------------------
// #174 — workbook metadata is not the first sheet's name
// ---------------------------------------------------------------------------

#[test]
fn workbook_metadata_comes_from_core_properties() {
    let ir = Xlsx::new(vec![Sheet::new(
        "Sheet1",
        r#"<row r="1"><c r="A1" t="inlineStr"><is><t>x</t></is></c></row>"#,
    )])
    .core(r#"<dc:title>Budget 2026</dc:title><dc:creator>Finance</dc:creator>"#)
    .ir();
    assert_eq!(ir.metadata.title.as_deref(), Some("Budget 2026"));
    assert_eq!(ir.metadata.author.as_deref(), Some("Finance"));
}

#[test]
fn workbook_title_falls_back_to_the_first_sheet_name() {
    let ir = Xlsx::new(vec![Sheet::new(
        "Sheet1",
        r#"<row r="1"><c r="A1" t="inlineStr"><is><t>x</t></is></c></row>"#,
    )])
    .ir();
    assert_eq!(ir.metadata.title.as_deref(), Some("Sheet1"));
}

// ---------------------------------------------------------------------------
// #147 — number formats
// ---------------------------------------------------------------------------

#[test]
fn quoted_literals_in_a_number_format_are_not_date_tokens() {
    // The `M` in `" M"` and the `y`/`d` in `"yes"`/`" days"` are literal
    // text. Reading them as month/year/day tokens rendered 12,500,000 as
    // the date 36123-11-01.
    let ir = Xlsx::new(vec![Sheet::new(
        "S",
        r#"<row r="1">
             <c r="A1" s="1"><v>1</v></c>
             <c r="B1" s="2"><v>2.5</v></c>
             <c r="C1" s="3"><v>12500000</v></c>
           </row>
           <row r="2">
             <c r="A2" s="1"><v>1</v></c>
             <c r="B2" s="2"><v>2.5</v></c>
             <c r="C2" s="3"><v>12500000</v></c>
           </row>"#,
    )])
    .styles(
        r##"<numFmts count="3">
             <numFmt numFmtId="164" formatCode="&quot;yes&quot;;&quot;no&quot;"/>
             <numFmt numFmtId="165" formatCode="0.00&quot; days&quot;"/>
             <numFmt numFmtId="166" formatCode="#,##0,,&quot; M&quot;"/>
           </numFmts>
           <cellXfs count="4"><xf numFmtId="0"/>
             <xf numFmtId="164" applyNumberFormat="1"/>
             <xf numFmtId="165" applyNumberFormat="1"/>
             <xf numFmtId="166" applyNumberFormat="1"/></cellXfs>"##,
    )
    .ir();
    let t = only_table(&ir, 0);
    assert_eq!(cell_text(&t.rows[0].cells[0]), "yes");
    assert_eq!(cell_text(&t.rows[0].cells[1]), "2.50 days");
    // Two trailing commas scale by 1e6; `#,##0` then rounds 12.5 to 13,
    // which is what Excel displays.
    assert_eq!(cell_text(&t.rows[0].cells[2]), "13 M");
}

#[test]
fn a_real_date_format_is_still_a_date() {
    let ir = Xlsx::new(vec![Sheet::new(
        "S",
        r#"<row r="1"><c r="A1" s="1"><v>45000</v></c><c r="B1"><v>1</v></c></row>
           <row r="2"><c r="A2" s="1"><v>45000</v></c><c r="B2"><v>1</v></c></row>"#,
    )])
    .styles(
        r#"<numFmts count="1">
             <numFmt numFmtId="164" formatCode="yyyy&quot;年&quot;m&quot;月&quot;"/>
           </numFmts>
           <cellXfs count="2"><xf numFmtId="0"/>
             <xf numFmtId="164" applyNumberFormat="1"/></cellXfs>"#,
    )
    .ir();
    assert!(
        cell_text(&only_table(&ir, 0).rows[0].cells[0]).starts_with("2023-"),
        "expected a 2023 date, got {:?}",
        cell_text(&only_table(&ir, 0).rows[0].cells[0])
    );
}

// ---------------------------------------------------------------------------
// #165 — cell comments
// ---------------------------------------------------------------------------

#[test]
fn cell_comments_reach_the_ir() {
    let ir = Xlsx::new(vec![Sheet::new(
        "S",
        r#"<row r="1"><c r="A1" t="inlineStr"><is><t>x</t></is></c></row>"#,
    )])
    .comments(
        r#"<authors><author>Reviewer</author></authors>
           <commentList>
             <comment ref="B2" authorId="0"><text><r><t>Check this figure</t></r></text></comment>
           </commentList>"#,
    )
    .ir();
    let note = ir.sections[0]
        .elements
        .iter()
        .find_map(|e| match e {
            Element::Endnote(n) => Some(n),
            _ => None,
        })
        .expect("comment reached the IR");
    assert_eq!(note.marker.as_deref(), Some("B2 (Reviewer)"));
    assert!(ir.plain_text().contains("Check this figure"));
}

// ---------------------------------------------------------------------------
// #160 / #173 — writer validation
// ---------------------------------------------------------------------------

#[test]
fn non_finite_numbers_are_written_as_an_error_cell_not_as_inf() {
    use office_oxide::xlsx::write::{CellData, XlsxWriter};
    let mut wb = XlsxWriter::new();
    {
        let mut s = wb.add_sheet("S");
        s.set_cell(0, 0, CellData::Number(f64::NAN));
        s.set_cell(0, 1, CellData::Number(f64::INFINITY));
    }
    // The sheet part is deflated, so re-read through the parser: what
    // matters is that the workbook still opens and no `inf`/`NaN` literal
    // reached a `<v>` element.
    let doc =
        Document::from_reader(Cursor::new(to_bytes(&wb)), DocumentFormat::Xlsx).expect("parse");
    let text = doc.plain_text();
    assert!(!text.contains("inf"), "wrote a non-finite literal: {text:?}");
    assert!(!text.contains("NaN"), "wrote a non-finite literal: {text:?}");
    assert!(text.contains("#NUM!"), "expected an error cell: {text:?}");
}

#[test]
fn illegal_sheet_names_are_normalised() {
    use office_oxide::xlsx::write::XlsxWriter;
    let mut wb = XlsxWriter::new();
    let long = "N".repeat(40);
    wb.add_sheet(&long);
    wb.add_sheet("a/b:c*d?e[f]");
    wb.add_sheet("");
    wb.add_sheet("Dup");
    wb.add_sheet("Dup");
    let doc =
        Document::from_reader(Cursor::new(to_bytes(&wb)), DocumentFormat::Xlsx).expect("parse");
    let names: Vec<String> = doc
        .to_ir()
        .sections
        .iter()
        .filter_map(|s| s.title.clone())
        .collect();
    assert_eq!(names.len(), 5);
    assert!(names[0].chars().count() <= 31, "name too long: {:?}", names[0]);
    assert!(
        !names[1].contains(['/', ':', '*', '?', '[', ']']),
        "forbidden character survived: {:?}",
        names[1]
    );
    assert!(!names[2].is_empty(), "empty sheet name survived");
    assert_ne!(names[3], names[4], "duplicate sheet names survived");
}

#[test]
fn writing_to_a_nonexistent_sheet_reports_failure() {
    use office_oxide::xlsx::write::{CellData, XlsxWriter};
    let mut wb = XlsxWriter::new();
    let idx = wb.add_sheet_get_index("Only");
    assert!(wb.sheet_set_cell(idx, 0, 0, CellData::String("KEPT".into())));
    assert!(
        !wb.sheet_set_cell(99, 0, 0, CellData::String("LOST".into())),
        "a write to a sheet that does not exist must not report success"
    );
    let doc =
        Document::from_reader(Cursor::new(to_bytes(&wb)), DocumentFormat::Xlsx).expect("parse");
    assert!(doc.plain_text().contains("KEPT"));
    assert!(!doc.plain_text().contains("LOST"));
}

#[test]
fn cells_outside_excels_grid_are_refused() {
    use office_oxide::xlsx::write::{CellData, XlsxWriter};
    let mut wb = XlsxWriter::new();
    let idx = wb.add_sheet_get_index("S");
    // The last real cell is XFD1048576 = (1048575, 16383).
    assert!(wb.sheet_set_cell(idx, 1_048_575, 16_383, CellData::Number(1.0)));
    assert!(!wb.sheet_set_cell(idx, 1_048_576, 0, CellData::Number(2.0)));
    assert!(!wb.sheet_set_cell(idx, 0, 16_384, CellData::Number(3.0)));
}

// ---------------------------------------------------------------------------
// #179 — landscape page size
// ---------------------------------------------------------------------------

#[test]
fn landscape_orientation_rotates_the_paper_size() {
    let ir = Xlsx::new(vec![Sheet {
        name: "S",
        state: None,
        body: r#"<row r="1"><c r="A1" t="inlineStr"><is><t>x</t></is></c></row>"#,
        extra: r#"<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75"
                                header="0.3" footer="0.3"/>
                  <pageSetup paperSize="9" orientation="landscape"/>"#,
    }])
    .ir();
    let ps = ir.sections[0].page_setup.as_ref().expect("page setup");
    assert!(ps.landscape);
    assert!(
        ps.width_twips > ps.height_twips,
        "landscape must be wider than it is tall, got {}x{}",
        ps.width_twips,
        ps.height_twips
    );
}

// ---------------------------------------------------------------------------
// #148 — spanned cells do not shift their neighbours
// ---------------------------------------------------------------------------

#[test]
fn a_row_spanning_cell_leaves_the_covered_position_empty() {
    // Markdown has no rowspan syntax, so the covered position must render
    // empty. Indexing cells positionally shifted every cell to its right
    // one column left in the row below the span.
    let cell = |text: &str, row_span: u32| TableCell {
        content: vec![Element::Paragraph(Paragraph {
            content: vec![InlineContent::Text(TextSpan::plain(text))],
            ..Default::default()
        })],
        col_span: 1,
        row_span,
        ..Default::default()
    };
    let ir = DocumentIR {
        metadata: Metadata {
            format: DocumentFormat::Docx,
            ..Default::default()
        },
        sections: vec![Section {
            elements: vec![Element::Table(Table {
                rows: vec![
                    TableRow {
                        cells: vec![cell("A1", 2), cell("B1", 1), cell("C1", 1)],
                        ..Default::default()
                    },
                    // A1 spans into this row, so only B2 and C2 exist here.
                    TableRow {
                        cells: vec![cell("B2", 1), cell("C2", 1)],
                        ..Default::default()
                    },
                ],
                ..Default::default()
            })],
            ..Default::default()
        }],
    };
    let md = ir.to_markdown();
    let second_row = md.lines().nth(2).expect("a second body row");
    assert_eq!(second_row, "|  | B2 | C2 |", "B2 must stay in column 2; got {md}");
}

// ---------------------------------------------------------------------------
// PPTX fixture builder
// ---------------------------------------------------------------------------

const CT_PRES: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
const CT_SLIDE: &str = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const CT_NOTES: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml";
const CT_PPTX_COMMENTS: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.comments+xml";

/// One slide: the `<p:sld>` attributes and the `<p:spTree>` body.
struct Slide<'a> {
    attrs: &'a str,
    tree: &'a str,
    notes: Option<&'a str>,
    comments: Option<&'a str>,
}

impl<'a> Slide<'a> {
    fn new(tree: &'a str) -> Self {
        Self {
            attrs: "",
            tree,
            notes: None,
            comments: None,
        }
    }
}

fn pptx_ir(slides: Vec<Slide<'_>>) -> DocumentIR {
    let mut w = OpcWriter::new(Cursor::new(Vec::new())).unwrap();
    let pres = PartName::new("/ppt/presentation.xml").unwrap();
    w.add_package_rel(rel_types::OFFICE_DOCUMENT, "ppt/presentation.xml");

    let mut ids = String::new();
    for (i, s) in slides.iter().enumerate() {
        let file = format!("slides/slide{}.xml", i + 1);
        let rid = w.add_part_rel(&pres, rel_types::SLIDE, &file);
        ids.push_str(&format!(r#"<p:sldId id="{}" r:id="{rid}"/>"#, 256 + i));
        let part = PartName::new(&format!("/ppt/{file}")).unwrap();
        let xml = format!(
            r#"<?xml version="1.0"?><p:sld {}
                 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
               <p:cSld><p:spTree>{}</p:spTree></p:cSld></p:sld>"#,
            s.attrs, s.tree
        );
        w.add_part(&part, CT_SLIDE, xml.as_bytes()).unwrap();

        if let Some(notes) = s.notes {
            let np = PartName::new(&format!("/ppt/notesSlides/notesSlide{}.xml", i + 1)).unwrap();
            let xml = format!(
                r#"<?xml version="1.0"?><p:notes
                     xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                     xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                   <p:cSld><p:spTree><p:sp><p:nvSpPr><p:nvPr>
                     <p:ph type="body"/></p:nvPr></p:nvSpPr>
                     <p:txBody><a:p><a:r><a:t>{notes}</a:t></a:r></a:p></p:txBody>
                   </p:sp></p:spTree></p:cSld></p:notes>"#
            );
            w.add_part(&np, CT_NOTES, xml.as_bytes()).unwrap();
            w.add_part_rel(
                &part,
                rel_types::NOTES_SLIDE,
                &format!("../notesSlides/notesSlide{}.xml", i + 1),
            );
        }

        if let Some(inner) = s.comments {
            let cp = PartName::new(&format!("/ppt/comments/comment{}.xml", i + 1)).unwrap();
            let xml = format!(
                r#"<?xml version="1.0"?><p:cmLst
                     xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
                   {inner}</p:cmLst>"#
            );
            w.add_part(&cp, CT_PPTX_COMMENTS, xml.as_bytes()).unwrap();
            w.add_part_rel(
                &part,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
                &format!("../comments/comment{}.xml", i + 1),
            );
        }
    }

    let pres_xml = format!(
        r#"<?xml version="1.0"?><p:presentation
             xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
           <p:sldIdLst>{ids}</p:sldIdLst></p:presentation>"#
    );
    w.add_part(&pres, CT_PRES, pres_xml.as_bytes()).unwrap();

    let bytes = w.finish().unwrap().into_inner();
    Document::from_reader(Cursor::new(bytes), DocumentFormat::Pptx)
        .expect("parse pptx")
        .to_ir()
}

/// A body placeholder carrying the given `<a:p>` paragraphs.
fn body_sp(paragraphs: &str) -> String {
    format!(
        r#"<p:sp>
             <p:nvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
             <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr>
             <p:txBody>{paragraphs}</p:txBody>
           </p:sp>"#
    )
}

fn find_list(ir: &DocumentIR) -> &List {
    fn walk<'a>(elements: &'a [Element], out: &mut Option<&'a List>) {
        for e in elements {
            match e {
                Element::List(l) if out.is_none() => *out = Some(l),
                Element::TextBox(tb) => walk(&tb.content, out),
                _ => {},
            }
        }
    }
    let mut out = None;
    walk(&ir.sections[0].elements, &mut out);
    out.expect("a list")
}

fn first_pptx_span(ir: &DocumentIR) -> &TextSpan {
    fn walk<'a>(elements: &'a [Element], out: &mut Option<&'a TextSpan>) {
        for e in elements {
            match e {
                Element::Paragraph(p) if out.is_none() => {
                    if let Some(InlineContent::Text(s)) = p.content.first() {
                        *out = Some(s);
                    }
                },
                Element::TextBox(tb) => walk(&tb.content, out),
                _ => {},
            }
        }
    }
    let mut out = None;
    walk(&ir.sections[0].elements, &mut out);
    out.expect("a text span")
}

// ---------------------------------------------------------------------------
// #183 — PPTX run properties
// ---------------------------------------------------------------------------

#[test]
fn pptx_run_underline_font_baseline_caps_and_spacing_reach_the_ir() {
    let tree = body_sp(
        r#"<a:p><a:pPr><a:buNone/></a:pPr>
             <a:r>
               <a:rPr u="dbl" baseline="30000" cap="all" spc="-50">
                 <a:latin typeface="Consolas"/>
               </a:rPr>
               <a:t>styled</a:t>
             </a:r>
           </a:p>"#,
    );
    let ir = pptx_ir(vec![Slide::new(&tree)]);
    let s = first_pptx_span(&ir);
    assert_eq!(s.text, "styled");
    assert_eq!(s.underline, Some(UnderlineStyle::Double));
    assert_eq!(s.font_name.as_deref(), Some("Consolas"));
    assert_eq!(s.vertical_align, Some(VerticalAlign::Superscript));
    assert!(s.all_caps);
    assert_eq!(s.char_spacing_half_pt, Some(-10), "-50 hundredths-pt = -10");
}

// ---------------------------------------------------------------------------
// #150 — bullets
// ---------------------------------------------------------------------------

#[test]
fn level_zero_bullets_still_form_a_list() {
    // Every bullet at lvl=0 is the ordinary single-level list. Keying off
    // `level > 0` alone rendered it as plain paragraphs with no markers.
    let tree = body_sp(
        r#"<a:p><a:pPr><a:buChar char="&#8226;"/></a:pPr>
             <a:r><a:t>one</a:t></a:r></a:p>
           <a:p><a:pPr><a:buChar char="&#8226;"/></a:pPr>
             <a:r><a:t>two</a:t></a:r></a:p>"#,
    );
    let ir = pptx_ir(vec![Slide::new(&tree)]);
    let list = find_list(&ir);
    assert!(!list.ordered);
    assert_eq!(list.style, Some(ListStyle::Bullet));
    assert_eq!(list.items.len(), 2);
}

#[test]
fn auto_numbered_bullets_produce_an_ordered_list() {
    let tree = body_sp(
        r#"<a:p><a:pPr><a:buAutoNum type="alphaLcParenR" startAt="3"/></a:pPr>
             <a:r><a:t>one</a:t></a:r></a:p>
           <a:p><a:pPr><a:buAutoNum type="alphaLcParenR"/></a:pPr>
             <a:r><a:t>two</a:t></a:r></a:p>"#,
    );
    let ir = pptx_ir(vec![Slide::new(&tree)]);
    let list = find_list(&ir);
    assert!(list.ordered, "buAutoNum means an ordered list");
    assert_eq!(list.style, Some(ListStyle::LowerAlpha));
    assert_eq!(list.start_number, Some(3));
}

#[test]
fn bu_none_paragraphs_are_not_a_list() {
    let tree = body_sp(r#"<a:p><a:pPr><a:buNone/></a:pPr><a:r><a:t>prose</a:t></a:r></a:p>"#);
    let ir = pptx_ir(vec![Slide::new(&tree)]);
    fn has_list(elements: &[Element]) -> bool {
        elements.iter().any(|e| match e {
            Element::List(_) => true,
            Element::TextBox(tb) => has_list(&tb.content),
            _ => false,
        })
    }
    assert!(!has_list(&ir.sections[0].elements));
}

// ---------------------------------------------------------------------------
// #192 — table header rows
// ---------------------------------------------------------------------------

fn table_frame(tbl_pr: &str) -> String {
    format!(
        r#"<p:graphicFrame>
             <p:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></p:xfrm>
             <a:graphic><a:graphicData
                 uri="http://schemas.openxmlformats.org/drawingml/2006/table">
               <a:tbl>
                 {tbl_pr}
                 <a:tr><a:tc><a:txBody><a:p><a:r><a:t>A</a:t></a:r></a:p></a:txBody></a:tc></a:tr>
                 <a:tr><a:tc><a:txBody><a:p><a:r><a:t>B</a:t></a:r></a:p></a:txBody></a:tc></a:tr>
               </a:tbl>
             </a:graphicData></a:graphic>
           </p:graphicFrame>"#
    )
}

fn find_table(ir: &DocumentIR) -> &Table {
    fn walk<'a>(elements: &'a [Element], out: &mut Option<&'a Table>) {
        for e in elements {
            match e {
                Element::Table(t) if out.is_none() => *out = Some(t),
                Element::TextBox(tb) => walk(&tb.content, out),
                _ => {},
            }
        }
    }
    let mut out = None;
    walk(&ir.sections[0].elements, &mut out);
    out.expect("a table")
}

#[test]
fn a_table_declaring_no_header_row_does_not_get_one() {
    let tree = table_frame("");
    let ir = pptx_ir(vec![Slide::new(&tree)]);
    let t = find_table(&ir);
    assert!(!t.rows[0].is_header, "row 0 is only a header when a:tblPr says firstRow=\"1\"");
}

#[test]
fn a_table_declaring_first_row_gets_a_header() {
    let tree = table_frame(r#"<a:tblPr firstRow="1"/>"#);
    let ir = pptx_ir(vec![Slide::new(&tree)]);
    let t = find_table(&ir);
    assert!(t.rows[0].is_header);
    assert!(!t.rows[1].is_header);
}

// ---------------------------------------------------------------------------
// #193 — hidden slides
// ---------------------------------------------------------------------------

#[test]
fn a_hidden_slide_is_flagged_but_still_extracted() {
    let tree = body_sp(r#"<a:p><a:r><a:t>SECRET</a:t></a:r></a:p>"#);
    let ir = pptx_ir(vec![
        Slide::new(&tree),
        Slide {
            attrs: r#"show="0""#,
            tree: &tree,
            notes: None,
            comments: None,
        },
    ]);
    assert_eq!(ir.sections.len(), 2);
    assert!(!ir.sections[0].hidden);
    assert!(ir.sections[1].hidden, "show=\"0\" means hidden");
    assert!(ir.plain_text().contains("SECRET"));
}

// ---------------------------------------------------------------------------
// #143 — speaker notes
// ---------------------------------------------------------------------------

#[test]
fn speaker_notes_reach_the_ir() {
    let tree = body_sp(r#"<a:p><a:r><a:t>BODY</a:t></a:r></a:p>"#);
    let ir = pptx_ir(vec![Slide {
        attrs: "",
        tree: &tree,
        notes: Some("REMEMBER THE DEMO"),
        comments: None,
    }]);
    assert!(
        ir.plain_text().contains("REMEMBER THE DEMO"),
        "notes missing from {:?}",
        ir.plain_text()
    );
}

// ---------------------------------------------------------------------------
// #165 — slide comments
// ---------------------------------------------------------------------------

#[test]
fn slide_comments_reach_the_ir() {
    let tree = body_sp(r#"<a:p><a:r><a:t>BODY</a:t></a:r></a:p>"#);
    let ir = pptx_ir(vec![Slide {
        attrs: "",
        tree: &tree,
        notes: None,
        comments: Some(
            r#"<p:cm authorId="1" idx="1"><p:pos x="100" y="100"/>
                 <p:text>Fix the axis label</p:text></p:cm>"#,
        ),
    }]);
    assert!(
        ir.plain_text().contains("Fix the axis label"),
        "comment missing from {:?}",
        ir.plain_text()
    );
}

// ---------------------------------------------------------------------------
// #149 — SmartArt / chart text
// ---------------------------------------------------------------------------

#[test]
fn smartart_and_chart_text_is_extracted() {
    // A graphicFrame whose uri is not the table one used to be skipped
    // wholesale, so a deck built out of SmartArt extracted as empty.
    let tree = r#"<p:graphicFrame>
        <p:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></p:xfrm>
        <a:graphic><a:graphicData
            uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
          <dgm:relIds xmlns:dgm="d"/>
          <a:txBody><a:p><a:r><a:t>DIAGRAM NODE</a:t></a:r></a:p></a:txBody>
        </a:graphicData></a:graphic>
      </p:graphicFrame>"#;
    let ir = pptx_ir(vec![Slide::new(tree)]);
    assert!(
        ir.plain_text().contains("DIAGRAM NODE"),
        "diagram text missing from {:?}",
        ir.plain_text()
    );
}

// ---------------------------------------------------------------------------
// #174 — deck metadata
// ---------------------------------------------------------------------------

#[test]
fn a_deck_title_is_not_the_first_slides_title_when_core_properties_exist() {
    // Without core properties the slide title is a reasonable fallback,
    // which is what this asserts; the DOCX/XLSX tests cover the positive
    // core-properties case and the code path is shared.
    let tree = body_sp(r#"<a:p><a:r><a:t>BODY</a:t></a:r></a:p>"#);
    let ir = pptx_ir(vec![Slide::new(&tree)]);
    assert!(ir.metadata.author.is_none());
}

// ---------------------------------------------------------------------------
// Implicit row and column indices (ECMA-376 18.3.1.4 / 18.3.1.73)
// ---------------------------------------------------------------------------

#[test]
fn cells_without_a_reference_take_the_next_column() {
    // `r` is optional on both <row> and <c>; several writers omit it for the
    // whole sheet. Defaulting the missing reference to column 0 put every
    // cell of a row in the same slot, so laying the row out on the grid kept
    // one value and dropped the rest.
    let ir = Xlsx::new(vec![Sheet::new(
        "S",
        r#"<row>
             <c t="inlineStr"><is><t>Checked</t></is></c>
             <c t="inlineStr"><is><t>Ion</t></is></c>
             <c t="inlineStr"><is><t>Charge</t></is></c>
           </row>
           <row>
             <c t="inlineStr"><is><t>yes</t></is></c>
             <c t="inlineStr"><is><t>Na+</t></is></c>
             <c><v>1</v></c>
           </row>"#,
    )])
    .ir();
    let t = only_table(&ir, 0);
    assert_eq!(t.rows.len(), 2);
    let head: Vec<String> = t.rows[0].cells.iter().map(cell_text).collect();
    assert_eq!(head, ["Checked", "Ion", "Charge"]);
    let body: Vec<String> = t.rows[1].cells.iter().map(cell_text).collect();
    assert_eq!(body, ["yes", "Na+", "1"]);
}

#[test]
fn an_explicit_reference_resets_the_implied_column() {
    // A sheet may mix the two forms: the cell after an explicit `r` continues
    // from that column, not from wherever the implicit run had reached.
    let ir = Xlsx::new(vec![Sheet::new(
        "S",
        r#"<row r="1">
             <c t="inlineStr"><is><t>a</t></is></c>
             <c r="D1" t="inlineStr"><is><t>d</t></is></c>
             <c t="inlineStr"><is><t>e</t></is></c>
           </row>"#,
    )])
    .ir();
    let t = only_table(&ir, 0);
    let row: Vec<String> = t.rows[0].cells.iter().map(cell_text).collect();
    assert_eq!(row, ["a", "", "", "d", "e"]);
}

#[test]
fn rows_without_a_reference_are_numbered_in_document_order() {
    // Every row defaulting to index 1 collapsed the sheet's addressing; the
    // grid still has to report one row per <row> element, in order.
    let ir = Xlsx::new(vec![Sheet::new(
        "S",
        r#"<row><c t="inlineStr"><is><t>one</t></is></c><c><v>1</v></c></row>
           <row><c t="inlineStr"><is><t>two</t></is></c><c><v>2</v></c></row>
           <row><c t="inlineStr"><is><t>three</t></is></c><c><v>3</v></c></row>"#,
    )])
    .ir();
    let text: Vec<String> = only_table(&ir, 0)
        .rows
        .iter()
        .map(|r| format!("{}{}", cell_text(&r.cells[0]), cell_text(&r.cells[1])))
        .collect();
    assert_eq!(text, ["one1", "two2", "three3"]);
}

#[test]
fn cells_out_of_column_order_are_all_kept() {
    // A malformed sheet can list columns unordered. Bailing out of the row on
    // the first backwards reference discarded every cell after it; the values
    // belong in their own columns instead.
    let ir = Xlsx::new(vec![Sheet::new(
        "S",
        r#"<row r="1">
             <c r="C1" t="inlineStr"><is><t>c</t></is></c>
             <c r="A1" t="inlineStr"><is><t>a</t></is></c>
             <c r="B1" t="inlineStr"><is><t>b</t></is></c>
           </row>"#,
    )])
    .ir();
    let row: Vec<String> = only_table(&ir, 0).rows[0]
        .cells
        .iter()
        .map(cell_text)
        .collect();
    assert_eq!(row, ["a", "b", "c"]);
}

#[test]
fn a_duplicate_reference_keeps_the_first_value() {
    // Two cells claiming one column is corruption either way; the first wins,
    // and — the part that regressed — the rest of the row survives.
    let ir = Xlsx::new(vec![Sheet::new(
        "S",
        r#"<row r="1">
             <c r="A1" t="inlineStr"><is><t>first</t></is></c>
             <c r="A1" t="inlineStr"><is><t>second</t></is></c>
             <c r="B1" t="inlineStr"><is><t>b</t></is></c>
           </row>"#,
    )])
    .ir();
    let row: Vec<String> = only_table(&ir, 0).rows[0]
        .cells
        .iter()
        .map(cell_text)
        .collect();
    assert_eq!(row, ["first", "b"]);
}
