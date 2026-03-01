use std::io::Cursor;

use xlsx_oxide::{CellValue, SheetState, XlsxDocument};
use office_core::opc::{OpcWriter, PartName};
use office_core::relationships::{rel_types, TargetMode};

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

struct XlsxBuilder {
    writer: OpcWriter<Cursor<Vec<u8>>>,
    workbook_part: PartName,
    sheet_count: u32,
}

impl XlsxBuilder {
    fn new() -> Self {
        let cursor = Cursor::new(Vec::new());
        let mut writer = OpcWriter::new(cursor).unwrap();
        let workbook_part = PartName::new("/xl/workbook.xml").unwrap();
        writer.add_package_rel(rel_types::OFFICE_DOCUMENT, "xl/workbook.xml");
        Self {
            writer,
            workbook_part,
            sheet_count: 0,
        }
    }

    fn with_workbook(mut self, xml: &[u8]) -> Self {
        self.writer
            .add_part(
                &self.workbook_part,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                xml,
            )
            .unwrap();
        self
    }

    fn with_worksheet(mut self, rel_target: &str, xml: &[u8]) -> Self {
        self.sheet_count += 1;
        let part_path = format!("/xl/{}", rel_target);
        let part = PartName::new(&part_path).unwrap();
        self.writer
            .add_part(
                &part,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
                xml,
            )
            .unwrap();
        self.writer.add_part_rel(
            &self.workbook_part,
            rel_types::WORKSHEET,
            rel_target,
        );
        self
    }

    fn with_shared_strings(mut self, xml: &[u8]) -> Self {
        let part = PartName::new("/xl/sharedStrings.xml").unwrap();
        self.writer
            .add_part(
                &part,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
                xml,
            )
            .unwrap();
        self.writer.add_part_rel(
            &self.workbook_part,
            rel_types::SHARED_STRINGS,
            "sharedStrings.xml",
        );
        self
    }

    fn with_styles(mut self, xml: &[u8]) -> Self {
        let part = PartName::new("/xl/styles.xml").unwrap();
        self.writer
            .add_part(
                &part,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
                xml,
            )
            .unwrap();
        self.writer.add_part_rel(
            &self.workbook_part,
            rel_types::STYLES,
            "styles.xml",
        );
        self
    }

    fn with_sheet_hyperlink(mut self, sheet_rel_target: &str, url: &str) -> Self {
        let part_path = format!("/xl/{}", sheet_rel_target);
        let part = PartName::new(&part_path).unwrap();
        self.writer.add_part_rel_with_mode(
            &part,
            rel_types::HYPERLINK,
            url,
            TargetMode::External,
        );
        self
    }

    fn build(self) -> Vec<u8> {
        let result = self.writer.finish().unwrap();
        result.into_inner()
    }
}

fn parse(data: &[u8]) -> XlsxDocument {
    XlsxDocument::from_reader(Cursor::new(data.to_vec())).unwrap()
}

// ---------------------------------------------------------------------------
// 1. Simple worksheet: Numbers, strings, booleans
// ---------------------------------------------------------------------------

#[test]
fn simple_worksheet() {
    let wb_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let ws_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Name</t></is></c>
      <c r="B1" t="inlineStr"><is><t>Value</t></is></c>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>Pi</t></is></c>
      <c r="B2"><v>3.14159</v></c>
    </row>
    <row r="3">
      <c r="A3" t="inlineStr"><is><t>Active</t></is></c>
      <c r="B3" t="b"><v>1</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    let data = XlsxBuilder::new()
        .with_workbook(wb_xml)
        .with_worksheet("worksheets/sheet1.xml", ws_xml)
        .build();
    let doc = parse(&data);

    assert_eq!(doc.workbook.sheets.len(), 1);
    assert_eq!(doc.workbook.sheets[0].name, "Sheet1");
    assert_eq!(doc.worksheets.len(), 1);
    assert_eq!(doc.worksheets[0].rows.len(), 3);

    let text = doc.plain_text();
    assert!(text.contains("Name\tValue"), "text was: {text}");
    assert!(text.contains("Pi\t3.14159"), "text was: {text}");
    assert!(text.contains("Active\tTRUE"), "text was: {text}");
}

// ---------------------------------------------------------------------------
// 2. Shared strings: Multiple cells referencing SST
// ---------------------------------------------------------------------------

#[test]
fn shared_strings_round_trip() {
    let wb_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Data" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let sst_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
  <si><t>Hello</t></si>
  <si><t>World</t></si>
  <si><t>Shared</t></si>
</sst>"#;

    let ws_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1" t="s"><v>1</v></c>
    </row>
    <row r="2">
      <c r="A2" t="s"><v>2</v></c>
      <c r="B2" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    // Worksheet rel added first → rId1, SST rel → rId2
    let data = XlsxBuilder::new()
        .with_workbook(wb_xml)
        .with_worksheet("worksheets/sheet1.xml", ws_xml)
        .with_shared_strings(sst_xml)
        .build();
    let doc = parse(&data);

    // Shared strings should be dereferenced
    let text = doc.plain_text();
    assert!(text.contains("Hello\tWorld"), "text was: {text}");
    assert!(text.contains("Shared\tHello"), "text was: {text}");
}

// ---------------------------------------------------------------------------
// 3. Date detection: Numeric cells with date format -> ISO date output
// ---------------------------------------------------------------------------

#[test]
fn date_detection() {
    let wb_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Dates" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let styles_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>
  </numFmts>
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/></border></borders>
  <cellXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="14" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
  </cellXfs>
</styleSheet>"#;

    // 45306 = Jan 15, 2024 in 1900 system
    let ws_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Date (builtin)</t></is></c>
      <c r="B1" s="1"><v>45306</v></c>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>Date (custom)</t></is></c>
      <c r="B2" s="2"><v>45306</v></c>
    </row>
    <row r="3">
      <c r="A3" t="inlineStr"><is><t>Number (not date)</t></is></c>
      <c r="B3" s="0"><v>45306</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    // Worksheet rel added first → rId1, styles rel → rId2
    let data = XlsxBuilder::new()
        .with_workbook(wb_xml)
        .with_worksheet("worksheets/sheet1.xml", ws_xml)
        .with_styles(styles_xml)
        .build();
    let doc = parse(&data);

    let text = doc.plain_text();
    // Built-in date format (ID 14) should output ISO date
    assert!(text.contains("Date (builtin)\t2024-01-15"), "text was: {text}");
    // Custom date format (yyyy-mm-dd) should output ISO date
    assert!(text.contains("Date (custom)\t2024-01-15"), "text was: {text}");
    // Plain number (format 0) should NOT be converted to date
    assert!(text.contains("Number (not date)\t45306"), "text was: {text}");
}

// ---------------------------------------------------------------------------
// 4. Merged cells: Merge ranges parsed correctly
// ---------------------------------------------------------------------------

#[test]
fn merged_cells() {
    let wb_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let ws_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Merged Title</t></is></c>
    </row>
  </sheetData>
  <mergeCells count="2">
    <mergeCell ref="A1:C1"/>
    <mergeCell ref="A3:A5"/>
  </mergeCells>
</worksheet>"#;

    let data = XlsxBuilder::new()
        .with_workbook(wb_xml)
        .with_worksheet("worksheets/sheet1.xml", ws_xml)
        .build();
    let doc = parse(&data);

    assert_eq!(doc.worksheets[0].merged_cells.len(), 2);
    assert_eq!(doc.worksheets[0].merged_cells[0], "A1:C1");
    assert_eq!(doc.worksheets[0].merged_cells[1], "A3:A5");
}

// ---------------------------------------------------------------------------
// 5. Multiple sheets: Workbook with 3 sheets
// ---------------------------------------------------------------------------

#[test]
fn multiple_sheets() {
    let wb_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sales" sheetId="1" r:id="rId1"/>
    <sheet name="Expenses" sheetId="2" r:id="rId2"/>
    <sheet name="Summary" sheetId="3" r:id="rId3"/>
  </sheets>
</workbook>"#;

    let ws1 = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Sales Data</t></is></c></row>
  </sheetData>
</worksheet>"#;

    let ws2 = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Expense Data</t></is></c></row>
  </sheetData>
</worksheet>"#;

    let ws3 = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Summary</t></is></c></row>
  </sheetData>
</worksheet>"#;

    let data = XlsxBuilder::new()
        .with_workbook(wb_xml)
        .with_worksheet("worksheets/sheet1.xml", ws1)
        .with_worksheet("worksheets/sheet2.xml", ws2)
        .with_worksheet("worksheets/sheet3.xml", ws3)
        .build();
    let doc = parse(&data);

    assert_eq!(doc.worksheets.len(), 3);
    assert_eq!(doc.worksheets[0].name, "Sales");
    assert_eq!(doc.worksheets[1].name, "Expenses");
    assert_eq!(doc.worksheets[2].name, "Summary");

    let text = doc.plain_text();
    assert!(text.contains("Sales Data"), "text was: {text}");
    assert!(text.contains("Expense Data"), "text was: {text}");
    assert!(text.contains("Summary"), "text was: {text}");
}

// ---------------------------------------------------------------------------
// 6. Hyperlinks: External and internal
// ---------------------------------------------------------------------------

#[test]
fn hyperlinks() {
    let wb_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Links" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let ws_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Click me</t></is></c>
      <c r="B1" t="inlineStr"><is><t>Internal</t></is></c>
    </row>
  </sheetData>
  <hyperlinks>
    <hyperlink ref="A1" r:id="rId1" tooltip="Visit example"/>
    <hyperlink ref="B1" location="Sheet2!A1"/>
  </hyperlinks>
</worksheet>"#;

    let data = XlsxBuilder::new()
        .with_workbook(wb_xml)
        .with_worksheet("worksheets/sheet1.xml", ws_xml)
        .with_sheet_hyperlink("worksheets/sheet1.xml", "https://example.com")
        .build();
    let doc = parse(&data);

    assert_eq!(doc.worksheets[0].hyperlinks.len(), 2);

    let hl1 = &doc.worksheets[0].hyperlinks[0];
    assert_eq!(hl1.cell_ref, "A1");
    assert_eq!(hl1.tooltip.as_deref(), Some("Visit example"));
    match &hl1.target {
        xlsx_oxide::HyperlinkTarget::External(url) => assert_eq!(url, "https://example.com"),
        _ => panic!("expected external hyperlink"),
    }

    let hl2 = &doc.worksheets[0].hyperlinks[1];
    assert_eq!(hl2.cell_ref, "B1");
    match &hl2.target {
        xlsx_oxide::HyperlinkTarget::Internal(loc) => assert_eq!(loc, "Sheet2!A1"),
        _ => panic!("expected internal hyperlink"),
    }
}

// ---------------------------------------------------------------------------
// 7. Formulas: Cells with formula + cached value
// ---------------------------------------------------------------------------

#[test]
fn formulas_with_cached_values() {
    let wb_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Calc" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let ws_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>10</v></c>
      <c r="B1"><v>20</v></c>
      <c r="C1"><f>A1+B1</f><v>30</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    let data = XlsxBuilder::new()
        .with_workbook(wb_xml)
        .with_worksheet("worksheets/sheet1.xml", ws_xml)
        .build();
    let doc = parse(&data);

    let cell_c1 = &doc.worksheets[0].rows[0].cells[2];
    assert_eq!(cell_c1.formula.as_deref(), Some("A1+B1"));
    assert!(matches!(cell_c1.value, CellValue::Number(n) if n == 30.0));

    let text = doc.plain_text();
    assert!(text.contains("10\t20\t30"), "text was: {text}");
}

// ---------------------------------------------------------------------------
// 8. Styles resolution: Cell format -> font/number-format chain
// ---------------------------------------------------------------------------

#[test]
fn styles_resolution() {
    let wb_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Styled" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let styles_xml = br##"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="#,##0.00"/>
  </numFmts>
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="14"/><name val="Arial"/><color rgb="FFFF0000"/></font>
  </fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/></border></borders>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="164" fontId="1" fillId="0" borderId="0" applyNumberFormat="1"/>
  </cellXfs>
</styleSheet>"##;

    let ws_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" s="1"><v>1234.5</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    // Worksheet rel added first → rId1, styles rel → rId2
    let data = XlsxBuilder::new()
        .with_workbook(wb_xml)
        .with_worksheet("worksheets/sheet1.xml", ws_xml)
        .with_styles(styles_xml)
        .build();
    let doc = parse(&data);

    let styles = doc.styles.as_ref().unwrap();
    assert_eq!(styles.number_format_for(1), Some("#,##0.00"));

    let font = styles.font_for(1).unwrap();
    assert!(font.bold);
    assert_eq!(font.name.as_deref(), Some("Arial"));
    assert_eq!(font.size, Some(14.0));
}

// ---------------------------------------------------------------------------
// 9. CSV output: Correct quoting, escaping, CRLF
// ---------------------------------------------------------------------------

#[test]
fn csv_output() {
    let wb_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="CSV" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let sst_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4">
  <si><t>Name</t></si>
  <si><t>Value</t></si>
  <si><t>Hello, World</t></si>
  <si><t>Say "Hi"</t></si>
</sst>"#;

    let ws_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1" t="s"><v>1</v></c>
    </row>
    <row r="2">
      <c r="A2" t="s"><v>2</v></c>
      <c r="B2"><v>42</v></c>
    </row>
    <row r="3">
      <c r="A3" t="s"><v>3</v></c>
      <c r="B3"><v>100</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    // Worksheet rel added first → rId1, SST rel → rId2
    let data = XlsxBuilder::new()
        .with_workbook(wb_xml)
        .with_worksheet("worksheets/sheet1.xml", ws_xml)
        .with_shared_strings(sst_xml)
        .build();
    let doc = parse(&data);

    let csv = doc.to_csv();
    let lines: Vec<&str> = csv.split("\r\n").collect();
    assert_eq!(lines[0], "Name,Value", "csv was: {csv}");
    // "Hello, World" must be quoted because it contains a comma
    assert_eq!(lines[1], "\"Hello, World\",42", "csv was: {csv}");
    // Say "Hi" must have double-escaped quotes
    assert_eq!(lines[2], "\"Say \"\"Hi\"\"\",100", "csv was: {csv}");
}

// ---------------------------------------------------------------------------
// 10. Markdown output: Pipe tables with headers
// ---------------------------------------------------------------------------

#[test]
fn markdown_output() {
    let wb_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Products" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let sst_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4">
  <si><t>Product</t></si>
  <si><t>Price</t></si>
  <si><t>Widget</t></si>
  <si><t>Gadget</t></si>
</sst>"#;

    let ws_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1" t="s"><v>1</v></c>
    </row>
    <row r="2">
      <c r="A2" t="s"><v>2</v></c>
      <c r="B2"><v>9.99</v></c>
    </row>
    <row r="3">
      <c r="A3" t="s"><v>3</v></c>
      <c r="B3"><v>24.5</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    // Worksheet rel added first → rId1, SST rel → rId2
    let data = XlsxBuilder::new()
        .with_workbook(wb_xml)
        .with_worksheet("worksheets/sheet1.xml", ws_xml)
        .with_shared_strings(sst_xml)
        .build();
    let doc = parse(&data);

    let md = doc.to_markdown();
    assert!(md.contains("## Products"), "md was: {md}");
    assert!(md.contains("| Product | Price |"), "md was: {md}");
    assert!(md.contains("| --- | --- |"), "md was: {md}");
    assert!(md.contains("| Widget | 9.99 |"), "md was: {md}");
    assert!(md.contains("| Gadget | 24.5 |"), "md was: {md}");
}

// ---------------------------------------------------------------------------
// 11. Empty workbook: No data
// ---------------------------------------------------------------------------

#[test]
fn empty_workbook() {
    let wb_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Empty" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let ws_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"#;

    let data = XlsxBuilder::new()
        .with_workbook(wb_xml)
        .with_worksheet("worksheets/sheet1.xml", ws_xml)
        .build();
    let doc = parse(&data);

    assert_eq!(doc.worksheets.len(), 1);
    assert!(doc.worksheets[0].rows.is_empty());
    assert_eq!(doc.plain_text(), "");
}

// ---------------------------------------------------------------------------
// 12. Hidden sheets: SheetState::Hidden/VeryHidden
// ---------------------------------------------------------------------------

#[test]
fn hidden_sheets() {
    let wb_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Visible" sheetId="1" r:id="rId1"/>
    <sheet name="Hidden" sheetId="2" r:id="rId2" state="hidden"/>
    <sheet name="VeryHidden" sheetId="3" r:id="rId3" state="veryHidden"/>
  </sheets>
</workbook>"#;

    let ws = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"#;

    let data = XlsxBuilder::new()
        .with_workbook(wb_xml)
        .with_worksheet("worksheets/sheet1.xml", ws)
        .with_worksheet("worksheets/sheet2.xml", ws)
        .with_worksheet("worksheets/sheet3.xml", ws)
        .build();
    let doc = parse(&data);

    assert_eq!(doc.workbook.sheets[0].state, SheetState::Visible);
    assert_eq!(doc.workbook.sheets[1].state, SheetState::Hidden);
    assert_eq!(doc.workbook.sheets[2].state, SheetState::VeryHidden);
}

// ---------------------------------------------------------------------------
// 13. Large cell refs: XFD1048576 (max cell reference)
// ---------------------------------------------------------------------------

#[test]
fn large_cell_reference() {
    let wb_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Big" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let ws_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1048576">
      <c r="XFD1048576"><v>999</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    let data = XlsxBuilder::new()
        .with_workbook(wb_xml)
        .with_worksheet("worksheets/sheet1.xml", ws_xml)
        .build();
    let doc = parse(&data);

    let row = &doc.worksheets[0].rows[0];
    assert_eq!(row.index, 1048576);
    let cell = &row.cells[0];
    assert_eq!(cell.reference.col, 16383);
    assert_eq!(cell.reference.row, 1048575);
    assert!(matches!(cell.value, CellValue::Number(n) if n == 999.0));
}

// ---------------------------------------------------------------------------
// 14. Error cells: #DIV/0!, #REF! values
// ---------------------------------------------------------------------------

#[test]
fn error_cells() {
    let wb_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Errors" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let ws_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="e"><v>#DIV/0!</v></c>
      <c r="B1" t="e"><v>#REF!</v></c>
      <c r="C1" t="e"><v>#VALUE!</v></c>
      <c r="D1" t="e"><v>#N/A</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    let data = XlsxBuilder::new()
        .with_workbook(wb_xml)
        .with_worksheet("worksheets/sheet1.xml", ws_xml)
        .build();
    let doc = parse(&data);

    let cells = &doc.worksheets[0].rows[0].cells;
    assert!(matches!(&cells[0].value, CellValue::Error(e) if e == "#DIV/0!"));
    assert!(matches!(&cells[1].value, CellValue::Error(e) if e == "#REF!"));
    assert!(matches!(&cells[2].value, CellValue::Error(e) if e == "#VALUE!"));
    assert!(matches!(&cells[3].value, CellValue::Error(e) if e == "#N/A"));

    let text = doc.plain_text();
    assert!(text.contains("#DIV/0!\t#REF!\t#VALUE!\t#N/A"), "text was: {text}");
}

// ---------------------------------------------------------------------------
// 15. Missing parts: No shared strings (all inline), no styles
// ---------------------------------------------------------------------------

#[test]
fn missing_optional_parts() {
    let wb_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Minimal" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let ws_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Inline only</t></is></c>
      <c r="B1"><v>42</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    // No shared strings, no styles
    let data = XlsxBuilder::new()
        .with_workbook(wb_xml)
        .with_worksheet("worksheets/sheet1.xml", ws_xml)
        .build();
    let doc = parse(&data);

    assert!(doc.styles.is_none());
    assert_eq!(doc.shared_strings.strings.len(), 0);
    assert_eq!(doc.plain_text(), "Inline only\t42");
}

// ---------------------------------------------------------------------------
// Missing workbook error
// ---------------------------------------------------------------------------

#[test]
fn missing_workbook_part() {
    let cursor = Cursor::new(Vec::new());
    let mut writer = OpcWriter::new(cursor).unwrap();
    let part = PartName::new("/xl/other.xml").unwrap();
    writer
        .add_part(&part, "application/xml", b"<root/>")
        .unwrap();
    let result = writer.finish().unwrap();
    let data = result.into_inner();

    let result = XlsxDocument::from_reader(Cursor::new(data));
    assert!(result.is_err());
}
