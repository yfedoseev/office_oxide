# XLSX (Office Open XML SpreadsheetML) Format Specification Reference

## Table of Contents

1. [Official Specification References](#1-official-specification-references)
2. [Package Structure (OPC)](#2-package-structure-opc)
3. [XML Namespaces and Schemas](#3-xml-namespaces-and-schemas)
4. [Content Types](#4-content-types)
5. [Relationships](#5-relationships)
6. [Workbook](#6-workbook)
7. [Worksheets](#7-worksheets)
8. [Cell Data Model](#8-cell-data-model)
9. [Shared String Table](#9-shared-string-table)
10. [Styles](#10-styles)
11. [Formulas](#11-formulas)
12. [Defined Names](#12-defined-names)
13. [Tables](#13-tables)
14. [Conditional Formatting](#14-conditional-formatting)
15. [Charts](#15-charts)
16. [Pivot Tables](#16-pivot-tables)
17. [Additional Worksheet Features](#17-additional-worksheet-features)
18. [Implementation Notes](#18-implementation-notes)

---

## 1. Official Specification References

### ECMA-376: Office Open XML File Formats

The primary standard is **ECMA-376**, published by Ecma International. It consists of five parts:

| Part | Title | Content |
|------|-------|---------|
| **Part 1** | Fundamentals and Markup Language Reference | Complete element/attribute reference for SpreadsheetML, WordprocessingML, PresentationML, DrawingML, and shared markup |
| **Part 2** | Open Packaging Conventions (OPC) | ZIP-based package format, relationships, content types |
| **Part 3** | Markup Compatibility and Extensibility | Extension mechanisms for forward/backward compatibility |
| **Part 4** | Transitional Migration Features | Legacy features for backward compatibility with binary formats |
| **Part 5** | Markup Compatibility and Extensibility (additional) | Additional compatibility guidance |

Current edition: **ECMA-376, 5th Edition (December 2021)**
URL: https://ecma-international.org/publications-and-standards/standards/ecma-376/

### ISO/IEC 29500: Information Technology -- Office Open XML File Formats

The ISO counterpart, aligned with ECMA-376:

| Part | ISO Standard |
|------|-------------|
| Part 1 | ISO/IEC 29500-1:2016 -- Fundamentals and Markup Language Reference |
| Part 2 | ISO/IEC 29500-2:2021 -- Open Packaging Conventions |
| Part 3 | ISO/IEC 29500-3:2015 -- Markup Compatibility and Extensibility |
| Part 4 | ISO/IEC 29500-4:2016 -- Transitional Migration Features |

### Transitional vs. Strict Conformance

Two conformance classes exist:

- **Transitional** (most common): Uses `http://schemas.openxmlformats.org/` namespaces. Allows legacy features. This is what Excel produces by default.
- **Strict**: Uses `http://purl.oclc.org/ooxml/` namespaces. Removes deprecated features. Supported since Excel 2013.

An implementer should focus on **Transitional** conformance for maximum compatibility.

### Microsoft Extension Specifications

- **[MS-XLSX]**: Excel (.xlsx) Extensions to Office Open XML SpreadsheetML
- **[MS-OE376]**: Office Implementation Information for ECMA-376
- **[MS-OI29500]**: Office Implementation Information for ISO/IEC 29500

These document Microsoft-specific behaviors and deviations from the standard.

---

## 2. Package Structure (OPC)

An `.xlsx` file is a **ZIP archive** conforming to the Open Packaging Conventions (OPC). Renaming `.xlsx` to `.zip` reveals the internal structure.

### Minimum Package Layout

```
[Content_Types].xml
_rels/
  .rels
xl/
  _rels/
    workbook.xml.rels
  workbook.xml
  worksheets/
    sheet1.xml
```

### Typical Full Package Layout

```
[Content_Types].xml
_rels/
  .rels
docProps/
  app.xml                          # Application properties
  core.xml                         # Core properties (author, dates, etc.)
  custom.xml                       # Custom properties (optional)
xl/
  _rels/
    workbook.xml.rels
  workbook.xml                     # Workbook definition
  styles.xml                       # Style definitions
  sharedStrings.xml                # Shared string table
  theme/
    theme1.xml                     # Theme definitions
  worksheets/
    _rels/
      sheet1.xml.rels              # Per-sheet relationships
      sheet2.xml.rels
    sheet1.xml                     # Worksheet data
    sheet2.xml
  charts/
    chart1.xml                     # Chart definitions
  chartsheets/
    sheet1.xml                     # Chart sheets
  drawings/
    _rels/
      drawing1.xml.rels
    drawing1.xml                   # Drawing anchors
  tables/
    table1.xml                     # Table definitions
  pivotTables/
    pivotTable1.xml                # Pivot table definitions
  pivotCache/
    pivotCacheDefinition1.xml      # Pivot cache metadata
    pivotCacheRecords1.xml         # Pivot cache data
  printerSettings/
    printerSettings1.bin           # Binary printer settings
  calcChain.xml                    # Calculation chain
  externalLinks/
    externalLink1.xml              # External workbook references
  media/
    image1.png                     # Embedded images
```

### ZIP Requirements

- Standard ZIP (PKZIP) compression
- Files may be stored (no compression) or deflated
- UTF-8 encoding for XML parts
- No encryption at the ZIP level (encryption uses a different container)

---

## 3. XML Namespaces and Schemas

### Primary Namespaces (Transitional)

| Prefix | Namespace URI | Usage |
|--------|--------------|-------|
| (default) | `http://schemas.openxmlformats.org/spreadsheetml/2006/main` | SpreadsheetML elements |
| `r` | `http://schemas.openxmlformats.org/officeDocument/2006/relationships` | Relationship references |
| `rel` | `http://schemas.openxmlformats.org/package/2006/relationships` | Package relationships |
| `ct` | `http://schemas.openxmlformats.org/package/2006/content-types` | Content types |
| `a` | `http://schemas.openxmlformats.org/drawingml/2006/main` | DrawingML main |
| `c` | `http://schemas.openxmlformats.org/drawingml/2006/chart` | DrawingML charts |
| `xdr` | `http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing` | Spreadsheet drawing |
| `mc` | `http://schemas.openxmlformats.org/markup-compatibility/2006` | Markup compatibility |
| `dcterms` | `http://purl.org/dc/terms/` | Dublin Core terms (in core.xml) |
| `dc` | `http://purl.org/dc/elements/1.1/` | Dublin Core elements |
| `cp` | `http://schemas.openxmlformats.org/package/2006/metadata/core-properties` | Core properties |

### Microsoft Extension Namespaces

| Prefix | Namespace URI | Usage |
|--------|--------------|-------|
| `x14ac` | `http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac` | Excel 2010 extensions |
| `x15` | `http://schemas.microsoft.com/office/spreadsheetml/2010/11/main` | Excel 2013 extensions |
| `x16r2` | `http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2` | Rich data |

### Strict Namespaces

In Strict mode, the main namespace becomes:
`http://purl.oclc.org/ooxml/spreadsheetml/main`

And relationship namespaces change to:
`http://purl.oclc.org/ooxml/officeDocument/relationships`

---

## 4. Content Types

Every package must contain `[Content_Types].xml` at the root. It maps parts to MIME content types.

### XML Structure

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <!-- Default types by extension -->
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>
  <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"/>

  <!-- Override types for specific parts -->
  <Override PartName="/xl/workbook.xml"
    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml"
    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml"
    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml"
    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/theme/theme1.xml"
    ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/docProps/core.xml"
    ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml"
    ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
```

### Complete Content Type Reference

| Part | Content Type |
|------|-------------|
| Workbook | `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml` |
| Worksheet | `application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml` |
| Chartsheet | `application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml` |
| Dialogsheet | `application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml` |
| Macro sheet | `application/vnd.openxmlformats-officedocument.spreadsheetml.macrosheet+xml` |
| Shared Strings | `application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml` |
| Styles | `application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml` |
| Theme | `application/vnd.openxmlformats-officedocument.theme+xml` |
| Chart | `application/vnd.openxmlformats-officedocument.drawingml.chart+xml` |
| Drawing | `application/vnd.openxmlformats-officedocument.drawing+xml` |
| Table | `application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml` |
| Pivot Table | `application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml` |
| Pivot Cache Def | `application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml` |
| Pivot Cache Recs | `application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml` |
| Calc Chain | `application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml` |
| Comments | `application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml` |
| Core Properties | `application/vnd.openxmlformats-package.core-properties+xml` |
| Extended Properties | `application/vnd.openxmlformats-officedocument.extended-properties+xml` |
| Custom Properties | `application/vnd.openxmlformats-officedocument.custom-properties+xml` |
| Relationships | `application/vnd.openxmlformats-package.relationships+xml` |
| VBA Project | `application/vnd.ms-office.vbaProject` |
| Printer Settings | `application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings` |

**Note:** For `.xlsm` (macro-enabled), the workbook content type changes to:
`application/vnd.ms-excel.sheet.macroEnabled.main+xml`

---

## 5. Relationships

Relationships define how parts connect to each other. They are stored in `_rels/` subdirectories with `.rels` extension.

### Package-Level Relationships (`_rels/.rels`)

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="xl/workbook.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
    Target="docProps/core.xml"/>
  <Relationship Id="rId3"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
    Target="docProps/app.xml"/>
</Relationships>
```

### Workbook-Level Relationships (`xl/_rels/workbook.xml.rels`)

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    Target="styles.xml"/>
  <Relationship Id="rId4"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
    Target="sharedStrings.xml"/>
  <Relationship Id="rId5"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
    Target="theme/theme1.xml"/>
  <Relationship Id="rId6"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain"
    Target="calcChain.xml"/>
  <Relationship Id="rId7"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition"
    Target="pivotCache/pivotCacheDefinition1.xml"/>
</Relationships>
```

### Worksheet-Level Relationships (`xl/worksheets/_rels/sheet1.xml.rels`)

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
    Target="../drawings/drawing1.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"
    Target="../tables/table1.xml"/>
  <Relationship Id="rId3"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable"
    Target="../pivotTables/pivotTable1.xml"/>
  <Relationship Id="rId4"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings"
    Target="../printerSettings/printerSettings1.bin"/>
  <Relationship Id="rId5"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
    Target="../comments1.xml"/>
  <Relationship Id="rId6"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
    Target="https://example.com" TargetMode="External"/>
</Relationships>
```

### Relationship Type URIs

| Relationship | Type URI |
|-------------|----------|
| Office Document | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument` |
| Worksheet | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet` |
| Chartsheet | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet` |
| Shared Strings | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings` |
| Styles | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles` |
| Theme | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme` |
| Chart | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart` |
| Drawing | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing` |
| Table | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/table` |
| Pivot Table | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable` |
| Pivot Cache Def | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition` |
| Pivot Cache Recs | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords` |
| Calc Chain | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain` |
| Comments | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments` |
| Hyperlink | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink` |
| Printer Settings | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings` |
| Core Properties | `http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties` |
| Extended Properties | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties` |
| Custom Properties | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties` |
| Image | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/image` |
| External Link | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink` |
| VBA Project | `http://schemas.microsoft.com/office/2006/relationships/vbaProject` |

### Navigation Algorithm

1. Parse `[Content_Types].xml` to discover the workbook part location
2. Parse `_rels/.rels` to find the office document relationship (workbook)
3. Parse `xl/_rels/workbook.xml.rels` to locate all workbook-related parts
4. For each worksheet, parse `xl/worksheets/_rels/sheetN.xml.rels` for sheet-specific parts

**Critical:** The `r:id` attribute on `<sheet>` elements in the workbook is the **only reliable** way to map sheets to their worksheet XML files. Do not rely on `sheetId` or filename patterns.

---

## 6. Workbook

The workbook part (`xl/workbook.xml`) is the root of the spreadsheet document.

### Complete Workbook Example

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
          xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
          mc:Ignorable="x15"
          xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">

  <fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="27328"/>

  <workbookPr defaultThemeVersion="166925"
              date1904="false"
              filterPrivacy="false"/>

  <bookViews>
    <workbookView xWindow="-120" yWindow="-120"
                  windowWidth="29040" windowHeight="15720"
                  activeTab="0"/>
  </bookViews>

  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Sheet2" sheetId="2" r:id="rId2"/>
    <sheet name="Hidden Sheet" sheetId="3" state="hidden" r:id="rId3"/>
  </sheets>

  <definedNames>
    <definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$A$1:$G$50</definedName>
    <definedName name="_xlnm.Print_Titles" localSheetId="0">Sheet1!$1:$2</definedName>
    <definedName name="SalesData">Sheet1!$A$1:$D$100</definedName>
    <definedName name="TaxRate">0.08</definedName>
  </definedNames>

  <calcPr calcId="191029" calcMode="auto" fullCalcOnLoad="false"/>

  <pivotCaches>
    <pivotCache cacheId="0" r:id="rId7"/>
  </pivotCaches>
</workbook>
```

### Key Workbook Elements

| Element | Description |
|---------|-------------|
| `<fileVersion>` | Application version info |
| `<workbookPr>` | Workbook properties (date system, theme, etc.) |
| `<bookViews>` | Workbook view settings (window position, active tab) |
| `<sheets>` | Container for all sheet references |
| `<sheet>` | Individual sheet reference with name, ID, relationship ID |
| `<definedNames>` | Named ranges, print areas, print titles |
| `<calcPr>` | Calculation properties (auto/manual calc, iteration) |
| `<pivotCaches>` | Pivot cache references |
| `<externalReferences>` | References to external workbooks |
| `<functionGroups>` | Custom function group definitions |
| `<fileRecoveryPr>` | File recovery properties |

### Sheet Element Attributes

```xml
<sheet name="Sheet1"      <!-- Display name (max 31 chars, unique) -->
       sheetId="1"        <!-- Persistent unique ID (not an index) -->
       state="visible"    <!-- visible | hidden | veryHidden -->
       r:id="rId1"/>      <!-- Relationship ID -> worksheet part -->
```

The `state` attribute values:
- `visible` (default): Sheet tab is visible
- `hidden`: Hidden but can be unhidden via UI
- `veryHidden`: Can only be unhidden programmatically

### Date System (`workbookPr`)

The `date1904` attribute on `<workbookPr>` controls the date system:
- `false` (default): **1900 date system** -- serial date 1 = January 1, 1900
- `true`: **1904 date system** -- serial date 0 = January 1, 1904

**Important:** The 1900 date system has a known bug inherited from Lotus 1-2-3 where it incorrectly treats 1900 as a leap year. Serial date 60 = February 29, 1900 (which never existed). Serial dates 1-59 are off by one day.

---

## 7. Worksheets

Each worksheet is stored in a separate XML file (e.g., `xl/worksheets/sheet1.xml`).

### Complete Worksheet Structure

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
           xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
           mc:Ignorable="x14ac"
           xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">

  <sheetPr>
    <tabColor rgb="FF0000FF"/>
    <outlinePr summaryBelow="true" summaryRight="true"/>
    <pageSetUpPr fitToPage="false"/>
  </sheetPr>

  <dimension ref="A1:G50"/>

  <sheetViews>
    <sheetView tabSelected="true" workbookViewId="0">
      <pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"/>
      <selection pane="bottomRight" activeCell="B3" sqref="B3"/>
    </sheetView>
  </sheetViews>

  <sheetFormatPr defaultColWidth="8.43" defaultRowHeight="15" x14ac:dyDescent="0.25"/>

  <cols>
    <col min="1" max="1" width="20" style="1" customWidth="true"/>
    <col min="2" max="4" width="12.5" bestFit="true" customWidth="true"/>
    <col min="5" max="5" width="15" hidden="true"/>
  </cols>

  <sheetData>
    <!-- Row and cell data (see Cell Data Model section) -->
    <row r="1" spans="1:7" ht="20" customHeight="true" x14ac:dyDescent="0.25">
      <c r="A1" t="s" s="1"><v>0</v></c>
      <c r="B1" t="s" s="1"><v>1</v></c>
    </row>
    <row r="2" spans="1:7" x14ac:dyDescent="0.25">
      <c r="A2" t="s"><v>2</v></c>
      <c r="B2" s="2"><v>42500</v></c>
    </row>
  </sheetData>

  <sheetProtection sheet="true" password="ABCD" objects="true" scenarios="true"/>

  <autoFilter ref="A1:G50"/>

  <mergeCells count="1">
    <mergeCell ref="A1:C1"/>
  </mergeCells>

  <conditionalFormatting sqref="B2:B50">
    <cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan">
      <formula>50000</formula>
    </cfRule>
  </conditionalFormatting>

  <dataValidations count="1">
    <dataValidation type="list" allowBlank="true" showInputMessage="true"
                    showErrorMessage="true" sqref="D2:D50">
      <formula1>"Yes,No,Maybe"</formula1>
    </dataValidation>
  </dataValidations>

  <hyperlinks>
    <hyperlink ref="A10" r:id="rId6" tooltip="Click here"/>
    <hyperlink ref="A11" location="Sheet2!A1" display="Go to Sheet2"/>
  </hyperlinks>

  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>

  <pageSetup paperSize="1" orientation="landscape" r:id="rId4"/>

  <headerFooter>
    <oddHeader>&amp;C&amp;"Arial,Bold"Page &amp;P of &amp;N</oddHeader>
    <oddFooter>&amp;L&amp;D&amp;R&amp;F</oddFooter>
  </headerFooter>

  <drawing r:id="rId1"/>

  <tableParts count="1">
    <tablePart r:id="rId2"/>
  </tableParts>
</worksheet>
```

### Worksheet Child Element Order

Elements in a worksheet **must** appear in this order (all optional except `sheetData`):

1. `<sheetPr>` -- Sheet-level properties
2. `<dimension>` -- Used cell range
3. `<sheetViews>` -- View settings
4. `<sheetFormatPr>` -- Default row/column dimensions
5. `<cols>` -- Column definitions
6. **`<sheetData>`** -- **Required.** Cell data (rows and cells)
7. `<sheetCalcPr>` -- Sheet calculation properties
8. `<sheetProtection>` -- Protection settings
9. `<protectedRanges>` -- Protected cell ranges
10. `<scenarios>` -- What-if scenarios
11. `<autoFilter>` -- Auto-filter settings
12. `<sortState>` -- Sort settings
13. `<dataConsolidate>` -- Data consolidation
14. `<customSheetViews>` -- Custom views
15. `<mergeCells>` -- Merged cell ranges
16. `<phoneticPr>` -- Phonetic properties
17. `<conditionalFormatting>` -- Conditional format rules
18. `<dataValidations>` -- Data validation rules
19. `<hyperlinks>` -- Hyperlink definitions
20. `<printOptions>` -- Print options
21. `<pageMargins>` -- Page margin dimensions
22. `<pageSetup>` -- Page setup (paper size, orientation)
23. `<headerFooter>` -- Header/footer content
24. `<rowBreaks>` -- Row page breaks
25. `<colBreaks>` -- Column page breaks
26. `<customProperties>` -- Custom properties
27. `<cellWatches>` -- Cell watch items
28. `<ignoredErrors>` -- Ignored error rules
29. `<smartTags>` -- Smart tag data
30. `<drawing>` -- Drawing part reference
31. `<legacyDrawing>` -- VML drawing reference (comments)
32. `<picture>` -- Background image
33. `<oleObjects>` -- OLE objects
34. `<controls>` -- ActiveX controls
35. `<webPublishItems>` -- Web publish items
36. `<tableParts>` -- Table part references
37. `<extLst>` -- Extension list

### Row Element

```xml
<row r="5"                    <!-- 1-based row number -->
     spans="1:7"              <!-- Column range hint (min:max) -->
     s="2"                    <!-- Style index (from cellXfs) -->
     customFormat="true"      <!-- Row has custom format -->
     ht="20"                  <!-- Row height in points -->
     customHeight="true"      <!-- Height is explicitly set -->
     hidden="false"           <!-- Row is hidden -->
     outlineLevel="0"         <!-- Outline/grouping level (0-7) -->
     collapsed="false"        <!-- Outline group is collapsed -->
     thickTop="false"         <!-- Thick top border -->
     thickBot="false"         <!-- Thick bottom border -->
     ph="false"               <!-- Show phonetic info -->
     x14ac:dyDescent="0.25"> <!-- Font descent for row -->
  <!-- cell elements -->
</row>
```

**Important:** Rows and cells may be omitted when empty. A parser must not assume contiguous row numbers.

### Column Definitions

```xml
<cols>
  <!-- Columns 1-1 (A) with custom width -->
  <col min="1" max="1" width="25.5" customWidth="true"/>
  <!-- Columns 2-4 (B-D) with style -->
  <col min="2" max="4" width="12" style="3" bestFit="true" customWidth="true"/>
  <!-- Column 5 (E) hidden -->
  <col min="5" max="5" width="0" hidden="true" customWidth="true"/>
</cols>
```

Column `min` and `max` are 1-based column indices.

---

## 8. Cell Data Model

### Cell Element (`<c>`)

```xml
<c r="B5"       <!-- Cell reference (e.g., "A1", "XFD1048576") -->
   t="s"        <!-- Data type -->
   s="4"        <!-- Style index (0-based into cellXfs) -->
   cm="0"       <!-- Cell metadata index -->
   vm="0"       <!-- Value metadata index -->
   ph="false">  <!-- Show phonetic -->
  <f>...</f>    <!-- Formula (optional) -->
  <v>...</v>    <!-- Value -->
  <is>...</is>  <!-- Inline string (when t="inlineStr") -->
</c>
```

### Cell Reference Format

Cell references use column letters (A-XFD) followed by a 1-based row number (1-1048576).

Column letter conversion:
- A=1, B=2, ... Z=26, AA=27, AB=28, ... XFD=16384
- Maximum: column XFD (16,384), row 1,048,576

### Cell Types (`t` attribute)

| Value | Type | Value Storage | Description |
|-------|------|--------------|-------------|
| `n` | Number | `<v>` contains numeric value | **Default** when `t` is absent. Includes integers, floats, dates, times |
| `s` | Shared String | `<v>` contains 0-based index into shared string table | String stored in `sharedStrings.xml` |
| `inlineStr` | Inline String | `<is>` contains string directly | String stored in the cell itself (not shared) |
| `str` | Formula String | `<v>` contains the cached string result | Result of a formula that returns a string |
| `b` | Boolean | `<v>` contains `0` (false) or `1` (true) | Boolean value |
| `e` | Error | `<v>` contains error code string | Cell contains an error value |

### Error Values

When `t="e"`, the `<v>` element contains one of:

| Error Code | Meaning |
|-----------|---------|
| `#NULL!` | Intersection of two ranges that do not intersect |
| `#DIV/0!` | Division by zero |
| `#VALUE!` | Wrong type of operand or argument |
| `#REF!` | Invalid cell reference |
| `#NAME?` | Unrecognized formula name or text |
| `#NUM!` | Invalid numeric value |
| `#N/A` | Value not available |
| `#GETTING_DATA` | External data retrieval in progress |

### Cell Examples

```xml
<!-- Numeric value -->
<c r="A1">
  <v>42.5</v>
</c>

<!-- Shared string (index 3 in shared string table) -->
<c r="A2" t="s">
  <v>3</v>
</c>

<!-- Inline string -->
<c r="A3" t="inlineStr">
  <is>
    <t>Hello World</t>
  </is>
</c>

<!-- Boolean (TRUE) -->
<c r="A4" t="b">
  <v>1</v>
</c>

<!-- Error -->
<c r="A5" t="e">
  <v>#DIV/0!</v>
</c>

<!-- Formula with cached numeric result -->
<c r="A6">
  <f>SUM(B1:B5)</f>
  <v>150</v>
</c>

<!-- Formula with cached string result -->
<c r="A7" t="str">
  <f>CONCATENATE(B1," ",C1)</f>
  <v>Hello World</v>
</c>

<!-- Date stored as serial number with date format style -->
<c r="A8" s="14">
  <v>44927</v>
</c>

<!-- Styled numeric cell -->
<c r="A9" s="5">
  <v>1234.56</v>
</c>

<!-- Empty cell with style only -->
<c r="A10" s="2"/>
```

### Date and Time Values

Dates and times are stored as **serial numbers** (type `n`), not as dedicated date types. The data type is determined by examining the number format applied to the cell through its style.

**Serial Date Encoding (1900 system, default):**
- The integer part represents days since **December 30, 1899** (serial 1 = January 1, 1900)
- The fractional part represents time as a fraction of 24 hours
- 0.0 = midnight (00:00:00)
- 0.25 = 6:00 AM
- 0.5 = noon (12:00:00)
- 0.75 = 6:00 PM

**Examples:**
- `1` = January 1, 1900
- `44927` = January 1, 2023
- `44927.5` = January 1, 2023 at 12:00:00 PM
- `0.75` = 6:00:00 PM (time only)

**Lotus 1-2-3 Bug:** Serial date 60 maps to the non-existent February 29, 1900. Serial dates <= 59 are off by one day. Implementers must handle this quirk.

**Detecting Date Cells:** A cell is a date if its number format (from the style index) contains date/time tokens. See the [Number Format section](#built-in-number-format-ids) for how to identify date formats.

---

## 9. Shared String Table

The shared string table (`xl/sharedStrings.xml`) deduplicates string values across the entire workbook. A package contains **at most one** shared string table.

### Structure

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
     count="10"           <!-- Total string references in workbook -->
     uniqueCount="7">     <!-- Number of unique strings in table -->

  <!-- Index 0: Plain string -->
  <si>
    <t>Name</t>
  </si>

  <!-- Index 1: Plain string with whitespace preservation -->
  <si>
    <t xml:space="preserve"> Leading space</t>
  </si>

  <!-- Index 2: Rich text string (character-level formatting) -->
  <si>
    <r>
      <rPr>
        <b/>                      <!-- Bold -->
        <sz val="11"/>            <!-- Font size -->
        <color rgb="FFFF0000"/>   <!-- Red color -->
        <rFont val="Calibri"/>    <!-- Font name -->
        <family val="2"/>         <!-- Font family -->
        <scheme val="minor"/>     <!-- Theme font scheme -->
      </rPr>
      <t>Bold Red</t>
    </r>
    <r>
      <rPr>
        <sz val="11"/>
        <color theme="1"/>
        <rFont val="Calibri"/>
        <family val="2"/>
        <scheme val="minor"/>
      </rPr>
      <t xml:space="preserve"> Normal</t>
    </r>
  </si>

  <!-- Index 3-6: More strings... -->
  <si><t>Revenue</t></si>
  <si><t>Cost</t></si>
  <si><t>Profit</t></si>
  <si><t>Total</t></si>
</sst>
```

### SST Attributes

| Attribute | Type | Description |
|-----------|------|-------------|
| `count` | unsignedInt | Total number of string references in all cells across the workbook |
| `uniqueCount` | unsignedInt | Number of `<si>` elements (unique strings) in the table |

### String Item (`<si>`) Children

Each `<si>` contains either:

1. **Plain text**: A single `<t>` element
2. **Rich text**: One or more `<r>` (run) elements, each containing:
   - `<rPr>` (run properties): Font formatting for this run
   - `<t>` (text): The text content of this run
3. **Phonetic text** (optional): `<rPh>` elements for phonetic readings (e.g., Japanese furigana)
4. **Phonetic properties** (optional): `<phoneticPr>` for phonetic settings

### Run Properties (`<rPr>`) Elements

| Element | Description | Example |
|---------|-------------|---------|
| `<b/>` | Bold | `<b/>` or `<b val="true"/>` |
| `<i/>` | Italic | `<i/>` |
| `<u/>` | Underline | `<u/>` or `<u val="double"/>` |
| `<strike/>` | Strikethrough | `<strike/>` |
| `<vertAlign>` | Superscript/subscript | `<vertAlign val="superscript"/>` |
| `<sz>` | Font size in points | `<sz val="11"/>` |
| `<color>` | Text color | `<color rgb="FF0000FF"/>` |
| `<rFont>` | Font name | `<rFont val="Arial"/>` |
| `<family>` | Font family | `<family val="2"/>` |
| `<charset>` | Character set | `<charset val="1"/>` |
| `<scheme>` | Theme font scheme | `<scheme val="minor"/>` |
| `<outline/>` | Outline | `<outline/>` |
| `<shadow/>` | Shadow | `<shadow/>` |
| `<condense/>` | Condensed | `<condense/>` |
| `<extend/>` | Extended | `<extend/>` |

### Cell Reference to Shared String

A cell references a shared string via a zero-based index:

```xml
<!-- Cell contains the string at index 3 ("Revenue") -->
<c r="A1" t="s">
  <v>3</v>
</c>
```

### Inline Strings vs. Shared Strings

Inline strings (`t="inlineStr"`) store the text directly in the cell using `<is>`:

```xml
<c r="A1" t="inlineStr">
  <is>
    <t>Direct text</t>
  </is>
</c>
```

**When to use each:**
- **Shared strings**: Best for data with repeated values. Excel always uses shared strings.
- **Inline strings**: Simpler for programmatic generation with few/no repeated values. Avoids maintaining a separate string table.

Both are valid. A workbook can mix both approaches.

---

## 10. Styles

All styling is stored centrally in `xl/styles.xml`. Cells reference styles by a zero-based index into the `<cellXfs>` array.

### Complete Styles Structure

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            mc:Ignorable="x14ac x16r2"
            xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">

  <!-- 1. Number Formats (custom only; built-in are implicit) -->
  <numFmts count="2">
    <numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>
    <numFmt numFmtId="165" formatCode="#,##0.00\ &quot;USD&quot;"/>
  </numFmts>

  <!-- 2. Fonts -->
  <fonts count="3" x14ac:knownFonts="1">
    <!-- Index 0: Default font -->
    <font>
      <sz val="11"/>
      <color theme="1"/>
      <name val="Calibri"/>
      <family val="2"/>
      <scheme val="minor"/>
    </font>
    <!-- Index 1: Bold -->
    <font>
      <b/>
      <sz val="11"/>
      <color theme="1"/>
      <name val="Calibri"/>
      <family val="2"/>
      <scheme val="minor"/>
    </font>
    <!-- Index 2: Red italic -->
    <font>
      <i/>
      <sz val="11"/>
      <color rgb="FFFF0000"/>
      <name val="Calibri"/>
      <family val="2"/>
    </font>
  </fonts>

  <!-- 3. Fills -->
  <fills count="3">
    <!-- Index 0: No fill (REQUIRED) -->
    <fill>
      <patternFill patternType="none"/>
    </fill>
    <!-- Index 1: Gray 125 (REQUIRED) -->
    <fill>
      <patternFill patternType="gray125"/>
    </fill>
    <!-- Index 2: Custom solid fill -->
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FF92D050"/>   <!-- Foreground (fill) color -->
        <bgColor indexed="64"/>     <!-- Background color -->
      </patternFill>
    </fill>
  </fills>

  <!-- 4. Borders -->
  <borders count="2">
    <!-- Index 0: No borders -->
    <border>
      <left/><right/><top/><bottom/><diagonal/>
    </border>
    <!-- Index 1: Thin borders all around -->
    <border>
      <left style="thin"><color indexed="64"/></left>
      <right style="thin"><color indexed="64"/></right>
      <top style="thin"><color indexed="64"/></top>
      <bottom style="thin"><color indexed="64"/></bottom>
      <diagonal/>
    </border>
  </borders>

  <!-- 5. Cell Style Formats (master formatting records) -->
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>

  <!-- 6. Cell Formats (the main style array; cells reference this by index) -->
  <cellXfs count="5">
    <!-- Index 0: Default (Normal) -->
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>

    <!-- Index 1: Bold text -->
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="true"/>

    <!-- Index 2: Date format -->
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="true"/>

    <!-- Index 3: Currency with green fill and borders -->
    <xf numFmtId="165" fontId="0" fillId="2" borderId="1" xfId="0"
        applyNumberFormat="true" applyFill="true" applyBorder="true"/>

    <!-- Index 4: With alignment -->
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="true">
      <alignment horizontal="center" vertical="center" wrapText="true"/>
    </xf>
  </cellXfs>

  <!-- 7. Cell Styles (named styles like "Normal", "Heading 1") -->
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>

  <!-- 8. Differential Formats (for conditional formatting and tables) -->
  <dxfs count="1">
    <dxf>
      <font><color rgb="FF9C0006"/></font>
      <fill>
        <patternFill>
          <bgColor rgb="FFFFC7CE"/>
        </patternFill>
      </fill>
    </dxf>
  </dxfs>

  <!-- 9. Table Styles -->
  <tableStyles count="0" defaultTableStyle="TableStyleMedium2"
               defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>
```

### Style Resolution Chain

To determine the complete formatting of a cell:

1. Read the cell's `s` attribute (style index). Default is `0`.
2. Look up `cellXfs[s]` -- the `<xf>` element at that index.
3. The `<xf>` element has index references into each sub-collection:
   - `numFmtId` -- number format ID (built-in or custom from `<numFmts>`)
   - `fontId` -- index into `<fonts>`
   - `fillId` -- index into `<fills>`
   - `borderId` -- index into `<borders>`
   - `xfId` -- index into `<cellStyleXfs>` (parent named style)
4. `apply*` boolean attributes indicate which formatting categories are applied:
   - `applyNumberFormat`, `applyFont`, `applyFill`, `applyBorder`, `applyAlignment`, `applyProtection`
   - When `false` or absent, that aspect inherits from the parent `xfId` style.

### Built-in Number Format IDs

Number formats with IDs 0-163 are built-in. Custom formats start at ID 164. Built-in formats do not appear in `<numFmts>` -- they are implied.

| ID | Format Code | Category |
|----|-------------|----------|
| 0 | `General` | General |
| 1 | `0` | Number |
| 2 | `0.00` | Number |
| 3 | `#,##0` | Number |
| 4 | `#,##0.00` | Number |
| 5 | `$#,##0_);($#,##0)` | Currency |
| 6 | `$#,##0_);[Red]($#,##0)` | Currency |
| 7 | `$#,##0.00_);($#,##0.00)` | Currency |
| 8 | `$#,##0.00_);[Red]($#,##0.00)` | Currency |
| 9 | `0%` | Percentage |
| 10 | `0.00%` | Percentage |
| 11 | `0.00E+00` | Scientific |
| 12 | `# ?/?` | Fraction |
| 13 | `# ??/??` | Fraction |
| 14 | `m/d/yyyy` | **Date** |
| 15 | `d-mmm-yy` | **Date** |
| 16 | `d-mmm` | **Date** |
| 17 | `mmm-yy` | **Date** |
| 18 | `h:mm AM/PM` | **Time** |
| 19 | `h:mm:ss AM/PM` | **Time** |
| 20 | `h:mm` | **Time** |
| 21 | `h:mm:ss` | **Time** |
| 22 | `m/d/yyyy h:mm` | **Date+Time** |
| 37 | `#,##0_);(#,##0)` | Number |
| 38 | `#,##0_);[Red](#,##0)` | Number |
| 39 | `#,##0.00_);(#,##0.00)` | Number |
| 40 | `#,##0.00_);[Red](#,##0.00)` | Number |
| 41 | `_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)` | Accounting |
| 42 | `_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)` | Accounting |
| 43 | `_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)` | Accounting |
| 44 | `_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)` | Accounting |
| 45 | `mm:ss` | **Time** |
| 46 | `[h]:mm:ss` | **Time** (elapsed) |
| 47 | `mm:ss.0` | **Time** |
| 48 | `##0.0E+0` | Scientific |
| 49 | `@` | Text |

**IDs 23-36** and **50-81** are locale-specific date/time formats (CJK, Thai, etc.).

**Note on ID 14:** The standard defines it as `mm-dd-yy` but Excel implements it as `m/d/yyyy`. Use the Excel behavior for compatibility.

### Detecting Date/Time Formats

To determine if a number format represents a date or time, parse the format code string:

**Date indicators:** Contains `y` (year), `d` (day), `w` (weekday), or `q` (quarter) outside of quoted strings and bracketed sections.

**Time indicators:** Contains `h` (hour), `s` (second), `AM/PM`, or `A/P` outside of quoted strings and bracketed sections.

**Text format:** Contains `@`.

**Algorithm:**
1. Strip quoted strings (`"..."`) and escaped characters (`\x`)
2. Strip bracketed sections (`[...]`) -- these are colors or conditions
3. Check remaining text for date/time tokens
4. If date/time tokens found, the cell contains a date/time value

### Font Element

```xml
<font>
  <b/>                            <!-- Bold (val="true" implied) -->
  <i/>                            <!-- Italic -->
  <u val="single"/>               <!-- Underline: single, double, singleAccounting, doubleAccounting -->
  <strike/>                       <!-- Strikethrough -->
  <vertAlign val="superscript"/>  <!-- superscript | subscript | baseline -->
  <sz val="11"/>                  <!-- Size in points -->
  <color rgb="FF000000"/>         <!-- Color (see color model below) -->
  <name val="Calibri"/>           <!-- Font name -->
  <family val="2"/>               <!-- 0=N/A, 1=Roman, 2=Swiss, 3=Modern, 4=Script, 5=Decorative -->
  <charset val="1"/>              <!-- Character set -->
  <scheme val="minor"/>           <!-- none | major | minor (theme font) -->
  <outline/>                      <!-- Outline -->
  <shadow/>                       <!-- Shadow -->
  <condense/>                     <!-- Condensed -->
  <extend/>                       <!-- Extended -->
</font>
```

### Color Model

Colors can be specified in four ways:

```xml
<!-- 1. ARGB hex (most common) -->
<color rgb="FFFF0000"/>           <!-- Alpha + RGB, e.g., fully opaque red -->

<!-- 2. Theme color reference -->
<color theme="1"/>                <!-- Index into theme color scheme -->
<color theme="4" tint="0.39997"/> <!-- Theme color with lightness adjustment (-1.0 to 1.0) -->

<!-- 3. Indexed color (legacy palette) -->
<color indexed="10"/>             <!-- Index into default color palette (0-63) -->

<!-- 4. Auto color -->
<color auto="true"/>              <!-- System window text or background color -->
```

**Theme color indices:**
| Index | Meaning |
|-------|---------|
| 0 | Light 1 (typically White) |
| 1 | Dark 1 (typically Black) |
| 2 | Light 2 |
| 3 | Dark 2 |
| 4 | Accent 1 |
| 5 | Accent 2 |
| 6 | Accent 3 |
| 7 | Accent 4 |
| 8 | Accent 5 |
| 9 | Accent 6 |
| 10 | Hyperlink |
| 11 | Followed Hyperlink |

### Fill Element

```xml
<!-- Pattern fill -->
<fill>
  <patternFill patternType="solid">
    <fgColor rgb="FFFFFF00"/>     <!-- Pattern foreground color -->
    <bgColor rgb="FF000000"/>     <!-- Pattern background color -->
  </patternFill>
</fill>

<!-- Gradient fill -->
<fill>
  <gradientFill type="linear" degree="90">
    <stop position="0"><color rgb="FFFF0000"/></stop>
    <stop position="1"><color rgb="FF0000FF"/></stop>
  </gradientFill>
</fill>
```

**Pattern types:** `none`, `solid`, `mediumGray`, `darkGray`, `lightGray`, `darkHorizontal`, `darkVertical`, `darkDown`, `darkUp`, `darkGrid`, `darkTrellis`, `lightHorizontal`, `lightVertical`, `lightDown`, `lightUp`, `lightGrid`, `lightTrellis`, `gray125`, `gray0625`

**Important:** For `solid` fills, the actual fill color is in `fgColor`. The `bgColor` is irrelevant for solid patterns. The first two fills (index 0 = `none`, index 1 = `gray125`) are **required** and must always be present.

### Border Element

```xml
<border diagonalUp="false" diagonalDown="false">
  <left style="thin">
    <color rgb="FF000000"/>
  </left>
  <right style="medium">
    <color theme="1"/>
  </right>
  <top style="double">
    <color rgb="FFFF0000"/>
  </top>
  <bottom style="thick">
    <color indexed="64"/>
  </bottom>
  <diagonal style="dashDot">
    <color rgb="FF0000FF"/>
  </diagonal>
</border>
```

**Border styles:** `none`, `thin`, `medium`, `dashed`, `dotted`, `thick`, `double`, `hair`, `mediumDashed`, `dashDot`, `mediumDashDot`, `dashDotDot`, `mediumDashDotDot`, `slantDashDot`

### Alignment

```xml
<alignment horizontal="center"    <!-- general|left|center|right|fill|justify|centerContinuous|distributed -->
           vertical="center"      <!-- top|center|bottom|justify|distributed -->
           textRotation="45"      <!-- 0-180 degrees, or 255 for vertical text -->
           wrapText="true"        <!-- Enable text wrapping -->
           indent="2"             <!-- Indentation level -->
           shrinkToFit="false"    <!-- Shrink text to fit cell width -->
           readingOrder="0"/>     <!-- 0=context, 1=left-to-right, 2=right-to-left -->
```

### Protection

```xml
<protection locked="true"         <!-- Cell is locked when sheet is protected -->
            hidden="false"/>      <!-- Formula is hidden when sheet is protected -->
```

### Differential Formats (`<dxfs>`)

Differential formats are partial style records used by **conditional formatting** and **table styles**. Unlike `cellXfs` which are complete style records, a `dxf` only specifies the differences from the base style:

```xml
<dxf>
  <font>
    <b/>
    <color rgb="FF9C0006"/>
  </font>
  <fill>
    <patternFill>
      <bgColor rgb="FFFFC7CE"/>
    </patternFill>
  </fill>
  <border>
    <left style="thin"><color auto="true"/></left>
  </border>
</dxf>
```

---

## 11. Formulas

### Basic Formula

```xml
<c r="A6">
  <f>SUM(A1:A5)</f>
  <v>15</v>
</c>
```

The `<f>` element contains the formula text. The `<v>` element contains the **cached result** from the last calculation. A reader may use the cached value or recalculate.

### Formula Element Attributes

| Attribute | Type | Description |
|-----------|------|-------------|
| `t` | enum | Formula type: `normal` (default), `shared`, `array`, `dataTable` |
| `ref` | string | Range the formula applies to (required for shared/array/dataTable master cells) |
| `si` | unsignedInt | Shared formula group index |
| `ca` | boolean | Cell needs recalculation (volatile function, circular reference) |
| `aca` | boolean | Always calculate array (legacy) |
| `r1` | string | First input cell for data table |
| `r2` | string | Second input cell for data table |
| `dt2D` | boolean | True if two-dimensional data table |
| `dtr` | boolean | True if one-dimensional row data table |
| `del1` | boolean | First input cell deleted |
| `del2` | boolean | Second input cell deleted |
| `bx` | boolean | Formula assigns to a name |

### Normal Formula

```xml
<c r="C1">
  <f>A1+B1</f>
  <v>42</v>
</c>
```

### Shared Formula

Shared formulas optimize storage when the same formula pattern is applied across a range. Only the master cell stores the formula text; dependent cells reference it by group index.

```xml
<!-- Master cell: defines the shared formula -->
<c r="C1">
  <f t="shared" ref="C1:C100" si="0">A1*B1</f>
  <v>10</v>
</c>

<!-- Dependent cells: reference the shared group -->
<c r="C2">
  <f t="shared" si="0"/>
  <v>20</v>
</c>
<c r="C3">
  <f t="shared" si="0"/>
  <v>30</v>
</c>
```

The `ref` attribute on the master cell defines the range. The `si` attribute is the shared group identifier (unique within the sheet). Dependent cells omit `ref` and the formula text but include the `si` to identify the group. The actual formula for each dependent cell is derived by adjusting relative references from the master formula.

### Array Formula

Array formulas compute over a range and can return multiple values:

```xml
<!-- Array formula occupying C1:C5 -->
<c r="C1">
  <f t="array" ref="C1:C5">{A1:A5*B1:B5}</f>
  <v>10</v>
</c>
<c r="C2">
  <v>20</v>
</c>
<!-- C3, C4, C5 also just have <v> elements -->
```

Only the top-left cell of the array range contains the formula with `t="array"` and `ref`. Other cells in the range contain only their cached values. The formula text for array formulas entered with Ctrl+Shift+Enter is typically enclosed in braces `{}` in the XML.

### Data Table Formula

Data tables are what-if analysis tools:

```xml
<!-- One-input column data table -->
<c r="B2">
  <f t="dataTable" ref="B2:B10" r1="A1" dt2D="false" dtr="false"/>
  <v>100</v>
</c>

<!-- Two-input data table -->
<c r="B2">
  <f t="dataTable" ref="B2:F10" r1="A1" r2="B1" dt2D="true"/>
  <v>200</v>
</c>
```

| Attribute | Description |
|-----------|-------------|
| `r1` | Row input cell (one-input column table) or row input cell (two-input table) |
| `r2` | Column input cell (two-input table only) |
| `dt2D` | `true` for two-dimensional data table |
| `dtr` | `true` for one-dimensional row-oriented data table |

### Calculation Chain (`xl/calcChain.xml`)

Records the order cells were last calculated:

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <c r="C1" i="1"/>     <!-- i = sheet index -->
  <c r="C2" i="1"/>
  <c r="A6" i="1" l="1"/>  <!-- l="1" marks start of new dependency level -->
</calcChain>
```

**Implementation note:** The calculation chain is optional. When generating files, it can be omitted entirely -- Excel will rebuild it on first open.

---

## 12. Defined Names

Defined names are stored in the workbook's `<definedNames>` element.

### Structure

```xml
<definedNames>
  <!-- Global name (accessible from all sheets) -->
  <definedName name="SalesData">Sheet1!$A$1:$D$100</definedName>

  <!-- Sheet-scoped name (localSheetId = 0-based sheet index) -->
  <definedName name="Total" localSheetId="0">Sheet1!$E$101</definedName>

  <!-- Constant value -->
  <definedName name="TaxRate">0.08</definedName>

  <!-- Formula -->
  <definedName name="GrossProfit">Sheet1!$C$2-Sheet1!$D$2</definedName>

  <!-- Hidden name -->
  <definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="true">
    Sheet1!$A$1:$G$50
  </definedName>
</definedNames>
```

### Defined Name Attributes

| Attribute | Type | Description |
|-----------|------|-------------|
| `name` | string | **Required.** The name (case-insensitive, max 255 chars) |
| `localSheetId` | unsignedInt | Optional. 0-based sheet index for sheet-scoped names. Absent = global. |
| `hidden` | boolean | Optional. Default `false`. If `true`, name is hidden from UI. |
| `comment` | string | Optional. Comment for the defined name. |
| `function` | boolean | Optional. True if this is a user-defined function. |
| `vbProcedure` | boolean | Optional. True if this is a VBA procedure. |

### Reserved Built-in Names

| Name | Purpose |
|------|---------|
| `_xlnm.Print_Area` | Print area for a sheet |
| `_xlnm.Print_Titles` | Rows/columns to repeat on each page |
| `_xlnm._FilterDatabase` | AutoFilter range |
| `_xlnm.Criteria` | Advanced filter criteria range |
| `_xlnm.Extract` | Advanced filter output range |
| `_xlnm.Database` | Database range |
| `_xlnm.Sheet_Title` | Sheet title for headers/footers |

### Examples

```xml
<!-- Print area for Sheet1 -->
<definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$A$1:$G$50</definedName>

<!-- Print titles: repeat rows 1-2 on every page for Sheet1 -->
<definedName name="_xlnm.Print_Titles" localSheetId="0">Sheet1!$1:$2</definedName>

<!-- Print titles: repeat column A on every page -->
<definedName name="_xlnm.Print_Titles" localSheetId="1">Sheet2!$A:$A</definedName>

<!-- Multi-range reference -->
<definedName name="AllData">Sheet1!$A$1:$D$50,Sheet2!$A$1:$D$50</definedName>
```

---

## 13. Tables

Tables are stored as separate parts in `xl/tables/tableN.xml`, referenced from the worksheet.

### Table Part Structure

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
       id="1"                          <!-- Unique numeric ID across workbook -->
       name="Table1"                   <!-- Internal name (unique) -->
       displayName="SalesTable"        <!-- Display name (unique, used in structured refs) -->
       ref="A1:D10"                    <!-- Cell range including headers -->
       totalsRowCount="1"              <!-- Number of total rows (0 or 1) -->
       totalsRowShown="true"
       headerRowCount="1">             <!-- Number of header rows (default: 1) -->

  <autoFilter ref="A1:D10"/>

  <sortState ref="A2:D9">
    <sortCondition ref="B2:B9"/>       <!-- Sort by column B ascending -->
  </sortState>

  <tableColumns count="4">
    <tableColumn id="1" name="Name"
                 dataDxfId="1"/>
    <tableColumn id="2" name="Region"/>
    <tableColumn id="3" name="Revenue"
                 totalsRowFunction="sum"
                 dataDxfId="2"/>
    <tableColumn id="4" name="Cost"
                 totalsRowFunction="sum">
      <calculatedColumnFormula>[@Revenue]*0.6</calculatedColumnFormula>
    </tableColumn>
  </tableColumns>

  <tableStyleInfo name="TableStyleMedium2"
                  showFirstColumn="false"
                  showLastColumn="false"
                  showRowStripes="true"
                  showColumnStripes="false"/>
</table>
```

### Table Attributes

| Attribute | Description |
|-----------|-------------|
| `id` | Unique numeric ID (must be unique across all tables in the workbook) |
| `name` | Internal name (unique across tables and defined names) |
| `displayName` | Display name used in structured references like `SalesTable[Revenue]` |
| `ref` | Cell range of the table (including headers and totals) |
| `headerRowCount` | Number of header rows (default 1; set to 0 for headerless tables) |
| `totalsRowCount` | Number of total rows (0 = no totals row; 1 = totals row present) |
| `totalsRowShown` | Whether totals row is visible |

### Table Column Attributes

| Attribute | Description |
|-----------|-------------|
| `id` | Unique column ID within the table |
| `name` | Column header text (must match the cell value in the header row) |
| `totalsRowFunction` | Aggregation: `average`, `count`, `countNums`, `max`, `min`, `stdDev`, `sum`, `var`, `custom` |
| `totalsRowFormula` | Custom formula for totals (when function = `custom`) |
| `dataDxfId` | Differential format ID for data cells in this column |
| `headerRowDxfId` | Differential format ID for the header cell |
| `totalsRowDxfId` | Differential format ID for the totals cell |

### Calculated Column Formula

```xml
<tableColumn id="4" name="Profit">
  <calculatedColumnFormula>[@Revenue]-[@Cost]</calculatedColumnFormula>
</tableColumn>
```

Structured references use `[@ColumnName]` to refer to the current row's value in the named column.

### Worksheet Reference

Tables are linked from the worksheet:

```xml
<!-- In sheet1.xml -->
<tableParts count="1">
  <tablePart r:id="rId2"/>
</tableParts>
```

---

## 14. Conditional Formatting

Conditional formatting rules appear after `<sheetData>` in the worksheet.

### Structure

```xml
<conditionalFormatting sqref="B2:B100">
  <!-- Rule 1: Highlight cells > 50000 -->
  <cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan">
    <formula>50000</formula>
  </cfRule>

  <!-- Rule 2: Highlight cells with specific text -->
  <cfRule type="containsText" dxfId="1" priority="2" operator="containsText" text="Error">
    <formula>NOT(ISERROR(SEARCH("Error",B2)))</formula>
  </cfRule>
</conditionalFormatting>

<!-- Color Scale -->
<conditionalFormatting sqref="C2:C100">
  <cfRule type="colorScale" priority="3">
    <colorScale>
      <cfvo type="min"/>
      <cfvo type="percentile" val="50"/>
      <cfvo type="max"/>
      <color rgb="FFF8696B"/>    <!-- Red for min -->
      <color rgb="FFFFEB84"/>    <!-- Yellow for middle -->
      <color rgb="FF63BE7B"/>    <!-- Green for max -->
    </colorScale>
  </cfRule>
</conditionalFormatting>

<!-- Data Bar -->
<conditionalFormatting sqref="D2:D100">
  <cfRule type="dataBar" priority="4">
    <dataBar>
      <cfvo type="min"/>
      <cfvo type="max"/>
      <color rgb="FF638EC6"/>
    </dataBar>
  </cfRule>
</conditionalFormatting>

<!-- Icon Set -->
<conditionalFormatting sqref="E2:E100">
  <cfRule type="iconSet" priority="5">
    <iconSet iconSet="3TrafficLights1">
      <cfvo type="percent" val="0"/>
      <cfvo type="percent" val="33"/>
      <cfvo type="percent" val="67"/>
    </iconSet>
  </cfRule>
</conditionalFormatting>

<!-- Formula-based rule -->
<conditionalFormatting sqref="A2:G100">
  <cfRule type="expression" dxfId="2" priority="6">
    <formula>$A2="Urgent"</formula>
  </cfRule>
</conditionalFormatting>
```

### cfRule Type Values (`ST_CfType`)

| Type | Description |
|------|-------------|
| `cellIs` | Compare cell value (uses `operator` attribute) |
| `expression` | Boolean formula determines formatting |
| `colorScale` | Gradient color scale |
| `dataBar` | Data bar visualization |
| `iconSet` | Icon set visualization |
| `top10` | Top N or bottom N values |
| `aboveAverage` | Above/below average |
| `containsText` | Cell contains text |
| `notContainsText` | Cell does not contain text |
| `beginsWith` | Cell begins with text |
| `endsWith` | Cell ends with text |
| `containsBlanks` | Cell is blank |
| `notContainsBlanks` | Cell is not blank |
| `containsErrors` | Cell contains an error |
| `notContainsErrors` | Cell does not contain an error |
| `timePeriod` | Date falls within a time period |
| `duplicateValues` | Duplicate values |
| `uniqueValues` | Unique values |

### Comparison Operators (for `cellIs` type)

`lessThan`, `lessThanOrEqual`, `equal`, `notEqual`, `greaterThanOrEqual`, `greaterThan`, `between`, `notBetween`

For `between` and `notBetween`, two `<formula>` elements are required.

### Conditional Value Object (`<cfvo>`)

| Type | Description |
|------|-------------|
| `num` | Literal numeric value |
| `percent` | Percentage |
| `max` | Maximum value in the range |
| `min` | Minimum value in the range |
| `percentile` | Percentile value |
| `formula` | Formula-derived value |

---

## 15. Charts

Charts are stored as separate parts in `xl/charts/chartN.xml` and connected through drawings.

### Chart Relationship Chain

```
Worksheet
  -> Drawing (xl/drawings/drawingN.xml)
    -> Chart (xl/charts/chartN.xml)
```

### Drawing Anchor

In `xl/drawings/drawing1.xml`:

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
           xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">

  <xdr:twoCellAnchor>
    <xdr:from>
      <xdr:col>4</xdr:col>          <!-- 0-based column -->
      <xdr:colOff>0</xdr:colOff>    <!-- Offset in EMUs -->
      <xdr:row>1</xdr:row>          <!-- 0-based row -->
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>12</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>16</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:to>

    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="2" name="Chart 1"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm>
        <a:off x="0" y="0"/>
        <a:ext cx="0" cy="0"/>
      </xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                   r:id="rId1"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>

    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>
```

### Drawing Anchor Types

| Anchor Type | Description |
|-------------|-------------|
| `<xdr:twoCellAnchor>` | Anchored between two cells (moves and resizes with cells) |
| `<xdr:oneCellAnchor>` | Anchored to one cell with explicit size (moves but does not resize) |
| `<xdr:absoluteAnchor>` | Absolute position (does not move or resize) |

**EMU (English Metric Unit):** 1 inch = 914400 EMUs. Used for precise positioning.

### Chart Structure

In `xl/charts/chart1.xml`:

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">

  <c:chart>
    <c:title>
      <c:tx>
        <c:rich>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:t>Sales Report</a:t>
            </a:r>
          </a:p>
        </c:rich>
      </c:tx>
      <c:overlay val="0"/>
    </c:title>

    <c:autoTitleDeleted val="0"/>

    <c:plotArea>
      <c:layout/>

      <!-- Chart type element (one of many) -->
      <c:barChart>
        <c:barDir val="col"/>       <!-- col | bar -->
        <c:grouping val="clustered"/> <!-- clustered | stacked | percentStacked -->

        <c:ser>                      <!-- Data series -->
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx>                     <!-- Series name -->
            <c:strRef>
              <c:f>Sheet1!$B$1</c:f>
              <c:strCache>
                <c:ptCount val="1"/>
                <c:pt idx="0"><c:v>Revenue</c:v></c:pt>
              </c:strCache>
            </c:strRef>
          </c:tx>
          <c:cat>                    <!-- Categories (X axis) -->
            <c:strRef>
              <c:f>Sheet1!$A$2:$A$5</c:f>
              <c:strCache>
                <c:ptCount val="4"/>
                <c:pt idx="0"><c:v>Q1</c:v></c:pt>
                <c:pt idx="1"><c:v>Q2</c:v></c:pt>
                <c:pt idx="2"><c:v>Q3</c:v></c:pt>
                <c:pt idx="3"><c:v>Q4</c:v></c:pt>
              </c:strCache>
            </c:strRef>
          </c:cat>
          <c:val>                    <!-- Values (Y axis) -->
            <c:numRef>
              <c:f>Sheet1!$B$2:$B$5</c:f>
              <c:numCache>
                <c:formatCode>General</c:formatCode>
                <c:ptCount val="4"/>
                <c:pt idx="0"><c:v>10000</c:v></c:pt>
                <c:pt idx="1"><c:v>15000</c:v></c:pt>
                <c:pt idx="2"><c:v>12000</c:v></c:pt>
                <c:pt idx="3"><c:v>18000</c:v></c:pt>
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>

        <c:axId val="111111111"/>
        <c:axId val="222222222"/>
      </c:barChart>

      <!-- Axes -->
      <c:catAx>
        <c:axId val="111111111"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:crossAx val="222222222"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="222222222"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="l"/>
        <c:crossAx val="111111111"/>
      </c:valAx>
    </c:plotArea>

    <c:legend>
      <c:legendPos val="b"/>
      <c:overlay val="0"/>
    </c:legend>

    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>     <!-- gap | zero | span -->
  </c:chart>
</c:chartSpace>
```

### Chart Type Elements

| Element | Chart Type |
|---------|-----------|
| `<c:barChart>` | Bar/column chart |
| `<c:bar3DChart>` | 3D bar chart |
| `<c:lineChart>` | Line chart |
| `<c:line3DChart>` | 3D line chart |
| `<c:areaChart>` | Area chart |
| `<c:area3DChart>` | 3D area chart |
| `<c:pieChart>` | Pie chart |
| `<c:pie3DChart>` | 3D pie chart |
| `<c:doughnutChart>` | Doughnut chart |
| `<c:scatterChart>` | Scatter (XY) chart |
| `<c:bubbleChart>` | Bubble chart |
| `<c:radarChart>` | Radar chart |
| `<c:stockChart>` | Stock chart |
| `<c:surface3DChart>` | 3D surface chart |
| `<c:ofPieChart>` | Bar/pie of pie chart |

### Data Reference Types

| Element | Purpose |
|---------|---------|
| `<c:numRef>` | Reference to numeric data (contains `<c:f>` formula and `<c:numCache>`) |
| `<c:strRef>` | Reference to string data (contains `<c:f>` formula and `<c:strCache>`) |
| `<c:numLit>` | Literal numeric data (inline, no worksheet reference) |
| `<c:strLit>` | Literal string data (inline) |

---

## 16. Pivot Tables

Pivot tables involve three interconnected parts.

### Part Relationships

```
Workbook
  -> PivotCacheDefinition (xl/pivotCache/pivotCacheDefinition1.xml)
    -> PivotCacheRecords (xl/pivotCache/pivotCacheRecords1.xml)

Worksheet
  -> PivotTable (xl/pivotTables/pivotTable1.xml)
    -> PivotCacheDefinition (via r:id)
```

### Pivot Cache Definition

Defines the data source and field metadata:

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                      r:id="rId1"
                      refreshOnLoad="false"
                      recordCount="100">

  <cacheSource type="worksheet">
    <worksheetSource ref="A1:D101" sheet="Sheet1"/>
  </cacheSource>

  <cacheFields count="4">
    <cacheField name="Region" numFmtId="0">
      <sharedItems count="4">
        <s v="North"/>
        <s v="South"/>
        <s v="East"/>
        <s v="West"/>
      </sharedItems>
    </cacheField>
    <cacheField name="Product" numFmtId="0">
      <sharedItems count="3">
        <s v="Widget"/>
        <s v="Gadget"/>
        <s v="Doohickey"/>
      </sharedItems>
    </cacheField>
    <cacheField name="Revenue" numFmtId="0">
      <sharedItems containsSemiMixedTypes="false" containsString="false"
                   containsNumber="true" minValue="100" maxValue="50000"/>
    </cacheField>
    <cacheField name="Date" numFmtId="14">
      <sharedItems containsSemiMixedTypes="false" containsNonDate="false"
                   containsDate="true" containsString="false"
                   minDate="2023-01-01T00:00:00" maxDate="2023-12-31T00:00:00"/>
    </cacheField>
  </cacheFields>
</pivotCacheDefinition>
```

### Pivot Cache Records

Contains the actual cached data:

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                   count="100">
  <r>
    <x v="0"/>         <!-- Index into sharedItems of field 0 ("North") -->
    <x v="1"/>         <!-- Index into sharedItems of field 1 ("Gadget") -->
    <n v="5000"/>      <!-- Numeric value for field 2 -->
    <d v="2023-03-15T00:00:00"/>  <!-- Date value for field 3 -->
  </r>
  <r>
    <x v="2"/>
    <x v="0"/>
    <n v="3200"/>
    <d v="2023-06-22T00:00:00"/>
  </r>
  <!-- ... more records -->
</pivotCacheRecords>
```

### Cache Record Value Types

| Element | Type |
|---------|------|
| `<x>` | Shared item index |
| `<n>` | Number |
| `<s>` | String |
| `<b>` | Boolean |
| `<e>` | Error |
| `<d>` | Date (ISO 8601) |
| `<m/>` | Missing/empty |

### Pivot Table Definition

Defines the layout and configuration:

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                      name="PivotTable1"
                      cacheId="0"
                      dataOnRows="false"
                      applyNumberFormats="false"
                      applyBorderFormats="false"
                      applyFontFormats="false"
                      applyPatternFormats="false"
                      applyAlignmentFormats="false"
                      applyWidthHeightFormats="true"
                      dataCaption="Values"
                      updatedVersion="7"
                      minRefreshableVersion="3">

  <location ref="A3:C8" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>

  <pivotFields count="4">
    <!-- Field 0: Region on row axis -->
    <pivotField axis="axisRow" showAll="false">
      <items count="5">
        <item x="0"/><item x="1"/><item x="2"/><item x="3"/>
        <item t="default"/>   <!-- Default/total item -->
      </items>
    </pivotField>

    <!-- Field 1: Product (not on any axis) -->
    <pivotField showAll="false"/>

    <!-- Field 2: Revenue as data field -->
    <pivotField dataField="true" numFmtId="4" showAll="false"/>

    <!-- Field 3: Date (not used) -->
    <pivotField showAll="false"/>
  </pivotFields>

  <rowFields count="1">
    <field x="0"/>     <!-- Field index 0 (Region) -->
  </rowFields>

  <rowItems count="5">
    <i><x/></i>
    <i><x v="1"/></i>
    <i><x v="2"/></i>
    <i><x v="3"/></i>
    <i t="grand"><x/></i>
  </rowItems>

  <colItems count="1">
    <i/>
  </colItems>

  <dataFields count="1">
    <dataField name="Sum of Revenue" fld="2" subtotal="sum" baseField="0" baseItem="0"/>
  </dataFields>

  <pivotTableStyleInfo name="PivotStyleLight16"
                       showRowHeaders="true"
                       showColHeaders="true"
                       showRowStripes="false"
                       showColStripes="false"
                       showLastColumn="true"/>
</pivotTableDefinition>
```

### Pivot Field Axis Values

| Value | Description |
|-------|-------------|
| `axisRow` | Row labels area |
| `axisCol` | Column labels area |
| `axisPage` | Report filter area |
| `axisValues` | Values area |

### Data Field Subtotal Functions

`sum`, `count`, `average`, `max`, `min`, `product`, `countNums`, `stdDev`, `stdDevp`, `var`, `varp`

---

## 17. Additional Worksheet Features

### Merged Cells

```xml
<mergeCells count="2">
  <mergeCell ref="A1:C1"/>    <!-- Cells A1, B1, C1 merged -->
  <mergeCell ref="A5:A10"/>   <!-- Vertical merge -->
</mergeCells>
```

Only the top-left cell of a merged range should contain data. Other cells in the range should be empty.

### Hyperlinks

```xml
<hyperlinks>
  <!-- External URL (via relationship) -->
  <hyperlink ref="A1" r:id="rId1" tooltip="Visit website"/>

  <!-- Internal cell reference -->
  <hyperlink ref="A2" location="Sheet2!B5" display="Go to Sheet2 B5"/>

  <!-- Internal with defined name -->
  <hyperlink ref="A3" location="SalesData" display="Sales Data"/>
</hyperlinks>
```

External hyperlinks require a relationship entry with `TargetMode="External"`.

### Data Validations

```xml
<dataValidations count="3">
  <!-- Drop-down list from explicit values -->
  <dataValidation type="list" allowBlank="true"
                  showInputMessage="true" showErrorMessage="true"
                  sqref="B2:B100">
    <formula1>"Option A,Option B,Option C"</formula1>
  </dataValidation>

  <!-- Integer between 1 and 100 -->
  <dataValidation type="whole" operator="between"
                  allowBlank="true" sqref="C2:C100">
    <formula1>1</formula1>
    <formula2>100</formula2>
  </dataValidation>

  <!-- List from cell range -->
  <dataValidation type="list" allowBlank="true" sqref="D2:D100">
    <formula1>Sheet2!$A$1:$A$20</formula1>
  </dataValidation>
</dataValidations>
```

**Validation types:** `none`, `whole`, `decimal`, `list`, `date`, `time`, `textLength`, `custom`

**Operators:** `between`, `notBetween`, `equal`, `notEqual`, `lessThan`, `lessThanOrEqual`, `greaterThan`, `greaterThanOrEqual`

### Auto Filters

```xml
<autoFilter ref="A1:G50">
  <filterColumn colId="0">       <!-- 0-based column within the range -->
    <filters>
      <filter val="North"/>
      <filter val="South"/>
    </filters>
  </filterColumn>
  <filterColumn colId="2">
    <customFilters>
      <customFilter operator="greaterThanOrEqual" val="1000"/>
    </customFilters>
  </filterColumn>
</autoFilter>
```

### Sheet Protection

```xml
<sheetProtection sheet="true"
                 password="CC35"          <!-- Hashed password (legacy, weak) -->
                 algorithmName="SHA-512"  <!-- Modern hash algorithm -->
                 hashValue="..."          <!-- Base64 hash -->
                 saltValue="..."          <!-- Base64 salt -->
                 spinCount="100000"       <!-- Iteration count -->
                 objects="true"           <!-- Protect objects -->
                 scenarios="true"         <!-- Protect scenarios -->
                 formatCells="false"      <!-- Allow formatting cells -->
                 formatColumns="false"
                 formatRows="false"
                 insertColumns="false"
                 insertRows="false"
                 insertHyperlinks="false"
                 deleteColumns="false"
                 deleteRows="false"
                 selectLockedCells="false"
                 sort="false"
                 autoFilter="false"
                 pivotTables="false"
                 selectUnlockedCells="false"/>
```

### Freeze Panes

```xml
<sheetViews>
  <sheetView workbookViewId="0">
    <!-- Freeze at row 2, column B (rows 1 frozen, column A frozen) -->
    <pane xSplit="1" ySplit="1"
          topLeftCell="B2"
          activePane="bottomRight"
          state="frozen"/>
    <selection pane="bottomRight" activeCell="B2" sqref="B2"/>
  </sheetView>
</sheetViews>
```

**Pane states:** `frozen`, `frozenSplit`, `split`

### Page Setup

```xml
<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75"
             header="0.3" footer="0.3"/>

<pageSetup paperSize="1"              <!-- 1=Letter, 9=A4 -->
           scale="100"                 <!-- Print scale percentage -->
           fitToWidth="1"              <!-- Fit to N pages wide -->
           fitToHeight="1"             <!-- Fit to N pages tall -->
           orientation="landscape"     <!-- portrait | landscape -->
           horizontalDpi="600"
           verticalDpi="600"
           r:id="rId4"/>               <!-- Printer settings relationship -->
```

### Comments (Legacy)

Legacy comments (pre-threaded) use a separate comments part and VML drawing:

```xml
<!-- xl/comments1.xml -->
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors>
    <author>John Doe</author>
  </authors>
  <commentList>
    <comment ref="A1" authorId="0" shapeId="0">
      <text>
        <r>
          <rPr><b/><sz val="9"/><rFont val="Tahoma"/></rPr>
          <t>John Doe:</t>
        </r>
        <r>
          <rPr><sz val="9"/><rFont val="Tahoma"/></rPr>
          <t xml:space="preserve">
This is a comment.</t>
        </r>
      </text>
    </comment>
  </commentList>
</comments>
```

---

## 18. Implementation Notes

### Parser Implementation Order

For a minimal viable XLSX reader:

1. **Unzip** the package
2. **Parse `[Content_Types].xml`** to locate workbook
3. **Parse `_rels/.rels`** to find the workbook relationship
4. **Parse `xl/_rels/workbook.xml.rels`** to locate sheets, shared strings, styles
5. **Load shared strings** from `xl/sharedStrings.xml` into an indexed array
6. **Load styles** from `xl/styles.xml` (at minimum: `numFmts` and `cellXfs` for date detection)
7. **Parse workbook** (`xl/workbook.xml`) to get sheet names and relationship IDs
8. **Parse each worksheet** using sheet relationship IDs to locate XML files
9. **Resolve cell values**: dereference shared strings, detect dates via style

### Writer Implementation Order

For a minimal viable XLSX writer:

1. **Build shared string table** by collecting all unique strings
2. **Build styles** (at minimum: default font, required fills, empty border, one cellXf)
3. **Write worksheet XML** for each sheet (rows, cells with string indices and style indices)
4. **Write workbook XML** with sheet references
5. **Write shared strings XML**
6. **Write styles XML**
7. **Write `[Content_Types].xml`** with all part types
8. **Write relationship files** (package-level and workbook-level)
9. **ZIP everything** into the output file

### Common Pitfalls

1. **Date detection**: Dates are stored as numbers. You must check the number format to distinguish dates from regular numbers. Both built-in format IDs (14-22, 27-36, 45-47, 50-58) and custom format strings must be checked.

2. **1900 date bug**: Serial dates 1-59 have an off-by-one error due to the phantom February 29, 1900. Dates on or after March 1, 1900 (serial 61+) are correct.

3. **Empty rows/cells**: Rows and cells may be omitted when empty. Do not assume contiguous numbering.

4. **Shared string vs. inline string**: Both are valid. Excel uses shared strings exclusively, but programmatic tools may use inline strings.

5. **Required fills**: The first two fills must always be `none` and `gray125`, in that order. Custom fills start at index 2.

6. **Relationship ID mapping**: Always use relationship IDs (not file naming patterns) to locate parts. The `r:id` attribute on `<sheet>` is the canonical way to find worksheet files.

7. **Column width units**: Column widths are in "character width" units based on the maximum digit width of the default font, not in pixels or points.

8. **Row height units**: Row heights are in points (1/72 inch).

9. **EMU units**: Drawing positions use English Metric Units (1 inch = 914400 EMUs).

10. **XML namespace prefixes**: Namespace prefixes are not fixed. Always match on namespace URI, never on prefix string.

11. **Calculation chain**: Optional in output. Excel will recalculate on open if missing.

12. **Maximum dimensions**: 16,384 columns (A to XFD) by 1,048,576 rows.

13. **String length**: Cell strings can be up to 32,767 characters.

14. **Formula length**: Formulas can be up to 8,192 characters.

15. **Format code length**: Must be less than 255 characters.

### Number Format Code Syntax

Format codes consist of up to four sections separated by semicolons:

```
positive;negative;zero;text
```

- If one section: applies to all numeric values
- If two sections: first for positive and zero, second for negative
- If three sections: positive, negative, zero
- If four sections: positive, negative, zero, text

**Placeholders:**
| Symbol | Meaning |
|--------|---------|
| `0` | Required digit (displays 0 if no digit) |
| `#` | Optional digit (no display if no digit) |
| `?` | Optional digit (displays space if no digit) |
| `.` | Decimal point |
| `,` | Thousands separator |
| `%` | Multiply by 100 and display percent sign |
| `E+`, `E-` | Scientific notation |
| `@` | Text placeholder |
| `*` | Repeat next character to fill column |
| `_` | Skip width of next character |
| `"text"` | Literal text |
| `\c` | Literal character |
| `[Color]` | Color name: `[Black]`, `[Blue]`, `[Cyan]`, `[Green]`, `[Magenta]`, `[Red]`, `[White]`, `[Yellow]`, `[Color1]`-`[Color56]` |
| `[condition]` | Conditional format: `[>100]`, `[<=0]` |
| `[$locale]` | Locale code: `[$-409]` for en-US |

**Date/Time tokens:**
| Token | Meaning |
|-------|---------|
| `yy` | Two-digit year |
| `yyyy` | Four-digit year |
| `m` | Month (1-12) or minutes (context-dependent) |
| `mm` | Month (01-12) or minutes (00-59) |
| `mmm` | Abbreviated month name (Jan-Dec) |
| `mmmm` | Full month name (January-December) |
| `mmmmm` | First letter of month (J, F, M, ...) |
| `d` | Day (1-31) |
| `dd` | Day (01-31) |
| `ddd` | Abbreviated weekday (Sun-Sat) |
| `dddd` | Full weekday (Sunday-Saturday) |
| `h` | Hour (0-23 or 1-12 with AM/PM) |
| `hh` | Hour with leading zero |
| `m` | Minutes (when following h or preceding s) |
| `mm` | Minutes with leading zero |
| `s` | Seconds (0-59) |
| `ss` | Seconds with leading zero |
| `.0`, `.00`, `.000` | Fractional seconds |
| `[h]` | Elapsed hours (can exceed 24) |
| `[m]` | Elapsed minutes |
| `[s]` | Elapsed seconds |
| `AM/PM` | 12-hour format with AM/PM |
| `A/P` | 12-hour format with A/P |

**Note on `m`/`mm` ambiguity:** These tokens represent months when they appear in a date context (after `y`/`d` or before `d`/`y`) and minutes when they appear in a time context (after `h` or before `s`).

---

## Appendix: File Extension Variants

| Extension | Description | Macro Support |
|-----------|-------------|---------------|
| `.xlsx` | Standard Excel workbook | No |
| `.xlsm` | Macro-enabled workbook | Yes (VBA) |
| `.xltx` | Template | No |
| `.xltm` | Macro-enabled template | Yes (VBA) |
| `.xlsb` | Binary workbook (different format) | Yes |
| `.xlam` | Add-in | Yes (VBA) |

The only difference between `.xlsx` and `.xlsm` in terms of OPC structure is:
- The workbook content type changes to `application/vnd.ms-excel.sheet.macroEnabled.main+xml`
- An additional `xl/vbaProject.bin` part is present

---

## References

- [ECMA-376 Standard](https://ecma-international.org/publications-and-standards/standards/ecma-376/)
- [ISO/IEC 29500-1:2016](https://www.iso.org/standard/71691.html)
- [Microsoft Open XML SDK Documentation](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/structure-of-a-spreadsheetml-document)
- [MS-XLSX: Excel Extensions](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/2c5dee00-eff2-4b22-92b6-0738acd4475e)
- [MS-OI29500: Office Implementation Info](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/17d11129-219b-4e2c-88db-45844d21e528)
- [MS-OE376: Office ECMA-376 Implementation](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/db9b9b72-b10b-4e7e-844c-09f88c972219)
- [Library of Congress XLSX Format Description](https://www.loc.gov/preservation/digital/formats/fdd/fdd000398.shtml)
- [Office Open XML on Wikipedia](https://en.wikipedia.org/wiki/Office_Open_XML)
