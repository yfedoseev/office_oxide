# DOCX (Office Open XML WordprocessingML) Technical Specification Reference

## Table of Contents

1. [Standards and Specifications](#1-standards-and-specifications)
2. [Package Structure (ZIP Archive)](#2-package-structure-zip-archive)
3. [Content Types](#3-content-types)
4. [Relationship System](#4-relationship-system)
5. [XML Namespaces and Schemas](#5-xml-namespaces-and-schemas)
6. [Document Structure Overview](#6-document-structure-overview)
7. [Paragraphs and Runs](#7-paragraphs-and-runs)
8. [Text Formatting (Run Properties)](#8-text-formatting-run-properties)
9. [Paragraph Formatting (Paragraph Properties)](#9-paragraph-formatting-paragraph-properties)
10. [Tables](#10-tables)
11. [Images and Drawings](#11-images-and-drawings)
12. [Styles](#12-styles)
13. [Fonts and Themes](#13-fonts-and-themes)
14. [Numbering and Lists](#14-numbering-and-lists)
15. [Headers and Footers](#15-headers-and-footers)
16. [Sections](#16-sections)
17. [Bookmarks, Hyperlinks, and Fields](#17-bookmarks-hyperlinks-and-fields)
18. [Document Settings](#18-document-settings)
19. [Document Properties (Metadata)](#19-document-properties-metadata)
20. [Strict vs. Transitional Conformance](#20-strict-vs-transitional-conformance)
21. [Implementer Notes](#21-implementer-notes)

---

## 1. Standards and Specifications

### Official Standards

| Standard | Description |
|----------|-------------|
| **ECMA-376** (Editions 1-5) | Original standard by Ecma International (TC45). Edition 1 (2006), current Edition 5 (2021). |
| **ISO/IEC 29500:2008-2016** | International standard, four parts. Defines both Strict and Transitional conformance. |

### ISO/IEC 29500 Parts

| Part | Title | Scope |
|------|-------|-------|
| **Part 1** | Fundamentals and Markup Language Reference | Core markup: WordprocessingML, SpreadsheetML, PresentationML, DrawingML, SharedML. Defines Strict conformance. |
| **Part 2** | Open Packaging Conventions (OPC) | ZIP-based container format, relationships, content types, digital signatures. |
| **Part 3** | Markup Compatibility and Extensibility (MCE) | Mechanisms for forward/backward compatibility of markup. |
| **Part 4** | Transitional Migration Features | Legacy elements/attributes retained for backward compatibility with older binary formats. |

### MIME Type

```
application/vnd.openxmlformats-officedocument.wordprocessingml.document
```

### File Extension

`.docx` (used by both Strict and Transitional variants)

### PRONOM Identifier

`fmt/412`

---

## 2. Package Structure (ZIP Archive)

A `.docx` file is a ZIP archive conforming to the Open Packaging Conventions (OPC). The archive contains XML parts, binary media, and relationship descriptors organized in a defined directory structure.

### Typical Directory Layout

```
my_document.docx (ZIP archive)
|
+-- [Content_Types].xml                    # REQUIRED: maps parts to MIME types
+-- _rels/
|   +-- .rels                              # REQUIRED: package-level relationships
+-- word/
|   +-- document.xml                       # REQUIRED: main document content
|   +-- styles.xml                         # Style definitions
|   +-- settings.xml                       # Document settings
|   +-- fontTable.xml                      # Font table
|   +-- numbering.xml                      # Numbering/list definitions
|   +-- footnotes.xml                      # Footnotes
|   +-- endnotes.xml                       # Endnotes
|   +-- comments.xml                       # Comments
|   +-- header1.xml                        # Header part(s)
|   +-- header2.xml
|   +-- footer1.xml                        # Footer part(s)
|   +-- footer2.xml
|   +-- glossary/                          # Glossary document (optional)
|   |   +-- document.xml
|   +-- _rels/
|   |   +-- document.xml.rels             # Relationships for document.xml
|   +-- theme/
|   |   +-- theme1.xml                    # Theme definitions (colors, fonts, effects)
|   +-- media/
|       +-- image1.png                     # Embedded media files
|       +-- image2.jpeg
+-- docProps/
    +-- core.xml                           # Core properties (Dublin Core metadata)
    +-- app.xml                            # Application-specific properties
    +-- custom.xml                         # Custom properties (optional)
    +-- thumbnail.jpeg                     # Thumbnail preview (optional)
```

### Required Parts (Minimum Valid Document)

A minimal valid `.docx` requires only:

1. `[Content_Types].xml`
2. `_rels/.rels`
3. `word/document.xml`

### Minimum `document.xml`

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Hello, World!</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>
```

---

## 3. Content Types

The `[Content_Types].xml` file at the package root is required. It maps every part in the package to a MIME content type. It uses two mechanisms: **Default** (by file extension) and **Override** (by specific part name).

### Structure

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">

  <!-- Default: maps file extensions to content types -->
  <Default Extension="rels"
           ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"
           ContentType="application/xml"/>
  <Default Extension="png"
           ContentType="image/png"/>
  <Default Extension="jpeg"
           ContentType="image/jpeg"/>

  <!-- Override: maps specific parts to content types -->
  <Override PartName="/word/document.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/fontTable.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
  <Override PartName="/word/numbering.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/word/footnotes.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
  <Override PartName="/word/endnotes.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
  <Override PartName="/word/comments.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
  <Override PartName="/word/header1.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  <Override PartName="/word/theme/theme1.xml"
            ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/docProps/core.xml"
            ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml"
            ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
```

### Lookup Rules

1. First check for an `Override` matching the exact part name.
2. If no override, fall back to a `Default` matching the file extension.
3. If neither matches, the part has no content type (invalid package).

### Common Content Types Reference

| Part | Content Type |
|------|-------------|
| Main Document | `application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml` |
| Styles | `application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml` |
| Settings | `application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml` |
| Font Table | `application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml` |
| Numbering | `application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml` |
| Header | `application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml` |
| Footer | `application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml` |
| Footnotes | `application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml` |
| Endnotes | `application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml` |
| Comments | `application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml` |
| Glossary Document | `application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml` |
| Theme | `application/vnd.openxmlformats-officedocument.theme+xml` |
| Core Properties | `application/vnd.openxmlformats-package.core-properties+xml` |
| Extended Properties | `application/vnd.openxmlformats-officedocument.extended-properties+xml` |
| Relationships | `application/vnd.openxmlformats-package.relationships+xml` |

---

## 4. Relationship System

Relationships define how parts within the package are connected to each other and to external resources. The relationship system is defined by OPC (ISO/IEC 29500-2).

### Relationship File Locations

Every part can have an associated `.rels` file in a `_rels` subdirectory relative to the part's location. The `.rels` file name is formed by appending `.rels` to the part's file name.

| Source Part | Relationship File |
|------------|-------------------|
| Package root | `_rels/.rels` |
| `word/document.xml` | `word/_rels/document.xml.rels` |
| `word/header1.xml` | `word/_rels/header1.xml.rels` |
| `word/footer1.xml` | `word/_rels/footer1.xml.rels` |

### Package-Level Relationships (`_rels/.rels`)

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                Target="word/document.xml"/>
  <Relationship Id="rId2"
                Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
                Target="docProps/core.xml"/>
  <Relationship Id="rId3"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
                Target="docProps/app.xml"/>
  <Relationship Id="rId4"
                Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail"
                Target="docProps/thumbnail.jpeg"/>
</Relationships>
```

### Document-Level Relationships (`word/_rels/document.xml.rels`)

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
                Target="styles.xml"/>
  <Relationship Id="rId2"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
                Target="settings.xml"/>
  <Relationship Id="rId3"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"
                Target="fontTable.xml"/>
  <Relationship Id="rId4"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
                Target="theme/theme1.xml"/>
  <Relationship Id="rId5"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
                Target="numbering.xml"/>
  <Relationship Id="rId6"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
                Target="footnotes.xml"/>
  <Relationship Id="rId7"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
                Target="endnotes.xml"/>
  <Relationship Id="rId8"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
                Target="header1.xml"/>
  <Relationship Id="rId9"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
                Target="footer1.xml"/>
  <Relationship Id="rId10"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                Target="media/image1.png"/>
  <Relationship Id="rId11"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
                Target="https://example.com"
                TargetMode="External"/>
</Relationships>
```

### Relationship Attributes

| Attribute | Description |
|-----------|-------------|
| `Id` | Unique identifier within the `.rels` file. Referenced by `r:id` attributes in the source XML. |
| `Type` | URI identifying the relationship type (see table below). |
| `Target` | Relative URI to the target part, or absolute URI for external targets. |
| `TargetMode` | `Internal` (default, omitted) or `External`. |

### Common Relationship Types (Transitional)

| Relationship | Type URI |
|-------------|----------|
| Office Document | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument` |
| Styles | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles` |
| Settings | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings` |
| Font Table | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable` |
| Theme | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme` |
| Numbering | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering` |
| Footnotes | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes` |
| Endnotes | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes` |
| Comments | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments` |
| Header | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/header` |
| Footer | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer` |
| Image | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/image` |
| Hyperlink | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink` |
| Glossary Document | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument` |
| Core Properties | `http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties` |
| Extended Properties | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties` |
| Custom Properties | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties` |
| Thumbnail | `http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail` |

### Explicit vs. Implicit Relationships

- **Explicit relationships**: The `r:id` attribute in source XML directly references the `Id` in the `.rels` file. The relationship target is the resource itself. Example: images (`r:embed="rId10"`), hyperlinks (`r:id="rId11"`), headers/footers (`r:id="rId8"`).

- **Implicit relationships**: The relationship target is implied by the containing element/context. The `id` in the source XML refers to an element *within* the implied target part, not to a relationship entry. Example: footnotes (the `<w:footnoteReference w:id="1"/>` refers to the footnote with `w:id="1"` in footnotes.xml).

---

## 5. XML Namespaces and Schemas

### Core Namespaces (Transitional)

| Prefix | Namespace URI | Usage |
|--------|--------------|-------|
| `w` | `http://schemas.openxmlformats.org/wordprocessingml/2006/main` | WordprocessingML elements and attributes |
| `r` | `http://schemas.openxmlformats.org/officeDocument/2006/relationships` | Relationship references (`r:id`, `r:embed`) |
| `wp` | `http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing` | Drawing positioning in word processing documents |
| `a` | `http://schemas.openxmlformats.org/drawingml/2006/main` | DrawingML core elements |
| `pic` | `http://schemas.openxmlformats.org/drawingml/2006/picture` | Picture-specific DrawingML |
| `m` | `http://schemas.openxmlformats.org/officeDocument/2006/math` | Office Math Markup Language (OMML) |
| `mc` | `http://schemas.openxmlformats.org/markup-compatibility/2006` | Markup Compatibility and Extensibility |
| `v` | `urn:schemas-microsoft-com:vml` | Vector Markup Language (legacy, Transitional only) |
| `o` | `urn:schemas-microsoft-com:office:office` | Office VML extensions (legacy) |
| `w10` | `urn:schemas-microsoft-com:office:word` | Word VML extensions (legacy) |
| `wne` | `http://schemas.microsoft.com/office/word/2006/wordml` | Word extensions |
| `w14` | `http://schemas.microsoft.com/office/word/2010/wordml` | Word 2010 extensions |
| `w15` | `http://schemas.microsoft.com/office/word/2012/wordml` | Word 2012/2013 extensions |
| `w16se` | `http://schemas.microsoft.com/office/word/2015/wordml/symex` | Symbol extensions |
| `wp14` | `http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing` | Drawing extensions |
| `a14` | `http://schemas.microsoft.com/office/drawing/2010/main` | DrawingML 2010 extensions |

### Package/OPC Namespaces

| Prefix | Namespace URI | Usage |
|--------|--------------|-------|
| (default) | `http://schemas.openxmlformats.org/package/2006/relationships` | Relationship parts |
| (default) | `http://schemas.openxmlformats.org/package/2006/content-types` | Content types |
| (default) | `http://schemas.openxmlformats.org/package/2006/metadata/core-properties` | Core properties wrapper |
| `dc` | `http://purl.org/dc/elements/1.1/` | Dublin Core metadata |
| `dcterms` | `http://purl.org/dc/terms/` | Dublin Core terms |
| `dcmitype` | `http://purl.org/dc/dcmitype/` | Dublin Core DCMI type |
| `cp` | `http://schemas.openxmlformats.org/package/2006/metadata/core-properties` | Core properties elements |
| `ep` | `http://schemas.openxmlformats.org/officeDocument/2006/extended-properties` | Extended properties |

### Strict Namespaces (ISO/IEC 29500 Strict)

Strict conformance uses different namespace URIs (see [Section 20](#20-strict-vs-transitional-conformance) for details):

| Prefix | Transitional URI | Strict URI |
|--------|-----------------|------------|
| `w` | `http://schemas.openxmlformats.org/wordprocessingml/2006/main` | `http://purl.oclc.org/ooxml/wordprocessingml/main` |
| `r` | `http://schemas.openxmlformats.org/officeDocument/2006/relationships` | `http://purl.oclc.org/ooxml/officeDocument/relationships` |
| `a` | `http://schemas.openxmlformats.org/drawingml/2006/main` | `http://purl.oclc.org/ooxml/drawingml/main` |
| `wp` | `http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing` | `http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing` |

### Key Schema Files

The normative schemas are distributed with ECMA-376 and ISO/IEC 29500. The primary WordprocessingML schema is `wml.xsd`. Supporting schemas include:

- `wml.xsd` -- WordprocessingML
- `dml-main.xsd` -- DrawingML core
- `dml-wordprocessingDrawing.xsd` -- DrawingML word processing anchoring
- `dml-picture.xsd` -- DrawingML pictures
- `shared-math.xsd` -- Office Math
- `opc-relationships.xsd` -- OPC relationships
- `opc-contentTypes.xsd` -- OPC content types

---

## 6. Document Structure Overview

### Stories

A WordprocessingML document is organized around the concept of **stories**. A story is a region of content. The main story types are:

- **Main document story** (word/document.xml) -- Required
- **Header story** (word/header{N}.xml)
- **Footer story** (word/footer{N}.xml)
- **Footnote story** (word/footnotes.xml)
- **Endnote story** (word/endnotes.xml)
- **Comment story** (word/comments.xml)
- **Text box story** (inline in document.xml)
- **Glossary document story** (word/glossary/document.xml)

### Block-Level vs. Inline Elements

**Block-level elements** (children of `w:body` or `w:tc`):
- `w:p` -- Paragraph
- `w:tbl` -- Table
- `w:sdt` -- Structured document tag (content control, block-level)
- `w:customXml` -- Custom XML block
- `w:altChunk` -- External content import

**Inline elements** (children of `w:p`):
- `w:r` -- Run
- `w:hyperlink` -- Hyperlink
- `w:sdt` -- Structured document tag (inline)
- `w:bookmarkStart` / `w:bookmarkEnd` -- Bookmark markers
- `w:commentRangeStart` / `w:commentRangeEnd` -- Comment range markers
- `w:fldSimple` -- Simple field
- `w:ins` / `w:del` -- Tracked changes (insertions/deletions)
- `w:smartTag` -- Smart tag

### High-Level Document XML

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
    xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
    xmlns:v="urn:schemas-microsoft-com:vml"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:w10="urn:schemas-microsoft-com:office:word"
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
    xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
    xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
    xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
    xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
    mc:Ignorable="w14 w15 wp14">
  <w:body>
    <!-- Block-level content: paragraphs, tables, etc. -->
    <w:p>...</w:p>
    <w:tbl>...</w:tbl>
    <w:p>...</w:p>

    <!-- Final section properties (for last section) -->
    <w:sectPr>...</w:sectPr>
  </w:body>
</w:document>
```

---

## 7. Paragraphs and Runs

### Paragraph (`w:p`)

A paragraph is the fundamental block-level element. It contains optional properties followed by inline content.

```xml
<w:p w:rsidR="00A77427" w:rsidRDefault="007A3403">
  <!-- Paragraph properties (must be first child if present) -->
  <w:pPr>
    <w:pStyle w:val="Heading1"/>
    <w:jc w:val="center"/>
  </w:pPr>

  <!-- Inline content -->
  <w:r>
    <w:t>Hello </w:t>
  </w:r>
  <w:r>
    <w:rPr>
      <w:b/>
    </w:rPr>
    <w:t>World</w:t>
  </w:r>
</w:p>
```

### Run (`w:r`)

A run is a contiguous region of text with a common set of formatting properties. A run contains optional run properties (`w:rPr`) followed by content elements.

```xml
<w:r>
  <w:rPr>
    <w:b/>              <!-- bold -->
    <w:i/>              <!-- italic -->
    <w:sz w:val="28"/>  <!-- font size in half-points (28 = 14pt) -->
  </w:rPr>
  <w:t>Formatted text</w:t>
</w:r>
```

### Run Content Elements

| Element | Description |
|---------|-------------|
| `w:t` | Text content. Use `xml:space="preserve"` to preserve leading/trailing whitespace. |
| `w:br` | Break (line break `w:type="textWrapping"`, page break `w:type="page"`, column break `w:type="column"`). |
| `w:tab` | Tab character. |
| `w:cr` | Carriage return. |
| `w:sym` | Symbol character (`w:font`, `w:char` attributes). |
| `w:drawing` | DrawingML object (images, shapes). |
| `w:pict` | VML picture (legacy, Transitional only). |
| `w:object` | Embedded OLE object. |
| `w:ruby` | Ruby (phonetic guide) text. |
| `w:footnoteReference` | Footnote reference. |
| `w:endnoteReference` | Endnote reference. |
| `w:commentReference` | Comment reference. |
| `w:fldChar` | Complex field character (begin/separate/end). |
| `w:instrText` | Field instruction text (within complex fields). |
| `w:lastRenderedPageBreak` | Marks where Word last inserted a page break during rendering. |
| `w:softHyphen` | Soft (optional) hyphen. |
| `w:noBreakHyphen` | Non-breaking hyphen. |

### Text Element Details

```xml
<!-- Whitespace preserved -->
<w:t xml:space="preserve"> text with spaces </w:t>

<!-- Default: leading/trailing whitespace may be stripped -->
<w:t>text without leading/trailing spaces</w:t>
```

**Implementer note**: Always emit `xml:space="preserve"` on `w:t` elements when the text begins or ends with whitespace, or contains only whitespace. Failure to do so will cause space loss.

---

## 8. Text Formatting (Run Properties)

Run properties (`w:rPr`) appear as the first child of `w:r` and control character-level formatting.

### Common Run Property Elements

| Element | Attribute(s) | Description |
|---------|-------------|-------------|
| `w:b` | `w:val` (optional, default true) | Bold. `<w:b/>` enables, `<w:b w:val="0"/>` disables. |
| `w:bCs` | `w:val` | Bold for complex script text. |
| `w:i` | `w:val` | Italic. |
| `w:iCs` | `w:val` | Italic for complex script text. |
| `w:u` | `w:val` (e.g., `single`, `double`, `wave`, `dotted`, `none`) | Underline. |
| `w:strike` | `w:val` | Strikethrough. |
| `w:dstrike` | `w:val` | Double strikethrough. |
| `w:outline` | `w:val` | Outline effect. |
| `w:shadow` | `w:val` | Shadow effect. |
| `w:emboss` | `w:val` | Emboss effect. |
| `w:imprint` | `w:val` | Imprint/engrave effect. |
| `w:caps` | `w:val` | All capitals display. |
| `w:smallCaps` | `w:val` | Small capitals display. |
| `w:vanish` | `w:val` | Hidden text. |
| `w:color` | `w:val` (hex RGB, e.g. `"FF0000"`), `w:themeColor`, `w:themeTint`, `w:themeShade` | Text color. |
| `w:sz` | `w:val` (half-points, e.g., `24` = 12pt) | Font size. |
| `w:szCs` | `w:val` | Font size for complex script. |
| `w:rFonts` | `w:ascii`, `w:hAnsi`, `w:eastAsia`, `w:cs`, `w:asciiTheme`, `w:hAnsiTheme`, `w:eastAsiaTheme`, `w:cstheme` | Font family. Can specify directly or via theme reference. |
| `w:highlight` | `w:val` (e.g., `yellow`, `green`, `cyan`) | Text highlight color (predefined palette). |
| `w:shd` | `w:val`, `w:color`, `w:fill` | Shading/background color. |
| `w:vertAlign` | `w:val` (`superscript`, `subscript`, `baseline`) | Vertical alignment (superscript/subscript). |
| `w:spacing` | `w:val` (twips) | Character spacing expansion/compression. |
| `w:kern` | `w:val` (half-points) | Minimum font size for kerning. |
| `w:position` | `w:val` (half-points) | Vertical position offset. |
| `w:rStyle` | `w:val` (style ID) | Character style reference. |
| `w:lang` | `w:val`, `w:eastAsia`, `w:bidi` | Language tag. |
| `w:noProof` | `w:val` | Disable spelling/grammar check. |
| `w:rtl` | `w:val` | Right-to-left text direction. |

### Example: Complex Run Formatting

```xml
<w:r>
  <w:rPr>
    <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
    <w:b/>
    <w:i/>
    <w:color w:val="FF0000"/>
    <w:sz w:val="28"/>           <!-- 14pt -->
    <w:u w:val="single"/>
    <w:highlight w:val="yellow"/>
    <w:lang w:val="en-US"/>
  </w:rPr>
  <w:t>Bold, italic, red, 14pt, underlined, highlighted text</w:t>
</w:r>
```

---

## 9. Paragraph Formatting (Paragraph Properties)

Paragraph properties (`w:pPr`) appear as the first child of `w:p` and control paragraph-level formatting.

### Common Paragraph Property Elements

| Element | Attribute(s) | Description |
|---------|-------------|-------------|
| `w:pStyle` | `w:val` (style ID) | Paragraph style reference. |
| `w:jc` | `w:val` (`left`, `center`, `right`, `both`, `distribute`) | Justification/alignment. `both` = justified. |
| `w:ind` | `w:left`, `w:right`, `w:hanging`, `w:firstLine` (twips) | Indentation. `w:hanging` and `w:firstLine` are mutually exclusive. In Strict, `w:start`/`w:end` replace `w:left`/`w:right`. |
| `w:spacing` | `w:before`, `w:after` (twips), `w:line` (twips or 240ths of a line), `w:lineRule` (`auto`, `exact`, `atLeast`) | Space before/after paragraph, line spacing. |
| `w:keepNext` | `w:val` | Keep with next paragraph (no page break between). |
| `w:keepLines` | `w:val` | Keep all lines on same page. |
| `w:widowControl` | `w:val` | Widow/orphan control. |
| `w:pageBreakBefore` | `w:val` | Force page break before paragraph. |
| `w:suppressAutoHyphens` | `w:val` | Suppress automatic hyphenation. |
| `w:outlineLvl` | `w:val` (0-9) | Outline level (0 = Heading 1). |
| `w:numPr` | child elements | Numbering/list properties (see Section 14). |
| `w:pBdr` | child elements | Paragraph borders (`w:top`, `w:bottom`, `w:left`, `w:right`, `w:between`, `w:bar`). |
| `w:shd` | `w:val`, `w:color`, `w:fill` | Paragraph shading. |
| `w:tabs` | child `w:tab` elements | Custom tab stops. |
| `w:rPr` | child elements | Default run properties for the paragraph mark. |
| `w:sectPr` | child elements | Section properties (non-final sections only). |
| `w:textAlignment` | `w:val` (`auto`, `top`, `center`, `baseline`, `bottom`) | Vertical text alignment within line. |
| `w:bidi` | `w:val` | Right-to-left paragraph. |
| `w:suppressLineNumbers` | `w:val` | Suppress line numbers. |
| `w:contextualSpacing` | `w:val` | Suppress spacing between paragraphs of same style. |

### Indentation Units

Indentation values are in **twips** (twentieths of a point). 1 inch = 1440 twips. 1 cm = 567 twips (approximately).

### Example: Paragraph with Full Formatting

```xml
<w:p>
  <w:pPr>
    <w:pStyle w:val="BodyText"/>
    <w:jc w:val="both"/>
    <w:ind w:left="720" w:right="720" w:firstLine="360"/>
    <w:spacing w:before="120" w:after="120" w:line="276" w:lineRule="auto"/>
    <w:keepNext/>
    <w:pBdr>
      <w:bottom w:val="single" w:sz="4" w:space="1" w:color="000000"/>
    </w:pBdr>
    <w:tabs>
      <w:tab w:val="left" w:pos="2880"/>
      <w:tab w:val="center" w:pos="4680"/>
      <w:tab w:val="right" w:pos="9360"/>
    </w:tabs>
  </w:pPr>
  <w:r>
    <w:t>Justified paragraph with borders, tabs, and indentation.</w:t>
  </w:r>
</w:p>
```

---

## 10. Tables

Tables are block-level elements. A table consists of rows, and rows consist of cells. Cells contain block-level elements (primarily paragraphs).

### Table Element Hierarchy

```
w:tbl
  w:tblPr          -- Table properties (first child)
  w:tblGrid        -- Column width definitions
    w:gridCol       -- One per column
  w:tr              -- Table row (one or more)
    w:trPr          -- Row properties (optional, first child of tr)
    w:tc            -- Table cell (one or more)
      w:tcPr        -- Cell properties (optional, first child of tc)
      w:p           -- Cell content (one or more paragraphs, REQUIRED)
```

### Complete Table Example

```xml
<w:tbl>
  <!-- Table properties -->
  <w:tblPr>
    <w:tblStyle w:val="TableGrid"/>
    <w:tblW w:w="5000" w:type="pct"/>  <!-- 100% width (5000 = 100%) -->
    <w:tblBorders>
      <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>
    </w:tblBorders>
    <w:tblLayout w:type="fixed"/>       <!-- fixed or autofit -->
    <w:tblCellMar>
      <w:top w:w="0" w:type="dxa"/>
      <w:left w:w="108" w:type="dxa"/>
      <w:bottom w:w="0" w:type="dxa"/>
      <w:right w:w="108" w:type="dxa"/>
    </w:tblCellMar>
  </w:tblPr>

  <!-- Column grid -->
  <w:tblGrid>
    <w:gridCol w:w="3000"/>
    <w:gridCol w:w="3000"/>
    <w:gridCol w:w="3000"/>
  </w:tblGrid>

  <!-- Row 1 -->
  <w:tr>
    <w:trPr>
      <w:tblHeader/>               <!-- Repeat as header row -->
    </w:trPr>
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="3000" w:type="dxa"/>
        <w:shd w:val="clear" w:color="auto" w:fill="CCCCCC"/>
      </w:tcPr>
      <w:p><w:r><w:t>Header 1</w:t></w:r></w:p>
    </w:tc>
    <w:tc>
      <w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>
      <w:p><w:r><w:t>Header 2</w:t></w:r></w:p>
    </w:tc>
    <w:tc>
      <w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>
      <w:p><w:r><w:t>Header 3</w:t></w:r></w:p>
    </w:tc>
  </w:tr>

  <!-- Row 2 -->
  <w:tr>
    <w:tc>
      <w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>
      <w:p><w:r><w:t>Cell A</w:t></w:r></w:p>
    </w:tc>
    <w:tc>
      <w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>
      <w:p><w:r><w:t>Cell B</w:t></w:r></w:p>
    </w:tc>
    <w:tc>
      <w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>
      <w:p><w:r><w:t>Cell C</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
</w:tbl>
```

### Table Width Types (`w:type` attribute)

| Value | Description |
|-------|-------------|
| `dxa` | Twentieths of a point (twips). Absolute width. |
| `pct` | Fiftieths of a percent. `5000` = 100%. |
| `auto` | Automatically determined. |
| `nil` | Zero width. |

### Cell Merging

**Horizontal merge** (column span): Use `w:gridSpan` in `w:tcPr`.

```xml
<w:tc>
  <w:tcPr>
    <w:tcW w:w="6000" w:type="dxa"/>
    <w:gridSpan w:val="2"/>            <!-- Spans 2 columns -->
  </w:tcPr>
  <w:p><w:r><w:t>Spans two columns</w:t></w:r></w:p>
</w:tc>
```

**Vertical merge** (row span): Use `w:vMerge` in `w:tcPr`.

```xml
<!-- First cell in vertical span (restart) -->
<w:tc>
  <w:tcPr>
    <w:vMerge w:val="restart"/>
  </w:tcPr>
  <w:p><w:r><w:t>Spans multiple rows</w:t></w:r></w:p>
</w:tc>

<!-- Continuation cells (in subsequent rows) -->
<w:tc>
  <w:tcPr>
    <w:vMerge/>                       <!-- val="continue" is implied default -->
  </w:tcPr>
  <w:p/>                               <!-- Empty paragraph required -->
</w:tc>
```

### Key Table Properties

| Element | Description |
|---------|-------------|
| `w:tblStyle` | Reference to a table style in styles.xml. |
| `w:tblW` | Table width. |
| `w:jc` | Table alignment (left, center, right). |
| `w:tblInd` | Table indentation from leading margin. |
| `w:tblBorders` | Table borders (top, bottom, left, right, insideH, insideV). |
| `w:tblCellMar` | Default cell margins. |
| `w:tblLayout` | Layout algorithm (`fixed` or `autofit`). |
| `w:tblLook` | Conditional formatting flags (first row, last row, banding, etc.). |

### Key Cell Properties

| Element | Description |
|---------|-------------|
| `w:tcW` | Cell width. |
| `w:gridSpan` | Number of grid columns spanned. |
| `w:vMerge` | Vertical merge (restart/continue). |
| `w:tcBorders` | Cell-specific borders. |
| `w:shd` | Cell shading/background. |
| `w:vAlign` | Vertical alignment within cell (`top`, `center`, `bottom`). |
| `w:textDirection` | Text direction within cell. |
| `w:noWrap` | Prevent text wrapping. |

**Implementer note**: Every `w:tc` must contain at least one `w:p` element. An empty cell still requires `<w:p/>`.

---

## 11. Images and Drawings

Images and other graphical objects are embedded using DrawingML within a `w:drawing` element inside a run. Legacy VML (`w:pict`) is also encountered in Transitional documents.

### DrawingML Image Placement

Images can be placed **inline** (flows with text) or **anchored** (positioned on the page independently).

#### Inline Image

```xml
<w:r>
  <w:drawing>
    <wp:inline distT="0" distB="0" distL="0" distR="0">
      <wp:extent cx="5486400" cy="3200400"/>   <!-- Size in EMUs -->
      <wp:effectExtent l="0" t="0" r="0" b="0"/>
      <wp:docPr id="1" name="Picture 1" descr="Alt text"/>
      <wp:cNvGraphicFramePr>
        <a:graphicFrameLocks noChangeAspect="1"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>
      </wp:cNvGraphicFramePr>
      <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:nvPicPr>
              <pic:cNvPr id="0" name="image1.png"/>
              <pic:cNvPicPr/>
            </pic:nvPicPr>
            <pic:blipFill>
              <a:blip r:embed="rId10"/>       <!-- Relationship ID to image -->
              <a:stretch>
                <a:fillRect/>
              </a:stretch>
            </pic:blipFill>
            <pic:spPr>
              <a:xfrm>
                <a:off x="0" y="0"/>
                <a:ext cx="5486400" cy="3200400"/>
              </a:xfrm>
              <a:prstGeom prst="rect">
                <a:avLst/>
              </a:prstGeom>
            </pic:spPr>
          </pic:pic>
        </a:graphicData>
      </a:graphic>
    </wp:inline>
  </w:drawing>
</w:r>
```

#### Anchored Image (Floating)

```xml
<w:r>
  <w:drawing>
    <wp:anchor distT="0" distB="0" distL="114300" distR="114300"
               simplePos="0" relativeHeight="251658240"
               behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
      <wp:simplePos x="0" y="0"/>
      <wp:positionH relativeFrom="column">
        <wp:posOffset>0</wp:posOffset>
      </wp:positionH>
      <wp:positionV relativeFrom="paragraph">
        <wp:posOffset>0</wp:posOffset>
      </wp:positionV>
      <wp:extent cx="2743200" cy="1828800"/>
      <wp:effectExtent l="0" t="0" r="0" b="0"/>
      <wp:wrapSquare wrapText="bothSides"/>    <!-- Text wrapping mode -->
      <wp:docPr id="2" name="Picture 2"/>
      <wp:cNvGraphicFramePr/>
      <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <!-- Same a:graphicData/pic:pic structure as inline -->
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:nvPicPr>
              <pic:cNvPr id="0" name="image2.png"/>
              <pic:cNvPicPr/>
            </pic:nvPicPr>
            <pic:blipFill>
              <a:blip r:embed="rId12"/>
              <a:stretch><a:fillRect/></a:stretch>
            </pic:blipFill>
            <pic:spPr>
              <a:xfrm>
                <a:off x="0" y="0"/>
                <a:ext cx="2743200" cy="1828800"/>
              </a:xfrm>
              <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            </pic:spPr>
          </pic:pic>
        </a:graphicData>
      </a:graphic>
    </wp:anchor>
  </w:drawing>
</w:r>
```

### Units: EMU (English Metric Units)

DrawingML uses EMUs for all measurements:
- 1 inch = 914400 EMU
- 1 cm = 360000 EMU
- 1 pt = 12700 EMU
- 1 pixel (96dpi) = 9525 EMU

### Image Reference Resolution

1. The `r:embed` attribute on `a:blip` contains a relationship ID (e.g., `rId10`).
2. Look up `rId10` in `word/_rels/document.xml.rels`.
3. The relationship target points to the media file (e.g., `media/image1.png`).

Images can alternatively be linked (not embedded) using `r:link` instead of `r:embed`, with `TargetMode="External"` in the relationship.

### Text Wrapping Modes (Anchored Images)

| Element | Description |
|---------|-------------|
| `wp:wrapNone` | No wrapping (image floats over text). |
| `wp:wrapSquare` | Text wraps in a rectangle around image. |
| `wp:wrapTight` | Text wraps tightly to image contour. |
| `wp:wrapThrough` | Text wraps through image contour. |
| `wp:wrapTopAndBottom` | Text above and below only. |

### Legacy VML Images (Transitional)

Older documents may use VML instead of DrawingML:

```xml
<w:r>
  <w:pict>
    <v:shape style="width:100pt;height:75pt">
      <v:imagedata r:id="rId10" o:title="image"/>
    </v:shape>
  </w:pict>
</w:r>
```

**Implementer note**: A parser should handle both DrawingML (`w:drawing`) and VML (`w:pict`) for full Transitional compatibility. In Strict conformance, VML is not allowed.

---

## 12. Styles

Styles are defined in `word/styles.xml`. They provide reusable formatting definitions that paragraphs, runs, tables, and lists can reference.

### Style Types

| Type Value | Description |
|-----------|-------------|
| `paragraph` | Paragraph formatting (can also include run properties as defaults). |
| `character` | Character/run formatting only. |
| `table` | Table formatting. |
| `numbering` | Numbering/list formatting. |

### Style Definition Structure

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">

  <!-- Document defaults -->
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorHAnsi"
                  w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>
        <w:sz w:val="22"/>       <!-- Default 11pt -->
        <w:szCs w:val="22"/>
        <w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:after="160" w:line="259" w:lineRule="auto"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>

  <!-- Latent style exceptions (toggles for built-in styles) -->
  <w:latentStyles w:defLockedState="0" w:defUIPriority="99"
                  w:defSemiHidden="0" w:defUnhideWhenUsed="0"
                  w:defQFormat="0" w:count="376">
    <w:lsdException w:name="Normal" w:uiPriority="0" w:qFormat="1"/>
    <w:lsdException w:name="heading 1" w:uiPriority="9" w:qFormat="1"/>
    <!-- ... -->
  </w:latentStyles>

  <!-- Normal style (base for most paragraph styles) -->
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>

  <!-- Heading 1 -->
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:link w:val="Heading1Char"/>
    <w:uiPriority w:val="9"/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:keepLines/>
      <w:spacing w:before="240" w:after="0"/>
      <w:outlineLvl w:val="0"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"
                w:eastAsiaTheme="majorEastAsia" w:cstheme="majorBidi"/>
      <w:color w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"/>
      <w:sz w:val="32"/>
    </w:rPr>
  </w:style>

  <!-- Character style (linked to Heading1) -->
  <w:style w:type="character" w:customStyle="1" w:styleId="Heading1Char">
    <w:name w:val="Heading 1 Char"/>
    <w:basedOn w:val="DefaultParagraphFont"/>
    <w:link w:val="Heading1"/>
    <w:uiPriority w:val="9"/>
    <w:rPr>
      <w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/>
      <w:color w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"/>
      <w:sz w:val="32"/>
    </w:rPr>
  </w:style>

  <!-- Table style -->
  <w:style w:type="table" w:styleId="TableGrid">
    <w:name w:val="Table Grid"/>
    <w:basedOn w:val="TableNormal"/>
    <w:uiPriority w:val="39"/>
    <w:tblPr>
      <w:tblBorders>
        <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
        <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
        <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
        <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
        <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
        <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      </w:tblBorders>
    </w:tblPr>
  </w:style>

</w:styles>
```

### Style Properties

| Element | Description |
|---------|-------------|
| `w:name` | Human-readable style name. |
| `w:styleId` | Machine-readable identifier (referenced by `w:pStyle`, `w:rStyle`, `w:tblStyle`). |
| `w:type` | Style type: `paragraph`, `character`, `table`, `numbering`. |
| `w:basedOn` | Parent style ID for inheritance. |
| `w:next` | Style ID to apply to the next paragraph after pressing Enter. |
| `w:link` | Links a paragraph style to its companion character style (and vice versa). |
| `w:default` | If `"1"`, this is the default style for its type. |
| `w:customStyle` | If `"1"`, this is a user-defined (non-built-in) style. |
| `w:qFormat` | Style appears in the Quick Styles gallery. |
| `w:uiPriority` | Sort order in the UI. |
| `w:semiHidden` | Hidden from the main UI. |
| `w:unhideWhenUsed` | Become visible once applied. |
| `w:pPr` | Paragraph properties for this style. |
| `w:rPr` | Run properties for this style. |
| `w:tblPr` | Table properties (table styles only). |
| `w:trPr` | Row properties (table styles only). |
| `w:tcPr` | Cell properties (table styles only). |
| `w:tblStylePr` | Conditional table formatting (first row, last column, banding, etc.). |

### Style Inheritance and Resolution

Effective formatting is resolved by layering properties in this order (later overrides earlier):

1. **Document defaults** (`w:docDefaults`)
2. **Table style** (if within a table)
3. **Numbering style** (if part of a list)
4. **Paragraph style** (`w:pStyle`)
5. **Character style** (`w:rStyle`)
6. **Direct formatting** (inline `w:pPr` / `w:rPr` on the element)

Within the style hierarchy, `w:basedOn` creates an inheritance chain. Properties not explicitly set in a child style are inherited from the parent. The chain terminates at a style with no `w:basedOn` or at the implicit base style for its type.

### Linked Styles

A linked style pairs a paragraph style with a character style. When applied to an entire paragraph, the paragraph style is used. When applied to a text selection within a paragraph, the character style is used. They share the same run properties but the paragraph style additionally includes paragraph properties.

---

## 13. Fonts and Themes

### Font Table (`word/fontTable.xml`)

The font table enumerates all fonts referenced in the document and provides substitution metadata.

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:font w:name="Calibri">
    <w:panose1 w:val="020F0502020204030204"/>
    <w:charset w:val="00"/>
    <w:family w:val="swiss"/>
    <w:pitch w:val="variable"/>
    <w:sig w:usb0="E4002EFF" w:usb1="C000247B" w:usb2="00000009"
           w:usb3="00000000" w:csb0="000001FF" w:csb1="00000000"/>
  </w:font>
  <w:font w:name="Times New Roman">
    <w:panose1 w:val="02020603050405020304"/>
    <w:charset w:val="00"/>
    <w:family w:val="roman"/>
    <w:pitch w:val="variable"/>
    <w:sig w:usb0="E0002EFF" w:usb1="C000785B" w:usb2="00000009"
           w:usb3="00000000" w:csb0="000001FF" w:csb1="00000000"/>
  </w:font>
  <!-- Embedded font (optional) -->
  <w:font w:name="CustomFont">
    <w:embedRegular r:id="rId1"/>
    <w:embedBold r:id="rId2"/>
    <!-- ... -->
  </w:font>
</w:fonts>
```

### Font Properties

| Element | Description |
|---------|-------------|
| `w:panose1` | PANOSE classification (10-byte font classification system). |
| `w:charset` | Character set code. |
| `w:family` | Font family (`roman`, `swiss`, `modern`, `script`, `decorative`, `auto`). |
| `w:pitch` | Pitch (`fixed`, `variable`, `default`). |
| `w:sig` | Font signature (Unicode and code page coverage bitmasks). |
| `w:embedRegular` | Embedded regular font (relationship reference). |
| `w:embedBold` | Embedded bold font. |
| `w:embedItalic` | Embedded italic font. |
| `w:embedBoldItalic` | Embedded bold-italic font. |

### Font Selection in Runs (`w:rFonts`)

The `w:rFonts` element specifies fonts for four Unicode ranges:

| Attribute | Unicode Range |
|-----------|--------------|
| `w:ascii` / `w:asciiTheme` | Basic Latin (U+0000-U+007F) |
| `w:hAnsi` / `w:hAnsiTheme` | All other characters not covered by eastAsia or cs |
| `w:eastAsia` / `w:eastAsiaTheme` | East Asian characters |
| `w:cs` / `w:cstheme` | Complex script characters (Arabic, Hebrew, etc.) |

Direct font names (e.g., `w:ascii="Arial"`) override theme references (e.g., `w:asciiTheme="minorHAnsi"`).

### Theme (`word/theme/theme1.xml`)

A theme defines the document's visual identity: color scheme, font scheme, and effect scheme. It uses DrawingML namespace.

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
         name="Office Theme">
  <a:themeElements>

    <!-- Color Scheme: 12 named colors -->
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>

    <!-- Font Scheme: major (headings) and minor (body) fonts -->
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Calibri Light" panose="020F0302020204030204"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
        <!-- Script-specific overrides -->
        <a:font script="Jpan" typeface="Yu Gothic Light"/>
        <a:font script="Hans" typeface="DengXian Light"/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri" panose="020F0502020204030204"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
        <a:font script="Jpan" typeface="Yu Gothic"/>
        <a:font script="Hans" typeface="DengXian"/>
      </a:minorFont>
    </a:fontScheme>

    <!-- Format/Effects Scheme -->
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <!-- Fill styles for subtle, moderate, intense -->
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:gradFill rotWithShape="1"><!-- ... --></a:gradFill>
        <a:gradFill rotWithShape="1"><!-- ... --></a:gradFill>
      </a:fillStyleLst>
      <a:lnStyleLst><!-- Line styles --></a:lnStyleLst>
      <a:effectStyleLst><!-- Effect styles --></a:effectStyleLst>
      <a:bgFillStyleLst><!-- Background fill styles --></a:bgFillStyleLst>
    </a:fmtScheme>

  </a:themeElements>
</a:theme>
```

### Theme Color References

When a color property references a theme color:

```xml
<w:color w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"/>
```

- `w:val` -- Computed RGB value (for consumers that do not resolve themes).
- `w:themeColor` -- References a named color from the theme's `a:clrScheme`.
- `w:themeTint` -- Lightens the theme color (hex 00-FF; higher = lighter).
- `w:themeShade` -- Darkens the theme color (hex 00-FF; higher = darker).

### Theme Font References

| Theme Identifier | Maps To |
|-----------------|---------|
| `majorHAnsi` | `a:majorFont/a:latin/@typeface` |
| `majorEastAsia` | `a:majorFont/a:ea/@typeface` |
| `majorBidi` | `a:majorFont/a:cs/@typeface` |
| `minorHAnsi` | `a:minorFont/a:latin/@typeface` |
| `minorEastAsia` | `a:minorFont/a:ea/@typeface` |
| `minorBidi` | `a:minorFont/a:cs/@typeface` |

---

## 14. Numbering and Lists

Numbering definitions are stored in `word/numbering.xml`. The system uses a two-level indirection: abstract numbering definitions define the format, and numbering instances reference them.

### Architecture

```
numbering.xml
  w:abstractNum (defines format for up to 9 levels)
    w:lvl (level 0-8)
  w:num (instance referencing an abstractNum)
    w:abstractNumId
    w:lvlOverride (optional per-instance overrides)

document.xml
  w:p / w:pPr / w:numPr
    w:ilvl (indentation level 0-8)
    w:numId (references a w:num)
```

### Numbering XML Example

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">

  <!-- Abstract numbering definition -->
  <w:abstractNum w:abstractNumId="0">
    <w:nsid w:val="12345678"/>
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:tmpl w:val="ABCD1234"/>

    <!-- Level 0: Numbered -->
    <w:lvl w:ilvl="0" w:tplc="04090001">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="720" w:hanging="360"/>
      </w:pPr>
      <w:rPr>
        <w:rFonts w:hint="default"/>
      </w:rPr>
    </w:lvl>

    <!-- Level 1: Lettered -->
    <w:lvl w:ilvl="1" w:tplc="04090003">
      <w:start w:val="1"/>
      <w:numFmt w:val="lowerLetter"/>
      <w:lvlText w:val="%2."/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="1440" w:hanging="360"/>
      </w:pPr>
    </w:lvl>

    <!-- Level 2: Roman numerals -->
    <w:lvl w:ilvl="2" w:tplc="04090005">
      <w:start w:val="1"/>
      <w:numFmt w:val="lowerRoman"/>
      <w:lvlText w:val="%3."/>
      <w:lvlJc w:val="right"/>
      <w:pPr>
        <w:ind w:left="2160" w:hanging="180"/>
      </w:pPr>
    </w:lvl>

    <!-- Levels 3-8 follow the same pattern -->
  </w:abstractNum>

  <!-- Bullet list abstract definition -->
  <w:abstractNum w:abstractNumId="1">
    <w:nsid w:val="87654321"/>
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="\uF0B7"/>          <!-- Bullet character -->
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="720" w:hanging="360"/>
      </w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/>
      </w:rPr>
    </w:lvl>
  </w:abstractNum>

  <!-- Numbering instance -->
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>

  <w:num w:numId="2">
    <w:abstractNumId w:val="1"/>
  </w:num>

</w:numbering>
```

### Referencing from Paragraphs

```xml
<w:p>
  <w:pPr>
    <w:numPr>
      <w:ilvl w:val="0"/>     <!-- Level 0 -->
      <w:numId w:val="1"/>    <!-- Numbering instance 1 -->
    </w:numPr>
  </w:pPr>
  <w:r><w:t>First item</w:t></w:r>
</w:p>
<w:p>
  <w:pPr>
    <w:numPr>
      <w:ilvl w:val="1"/>     <!-- Level 1 (sub-item) -->
      <w:numId w:val="1"/>
    </w:numPr>
  </w:pPr>
  <w:r><w:t>Sub-item</w:t></w:r>
</w:p>
```

### Key Numbering Elements

| Element | Description |
|---------|-------------|
| `w:abstractNum` | Defines the abstract numbering format. `w:abstractNumId` is the unique key. |
| `w:nsid` | Unique number identifier (used for list continuity across saves). |
| `w:multiLevelType` | `singleLevel`, `multilevel`, or `hybridMultilevel`. |
| `w:lvl` | Level definition (0-8). `w:ilvl` attribute specifies level index. |
| `w:start` | Starting number. |
| `w:numFmt` | Number format: `decimal`, `lowerLetter`, `upperLetter`, `lowerRoman`, `upperRoman`, `bullet`, `none`, `ordinal`, etc. |
| `w:lvlText` | Display format string. `%1` = level 1 value, `%2` = level 2 value, etc. |
| `w:lvlJc` | Number justification: `left`, `center`, `right`. |
| `w:isLgl` | Use legal numbering (Arabic numerals for all ancestor levels). |
| `w:lvlRestart` | Level at which numbering restarts. |
| `w:num` | Numbering instance. `w:numId` is the unique key referenced by paragraphs. |
| `w:abstractNumId` | Child of `w:num`; references the abstract definition. |
| `w:lvlOverride` | Per-instance level override within `w:num`. |

### Style-Based Numbering

Numbering can also be applied through styles. The `w:numPr` can appear within a style's `w:pPr`:

```xml
<w:style w:type="paragraph" w:styleId="ListBullet">
  <w:name w:val="List Bullet"/>
  <w:basedOn w:val="Normal"/>
  <w:pPr>
    <w:numPr>
      <w:numId w:val="2"/>
    </w:numPr>
  </w:pPr>
</w:style>
```

---

## 15. Headers and Footers

Headers and footers are stored in separate parts (`word/header{N}.xml`, `word/footer{N}.xml`). They are referenced from section properties via explicit relationships.

### Header/Footer Part Structure

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:p>
    <w:pPr>
      <w:pStyle w:val="Header"/>
      <w:jc w:val="center"/>
    </w:pPr>
    <w:r>
      <w:t>Document Title</w:t>
    </w:r>
  </w:p>
</w:hdr>
```

Footer parts use `w:ftr` as the root element with the same internal structure.

### Referencing from Section Properties

```xml
<w:sectPr>
  <!-- Each type independently controlled -->
  <w:headerReference w:type="default" r:id="rId8"/>
  <w:headerReference w:type="first" r:id="rId9"/>
  <w:headerReference w:type="even" r:id="rId10"/>
  <w:footerReference w:type="default" r:id="rId11"/>
  <w:footerReference w:type="first" r:id="rId12"/>
  <w:footerReference w:type="even" r:id="rId13"/>
  <w:titlePg/>           <!-- Enable different first page header/footer -->
  <!-- ... other section properties ... -->
</w:sectPr>
```

### Header/Footer Types

| `w:type` Value | Description |
|---------------|-------------|
| `default` | Used on all pages unless overridden by `first` or `even`. |
| `first` | Used on the first page of the section. Requires `w:titlePg` in `w:sectPr`. |
| `even` | Used on even-numbered pages. Requires even/odd header setting in settings.xml (`w:evenAndOddHeaders`). |

### Key Behaviors

- If a section does not specify its own headers/footers, it inherits from the previous section.
- Header/footer parts can contain the same content as the document body: paragraphs, tables, images, etc.
- Page numbers are typically implemented using field codes within headers/footers (see Section 17).
- Each header/footer part can have its own relationships (e.g., for images within headers).

---

## 16. Sections

Sections partition the document into regions with distinct page layout properties. Properties include page size, margins, orientation, columns, headers/footers, and page numbering.

### Section Properties Location

- **Non-final sections**: `w:sectPr` is a child of the last `w:pPr` in the section.
- **Final section**: `w:sectPr` is a direct child of `w:body`.

```xml
<w:body>
  <!-- Section 1 content -->
  <w:p>
    <w:r><w:t>Section 1 content</w:t></w:r>
  </w:p>
  <w:p>
    <w:pPr>
      <!-- Section 1 properties (non-final: in last paragraph's pPr) -->
      <w:sectPr>
        <w:type w:val="nextPage"/>
        <w:pgSz w:w="12240" w:h="15840"/>
        <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
                 w:header="720" w:footer="720" w:gutter="0"/>
      </w:sectPr>
    </w:pPr>
  </w:p>

  <!-- Section 2 content -->
  <w:p>
    <w:r><w:t>Section 2 content (landscape)</w:t></w:r>
  </w:p>

  <!-- Final section properties (direct child of body) -->
  <w:sectPr>
    <w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>
    <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
             w:header="720" w:footer="720" w:gutter="0"/>
    <w:headerReference w:type="default" r:id="rId8"/>
    <w:footerReference w:type="default" r:id="rId9"/>
    <w:cols w:space="720"/>
    <w:docGrid w:linePitch="360"/>
  </w:sectPr>
</w:body>
```

### Section Break Types (`w:type`)

| Value | Description |
|-------|-------------|
| `nextPage` | New section starts on the next page (default if `w:type` is omitted). |
| `continuous` | New section starts on the same page. |
| `evenPage` | New section starts on the next even-numbered page. |
| `oddPage` | New section starts on the next odd-numbered page. |
| `nextColumn` | New section starts in the next column (multi-column layouts). |

### Key Section Property Elements

| Element | Description |
|---------|-------------|
| `w:pgSz` | Page size. `w:w` and `w:h` in twips. `w:orient` = `portrait` or `landscape`. |
| `w:pgMar` | Page margins. `w:top`, `w:bottom`, `w:left`, `w:right`, `w:header`, `w:footer`, `w:gutter` (all in twips). |
| `w:cols` | Column layout. `w:num` = number of columns, `w:space` = spacing between columns. |
| `w:type` | Section break type. |
| `w:pgNumType` | Page numbering. `w:start` = starting page number, `w:fmt` = format. |
| `w:headerReference` | Header reference (see Section 15). |
| `w:footerReference` | Footer reference (see Section 15). |
| `w:titlePg` | Different first page header/footer. |
| `w:pgBorders` | Page borders. |
| `w:lnNumType` | Line numbering. |
| `w:docGrid` | Document grid (line pitch, char pitch for East Asian layout). |
| `w:textDirection` | Text flow direction. |
| `w:bidi` | Right-to-left section. |
| `w:rtlGutter` | Gutter on the right side. |
| `w:formProt` | Form field protection. |
| `w:vAlign` | Vertical alignment of text on page (`top`, `center`, `both`, `bottom`). |

### Common Page Sizes (in twips)

| Paper Size | Width (`w:w`) | Height (`w:h`) |
|-----------|--------------|----------------|
| US Letter | 12240 | 15840 |
| US Legal | 12240 | 20160 |
| A4 | 11906 | 16838 |
| A3 | 16838 | 23811 |

---

## 17. Bookmarks, Hyperlinks, and Fields

### Bookmarks

Bookmarks mark named ranges in the document.

```xml
<w:p>
  <w:bookmarkStart w:id="0" w:name="introduction"/>
  <w:r><w:t>This text is bookmarked.</w:t></w:r>
  <w:bookmarkEnd w:id="0"/>
</w:p>
```

- `w:id` must be unique within the document and match between start and end.
- `w:name` is the human-readable bookmark name.
- Bookmarks can span multiple paragraphs (start and end in different `w:p` elements).
- `_GoBack` is a special bookmark placed by Word at the last editing position.

### Hyperlinks

Hyperlinks can target external URLs or internal bookmarks.

**External hyperlink** (using relationship):

```xml
<w:hyperlink r:id="rId11" w:history="1">
  <w:r>
    <w:rPr>
      <w:rStyle w:val="Hyperlink"/>
    </w:rPr>
    <w:t>Click here</w:t>
  </w:r>
</w:hyperlink>
```

The `r:id` references a relationship with `TargetMode="External"`:

```xml
<Relationship Id="rId11"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
              Target="https://example.com"
              TargetMode="External"/>
```

**Internal hyperlink** (to bookmark):

```xml
<w:hyperlink w:anchor="introduction">
  <w:r>
    <w:rPr>
      <w:rStyle w:val="Hyperlink"/>
    </w:rPr>
    <w:t>Go to Introduction</w:t>
  </w:r>
</w:hyperlink>
```

### Fields

Fields are dynamic content placeholders (page numbers, dates, table of contents, etc.).

**Simple field** (`w:fldSimple`):

```xml
<w:p>
  <w:fldSimple w:instr=" PAGE ">
    <w:r>
      <w:t>1</w:t>              <!-- Cached/last-calculated result -->
    </w:r>
  </w:fldSimple>
</w:p>
```

**Complex field** (`w:fldChar` + `w:instrText`):

```xml
<w:p>
  <!-- Field begin -->
  <w:r>
    <w:fldChar w:fldCharType="begin"/>
  </w:r>
  <!-- Field instruction -->
  <w:r>
    <w:instrText xml:space="preserve"> PAGE </w:instrText>
  </w:r>
  <!-- Separator between instruction and result -->
  <w:r>
    <w:fldChar w:fldCharType="separate"/>
  </w:r>
  <!-- Cached result -->
  <w:r>
    <w:t>1</w:t>
  </w:r>
  <!-- Field end -->
  <w:r>
    <w:fldChar w:fldCharType="end"/>
  </w:r>
</w:p>
```

### Common Field Codes

| Field Code | Description |
|-----------|-------------|
| `PAGE` | Current page number. |
| `NUMPAGES` | Total number of pages. |
| `DATE` | Current date. |
| `TIME` | Current time. |
| `TOC` | Table of contents. |
| `REF bookmarkName` | Cross-reference to a bookmark. |
| `HYPERLINK "url"` | Hyperlink. |
| `SEQ figureName` | Sequence number (for figure/table numbering). |
| `MERGEFIELD fieldName` | Mail merge field. |
| `IF` | Conditional field. |
| `AUTHOR` | Document author. |
| `TITLE` | Document title. |
| `FILENAME` | File name. |

---

## 18. Document Settings

Document settings are stored in `word/settings.xml` under the root element `w:settings`.

### Key Settings Elements

| Element | Description |
|---------|-------------|
| `w:zoom` | Default zoom level (`w:percent`). |
| `w:defaultTabStop` | Default tab stop interval (twips). |
| `w:evenAndOddHeaders` | Enable different headers for even/odd pages. |
| `w:characterSpacingControl` | Character spacing control method. |
| `w:compat` | Compatibility settings (see below). |
| `w:rsids` | Revision save IDs (tracking edit sessions). |
| `w:mathPr` | Math properties. |
| `w:themeFontLang` | Theme font languages. |
| `w:clrSchemeMapping` | Maps theme color names to document color roles. |
| `w:decimalSymbol` | Decimal separator character. |
| `w:listSeparator` | List separator character. |
| `w:trackRevisions` | Enable change tracking. |
| `w:doNotTrackMoves` | Disable move tracking. |
| `w:documentProtection` | Document protection settings (password hash, type). |
| `w:autoHyphenation` | Enable automatic hyphenation. |
| `w:displayBackgroundShape` | Display background shapes. |

### Compatibility Settings (`w:compat`)

The `w:compat` element contains compatibility flags that control layout behavior matching older word processors. The `w:compatSetting` children are the modern approach:

```xml
<w:compat>
  <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
  <w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
  <!-- Legacy compatibility flags -->
  <w:useFELayout/>
  <w:doNotExpandShiftReturn/>
</w:compat>
```

The `compatibilityMode` value corresponds to Word versions:
- `11` = Word 2003
- `12` = Word 2007
- `14` = Word 2010
- `15` = Word 2013+ (current)

---

## 19. Document Properties (Metadata)

### Core Properties (`docProps/core.xml`)

Uses Dublin Core metadata. Namespace: `http://schemas.openxmlformats.org/package/2006/metadata/core-properties`.

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties
    xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:dcterms="http://purl.org/dc/terms/"
    xmlns:dcmitype="http://purl.org/dc/dcmitype/"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Document Title</dc:title>
  <dc:subject>Subject</dc:subject>
  <dc:creator>Author Name</dc:creator>
  <cp:keywords>keyword1, keyword2</cp:keywords>
  <dc:description>Description/comments</dc:description>
  <cp:lastModifiedBy>Editor Name</cp:lastModifiedBy>
  <cp:revision>3</cp:revision>
  <dcterms:created xsi:type="dcterms:W3CDTF">2024-01-15T10:30:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2024-01-20T14:00:00Z</dcterms:modified>
  <cp:category>Category</cp:category>
  <cp:contentStatus>Draft</cp:contentStatus>
</cp:coreProperties>
```

### Extended Properties (`docProps/app.xml`)

Application-specific metadata. Namespace: `http://schemas.openxmlformats.org/officeDocument/2006/extended-properties`.

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Template>Normal.dotm</Template>
  <TotalTime>45</TotalTime>
  <Pages>5</Pages>
  <Words>1250</Words>
  <Characters>7125</Characters>
  <Application>Microsoft Office Word</Application>
  <DocSecurity>0</DocSecurity>
  <Lines>60</Lines>
  <Paragraphs>16</Paragraphs>
  <Company>Company Name</Company>
  <AppVersion>16.0000</AppVersion>
</Properties>
```

### Custom Properties (`docProps/custom.xml`)

User-defined name/value pairs:

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="CustomProp1">
    <vt:lpwstr>Custom Value</vt:lpwstr>
  </property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="CustomBool">
    <vt:bool>true</vt:bool>
  </property>
</Properties>
```

---

## 20. Strict vs. Transitional Conformance

ISO/IEC 29500 defines two conformance classes. Virtually all `.docx` files in the wild use Transitional. A robust parser must handle both.

### Summary of Differences

| Aspect | Transitional | Strict |
|--------|-------------|--------|
| **Spec Part** | Defined by Parts 1 + 4 | Defined by Part 1 only |
| **Namespaces** | `schemas.openxmlformats.org/...` | `purl.oclc.org/ooxml/...` |
| **VML** | Allowed (`v:`, `o:`, `w10:` namespaces) | Not allowed; DrawingML only |
| **Legacy names** | `left`/`right` attributes allowed | Must use `start`/`end` (bidi-aware) |
| **Legacy numbering** | Old numbering elements allowed | Deprecated numbering removed |
| **Non-Unicode** | Legacy character set attributes allowed | Unicode only |
| **Hash algorithms** | Legacy hash mechanisms allowed | Only modern hash algorithms |
| **Compatibility** | Layout compatibility settings allowed | No legacy compatibility settings |
| **Prevalence** | Nearly all real-world documents | Rare; Office 2013+ can write on request |

### Namespace Mapping (Transitional to Strict)

| Component | Transitional | Strict |
|-----------|-------------|--------|
| WordprocessingML | `http://schemas.openxmlformats.org/wordprocessingml/2006/main` | `http://purl.oclc.org/ooxml/wordprocessingml/main` |
| Relationships (markup) | `http://schemas.openxmlformats.org/officeDocument/2006/relationships` | `http://purl.oclc.org/ooxml/officeDocument/relationships` |
| DrawingML | `http://schemas.openxmlformats.org/drawingml/2006/main` | `http://purl.oclc.org/ooxml/drawingml/main` |
| WP Drawing | `http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing` | `http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing` |
| Picture | `http://schemas.openxmlformats.org/drawingml/2006/picture` | `http://purl.oclc.org/ooxml/drawingml/picture` |
| Math | `http://schemas.openxmlformats.org/officeDocument/2006/math` | `http://purl.oclc.org/ooxml/officeDocument/math` |

### Relationship Type URI Differences

The same pattern applies to relationship type URIs:

| Relationship | Transitional | Strict |
|-------------|-------------|--------|
| Office Document | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument` | `http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument` |
| Styles | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles` | `http://purl.oclc.org/ooxml/officeDocument/relationships/styles` |
| (others) | Same pattern: `schemas.openxmlformats.org/...` | Same pattern: `purl.oclc.org/ooxml/...` |

### Attribute Name Differences (Strict Bidi-Aware Naming)

In Strict, directional attribute names are replaced with logical equivalents:

| Transitional | Strict |
|-------------|--------|
| `w:left` (on `w:ind`, `w:pgMar`, etc.) | `w:start` |
| `w:right` (on `w:ind`, `w:pgMar`, etc.) | `w:end` |

### Implementation Strategy

For a parser:
1. Detect whether the document is Strict or Transitional by checking the namespace on the root `w:document` element.
2. Normalize namespace URIs internally so the rest of the parsing logic does not need to branch.
3. When encountering VML (`w:pict`), either convert to DrawingML representation or handle as a separate code path.
4. Map `left`/`right` to `start`/`end` and vice versa depending on the conformance mode.

For a writer:
1. Default to Transitional for maximum compatibility.
2. Support Strict output as an option by remapping namespaces, relationship types, and attribute names.
3. Never emit VML in Strict mode.

---

## 21. Implementer Notes

### Unit Reference Table

| Context | Unit | Description |
|---------|------|-------------|
| Page/margin sizes (`w:pgSz`, `w:pgMar`) | Twips (1/20 pt) | 1 inch = 1440 twips |
| Indentation (`w:ind`) | Twips | |
| Spacing (`w:spacing before/after`) | Twips | |
| Line spacing (`w:spacing line`, `lineRule="auto"`) | 240ths of a line | 240 = single, 480 = double |
| Line spacing (`lineRule="exact"` or `"atLeast"`) | Twips | |
| Font size (`w:sz`) | Half-points | 24 = 12pt |
| Table width (`w:type="dxa"`) | Twips | |
| Table width (`w:type="pct"`) | 50ths of a percent | 5000 = 100% |
| Border width (`w:sz`) | Eighths of a point | 4 = 0.5pt, 12 = 1.5pt |
| DrawingML dimensions | EMU | 1 inch = 914400 EMU |
| Character spacing (`w:spacing` in rPr) | Twips | |

### Common Pitfalls

1. **Whitespace in text**: Always use `xml:space="preserve"` on `w:t` elements when whitespace matters. Omitting it causes Word to strip leading/trailing spaces.

2. **Empty table cells**: Every `w:tc` must contain at least one `w:p`. An empty table cell must still have `<w:p/>`.

3. **Section properties ownership**: The last paragraph in a section "owns" the `w:sectPr` in its `w:pPr`. The body's direct `w:sectPr` child defines the final section. Confusing these causes layout corruption.

4. **Relationship ID uniqueness**: `r:id` values must be unique within each `.rels` file but are scoped to that file. `rId1` in `_rels/.rels` is independent of `rId1` in `word/_rels/document.xml.rels`.

5. **Content type registration**: Every part in the ZIP must have a content type registered in `[Content_Types].xml` either via Default or Override. Missing entries cause load failures.

6. **Numbering indirection**: Paragraphs reference `w:numId` (numbering instance), not `w:abstractNumId` directly. The indirection through `w:num` allows multiple lists to share the same format but track separate counters.

7. **Font resolution order**: When both direct font name (`w:ascii`) and theme reference (`w:asciiTheme`) are present, the direct name takes precedence.

8. **Boolean toggle properties**: Elements like `w:b` (bold) are toggles. `<w:b/>` is equivalent to `<w:b w:val="true"/>`. To explicitly disable, use `<w:b w:val="0"/>` or `<w:b w:val="false"/>`. The `w:val` attribute accepts `true`, `false`, `1`, `0`, `on`, `off`.

9. **Part naming**: Part names in `[Content_Types].xml` overrides must start with `/`. Relationship targets are relative to the source part's directory unless `TargetMode="External"`.

10. **MCE (Markup Compatibility and Extensibility)**: Documents often include `mc:Ignorable` attributes listing namespace prefixes that non-supporting consumers should skip. A parser should respect these or implement full MCE processing per ISO/IEC 29500-3.

### Revision Save IDs (rsid)

Various elements carry `w:rsidR`, `w:rsidRDefault`, `w:rsidRPr`, etc. attributes. These are revision save identifiers that Word uses internally to track editing sessions. They are optional and a parser may ignore them. A writer should either generate them consistently or omit them entirely.

### Markup Compatibility and Extensibility (MCE)

The `mc:Ignorable` attribute on the document root lists namespace prefixes whose elements and attributes can be safely ignored by consumers that do not understand them. The `mc:AlternateContent` / `mc:Choice` / `mc:Fallback` pattern provides versioned content:

```xml
<mc:AlternateContent>
  <mc:Choice Requires="w14">
    <!-- Word 2010+ specific content -->
  </mc:Choice>
  <mc:Fallback>
    <!-- Fallback for older consumers -->
  </mc:Fallback>
</mc:AlternateContent>
```

---

## References

- **ECMA-376** (5th Edition, 2021): https://ecma-international.org/publications-and-standards/standards/ecma-376/
- **ISO/IEC 29500:2016**: https://www.iso.org/standard/71691.html (Part 1), https://www.iso.org/standard/71690.html (Part 2)
- **Library of Congress DOCX Transitional Format Description**: https://www.loc.gov/preservation/digital/formats/fdd/fdd000397.shtml
- **Library of Congress DOCX Strict Format Description**: https://www.loc.gov/preservation/digital/formats/fdd/fdd000400.shtml
- **Microsoft Open XML SDK Documentation**: https://learn.microsoft.com/en-us/office/open-xml/word/structure-of-a-wordprocessingml-document
- **officeopenxml.com** (community reference): http://officeopenxml.com/anatomyofOOXML.php
- **OOXML Info** (spec browser): https://ooxml.info/
