# Open Packaging Conventions (OPC) & Shared Format Specification

> **Reference document for office_oxide implementers.**
> Covers everything shared across DOCX, XLSX, and PPTX, plus legacy and auxiliary formats.

## Table of Contents

1. [Normative References](#1-normative-references)
2. [ZIP Archive Structure](#2-zip-archive-structure)
3. [Content Types](#3-content-types)
4. [Relationships](#4-relationships)
5. [Part Naming Conventions](#5-part-naming-conventions)
6. [Core Properties](#6-core-properties)
7. [App (Extended) Properties](#7-app-extended-properties)
8. [Thumbnails](#8-thumbnails)
9. [Digital Signatures](#9-digital-signatures)
10. [DrawingML Shared Components](#10-drawingml-shared-components)
11. [Measurement Units](#11-measurement-units)
12. [Legacy Binary Formats (CFBF/OLE2)](#12-legacy-binary-formats-cfbfole2)
13. [RTF Format](#13-rtf-format)
14. [CSV/TSV for Spreadsheets](#14-csvtsv-for-spreadsheets)
15. [Namespace Reference](#15-namespace-reference)
16. [Implementation Priorities](#16-implementation-priorities)

---

## 1. Normative References

### Primary Standards

| Standard | Scope | Notes |
|----------|-------|-------|
| **ISO/IEC 29500-1** | Fundamentals and Markup Language Reference | Defines WordprocessingML, SpreadsheetML, PresentationML, DrawingML |
| **ISO/IEC 29500-2** | Open Packaging Conventions (OPC) | The packaging layer: ZIP, content types, relationships, core properties, digital signatures |
| **ISO/IEC 29500-3** | Markup Compatibility and Extensibility | Rules for versioning and forward/backward compatibility |
| **ISO/IEC 29500-4** | Transitional Migration Features | Legacy compatibility elements (VML, etc.) |
| **ECMA-376 5th Edition** | Office Open XML File Formats | Freely downloadable; substantially equivalent to ISO 29500 |

### Supporting Standards

| Standard | Scope |
|----------|-------|
| **APPNOTE (PKWARE) 6.3.x** | ZIP File Format Specification |
| **ISO/IEC 21320-1** | Document Container File (interoperable ZIP subset) |
| **W3C XML Digital Signature** | XML Signature Syntax and Processing |
| **RFC 3986** | URI Generic Syntax (part naming) |
| **Dublin Core (ISO 15836)** | Metadata element set used in core properties |
| **RFC 4180** | CSV format |

### Where to Get the Specs

- ECMA-376 (free): https://ecma-international.org/publications-and-standards/standards/ecma-376/
- ISO 29500 (free from ITTF): https://standards.iso.org/ittf/PubliclyAvailableStandards/
- MS-OE376 (implementation notes): https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/

---

## 2. ZIP Archive Structure

An OOXML file (.docx, .xlsx, .pptx) is a ZIP archive. The OPC specification normatively references PKWARE's ZIP File Format Specification version 6.2.0 with additional constraints.

### ZIP Requirements for OPC

| Requirement | Detail |
|-------------|--------|
| **Compression** | Only `STORED` (method 0) and `DEFLATED` (method 8) are permitted |
| **Encryption** | No ZIP-level encryption allowed (encryption is handled at the OPC layer if needed) |
| **Filenames** | UTF-8 encoded |
| **Spanning** | Multi-disk spanning is not permitted |
| **ZIP64** | Supported for packages larger than 4 GB |
| **Timestamps** | Conforming producers may set timestamps to zero (`0x00002100`); consumers must not rely on them |

### Physical Structure

```
my_document.docx  (ZIP archive)
|
+-- [Content_Types].xml          # REQUIRED: content type declarations
+-- _rels/
|   +-- .rels                    # REQUIRED: package-level relationships
+-- word/                        # (DOCX-specific main content)
|   +-- document.xml
|   +-- _rels/
|   |   +-- document.xml.rels    # Part-level relationships for document.xml
|   +-- styles.xml
|   +-- settings.xml
|   +-- fontTable.xml
|   +-- media/
|   |   +-- image1.png
|   +-- theme/
|       +-- theme1.xml
+-- docProps/
    +-- core.xml                 # Core properties (Dublin Core metadata)
    +-- app.xml                  # Extended/app properties
    +-- thumbnail.jpeg           # Optional thumbnail
```

For XLSX, the main directory is `xl/`; for PPTX, it is `ppt/`.

### Implementation Notes

- Always read the ZIP central directory first for efficient random access.
- Parts may appear in any order within the ZIP; do not assume a specific ordering.
- Interleaved (chunked) parts use the naming pattern `partname/[piece-index].piece` (starting at index 0). This is rare in practice but must be handled.
- Some producers create ZIP entries for directories (zero-length entries ending in `/`). These are not OPC parts and should be ignored.
- The `[Content_Types].xml` file and `_rels/.rels` file are always at the root of the archive.

### Minimum Valid OPC Package

A minimal valid OPC package requires exactly:
1. `[Content_Types].xml` -- even if empty of part declarations
2. At least one part (the main document part, referenced via a package relationship)
3. `_rels/.rels` -- containing the relationship to the main part

---

## 3. Content Types

The `[Content_Types].xml` file is a **mandatory** part at the root of every OPC package. It maps every part in the package to a MIME content type. Consumers must use this file -- not file extensions -- to determine part types.

### XML Namespace

```
http://schemas.openxmlformats.org/package/2006/content-types
```

### Structure

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <!-- Default: maps a file extension to a content type -->
  <Default Extension="rels"
           ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"
           ContentType="application/xml"/>
  <Default Extension="jpeg"
           ContentType="image/jpeg"/>
  <Default Extension="png"
           ContentType="image/png"/>

  <!-- Override: maps a specific part name to a content type -->
  <Override PartName="/word/document.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/core.xml"
            ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml"
            ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
```

### Resolution Algorithm

1. Look for an `<Override>` whose `PartName` matches the part's name (case-insensitive comparison).
2. If no override found, extract the file extension from the part name and look for a `<Default>` whose `Extension` attribute matches (case-insensitive, without the leading dot).
3. If neither matches, the part has no content type and is considered malformed.

### Common Content Types

| Content Type | Usage |
|-------------|-------|
| `application/vnd.openxmlformats-package.relationships+xml` | All `.rels` files |
| `application/vnd.openxmlformats-package.core-properties+xml` | `docProps/core.xml` |
| `application/vnd.openxmlformats-officedocument.extended-properties+xml` | `docProps/app.xml` |
| `application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml` | DOCX main document |
| `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml` | XLSX main workbook |
| `application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml` | PPTX main presentation |
| `application/vnd.openxmlformats-officedocument.theme+xml` | DrawingML theme |
| `application/vnd.openxmlformats-officedocument.drawingml.chart+xml` | DrawingML chart |
| `image/png` | PNG images |
| `image/jpeg` | JPEG images |
| `image/gif` | GIF images |
| `image/tiff` | TIFF images |
| `image/x-emf` | EMF metafiles |
| `image/x-wmf` | WMF metafiles |

### Macro-Enabled Variants

| Extension | Content Type (main part) |
|-----------|------------------------|
| `.docm` | `application/vnd.ms-word.document.macroEnabled.main+xml` |
| `.xlsm` | `application/vnd.ms-excel.sheet.macroEnabled.main+xml` |
| `.pptm` | `application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml` |
| `.dotx` | `application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml` |
| `.xltx` | `application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml` |
| `.potx` | `application/vnd.openxmlformats-officedocument.presentationml.template.main+xml` |

---

## 4. Relationships

Relationships are the **directed graph edges** of an OPC package. They connect a source (the package itself or a part) to a target (another part or an external resource). Relationships decouple navigation from content -- a consumer never needs to parse XML content to discover which parts reference which other parts.

### XML Namespace

```
http://schemas.openxmlformats.org/package/2006/relationships
```

### Relationship XML Structure

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

### Relationship Attributes

| Attribute | Required | Description |
|-----------|----------|-------------|
| `Id` | Yes | Unique identifier within this relationships part. Conventionally `rId1`, `rId2`, etc. Referenced from content XML via `r:id` attributes. |
| `Type` | Yes | URI identifying the semantic type of the relationship. Determines what the target represents. |
| `Target` | Yes | URI of the target resource, resolved relative to the source part's location (not the `.rels` file location). |
| `TargetMode` | No | `Internal` (default) or `External`. External targets are absolute URIs to resources outside the package (e.g., hyperlinks). |

### Relationship Storage Rules

Relationships are stored in dedicated `.rels` files inside `_rels` directories:

| Source | Relationships Part Location |
|--------|-----------------------------|
| Package itself | `/_rels/.rels` |
| `/word/document.xml` | `/word/_rels/document.xml.rels` |
| `/xl/workbook.xml` | `/xl/_rels/workbook.xml.rels` |
| `/ppt/presentation.xml` | `/ppt/_rels/presentation.xml.rels` |
| `/xl/worksheets/sheet1.xml` | `/xl/worksheets/_rels/sheet1.xml.rels` |

**General formula**: For a part at path `/dir/file.ext`, its relationships are stored at `/dir/_rels/file.ext.rels`.

### Target URI Resolution

Target URIs in relationships are resolved relative to the **source part**, not the `.rels` file:

```
Source part:    /word/document.xml
Target value:   media/image1.png
Resolved:       /word/media/image1.png

Source part:    /word/document.xml
Target value:   ../docProps/core.xml
Resolved:       /docProps/core.xml
```

For package relationships (source is the package root), targets are resolved relative to the package root `/`.

### Standard Relationship Types

#### Package-Level (defined by OPC, in `/_rels/.rels`)

| Relationship Type URI | Target | Description |
|-----------------------|--------|-------------|
| `.../officeDocument/2006/relationships/officeDocument` | Main document part | Entry point to the document content |
| `.../package/2006/relationships/metadata/core-properties` | `docProps/core.xml` | Dublin Core metadata |
| `.../officeDocument/2006/relationships/extended-properties` | `docProps/app.xml` | Application metadata |
| `.../package/2006/relationships/metadata/thumbnail` | `docProps/thumbnail.jpeg` | Package thumbnail image |
| `.../package/2006/relationships/digital-signature/origin` | `_xmlsignatures/origin.sigs` | Digital signature origin |

(Where `...` = `http://schemas.openxmlformats.org`)

#### Common Part-Level Relationship Types

| Short Name | Full Type URI Suffix | Description |
|------------|---------------------|-------------|
| `styles` | `.../relationships/styles` | Styles part |
| `theme` | `.../relationships/theme` | Theme part |
| `settings` | `.../relationships/settings` | Settings part |
| `fontTable` | `.../relationships/fontTable` | Font table |
| `image` | `.../relationships/image` | Embedded image |
| `hyperlink` | `.../relationships/hyperlink` | External hyperlink (TargetMode=External) |
| `chart` | `.../relationships/chart` | Embedded chart |
| `oleObject` | `.../relationships/oleObject` | Embedded OLE object |
| `comments` | `.../relationships/comments` | Comments part |
| `header` | `.../relationships/header` | Header part (DOCX) |
| `footer` | `.../relationships/footer` | Footer part (DOCX) |
| `worksheet` | `.../relationships/worksheet` | Worksheet (XLSX) |
| `sharedStrings` | `.../relationships/sharedStrings` | Shared strings table (XLSX) |
| `slide` | `.../relationships/slide` | Slide (PPTX) |
| `slideLayout` | `.../relationships/slideLayout` | Slide layout (PPTX) |
| `slideMaster` | `.../relationships/slideMaster` | Slide master (PPTX) |
| `notesSlide` | `.../relationships/notesSlide` | Notes slide (PPTX) |

### Navigation Algorithm

To open a package:

1. Parse `[Content_Types].xml` to build the content type map.
2. Parse `/_rels/.rels` to find package relationships.
3. Find the relationship with type `…/officeDocument` to locate the main document part.
4. Parse the main document part's `.rels` file to discover its related parts.
5. Recursively follow relationships as needed.

**Critical rule**: Never hardcode part paths. Always discover parts through relationships. Different producers may use different directory structures.

---

## 5. Part Naming Conventions

### Part Name Grammar (from ECMA-376 Part 2, Section 9.1.1)

A valid part name must satisfy all of the following:

| Rule | Constraint |
|------|-----------|
| **Starts with `/`** | Part names are absolute paths rooted at the package root. |
| **No trailing `/`** | A part name must not end with a forward slash. |
| **No empty segments** | Consecutive slashes (`//`) are not permitted. |
| **No dot segments** | Segments `.` and `..` are not permitted in the resolved form. |
| **No percent-encoding** | Part names must not contain percent-encoded characters (unlike general URIs). |
| **No query/fragment** | No `?` query strings or `#` fragment identifiers. |
| **Segment constraints** | Each segment must not end with a dot (`.`). |
| **Case-insensitive** | Part names are compared case-insensitively (per ISO 29500-2:2021). Two names differing only in case refer to the same part. |
| **No backslashes** | Only forward slashes (`/`) are valid separators. |

### Reserved Names

| Name | Purpose |
|------|---------|
| `[Content_Types].xml` | Content type declarations (not technically a "part" but a special stream) |
| `_rels/` | Directory prefix reserved for relationships parts |
| `*.rels` files within `_rels/` | Relationships parts |

No other names are reserved. The bracket characters in `[Content_Types].xml` are literal and required.

### Part Name Examples

```
Valid:
  /word/document.xml
  /xl/worksheets/sheet1.xml
  /ppt/slides/slide1.xml
  /docProps/core.xml
  /word/media/image1.png

Invalid:
  word/document.xml          (no leading /)
  /word/document.xml/        (trailing /)
  /word//document.xml        (empty segment)
  /word/./document.xml       (dot segment)
  /word/document.xml?v=1     (query string)
  /word/my%20doc.xml         (percent encoding)
```

### Conventional Directory Layout

Although not required by the spec, all major producers follow these conventions:

| Format | Main Content Directory | Properties Directory |
|--------|----------------------|---------------------|
| DOCX | `/word/` | `/docProps/` |
| XLSX | `/xl/` | `/docProps/` |
| PPTX | `/ppt/` | `/docProps/` |
| All | `/word/media/`, `/xl/media/`, `/ppt/media/` | Media assets |
| All | `/word/theme/`, `/xl/theme/`, `/ppt/theme/` | Theme definitions |

---

## 6. Core Properties

Core properties provide Dublin Core and OPC-specific metadata about the document. They are stored in `docProps/core.xml` (by convention) and accessed via the package relationship of type `…/metadata/core-properties`.

### XML Namespaces

| Prefix | URI | Defines |
|--------|-----|---------|
| `cp` | `http://schemas.openxmlformats.org/package/2006/metadata/core-properties` | OPC core properties elements |
| `dc` | `http://purl.org/dc/elements/1.1/` | Dublin Core elements |
| `dcterms` | `http://purl.org/dc/terms/` | Dublin Core terms (dates) |
| `dcmitype` | `http://purl.org/dc/dcmitype/` | DCMI Type Vocabulary |
| `xsi` | `http://www.w3.org/2001/XMLSchema-instance` | Schema instance (for `xsi:type`) |

### Content Type

```
application/vnd.openxmlformats-package.core-properties+xml
```

### Complete Element Reference

| Element | Namespace | Type | Description |
|---------|-----------|------|-------------|
| `dc:title` | DC | `xsd:string` | Document title |
| `dc:subject` | DC | `xsd:string` | Document subject/summary |
| `dc:creator` | DC | `xsd:string` | Primary author |
| `dc:description` | DC | `xsd:string` | Document description/comments |
| `cp:keywords` | CP | `xsd:string` | Semicolon or comma-delimited keywords |
| `cp:category` | CP | `xsd:string` | Document category |
| `dc:language` | DC | `xsd:string` | Language code (e.g., `en-US`) |
| `cp:contentType` | CP | `xsd:string` | Document type classification |
| `cp:contentStatus` | CP | `xsd:string` | Status (e.g., "Draft", "Final", "Reviewed") |
| `cp:version` | CP | `xsd:string` | Freeform version string |
| `cp:revision` | CP | `xsd:string` | Revision number (incremented on save by some producers) |
| `dc:identifier` | DC | `xsd:string` | Unique resource identifier |
| `cp:lastModifiedBy` | CP | `xsd:string` | Name of the last person to modify |
| `cp:lastPrinted` | CP | `dcterms:W3CDTF` | Date/time of last print |
| `dcterms:created` | DCTerms | `dcterms:W3CDTF` | Creation date/time |
| `dcterms:modified` | DCTerms | `dcterms:W3CDTF` | Last modification date/time |

### Date Format (W3CDTF)

Dates use the W3C Date-Time Format profile of ISO 8601. The `xsi:type="dcterms:W3CDTF"` attribute is required on date elements.

```
YYYY-MM-DDThh:mm:ssZ          (UTC)
YYYY-MM-DDThh:mm:ss+hh:mm     (with timezone offset)
YYYY-MM-DD                     (date only)
```

### Complete Example

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties
    xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:dcterms="http://purl.org/dc/terms/"
    xmlns:dcmitype="http://purl.org/dc/dcmitype/"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Quarterly Report</dc:title>
  <dc:subject>Q4 2024 Financial Summary</dc:subject>
  <dc:creator>Jane Smith</dc:creator>
  <dc:description>Fourth quarter financial report for FY2024.</dc:description>
  <cp:keywords>finance; quarterly; report; 2024</cp:keywords>
  <cp:category>Report</cp:category>
  <dc:language>en-US</dc:language>
  <cp:contentStatus>Final</cp:contentStatus>
  <cp:version>2.0</cp:version>
  <cp:revision>4</cp:revision>
  <cp:lastModifiedBy>John Doe</cp:lastModifiedBy>
  <cp:lastPrinted xsi:type="dcterms:W3CDTF">2024-12-20T14:30:00Z</cp:lastPrinted>
  <dcterms:created xsi:type="dcterms:W3CDTF">2024-10-01T09:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2024-12-22T16:45:00Z</dcterms:modified>
</cp:coreProperties>
```

### Implementation Notes

- All elements are optional. Producers may include any subset.
- Elements are non-repeatable (at most one of each).
- Order of elements within `<cp:coreProperties>` is not significant.
- Unknown elements should be preserved (round-trip fidelity).

---

## 7. App (Extended) Properties

Extended properties are application-level metadata stored in `docProps/app.xml`. They describe the generating application, document statistics, and other application-specific data.

### XML Namespace

```
http://schemas.openxmlformats.org/officeDocument/2006/extended-properties
```

Conventional prefix: `ap` (or unprefixed as default namespace).

### Content Type

```
application/vnd.openxmlformats-officedocument.extended-properties+xml
```

### Relationship Type

```
http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties
```

### Element Reference

| Element | Type | Description | Example |
|---------|------|-------------|---------|
| `Application` | `xsd:string` | Name of the application that created the document | `Microsoft Office Word` |
| `AppVersion` | `xsd:string` | Version of the application (format: `XX.YYYY`) | `16.0000` |
| `Template` | `xsd:string` | Template name used | `Normal.dotm` |
| `TotalTime` | `xsd:int` | Total editing time in minutes | `120` |
| `Pages` | `xsd:int` | Number of pages (DOCX) | `5` |
| `Words` | `xsd:int` | Word count | `2450` |
| `Characters` | `xsd:int` | Character count (without spaces) | `14200` |
| `CharactersWithSpaces` | `xsd:int` | Character count (with spaces) | `16650` |
| `Lines` | `xsd:int` | Line count | `98` |
| `Paragraphs` | `xsd:int` | Paragraph count | `34` |
| `Slides` | `xsd:int` | Slide count (PPTX) | `12` |
| `Notes` | `xsd:int` | Notes slide count (PPTX) | `3` |
| `HiddenSlides` | `xsd:int` | Hidden slide count (PPTX) | `1` |
| `Company` | `xsd:string` | Company name | `Acme Corp` |
| `Manager` | `xsd:string` | Manager name | `Alice Manager` |
| `DocSecurity` | `xsd:int` | Security level (0=none, 1=password protected, 2=read-only recommended, 4=read-only enforced, 8=locked for annotation) | `0` |
| `ScaleCrop` | `xsd:boolean` | Whether thumbnail display is scaled or cropped | `false` |
| `LinksUpToDate` | `xsd:boolean` | Whether hyperlinks are up to date | `false` |
| `SharedDoc` | `xsd:boolean` | Whether document is shared | `false` |
| `HyperlinksChanged` | `xsd:boolean` | Whether hyperlinks were updated | `false` |
| `HeadingPairs` | Vector | Pairs of group names and counts (e.g., "Theme" / 1, "Slide Titles" / 12) | Complex |
| `TitlesOfParts` | Vector | Names of document parts (sheet names, slide titles) | Complex |

### Vector Elements (HeadingPairs, TitlesOfParts)

These use the `vt:` (variant types) namespace:

```xml
<HeadingPairs>
  <vt:vector size="4" baseType="variant">
    <vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant>
    <vt:variant><vt:i4>3</vt:i4></vt:variant>
    <vt:variant><vt:lpstr>Named Ranges</vt:lpstr></vt:variant>
    <vt:variant><vt:i4>1</vt:i4></vt:variant>
  </vt:vector>
</HeadingPairs>
<TitlesOfParts>
  <vt:vector size="4" baseType="lpstr">
    <vt:lpstr>Sheet1</vt:lpstr>
    <vt:lpstr>Sheet2</vt:lpstr>
    <vt:lpstr>Sheet3</vt:lpstr>
    <vt:lpstr>Print_Area</vt:lpstr>
  </vt:vector>
</TitlesOfParts>
```

Variant types namespace: `http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes`

### Complete Example

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Template>Normal.dotm</Template>
  <TotalTime>45</TotalTime>
  <Pages>3</Pages>
  <Words>1250</Words>
  <Characters>7125</Characters>
  <Application>Microsoft Office Word</Application>
  <DocSecurity>0</DocSecurity>
  <Lines>62</Lines>
  <Paragraphs>17</Paragraphs>
  <ScaleCrop>false</ScaleCrop>
  <Company>Acme Corp</Company>
  <LinksUpToDate>false</LinksUpToDate>
  <CharactersWithSpaces>8358</CharactersWithSpaces>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>16.0000</AppVersion>
</Properties>
```

### Implementation Notes

- All elements are optional and non-repeatable.
- `AppVersion` format is `XX.YYYY` where `XX` is major and `YYYY` is minor (zero-padded).
- When we produce documents, set `Application` to `office_oxide` and `AppVersion` to our version.
- The `DocSecurity` integer is a bitmask.

---

## 8. Thumbnails

OPC packages may include a thumbnail image for preview purposes.

### Relationship Type

```
http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail
```

### Storage

- The thumbnail is typically stored at `docProps/thumbnail.jpeg` or `docProps/thumbnail.emf`.
- It is declared as a package-level relationship in `/_rels/.rels`.
- JPEG and EMF are the most common formats. PNG is also acceptable.
- The content type must be declared in `[Content_Types].xml` (usually via a `<Default>` for the extension).

### Example Relationship

```xml
<Relationship Id="rId4"
              Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail"
              Target="docProps/thumbnail.jpeg"/>
```

### Implementation Notes

- Thumbnails are optional. Many documents do not include them.
- For reading: extract and expose through API if present.
- For writing: generating thumbnails requires rendering capability. Defer to Phase 5 or later.
- Typical thumbnail size: 200x200 pixels or smaller.

---

## 9. Digital Signatures

OPC defines a digital signature framework built on W3C XML Digital Signature (XMLDSig) with OPC-specific extensions.

### Architecture

```
/_xmlsignatures/
    origin.sigs                        # Digital Signature Origin part
    sig1.xml                           # Signature part (one per signature)
    sig2.xml
    _rels/
        origin.sigs.rels               # Relationships from origin to signatures
```

### Key Components

| Component | Description |
|-----------|-------------|
| **Digital Signature Origin Part** | Located at `/_xmlsignatures/origin.sigs` (conventionally). Starting point for locating all signatures. Does not contain signature markup itself. |
| **Signature Parts** | Each contains one XML Digital Signature (`<Signature>` element per W3C XMLDSig). |
| **X.509 Certificates** | Embedded in signature markup via `<KeyInfo>/<X509Data>/<X509Certificate>`. |

### Relationship Types

| Type URI | Purpose |
|----------|---------|
| `.../package/2006/relationships/digital-signature/origin` | Package to signature origin |
| `.../package/2006/relationships/digital-signature/signature` | Origin to individual signature |
| `.../package/2006/relationships/digital-signature/certificate` | Signature to external certificate |

### Signature XML Structure (Simplified)

```xml
<Signature Id="SignatureId" xmlns="http://www.w3.org/2000/09/xmldsig#">
  <SignedInfo>
    <CanonicalizationMethod Algorithm="http://www.w3.org/TR/2001/REC-xml-c14n-20010315"/>
    <SignatureMethod Algorithm="http://www.w3.org/2001/04/xmldsig-more#rsa-sha256"/>
    <!-- References to the package-specific Object -->
    <Reference URI="#idPackageObject"
               Type="http://www.w3.org/2000/09/xmldsig#Object">
      <DigestMethod Algorithm="http://www.w3.org/2001/04/xmlenc#sha256"/>
      <DigestValue>...</DigestValue>
    </Reference>
  </SignedInfo>
  <SignatureValue>...</SignatureValue>
  <KeyInfo>
    <X509Data>
      <X509Certificate>...</X509Certificate>
    </X509Data>
  </KeyInfo>
  <!-- OPC-specific Object: lists signed parts and relationships -->
  <Object Id="idPackageObject">
    <Manifest>
      <!-- Reference to a signed part -->
      <Reference URI="/word/document.xml?ContentType=application/...">
        <DigestMethod Algorithm="http://www.w3.org/2001/04/xmlenc#sha256"/>
        <DigestValue>...</DigestValue>
      </Reference>
      <!-- Reference to signed relationships -->
      <Reference URI="/_rels/.rels?ContentType=application/...">
        <Transforms>
          <Transform Algorithm="http://schemas.openxmlformats.org/package/2006/RelationshipTransform">
            <mdssi:RelationshipReference SourceId="rId1"/>
          </Transform>
          <Transform Algorithm="http://www.w3.org/TR/2001/REC-xml-c14n-20010315"/>
        </Transforms>
        <DigestMethod Algorithm="http://www.w3.org/2001/04/xmlenc#sha256"/>
        <DigestValue>...</DigestValue>
      </Reference>
    </Manifest>
    <SignatureProperties>
      <SignatureProperty Id="idSignatureTime" Target="#SignatureId">
        <mdssi:SignatureTime>
          <mdssi:Format>YYYY-MM-DDThh:mm:ssTZD</mdssi:Format>
          <mdssi:Value>2024-12-20T14:30:00Z</mdssi:Value>
        </mdssi:SignatureTime>
      </SignatureProperty>
    </SignatureProperties>
  </Object>
</Signature>
```

### What Can Be Signed

- Individual parts (by URI + content type)
- Specific relationships (by relationship `Id` or by relationship type)
- Application-specific XML in `<Object>` elements

### Validation Process

1. Recompute the digest of each referenced part/relationship.
2. Compare against stored `<DigestValue>` values.
3. Recompute the signature over `<SignedInfo>`.
4. Compare against stored `<SignatureValue>`.
5. Validate certificate chain (caller responsibility).

### Recommended Algorithms (ISO 29500-2:2021)

| Purpose | Algorithm | URI |
|---------|-----------|-----|
| Digest | SHA-256 | `http://www.w3.org/2001/04/xmlenc#sha256` |
| Signature | RSA-SHA256 | `http://www.w3.org/2001/04/xmldsig-more#rsa-sha256` |
| Canonicalization | C14N | `http://www.w3.org/TR/2001/REC-xml-c14n-20010315` |

SHA-1 is deprecated but must be supported for reading legacy documents.

### Implementation Priority

**Low priority for initial release.** Digital signatures are important for enterprise use cases but not for LLM ingestion or basic document processing. Plan to support:
- Phase 1: Detect and report presence of signatures (read-only).
- Phase 2: Validate signatures.
- Phase 3: Generate signatures.

---

## 10. DrawingML Shared Components

DrawingML is the shared drawing/graphics framework used across all three OOXML formats. It defines themes, colors, fonts, shapes, charts, diagrams, tables, and images.

### Namespaces

| Prefix | URI | Scope |
|--------|-----|-------|
| `a` | `http://schemas.openxmlformats.org/drawingml/2006/main` | Core DrawingML types (colors, fills, text, effects, shapes, tables) |
| `r` | `http://schemas.openxmlformats.org/officeDocument/2006/relationships` | Relationship references |
| `wp` | `http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing` | Drawing positioning in WordprocessingML |
| `xdr` | `http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing` | Drawing positioning in SpreadsheetML |
| `p` | `http://schemas.openxmlformats.org/presentationml/2006/main` | PresentationML (uses DrawingML types directly) |
| `c` | `http://schemas.openxmlformats.org/drawingml/2006/chart` | Charts |
| `dgm` | `http://schemas.openxmlformats.org/drawingml/2006/diagram` | SmartArt/Diagrams |
| `pic` | `http://schemas.openxmlformats.org/drawingml/2006/picture` | Pictures |

### Theme Structure

Every OOXML document includes at least one theme part (e.g., `/word/theme/theme1.xml`). The theme defines the visual identity of the document.

```xml
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
         name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">...</a:clrScheme>
    <a:fontScheme name="Office">...</a:fontScheme>
    <a:fmtScheme name="Office">...</a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults/>
  <a:extraClrSchemeLst/>
</a:theme>
```

#### XML Schema (CT_OfficeStyleSheet)

```xml
<xsd:complexType name="CT_OfficeStyleSheet">
  <xsd:sequence>
    <xsd:element name="themeElements" type="CT_BaseStyles"/>
    <xsd:element name="objectDefaults" type="CT_ObjectStyleDefaults" minOccurs="0"/>
    <xsd:element name="extraClrSchemeLst" type="CT_ColorSchemeList" minOccurs="0"/>
    <xsd:element name="custClrLst" type="CT_CustomColorList" minOccurs="0"/>
    <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0"/>
  </xsd:sequence>
  <xsd:attribute name="name" type="xsd:string" default=""/>
</xsd:complexType>
```

### Color Scheme (a:clrScheme)

Defines 12 named theme colors:

| Element | Semantic Name | Typical Default |
|---------|--------------|-----------------|
| `a:dk1` | Dark 1 (text/background) | Black `000000` |
| `a:lt1` | Light 1 (text/background) | White `FFFFFF` |
| `a:dk2` | Dark 2 | Dark gray `44546A` |
| `a:lt2` | Light 2 | Light gray `E7E6E6` |
| `a:accent1` | Accent 1 | Blue `4472C4` |
| `a:accent2` | Accent 2 | Orange `ED7D31` |
| `a:accent3` | Accent 3 | Gray `A5A5A5` |
| `a:accent4` | Accent 4 | Gold `FFC000` |
| `a:accent5` | Accent 5 | Blue-gray `5B9BD5` |
| `a:accent6` | Accent 6 | Green `70AD47` |
| `a:hlink` | Hyperlink | Blue `0563C1` |
| `a:folHlink` | Followed Hyperlink | Purple `954F72` |

Each color element contains a color value using one of:

```xml
<a:dk1>
  <a:sysClr val="windowText" lastClr="000000"/>  <!-- System color -->
</a:dk1>
<a:accent1>
  <a:srgbClr val="4472C4"/>                       <!-- sRGB hex color -->
</a:accent1>
```

#### Color Reference System

Content XML references theme colors by index rather than literal values:

```xml
<a:solidFill>
  <a:schemeClr val="accent1"/>    <!-- Resolved via theme -->
</a:solidFill>
```

Scheme color values: `bg1`, `tx1`, `bg2`, `tx2`, `accent1`..`accent6`, `hlink`, `folHlink`, `dk1`, `lt1`, `dk2`, `lt2`.

Color modifications can be applied inline:

```xml
<a:schemeClr val="accent1">
  <a:lumMod val="75000"/>    <!-- 75% luminance -->
  <a:lumOff val="25000"/>    <!-- +25% luminance offset -->
</a:schemeClr>
```

### Font Scheme (a:fontScheme)

Defines two font families -- **Major** (headings) and **Minor** (body):

```xml
<a:fontScheme name="Office">
  <a:majorFont>
    <a:latin typeface="Calibri Light" panose="020F0302020204030204"/>
    <a:ea typeface=""/>        <!-- East Asian: empty = use default -->
    <a:cs typeface=""/>        <!-- Complex Script: empty = use default -->
    <!-- Optional per-language overrides -->
    <a:font script="Jpan" typeface="Yu Gothic Light"/>
    <a:font script="Hang" typeface="Malgun Gothic"/>
  </a:majorFont>
  <a:minorFont>
    <a:latin typeface="Calibri" panose="020F0502020204030204"/>
    <a:ea typeface=""/>
    <a:cs typeface=""/>
    <a:font script="Jpan" typeface="Yu Gothic"/>
    <a:font script="Hang" typeface="Malgun Gothic"/>
  </a:minorFont>
</a:fontScheme>
```

Content XML references theme fonts via:

```xml
<a:latin typeface="+mj-lt"/>   <!-- Major Latin font from theme -->
<a:latin typeface="+mn-lt"/>   <!-- Minor Latin font from theme -->
```

The `+mj-` prefix means major font; `+mn-` means minor font. Suffixes: `-lt` (Latin), `-ea` (East Asian), `-cs` (Complex Script).

### Format Scheme (a:fmtScheme)

Defines four style lists for consistent visual formatting:

| Element | Purpose | Contains |
|---------|---------|----------|
| `a:fillStyleLst` | Fill styles (3 entries: subtle, moderate, intense) | Solid, gradient, pattern fills |
| `a:lnStyleLst` | Line/border styles (3 entries) | Width, dash, join styles |
| `a:effectStyleLst` | Shadow/glow/reflection effects (3 entries) | Effect chains |
| `a:bgFillStyleLst` | Background fill styles (3 entries) | Background-specific fills |

Content references these by 1-based index: style index 1 = subtle, 2 = moderate, 3 = intense.

### Shared Shape Types

DrawingML defines preset geometry shapes via `a:prstGeom`:

```xml
<a:prstGeom prst="rect">       <!-- Rectangle -->
  <a:avLst/>                    <!-- Adjustment values (empty = defaults) -->
</a:prstGeom>
```

Common preset values: `rect`, `roundRect`, `ellipse`, `triangle`, `diamond`, `star5`, `rightArrow`, `line`, `flowChartProcess`, etc. There are 187 preset geometries in the specification.

Custom geometry is defined via `a:custGeom` with explicit path commands (moveTo, lineTo, arcTo, cubicBezTo, close).

### Embedded Images

Images are stored as binary parts in `media/` directories and referenced via relationships:

```xml
<!-- In content XML -->
<a:blip r:embed="rId5"/>   <!-- r:embed references a relationship Id -->

<!-- In the .rels file -->
<Relationship Id="rId5" Type=".../image" Target="../media/image1.png"/>
```

Supported image formats: PNG, JPEG, GIF, TIFF, BMP, EMF, WMF, SVG (Office 2016+).

### Charts

Charts use the `c:` namespace and are stored as separate parts:

```
/word/charts/chart1.xml     (or /xl/charts/, /ppt/charts/)
```

Chart types include: `c:barChart`, `c:lineChart`, `c:pieChart`, `c:scatterChart`, `c:areaChart`, `c:doughnutChart`, `c:radarChart`, `c:surfaceChart`, `c:bubbleChart`, `c:stockChart`.

### Content Type

```
application/vnd.openxmlformats-officedocument.theme+xml
```

---

## 11. Measurement Units

OOXML uses multiple measurement systems depending on context.

### English Metric Units (EMU)

The primary coordinate system for DrawingML. EMU is an integer-based unit designed to convert exactly between inches and centimeters without floating-point rounding.

| Conversion | Value |
|------------|-------|
| 1 inch | 914,400 EMU |
| 1 cm | 360,000 EMU |
| 1 point (1/72 inch) | 12,700 EMU |
| 1 pixel (at 96 DPI) | 9,525 EMU |

**Why EMU?** `914400 = LCM(100, 254) * 72 / 2 = 12700 * 72`. This integer ensures exact conversion between metric, imperial, and typographic units.

EMU is used for: shape positions, shape sizes, image dimensions, margins in DrawingML, and all coordinate attributes in `wp:`, `xdr:`, and drawing anchors.

### Half-Points

Used in WordprocessingML for font sizes:

```xml
<w:sz w:val="24"/>   <!-- 24 half-points = 12pt font -->
```

| Value | Font Size |
|-------|-----------|
| 20 | 10pt |
| 24 | 12pt |
| 28 | 14pt |
| 48 | 24pt |

### Twentieths of a Point (Twips)

Used in WordprocessingML for page dimensions, margins, spacing, and indentation:

```xml
<w:pgSz w:w="12240" w:h="15840"/>  <!-- US Letter: 8.5" x 11" -->
```

| Conversion | Value |
|------------|-------|
| 1 inch | 1,440 twips |
| 1 cm | 567 twips (approx) |
| 1 point | 20 twips |

### Hundredths of a Character / Fiftieths of a Line

Used in SpreadsheetML for column widths and row heights. Column width is measured in units of the width of the zero character (`0`) in the default font, expressed in 1/256ths of that width.

### Percentages

Many attributes use percentages scaled by a factor:

| Scale | Usage |
|-------|-------|
| 1/1000th of a percent (100000 = 100%) | Color modifications (lumMod, satMod), transparency |
| 1/100th of a percent (10000 = 100%) | Some spacing values |
| Direct percentage (100 = 100%) | Zoom levels, scale |

---

## 12. Legacy Binary Formats (CFBF/OLE2)

### Overview

Before OOXML (Office 2007+), Microsoft Office used the **Compound File Binary Format (CFBF)**, also known as **OLE2 Structured Storage** or **COM Structured Storage**. These are the `.doc`, `.xls`, and `.ppt` formats.

### File Structure

CFBF implements a FAT-like filesystem within a single file:

```
+----------------------------------+
| Header (512 bytes)               |  Magic: D0 CF 11 E0 A1 B1 1A E1
+----------------------------------+
| DIFAT (Double-Indirect FAT)      |  Locates FAT sectors
+----------------------------------+
| FAT (File Allocation Table)      |  Sector chains (like FAT32)
+----------------------------------+
| MiniFAT                          |  Chain table for small streams
+----------------------------------+
| Directory Entries                 |  Red-black tree of storage/stream entries
+----------------------------------+
| Stream Data (sectors)            |  Actual file content
+----------------------------------+
| Mini Stream                      |  Small streams (<4096 bytes) packed together
+----------------------------------+
```

### Key Parameters

| Parameter | Version 3 | Version 4 |
|-----------|-----------|-----------|
| Magic bytes | `D0 CF 11 E0 A1 B1 1A E1` | Same |
| Sector size | 512 bytes | 4096 bytes |
| Mini sector size | 64 bytes | 64 bytes |
| Mini stream cutoff | 4096 bytes | 4096 bytes |
| Max directory entries | No limit | No limit |
| Max file size | ~2 GB | ~2 GB (practical) |

### Directory Entry Structure (128 bytes each)

| Field | Size | Description |
|-------|------|-------------|
| Name | 64 bytes | UTF-16LE entry name (max 31 chars + null) |
| Name Length | 2 bytes | In bytes, including null terminator |
| Object Type | 1 byte | 0=unknown, 1=storage, 2=stream, 5=root |
| Color Flag | 1 byte | 0=red, 1=black (red-black tree) |
| Left Sibling ID | 4 bytes | Directory entry index |
| Right Sibling ID | 4 bytes | Directory entry index |
| Child ID | 4 bytes | First child directory entry |
| CLSID | 16 bytes | Class identifier |
| State Bits | 4 bytes | User-defined flags |
| Creation Time | 8 bytes | FILETIME |
| Modified Time | 8 bytes | FILETIME |
| Starting Sector | 4 bytes | First sector of stream data |
| Stream Size | 8 bytes | Total stream size in bytes |

### What Is Inside Each Format

| Format | Root Streams | Key Content |
|--------|-------------|-------------|
| `.doc` | `WordDocument`, `1Table`/`0Table`, `Data` | Binary word processor stream with FIB (File Information Block) |
| `.xls` | `Workbook` (or `Book`) | BIFF8 record stream (opcode + length + data) |
| `.ppt` | `PowerPoint Document`, `Current User` | Record-based binary stream |
| All | `\x05SummaryInformation` | OLE property set (title, author, etc.) |
| All | `\x05DocumentSummaryInformation` | Extended OLE property set |

### Should We Support CFBF?

**Recommendation: Yes, read-only support, phased.**

**Arguments for:**
- Legacy `.doc`/`.xls`/`.ppt` files remain extremely common in enterprise document archives.
- LLM ingestion pipelines encounter these formats frequently.
- Text extraction from legacy formats is a major use case.
- Competitors (Apache POI, Aspose) support them.

**Arguments against:**
- Significantly more complex than OOXML (no XML, custom binary formats per application).
- Write support is extremely difficult and error-prone.
- Microsoft themselves have moved away from these formats.

**Proposed approach:**
1. Implement a CFBF container reader in `office_core` (parse sectors, FAT, directory).
2. Implement read-only text extraction for `.doc` (Word Binary Format).
3. Implement read-only text extraction for `.xls` (BIFF8).
4. Implement read-only text extraction for `.ppt` (PowerPoint Binary).
5. No write support for binary formats. Convert to OOXML instead.

### Specification Reference

- **[MS-CFB]**: Compound File Binary File Format -- https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/
- **[MS-DOC]**: Word Binary File Format -- https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/
- **[MS-XLS]**: Excel Binary File Format -- https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/
- **[MS-PPT]**: PowerPoint Binary File Format -- https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ppt/

---

## 13. RTF Format

### Overview

Rich Text Format (RTF) is a Microsoft-developed document format that uses 7-bit ASCII text with control words for formatting. The final version is **RTF 1.9.1** (March 2008, covering Word 2007 features).

### File Structure

```
{\rtf1\ansi\deff0             RTF version 1, ANSI charset, default font 0
  {\fonttbl                   Font table (required)
    {\f0\froman Times New Roman;}
    {\f1\fswiss Arial;}
  }
  {\colortbl ;                Color table (optional, first entry is "auto")
    \red255\green0\blue0;     Color 1: red
    \red0\green0\blue255;     Color 2: blue
  }
  {\stylesheet                Style table (optional)
    {\s0 Normal;}
    {\s1\b Heading 1;}
  }
  {\info                      Document info (optional)
    {\title My Document}
    {\author Jane Smith}
    {\creatim\yr2024\mo12\dy20\hr14\min30}
  }
  \pard\plain                 Default paragraph, reset formatting
  Hello, {\b bold} and        "Hello, bold and"
  {\i italic} world.\par      "italic world."
}
```

### Control Word Syntax

```
\keyword[N]         N is an optional signed 16-bit integer (-32768 to 32767)
\keyword-N          Negative values (e.g., \li-720 for negative indent)
\'HH                8-bit character by hex code (e.g., \'e9 for e-acute)
\uN?                Unicode character N, followed by ANSI fallback character(s)
\\                  Literal backslash
\{                  Literal open brace
\}                  Literal close brace
\~                  Non-breaking space
\-                  Optional hyphen
\_                  Non-breaking hyphen
```

### Key Control Words

| Category | Control Words |
|----------|--------------|
| **Character set** | `\ansi`, `\mac`, `\pc`, `\pca`, `\ansicpg1252` |
| **Font** | `\f0` (select font), `\fs24` (font size in half-points), `\fcharset0` |
| **Character formatting** | `\b` (bold), `\i` (italic), `\ul` (underline), `\strike`, `\super`, `\sub` |
| **Color** | `\cf1` (foreground color index), `\cb2` (background color index), `\highlight3` |
| **Paragraph** | `\pard` (reset paragraph), `\par` (new paragraph), `\line` (line break) |
| **Alignment** | `\ql` (left), `\qr` (right), `\qc` (center), `\qj` (justify) |
| **Indentation** | `\li720` (left indent, twips), `\ri720` (right), `\fi-360` (first line/hanging) |
| **Spacing** | `\sb120` (space before), `\sa120` (space after), `\sl240` (line spacing) |
| **Tables** | `\trowd`, `\cellx`, `\cell`, `\row`, `\intbl` |
| **Images** | `\pict`, `\pngblip`, `\jpegblip`, `\emfblip`, `\wmetafile` |
| **Unicode** | `\uc1` (1-byte ANSI fallback), `\u8364?` (Euro sign, `?` as fallback) |
| **Sections** | `\sect`, `\sectd`, `\sbknone`, `\sbkpage` |
| **Page** | `\paperw12240` (width twips), `\paperh15840` (height), `\margl1800` (left margin) |

### Unicode Handling

RTF uses `\uN` for Unicode characters where N is a signed 16-bit value:

```
\u8364?     Euro sign (U+20AC = 8364 decimal), ? is ANSI fallback
\u-4064?    Same as \u61472 (values > 32767 expressed as negative: 61472 - 65536 = -4064)
```

The `\uc` keyword specifies how many bytes of ANSI fallback follow each `\u`:

```
\uc1\u8364?          1 fallback byte
\uc2\u12354AB        2 fallback bytes
\uc0\u8364           No fallback bytes
```

### Should We Support RTF?

**Recommendation: Yes, read-only, moderate priority.**

- RTF is common in legacy systems, legal documents, and clipboard interchange.
- Text extraction is straightforward (parse control words, ignore unknown ones).
- The format is well-documented and relatively simple compared to CFBF.
- Write support is feasible but lower priority than OOXML write.

### Specification Reference

- RTF 1.9.1 Specification (Microsoft, 2008): https://learn.microsoft.com/en-us/archive/blogs/microsoft_office_word/

---

## 14. CSV/TSV for Spreadsheets

### CSV (RFC 4180)

MIME type: `text/csv`

#### Rules

1. Each record is on a separate line, terminated by CRLF (`\r\n`).
2. The last record may or may not end with CRLF.
3. An optional header line may appear as the first record.
4. Each record contains the same number of fields, separated by commas.
5. Fields may be enclosed in double quotes (`"`). Fields containing commas, double quotes, or line breaks must be quoted.
6. A double quote within a quoted field is escaped by doubling it: `""`.

#### Examples

```csv
Name,Age,City
"Smith, John",30,"New York"
Jane Doe,25,Chicago
"He said ""hello""",40,Boston
"Multi
line field",50,Denver
```

#### Practical Considerations

- **Encoding**: RFC 4180 does not specify encoding. In practice: UTF-8 (preferred), UTF-8 with BOM (`EF BB BF`), Windows-1252, or locale-dependent.
- **Delimiter detection**: Some producers use semicolons (`;`) instead of commas (common in European locales where comma is the decimal separator). Detection heuristics are needed.
- **Line endings**: Accept `\r\n`, `\n`, or `\r`.
- **Empty fields**: `,,` represents an empty field. `""` also represents an empty field.
- **Numeric types**: All fields are strings in CSV. Type inference is the consumer's responsibility.

### TSV (Tab-Separated Values)

MIME type: `text/tab-separated-values`

#### Rules

- Same as CSV but delimiter is TAB (`\t`) instead of comma.
- No formal quoting mechanism in the original spec. In practice, double-quote quoting (as in CSV) is often supported.
- Tabs and newlines within field values are typically not allowed (or must be escaped).

### Implementation for office_oxide

For XLSX export/import, support:

| Feature | Priority |
|---------|----------|
| CSV read (text extraction from XLSX to CSV) | High |
| CSV write (XLSX to CSV conversion) | High |
| TSV read/write | Medium |
| Delimiter auto-detection | Medium |
| Encoding detection (BOM, heuristic) | Medium |
| Type inference (numbers, dates, booleans) | Medium |
| Large file streaming (no full load into memory) | High |

---

## 15. Namespace Reference

### Package-Level Namespaces (OPC)

| Prefix | URI |
|--------|-----|
| (content types) | `http://schemas.openxmlformats.org/package/2006/content-types` |
| (relationships) | `http://schemas.openxmlformats.org/package/2006/relationships` |
| `cp` | `http://schemas.openxmlformats.org/package/2006/metadata/core-properties` |
| `dc` | `http://purl.org/dc/elements/1.1/` |
| `dcterms` | `http://purl.org/dc/terms/` |
| `dcmitype` | `http://purl.org/dc/dcmitype/` |
| `xsi` | `http://www.w3.org/2001/XMLSchema-instance` |

### Office Document Namespaces (Shared)

| Prefix | URI | Scope |
|--------|-----|-------|
| `r` | `http://schemas.openxmlformats.org/officeDocument/2006/relationships` | Relationship references in content XML |
| `ap` | `http://schemas.openxmlformats.org/officeDocument/2006/extended-properties` | App properties |
| `vt` | `http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes` | Variant types in properties |
| `m` | `http://schemas.openxmlformats.org/officeDocument/2006/math` | Office Math (OMML) |

### DrawingML Namespaces

| Prefix | URI | Scope |
|--------|-----|-------|
| `a` | `http://schemas.openxmlformats.org/drawingml/2006/main` | Core DrawingML |
| `wp` | `http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing` | Drawing anchors in DOCX |
| `xdr` | `http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing` | Drawing anchors in XLSX |
| `c` | `http://schemas.openxmlformats.org/drawingml/2006/chart` | Charts |
| `dgm` | `http://schemas.openxmlformats.org/drawingml/2006/diagram` | SmartArt diagrams |
| `pic` | `http://schemas.openxmlformats.org/drawingml/2006/picture` | Pictures |

### Format-Specific Namespaces

| Prefix | URI | Scope |
|--------|-----|-------|
| `w` | `http://schemas.openxmlformats.org/wordprocessingml/2006/main` | WordprocessingML (DOCX) |
| `x` | `http://schemas.openxmlformats.org/spreadsheetml/2006/main` | SpreadsheetML (XLSX) |
| `p` | `http://schemas.openxmlformats.org/presentationml/2006/main` | PresentationML (PPTX) |

### Microsoft Extension Namespaces (Common)

| Prefix | URI | Scope |
|--------|-----|-------|
| `mc` | `http://schemas.openxmlformats.org/markup-compatibility/2006` | Markup compatibility |
| `w14` | `http://schemas.microsoft.com/office/word/2010/wordml` | Word 2010 extensions |
| `w15` | `http://schemas.microsoft.com/office/word/2012/wordml` | Word 2013 extensions |
| `x14` | `http://schemas.microsoft.com/office/spreadsheetml/2009/9/main` | Excel 2010 extensions |
| `a14` | `http://schemas.microsoft.com/office/drawing/2010/main` | DrawingML 2010 extensions |

These extension namespaces are used with `mc:Ignorable` for forward compatibility. Consumers that do not understand them should ignore them gracefully (per ISO 29500-3).

---

## 16. Implementation Priorities

Based on the office_oxide mission (LLM ingestion, speed, text extraction first):

### Phase 1: office_core Foundation

| Component | Priority | Notes |
|-----------|----------|-------|
| ZIP reader (DEFLATE + STORED) | **P0** | Use `zip` crate or custom implementation for streaming |
| `[Content_Types].xml` parser | **P0** | Simple XML, build content type map |
| Relationships parser (`.rels`) | **P0** | Core navigation mechanism |
| Part name validation | **P0** | Enforce OPC naming rules |
| Core properties reader | **P1** | Metadata extraction for LLM |
| App properties reader | **P1** | Application metadata |
| Theme parser (DrawingML) | **P1** | Needed for color/font resolution |
| ZIP writer (for OOXML creation) | **P1** | Required for document creation |
| Content types writer | **P1** | Required for document creation |
| Relationships writer | **P1** | Required for document creation |

### Phase 2: Read Support

| Component | Priority | Notes |
|-----------|----------|-------|
| DrawingML color resolution | **P1** | Theme colors, system colors, sRGB |
| DrawingML font resolution | **P1** | Theme fonts, fallback chains |
| Image extraction | **P1** | Follow relationships to media parts |
| Chart data extraction | **P2** | Extract underlying data from charts |
| CFBF container reader | **P2** | For legacy format support |
| RTF text extraction | **P2** | Straightforward parsing |

### Phase 3: Advanced Features

| Component | Priority | Notes |
|-----------|----------|-------|
| Digital signature detection | **P3** | Report presence |
| Digital signature validation | **P3** | Verify integrity |
| Custom properties reader | **P3** | User-defined properties |
| Thumbnail extraction | **P3** | Low value for LLM use cases |

### Key Design Decisions

1. **Streaming vs. DOM**: Prefer streaming/SAX-style XML parsing for large files. Only build DOM trees for small parts (relationships, properties, themes).

2. **Zero-copy ZIP**: Use memory-mapped I/O and parse ZIP central directory first. Decompress parts on demand, not eagerly.

3. **Relationship-driven navigation**: Always discover parts through relationships. Never hardcode paths like `/word/document.xml`.

4. **Content type checking**: Verify content types match expectations. Reject parts with unexpected content types for security.

5. **Namespace-aware XML parsing**: Use namespace URIs, not prefixes, for element matching. Prefixes are arbitrary and vary between producers.

6. **Round-trip fidelity**: When reading and re-writing, preserve unknown elements and attributes. This is critical for compatibility with features we do not yet support.

7. **Error tolerance**: Real-world documents frequently violate the spec. Accept common violations (wrong case, missing optional attributes, extra whitespace) while rejecting truly malformed content.

---

## Appendix A: Typical Package Relationships (/_rels/.rels)

### DOCX

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
</Relationships>
```

### XLSX

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

### PPTX

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="ppt/presentation.xml"/>
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

---

## Appendix B: Content Types Quick Reference

### Minimal [Content_Types].xml for DOCX

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/core.xml"
            ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml"
            ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
```

### Minimal [Content_Types].xml for XLSX

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml"
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml"
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml"
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml"
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/docProps/core.xml"
            ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml"
            ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
```

---

## Appendix C: CFBF Magic Bytes and Detection

```
Offset  Bytes                          Meaning
------  ----------------------------   -------
0x00    D0 CF 11 E0 A1 B1 1A E1       CFBF magic signature
0x1A    03 00  or  04 00               Minor version (3 or 4)
0x1C    03 00  or  04 00               Major version (3=512-byte sectors, 4=4096-byte sectors)
0x1E    FE FF                          Byte order (little-endian)
0x20    09 00  or  0C 00               Sector size power (9=512, 12=4096)
0x22    06 00                          Mini sector size power (6=64)
```

### Format Detection Strategy

```
if file[0..8] == [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]:
    -> CFBF (OLE2): .doc, .xls, .ppt, or other OLE format
    -> Parse directory to determine specific format

if file[0..4] == [0x50, 0x4B, 0x03, 0x04]:  // "PK\x03\x04"
    -> ZIP archive: .docx, .xlsx, .pptx, or other OPC format
    -> Check [Content_Types].xml for specific format

if file[0..5] == b"{\\rtf":
    -> RTF document

if file is valid UTF-8/ASCII with comma/tab delimiters:
    -> CSV/TSV (heuristic detection)
```
