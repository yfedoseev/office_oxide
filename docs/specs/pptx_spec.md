# PPTX (Office Open XML PresentationML) Technical Specification Reference

> **Reference document for office_oxide implementers.**
> Covers everything specific to the PPTX format. For shared OPC concerns (ZIP, content types basics, core/app properties, digital signatures), see `opc_shared_spec.md`.

## Table of Contents

1. [Official Specification References](#1-official-specification-references)
2. [Package Structure (ZIP Layout)](#2-package-structure-zip-layout)
3. [XML Namespaces](#3-xml-namespaces)
4. [Content Types](#4-content-types)
5. [Relationship System](#5-relationship-system)
6. [Presentation Part (presentation.xml)](#6-presentation-part-presentationxml)
7. [Slide Structure (slide.xml)](#7-slide-structure-slidexml)
8. [Shape Types](#8-shape-types)
9. [Text Body System](#9-text-body-system)
10. [Slide Masters and Slide Layouts](#10-slide-masters-and-slide-layouts)
11. [Themes (DrawingML)](#11-themes-drawingml)
12. [Transitions and Animations](#12-transitions-and-animations)
13. [Notes Slides and Handout Masters](#13-notes-slides-and-handout-masters)
14. [Embedded Media](#14-embedded-media)
15. [Tables in Slides](#15-tables-in-slides)
16. [Comments](#16-comments)
17. [Strict vs. Transitional Conformance](#17-strict-vs-transitional-conformance)
18. [Implementation Notes and Common Pitfalls](#18-implementation-notes-and-common-pitfalls)

---

## 1. Official Specification References

### Primary Standards

| Standard | Description |
|----------|-------------|
| **ECMA-376** | Office Open XML File Formats. Published by Ecma International. 5 editions (2006-2021). |
| **ISO/IEC 29500:2008-2016** | International standard derived from ECMA-376. 4 parts. |

### Standard Parts

| Part | Title | Scope |
|------|-------|-------|
| **Part 1** | Fundamentals and Markup Language Reference | Element/attribute definitions for PresentationML, DrawingML, etc. |
| **Part 2** | Open Packaging Conventions (OPC) | ZIP container, relationships, content types. |
| **Part 3** | Markup Compatibility and Extensibility | Versioning, ignorable namespaces. |
| **Part 4** | Transitional Migration Features | Legacy compatibility (VML, etc.). |

### Conformance Classes

The standard defines two conformance levels:

- **Strict** -- uses `http://purl.oclc.org/ooxml/...` namespaces. No legacy features.
- **Transitional** -- uses `http://schemas.openxmlformats.org/...` namespaces. Permits legacy markup (VML, deprecated attributes). This is what virtually all real-world `.pptx` files use.

### Key Online References

- ECMA-376 downloads: `https://ecma-international.org/publications-and-standards/standards/ecma-376/`
- ISO/IEC 29500: `https://standards.iso.org/ittf/PubliclyAvailableStandards/`
- Microsoft Open XML SDK docs: `https://learn.microsoft.com/en-us/office/open-xml/`
- Unofficial reference site: `http://officeopenxml.com/`

### Microsoft Extension Specifications

- **[MS-PPTX]**: PowerPoint (.pptx) Extensions to Office Open XML PresentationML
- **[MS-OE376]**: Office Implementation Information for ECMA-376
- **[MS-OI29500]**: Office Implementation Information for ISO/IEC 29500

URL: https://learn.microsoft.com/en-us/openspecs/office_standards/

### Key Schema Files (from ECMA-376)

- `pml.xsd` -- PresentationML main schema
- `dml-main.xsd` -- DrawingML core
- `dml-chart.xsd` -- DrawingML charts
- `dml-diagram.xsd` -- DrawingML SmartArt/diagrams
- `dml-picture.xsd` -- DrawingML pictures
- `shared-commonSimpleTypes.xsd` -- Shared simple types (ST_Coordinate, ST_Percentage, etc.)

### MIME Types

| Extension | MIME Type |
|-----------|----------|
| `.pptx` | `application/vnd.openxmlformats-officedocument.presentationml.presentation` |
| `.pptm` | `application/vnd.ms-powerpoint.presentation.macroEnabled.12` |
| `.potx` | `application/vnd.openxmlformats-officedocument.presentationml.template` |
| `.potm` | `application/vnd.ms-powerpoint.template.macroEnabled.12` |
| `.ppsx` | `application/vnd.openxmlformats-officedocument.presentationml.slideshow` |
| `.ppsm` | `application/vnd.ms-powerpoint.slideshow.macroEnabled.12` |

---

## 2. Package Structure (ZIP Layout)

A `.pptx` file is a **ZIP archive** conforming to the Open Packaging Conventions (OPC, ISO/IEC 29500 Part 2). Renaming `.pptx` to `.zip` allows direct extraction.

### Typical Directory Layout

```
[Content_Types].xml              # REQUIRED: content type mappings
_rels/
  .rels                          # REQUIRED: package-level relationships
docProps/
  app.xml                        # Application properties (optional)
  core.xml                       # Dublin Core metadata (optional)
  thumbnail.jpeg                 # Thumbnail image (optional)
ppt/
  presentation.xml               # REQUIRED: main presentation part
  presProps.xml                   # Presentation properties
  viewProps.xml                   # View properties
  tableStyles.xml                # Table style definitions
  commentAuthors.xml             # Comment author list (if comments exist)
  _rels/
    presentation.xml.rels        # Relationships for presentation.xml
  slides/
    slide1.xml                   # Slide content
    slide2.xml
    _rels/
      slide1.xml.rels            # Relationships for slide1
      slide2.xml.rels
  slideLayouts/
    slideLayout1.xml             # Layout definitions
    slideLayout2.xml
    _rels/
      slideLayout1.xml.rels
      slideLayout2.xml.rels
  slideMasters/
    slideMaster1.xml             # Master slide definitions
    _rels/
      slideMaster1.xml.rels
  notesMasters/
    notesMaster1.xml             # Notes master (optional)
    _rels/
      notesMaster1.xml.rels
  notesSlides/
    notesSlide1.xml              # Per-slide notes (optional)
    _rels/
      notesSlide1.xml.rels
  handoutMasters/
    handoutMaster1.xml           # Handout master (optional)
  comments/
    comment1.xml                 # Per-slide comments (optional)
  theme/
    theme1.xml                   # Theme definition
  media/
    image1.png                   # Embedded media files
    audio1.wav
    video1.mp4
  embeddings/
    oleObject1.bin               # OLE embedded objects
  charts/
    chart1.xml                   # Embedded chart parts
  diagrams/
    data1.xml                    # SmartArt diagram data
  tags/
    tag1.xml                     # Custom tags (key-value metadata)
```

### Part Numbering Convention

Parts are numbered sequentially: `slide1.xml`, `slide2.xml`, etc. However, the actual slide order is determined by the `<p:sldIdLst>` element in `presentation.xml`, NOT by file numbering. After deletions and reorderings, file numbers may have gaps (e.g., slide1.xml, slide3.xml, slide7.xml).

### Implementation Notes

- The ZIP must use Deflate or Store compression only.
- File paths inside the ZIP are case-insensitive per OPC.
- The `[Content_Types].xml` file must be at the archive root (not in any folder).
- Relationship files (`.rels`) are always in a `_rels` subfolder relative to the part they describe.
- Some producers create ZIP entries for directories (zero-length entries ending in `/`). These are not OPC parts and should be ignored.
- Parts may appear in any order within the ZIP; do not assume a specific ordering.

---

## 3. XML Namespaces

### Primary Namespaces (Transitional)

| Prefix | URI | Usage |
|--------|-----|-------|
| `p` | `http://schemas.openxmlformats.org/presentationml/2006/main` | PresentationML elements |
| `a` | `http://schemas.openxmlformats.org/drawingml/2006/main` | DrawingML elements |
| `r` | `http://schemas.openxmlformats.org/officeDocument/2006/relationships` | Relationship references |

### Microsoft Extension Namespaces

| Prefix | URI | Usage |
|--------|-----|-------|
| `p14` | `http://schemas.microsoft.com/office/powerpoint/2010/main` | PowerPoint 2010 extensions |
| `p15` | `http://schemas.microsoft.com/office/powerpoint/2012/main` | PowerPoint 2013 extensions |
| `p16` | `http://schemas.microsoft.com/office/powerpoint/2015/main` | PowerPoint 2016 extensions |
| `p228` | `http://schemas.microsoft.com/office/powerpoint/2022/08/main` | Modern comments (Office 2021+) |
| `a14` | `http://schemas.microsoft.com/office/drawing/2010/main` | Drawing 2010 extensions |
| `a16` | `http://schemas.microsoft.com/office/drawing/2014/main` | Drawing 2014 extensions |

### DrawingML Sub-Namespaces

| Prefix | URI | Usage |
|--------|-----|-------|
| `c` | `http://schemas.openxmlformats.org/drawingml/2006/chart` | Chart elements |
| `dgm` | `http://schemas.openxmlformats.org/drawingml/2006/diagram` | SmartArt diagrams |

### Package and Metadata Namespaces

| Prefix | URI | Usage |
|--------|-----|-------|
| (none) | `http://schemas.openxmlformats.org/package/2006/relationships` | .rels files |
| (none) | `http://schemas.openxmlformats.org/package/2006/content-types` | [Content_Types].xml |
| `mc` | `http://schemas.openxmlformats.org/markup-compatibility/2006` | Markup compatibility |
| `cp` | `http://schemas.openxmlformats.org/package/2006/metadata/core-properties` | Core properties |
| `dc` | `http://purl.org/dc/elements/1.1/` | Dublin Core metadata |
| `dcterms` | `http://purl.org/dc/terms/` | Dublin Core terms |
| `dcmitype` | `http://purl.org/dc/dcmitype/` | Dublin Core types |
| `ep` | `http://schemas.openxmlformats.org/officeDocument/2006/extended-properties` | Extended properties |

### Legacy Namespaces (Transitional Only)

| Prefix | URI | Usage |
|--------|-----|-------|
| `v` | `urn:schemas-microsoft-com:vml` | Vector Markup Language (legacy drawings) |
| `o` | `urn:schemas-microsoft-com:office:office` | Office VML extensions |

### Typical Root Element Declaration

```xml
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
```

---

## 4. Content Types

The file `[Content_Types].xml` maps parts to their MIME-like content types. It is **required** at the package root.

### Example

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="wav" ContentType="audio/wav"/>
  <Override PartName="/ppt/presentation.xml"
            ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml"
            ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml"
            ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml"
            ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/theme/theme1.xml"
            ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/ppt/presProps.xml"
            ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>
  <Override PartName="/ppt/viewProps.xml"
            ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>
  <Override PartName="/ppt/tableStyles.xml"
            ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>
  <Override PartName="/docProps/core.xml"
            ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml"
            ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
```

### Complete Content Type Reference

| Part | Content Type |
|------|-------------|
| Presentation (main) | `application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml` |
| Presentation (macro-enabled) | `application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml` |
| Slideshow | `application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml` |
| Template | `application/vnd.openxmlformats-officedocument.presentationml.template.main+xml` |
| Slide | `application/vnd.openxmlformats-officedocument.presentationml.slide+xml` |
| Slide Layout | `application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml` |
| Slide Master | `application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml` |
| Theme | `application/vnd.openxmlformats-officedocument.theme+xml` |
| Theme Override | `application/vnd.openxmlformats-officedocument.themeOverride+xml` |
| Notes Slide | `application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml` |
| Notes Master | `application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml` |
| Handout Master | `application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml` |
| Comments | `application/vnd.openxmlformats-officedocument.presentationml.comments+xml` |
| Comment Authors | `application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml` |
| Presentation Properties | `application/vnd.openxmlformats-officedocument.presentationml.presProps+xml` |
| View Properties | `application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml` |
| Table Styles | `application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml` |
| Tags | `application/vnd.openxmlformats-officedocument.presentationml.tags+xml` |
| Chart | `application/vnd.openxmlformats-officedocument.drawingml.chart+xml` |
| Diagram Data | `application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml` |
| Diagram Layout | `application/vnd.openxmlformats-officedocument.drawingml.diagramLayoutDefinition+xml` |
| Diagram Style | `application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml` |
| Diagram Colors | `application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml` |

---

## 5. Relationship System

Relationships define how parts connect to each other. They are stored in `.rels` files inside `_rels` subdirectories.

### Package-Level Relationships (`_rels/.rels`)

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

### Presentation-Level Relationships (`ppt/_rels/presentation.xml.rels`)

```xml
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
    Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
    Target="slides/slide1.xml"/>
  <Relationship Id="rId3"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
    Target="slides/slide2.xml"/>
  <Relationship Id="rId4"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster"
    Target="notesMasters/notesMaster1.xml"/>
  <Relationship Id="rId5"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster"
    Target="handoutMasters/handoutMaster1.xml"/>
  <Relationship Id="rId6"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps"
    Target="presProps.xml"/>
  <Relationship Id="rId7"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps"
    Target="viewProps.xml"/>
  <Relationship Id="rId8"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
    Target="theme/theme1.xml"/>
  <Relationship Id="rId9"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles"
    Target="tableStyles.xml"/>
  <Relationship Id="rId10"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors"
    Target="commentAuthors.xml"/>
</Relationships>
```

### Slide-Level Relationships (`ppt/slides/_rels/slide1.xml.rels`)

```xml
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
    Target="../slideLayouts/slideLayout2.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
    Target="../notesSlides/notesSlide1.xml"/>
  <Relationship Id="rId3"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    Target="../media/image1.png"/>
  <Relationship Id="rId4"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
    Target="https://example.com" TargetMode="External"/>
</Relationships>
```

### Relationship Type URI Reference

| Relationship | Type URI |
|-------------|----------|
| Office Document | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument` |
| Slide | `...relationships/slide` |
| Slide Layout | `...relationships/slideLayout` |
| Slide Master | `...relationships/slideMaster` |
| Notes Slide | `...relationships/notesSlide` |
| Notes Master | `...relationships/notesMaster` |
| Handout Master | `...relationships/handoutMaster` |
| Theme | `...relationships/theme` |
| Theme Override | `...relationships/themeOverride` |
| Presentation Props | `...relationships/presProps` |
| View Props | `...relationships/viewProps` |
| Table Styles | `...relationships/tableStyles` |
| Comment Authors | `...relationships/commentAuthors` |
| Comments | `...relationships/comments` |
| Image | `...relationships/image` |
| Audio | `...relationships/audio` |
| Video | `...relationships/video` |
| Hyperlink | `...relationships/hyperlink` |
| Chart | `...relationships/chart` |
| OLE Object | `...relationships/oleObject` |
| Package (embedded) | `...relationships/package` |
| Tags | `...relationships/tags` |
| Font | `...relationships/font` |
| Core Properties | `http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties` |
| Extended Properties | `...relationships/extended-properties` |
| Thumbnail | `http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail` |

(All `...relationships/` entries are prefixed with `http://schemas.openxmlformats.org/officeDocument/2006/` unless otherwise shown.)

### Relationship Attributes

| Attribute | Required | Description |
|-----------|----------|-------------|
| `Id` | Yes | Unique identifier within the .rels file (e.g., `rId1`). Referenced via `r:id` in markup. |
| `Type` | Yes | URI identifying the relationship type. |
| `Target` | Yes | Relative URI to the target part or absolute URI for external targets. |
| `TargetMode` | No | `Internal` (default) for package parts, `External` for outside resources. |

### Relationship Navigation Chain

Understanding the relationship graph is critical for resolving the inheritance chain:

```
Package (_rels/.rels)
  -> presentation.xml
       -> slide1.xml (via r:id in sldIdLst)
       |    -> slideLayout2.xml (via slide rels)
       |    |    -> slideMaster1.xml (via layout rels)
       |    |         -> theme1.xml (via master rels)
       |    -> notesSlide1.xml (via slide rels)
       |    -> media/image1.png (via slide rels)
       -> slideMaster1.xml
       |    -> slideLayout1..N.xml (via master rels)
       |    -> theme1.xml (via master rels)
       -> notesMaster1.xml
       -> theme1.xml
```

---

## 6. Presentation Part (presentation.xml)

The `presentation.xml` file is the root part of a PresentationML document. Its root element is `<p:presentation>`.

### Full Structure

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                saveSubsetFonts="1"
                autoCompressPictures="0">

  <!-- Slide master references -->
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId1"/>
  </p:sldMasterIdLst>

  <!-- Notes master (optional) -->
  <p:notesMasterIdLst>
    <p:notesMasterId r:id="rId4"/>
  </p:notesMasterIdLst>

  <!-- Handout master (optional) -->
  <p:handoutMasterIdLst>
    <p:handoutMasterId r:id="rId5"/>
  </p:handoutMasterIdLst>

  <!-- Ordered slide list -- THIS DEFINES SLIDE ORDER -->
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId2"/>
    <p:sldId id="257" r:id="rId3"/>
  </p:sldIdLst>

  <!-- Slide size in EMUs -->
  <p:sldSz cx="12192000" cy="6858000" type="custom"/>

  <!-- Notes slide size -->
  <p:notesSz cx="6858000" cy="9144000"/>

  <!-- Default text styling for new text -->
  <p:defaultTextStyle>
    <a:defPPr>
      <a:defRPr lang="en-US"/>
    </a:defPPr>
    <a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1"
               latinLnBrk="0" hangingPunct="1">
      <a:defRPr sz="1800" kern="1200">
        <a:solidFill><a:schemeClr val="tx1"/></a:solidFill>
        <a:latin typeface="+mn-lt"/>
        <a:ea typeface="+mn-ea"/>
        <a:cs typeface="+mn-cs"/>
      </a:defRPr>
    </a:lvl1pPr>
    <!-- lvl2pPr through lvl9pPr follow the same pattern -->
  </p:defaultTextStyle>

</p:presentation>
```

### Key `<p:presentation>` Attributes

| Attribute | Type | Description |
|-----------|------|-------------|
| `saveSubsetFonts` | boolean | Embed only used character subsets |
| `autoCompressPictures` | boolean | Auto-compress images on save |
| `bookmarkIdSeed` | uint32 | Next bookmark ID to use |
| `firstSlideNum` | int32 | Starting slide number (default 1) |
| `rtl` | boolean | Right-to-left presentation |
| `removePersonalInfoOnSave` | boolean | Strip personal info on save |
| `showSpecialPlsOnTitleSld` | boolean | Show special placeholders on title slide |
| `strictFirstAndLastChars` | boolean | Strict East Asian line-break rules |

### Key Child Elements

| Element | Description |
|---------|-------------|
| `p:sldMasterIdLst` | List of slide master references |
| `p:notesMasterIdLst` | List of notes master references |
| `p:handoutMasterIdLst` | List of handout master references |
| `p:sldIdLst` | Ordered list of slides (determines presentation order) |
| `p:sldSz` | Slide dimensions |
| `p:notesSz` | Notes page dimensions |
| `p:defaultTextStyle` | Default text formatting |
| `p:embeddedFontLst` | Embedded fonts list |
| `p:custShowLst` | Custom slide show definitions |
| `p:kinsoku` | East Asian line-break settings |
| `p:modifyVerifier` | Document protection hash |

### Standard Slide Sizes

| Type | cx (EMU) | cy (EMU) | Aspect |
|------|----------|----------|--------|
| `screen4x3` | 9144000 | 6858000 | 4:3 (10" x 7.5") |
| `screen16x9` | 12192000 | 6858000 | 16:9 (13.33" x 7.5") |
| `screen16x10` | 12192000 | 7620000 | 16:10 |
| `letter` | 9144000 | 6858000 | Letter |
| `A4` | 9906000 | 6858000 | A4 |
| `custom` | (any) | (any) | Custom dimensions |

### EMU Quick Reference

1 inch = 914400 EMU, 1 cm = 360000 EMU, 1 pt = 12700 EMU, 1 pixel (96 dpi) = 9525 EMU.

---

## 7. Slide Structure (slide.xml)

Each slide is stored as a separate XML part (e.g., `ppt/slides/slide1.xml`). The root element is `<p:sld>`.

### Slide Structure

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">

  <!-- Common slide data (shapes, text, etc.) -->
  <p:cSld name="">
    <!-- Optional background override -->
    <p:bg>
      <p:bgPr>
        <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
        <a:effectLst/>
      </p:bgPr>
    </p:bg>

    <!-- Shape tree -- ALL visual objects on the slide -->
    <p:spTree>
      <!-- Shape tree root (itself a group shape) -->
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>

      <!-- Individual shapes go here (sp, pic, grpSp, graphicFrame, cxnSp) -->
    </p:spTree>
  </p:cSld>

  <!-- Optional: color map override -->
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>

  <!-- Optional: slide transition -->
  <p:transition spd="med" advClick="1" advTm="3000">
    <p:fade/>
  </p:transition>

  <!-- Optional: animation timing -->
  <p:timing>
    <p:tnLst>
      <p:par>
        <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"/>
      </p:par>
    </p:tnLst>
  </p:timing>

</p:sld>
```

### `<p:sld>` Attributes

| Attribute | Type | Default | Description |
|-----------|------|---------|-------------|
| `show` | boolean | `true` | Whether slide is shown during slideshow |
| `showMasterSp` | boolean | `true` | Show master slide shapes |
| `showMasterPhAnim` | boolean | `true` | Show master placeholder animations |

### Common Slide Data (`p:cSld`)

This element is shared by slides, slide layouts, slide masters, notes slides, notes masters, and handout masters. Child elements:

| Element | Description |
|---------|-------------|
| `p:bg` | Slide background |
| `p:spTree` | Shape tree (required) -- contains all visual content |
| `p:custDataLst` | Custom data |
| `p:controls` | ActiveX controls |
| `p:extLst` | Extension list |

### Slide Background (`p:bg`)

```xml
<p:bg>
  <!-- Explicit background properties -->
  <p:bgPr>
    <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
    <a:effectLst/>
  </p:bgPr>
  <!-- OR theme background reference -->
  <!-- <p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef> -->
</p:bg>
```

Background fill options: `a:solidFill`, `a:gradFill`, `a:pattFill`, `a:blipFill` (image), `a:noFill`.

### Z-Order

Shapes in the shape tree are rendered in document order: first child is at the back (lowest z-index), last child is at the front (highest z-index).

---

## 8. Shape Types

All shapes share a common three-part non-visual property pattern:

1. **`p:cNvPr`** (Common Non-Visual Properties): `id` (unique uint32 within slide), `name`, `descr` (alt text), `hidden`, `title`, hyperlink click/hover actions.
2. **`p:cNvXxxPr`** (type-specific non-visual properties): Constraints (e.g., `noGrp`, `noRot`, `noResize` for shapes; `noChangeAspect` for pictures).
3. **`p:nvPr`** (PresentationML non-visual properties): Placeholder info (`p:ph`), audio/video file refs, customer data.

### 8.1 AutoShapes (`p:sp`)

The most common shape type. Used for text boxes, rectangles, arrows, callouts, and all preset geometry shapes.

```xml
<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="4" name="Rectangle 3" descr="Alt text"/>
    <p:cNvSpPr>
      <a:spLocks noGrp="1"/>
    </p:cNvSpPr>
    <p:nvPr>
      <p:ph type="body" idx="1"/>  <!-- Placeholder info (if placeholder) -->
    </p:nvPr>
  </p:nvSpPr>

  <p:spPr>
    <a:xfrm>
      <a:off x="457200" y="1600200"/>       <!-- position in EMU -->
      <a:ext cx="8229600" cy="4525963"/>    <!-- width and height in EMU -->
    </a:xfrm>
    <a:prstGeom prst="rect">               <!-- preset geometry -->
      <a:avLst/>                            <!-- adjust values -->
    </a:prstGeom>
    <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>
    <a:ln w="12700">                        <!-- 1pt line -->
      <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
    </a:ln>
  </p:spPr>

  <!-- Theme style reference (optional) -->
  <p:style>
    <a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef>
    <a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>
    <a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>
    <a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>
  </p:style>

  <!-- Text content -->
  <p:txBody>
    <a:bodyPr/>
    <a:lstStyle/>
    <a:p>
      <a:r>
        <a:rPr lang="en-US" dirty="0"/>
        <a:t>Hello World</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>
```

#### Placeholder Types (`p:ph`)

| `type` | Description | Typical `idx` |
|--------|-------------|---------------|
| `title` | Slide title | 0 |
| `ctrTitle` | Centered title (title slide) | 0 |
| `subTitle` | Subtitle (title slide) | 1 |
| `body` | Body/content area | 1 |
| `dt` | Date/time | 10 |
| `ftr` | Footer | 11 |
| `sldNum` | Slide number | 12 |
| `hdr` | Header (notes/handouts only) | varies |
| `obj` | Generic object | varies |
| `chart` | Chart | varies |
| `tbl` | Table | varies |
| `clipArt` | Clip art | varies |
| `dgm` | Diagram (SmartArt) | varies |
| `media` | Media | varies |
| `sldImg` | Slide image (notes/handout) | varies |
| `pic` | Picture | varies |

The `idx` attribute on `p:ph` links a slide placeholder to its corresponding layout/master placeholder. Matching is done first by `idx`, then by `type`.

#### Common Preset Geometries (`a:prstGeom prst="..."`)

The complete enumeration has **187 values** defined in `ST_ShapeType`. Common ones:

| Category | Values |
|----------|--------|
| **Basic** | `rect`, `roundRect`, `ellipse`, `triangle`, `rtTriangle`, `parallelogram`, `trapezoid`, `diamond`, `pentagon`, `hexagon`, `octagon` |
| **Arrows** | `rightArrow`, `leftArrow`, `upArrow`, `downArrow`, `leftRightArrow`, `upDownArrow`, `chevron`, `notchedRightArrow` |
| **Flowchart** | `flowChartProcess`, `flowChartDecision`, `flowChartTerminator`, `flowChartDocument`, `flowChartConnector` |
| **Stars** | `star4`, `star5`, `star6`, `star8`, `star10`, `star12`, `star16`, `star24`, `star32` |
| **Callouts** | `wedgeRoundRectCallout`, `wedgeEllipseCallout`, `cloudCallout`, `borderCallout1`, `borderCallout2` |
| **Connectors** | `line`, `straightConnector1`, `bentConnector2`-`5`, `curvedConnector2`-`5` |
| **Action Buttons** | `actionButtonBlank`, `actionButtonHome`, `actionButtonHelp`, `actionButtonForwardNext`, `actionButtonBackPrevious` |
| **Other** | `heart`, `lightningBolt`, `sun`, `moon`, `cloud`, `gear6`, `donut`, `pie`, `arc`, `smileyFace`, `frame`, `cube`, `bevel` |

### 8.2 Pictures (`p:pic`)

```xml
<p:pic>
  <p:nvPicPr>
    <p:cNvPr id="5" name="Picture 4" descr="Photo description"/>
    <p:cNvPicPr>
      <a:picLocks noChangeAspect="1"/>
    </p:cNvPicPr>
    <p:nvPr/>
  </p:nvPicPr>

  <p:blipFill>
    <a:blip r:embed="rId2" cstate="print">
      <a:alphaModFix amt="50000"/>  <!-- 50% opacity -->
    </a:blip>
    <a:srcRect l="10000" t="10000" r="10000" b="10000"/>  <!-- crop: 1/1000 of % -->
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>

  <p:spPr>
    <a:xfrm>
      <a:off x="1524000" y="1397000"/>
      <a:ext cx="6096000" cy="4064000"/>
    </a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>
```

Key points:
- `r:embed` in `a:blip` references the image binary via the slide's `.rels` file.
- `r:link` can be used instead for externally-linked images (rare).
- `a:srcRect` attributes are in 1/1000th of a percent (e.g., `l="10000"` = crop 10% from left).

### 8.3 Group Shapes (`p:grpSp`)

Recursive container. Uses its own coordinate space for children.

```xml
<p:grpSp>
  <p:nvGrpSpPr>
    <p:cNvPr id="8" name="Group 7"/>
    <p:cNvGrpSpPr/>
    <p:nvPr/>
  </p:nvGrpSpPr>

  <p:grpSpPr>
    <a:xfrm>
      <!-- Position and size on the slide -->
      <a:off x="1000000" y="1000000"/>
      <a:ext cx="5000000" cy="3000000"/>
      <!-- Internal coordinate space for children -->
      <a:chOff x="0" y="0"/>
      <a:chExt cx="5000000" cy="3000000"/>
    </a:xfrm>
  </p:grpSpPr>

  <!-- Children use coordinates relative to chOff/chExt -->
  <p:sp>...</p:sp>
  <p:pic>...</p:pic>
  <p:grpSp><!-- nested groups are allowed --></p:grpSp>
</p:grpSp>
```

The coordinate transform maps `(chOff, chExt)` to `(off, ext)`. To compute absolute position: `abs_x = off.x + (child.x - chOff.x) * (ext.cx / chExt.cx)`.

### 8.4 Graphic Frames (`p:graphicFrame`)

Container for tables, charts, SmartArt diagrams, and OLE objects.

```xml
<p:graphicFrame>
  <p:nvGraphicFramePr>
    <p:cNvPr id="9" name="Table 8"/>
    <p:cNvGraphicFramePr>
      <a:graphicFrameLocks noGrp="1"/>
    </p:cNvGraphicFramePr>
    <p:nvPr/>
  </p:nvGraphicFramePr>

  <p:xfrm>
    <a:off x="838200" y="1825625"/>
    <a:ext cx="10515600" cy="3403600"/>
  </p:xfrm>

  <a:graphic>
    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
      <a:tbl>...</a:tbl>
    </a:graphicData>
  </a:graphic>
</p:graphicFrame>
```

**GraphicData URI values** determine the content type:

| URI | Content Type |
|-----|-------------|
| `http://schemas.openxmlformats.org/drawingml/2006/table` | Table |
| `http://schemas.openxmlformats.org/drawingml/2006/chart` | Chart |
| `http://schemas.openxmlformats.org/drawingml/2006/diagram` | SmartArt diagram |
| `http://schemas.openxmlformats.org/presentationml/2006/ole` | OLE object |

### 8.5 Connectors (`p:cxnSp`)

Lines that connect two shapes, with automatic routing.

```xml
<p:cxnSp>
  <p:nvCxnSpPr>
    <p:cNvPr id="6" name="Straight Connector 5"/>
    <p:cNvCxnSpPr>
      <a:stCxn id="4" idx="2"/>   <!-- start: shape id=4, connection site 2 -->
      <a:endCxn id="7" idx="0"/>  <!-- end: shape id=7, connection site 0 -->
    </p:cNvCxnSpPr>
    <p:nvPr/>
  </p:nvCxnSpPr>

  <p:spPr>
    <a:xfrm>
      <a:off x="3048000" y="2286000"/>
      <a:ext cx="1524000" cy="762000"/>
    </a:xfrm>
    <a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>
    <a:ln w="19050">
      <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>
      <a:tailEnd type="triangle"/>  <!-- arrowhead -->
    </a:ln>
  </p:spPr>
</p:cxnSp>
```

Connection site indices are defined by the target shape's geometry. For a rectangle: 0=top, 1=right, 2=bottom, 3=left.

---

## 9. Text Body System

Text in PresentationML uses the DrawingML text body (`<p:txBody>` or `<a:txBody>`), which appears inside `<p:sp>` elements.

### Text Body Hierarchy

```
txBody
  +-- bodyPr          (body properties: margins, autofit, columns, anchor, etc.)
  +-- lstStyle        (list style: up to 9 indent-level default paragraph props)
  +-- p*              (one or more paragraphs)
        +-- pPr        (paragraph properties: alignment, indent, spacing, bullets)
        +-- r*         (text runs)
        |     +-- rPr  (run properties: font, size, bold, italic, color, etc.)
        |     +-- t    (text content)
        +-- br         (line break within paragraph)
        +-- fld        (field: slide number, date, etc.)
        +-- endParaRPr (end-of-paragraph run properties for empty paragraphs)
```

### Body Properties (`a:bodyPr`)

```xml
<a:bodyPr wrap="square"
          lIns="91440" tIns="45720" rIns="91440" bIns="45720"
          anchor="t" anchorCtr="0"
          rtlCol="0" numCol="1" spcCol="0"
          vert="horz" rot="0">
  <!-- Auto-fit options (choose one): -->
  <a:normAutofit fontScale="92500" lnSpcReduction="10000"/>
  <!-- OR: <a:spAutoFit/> -->
  <!-- OR: <a:noAutofit/> -->
</a:bodyPr>
```

| Attribute | Type | Description |
|-----------|------|-------------|
| `wrap` | enum | `none`, `square` -- text wrapping mode |
| `lIns`, `tIns`, `rIns`, `bIns` | EMU | Left, top, right, bottom text insets |
| `anchor` | enum | Vertical anchor: `t` (top), `ctr` (center), `b` (bottom), `just`, `dist` |
| `anchorCtr` | boolean | Center text horizontally in shape |
| `vert` | enum | Text direction: `horz`, `vert`, `vert270`, `wordArtVert`, `eaVert`, `mongolianVert` |
| `rot` | int | Text rotation in 60,000ths of a degree |
| `numCol` | int | Number of text columns |
| `spcCol` | EMU | Column spacing |
| `rtlCol` | boolean | Right-to-left columns |
| `upright` | boolean | Keep text upright when shape is rotated |
| `horzOverflow` | enum | `overflow` or `clip` |
| `vertOverflow` | enum | `overflow`, `clip`, or `ellipsis` |

#### Auto-Fit Options

| Element | Behavior |
|---------|----------|
| `<a:noAutofit/>` | Fixed text box -- text can overflow |
| `<a:spAutoFit/>` | Shape resizes to fit text |
| `<a:normAutofit fontScale="..." lnSpcReduction="..."/>` | Text shrinks to fit (fontScale in 1/1000 of percent, e.g., `92500` = 92.5%) |

### List Style (`a:lstStyle`)

Defines default paragraph properties for 9 indent levels:

```xml
<a:lstStyle>
  <a:lvl1pPr marL="0" indent="0" algn="l">
    <a:buNone/>
    <a:defRPr sz="3200"/>
  </a:lvl1pPr>
  <a:lvl2pPr marL="457200" indent="-228600" algn="l">
    <a:buChar char="&#x2013;"/>
    <a:defRPr sz="2800"/>
  </a:lvl2pPr>
  <!-- lvl3pPr through lvl9pPr -->
</a:lstStyle>
```

### Paragraph Properties (`a:pPr`)

```xml
<a:pPr lvl="0" algn="ctr" marL="0" marR="0" indent="0"
       defTabSz="914400" rtl="0" fontAlgn="auto">
  <a:lnSpc><a:spcPct val="100000"/></a:lnSpc>    <!-- Line spacing: 100% -->
  <a:spcBef><a:spcPts val="0"/></a:spcBef>         <!-- Space before: 0 pt -->
  <a:spcAft><a:spcPts val="600"/></a:spcAft>       <!-- Space after: 6 pt -->

  <!-- Bullet/numbering -->
  <a:buFont typeface="Arial"/>
  <a:buChar char="&#x2022;"/>
  <!-- OR: <a:buAutoNum type="arabicPeriod"/> -->
  <!-- OR: <a:buBlip><a:blip r:embed="rId3"/></a:buBlip> -->
  <!-- OR: <a:buNone/> -->

  <a:buClr><a:schemeClr val="accent1"/></a:buClr>
  <a:buSzPct val="100000"/>

  <a:tabLst><a:tab pos="914400" algn="l"/></a:tabLst>
</a:pPr>
```

| Attribute | Description |
|-----------|-------------|
| `algn` | Alignment: `l` (left), `ctr` (center), `r` (right), `just` (justified), `dist` (distributed) |
| `marL` | Left margin in EMU |
| `indent` | First-line indent in EMU (negative = hanging indent) |
| `lvl` | Outline level (0-8) |
| `rtl` | Right-to-left paragraph |
| `defTabSz` | Default tab stop spacing in EMU |

### Run Properties (`a:rPr`)

```xml
<a:rPr lang="en-US" sz="2400" b="1" i="0" u="sng" strike="noStrike"
       kern="1200" spc="0" baseline="0" dirty="0">
  <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
  <a:latin typeface="Arial"/>
  <a:ea typeface="+mn-ea"/>
  <a:cs typeface="+mn-cs"/>
  <a:hlinkClick r:id="rId1"/>
</a:rPr>
```

| Attribute | Type | Description |
|-----------|------|-------------|
| `sz` | int | Font size in hundredths of a point (2400 = 24pt) |
| `b` | boolean | Bold |
| `i` | boolean | Italic |
| `u` | enum | Underline: `none`, `sng`, `dbl`, `heavy`, `dotted`, `dash`, `wavy`, `wavyHeavy`, `words`, etc. |
| `strike` | enum | `noStrike`, `sngStrike`, `dblStrike` |
| `kern` | int | Kerning threshold in hundredths of a point |
| `spc` | int | Character spacing in hundredths of a point |
| `baseline` | int | Superscript/subscript (30000 = superscript 30%, -25000 = subscript) |
| `cap` | enum | Capitalization: `none`, `small`, `all` |
| `lang` | string | Language tag (e.g., `en-US`) |
| `dirty` | boolean | Needs spell-check (implementation artifact, ignore on read) |
| `err` | boolean | Text has error |
| `noProof` | boolean | Skip proofing |

#### Run Property Child Elements

| Element | Description |
|---------|-------------|
| `a:solidFill` / `a:gradFill` / `a:noFill` | Text fill color |
| `a:ln` | Text outline |
| `a:effectLst` | Text effects (shadow, glow, etc.) |
| `a:highlight` | Highlight color |
| `a:latin` | Latin font (`typeface`, `panose`, `pitchFamily`, `charset`) |
| `a:ea` | East Asian font |
| `a:cs` | Complex Script font |
| `a:sym` | Symbol font |
| `a:hlinkClick` | Hyperlink on click (`r:id`, `action`, `tooltip`) |
| `a:hlinkMouseOver` | Hyperlink on mouse-over |

### Font References

Font `typeface` values starting with `+` reference theme fonts:
- `+mj-lt` = Major Latin (heading font)
- `+mn-lt` = Minor Latin (body font)
- `+mj-ea` / `+mn-ea` = Major/Minor East Asian
- `+mj-cs` / `+mn-cs` = Major/Minor Complex Script

### Field Element (`a:fld`)

```xml
<a:fld id="{B6F15528-F159-4107-2D4E-6B91C1C71221}" type="slidenum">
  <a:rPr lang="en-US"/>
  <a:t>1</a:t>    <!-- Current/cached value -->
</a:fld>
```

Common field types: `slidenum`, `datetime1` through `datetime13`.

### Text Formatting Inheritance Chain

Text formatting resolves through a multi-level inheritance chain (first match wins):

1. **Run-level** `a:rPr` on the specific run
2. **Paragraph-level** `a:pPr/a:defRPr` on the containing paragraph
3. **Shape-level** `p:txBody/a:lstStyle/a:lvlNpPr` on the shape
4. **Layout-level** matching placeholder's `a:lstStyle` on the slide layout
5. **Master-level** matching placeholder's `a:lstStyle` on the slide master
6. **Master `txStyles`** (`p:txStyles/p:titleStyle` or `p:txStyles/p:bodyStyle` on master)
7. **Theme `fmtScheme`** font scheme defaults
8. **Presentation `defaultTextStyle`** in `presentation.xml`

---

## 10. Slide Masters and Slide Layouts

### Inheritance Chain

```
Slide  ->  Slide Layout  ->  Slide Master  ->  Theme
```

Each level can define:
- **Background**: Slide overrides layout, layout overrides master
- **Shapes**: Master shapes appear on all slides (unless `showMasterSp="0"`)
- **Placeholders**: Matched by `idx` and `type` attributes through the chain
- **Text styles**: Each level can provide defaults that are inherited

### Slide Master (`slideMaster1.xml`)

```xml
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">

  <p:cSld>
    <p:bg>
      <p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef>
    </p:bg>
    <p:spTree>
      <!-- Master-level placeholder shapes:
           title, body, date, footer, slide number, etc. -->
    </p:spTree>
  </p:cSld>

  <!-- REQUIRED: color map (maps logical colors to theme colors) -->
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2"
            accent1="accent1" accent2="accent2" accent3="accent3"
            accent4="accent4" accent5="accent5" accent6="accent6"
            hlink="hlink" folHlink="folHlink"/>

  <!-- Slide layout references -->
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="2147483649" r:id="rId1"/>
    <p:sldLayoutId id="2147483650" r:id="rId2"/>
    <p:sldLayoutId id="2147483651" r:id="rId3"/>
  </p:sldLayoutIdLst>

  <!-- Text styles for title, body, and other text -->
  <p:txStyles>
    <p:titleStyle>
      <a:lvl1pPr algn="l" defTabSz="914400" rtl="0" eaLnBrk="1">
        <a:spcBef><a:spcPct val="0"/></a:spcBef>
        <a:defRPr sz="4400" kern="1200">
          <a:solidFill><a:schemeClr val="tx1"/></a:solidFill>
          <a:latin typeface="+mj-lt"/>
          <a:ea typeface="+mj-ea"/>
          <a:cs typeface="+mj-cs"/>
        </a:defRPr>
      </a:lvl1pPr>
    </p:titleStyle>
    <p:bodyStyle>
      <a:lvl1pPr marL="228600" indent="-228600" algn="l">
        <a:buFont typeface="Arial"/>
        <a:buChar char="&#x2022;"/>
        <a:defRPr sz="3200" kern="1200">
          <a:solidFill><a:schemeClr val="tx1"/></a:solidFill>
          <a:latin typeface="+mn-lt"/>
        </a:defRPr>
      </a:lvl1pPr>
      <!-- lvl2pPr through lvl5pPr with decreasing sizes -->
    </p:bodyStyle>
    <p:otherStyle>
      <!-- Default styles for shapes that are not placeholders -->
    </p:otherStyle>
  </p:txStyles>

</p:sldMaster>
```

### Color Map (`p:clrMap`)

Maps 12 logical color names to theme color slots:

| Logical Name | Description | Typical Mapping |
|-------------|-------------|-----------------|
| `bg1` | Background 1 | `lt1` |
| `tx1` | Text 1 | `dk1` |
| `bg2` | Background 2 | `lt2` |
| `tx2` | Text 2 | `dk2` |
| `accent1`-`accent6` | Accent colors | `accent1`-`accent6` |
| `hlink` | Hyperlink | `hlink` |
| `folHlink` | Followed hyperlink | `folHlink` |

### Slide Layout (`slideLayout1.xml`)

Each layout is associated with exactly one slide master. Root element: `<p:sldLayout>`.

```xml
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             type="twoObj" preserve="1" showMasterSp="1">

  <p:cSld name="Two Content">
    <p:spTree>
      <p:nvGrpSpPr>...</p:nvGrpSpPr>
      <p:grpSpPr>...</p:grpSpPr>

      <!-- Placeholders inherit from master but can override position/size -->
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title 1"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="title"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr/>  <!-- empty = inherit from master -->
        <p:txBody>...</p:txBody>
      </p:sp>

      <!-- Content placeholder with explicit position -->
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="Content Placeholder 2"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph idx="1" sz="half"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="838200" y="1825625"/>
            <a:ext cx="5181600" cy="4351338"/>
          </a:xfrm>
        </p:spPr>
        <p:txBody>...</p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>

  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>
```

### Layout Attributes

| Attribute | Type | Description |
|-----------|------|-------------|
| `type` | ST_SlideLayoutType | Layout type enum (see below) |
| `matchingName` | string | Matching name for custom layouts |
| `preserve` | boolean | Prevent removal when not in use |
| `showMasterSp` | boolean | Show shapes from slide master |
| `userDrawn` | boolean | User-created layout |

### Standard Layout Types (`ST_SlideLayoutType`)

| Value | Description |
|-------|-------------|
| `title` | Title Slide |
| `obj` | Title and Content (most common) |
| `twoObj` | Two Content |
| `titleOnly` | Title Only |
| `blank` | Blank |
| `cust` | Custom |
| `tx` | Text |
| `twoTxTwoObj` | Comparison |
| `txAndObj` | Title, Text, and Content |
| `objAndTx` | Title, Content, and Text |
| `secHead` | Section Header |
| `twoObjAndObj` | Title and Content over Text |
| `objOverTx` | Content over Text |
| `picTx` | Picture with Caption |
| `vertTitleAndTx` | Vertical Title and Text |
| `vertTx` | Vertical Text |
| `mediaAndTx` | Media and Text |
| `dgm` | Diagram |
| `chart` | Chart |
| `txAndChart` | Text and Chart |
| `fourObj` | Four Objects |

### Placeholder Matching Algorithm

When rendering a slide, placeholders are resolved by matching through the chain:

1. Find placeholder on slide by `idx` attribute.
2. If `idx` not found, match by `type` attribute.
3. Look up matching placeholder on the slide layout (same `idx` or `type`).
4. Look up matching placeholder on the slide master.
5. Merge properties: slide placeholder properties override layout, which override master.

If a placeholder on the slide has no `a:xfrm` (empty `p:spPr`), it inherits its position/size from the layout. If the layout placeholder also has no transform, it inherits from the master.

---

## 11. Themes (DrawingML)

A theme (`ppt/theme/theme1.xml`) defines the visual identity: colors, fonts, and effects. Themes are shared DrawingML components (used identically across DOCX, XLSX, and PPTX). In PPTX, the theme is referenced from the slide master.

### Theme Structure

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
         name="Office Theme">

  <a:themeElements>

    <!-- Color scheme: 12 named color slots -->
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

    <!-- Font scheme: major (headings) and minor (body) -->
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Calibri Light" panose="020F0302020204030204"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
        <a:font script="Jpan" typeface="Yu Gothic Light"/>
        <a:font script="Hang" typeface="&#xB9D1;&#xC740; &#xACE0;&#xB515;"/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri" panose="020F0502020204030204"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:minorFont>
    </a:fontScheme>

    <!-- Format scheme: fill, line, effect, and background fill styles (3 levels each) -->
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:gradFill rotWithShape="1">...</a:gradFill>
        <a:gradFill rotWithShape="1">...</a:gradFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="6350">...</a:ln>
        <a:ln w="12700">...</a:ln>
        <a:ln w="19050">...</a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst>...</a:effectLst></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr">...</a:schemeClr></a:solidFill>
        <a:gradFill rotWithShape="1">...</a:gradFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>

  </a:themeElements>

  <a:objectDefaults/>
  <a:extraClrSchemeLst/>
</a:theme>
```

### Theme Color Slots

| Slot | Semantic | Typical Use |
|------|----------|-------------|
| `dk1` | Dark 1 | Primary text color |
| `lt1` | Light 1 | Primary background |
| `dk2` | Dark 2 | Secondary dark |
| `lt2` | Light 2 | Secondary light |
| `accent1` - `accent6` | Accent colors | Charts, shapes, emphasis |
| `hlink` | Hyperlink | Unvisited hyperlinks |
| `folHlink` | Followed hyperlink | Visited hyperlinks |

Additional logical names used in markup (resolved via `p:clrMap`):
- `tx1`, `tx2` -- text colors (typically map to `dk1`, `dk2`)
- `bg1`, `bg2` -- background colors (typically map to `lt1`, `lt2`)
- `phClr` -- placeholder color (replaced by context-specific color at render time)

### Color Transforms

Scheme colors can be modified with transforms:

```xml
<a:schemeClr val="accent1">
  <a:tint val="40000"/>       <!-- 40% tint (lighter) -->
  <a:shade val="75000"/>      <!-- 75% shade (darker) -->
  <a:satMod val="120000"/>    <!-- 120% saturation -->
  <a:lumMod val="80000"/>     <!-- 80% luminance -->
  <a:alpha val="50000"/>      <!-- 50% opacity -->
</a:schemeClr>
```

### Color Specification Models

Six color models are available, each supporting child transform elements:

| Element | Description | Example |
|---------|-------------|---------|
| `a:srgbClr` | sRGB hex | `val="4472C4"` |
| `a:schemeClr` | Theme/scheme color | `val="accent1"` |
| `a:sysClr` | System color | `val="windowText" lastClr="000000"` |
| `a:hslClr` | HSL color | `hue="0" sat="100000" lum="50000"` |
| `a:prstClr` | Preset color name | `val="red"` |
| `a:scrgbClr` | scRGB percentages | `r="50000" g="50000" b="50000"` |

### Style Matrix References

Shapes reference theme styles via `<p:style>`:

```xml
<p:style>
  <a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef>     <!-- Line style 1-3 -->
  <a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>   <!-- Fill style 1-3, or 1001-1003 for bg -->
  <a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef> <!-- Effect style 0-2 -->
  <a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>    <!-- "major" or "minor" -->
</p:style>
```

### Theme Color Resolution Algorithm

To resolve a scheme color to RGB:
1. Read the `p:clrMap` from the slide master (or override from layout/slide).
2. Map the scheme color name (e.g., `tx1`) to a theme color slot (e.g., `dk1`).
3. Look up the actual color value in the theme's `a:clrScheme`.
4. Apply any color transforms (tint, shade, satMod, lumMod, alpha) in order.

---

## 12. Transitions and Animations

### Slide Transitions

Transitions are specified within the `<p:transition>` element, a direct child of `<p:sld>`.

```xml
<p:transition spd="med" advClick="1" advTm="5000">
  <p:fade thruBlk="0"/>
</p:transition>
```

| Attribute | Type | Description |
|-----------|------|-------------|
| `spd` | enum | Speed: `slow`, `med`, `fast` |
| `advClick` | boolean | Advance on mouse click (default: `true`) |
| `advTm` | uint32 | Auto-advance time in milliseconds |

### Transition Type Elements

| Element | Attributes | Description |
|---------|-----------|-------------|
| `<p:blinds>` | `dir="horz\|vert"` | Blinds |
| `<p:checker>` | `dir="horz\|vert"` | Checkerboard |
| `<p:circle/>` | | Circle |
| `<p:comb>` | `dir="horz\|vert"` | Comb |
| `<p:cover>` | `dir="l\|r\|u\|d\|lu\|ru\|ld\|rd"` | Cover |
| `<p:cut>` | `thruBlk="0\|1"` | Cut (optionally through black) |
| `<p:diamond/>` | | Diamond |
| `<p:dissolve/>` | | Dissolve |
| `<p:fade>` | `thruBlk="0\|1"` | Fade (optionally through black) |
| `<p:newsflash/>` | | Newsflash |
| `<p:plus/>` | | Plus |
| `<p:pull>` | `dir="l\|r\|u\|d\|lu\|ru\|ld\|rd"` | Pull |
| `<p:push>` | `dir="l\|r\|u\|d"` | Push |
| `<p:random/>` | | Random |
| `<p:randomBar>` | `dir="horz\|vert"` | Random bars |
| `<p:split>` | `orient="horz\|vert" dir="in\|out"` | Split |
| `<p:strips>` | `dir="lu\|ru\|ld\|rd"` | Strips |
| `<p:wedge/>` | | Wedge |
| `<p:wheel>` | `spokes="4"` | Wheel |
| `<p:wipe>` | `dir="l\|r\|u\|d"` | Wipe |
| `<p:zoom>` | `dir="in\|out"` | Zoom |

### Transition with Sound

```xml
<p:transition spd="slow" advClick="1">
  <p:fade/>
  <p:sndAc>
    <p:stSnd>
      <p:snd r:embed="rId2" name="applause.wav"/>
    </p:stSnd>
  </p:sndAc>
</p:transition>
```

### Animations (`p:timing`)

Animations are stored in the `<p:timing>` element. The model is based on **SMIL** (Synchronized Multimedia Integration Language).

```xml
<p:timing>
  <p:tnLst>
    <p:par>
      <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
        <p:childTnLst>
          <!-- Main sequence (click-triggered animations) -->
          <p:seq concurrent="1" nextAc="seek">
            <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
              <p:childTnLst>
                <!-- One click-group per click -->
                <p:par>
                  <p:cTn id="3" fill="hold">
                    <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                    <p:childTnLst>
                      <p:par>
                        <p:cTn id="4" presetID="10" presetClass="entr"
                               presetSubtype="0" fill="hold"
                               grpId="0" nodeType="clickEffect">
                          <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                          <p:childTnLst>
                            <p:set>
                              <p:cBhvr>
                                <p:cTn id="5" dur="1" fill="hold">
                                  <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                                </p:cTn>
                                <p:tgtEl><p:spTgt spid="4"/></p:tgtEl>
                                <p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>
                              </p:cBhvr>
                              <p:to><p:strVal val="visible"/></p:to>
                            </p:set>
                            <p:animEffect transition="in" filter="fade">
                              <p:cBhvr>
                                <p:cTn id="6" dur="500"/>
                                <p:tgtEl><p:spTgt spid="4"/></p:tgtEl>
                              </p:cBhvr>
                            </p:animEffect>
                          </p:childTnLst>
                        </p:cTn>
                      </p:par>
                    </p:childTnLst>
                  </p:cTn>
                </p:par>
              </p:childTnLst>
            </p:cTn>
            <p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>
            <p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>
          </p:seq>
        </p:childTnLst>
      </p:cTn>
    </p:par>
  </p:tnLst>

  <p:bldLst>
    <p:bldP spid="3" grpId="0" build="p"/>  <!-- Build paragraphs one by one -->
  </p:bldLst>
</p:timing>
```

### Time Node Container Types

| Element | Description |
|---------|-------------|
| `p:par` | Parallel -- children play simultaneously |
| `p:seq` | Sequence -- children play one after another |
| `p:excl` | Exclusive -- only one child plays at a time |

### Common Time Node (`p:cTn`) Attributes

| Attribute | Type | Description |
|-----------|------|-------------|
| `id` | uint32 | Unique time node ID within the slide |
| `dur` | string/uint32 | Duration: ms value or `indefinite` |
| `fill` | enum | `hold`, `remove`, `freeze`, `transition` |
| `restart` | enum | `always`, `whenNotActive`, `never` |
| `nodeType` | enum | `tmRoot`, `mainSeq`, `interactiveSeq`, `clickEffect`, `withEffect`, `afterEffect`, `afterGroup` |
| `presetID` | int | Animation preset identifier |
| `presetClass` | enum | `entr` (entrance), `exit`, `emph` (emphasis), `path` (motion path), `mediacall` |
| `presetSubtype` | int | Preset subtype/direction |
| `grpId` | uint32 | Group identifier (links animation to build list) |

### Animation Behavior Elements

| Element | Description |
|---------|-------------|
| `p:anim` | Generic property animation (CSS-style property names) |
| `p:animClr` | Color animation |
| `p:animEffect` | Filter/transition effect (e.g., `blinds(horizontal)`) |
| `p:animMotion` | Motion path animation |
| `p:animRot` | Rotation animation |
| `p:animScale` | Scale animation |
| `p:set` | Discrete property set (e.g., visibility on/off) |
| `p:cmd` | Command (call, play/pause media) |
| `p:audio` | Audio playback |
| `p:video` | Video playback |

### Target Element Specification

```xml
<!-- Target a shape -->
<p:tgtEl><p:spTgt spid="4"/></p:tgtEl>

<!-- Target text within a shape (paragraph range) -->
<p:tgtEl>
  <p:spTgt spid="3">
    <p:txEl><p:pRg st="0" end="0"/></p:txEl>
  </p:spTgt>
</p:tgtEl>

<!-- Target the slide itself -->
<p:tgtEl><p:sldTgt/></p:tgtEl>
```

### Implementation Note

The animation timing tree is deeply nested and complex. For a first implementation, consider:
1. Parsing transitions (simple, high value)
2. Preserving animation XML on round-trip without full interpretation
3. Only later implementing full animation playback

---

## 13. Notes Slides and Handout Masters

### Notes Slide (`ppt/notesSlides/notesSlide1.xml`)

Each slide can have one associated notes slide. Root element: `<p:notes>`.

```xml
<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
         xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
         xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>...</p:nvGrpSpPr>
      <p:grpSpPr>...</p:grpSpPr>

      <!-- Slide image placeholder -->
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Slide Image Placeholder 1"/>
          <p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="sldImg"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
      </p:sp>

      <!-- Notes text placeholder -->
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="Notes Placeholder 2"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="en-US" dirty="0"/>
              <a:t>Speaker notes text goes here.</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>

      <!-- Slide number placeholder (optional) -->
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="4" name="Slide Number Placeholder 3"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="sldNum" sz="quarter" idx="10"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/><a:lstStyle/>
          <a:p>
            <a:fld id="{GUID}" type="slidenum">
              <a:rPr lang="en-US"/>
              <a:t>1</a:t>
            </a:fld>
            <a:endParaRPr lang="en-US"/>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:notes>
```

The notes slide has an implicit relationship to the Notes Master and an implicit relationship from its parent Slide.

### Notes Master (`ppt/notesMasters/notesMaster1.xml`)

Root element: `<p:notesMaster>`. Defines default formatting for all notes pages. Structure mirrors slide masters. At most one per package. Contains default placeholders: header, date, sldImg, body, footer, sldNum. Has its own `p:clrMap` and `p:notesStyle` (default text styles for notes).

### Handout Master (`ppt/handoutMasters/handoutMaster1.xml`)

Root element: `<p:handoutMaster>`. Defines how printed handout pages look. At most one per package. Contains placeholders for header, footer, date, and slide number. Slide image positions are implicit based on the handout layout type.

### Text Extraction from Notes

For LLM pipelines, notes text is extracted by:
1. Finding the notes slide relationship for each slide.
2. Locating the body placeholder (`p:ph type="body"`).
3. Extracting text from its `p:txBody`.

---

## 14. Embedded Media

### Images

Images are stored in `ppt/media/` and referenced via relationships. Supported formats: PNG, JPEG, GIF, BMP, TIFF, WMF, EMF, SVG.

**In slide .rels:**
```xml
<Relationship Id="rId2"
  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
  Target="../media/image1.png"/>
```

**Referenced in markup via `r:embed`:**
```xml
<a:blip r:embed="rId2"/>
```

### Audio

```xml
<!-- In the shape's non-visual properties -->
<p:nvPr>
  <a:audioFile r:link="rId3"/>
  <!-- OR for embedded audio: <a:audioFile r:embed="rId3"/> -->
</p:nvPr>
```

### Video

```xml
<p:nvPr>
  <a:videoFile r:link="rId4"/>
  <!-- Modern PowerPoint also uses p14:media extension -->
  <p:extLst>
    <p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}">
      <p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"
                 r:embed="rId5"/>
    </p:ext>
  </p:extLst>
</p:nvPr>
```

Note: Modern PowerPoint uses a `p14:media` extension with an embedded relationship for the actual video binary, while the `a:videoFile r:link` may point to a fallback or external source.

### OLE Embedded Objects

OLE objects use `<p:graphicFrame>` with a special graphic data URI:

```xml
<a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole">
  <p:oleObj spid="_x0000_s1027" name="Worksheet" r:id="rId5"
            imgW="3048000" imgH="2286000" progId="Excel.Sheet.12">
    <p:embed/>
    <!-- OR for linked: <p:link autoUpdate="1"/> -->
  </p:oleObj>
</a:graphicData>
```

| Attribute | Description |
|-----------|-------------|
| `r:id` | Relationship to the embedded object binary |
| `progId` | OLE program identifier (e.g., `Excel.Sheet.12`, `Word.Document.12`) |
| `imgW`, `imgH` | Display image dimensions (EMU) |

The binary data is stored in `ppt/embeddings/`. An associated image representation is typically also stored for preview rendering.

### Charts

Charts are embedded via `<p:graphicFrame>`:

```xml
<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
           r:id="rId6"/>
</a:graphicData>
```

Chart XML is stored in `ppt/charts/chart1.xml` with its own relationship and content type.

### Media Content Types

| Format | Content Type | Extension |
|--------|-------------|-----------|
| MP4 | `video/mp4` | `.mp4` |
| MP3 | `audio/mpeg` | `.mp3` |
| WAV | `audio/wav` | `.wav` |
| WMV | `video/x-ms-wmv` | `.wmv` |
| M4A | `audio/mp4` | `.m4a` |

---

## 15. Tables in Slides

Tables in PresentationML are DrawingML constructs embedded within a `p:graphicFrame`. Unlike WordprocessingML tables, they are drawn objects with absolute positioning.

### Table Structure

```xml
<p:graphicFrame>
  <p:nvGraphicFramePr>
    <p:cNvPr id="4" name="Table 3"/>
    <p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>
    <p:nvPr/>
  </p:nvGraphicFramePr>
  <p:xfrm>
    <a:off x="457200" y="1600200"/>
    <a:ext cx="8229600" cy="2743200"/>
  </p:xfrm>
  <a:graphic>
    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
      <a:tbl>
        <a:tblPr firstRow="1" bandRow="1">
          <a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId>
        </a:tblPr>

        <!-- Column widths -->
        <a:tblGrid>
          <a:gridCol w="2743200"/>
          <a:gridCol w="2743200"/>
          <a:gridCol w="2743200"/>
        </a:tblGrid>

        <!-- Rows -->
        <a:tr h="370840">
          <a:tc>
            <a:txBody>
              <a:bodyPr/><a:lstStyle/>
              <a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>Header 1</a:t></a:r></a:p>
            </a:txBody>
            <a:tcPr>
              <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>
              <a:lnL w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:lnL>
              <a:lnR w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:lnR>
              <a:lnT w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:lnT>
              <a:lnB w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:lnB>
            </a:tcPr>
          </a:tc>
          <a:tc>...</a:tc>
          <a:tc>...</a:tc>
        </a:tr>
        <a:tr h="370840">
          <a:tc>...</a:tc>
          <a:tc>...</a:tc>
          <a:tc>...</a:tc>
        </a:tr>
      </a:tbl>
    </a:graphicData>
  </a:graphic>
</p:graphicFrame>
```

### Cell Spanning

```xml
<!-- Cell spanning 2 columns -->
<a:tc gridSpan="2">
  <a:txBody>...</a:txBody>
  <a:tcPr/>
</a:tc>
<a:tc hMerge="1"/>  <!-- horizontally merged placeholder -->

<!-- Cell spanning 2 rows -->
<a:tc rowSpan="2">
  <a:txBody>...</a:txBody>
  <a:tcPr/>
</a:tc>
<!-- In the next row, the corresponding cell: -->
<a:tc vMerge="1"/>  <!-- vertically merged placeholder -->
```

The `hMerge="1"` and `vMerge="1"` cells are placeholders that must still be present in the XML to maintain the grid structure, but they are not rendered.

### Cell Properties (`a:tcPr`)

| Attribute/Element | Description |
|-------------------|-------------|
| `marL`, `marR`, `marT`, `marB` | Cell margins in EMU |
| `anchor` | Vertical alignment: `t`, `ctr`, `b` |
| `vert` | Text direction (same as `a:bodyPr` vert) |
| `a:lnL`, `a:lnR`, `a:lnT`, `a:lnB` | Border lines (left, right, top, bottom) |
| `a:lnTlToBr`, `a:lnBlToTr` | Diagonal borders |
| `a:solidFill` / `a:gradFill` / `a:noFill` | Cell fill |

### Table Style Properties (`a:tblPr`)

| Attribute | Description |
|-----------|-------------|
| `firstRow` | Apply first-row formatting (header row) |
| `lastRow` | Apply last-row formatting |
| `firstCol` | Apply first-column formatting |
| `lastCol` | Apply last-column formatting |
| `bandRow` | Apply banded row formatting |
| `bandCol` | Apply banded column formatting |
| `rtl` | Right-to-left table |

Table styles are identified by GUID and defined in `ppt/tableStyles.xml` or reference built-in styles. The `tableStyles.xml` file contains the default style GUID and custom style definitions.

---

## 16. Comments

### Legacy Comments (Pre-Office 2021)

#### Comment Authors (`ppt/commentAuthors.xml`)

Root element: `<p:cmAuthorLst>`. One per package (if any comments exist).

```xml
<p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cmAuthor id="0" name="John Doe" initials="JD" lastIdx="3" clrIdx="0"/>
  <p:cmAuthor id="1" name="Jane Smith" initials="JS" lastIdx="1" clrIdx="1"/>
</p:cmAuthorLst>
```

| Attribute | Description |
|-----------|-------------|
| `id` | Unique author ID |
| `name` | Display name |
| `initials` | Author initials |
| `lastIdx` | Last comment index used by this author |
| `clrIdx` | Color index for display |

#### Comments (`ppt/comments/comment1.xml`)

One comments part per slide. Root element: `<p:cmLst>`.

```xml
<p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cm authorId="0" dt="2024-01-15T10:30:00.000" idx="1">
    <p:pos x="4486" y="1342"/>              <!-- Position in slide coordinate hundredths -->
    <p:text>Review this slide layout.</p:text>
  </p:cm>
  <p:cm authorId="1" dt="2024-01-16T14:15:00.000" idx="1">
    <p:pos x="2100" y="3200"/>
    <p:text>Looks good to me.</p:text>
  </p:cm>
</p:cmLst>
```

| Attribute | Description |
|-----------|-------------|
| `authorId` | References `p:cmAuthor/@id` |
| `dt` | ISO 8601 datetime |
| `idx` | Comment index (unique per author) |

### Modern Comments (Office 2021+)

Modern comments use a threaded model stored with the `p228` namespace extension. They support:
- **Threading**: replies to comments form a thread
- **Mentions**: @-mention users within comment text
- **Rich text**: formatting within comments
- **Resolved state**: comments can be marked as resolved
- **Anchoring**: comments can anchor to specific shapes or text ranges

These are stored as extensions (in `p:extLst` elements or separate modern comment parts) and may coexist with legacy comments for backward compatibility.

### Implementation Priority

For initial implementation, focus on legacy comments as they are simpler and more widely supported. Preserve modern comment extensions on round-trip without modification.

---

## 17. Strict vs. Transitional Conformance

### Overview

| Aspect | Transitional | Strict |
|--------|-------------|--------|
| Namespace base | `http://schemas.openxmlformats.org/` | `http://purl.oclc.org/ooxml/` |
| Legacy features | VML, compatibility settings allowed | Removed |
| Default in PowerPoint | Yes (all versions) | Must be explicitly selected (since PP 2013) |
| Real-world prevalence | ~99%+ of files | Very rare |

### Namespace Mapping

| Prefix | Transitional | Strict |
|--------|-------------|--------|
| `p` | `http://schemas.openxmlformats.org/presentationml/2006/main` | `http://purl.oclc.org/ooxml/presentationml/main` |
| `a` | `http://schemas.openxmlformats.org/drawingml/2006/main` | `http://purl.oclc.org/ooxml/drawingml/main` |
| `r` | `http://schemas.openxmlformats.org/officeDocument/2006/relationships` | `http://purl.oclc.org/ooxml/officeDocument/relationships` |

### Relationship Type Mapping

Relationship type URIs also change. The base changes from `http://schemas.openxmlformats.org/officeDocument/2006/relationships/` to `http://purl.oclc.org/ooxml/officeDocument/relationships/`.

### Key Differences in Strict Mode

- VML elements (`v:`, `o:` namespaces) are not permitted.
- Certain deprecated attributes are removed (e.g., legacy color values).
- Some element structures are simplified or constrained.
- Date values must use ISO 8601 format (no legacy date serial numbers).

### Implementation Strategy

1. **Read path**: Detect Strict by checking the root element namespace of `<p:presentation>`. If it starts with `http://purl.oclc.org/`, it is Strict. Maintain a namespace mapping table to normalize Strict namespaces to Transitional equivalents during parsing.
2. **Write path**: Default to Transitional for maximum compatibility. Only produce Strict if explicitly requested.
3. **Relationship resolution**: When parsing Strict files, map relationship type URIs to their Transitional equivalents to use common lookup logic.

---

## 18. Implementation Notes and Common Pitfalls

### Slide Order

The canonical slide order is defined by `<p:sldIdLst>` in `presentation.xml`, NOT by file names or ZIP entry order. File numbers can have gaps after slide deletions. Always use the `sldIdLst` order.

### Placeholder Inheritance is Complex

The placeholder matching and property inheritance chain (slide -> layout -> master -> theme -> presentation defaults) is one of the most complex aspects of PPTX. Key rules:
- Match placeholders by `idx` attribute first, then by `type`.
- An empty `<p:spPr/>` means "inherit everything from the parent level."
- An empty `<a:xfrm/>` inside `p:spPr` means "inherit position/size."
- Text formatting inherits through 8 levels (see Section 9).
- PowerPoint may omit the `type` attribute on slide-level placeholders if `idx` is sufficient.

### Coordinate System

All positions and sizes use EMU (English Metric Units). 1 inch = 914400 EMU. Shapes use absolute positioning from the slide's top-left origin. There is no flow layout -- every shape has explicit coordinates.

### Shape IDs Must Be Unique Per Slide

The `id` attribute on `p:cNvPr` must be unique within each slide part (not globally). IDs start at 1 (the shape tree root). Animation targets reference shapes by these IDs via `spid`.

### Relationship ID Generation

Relationship IDs (`rId1`, `rId2`, etc.) must be unique within each `.rels` file but can repeat across different `.rels` files. When writing, generate sequential IDs. When reading, always resolve IDs relative to the containing part's `.rels`.

### Unit Systems Summary

| Context | Unit | Conversion |
|---------|------|------------|
| Position/size | EMU | 1" = 914400, 1pt = 12700, 1px@96dpi = 9525 |
| Font size (`a:rPr/@sz`) | 1/100 pt | `sz="2400"` = 24pt |
| Rotation (`a:xfrm/@rot`) | 1/60000 degree | 90deg = 5400000 (clockwise) |
| Percentages (color, scale) | 1/1000 % | 100000 = 100% |
| Line spacing (`a:spcPct`) | 1/1000 % | 100000 = single spacing |
| Space points (`a:spcPts`) | 1/100 pt | `val="600"` = 6pt |

### Z-Order

Shapes in the shape tree are rendered in document order: first child = bottom (back), last child = top (front).

### Empty Text Bodies

Shapes that have no visible text may still contain empty `<p:txBody>` elements with a single empty `<a:p/>`. This is normal. An `<a:endParaRPr>` element in an empty paragraph stores the formatting that would apply if the user typed in that shape.

### Markup Compatibility (MCE)

PowerPoint files frequently use `mc:AlternateContent` for backward compatibility:

```xml
<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <mc:Choice Requires="p14">
    <!-- Modern content using p14 namespace -->
  </mc:Choice>
  <mc:Fallback>
    <!-- Fallback for older readers -->
  </mc:Fallback>
</mc:AlternateContent>
```

Strategy: if you support the required namespace, use the `Choice` branch; otherwise, use `Fallback`. If no `Fallback` is present and you do not support the namespace, skip the element.

### Round-Trip Fidelity

For write-back scenarios, preserve unrecognized elements and attributes. PowerPoint extensions (p14, p15, p16, a14, a16, etc.) add features beyond the base spec. Dropping these on round-trip will cause data loss. Strategy: store unknown XML subtrees as opaque blobs keyed by their location in the document tree.

### Default Values

Many attributes have implicit defaults when omitted. Always apply these during parsing rather than treating omitted attributes as null:

| Attribute | Default |
|-----------|---------|
| `show` (slide) | `true` |
| `showMasterSp` | `true` |
| `TargetMode` (relationship) | `Internal` |
| `algn` (paragraph) | `l` (left) |
| `b` (bold) | `0` (false) |
| `i` (italic) | `0` (false) |
| `u` (underline) | `none` |
| `strike` | `noStrike` |
| `advClick` (transition) | `true` |

### Text Extraction Strategy

For LLM/text extraction:
1. Iterate slides in `sldIdLst` order.
2. For each slide, walk the shape tree.
3. Shapes have no guaranteed reading order. Sort spatially (top-to-bottom, left-to-right) for coherent text output.
4. Extract text from `p:txBody` of each shape.
5. Include notes text from associated notes slides.
6. Table cells: iterate row by row, cell by cell.
7. Group shapes: recursively descend into `p:grpSp`.

### Minimum Valid PPTX for Writing

To create a minimal valid PPTX that opens in PowerPoint, you need these 11 parts:

1. `[Content_Types].xml` with overrides for presentation, slide, layout, master, and theme.
2. `_rels/.rels` pointing to `ppt/presentation.xml`.
3. `ppt/presentation.xml` with `sldMasterIdLst`, `sldIdLst`, and `sldSz`.
4. `ppt/_rels/presentation.xml.rels` linking to slide, master, and theme.
5. `ppt/slides/slide1.xml` with at least an empty `p:spTree`.
6. `ppt/slides/_rels/slide1.xml.rels` linking to a layout.
7. `ppt/slideLayouts/slideLayout1.xml` linking to a master.
8. `ppt/slideLayouts/_rels/slideLayout1.xml.rels` linking to master.
9. `ppt/slideMasters/slideMaster1.xml` with `clrMap` and `sldLayoutIdLst`.
10. `ppt/slideMasters/_rels/slideMaster1.xml.rels` linking to layout(s) and theme.
11. `ppt/theme/theme1.xml` with `clrScheme`, `fontScheme`, and `fmtScheme`.

### Common File Corruption Causes

- Missing or incorrect `[Content_Types].xml` entries.
- Relationship ID mismatch (referencing an `rId` that does not exist in the `.rels`).
- Duplicate shape IDs within a single slide.
- Invalid XML (unclosed tags, invalid characters).
- Missing required parts (theme, slide master).
- Incorrect relative paths in relationships (e.g., forgetting `../` when referencing from `slides/` to `slideLayouts/`).

### Differences from DOCX/XLSX

| Aspect | PPTX | DOCX | XLSX |
|--------|------|------|------|
| Content model | Shape tree (absolute positioning) | Document flow (paragraphs, tables) | Cell grid (rows, columns) |
| Text container | `txBody` in shapes | Document body with `w:p`/`w:r` | Shared string table + cell values |
| Inheritance | Slide -> Layout -> Master -> Theme | Styles + direct formatting | Cell styles + number formats |
| Main namespace | PresentationML (`p:`) | WordprocessingML (`w:`) | SpreadsheetML (`sml:`) |
| Tables | DrawingML `a:tbl` in graphic frames | WordprocessingML `w:tbl` | Native cell grid |
| Images | `p:pic` in shape tree | `w:drawing` + DrawingML | Drawing overlay |

### ID Value Constraints

| ID Type | Constraint |
|---------|-----------|
| `p:sldId/@id` | Unique, >= 256 |
| `p:sldMasterId/@id` | Unique, >= 2147483648 (0x80000000) |
| `p:sldLayoutId/@id` | Unique, >= 2147483649 |
| `p:cNvPr/@id` | Unique within each slide part, >= 1 |
| Relationship `Id` | Unique within each `.rels` file |
| Time node `id` | Unique within each slide's timing tree |

---

## Appendix A: File Extension Variants

| Extension | Content Type (main part) | Description |
|-----------|------------------------|-------------|
| `.pptx` | `...presentationml.presentation.main+xml` | Standard presentation |
| `.pptm` | `...ms-powerpoint.presentation.macroEnabled.main+xml` | Macro-enabled presentation |
| `.potx` | `...presentationml.template.main+xml` | Presentation template |
| `.potm` | `...ms-powerpoint.template.macroEnabled.main+xml` | Macro-enabled template |
| `.ppsx` | `...presentationml.slideshow.main+xml` | Slideshow (opens in show mode) |
| `.ppsm` | `...ms-powerpoint.slideshow.macroEnabled.main+xml` | Macro-enabled slideshow |
| `.ppam` | `...ms-powerpoint.addin.macroEnabled.main+xml` | PowerPoint add-in |
| `.sldx` | `...presentationml.slide+xml` | Single slide |

## Appendix B: Element Hierarchy Quick Reference

```
p:presentation
+-- p:sldMasterIdLst / p:sldMasterId [@id, @r:id]
+-- p:notesMasterIdLst / p:notesMasterId [@r:id]
+-- p:handoutMasterIdLst / p:handoutMasterId [@r:id]
+-- p:sldIdLst / p:sldId [@id, @r:id]
+-- p:sldSz [@cx, @cy]
+-- p:notesSz [@cx, @cy]
+-- p:defaultTextStyle / a:lvl1pPr..a:lvl9pPr

p:sld / p:sldLayout / p:sldMaster
+-- p:cSld
|   +-- p:bg (p:bgPr or p:bgRef)
|   +-- p:spTree
|       +-- p:sp (auto shape)
|       |   +-- p:nvSpPr (p:cNvPr, p:cNvSpPr, p:nvPr[p:ph])
|       |   +-- p:spPr (a:xfrm, a:prstGeom, fill, line, effects)
|       |   +-- p:style (theme style ref)
|       |   +-- p:txBody (a:bodyPr, a:lstStyle, a:p[])
|       +-- p:pic
|       |   +-- p:nvPicPr / p:blipFill / p:spPr
|       +-- p:grpSp (recursive)
|       |   +-- p:nvGrpSpPr / p:grpSpPr (a:xfrm with chOff/chExt)
|       |   +-- [child shapes]
|       +-- p:graphicFrame (table, chart, SmartArt, OLE)
|       |   +-- p:nvGraphicFramePr / p:xfrm / a:graphic / a:graphicData[@uri]
|       +-- p:cxnSp (connector)
|           +-- p:nvCxnSpPr (a:stCxn, a:endCxn) / p:spPr
+-- p:clrMapOvr / p:clrMap
+-- p:transition
+-- p:timing
```

## Appendix C: Implementation Checklist

### Parser Minimum Requirements

- [ ] Read ZIP archive (Deflate + Store)
- [ ] Parse `[Content_Types].xml`
- [ ] Parse `.rels` files and build relationship graph
- [ ] Parse `presentation.xml` (slide list, masters, size)
- [ ] Parse slides (`p:sld` with `p:cSld` and `p:spTree`)
- [ ] Parse shape tree: `sp`, `pic`, `grpSp`, `cxnSp`, `graphicFrame`
- [ ] Parse text body: `txBody`, `bodyPr`, `p`, `r`, `rPr`, `t`
- [ ] Parse shape properties: `spPr`, `xfrm`, `prstGeom`, fills, lines
- [ ] Parse slide masters and layouts
- [ ] Parse theme (colors, fonts, format scheme)
- [ ] Resolve style inheritance chain
- [ ] Handle placeholder inheritance (idx/type matching)

### Writer Minimum Requirements

- [ ] Create valid ZIP archive
- [ ] Generate `[Content_Types].xml` with correct content types
- [ ] Generate all `.rels` files with unique relationship IDs
- [ ] Generate `presentation.xml` with slide list and master references
- [ ] Generate at least one slide master, layout, and theme
- [ ] Generate slides with valid shape trees
- [ ] Ensure all `id` attributes on `p:cNvPr` are unique per slide
- [ ] Ensure `p:sldId/@id` values are unique and >= 256
- [ ] Ensure `p:sldMasterId/@id` values are unique and >= 2147483648

---

*This document is derived from public specifications: ISO/IEC 29500:2008-2016 and ECMA-376 5th Edition. For the authoritative and complete reference, consult the full standards.*
