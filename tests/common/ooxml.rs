//! Small OOXML packages built in code: `.docx`, `.pptx` and `.xlsx` that
//! draw pictures, and the parts they are made of.

use std::io::{Cursor, Write};

#[allow(dead_code)]
pub const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
#[allow(dead_code)]
pub const PKG_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
#[allow(dead_code)]
pub const WML: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
#[allow(dead_code)]
pub const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
#[allow(dead_code)]
pub const DML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
#[allow(dead_code)]
pub const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

/// The inflated size of each picture.
#[allow(dead_code)]
pub const PICTURE_LEN: usize = 2 << 20;

/// A "PNG" that deflates to a few kilobytes: a signature, then a pattern.
#[allow(dead_code)]
pub fn picture(n: usize) -> Vec<u8> {
    let mut data = vec![0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A];
    data.extend((0..PICTURE_LEN - 8).map(|i| ((i / 4096 + n) % 251) as u8));
    data
}

#[allow(dead_code)]
pub fn zip_of(parts: Vec<(String, Vec<u8>)>) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let deflated = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);
    for (name, data) in parts {
        zip.start_file(name, deflated).unwrap();
        zip.write_all(&data).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

#[allow(dead_code)]
pub fn rels(entries: &[(String, &str, String)]) -> Vec<u8> {
    let mut s = format!(r#"<Relationships xmlns="{PKG_REL}">"#);
    for (id, ty, target) in entries {
        s.push_str(&format!(r#"<Relationship Id="{id}" Type="{REL}/{ty}" Target="{target}"/>"#));
    }
    s.push_str("</Relationships>");
    s.into_bytes()
}

#[allow(dead_code)]
pub const CT: &str = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#;

/// A `.docx` whose body draws picture `i % pictures` in each of `draws`
/// paragraphs, between lines of text.
#[allow(dead_code)]
pub fn docx(pictures: usize, draws: usize) -> Vec<u8> {
    docx_with_media_names(pictures, draws, ".png")
}

/// [`docx`] with the media parts named `img{n}{suffix}` (`""` for parts
/// with no extension, whose format is told from their bytes).
#[allow(dead_code)]
pub fn docx_with_media_names(pictures: usize, draws: usize, suffix: &str) -> Vec<u8> {
    let mut body = String::new();
    for i in 0..draws {
        body.push_str(&format!(
            r#"<w:p><w:r><w:t>Paragraph {i} before a picture.</w:t></w:r></w:p><w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="100" cy="100"/><wp:docPr id="{}" name="p" descr="picture {i}"/><a:graphic xmlns:a="{DML}"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:blipFill><a:blip r:embed="rIdP{}"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"#,
            i + 1,
            i % pictures
        ));
    }
    let doc_rels: Vec<_> = (0..pictures)
        .map(|n| (format!("rIdP{n}"), "image", format!("media/img{n}{suffix}")))
        .collect();
    let mut parts = vec![
        ("[Content_Types].xml".to_string(), CT.as_bytes().to_vec()),
        (
            "_rels/.rels".to_string(),
            rels(&[("rId1".into(), "officeDocument", "word/document.xml".into())]),
        ),
        (
            "word/document.xml".to_string(),
            format!(r#"<w:document xmlns:w="{WML}" xmlns:r="{REL}"><w:body>{body}</w:body></w:document>"#)
                .into_bytes(),
        ),
        ("word/_rels/document.xml.rels".to_string(), rels(&doc_rels)),
    ];
    for n in 0..pictures {
        parts.push((format!("word/media/img{n}{suffix}"), picture(n)));
    }
    zip_of(parts)
}

/// A one-slide `.pptx` drawing picture `i % pictures` `draws` times.
#[allow(dead_code)]
pub fn pptx(pictures: usize, draws: usize) -> Vec<u8> {
    let mut shapes = String::new();
    for i in 0..draws {
        shapes.push_str(&format!(
            r#"<p:pic><p:nvPicPr><p:cNvPr id="{}" name="P{i}" descr="picture {i}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdP{}"/></p:blipFill><p:spPr><a:xfrm><a:off x="{}" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr></p:pic>"#,
            i + 2,
            i % pictures,
            i * 1000
        ));
    }
    let slide = format!(
        r#"<p:sld xmlns:p="{PML}" xmlns:a="{DML}" xmlns:r="{REL}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="999" name="T"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>Slide text</a:t></a:r></a:p></p:txBody></p:sp>{shapes}</p:spTree></p:cSld></p:sld>"#
    );
    let slide_rels: Vec<_> = (0..pictures)
        .map(|n| (format!("rIdP{n}"), "image", format!("../media/img{n}.png")))
        .collect();
    let mut parts = vec![
        ("[Content_Types].xml".to_string(), CT.as_bytes().to_vec()),
        (
            "_rels/.rels".to_string(),
            rels(&[("rId1".into(), "officeDocument", "ppt/presentation.xml".into())]),
        ),
        (
            "ppt/presentation.xml".to_string(),
            format!(r#"<p:presentation xmlns:p="{PML}" xmlns:r="{REL}"><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>"#)
                .into_bytes(),
        ),
        (
            "ppt/_rels/presentation.xml.rels".to_string(),
            rels(&[("rId2".into(), "slide", "slides/slide1.xml".into())]),
        ),
        ("ppt/slides/slide1.xml".to_string(), slide.into_bytes()),
        ("ppt/slides/_rels/slide1.xml.rels".to_string(), rels(&slide_rels)),
    ];
    for n in 0..pictures {
        parts.push((format!("ppt/media/img{n}.png"), picture(n)));
    }
    zip_of(parts)
}

/// A one-sheet `.xlsx` whose drawing anchors picture `i % pictures` `draws`
/// times.
#[allow(dead_code)]
pub fn xlsx(pictures: usize, draws: usize) -> Vec<u8> {
    let mut drawing = String::from(
        r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
    );
    for i in 0..draws {
        drawing.push_str(&format!(
            r#"<xdr:oneCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{i}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="100" cy="100"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{}" name="P{i}" descr="picture {i}"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdP{}"/></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>"#,
            i + 2,
            i % pictures
        ));
    }
    drawing.push_str("</xdr:wsDr>");
    let drawing_rels: Vec<_> = (0..pictures)
        .map(|n| (format!("rIdP{n}"), "image", format!("../media/img{n}.png")))
        .collect();
    let mut parts = vec![
        ("[Content_Types].xml".to_string(), CT.as_bytes().to_vec()),
        (
            "_rels/.rels".to_string(),
            rels(&[("rId1".into(), "officeDocument", "xl/workbook.xml".into())]),
        ),
        (
            "xl/workbook.xml".to_string(),
            format!(r#"<workbook xmlns="{SML}" xmlns:r="{REL}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>"#)
                .into_bytes(),
        ),
        (
            "xl/_rels/workbook.xml.rels".to_string(),
            rels(&[("rId1".into(), "worksheet", "worksheets/sheet1.xml".into())]),
        ),
        (
            "xl/worksheets/sheet1.xml".to_string(),
            format!(r#"<worksheet xmlns="{SML}" xmlns:r="{REL}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>cell text</t></is></c></row></sheetData><drawing r:id="rId1"/></worksheet>"#)
                .into_bytes(),
        ),
        (
            "xl/worksheets/_rels/sheet1.xml.rels".to_string(),
            rels(&[("rId1".into(), "drawing", "../drawings/drawing1.xml".into())]),
        ),
        ("xl/drawings/drawing1.xml".to_string(), drawing.into_bytes()),
        ("xl/drawings/_rels/drawing1.xml.rels".to_string(), rels(&drawing_rels)),
    ];
    for n in 0..pictures {
        parts.push((format!("xl/media/img{n}.png"), picture(n)));
    }
    zip_of(parts)
}

/// A `.docx` whose body is `body` (WordprocessingML, `w:` and `r:` bound).
#[allow(dead_code)]
pub fn docx_with_body(body: &str) -> Vec<u8> {
    zip_of(vec![
        ("[Content_Types].xml".to_string(), CT.as_bytes().to_vec()),
        (
            "_rels/.rels".to_string(),
            rels(&[("rId1".into(), "officeDocument", "word/document.xml".into())]),
        ),
        (
            "word/document.xml".to_string(),
            format!(r#"<w:document xmlns:w="{WML}" xmlns:r="{REL}"><w:body>{body}</w:body></w:document>"#)
                .into_bytes(),
        ),
    ])
}

/// An `.xlsx` with one sheet per entry of `sheet_data`, each the content of
/// that sheet's `<sheetData>`.
#[allow(dead_code)]
pub fn xlsx_with_sheets(sheet_data: &[String]) -> Vec<u8> {
    let mut names = String::new();
    let mut book_rels = Vec::new();
    let mut parts = vec![
        ("[Content_Types].xml".to_string(), CT.as_bytes().to_vec()),
        (
            "_rels/.rels".to_string(),
            rels(&[("rId1".into(), "officeDocument", "xl/workbook.xml".into())]),
        ),
    ];
    for (i, data) in sheet_data.iter().enumerate() {
        let n = i + 1;
        names.push_str(&format!(r#"<sheet name="S{n}" sheetId="{n}" r:id="rId{n}"/>"#));
        book_rels.push((format!("rId{n}"), "worksheet", format!("worksheets/sheet{n}.xml")));
        parts.push((
            format!("xl/worksheets/sheet{n}.xml"),
            format!(r#"<worksheet xmlns="{SML}" xmlns:r="{REL}"><sheetData>{data}</sheetData></worksheet>"#)
                .into_bytes(),
        ));
    }
    parts.push((
        "xl/workbook.xml".to_string(),
        format!(r#"<workbook xmlns="{SML}" xmlns:r="{REL}"><sheets>{names}</sheets></workbook>"#)
            .into_bytes(),
    ));
    parts.push(("xl/_rels/workbook.xml.rels".to_string(), rels(&book_rels)));
    zip_of(parts)
}

/// A `.pptx` of `slides` slides, each drawing picture `s % pictures` once
/// besides a line of text.
#[allow(dead_code)]
pub fn pptx_with_slides(slides: usize, pictures: usize) -> Vec<u8> {
    let mut ct = String::from(
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>"#,
    );
    let mut ids = String::new();
    let mut pres_rels = Vec::new();
    let mut parts = Vec::new();
    for s in 1..=slides {
        ct.push_str(&format!(r#"<Override PartName="/ppt/slides/slide{s}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>"#));
        ids.push_str(&format!(r#"<p:sldId id="{}" r:id="rId{s}"/>"#, 255 + s));
        pres_rels.push((format!("rId{s}"), "slide", format!("slides/slide{s}.xml")));
        let slide = format!(
            r#"<p:sld xmlns:p="{PML}" xmlns:a="{DML}" xmlns:r="{REL}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="T"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>Slide {s} text</a:t></a:r></a:p></p:txBody></p:sp><p:pic><p:nvPicPr><p:cNvPr id="3" name="P" descr="picture on slide {s}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdP"/></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr></p:pic></p:spTree></p:cSld></p:sld>"#
        );
        parts.push((format!("ppt/slides/slide{s}.xml"), slide.into_bytes()));
        parts.push((
            format!("ppt/slides/_rels/slide{s}.xml.rels"),
            rels(&[("rIdP".into(), "image", format!("../media/img{}.png", s % pictures))]),
        ));
    }
    ct.push_str("</Types>");
    parts.push(("[Content_Types].xml".to_string(), ct.into_bytes()));
    parts.push((
        "_rels/.rels".to_string(),
        rels(&[("rId1".into(), "officeDocument", "ppt/presentation.xml".into())]),
    ));
    parts.push((
        "ppt/presentation.xml".to_string(),
        format!(r#"<p:presentation xmlns:p="{PML}" xmlns:r="{REL}"><p:sldIdLst>{ids}</p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>"#)
            .into_bytes(),
    ));
    parts.push(("ppt/_rels/presentation.xml.rels".to_string(), rels(&pres_rels)));
    for n in 0..pictures {
        parts.push((format!("ppt/media/img{n}.png"), picture(n)));
    }
    zip_of(parts)
}

/// An `.xlsx` of `sheets` sheets, each holding a cell and a drawing that
/// anchors picture `s % pictures` once.
#[allow(dead_code)]
pub fn xlsx_with_picture_sheets(sheets: usize, pictures: usize) -> Vec<u8> {
    let mut names = String::new();
    let mut book_rels = Vec::new();
    let mut parts = vec![
        ("[Content_Types].xml".to_string(), CT.as_bytes().to_vec()),
        (
            "_rels/.rels".to_string(),
            rels(&[("rId1".into(), "officeDocument", "xl/workbook.xml".into())]),
        ),
    ];
    for s in 1..=sheets {
        names.push_str(&format!(r#"<sheet name="S{s}" sheetId="{s}" r:id="rId{s}"/>"#));
        book_rels.push((format!("rId{s}"), "worksheet", format!("worksheets/sheet{s}.xml")));
        parts.push((
            format!("xl/worksheets/sheet{s}.xml"),
            format!(r#"<worksheet xmlns="{SML}" xmlns:r="{REL}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>sheet {s}</t></is></c></row></sheetData><drawing r:id="rIdD"/></worksheet>"#)
                .into_bytes(),
        ));
        parts.push((
            format!("xl/worksheets/_rels/sheet{s}.xml.rels"),
            rels(&[("rIdD".into(), "drawing", format!("../drawings/drawing{s}.xml"))]),
        ));
        parts.push((
            format!("xl/drawings/drawing{s}.xml"),
            format!(r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="{DML}" xmlns:r="{REL}"><xdr:oneCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="100" cy="100"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="P" descr="picture on sheet {s}"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdP"/></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>"#)
                .into_bytes(),
        ));
        parts.push((
            format!("xl/drawings/_rels/drawing{s}.xml.rels"),
            rels(&[("rIdP".into(), "image", format!("../media/img{}.png", s % pictures))]),
        ));
    }
    parts.push((
        "xl/workbook.xml".to_string(),
        format!(r#"<workbook xmlns="{SML}" xmlns:r="{REL}"><sheets>{names}</sheets></workbook>"#)
            .into_bytes(),
    ));
    parts.push(("xl/_rels/workbook.xml.rels".to_string(), rels(&book_rels)));
    for n in 0..pictures {
        parts.push((format!("xl/media/img{n}.png"), picture(n)));
    }
    zip_of(parts)
}

/// A one-slide `.pptx` of `frames` OLE objects, each previewed by picture 0.
#[allow(dead_code)]
pub fn pptx_with_ole_previews(frames: usize) -> Vec<u8> {
    let mut shapes = String::new();
    for i in 0..frames {
        shapes.push_str(&format!(
            r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="{}" name="O{i}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole"><p:oleObj name="W" progId="Excel.Sheet.12"><p:embed/><p:pic><p:nvPicPr><p:cNvPr id="0" name=""/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdP0"/></p:blipFill><p:spPr/></p:pic></p:oleObj></a:graphicData></a:graphic></p:graphicFrame>"#,
            i + 2
        ));
    }
    let slide = format!(
        r#"<p:sld xmlns:p="{PML}" xmlns:a="{DML}" xmlns:r="{REL}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{shapes}</p:spTree></p:cSld></p:sld>"#
    );
    zip_of(vec![
        ("[Content_Types].xml".to_string(), CT.as_bytes().to_vec()),
        (
            "_rels/.rels".to_string(),
            rels(&[("rId1".into(), "officeDocument", "ppt/presentation.xml".into())]),
        ),
        (
            "ppt/presentation.xml".to_string(),
            format!(r#"<p:presentation xmlns:p="{PML}" xmlns:r="{REL}"><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>"#)
                .into_bytes(),
        ),
        (
            "ppt/_rels/presentation.xml.rels".to_string(),
            rels(&[("rId2".into(), "slide", "slides/slide1.xml".into())]),
        ),
        ("ppt/slides/slide1.xml".to_string(), slide.into_bytes()),
        (
            "ppt/slides/_rels/slide1.xml.rels".to_string(),
            rels(&[("rIdP0".into(), "image", "../media/img0.png".into())]),
        ),
        ("ppt/media/img0.png".to_string(), picture(0)),
    ])
}
