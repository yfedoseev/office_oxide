"""Generate minimal synthetic OOXML documents that scale one structure at a
time, so a superlinear parser path shows up as a log-log slope > 1 instead
of hiding behind "big file = slow".

  python3 gen_scaling.py OUTDIR DIM N        # one file
  python3 gen_scaling.py --list              # every dimension

Each dimension isolates one thing a parser might look up per element:
docx  para runs table tables hyperlinks bookmarks styles comments footnotes
      images lists sections
xlsx  rows sst cols merges styles comments hyperlinks sheets formulas dv cf
      dates sparse
pptx  slides shapes runs table notes
"""
import os, struct, sys, zipfile, zlib

CT = "http://schemas.openxmlformats.org/package/2006/content-types"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

def rels(items):
    body = "".join(f'<Relationship Id="{i}" Type="{t}" Target="{g}"{" TargetMode=\"External\"" if ext else ""}/>'
                   for i, t, g, ext in items)
    return f'{XML}<Relationships xmlns="{PR}">{body}</Relationships>'

def ctypes(defaults, overrides):
    d = "".join(f'<Default Extension="{e}" ContentType="{c}"/>' for e, c in defaults)
    o = "".join(f'<Override PartName="{p}" ContentType="{c}"/>' for p, c in overrides)
    return f'{XML}<Types xmlns="{CT}">{d}{o}</Types>'

def png1x1():
    def chunk(t, d):
        return struct.pack(">I", len(d)) + t + d + struct.pack(">I", zlib.crc32(t + d) & 0xffffffff)
    ihdr = struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)
    idat = zlib.compress(b"\x00\xff\x00\x00")
    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr) + chunk(b"IDAT", idat) + chunk(b"IEND", b"")

def write(path, parts):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts:
            z.writestr(name, data)

# ------------------------------------------------------------------ DOCX
def docx(dim, n):
    body = []; extra_parts = []; doc_rels = []; ct_over = []
    styles = f'{XML}<w:styles xmlns:w="{W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
    p = lambda text, ppr="": f'<w:p>{ppr}<w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
    if dim == "para":
        body = [p(f"Paragraph {i} with a few words of body text.") for i in range(n)]
    elif dim == "runs":
        body = ['<w:p>' + "".join(f'<w:r>{"<w:rPr><w:b/></w:rPr>" if i % 2 else ""}<w:t xml:space="preserve">run{i} </w:t></w:r>' for i in range(n)) + '</w:p>']
    elif dim == "table":
        rows = "".join('<w:tr>' + "".join(f'<w:tc><w:p><w:r><w:t>r{i}c{j}</w:t></w:r></w:p></w:tc>' for j in range(8)) + '</w:tr>' for i in range(n))
        body = [f'<w:tbl><w:tblPr/><w:tblGrid>{"<w:gridCol w:w=\"1000\"/>" * 8}</w:tblGrid>{rows}</w:tbl>']
    elif dim == "tables":
        one = '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>' + "".join('<w:tr>' + "".join(f'<w:tc><w:p><w:r><w:t>t{{i}}r{r}c{c}</w:t></w:r></w:p></w:tc>' for c in range(3)) + '</w:tr>' for r in range(3)) + '</w:tbl>'
        body = [one.replace("{i}", str(i)) + p("between") for i in range(n)]
    elif dim == "hyperlinks":
        for i in range(n):
            doc_rels.append((f"rIdH{i}", R + "/hyperlink", f"https://example.com/{i}", True))
            body.append(f'<w:p><w:hyperlink r:id="rIdH{i}"><w:r><w:t>link {i}</w:t></w:r></w:hyperlink></w:p>')
    elif dim == "bookmarks":
        body = [f'<w:p><w:bookmarkStart w:id="{i}" w:name="bm{i}"/><w:r><w:t>para {i}</w:t></w:r><w:bookmarkEnd w:id="{i}"/></w:p>' for i in range(n)]
    elif dim == "styles":
        styles += "".join(f'<w:style w:type="paragraph" w:styleId="S{i}"><w:name w:val="Style {i}"/><w:basedOn w:val="Normal"/><w:rPr><w:b/></w:rPr></w:style>' for i in range(n))
        body = [p(f"para {i}", f'<w:pPr><w:pStyle w:val="S{i}"/></w:pPr>') for i in range(n)]
    elif dim == "comments":
        cm = "".join(f'<w:comment w:id="{i}" w:author="a"><w:p><w:r><w:t>comment {i}</w:t></w:r></w:p></w:comment>' for i in range(n))
        extra_parts.append(("word/comments.xml", f'{XML}<w:comments xmlns:w="{W}">{cm}</w:comments>'))
        doc_rels.append(("rIdC", R + "/comments", "comments.xml", False))
        ct_over.append(("/word/comments.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"))
        body = [f'<w:p><w:commentRangeStart w:id="{i}"/><w:r><w:t>para {i}</w:t></w:r><w:commentRangeEnd w:id="{i}"/><w:r><w:commentReference w:id="{i}"/></w:r></w:p>' for i in range(n)]
    elif dim == "footnotes":
        fn = "".join(f'<w:footnote w:id="{i+1}"><w:p><w:r><w:t>footnote {i}</w:t></w:r></w:p></w:footnote>' for i in range(n))
        extra_parts.append(("word/footnotes.xml", f'{XML}<w:footnotes xmlns:w="{W}">{fn}</w:footnotes>'))
        doc_rels.append(("rIdF", R + "/footnotes", "footnotes.xml", False))
        ct_over.append(("/word/footnotes.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"))
        body = [f'<w:p><w:r><w:t>para {i}</w:t></w:r><w:r><w:footnoteReference w:id="{i+1}"/></w:r></w:p>' for i in range(n)]
    elif dim == "images":
        png = png1x1()
        for i in range(n):
            extra_parts.append((f"word/media/image{i}.png", png))
            doc_rels.append((f"rIdI{i}", R + "/image", f"media/image{i}.png", False))
            body.append(f'<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="100" cy="100"/><wp:docPr id="{i+1}" name="img{i}"/><a:graphic xmlns:a="{A}"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="{i+1}" name="img{i}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdI{i}"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>')
    elif dim == "lists":
        num = f'{XML}<w:numbering xmlns:w="{W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>'
        extra_parts.append(("word/numbering.xml", num))
        doc_rels.append(("rIdN", R + "/numbering", "numbering.xml", False))
        ct_over.append(("/word/numbering.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"))
        body = [p(f"item {i}", '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>') for i in range(n)]
    elif dim == "sections":
        for i in range(n):
            extra_parts.append((f"word/header{i}.xml", f'{XML}<w:hdr xmlns:w="{W}"><w:p><w:r><w:t>header {i}</w:t></w:r></w:p></w:hdr>'))
            doc_rels.append((f"rIdS{i}", R + "/header", f"header{i}.xml", False))
            ct_over.append((f"/word/header{i}.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"))
            body.append(p(f"section {i}", f'<w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rIdS{i}"/></w:sectPr></w:pPr>'))
    else:
        raise SystemExit(f"unknown docx dim {dim}")
    styles += "</w:styles>"
    doc = f'{XML}<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body>{"".join(body)}<w:sectPr/></w:body></w:document>'
    doc_rels.append(("rIdSt", R + "/styles", "styles.xml", False))
    parts = [
        ("[Content_Types].xml", ctypes([("rels", "application/vnd.openxmlformats-package.relationships+xml"), ("xml", "application/xml"), ("png", "image/png")],
                                       [("/word/document.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"),
                                        ("/word/styles.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml")] + ct_over)),
        ("_rels/.rels", rels([("rId1", R + "/officeDocument", "word/document.xml", False)])),
        ("word/document.xml", doc),
        ("word/styles.xml", styles),
        ("word/_rels/document.xml.rels", rels(doc_rels)),
    ] + extra_parts
    return parts

# ------------------------------------------------------------------ XLSX
def col(i):
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s

def xlsx(dim, n):
    sheets = []  # (name, sheetdata xml, extra sheet xml after sheetData, sheet rels)
    sst = None; styles_xml = None; extra_parts = []; ct_over = []
    ncols = 10
    def grid(rows, cols, cell):
        return "".join(f'<row r="{r+1}">' + "".join(cell(r, c) for c in range(cols)) + '</row>' for r in range(rows))
    if dim == "rows":
        sheets.append(("S", grid(n, ncols, lambda r, c: f'<c r="{col(c)}{r+1}"><v>{r*ncols+c}</v></c>'), "", []))
    elif dim == "sst":
        k = n * ncols
        sst = f'{XML}<sst xmlns="{S}" count="{k}" uniqueCount="{k}">' + "".join(f'<si><t>string {i}</t></si>' for i in range(k)) + '</sst>'
        sheets.append(("S", grid(n, ncols, lambda r, c: f'<c r="{col(c)}{r+1}" t="s"><v>{r*ncols+c}</v></c>'), "", []))
    elif dim == "inline":
        sheets.append(("S", grid(n, ncols, lambda r, c: f'<c r="{col(c)}{r+1}" t="inlineStr"><is><t>s{r}_{c}</t></is></c>'), "", []))
    elif dim == "cols":
        sheets.append(("S", grid(1, n, lambda r, c: f'<c r="{col(c)}{r+1}"><v>{c}</v></c>'), "", []))
    elif dim == "merges":
        m = f'<mergeCells count="{n}">' + "".join(f'<mergeCell ref="A{i+1}:B{i+1}"/>' for i in range(n)) + '</mergeCells>'
        sheets.append(("S", grid(n, 3, lambda r, c: f'<c r="{col(c)}{r+1}"><v>{r}</v></c>'), m, []))
    elif dim == "styles":
        styles_xml = f'{XML}<styleSheet xmlns="{S}"><fonts count="{n}">' + "".join(f'<font><sz val="{8+i%20}"/><name val="Arial"/></font>' for i in range(n)) + f'</fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellXfs count="{n}">' + "".join(f'<xf numFmtId="0" fontId="{i}" applyFont="1"/>' for i in range(n)) + '</cellXfs></styleSheet>'
        sheets.append(("S", grid(n, 1, lambda r, c: f'<c r="A{r+1}" s="{r}"><v>{r}</v></c>'), "", []))
    elif dim == "dates":
        styles_xml = f'{XML}<styleSheet xmlns="{S}"><fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14" applyNumberFormat="1"/></cellXfs></styleSheet>'
        sheets.append(("S", grid(n, 1, lambda r, c: f'<c r="A{r+1}" s="1"><v>{40000 + r % 3000}</v></c>'), "", []))
    elif dim == "comments":
        cm = f'{XML}<comments xmlns="{S}"><authors><author>a</author></authors><commentList>' + "".join(f'<comment ref="A{i+1}" authorId="0"><text><t>c{i}</t></text></comment>' for i in range(n)) + '</commentList></comments>'
        extra_parts.append(("xl/comments1.xml", cm))
        ct_over.append(("/xl/comments1.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"))
        sheets.append(("S", grid(n, 1, lambda r, c: f'<c r="A{r+1}"><v>{r}</v></c>'), '<legacyDrawing r:id="rIdV"/>', [("rIdCm", R + "/comments", "../comments1.xml", False)]))
    elif dim == "hyperlinks":
        hl = "<hyperlinks>" + "".join(f'<hyperlink ref="A{i+1}" r:id="rIdH{i}"/>' for i in range(n)) + "</hyperlinks>"
        sheets.append(("S", grid(n, 1, lambda r, c: f'<c r="A{r+1}"><v>{r}</v></c>'), hl, [(f"rIdH{i}", R + "/hyperlink", f"https://example.com/{i}", True) for i in range(n)]))
    elif dim == "sheets":
        for i in range(n):
            sheets.append((f"Sheet{i}", grid(2, 2, lambda r, c: f'<c r="{col(c)}{r+1}"><v>{r+c}</v></c>'), "", []))
    elif dim == "formulas":
        sheets.append(("S", grid(n, 2, lambda r, c: (f'<c r="A{r+1}"><v>{r}</v></c>' if c == 0 else f'<c r="B{r+1}"><f>A{r+1}*2+SUM(A$1:A{r+1})</f><v>{r*2}</v></c>')), "", []))
    elif dim == "dv":
        dv = f'<dataValidations count="{n}">' + "".join(f'<dataValidation type="list" sqref="A{i+1}"><formula1>"a,b,c"</formula1></dataValidation>' for i in range(n)) + '</dataValidations>'
        sheets.append(("S", grid(n, 1, lambda r, c: f'<c r="A{r+1}"><v>{r}</v></c>'), dv, []))
    elif dim == "cf":
        cf = "".join(f'<conditionalFormatting sqref="A{i+1}"><cfRule type="cellIs" priority="{i+1}" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>' for i in range(n))
        sheets.append(("S", grid(n, 1, lambda r, c: f'<c r="A{r+1}"><v>{r}</v></c>'), cf, []))
    elif dim == "sparse":
        # n cells scattered down the diagonal of a huge declared range.
        sd = "".join(f'<row r="{i*50+1}"><c r="{col(i % 200)}{i*50+1}"><v>{i}</v></c></row>' for i in range(n))
        sheets.append(("S", sd, "", []))
    else:
        raise SystemExit(f"unknown xlsx dim {dim}")
    if styles_xml is None:
        styles_xml = f'{XML}<styleSheet xmlns="{S}"><fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders><cellXfs count="1"><xf numFmtId="0"/></cellXfs></styleSheet>'
    parts = []
    wb_sheets = ""; wb_rels = [("rIdSt", R + "/styles", "styles.xml", False)]
    over = [("/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"),
            ("/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")]
    if sst:
        parts.append(("xl/sharedStrings.xml", sst)); wb_rels.append(("rIdSs", R + "/sharedStrings", "sharedStrings.xml", False))
        over.append(("/xl/sharedStrings.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"))
    for i, (name, sd, after, srels) in enumerate(sheets):
        pn = f"xl/worksheets/sheet{i+1}.xml"
        parts.append((pn, f'{XML}<worksheet xmlns="{S}" xmlns:r="{R}"><sheetData>{sd}</sheetData>{after}</worksheet>'))
        if srels:
            parts.append((f"xl/worksheets/_rels/sheet{i+1}.xml.rels", rels(srels)))
        wb_sheets += f'<sheet name="{name}" sheetId="{i+1}" r:id="rIdW{i}"/>'
        wb_rels.append((f"rIdW{i}", R + "/worksheet", f"worksheets/sheet{i+1}.xml", False))
        over.append((f"/{pn}", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"))
    parts += [
        ("[Content_Types].xml", ctypes([("rels", "application/vnd.openxmlformats-package.relationships+xml"), ("xml", "application/xml")], over + ct_over)),
        ("_rels/.rels", rels([("rId1", R + "/officeDocument", "xl/workbook.xml", False)])),
        ("xl/workbook.xml", f'{XML}<workbook xmlns="{S}" xmlns:r="{R}"><sheets>{wb_sheets}</sheets></workbook>'),
        ("xl/_rels/workbook.xml.rels", rels(wb_rels)),
        ("xl/styles.xml", styles_xml),
    ] + extra_parts
    return parts

# ------------------------------------------------------------------ PPTX
def pptx(dim, n):
    def sp(i, text, runs=1):
        rs = "".join(f'<a:r><a:rPr lang="en-US"{" b=\"1\"" if k % 2 else ""}/><a:t>{text} run{k} </a:t></a:r>' for k in range(runs))
        return f'<p:sp><p:nvSpPr><p:cNvPr id="{i+2}" name="TextBox {i}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="{i*10}"/><a:ext cx="1000000" cy="100000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:p>{rs}</a:p></p:txBody></p:sp>'
    slides = []  # (shapes xml, notes text or None)
    if dim == "slides":
        slides = [(sp(0, f"slide {i}"), None) for i in range(n)]
    elif dim == "shapes":
        slides = ["".join(sp(i, f"shape {i}") for i in range(n))]; slides = [(slides[0], None)]
    elif dim == "runs":
        slides = [(sp(0, "t", runs=n), None)]
    elif dim == "table":
        rows = "".join('<a:tr h="100">' + "".join(f'<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>r{r}c{c}</a:t></a:r></a:p></a:txBody></a:tc>' for c in range(6)) + '</a:tr>' for r in range(n))
        gf = f'<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblGrid>{"<a:gridCol w=\"100\"/>" * 6}</a:tblGrid>{rows}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>'
        slides = [(gf, None)]
    elif dim == "notes":
        slides = [(sp(0, f"slide {i}"), f"notes for slide {i}") for i in range(n)]
    else:
        raise SystemExit(f"unknown pptx dim {dim}")
    parts = []; over = []
    sld_ids = ""; pres_rels = [("rIdM", R + "/slideMaster", "slideMasters/slideMaster1.xml", False)]
    for i, (shapes, notes) in enumerate(slides):
        pn = f"ppt/slides/slide{i+1}.xml"
        parts.append((pn, f'{XML}<p:sld xmlns:a="{A}" xmlns:r="{R}" xmlns:p="{P}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{shapes}</p:spTree></p:cSld></p:sld>'))
        srels = [("rIdL", R + "/slideLayout", "../slideLayouts/slideLayout1.xml", False)]
        if notes is not None:
            nn = f"ppt/notesSlides/notesSlide{i+1}.xml"
            parts.append((nn, f'{XML}<p:notes xmlns:a="{A}" xmlns:r="{R}" xmlns:p="{P}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>{notes}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>'))
            parts.append((f"ppt/notesSlides/_rels/notesSlide{i+1}.xml.rels", rels([("rIdS", R + "/slide", f"../slides/slide{i+1}.xml", False)])))
            over.append((f"/{nn}", "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"))
            srels.append(("rIdN", R + "/notesSlide", f"../notesSlides/notesSlide{i+1}.xml", False))
        parts.append((f"ppt/slides/_rels/slide{i+1}.xml.rels", rels(srels)))
        over.append((f"/{pn}", "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"))
        sld_ids += f'<p:sldId id="{256+i}" r:id="rIdS{i}"/>'
        pres_rels.append((f"rIdS{i}", R + "/slide", f"slides/slide{i+1}.xml", False))
    empty_tree = '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>'
    parts += [
        ("[Content_Types].xml", ctypes([("rels", "application/vnd.openxmlformats-package.relationships+xml"), ("xml", "application/xml")],
            [("/ppt/presentation.xml", "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"),
             ("/ppt/slideMasters/slideMaster1.xml", "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"),
             ("/ppt/slideLayouts/slideLayout1.xml", "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml")] + over)),
        ("_rels/.rels", rels([("rId1", R + "/officeDocument", "ppt/presentation.xml", False)])),
        ("ppt/presentation.xml", f'{XML}<p:presentation xmlns:a="{A}" xmlns:r="{R}" xmlns:p="{P}"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdM"/></p:sldMasterIdLst><p:sldIdLst>{sld_ids}</p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>'),
        ("ppt/_rels/presentation.xml.rels", rels(pres_rels)),
        ("ppt/slideMasters/slideMaster1.xml", f'{XML}<p:sldMaster xmlns:a="{A}" xmlns:r="{R}" xmlns:p="{P}">{empty_tree}<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rIdL"/></p:sldLayoutIdLst></p:sldMaster>'),
        ("ppt/slideMasters/_rels/slideMaster1.xml.rels", rels([("rIdL", R + "/slideLayout", "../slideLayouts/slideLayout1.xml", False)])),
        ("ppt/slideLayouts/slideLayout1.xml", f'{XML}<p:sldLayout xmlns:a="{A}" xmlns:r="{R}" xmlns:p="{P}" type="blank">{empty_tree}</p:sldLayout>'),
        ("ppt/slideLayouts/_rels/slideLayout1.xml.rels", rels([("rIdM", R + "/slideMaster", "../slideMasters/slideMaster1.xml", False)])),
    ]
    return parts

DIMS = {
    "docx": ["para", "runs", "table", "tables", "hyperlinks", "bookmarks", "styles", "comments", "footnotes", "images", "lists", "sections"],
    "xlsx": ["rows", "sst", "inline", "cols", "merges", "styles", "dates", "comments", "hyperlinks", "sheets", "formulas", "dv", "cf", "sparse"],
    "pptx": ["slides", "shapes", "runs", "table", "notes"],
}
GEN = {"docx": docx, "xlsx": xlsx, "pptx": pptx}

def generate(outdir, fmt, dim, n):
    os.makedirs(outdir, exist_ok=True)
    path = os.path.join(outdir, f"{fmt}_{dim}_{n}.{fmt}")
    write(path, GEN[fmt](dim, n))
    return path

if __name__ == "__main__":
    if sys.argv[1:] == ["--list"]:
        for f, ds in DIMS.items():
            for d in ds: print(f, d)
        sys.exit(0)
    outdir, fmt_dim, n = sys.argv[1], sys.argv[2], int(sys.argv[3])
    fmt, dim = fmt_dim.split(".")
    print(generate(outdir, fmt, dim, n))
