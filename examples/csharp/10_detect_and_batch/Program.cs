// 10_detect_and_batch — detect each document's format, then process a mixed batch and print one line per document.
//
// Run: dotnet run --project examples/csharp/10_detect_and_batch/10_detect_and_batch.csproj
//      (set LD_LIBRARY_PATH to the directory containing liboffice_oxide.so)
using System.IO.Compression;
using OfficeOxide;

// There is no DOCX writer in the binding, so the DOCX is a minimal OPC package
// zipped in memory: nothing touches the disk.
static byte[] BuildDocxBytes(string title = "Quarterly Report")
{
    static string Para(string t) => $"<w:p><w:r><w:t>{t}</w:t></w:r></w:p>";
    static string Row(params string[] cells) =>
        "<w:tr>" + string.Concat(cells.Select(c => $"<w:tc>{Para(c)}</w:tc>")) + "</w:tr>";
    var body = Para(title) + Para("Revenue grew in every region.")
        + "<w:tbl>" + Row("Region", "Q1", "Q2") + Row("North", "120", "135") + Row("South", "95", "110") + "</w:tbl>";
    var parts = new Dictionary<string, string>
    {
        ["[Content_Types].xml"] =
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
            + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
            + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
            + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>",
        ["_rels/.rels"] =
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>",
        ["word/document.xml"] =
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?><w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
            + $"<w:body>{body}</w:body></w:document>",
    };
    using var ms = new MemoryStream();
    using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        foreach (var (name, xml) in parts)
        {
            using var w = new StreamWriter(zip.CreateEntry(name).Open());
            w.Write(xml);
        }
    return ms.ToArray();
}

// Rows and columns are zero-based; writing outside Excel's grid throws.
static byte[] BuildXlsxBytes()
{
    using var x = new XlsxWriter();
    var s = x.AddSheet("Sales");
    string[] headers = ["Region", "Q1", "Q2", "Growth"];
    for (uint c = 0; c < headers.Length; c++)
        x.SetCellStyled(s, 0, c, headers[c], bold: true, bgColor: "DDDDDD");   // styled header row
    x.SetCell(s, 1, 0, "North"); x.SetCell(s, 1, 1, 120); x.SetCell(s, 1, 2, 135); x.SetCell(s, 1, 3, true);
    x.SetCell(s, 2, 0, "South"); x.SetCell(s, 2, 1, 95);  x.SetCell(s, 2, 2, 110);
    x.SetCell(s, 3, 0, "Total");
    x.SetFormula(s, 3, 1, "SUM(B2:B3)");       // formulas are written without the leading '='
    x.SetFormula(s, 3, 2, "SUM(C2:C3)");
    x.SetCell(s, 5, 0, "Notes");
    x.MergeCells(s, 5, 0, 1, 3);               // A6:C6
    x.SetColumnWidth(s, 0, 18.0);
    return x.ToBytes();
}

// Sizes and offsets are EMU (914400 per inch).
static byte[] BuildPptxBytes()
{
    var png = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==");
    using var p = new PptxWriter();
    p.SetPresentationSize(12192000, 6858000);   // 16:9
    var first = p.AddSlide();
    p.SetSlideTitle(first, "Welcome");
    p.AddSlideText(first, "Quarterly Report");
    var second = p.AddSlide();
    p.SetSlideTitle(second, "Results");
    p.AddSlideText(second, "Revenue grew in every region.");
    p.AddSlideImage(second, png, "png", 914400, 1828800, 914400, 914400);
    return p.ToBytes();
}

static void Check(bool condition, string message)
{
    if (!condition) throw new Exception($"check failed: {message}");
}

Console.WriteLine("step 1: build a DOCX, an XLSX and a PPTX in memory");
var batch = new (string Ext, byte[] Data)[]
{
    ("docx", BuildDocxBytes()), ("xlsx", BuildXlsxBytes()), ("pptx", BuildPptxBytes()),
};

// The binding's detector looks at the file name's extension only; there is no
// sniffing of raw bytes, so each document is staged in a temp directory.
Console.WriteLine("step 2: detect the format of each document");
var dir = Path.Combine(Path.GetTempPath(), $"oo_10_detect_{Environment.ProcessId}");
Directory.CreateDirectory(dir);
try
{
    foreach (var (ext, data) in batch)
    {
        var path = Path.Combine(dir, $"report.{ext}");
        File.WriteAllBytes(path, data);
        var detected = Document.DetectFormat(path);
        Console.WriteLine($"  report.{ext}: {detected}");
        Check(detected == ext, $"report.{ext} is detected as {ext}");
    }

    Console.WriteLine("step 3: process the batch");
    foreach (var (ext, _) in batch)
    {
        using var doc = Document.Open(Path.Combine(dir, $"report.{ext}"));
        var chars = doc.PlainText().Length;
        Console.WriteLine($"  report.{ext}: format={doc.Format} chars={chars}");
        Check(doc.Format == ext, $"report.{ext} opens as {ext}");
        Check(chars > 0, $"report.{ext} has text");
    }
}
finally
{
    Directory.Delete(dir, recursive: true);
}

Console.WriteLine("OK 10_detect_and_batch");
