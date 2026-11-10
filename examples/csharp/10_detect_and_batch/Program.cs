// 10_detect_and_batch: detect the format of bytes, then process a mixed batch.
//
// Run: dotnet run --project examples/csharp/10_detect_and_batch/10_detect_and_batch.csproj
using System.Text;
using OfficeOxide;

const string Markdown = """
    # Quarterly Report

    Revenue grew in every region.

    | Region | Q1  | Q2  |
    |--------|-----|-----|
    | North  | 120 | 135 |
    | South  | 95  | 110 |

    """;

static void Check(bool condition, string message)
{
    if (condition) return;
    Console.Error.WriteLine($"FAIL: {message}");
    Environment.Exit(1);
}

// CreateFromMarkdown writes to a path, so the file lives in a private temp
// file just long enough to read its bytes back.
static byte[] MarkdownToBytes(string markdown, string format)
{
    var path = Path.Combine(Path.GetTempPath(), $"oo_example_{Guid.NewGuid():N}.{format}");
    try
    {
        Document.CreateFromMarkdown(markdown, format, path);
        return File.ReadAllBytes(path);
    }
    finally
    {
        File.Delete(path);
    }
}

static byte[] BuildDocxBytes(string title = "Quarterly Report") =>
    MarkdownToBytes(Markdown.Replace("Quarterly Report", title), "docx");

// Opens bytes as the given format and returns the plain text.
static string TextOf(byte[] data, string format)
{
    using var doc = Document.FromBytes(data, format);
    return doc.PlainText();
}

static byte[] BuildXlsxBytes()
{
    using var x = new XlsxWriter();
    var s = x.AddSheet("Sales");
    string[] headers = ["Region", "Q1", "Q2", "Growth"];
    for (uint c = 0; c < headers.Length; c++)
        x.SetCellStyled(s, 0, c, headers[c], bold: true); // bold header row
    x.SetCell(s, 1, 0, "North"); x.SetCell(s, 1, 1, 120); x.SetCell(s, 1, 2, 135); x.SetCell(s, 1, 3, true); // D2
    x.SetCell(s, 2, 0, "South"); x.SetCell(s, 2, 1, 95); x.SetCell(s, 2, 2, 110);
    x.SetCell(s, 3, 0, "Total");
    x.SetFormula(s, 3, 1, "SUM(B2:B3)");
    x.SetFormula(s, 3, 2, "SUM(C2:C3)");
    x.SetCell(s, 5, 0, "Notes");
    x.MergeCells(s, 5, 0, 1, 3); // A6:C6
    x.SetColumnWidth(s, 0, 18.0);
    return x.ToBytes();
}

static byte[] BuildPptxBytes()
{
    var png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==");
    using var p = new PptxWriter();
    p.SetPresentationSize(12192000, 6858000); // 16:9, in EMU (914400 per inch)
    var one = p.AddSlide();
    p.SetSlideTitle(one, "Welcome");
    p.AddSlideText(one, "Quarterly Report");
    var two = p.AddSlide();
    p.SetSlideTitle(two, "Results");
    p.AddSlideText(two, "Revenue grew in every region.");
    p.AddSlideImage(two, png, "png", 914400, 2743200, 1828800, 1828800);
    return p.ToBytes();
}

// The binding detects from a path only. A ZIP stores each part name in the
// clear, and the main part's name tells the three formats apart.
static string? DetectFormat(byte[] data)
{
    if (data.Length < 2 || data[0] != (byte)'P' || data[1] != (byte)'K') return null;
    (string Part, string Format)[] parts = [("word/document.xml", "docx"), ("xl/workbook.xml", "xlsx"), ("ppt/presentation.xml", "pptx")];
    foreach (var (part, format) in parts)
        if (data.AsSpan().IndexOf(Encoding.ASCII.GetBytes(part)) >= 0) return format;
    return null;
}

Console.WriteLine("step 1: build a DOCX, an XLSX and a PPTX in memory");
var batch = new (string Name, byte[] Data)[]
{
    ("report.docx", BuildDocxBytes()),
    ("report.xlsx", BuildXlsxBytes()),
    ("report.pptx", BuildPptxBytes()),
};

Console.WriteLine("step 2: detect each format and open it from bytes");
foreach (var (name, data) in batch)
{
    var format = DetectFormat(data);
    Check(format == Path.GetExtension(name)[1..], "the detected format differs from the one built");
    var text = TextOf(data, format!);
    Check(text.Trim() != "", "a document has no text");
    Console.WriteLine($"  {name}: format {format}, chars {text.Length}");
}

Console.WriteLine("step 3: check that bytes that are not a document are not detected");
Check(DetectFormat("not an office file"u8.ToArray()) is null, "bytes that are not a document were detected as one");

Console.WriteLine("OK 10_detect_and_batch");
