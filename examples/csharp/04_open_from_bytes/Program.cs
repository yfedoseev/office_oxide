// 04_open_from_bytes: build a DOCX, an XLSX and a PPTX in memory and open each from its bytes.
//
// Run: dotnet run --project examples/csharp/04_open_from_bytes/04_open_from_bytes.csproj
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

static void Has(string text, string expected) =>
    Check(text.Contains(expected), $"\"{expected}\" is missing from the text");

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

Console.WriteLine("step 1: build a DOCX, an XLSX and a PPTX in memory");
var built = new Dictionary<string, byte[]> { ["docx"] = BuildDocxBytes(), ["xlsx"] = BuildXlsxBytes(), ["pptx"] = BuildPptxBytes() };

Console.WriteLine("step 2: open each from bytes");
var texts = new Dictionary<string, string>();
foreach (var (format, data) in built)
{
    using var doc = Document.FromBytes(data, format);
    Check(doc.Format == format, "the opened format differs from the requested one");
    texts[format] = doc.PlainText();
    Console.WriteLine($"  format: {format}");
    Console.WriteLine($"  chars: {texts[format].Length}");
}

Console.WriteLine("step 3: check the text");
Has(texts["docx"], "Revenue grew in every region.");
Has(texts["xlsx"], "North");
Has(texts["xlsx"], "Total");
Has(texts["pptx"], "Welcome");
Has(texts["pptx"], "Results");

Console.WriteLine("step 4: check that a wrong format is an error");
string? error = null;
try
{
    using var wrong = Document.FromBytes(built["xlsx"], "pptx");
}
catch (OfficeOxideException e)
{
    error = e.Message;
}
Check(error is not null, "an XLSX opened as a PPTX was accepted");
Console.WriteLine($"  error: {error}");

Console.WriteLine("OK 04_open_from_bytes");
