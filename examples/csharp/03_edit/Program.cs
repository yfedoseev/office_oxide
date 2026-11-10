// 03_edit: replace text in a DOCX and set cells in an XLSX, in memory.
//
// Run: dotnet run --project examples/csharp/03_edit/03_edit.csproj
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

Console.WriteLine("step 1: build a DOCX in memory");
var docx = BuildDocxBytes();

Console.WriteLine("step 2: replace text");
byte[] edited;
using (var ed = EditableDocument.FromBytes(docx, "docx"))
{
    var replaced = ed.ReplaceText("Quarterly Report", "Annual Report");
    Console.WriteLine($"  replacements: {replaced}");
    Check(replaced == 1, "replace_text did not make exactly one replacement");

    Console.WriteLine("step 3: save to bytes and check the text");
    edited = ed.SaveToBytes();
}
var text = TextOf(edited, "docx");
Has(text, "Annual Report");
Check(!text.Contains("Quarterly Report"), "the old title is still present");

Console.WriteLine("step 4: build an XLSX in memory");
var xlsx = BuildXlsxBytes();

Console.WriteLine("step 5: set cells");
using (var ed = EditableDocument.FromBytes(xlsx, "xlsx"))
{
    ed.SetCell(0, "B2", 999.0);
    ed.SetCell(0, "A2", "Central");

    Console.WriteLine("step 6: save to bytes and check the cells");
    edited = ed.SaveToBytes();
}
text = TextOf(edited, "xlsx");
Has(text, "Central");
Has(text, "999");
Check(!text.Contains("North"), "the old A2 value is still present");

Console.WriteLine("OK 03_edit");
