// 05_xlsx_writer: write a workbook with typed cells, a formula, a style, a merge and a column width, then read it back.
//
// Run: dotnet run --project examples/csharp/05_xlsx_writer/05_xlsx_writer.csproj
using System.Text.Json;
using OfficeOxide;

static void Check(bool condition, string message)
{
    if (condition) return;
    Console.Error.WriteLine($"FAIL: {message}");
    Environment.Exit(1);
}

static void Has(string text, string expected) =>
    Check(text.Contains(expected), $"\"{expected}\" is missing from the text");

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

// Counts the sheets or slides in the IR tree.
static int CountSections(Document doc)
{
    using var ir = JsonDocument.Parse(doc.ToIrJson());
    return ir.RootElement.GetProperty("sections").GetArrayLength();
}

Console.WriteLine("step 1: build an XLSX in memory");
var data = BuildXlsxBytes();

Console.WriteLine("step 2: open it from bytes");
using var doc = Document.FromBytes(data, "xlsx");
var text = doc.PlainText();
var sheets = CountSections(doc);
Console.WriteLine($"  format: {doc.Format}");
Console.WriteLine($"  chars: {text.Length}");
Console.WriteLine($"  sheets: {sheets}");

Console.WriteLine("step 3: check the cells");
Check(sheets == 1, "expected one sheet");
foreach (var expected in new[] { "Sales", "Region", "Q1", "Q2", "Growth", "North", "120", "135", "South", "95", "110", "Total", "Notes" })
    Has(text, expected);
Check(text.Contains("TRUE"), "the boolean cell is missing");
Check(text.Contains("SUM(B2:B3)") && text.Contains("SUM(C2:C3)"), "a formula is missing");

Console.WriteLine("OK 05_xlsx_writer");
