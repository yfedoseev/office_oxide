// 05_xlsx_writer — write a workbook with typed cells, a formula, a styled cell, merged cells and a column width; read it back.
//
// Run: dotnet run --project examples/csharp/05_xlsx_writer/05_xlsx_writer.csproj
//      (set LD_LIBRARY_PATH to the directory containing liboffice_oxide.so)
using System.Text.Json;
using OfficeOxide;

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

static int CountSections(Document doc)
{
    using var ir = JsonDocument.Parse(doc.ToIrJson());
    return ir.RootElement.GetProperty("sections").GetArrayLength();
}

static void Check(bool condition, string message)
{
    if (!condition) throw new Exception($"check failed: {message}");
}

Console.WriteLine("step 1: build a workbook");
var bytes = BuildXlsxBytes();
Console.WriteLine($"  chars: {bytes.Length}");

Console.WriteLine("step 2: read it back from bytes");
using var doc = Document.FromBytes(bytes, "xlsx");
var text = doc.PlainText();
Console.WriteLine($"  format: {doc.Format}");
Console.WriteLine($"  sheets: {CountSections(doc)}");

Console.WriteLine("step 3: check the cells");
Check(CountSections(doc) == 1, "the workbook has one sheet");
foreach (var want in new[] { "Region", "Q1", "Q2", "Growth", "North", "120", "135", "South", "95", "110", "Total", "Notes" })
    Check(text.Contains(want), $"text contains {want}");
Check(text.Contains("TRUE"), "the boolean cell reads back as TRUE");
Check(doc.ToIrJson().Contains("SUM(B2:B3)"), "the formula survives the round trip");

Console.WriteLine("step 4: check that an off-grid write is rejected");
using (var x = new XlsxWriter())
{
    var s = x.AddSheet("Sales");
    var threw = false;
    try { x.SetCell(s, 2_000_000, 0, "off-grid"); } catch (InvalidOperationException) { threw = true; }
    Check(threw, "an off-grid write throws");
}

Console.WriteLine("OK 05_xlsx_writer");
