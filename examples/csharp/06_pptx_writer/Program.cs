// 06_pptx_writer — write a deck with a custom size, titles, text and an image; read it back.
//
// Run: dotnet run --project examples/csharp/06_pptx_writer/06_pptx_writer.csproj
//      (set LD_LIBRARY_PATH to the directory containing liboffice_oxide.so)
using System.Text.Json;
using OfficeOxide;

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

static int CountSections(Document doc)
{
    using var ir = JsonDocument.Parse(doc.ToIrJson());
    return ir.RootElement.GetProperty("sections").GetArrayLength();
}

static void Check(bool condition, string message)
{
    if (!condition) throw new Exception($"check failed: {message}");
}

Console.WriteLine("step 1: build a deck");
var bytes = BuildPptxBytes();
Console.WriteLine($"  chars: {bytes.Length}");

Console.WriteLine("step 2: read it back from bytes");
using var doc = Document.FromBytes(bytes, "pptx");
var text = doc.PlainText();
Console.WriteLine($"  format: {doc.Format}");
Console.WriteLine($"  slides: {CountSections(doc)}");

Console.WriteLine("step 3: check the slides");
Check(CountSections(doc) == 2, "the deck has two slides");
foreach (var want in new[] { "Welcome", "Quarterly Report", "Results", "Revenue grew in every region." })
    Check(text.Contains(want), $"text contains {want}");
Check(text.IndexOf("Welcome", StringComparison.Ordinal) < text.IndexOf("Results", StringComparison.Ordinal),
    "the slides are in order");
Check(doc.ToIrJson().Contains("image", StringComparison.OrdinalIgnoreCase), "the image is in the IR");

Console.WriteLine("OK 06_pptx_writer");
