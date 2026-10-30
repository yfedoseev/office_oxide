// 06_pptx_writer: write a presentation with a size, slides, titles, text and an image, then read it back.
//
// Run: dotnet run --project examples/csharp/06_pptx_writer/06_pptx_writer.csproj
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

// Counts the sheets or slides in the IR tree.
static int CountSections(Document doc)
{
    using var ir = JsonDocument.Parse(doc.ToIrJson());
    return ir.RootElement.GetProperty("sections").GetArrayLength();
}

Console.WriteLine("step 1: build a PPTX in memory");
var data = BuildPptxBytes();

Console.WriteLine("step 2: open it from bytes");
using var doc = Document.FromBytes(data, "pptx");
var text = doc.PlainText();
var slides = CountSections(doc);
// Plain Markdown drops images; the WithImages variant keeps them in place.
var withImages = doc.ToMarkdownWithImages();
Console.WriteLine($"  format: {doc.Format}");
Console.WriteLine($"  chars: {text.Length}");
Console.WriteLine($"  slides: {slides}");

Console.WriteLine("step 3: check the slides");
Check(slides == 2, "expected two slides");
foreach (var expected in new[] { "Welcome", "Quarterly Report", "Results", "Revenue grew in every region." })
    Has(text, expected);
Check(text.IndexOf("Welcome", StringComparison.Ordinal) < text.IndexOf("Results", StringComparison.Ordinal), "the slides are out of order");
Check(withImages.Contains("[image-base64:"), "the slide image is missing");

Console.WriteLine("OK 06_pptx_writer");
