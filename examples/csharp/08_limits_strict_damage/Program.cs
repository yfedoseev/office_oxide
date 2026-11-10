// 08_limits_strict_damage: strict mode refuses a damaged archive; a lenient open reports the damage.
//
// Run: dotnet run --project examples/csharp/08_limits_strict_damage/08_limits_strict_damage.csproj
using System.Buffers.Binary;
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

const string Part = "word/_rels/document.xml.rels";

// Overwrites the start of one part's compressed data. A first deflate byte of
// 0xFF declares a reserved block type, so every inflater rejects the part, and
// the central directory stays intact. The part is found by name, not by
// offset, so the recipe does not depend on how big the other parts are.
static byte[] DamagePart(byte[] data)
{
    var at = data.AsSpan().IndexOf(Encoding.ASCII.GetBytes(Part)); // the first hit is the part's local header
    Check(at > 0, "the part to damage is missing from the archive");
    // The extra-field length sits in the two bytes just before the name.
    var start = at + Part.Length + BinaryPrimitives.ReadUInt16LittleEndian(data.AsSpan(at - 2));
    var damaged = (byte[])data.Clone();
    damaged.AsSpan(start, 16).Fill(0xFF);
    return damaged;
}

// Opens bytes as a DOCX; returns the error text, or null when the open succeeded.
static string? OpenError(byte[] data, ParseOptions options)
{
    try
    {
        using var doc = Document.FromBytes(data, "docx", options);
        return null;
    }
    catch (OfficeOxideException e)
    {
        return e.Message;
    }
}

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

Console.WriteLine("step 1: build a DOCX in memory");
var good = BuildDocxBytes();

Console.WriteLine($"step 2: damage one part: overwrite 16 bytes of {Part}");
var damaged = DamagePart(good);

Console.WriteLine("step 3: open it in strict mode");
using var strict = new ParseOptions(strict: true);
var error = OpenError(damaged, strict);
Check(error is not null, "strict mode accepted a damaged part");
Console.WriteLine($"  error: {error}");

Console.WriteLine("step 4: open it leniently");
using var lenient = new ParseOptions();
using (var doc = Document.FromBytes(damaged, "docx", lenient))
{
    var kinds = doc.Damage().Select(d => d.Kind).ToArray();
    Console.WriteLine($"  damage: {string.Join(", ", kinds)}");
    Check(kinds.Contains("unreadable_part"), "the damage report does not name an unreadable part");
    Has(doc.PlainText(), "Quarterly Report");
}

Console.WriteLine("step 5: truncate the archive to its first half and open it both ways");
var half = good[..(good.Length / 2)];
foreach (var (options, message) in new[] { (strict, "strict mode accepted a truncated archive"), (lenient, "a lenient open accepted a truncated archive") })
{
    error = OpenError(half, options);
    Check(error is not null, message);
    Console.WriteLine($"  error: {error}");
}

Console.WriteLine("step 6: a limit applies to one document only");
using var limited = new ParseOptions().Set("part_bytes", 100UL);
error = OpenError(good, limited);
Check(error is not null, "the part_bytes limit was not applied");
Console.WriteLine($"  error: {error}");
Has(TextOf(good, "docx"), "Quarterly Report"); // the default limits still accept it

Console.WriteLine("OK 08_limits_strict_damage");
