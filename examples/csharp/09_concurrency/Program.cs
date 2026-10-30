// 09_concurrency: open and read 32 documents from 8 workers at once, with logging on.
//
// Run: dotnet run --project examples/csharp/09_concurrency/09_concurrency.csproj
using OfficeOxide;

const string Markdown = """
    # Quarterly Report

    Revenue grew in every region.

    | Region | Q1  | Q2  |
    |--------|-----|-----|
    | North  | 120 | 135 |
    | South  | 95  | 110 |

    """;

const int Workers = 8, Documents = 32;

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

Console.WriteLine("step 1: build 32 DOCX files in memory");
var inputs = Enumerable.Range(0, Documents).Select(i => BuildDocxBytes($"Quarterly Report {i}")).ToArray();
Console.WriteLine($"  documents: {Documents}");

Console.WriteLine("step 2: switch logging on");
// A plain TextWriter keeps the example free of logging packages; it is
// synchronized because eight workers write to it.
var log = TextWriter.Synchronized(Console.Error);

Console.WriteLine("step 3: parse them on 8 workers");
var texts = new string[Documents];
await Parallel.ForEachAsync(Enumerable.Range(0, Documents), new ParallelOptions { MaxDegreeOfParallelism = Workers },
    (i, _) =>
    {
        log.WriteLine($"document {i}: parse start");
        texts[i] = TextOf(inputs[i], "docx");
        log.WriteLine($"document {i}: parse done");
        return ValueTask.CompletedTask;
    }).WaitAsync(TimeSpan.FromSeconds(60)); // a deadlock fails the example instead of hanging it
Console.WriteLine($"  workers: {Workers}");

Console.WriteLine("step 4: check every result");
Check(texts.All(t => !string.IsNullOrEmpty(t)), "not every document was read");
for (var i = 0; i < Documents; i++)
    Check(texts[i].Contains($"Quarterly Report {i}"), $"document {i} came back with the wrong text");

Console.WriteLine("OK 09_concurrency");
