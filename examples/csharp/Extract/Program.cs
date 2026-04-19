using OfficeOxide;

if (args.Length != 1)
{
    Console.Error.WriteLine("usage: Extract <file>");
    Environment.Exit(1);
}

using var doc = Document.Open(args[0]);
Console.WriteLine($"format: {doc.Format}");
Console.WriteLine("--- plain text ---");
Console.WriteLine(doc.PlainText());
Console.WriteLine("--- markdown ---");
Console.WriteLine(doc.ToMarkdown());
