using OfficeOxide;

if (args.Length != 2)
{
    Console.Error.WriteLine("usage: Replace <template> <output>");
    Environment.Exit(1);
}

using var ed = EditableDocument.Open(args[0]);
var n = ed.ReplaceText("{{NAME}}", "Alice");
n += ed.ReplaceText("{{DATE}}", "2026-04-18");
Console.WriteLine($"replacements: {n}");
ed.Save(args[1]);
