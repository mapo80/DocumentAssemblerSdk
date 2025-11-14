using DocumentAssembler.Core;

var templatePath = Path.Combine(AppContext.BaseDirectory, "Template.docx");
if (!File.Exists(templatePath))
{
    Console.Error.WriteLine($"Template not found at {templatePath}");
    return 1;
}

var wml = new WmlDocument(templatePath);
var fields = TemplateSchemaExtractor.ExtractMailMergeFields(wml);
var actual = fields.Select(f => f.FieldName).OrderBy(x => x, StringComparer.Ordinal).ToArray();
var expected = new[] { "FirstName", "LastName" };

if (!actual.SequenceEqual(expected))
{
    Console.Error.WriteLine("Mail merge fields mismatch.");
    Console.Error.WriteLine("Expected : " + string.Join(", ", expected));
    Console.Error.WriteLine("Actual   : " + (actual.Length == 0 ? "(none)" : string.Join(", ", actual)));
    return 2;
}

Console.WriteLine("Detected mail merge fields:");
foreach (var field in fields)
{
    Console.WriteLine($" - {field.FieldName} => {field.XPath}");
}
Console.WriteLine("All expected fields are present.");
return 0;
