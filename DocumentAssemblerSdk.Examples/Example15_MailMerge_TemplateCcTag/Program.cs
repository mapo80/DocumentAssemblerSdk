using DocumentAssembler.Core;

var templatePath = Path.Combine(AppContext.BaseDirectory, "Template.docx");
if (!File.Exists(templatePath))
{
    Console.Error.WriteLine($"Template not found at {templatePath}");
    return 1;
}

var wml = new WmlDocument(templatePath);
var fields = TemplateSchemaExtractor.ExtractMailMergeFields(wml);

if (fields.Count != 0)
{
    Console.Error.WriteLine("No MailMerge fields were expected in this template.");
    foreach (var field in fields)
    {
        Console.Error.WriteLine($" - {field.FieldName} => {field.XPath}");
    }
    return 2;
}

Console.WriteLine("The template does not contain MailMerge fields, as expected.");
return 0;
