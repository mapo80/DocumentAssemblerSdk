using DocumentAssembler.Core;
using System.Xml.Linq;

namespace Example09_AllTags;

/// <summary>
/// Example 09: end-to-end test for TemplateAllTagsDocument.
/// This example feeds a sample XML covering every tag supported by the SDK
/// and produces an assembled DOCX for manual inspection.
/// </summary>
internal class Program
{
    static void Main()
    {
        Console.WriteLine("===========================================");
        Console.WriteLine("Example 09: All Tags Template");
        Console.WriteLine("===========================================\n");

        var data = Example09DataFactory.CreateSampleData();
        GenerateDocument("TemplateAllTagsDocument.docx", data, "Output_AllTagsDemo.docx");

        Console.WriteLine("\nDone! Check Output_AllTagsDemo.docx in the output folder.");
    }

    private static void GenerateDocument(string templateName, XElement data, string outputName)
    {
        try
        {
            var templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, templateName);
            var wmlTemplate = new WmlDocument(templatePath);

            var assembled = DocumentAssembler.Core.DocumentAssembler.AssembleDocument(
                wmlTemplate,
                data,
                out bool templateError,
                out string? templateSummary);

            var outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, outputName);
            assembled.SaveAs(outputPath);

            if (templateError)
            {
                Console.WriteLine($"⚠️  Template warnings: {templateSummary}");
            }

            Console.WriteLine($"✓ Generated {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Error generating document: {ex.Message}");
        }
    }

}
