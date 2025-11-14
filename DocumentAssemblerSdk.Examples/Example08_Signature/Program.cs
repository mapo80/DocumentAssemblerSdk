using DocumentAssembler.Core;
using System.Xml.Linq;

namespace Example08_Signature;

/// <summary>
/// Example 08: Signature placeholders.
/// Demonstrates how to place <Signature> tags inside a template
/// and assemble the document so that downstream PDF conversion
/// can create signature fields automatically.
/// </summary>
internal class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("===========================================");
        Console.WriteLine("Example 08: Signature placeholders");
        Console.WriteLine("===========================================\n");

        GenerateDocument("TemplateSignatureDocument.docx", new XElement("Modulo"), "Output_SignatureDemo.docx");

        Console.WriteLine("\nDone! The assembled document is in the output directory.");
        Console.WriteLine("You can now run the usual DOCX -> PDF conversion pipeline to");
        Console.WriteLine("see the AcroForm signature fields generated.");
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
