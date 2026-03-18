using DocumentAssembler.Core;
using DocumentFormat.OpenXml.Packaging;
using System.Xml.Linq;

namespace Example17_ConditionalIfElse_SDT
{
    /// <summary>
    /// Example 17: Conditional If/Else with SDT-based templates
    ///
    /// This example tests Conditional/Else/EndConditional tags created via
    /// SDT content controls (as produced by LibreOffice or the web editor),
    /// rather than the <#...#> text format.
    ///
    /// Template (docs/test1.docx) structure:
    ///   <Conditional Select="status" Match="OK" />
    ///   A
    ///   <Else />
    ///   B
    ///   <EndConditional />
    ///
    /// Expected:
    ///   status=OK  -> output contains "A" only
    ///   status=KO  -> output contains "B" only
    /// </summary>
    class Program
    {
        static int Main(string[] args)
        {
            Console.WriteLine("===========================================");
            Console.WriteLine("Example 17: Conditional If/Else (SDT tags)");
            Console.WriteLine("===========================================\n");

            var hasErrors = false;

            // Test 1: status = OK -> should show A only
            Console.WriteLine("Test 1: status = OK");
            Console.WriteLine("------------------------------------------");
            var result1 = GenerateAndValidate(
                "TemplateDocument.docx",
                new XElement("Data", new XElement("status", "OK")),
                "Output_StatusOK.docx",
                expectedContains: "A",
                expectedNotContains: "B"
            );
            if (!result1) hasErrors = true;

            // Test 2: status = KO -> should show B only
            Console.WriteLine("\nTest 2: status = KO");
            Console.WriteLine("------------------------------------------");
            var result2 = GenerateAndValidate(
                "TemplateDocument.docx",
                new XElement("Data", new XElement("status", "KO")),
                "Output_StatusKO.docx",
                expectedContains: "B",
                expectedNotContains: "A"
            );
            if (!result2) hasErrors = true;

            Console.WriteLine("\n===========================================");
            if (hasErrors)
            {
                Console.WriteLine("SOME TESTS FAILED!");
                return 1;
            }

            Console.WriteLine("All tests passed!");
            Console.WriteLine("===========================================");
            return 0;
        }

        static bool GenerateAndValidate(
            string templateName,
            XElement data,
            string outputName,
            string expectedContains,
            string expectedNotContains)
        {
            try
            {
                var templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, templateName);
                var wmlTemplate = new WmlDocument(templatePath);

                var wmlAssembled = DocumentAssembler.Core.DocumentAssembler.AssembleDocument(
                    wmlTemplate,
                    data,
                    out bool templateError
                );

                var outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, outputName);
                wmlAssembled.SaveAs(outputPath);

                if (templateError)
                {
                    Console.WriteLine($"  FAIL: Template errors detected");
                    Console.WriteLine($"  Output: {outputPath}");
                    return false;
                }

                // Read the generated document and validate content
                using var doc = WordprocessingDocument.Open(outputPath, false);
                var body = doc.MainDocumentPart?.Document.Body;
                if (body == null)
                {
                    Console.WriteLine($"  FAIL: Document body is null");
                    return false;
                }

                var text = body.InnerText.Trim();
                Console.WriteLine($"  Data: {data}");
                Console.WriteLine($"  Output text: \"{text}\"");

                var containsExpected = text.Contains(expectedContains);
                var containsUnexpected = text.Contains(expectedNotContains);

                if (containsExpected && !containsUnexpected)
                {
                    Console.WriteLine($"  PASS: Contains \"{expectedContains}\", does not contain \"{expectedNotContains}\"");
                    return true;
                }

                if (!containsExpected)
                {
                    Console.WriteLine($"  FAIL: Expected \"{expectedContains}\" not found in output");
                }
                if (containsUnexpected)
                {
                    Console.WriteLine($"  FAIL: Unexpected \"{expectedNotContains}\" found in output");
                }
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  ERROR: {ex.Message}");
                return false;
            }
        }
    }
}
