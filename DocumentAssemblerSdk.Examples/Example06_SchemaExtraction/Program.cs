using DocumentAssembler.Core;
using System;
using System.IO;
using System.Linq;

namespace Example06_SchemaExtraction
{
    /// <summary>
    /// Example 06: Schema Extraction
    ///
    /// This example demonstrates how to extract the required XML schema
    /// from a DocumentAssembler template. This is useful when you have
    /// a template and need to know what XML data structure is required.
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("=== Example 06: XML Schema Extraction ===\n");

            try
            {
                // Example 1: Extract schema from a basic template
                Console.WriteLine("Example 1: Basic Template Schema Extraction");
                Console.WriteLine("--------------------------------------------");
                ExtractSchemaFromTemplate("TemplateDocument.docx");

                Console.WriteLine("\n\n");

                // Example 2: Extract schema and analyze fields
                Console.WriteLine("Example 2: Detailed Field Analysis");
                Console.WriteLine("-----------------------------------");
                AnalyzeTemplateFields("TemplateDocument.docx");

                Console.WriteLine("\n\nAll examples completed successfully!");
                Console.WriteLine("\nCheck the output directory for generated XML schema files.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\nError: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }

        /// <summary>
        /// Extracts and displays the XML schema from a template
        /// </summary>
        static void ExtractSchemaFromTemplate(string templateFileName)
        {
            // Load the template
            var templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, templateFileName);
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template not found: {templatePath}");
                Console.WriteLine("Please ensure the template file exists.");
                return;
            }

            var templateDoc = new WmlDocument(templatePath);

            // Extract the schema
            Console.WriteLine($"Analyzing template: {templateFileName}");
            var result = TemplateSchemaExtractor.ExtractXmlSchema(templateDoc);

            Console.WriteLine($"\nDiscovered {result.Fields.Count} field(s)");
            Console.WriteLine($"Root element: {result.RootElementName}");

            // Display the generated XML template
            Console.WriteLine("\nGenerated XML Template:");
            Console.WriteLine("----------------------");
            Console.WriteLine(result.XmlTemplate);

            // Save to file
            var outputPath = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                $"Schema_{Path.GetFileNameWithoutExtension(templateFileName)}.xml"
            );
            File.WriteAllText(outputPath, result.ToFormattedXml());
            Console.WriteLine($"\nSchema saved to: {Path.GetFileName(outputPath)}");

            var formattedXsd = result.ToFormattedXsd();
            var xsdOutputPath = Path.ChangeExtension(outputPath, ".xsd");
            File.WriteAllText(xsdOutputPath, formattedXsd);
            Console.WriteLine($"\nXSD saved to: {Path.GetFileName(xsdOutputPath)}");

            Console.WriteLine("\nGenerated XSD Schema:");
            Console.WriteLine("---------------------");
            Console.WriteLine(formattedXsd);
        }

        /// <summary>
        /// Analyzes and displays detailed information about template fields
        /// </summary>
        static void AnalyzeTemplateFields(string templateFileName)
        {
            var templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, templateFileName);
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template not found: {templatePath}");
                return;
            }

            var templateDoc = new WmlDocument(templatePath);
            var result = TemplateSchemaExtractor.ExtractXmlSchema(templateDoc);

            Console.WriteLine($"Analyzing template: {templateFileName}\n");

            // Group fields by type
            var fieldsByType = result.Fields.GroupBy(f => f.TagType);

            foreach (var group in fieldsByType.OrderBy(g => g.Key))
            {
                Console.WriteLine($"\n{group.Key} Fields ({group.Count()}):");
                Console.WriteLine(new string('-', 40));

                foreach (var field in group.OrderBy(f => f.XPath))
                {
                    Console.WriteLine($"  XPath: {field.XPath}");
                    Console.WriteLine($"    - Element: {field.ElementName}");
                    Console.WriteLine($"    - Optional: {field.IsOptional}");
                    Console.WriteLine($"    - Repeating: {field.IsRepeating}");

                    if (field.Attributes.Count > 0)
                    {
                        Console.WriteLine($"    - Attributes:");
                        foreach (var attr in field.Attributes)
                        {
                            if (attr.Key != "Select")
                            {
                                Console.WriteLine($"        {attr.Key} = {attr.Value}");
                            }
                        }
                    }
                    Console.WriteLine();
                }
            }

            // Summary
            Console.WriteLine("\nSummary:");
            Console.WriteLine("--------");
            Console.WriteLine($"Total fields: {result.Fields.Count}");
            Console.WriteLine($"Optional fields: {result.Fields.Count(f => f.IsOptional)}");
            Console.WriteLine($"Required fields: {result.Fields.Count(f => !f.IsOptional)}");
            Console.WriteLine($"Repeating structures: {result.Fields.Count(f => f.IsRepeating)}");
            Console.WriteLine($"Content fields: {result.Fields.Count(f => f.TagType == "Content")}");
            Console.WriteLine($"Image fields: {result.Fields.Count(f => f.TagType == "Image")}");
            Console.WriteLine($"Table fields: {result.Fields.Count(f => f.TagType == "Table")}");
            Console.WriteLine($"Repeat fields: {result.Fields.Count(f => f.TagType == "Repeat")}");
            Console.WriteLine($"Conditional fields: {result.Fields.Count(f => f.TagType == "Conditional")}");
        }
    }
}
