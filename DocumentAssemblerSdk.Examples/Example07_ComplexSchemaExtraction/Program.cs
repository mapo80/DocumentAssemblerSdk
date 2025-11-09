using DocumentAssembler.Core;
using System;
using System.IO;
using System.Linq;

namespace Example07_ComplexSchemaExtraction
{
    /// <summary>
    /// Example 07: Complex Schema Extraction
    ///
    /// This example demonstrates schema extraction from a comprehensive template
    /// that includes all DocumentAssembler tag types:
    /// - Content (simple fields)
    /// - Image (with sizing metadata)
    /// - Table (dynamic rows)
    /// - Repeat (repeating blocks)
    /// - Conditional with Else (if-else logic)
    /// - Nested structures (conditionals within conditionals)
    /// - Optional fields
    ///
    /// Template represents an E-Commerce Order Report with realistic complexity.
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("=== Example 07: Complex Schema Extraction ===\n");

            try
            {
                string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ComplexTemplate.docx");

                // Check if template exists, if not create it
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine("Template not found. Creating complex template...");
                    CreateComplexTemplate(templatePath);
                    Console.WriteLine();
                }

                // Load the complex template
                Console.WriteLine($"Loading template: {Path.GetFileName(templatePath)}");
                Console.WriteLine($"Template size: {new FileInfo(templatePath).Length:N0} bytes\n");

                var templateDoc = new WmlDocument(templatePath);

                // Extract the schema
                Console.WriteLine("Extracting XML schema...");
                var sw = System.Diagnostics.Stopwatch.StartNew();
                var result = TemplateSchemaExtractor.ExtractXmlSchema(templateDoc);
                sw.Stop();

                Console.WriteLine($"✓ Extraction completed in {sw.ElapsedMilliseconds}ms\n");

                // Display summary statistics
                Console.WriteLine("═══════════════════════════════════════");
                Console.WriteLine("           EXTRACTION SUMMARY           ");
                Console.WriteLine("═══════════════════════════════════════");
                Console.WriteLine($"Total fields discovered: {result.Fields.Count}");
                Console.WriteLine($"Root element name: {result.RootElementName}");
                Console.WriteLine();

                // Group and display fields by type
                var fieldsByType = result.Fields.GroupBy(f => f.TagType).OrderBy(g => g.Key);

                Console.WriteLine("Fields by Type:");
                Console.WriteLine("───────────────────────────────────────");
                foreach (var group in fieldsByType)
                {
                    Console.WriteLine($"  {group.Key,-15} : {group.Count(),3} field(s)");
                }
                Console.WriteLine();

                // Display optional vs required
                var optionalCount = result.Fields.Count(f => f.IsOptional);
                var requiredCount = result.Fields.Count(f => !f.IsOptional);
                Console.WriteLine($"Optional fields: {optionalCount}");
                Console.WriteLine($"Required fields: {requiredCount}");
                Console.WriteLine();

                // Display repeating structures
                var repeatingCount = result.Fields.Count(f => f.IsRepeating);
                Console.WriteLine($"Repeating structures: {repeatingCount}");
                Console.WriteLine();

                // Show detailed field analysis
                Console.WriteLine("═══════════════════════════════════════");
                Console.WriteLine("         DETAILED FIELD ANALYSIS        ");
                Console.WriteLine("═══════════════════════════════════════\n");

                foreach (var group in fieldsByType)
                {
                    Console.WriteLine($"╔═══ {group.Key} Fields ({group.Count()}) ═══");

                    foreach (var field in group.OrderBy(f => f.XPath))
                    {
                        Console.WriteLine($"║ XPath: {field.XPath}");

                        var properties = new System.Collections.Generic.List<string>();
                        if (!field.IsOptional) properties.Add("REQUIRED");
                        if (field.IsRepeating) properties.Add("REPEATING");

                        if (properties.Any())
                        {
                            Console.WriteLine($"║   Properties: {string.Join(", ", properties)}");
                        }

                        if (field.Attributes.Count > 1) // More than just Select
                        {
                            Console.WriteLine($"║   Attributes:");
                            foreach (var attr in field.Attributes.Where(a => a.Key != "Select"))
                            {
                                Console.WriteLine($"║     • {attr.Key} = \"{attr.Value}\"");
                            }
                        }
                        Console.WriteLine("║");
                    }
                    Console.WriteLine("╚═══════════════════════════════════════\n");
                }

                // Generate and save XML template
                Console.WriteLine("═══════════════════════════════════════");
                Console.WriteLine("        GENERATED XML TEMPLATE          ");
                Console.WriteLine("═══════════════════════════════════════\n");

                string xmlTemplate = result.ToFormattedXml();
                Console.WriteLine(xmlTemplate);
                Console.WriteLine();

                // Save to file
                string outputPath = Path.Combine(
                    AppDomain.CurrentDomain.BaseDirectory,
                    "Schema_ComplexTemplate.xml"
                );
                File.WriteAllText(outputPath, xmlTemplate);
                Console.WriteLine($"✓ Schema saved to: {Path.GetFileName(outputPath)}");
                Console.WriteLine();

                // Validation summary
                Console.WriteLine("═══════════════════════════════════════");
                Console.WriteLine("           VALIDATION SUMMARY           ");
                Console.WriteLine("═══════════════════════════════════════");

                // Check for expected tag types
                bool hasContent = result.Fields.Any(f => f.TagType == "Content");
                bool hasImage = result.Fields.Any(f => f.TagType == "Image");
                bool hasTable = result.Fields.Any(f => f.TagType == "Table");
                bool hasRepeat = result.Fields.Any(f => f.TagType == "Repeat");
                bool hasConditional = result.Fields.Any(f => f.TagType == "Conditional");

                Console.WriteLine($"✓ Content tags:      {(hasContent ? "YES" : "NO")}");
                Console.WriteLine($"✓ Image tags:        {(hasImage ? "YES" : "NO")}");
                Console.WriteLine($"✓ Table tags:        {(hasTable ? "YES" : "NO")}");
                Console.WriteLine($"✓ Repeat tags:       {(hasRepeat ? "YES" : "NO")}");
                Console.WriteLine($"✓ Conditional tags:  {(hasConditional ? "YES" : "NO")}");
                Console.WriteLine();

                // Check XML validity
                try
                {
                    var xmlDoc = new System.Xml.XmlDocument();
                    xmlDoc.LoadXml(xmlTemplate);
                    Console.WriteLine("✓ Generated XML is well-formed");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"✗ XML validation failed: {ex.Message}");
                }

                Console.WriteLine();
                Console.WriteLine("═══════════════════════════════════════");
                Console.WriteLine("    All validations completed! ✓");
                Console.WriteLine("═══════════════════════════════════════");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\n✗ Error: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
                Environment.Exit(1);
            }
        }

        /// <summary>
        /// Creates the complex template using Python script
        /// </summary>
        static void CreateComplexTemplate(string outputPath)
        {
            var pythonScript = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "create_complex_template.py");

            if (!File.Exists(pythonScript))
            {
                throw new FileNotFoundException("Python script not found", pythonScript);
            }

            var process = new System.Diagnostics.Process
            {
                StartInfo = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = "python3",
                    Arguments = $"\"{pythonScript}\" \"{outputPath}\"",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                }
            };

            process.Start();
            string output = process.StandardOutput.ReadToEnd();
            string error = process.StandardError.ReadToEnd();
            process.WaitForExit();

            if (process.ExitCode != 0)
            {
                throw new InvalidOperationException($"Failed to create template: {error}");
            }

            Console.WriteLine(output);
        }
    }
}
