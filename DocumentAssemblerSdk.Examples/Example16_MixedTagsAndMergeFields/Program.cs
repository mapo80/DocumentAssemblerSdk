using DocumentAssembler.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Xml.Linq;

namespace Example16_MixedTagsAndMergeFields
{
    /// <summary>
    /// Example 16: Mixed DA Tags and Merge Fields
    ///
    /// This example demonstrates:
    /// 1. Conditional/Else/EndConditional tags created by LibreOffice as inline SDTs
    /// 2. Coexistence of DocumentAssembler tags and Word MERGEFIELD fields
    ///
    /// The SDK should process DA tags while preserving merge fields untouched.
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("===========================================");
            Console.WriteLine("Example 16: Custom Font + Mixed Tags Test");
            Console.WriteLine("===========================================\n");

            // --- Part 1: Original custom font conditional test ---
            Console.WriteLine("--- Part 1: Inline SDT Conditional ---\n");

            Console.WriteLine("Test 1a: Condition TRUE (x=1) -> should show '11111'");
            Console.WriteLine("------------------------------------------");
            GenerateDocument(
                "TemplateDocument.docx",
                new XElement("Root", new XElement("x", "1")),
                "Output_ConditionTrue.docx"
            );

            Console.WriteLine("\nTest 1b: Condition FALSE (x=0) -> should show '22222'");
            Console.WriteLine("------------------------------------------");
            GenerateDocument(
                "TemplateDocument.docx",
                new XElement("Root", new XElement("x", "0")),
                "Output_ConditionFalse.docx"
            );

            // --- Part 2: Mixed DA tags + Merge Fields ---
            Console.WriteLine("\n--- Part 2: Mixed DA Tags + Merge Fields ---\n");

            var mixedTemplatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MixedTemplate.docx");
            CreateMixedTemplate(mixedTemplatePath);

            Console.WriteLine("Test 2a: Premium member -> should show premium content + merge fields preserved");
            Console.WriteLine("------------------------------------------");
            var output2a = GenerateDocument(
                "MixedTemplate.docx",
                new XElement("Customer",
                    new XElement("Name", "Mario Rossi"),
                    new XElement("Type", "Premium"),
                    new XElement("Discount", "20%")),
                "Output_Mixed_Premium.docx"
            );

            Console.WriteLine("\nTest 2b: Standard member -> should show standard content + merge fields preserved");
            Console.WriteLine("------------------------------------------");
            var output2b = GenerateDocument(
                "MixedTemplate.docx",
                new XElement("Customer",
                    new XElement("Name", "Luigi Bianchi"),
                    new XElement("Type", "Standard"),
                    new XElement("Discount", "5%")),
                "Output_Mixed_Standard.docx"
            );

            // --- Verification ---
            Console.WriteLine("\n--- Verification ---\n");
            VerifyOutput("Output_Mixed_Premium.docx", "Mario Rossi", "Premium", hasMergeFields: true);
            VerifyOutput("Output_Mixed_Standard.docx", "Luigi Bianchi", "Standard", hasMergeFields: true);

            Console.WriteLine("\n===========================================");
            Console.WriteLine("All tests completed!");
            Console.WriteLine("===========================================");
        }

        static string GenerateDocument(string templateName, XElement data, string outputName)
        {
            var outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, outputName);
            try
            {
                var templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, templateName);
                var wmlTemplate = new WmlDocument(templatePath);

                var wmlAssembled = DocumentAssembler.Core.DocumentAssembler.AssembleDocument(
                    wmlTemplate,
                    data,
                    out bool templateError
                );

                wmlAssembled.SaveAs(outputPath);

                if (templateError)
                {
                    Console.WriteLine($"  WARNING: Template errors detected! Output: {outputPath}");
                }
                else
                {
                    Console.WriteLine($"  OK: Generated {outputPath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  ERROR: {ex.Message}");
            }
            return outputPath;
        }

        /// <summary>
        /// Creates a docx template that mixes:
        /// - DA Content tags (inline SDT, like LibreOffice creates)
        /// - DA Conditional/Else/EndConditional tags (inline SDT)
        /// - Word MERGEFIELD fields (complex field: fldChar + instrText)
        /// </summary>
        static void CreateMixedTemplate(string path)
        {
            using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = new Body();

            // Paragraph 1: "Gentile [Content:Name],"
            body.Append(CreateParagraphWithInlineSdt("Gentile ", "<Content Select=\"./Name\" />", ","));

            // Paragraph 2: empty
            body.Append(new Paragraph());

            // Paragraph 3: "Indirizzo: «Address»" (merge field)
            body.Append(CreateParagraphWithMergeField("Indirizzo: ", "Address"));

            // Paragraph 4: "Citta: «City»" (merge field)
            body.Append(CreateParagraphWithMergeField("Citta: ", "City"));

            // Paragraph 5: empty
            body.Append(new Paragraph());

            // Conditional: Select="./Type" Match="Premium"
            body.Append(CreateSdtOnlyParagraph("<Conditional Select=\"./Type\" Match=\"Premium\" />"));

            // Premium content
            body.Append(new Paragraph(new Run(new Text("Sei un cliente PREMIUM! Il tuo sconto e' del "))));
            body.Append(CreateParagraphWithInlineSdt("", "<Content Select=\"./Discount\" />", "."));

            // Paragraph with merge field INSIDE the conditional true branch
            body.Append(CreateParagraphWithMergeField("Codice promozione: ", "PromoCode"));

            // Else
            body.Append(CreateSdtOnlyParagraph("<Else />"));

            // Standard content
            body.Append(new Paragraph(new Run(new Text("Sei un cliente standard. Il tuo sconto e' del "))));
            body.Append(CreateParagraphWithInlineSdt("", "<Content Select=\"./Discount\" />", "."));

            // Paragraph with merge field INSIDE the conditional false branch
            body.Append(CreateParagraphWithMergeField("Riferimento: ", "RefCode"));

            // EndConditional
            body.Append(CreateSdtOnlyParagraph("<EndConditional />"));

            // Paragraph after conditional with another merge field
            body.Append(new Paragraph());
            body.Append(CreateParagraphWithMergeField("Data: ", "Date"));

            // Section properties
            body.Append(new SectionProperties(
                new PageSize { Width = 12240, Height = 15840 },
                new PageMargin { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440 }));

            mainPart.Document.Append(body);
            mainPart.Document.Save();
        }

        /// <summary>
        /// Creates a paragraph with an inline SDT containing a DA tag.
        /// This mimics how LibreOffice creates content controls.
        /// Format: [prefix text] [SDT with tag] [suffix text]
        /// </summary>
        static Paragraph CreateParagraphWithInlineSdt(string prefix, string tagXml, string suffix)
        {
            var para = new Paragraph();
            if (!string.IsNullOrEmpty(prefix))
            {
                para.Append(new Run(new Text(prefix) { Space = SpaceProcessingModeValues.Preserve }));
            }

            var sdt = new SdtRun();
            sdt.Append(new SdtProperties(new SdtContentText()));
            var sdtContent = new SdtContentRun();
            sdtContent.Append(new Run(new Text(tagXml)));
            sdt.Append(sdtContent);
            para.Append(sdt);

            if (!string.IsNullOrEmpty(suffix))
            {
                para.Append(new Run(new Text(suffix) { Space = SpaceProcessingModeValues.Preserve }));
            }
            return para;
        }

        /// <summary>
        /// Creates a paragraph containing only an inline SDT with a DA tag.
        /// Used for block-level tags like Conditional, Else, EndConditional.
        /// </summary>
        static Paragraph CreateSdtOnlyParagraph(string tagXml)
        {
            var para = new Paragraph();
            var sdt = new SdtRun();
            sdt.Append(new SdtProperties(new SdtContentText()));
            var sdtContent = new SdtContentRun();
            sdtContent.Append(new Run(new Text(tagXml)));
            sdt.Append(sdtContent);
            para.Append(sdt);
            return para;
        }

        /// <summary>
        /// Creates a paragraph with a complex MERGEFIELD.
        /// Format: [prefix] «FieldName»
        /// Uses fldChar begin/separate/end pattern.
        /// </summary>
        static Paragraph CreateParagraphWithMergeField(string prefix, string fieldName)
        {
            var para = new Paragraph();

            if (!string.IsNullOrEmpty(prefix))
            {
                para.Append(new Run(new Text(prefix) { Space = SpaceProcessingModeValues.Preserve }));
            }

            // fldChar begin
            para.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }));

            // instrText
            para.Append(new Run(new FieldCode($" MERGEFIELD {fieldName} ") { Space = SpaceProcessingModeValues.Preserve }));

            // fldChar separate
            para.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }));

            // display value (placeholder)
            para.Append(new Run(new Text($"\u00AB{fieldName}\u00BB")));

            // fldChar end
            para.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.End }));

            return para;
        }

        /// <summary>
        /// Verifies the output document:
        /// - DA Content tags replaced with actual values
        /// - Conditional logic applied correctly
        /// - Merge fields preserved (instrText still present)
        /// </summary>
        static void VerifyOutput(string outputName, string expectedName, string expectedType, bool hasMergeFields)
        {
            var outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, outputName);
            if (!File.Exists(outputPath))
            {
                Console.WriteLine($"  FAIL: {outputName} not found");
                return;
            }

            Console.WriteLine($"Verifying {outputName}:");

            using var doc = WordprocessingDocument.Open(outputPath, false);
            var body = doc.MainDocumentPart?.Document?.Body;
            if (body == null)
            {
                Console.WriteLine($"  FAIL: No body found");
                return;
            }

            var fullText = body.InnerText;

            // Check Content tag was replaced
            if (fullText.Contains(expectedName))
            {
                Console.WriteLine($"  OK: Content tag replaced -> found '{expectedName}'");
            }
            else
            {
                Console.WriteLine($"  FAIL: Content tag NOT replaced -> '{expectedName}' not found in text");
            }

            // Check conditional logic
            if (expectedType == "Premium")
            {
                if (fullText.Contains("PREMIUM"))
                    Console.WriteLine($"  OK: Premium branch shown");
                else
                    Console.WriteLine($"  FAIL: Premium branch NOT found");

                if (fullText.Contains("cliente standard"))
                    Console.WriteLine($"  FAIL: Standard branch should NOT appear for Premium");
                else
                    Console.WriteLine($"  OK: Standard branch correctly hidden");
            }
            else
            {
                if (fullText.Contains("cliente standard"))
                    Console.WriteLine($"  OK: Standard branch shown");
                else
                    Console.WriteLine($"  FAIL: Standard branch NOT found");

                if (fullText.Contains("PREMIUM"))
                    Console.WriteLine($"  FAIL: Premium branch should NOT appear for Standard");
                else
                    Console.WriteLine($"  OK: Premium branch correctly hidden");
            }

            // Check merge fields preserved
            if (hasMergeFields)
            {
                var xDoc = doc.MainDocumentPart!.GetXDocument();
                var ns = xDoc.Root!.GetDefaultNamespace();
                if (ns == XNamespace.None)
                {
                    ns = xDoc.Root.GetNamespaceOfPrefix("w") ?? XNamespace.None;
                }
                var W_instrText = ns + "instrText";
                var W_fldChar = ns + "fldChar";

                var instrTexts = xDoc.Descendants(W_instrText).ToList();
                var fldChars = xDoc.Descendants(W_fldChar).ToList();

                if (instrTexts.Any(t => t.Value.Contains("MERGEFIELD")))
                {
                    Console.WriteLine($"  OK: Merge fields preserved ({instrTexts.Count} instrText elements found)");

                    // Check which merge fields survived (some may have been removed by conditional)
                    foreach (var instr in instrTexts)
                    {
                        var fieldText = instr.Value.Trim();
                        Console.WriteLine($"       -> {fieldText}");
                    }
                }
                else
                {
                    Console.WriteLine($"  FAIL: Merge fields NOT found in output (instrText count: {instrTexts.Count}, fldChar count: {fldChars.Count})");
                }
            }
        }
    }
}
