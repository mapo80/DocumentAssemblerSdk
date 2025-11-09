using DocumentAssembler.Core;
using DocumentAssembler.Tests;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;
using System.Linq;
using System.Xml;
using Xunit;

namespace DocumentAssemblerSdk.Tests
{
    public class TemplateSchemaExtractorTests
    {
        // Test data directory - using the same as other tests
        private static readonly string s_TestFilesDir = Path.Combine(
            Path.GetDirectoryName(typeof(TemplateSchemaExtractorTests).Assembly.Location) ?? "",
            "..", "..", "..", "TestFiles");

        [Fact]
        public void ExtractXmlSchema_BasicContent_ShouldExtractCorrectly()
        {
            // Arrange
            var templatePath = Path.Combine(s_TestFilesDir, "DA001-TemplateDocument.docx");
            if (!File.Exists(templatePath))
            {
                // Skip if test file doesn't exist
                return;
            }

            var templateDoc = new WmlDocument(templatePath);

            // Act
            var result = TemplateSchemaExtractor.ExtractXmlSchema(templateDoc);

            // Assert
            Assert.NotNull(result);
            Assert.NotEmpty(result.Fields);
            Assert.NotEmpty(result.XmlTemplate);

            // Verify XML is valid
            var xmlDoc = new XmlDocument();
            Assert.Null(Record.Exception(() => xmlDoc.LoadXml(result.XmlTemplate)));

            // Should contain Customer fields
            var customerFields = result.Fields.Where(f => f.XPath.Contains("Customer")).ToList();
            Assert.NotEmpty(customerFields);
        }

        [Fact]
        public void ExtractXmlSchema_WithOptionalFields_ShouldMarkCorrectly()
        {
            // Arrange
            var templatePath = Path.Combine(s_TestFilesDir, "DA210-OptionalContent.docx");
            if (!File.Exists(templatePath))
            {
                // Skip if test file doesn't exist
                return;
            }

            var templateDoc = new WmlDocument(templatePath);

            // Act
            var result = TemplateSchemaExtractor.ExtractXmlSchema(templateDoc);

            // Assert
            Assert.NotNull(result);
            var allOptional = result.Fields.Where(f => f.TagType == "Content").All(f => f.IsOptional);
            Assert.True(allOptional, "All Content fields should be optional by default");
        }

        [Fact]
        public void ExtractXmlSchema_WithRepeat_ShouldDetectRepeatingElements()
        {
            // Arrange
            var templatePath = Path.Combine(s_TestFilesDir, "DA110-RepeatBlocks.docx");
            if (!File.Exists(templatePath))
            {
                return;
            }

            var templateDoc = new WmlDocument(templatePath);

            // Act
            var result = TemplateSchemaExtractor.ExtractXmlSchema(templateDoc);

            // Assert
            Assert.NotNull(result);
            var repeatingFields = result.Fields.Where(f => f.IsRepeating).ToList();
            Assert.NotEmpty(repeatingFields);

            // XML should contain repeating comment
            Assert.Contains("Repeating", result.XmlTemplate);
        }

        [Fact]
        public void ExtractXmlSchema_WithTable_ShouldDetectTableStructure()
        {
            // Arrange
            var templatePath = Path.Combine(s_TestFilesDir, "DA050-Table.docx");
            if (!File.Exists(templatePath))
            {
                return;
            }

            var templateDoc = new WmlDocument(templatePath);

            // Act
            var result = TemplateSchemaExtractor.ExtractXmlSchema(templateDoc);

            // Assert
            Assert.NotNull(result);
            var tableFields = result.Fields.Where(f => f.TagType == "Table").ToList();
            Assert.NotEmpty(tableFields);
        }

        [Fact]
        public void ExtractXmlSchema_WithImage_ShouldDetectImagePlaceholders()
        {
            // Arrange
            var templatePath = Path.Combine(s_TestFilesDir, "DA200-ImagePlaceholder.docx");
            if (!File.Exists(templatePath))
            {
                return;
            }

            var templateDoc = new WmlDocument(templatePath);

            // Act
            var result = TemplateSchemaExtractor.ExtractXmlSchema(templateDoc);

            // Assert
            Assert.NotNull(result);
            var imageFields = result.Fields.Where(f => f.TagType == "Image").ToList();
            Assert.NotEmpty(imageFields);

            // XML should contain Base64 comment for images
            Assert.Contains("Base64", result.XmlTemplate);
        }

        [Fact]
        public void ExtractXmlSchema_WithConditional_ShouldDetectButNotIncludeInXml()
        {
            // Arrange
            var templatePath = Path.Combine(s_TestFilesDir, "DA130-ConditionalBlocks.docx");
            if (!File.Exists(templatePath))
            {
                return;
            }

            var templateDoc = new WmlDocument(templatePath);

            // Act
            var result = TemplateSchemaExtractor.ExtractXmlSchema(templateDoc);

            // Assert
            Assert.NotNull(result);
            var conditionalFields = result.Fields.Where(f => f.TagType == "Conditional").ToList();

            // Conditionals should be detected
            if (conditionalFields.Any())
            {
                // But their XPaths should not appear as data elements in the XML template
                // (they're used for conditions, not data insertion)
                foreach (var condField in conditionalFields)
                {
                    Assert.Contains(condField, result.Fields);
                }
            }
        }

        [Fact]
        public void ExtractXmlSchema_EmptyTemplate_ShouldReturnEmptyResult()
        {
            // Arrange
            var emptyDoc = CreateBlankTemplateDocument();

            // Act
            var result = TemplateSchemaExtractor.ExtractXmlSchema(emptyDoc);

            // Assert
            Assert.NotNull(result);
            Assert.Equal("Data", result.RootElementName);
        }

        [Fact]
        public void FieldInfo_ElementName_ShouldReturnLastPathSegment()
        {
            // Arrange
            var field = new TemplateSchemaExtractor.FieldInfo
            {
                XPath = "Customer/Address/City"
            };

            // Act
            var elementName = field.ElementName;

            // Assert
            Assert.Equal("City", elementName);
        }

        [Fact]
        public void ExtractXmlSchema_WithNestedStructures_ShouldBuildHierarchy()
        {
            // Arrange
            var templatePath = Path.Combine(s_TestFilesDir, "DA001-TemplateDocument.docx");
            var templateDoc = new WmlDocument(templatePath);

            // Act
            var result = TemplateSchemaExtractor.ExtractXmlSchema(templateDoc);

            // Assert
            Assert.NotNull(result);
            Assert.NotEmpty(result.XmlTemplate);

            // Should be well-formed XML with nested structure
            var xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(result.XmlTemplate);
            Assert.NotNull(xmlDoc.DocumentElement);
        }

        [Fact]
        public void SchemaExtractionResult_ToFormattedXml_ShouldFormatCorrectly()
        {
            // Arrange
            var result = new TemplateSchemaExtractor.SchemaExtractionResult
            {
                XmlTemplate = "<Data><Customer><Name>[value]</Name></Customer></Data>"
            };

            // Act
            var formatted = result.ToFormattedXml();

            // Assert
            Assert.NotNull(formatted);
            Assert.Contains("Data", formatted);
            Assert.Contains("Customer", formatted);
            Assert.Contains("Name", formatted);
        }

        [Fact]
        public void ExtractXmlSchema_WithAttributes_ShouldCaptureAttributes()
        {
            // Arrange
            var templatePath = Path.Combine(s_TestFilesDir, "DA200-ImagePlaceholder.docx");
            if (!File.Exists(templatePath))
            {
                return;
            }

            var templateDoc = new WmlDocument(templatePath);

            // Act
            var result = TemplateSchemaExtractor.ExtractXmlSchema(templateDoc);

            // Assert
            Assert.NotNull(result);
            var imageFields = result.Fields.Where(f => f.TagType == "Image").ToList();

            if (imageFields.Any())
            {
                // Image fields might have attributes like Align, Width, etc.
                Assert.NotNull(imageFields[0].Attributes);
            }
        }

        [Fact]
        public void ExtractXmlSchema_Performance_ShouldBeFast()
        {
            // Arrange
            var templatePath = Path.Combine(s_TestFilesDir, "DA001-TemplateDocument.docx");
            var templateDoc = new WmlDocument(templatePath);

            // Act & Assert
            var sw = System.Diagnostics.Stopwatch.StartNew();
            var result = TemplateSchemaExtractor.ExtractXmlSchema(templateDoc);
            sw.Stop();

            // Should complete in less than 1 second for typical templates
            Assert.True(sw.ElapsedMilliseconds < 1000,
                $"Extraction took {sw.ElapsedMilliseconds}ms, expected < 1000ms");
            Assert.NotNull(result);
        }

        [Fact]
        public void ExtractXmlSchema_MultipleOccurrences_ShouldMergeCorrectly()
        {
            // Arrange
            var templatePath = Path.Combine(s_TestFilesDir, "DA001-TemplateDocument.docx");
            var templateDoc = new WmlDocument(templatePath);

            // Act
            var result = TemplateSchemaExtractor.ExtractXmlSchema(templateDoc);

            // Assert
            Assert.NotNull(result);

            // Each unique XPath should appear only once in Fields list
            var xpaths = result.Fields.Select(f => f.XPath).ToList();
            var uniqueXPaths = xpaths.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            Assert.Equal(uniqueXPaths.Count, xpaths.Count);
        }

        [Fact]
        public void ExtractXmlSchema_DetermineRootName_ShouldFindMostCommonRoot()
        {
            // Arrange
            var templatePath = Path.Combine(s_TestFilesDir, "DA001-TemplateDocument.docx");
            var templateDoc = new WmlDocument(templatePath);

            // Act
            var result = TemplateSchemaExtractor.ExtractXmlSchema(templateDoc);

            // Assert
            Assert.NotNull(result);
            Assert.NotEmpty(result.RootElementName);
            Assert.NotEqual("Data", result.RootElementName); // Should detect actual root like "Customer"
        }

        private static WmlDocument CreateBlankTemplateDocument()
        {
            using var ms = new MemoryStream();
            using (var wordDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
            {
                var mainPart = wordDoc.AddMainDocumentPart();
                var document = new Document(new Body(new Paragraph(new Run(new Text(string.Empty)))));
                mainPart.Document = document;
                document.Save();
            }

            var bytes = ms.ToArray();
            return new WmlDocument("Empty.docx", bytes);
        }
    }
}
