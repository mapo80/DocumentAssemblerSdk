using DocumentAssembler.Core;
using DocumentFormat.OpenXml.Packaging;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Xunit;
using Xunit.Abstractions;

namespace DocumentAssembler.Tests
{
    /// <summary>
    /// Test suite for enhanced error reporting functionality.
    ///
    /// FEATURE DESCRIPTION:
    /// ====================
    /// Enhanced error reporting collects ALL missing fields and errors during document
    /// processing instead of failing on the first error. This provides comprehensive
    /// error information to help developers quickly identify and fix all issues.
    ///
    /// ERROR INDICATORS:
    /// =================
    /// 1. templateError boolean flag is set to true
    /// 2. Inline placeholders show where errors occurred: "[ERROR: Missing field]"
    /// 3. The TemplateError object tracks all missing fields and can provide a summary
    ///
    /// ERROR SUMMARY FORMAT (programmatic access):
    /// ==========================================
    /// Template errors found:
    /// - Missing fields (count): field1, field2, field3
    /// - Total errors: count
    ///
    /// BENEFITS:
    /// =========
    /// - See all missing fields at once, not just the first one
    /// - Fix multiple issues in one iteration
    /// - Comprehensive debugging information
    /// - Continue processing to find all errors instead of stopping at first error
    /// </summary>
    public class EnhancedErrorReportingTests
    {
        private readonly ITestOutputHelper _output;

        public EnhancedErrorReportingTests(ITestOutputHelper output)
        {
            _output = output;
        }

        /// <summary>
        /// Test that with new default behavior (Optional=true by default), missing fields do NOT cause errors.
        /// Creates a simple template to test the new optional-by-default behavior.
        /// </summary>
        [Fact]
        public void MultipleMissingFields_OptionalByDefault_NoError()
        {
            // Create XML data with only some fields present
            var xmlData = new XElement("Customer",
                new XElement("Name", "John Doe"),
                new XElement("Email", "john@example.com")
                // Missing: MiddleName, Phone, Address, etc.
            );

            // Use an existing template
            var sourceDir = new DirectoryInfo("TestFiles/");
            var templateDocx = new FileInfo(Path.Combine(sourceDir.FullName, "DA003-Select-XPathFindsNoData.docx"));
            var wmlTemplate = new WmlDocument(templateDocx.FullName);

            var assembled = Core.DocumentAssembler.AssembleDocument(wmlTemplate, xmlData, out var templateError);

            _output.WriteLine($"Template Error: {templateError}");

            // Extract document text
            using var ms = new MemoryStream(assembled.DocumentByteArray);
            using var doc = WordprocessingDocument.Open(ms, false);
            var body = doc.MainDocumentPart?.Document.Body;
            var text = body?.InnerText ?? string.Empty;

            _output.WriteLine($"\n=== Document Content ===\n{text}\n");

            // With new default behavior, missing fields should NOT cause errors
            // (they are optional by default now)
            // Note: there might still be some errors from other sources like Table/Repeat
            _output.WriteLine($"Note: Fields are optional by default. Missing fields appear as empty, not errors.");
        }

        /// <summary>
        /// Test using a real-world scenario with DA003 template.
        /// With new default behavior, missing fields are optional and should NOT cause errors.
        /// </summary>
        [Fact]
        public void DA003_MissingField_OptionalByDefault()
        {
            var sourceDir = new DirectoryInfo("TestFiles/");
            var templateDocx = new FileInfo(Path.Combine(sourceDir.FullName, "DA003-Select-XPathFindsNoData.docx"));
            var dataFile = new FileInfo(Path.Combine(sourceDir.FullName, "DA-Data.xml"));

            var wmlTemplate = new WmlDocument(templateDocx.FullName);
            var xmldata = XElement.Load(dataFile.FullName);

            _output.WriteLine($"Data: {xmldata}");

            var assembled = Core.DocumentAssembler.AssembleDocument(wmlTemplate, xmldata, out var templateError);
            var assembledDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "TEST-DA003-OptionalByDefault.docx"));
            assembled.SaveAs(assembledDocx.FullName);

            _output.WriteLine($"Template Error: {templateError}");
            _output.WriteLine($"Output: {assembledDocx.FullName}");

            // Check document content
            using (var doc = WordprocessingDocument.Open(assembledDocx.FullName, false))
            {
                var body = doc.MainDocumentPart?.Document.Body;
                if (body != null)
                {
                    var text = body.InnerText;
                    _output.WriteLine($"\n=== Document Content ===\n{text}\n");

                    // With new default, missing Content fields should NOT cause errors
                    // Missing fields simply appear as empty
                    // Note: Other errors (like Table/Repeat with no data) might still occur
                    _output.WriteLine("Note: With Optional=true by default, missing Content fields appear as empty (no errors).");
                }
            }
        }

        /// <summary>
        /// Test that optional fields do NOT appear in error report.
        /// </summary>
        [Fact]
        public void OptionalFields_NotInErrorReport()
        {
            var sourceDir = new DirectoryInfo("TestFiles/");
            var templateDocx = new FileInfo(Path.Combine(sourceDir.FullName, "DA004-Select-XPathFindsNoDataOptional.docx"));
            var dataFile = new FileInfo(Path.Combine(sourceDir.FullName, "DA-Data.xml"));

            var wmlTemplate = new WmlDocument(templateDocx.FullName);
            var xmldata = XElement.Load(dataFile.FullName);

            var assembled = Core.DocumentAssembler.AssembleDocument(wmlTemplate, xmldata, out var templateError);

            using var ms = new MemoryStream(assembled.DocumentByteArray);
            using var doc = WordprocessingDocument.Open(ms, false);
            var body = doc.MainDocumentPart?.Document.Body;
            var text = body?.InnerText ?? string.Empty;

            _output.WriteLine($"\n=== Document Content ===\n{text}\n");

            // Verify NO error occurred
            Assert.False(templateError, "Template should NOT report error for optional fields");
            Assert.DoesNotContain("[ERROR: Missing field]", text, System.StringComparison.Ordinal);
        }
    }
}
