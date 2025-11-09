using DocumentAssembler.Core;
using System.Xml.Linq;
using Xunit;
using Xunit.Abstractions;

namespace DocumentAssembler.Tests
{
    /// <summary>
    /// Test suite for missing fields behavior with and without Optional attribute.
    ///
    /// DEFAULT BEHAVIOR:
    /// =================
    /// By default, all fields are OPTIONAL. Missing fields in the XML data will NOT cause
    /// errors - they will simply appear as empty in the generated document.
    ///
    /// REQUIRING FIELDS:
    /// =================
    /// To make a field REQUIRED (causing an error if missing), explicitly set Optional="false":
    ///
    ///   <#Content Select="Customer/Name" Optional="false"#>
    ///
    /// This will cause an XPathException if Customer/Name is not present in the data.
    ///
    /// BACKWARD COMPATIBILITY:
    /// =======================
    /// Old templates with Optional="true" will continue to work as expected.
    /// Old templates WITHOUT the Optional attribute will now be more forgiving (optional by default).
    ///
    /// SUPPORTED TAGS:
    /// ===============
    /// The Optional attribute is supported on:
    /// - Content (text placeholders)
    /// - Image (image placeholders)
    /// - Repeat (repeating blocks)
    /// </summary>
    public class MissingFieldsTests
    {
        private readonly ITestOutputHelper _output;

        public MissingFieldsTests(ITestOutputHelper output)
        {
            _output = output;
        }

        /// <summary>
        /// Demonstrates that missing fields are now OPTIONAL by default.
        /// This test uses DA003 which expects a field that doesn't exist.
        /// Since Optional is now true by default, this should NOT cause an error.
        /// </summary>
        [Fact]
        public void MissingField_OptionalByDefault_NoError()
        {
            // DA003 expects Customer/MissingField which doesn't exist in DA-Data.xml
            // With the new default behavior (Optional=true by default), this should NOT cause an error
            var result = TestUtil_RunTest("DA003-Select-XPathFindsNoData.docx", "DA-Data.xml");

            _output.WriteLine("Test: Missing field (optional by default)");
            _output.WriteLine($"Template Error: {result.HasError}");
            _output.WriteLine($"Document contains error message: {result.ContainsErrorText}");

            // Verify that NO error occurred (fields are optional by default now)
            Assert.False(result.HasError, "Template should NOT fail when field is missing (optional by default)");
        }

        /// <summary>
        /// Demonstrates that missing OPTIONAL fields do NOT cause template errors.
        /// This test uses DA004 which has Optional="true" for a missing field.
        /// </summary>
        [Fact]
        public void MissingOptionalField_DoesNotCauseError()
        {
            // DA004 expects Customer/MissingField with Optional="true"
            var result = TestUtil_RunTest("DA004-Select-XPathFindsNoDataOptional.docx", "DA-Data.xml");

            _output.WriteLine("Test: Missing OPTIONAL field");
            _output.WriteLine($"Template Error: {result.HasError}");
            _output.WriteLine($"Document contains error message: {result.ContainsErrorText}");

            // Verify that NO error occurred
            Assert.False(result.HasError, "Template should succeed when optional field is missing");
            Assert.False(result.ContainsErrorText, "Document should NOT contain error message");
        }

        /// <summary>
        /// Demonstrates Optional behavior with Repeat tag.
        /// DA257 has Optional="true" on a Repeat that references missing data.
        /// </summary>
        [Fact]
        public void MissingOptionalRepeat_DoesNotCauseError()
        {
            // DA257 has a Repeat with Optional="true" for data that doesn't exist
            var result = TestUtil_RunTest("DA257-OptionalRepeat.docx", "DA-Data.xml");

            _output.WriteLine("Test: Missing OPTIONAL repeat data");
            _output.WriteLine($"Template Error: {result.HasError}");
            _output.WriteLine($"Document contains error message: {result.ContainsErrorText}");

            // Verify that NO error occurred
            Assert.False(result.HasError, "Template should succeed when optional repeat data is missing");
            Assert.False(result.ContainsErrorText, "Document should NOT contain error message");
        }

        /// <summary>
        /// Helper class to encapsulate test results
        /// </summary>
        private class TestResult
        {
            public bool HasError { get; set; }
            public bool ContainsErrorText { get; set; }
            public string DocumentText { get; set; } = string.Empty;
        }

        /// <summary>
        /// Helper method to run a test and return results
        /// </summary>
        private TestResult TestUtil_RunTest(string templateName, string dataName)
        {
            var sourceDir = new System.IO.DirectoryInfo("TestFiles/");
            var templateDocx = new System.IO.FileInfo(System.IO.Path.Combine(sourceDir.FullName, templateName));
            var dataFile = new System.IO.FileInfo(System.IO.Path.Combine(sourceDir.FullName, dataName));

            var wmlTemplate = new WmlDocument(templateDocx.FullName);
            var xmldata = XElement.Load(dataFile.FullName);

            var assembled = Core.DocumentAssembler.AssembleDocument(wmlTemplate, xmldata, out var templateError);

            // Extract document text
            using var ms = new System.IO.MemoryStream(assembled.DocumentByteArray);
            using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(ms, false);
            var body = doc.MainDocumentPart?.Document.Body;
            var text = body?.InnerText ?? string.Empty;

            return new TestResult
            {
                HasError = templateError,
                ContainsErrorText = text.Contains("Error", System.StringComparison.OrdinalIgnoreCase),
                DocumentText = text
            };
        }
    }
}
