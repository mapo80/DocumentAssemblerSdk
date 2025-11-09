using DocumentAssembler.Core;
using DocumentFormat.OpenXml.Packaging;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Xunit;
using Xunit.Abstractions;

namespace DocumentAssembler.Tests
{
    public class ElseDebugTest
    {
        private readonly ITestOutputHelper _output;

        public ElseDebugTest(ITestOutputHelper output)
        {
            _output = output;
        }

        [Fact]
        public void Debug_DA270_Premium()
        {
            var sourceDir = new DirectoryInfo("TestFiles/");
            var templateDocx = new FileInfo(Path.Combine(sourceDir.FullName, "DA270-ConditionalWithElse.docx"));
            var dataFile = new FileInfo(Path.Combine(sourceDir.FullName, "DA-ElseTestPremium.xml"));

            var wmlTemplate = new WmlDocument(templateDocx.FullName);
            var xmldata = XElement.Load(dataFile.FullName);

            _output.WriteLine($"Data: {xmldata}");

            var afterAssembling = Core.DocumentAssembler.AssembleDocument(wmlTemplate, xmldata, out var returnedTemplateError);
            var assembledDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "DEBUG-DA270-Premium.docx"));
            afterAssembling.SaveAs(assembledDocx.FullName);

            _output.WriteLine($"Template Error: {returnedTemplateError}");
            _output.WriteLine($"Output: {assembledDocx.FullName}");

            // Check document content for errors
            using (var doc = WordprocessingDocument.Open(assembledDocx.FullName, false))
            {
                var body = doc.MainDocumentPart?.Document.Body;
                if (body != null)
                {
                    var text = body.InnerText;
                    _output.WriteLine($"\n=== Document Content ===\n{text}\n");

                    // Find error messages
                    if (text.Contains("Error:") || text.Contains("error") || text.Contains("Error") || returnedTemplateError)
                    {
                        _output.WriteLine("\n⚠️  ERRORS FOUND IN DOCUMENT:");

                        var paragraphs = body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>();
                        foreach (var para in paragraphs)
                        {
                            var paraText = para.InnerText;
                            if (paraText.Contains("rror") || paraText.Contains("RROR"))
                            {
                                _output.WriteLine($"  - {paraText}");
                            }
                        }
                    }
                }
            }

            Assert.False(returnedTemplateError, "Template processing should not have errors");
        }

        [Fact]
        public void Debug_DA270_Standard()
        {
            var sourceDir = new DirectoryInfo("TestFiles/");
            var templateDocx = new FileInfo(Path.Combine(sourceDir.FullName, "DA270-ConditionalWithElse.docx"));
            var dataFile = new FileInfo(Path.Combine(sourceDir.FullName, "DA-ElseTestStandard.xml"));

            var wmlTemplate = new WmlDocument(templateDocx.FullName);
            var xmldata = XElement.Load(dataFile.FullName);

            _output.WriteLine($"Data: {xmldata}");

            var afterAssembling = Core.DocumentAssembler.AssembleDocument(wmlTemplate, xmldata, out var returnedTemplateError);
            var assembledDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "DEBUG-DA270-Standard.docx"));
            afterAssembling.SaveAs(assembledDocx.FullName);

            _output.WriteLine($"Template Error: {returnedTemplateError}");
            _output.WriteLine($"Output: {assembledDocx.FullName}");

            // Check document content
            using (var doc = WordprocessingDocument.Open(assembledDocx.FullName, false))
            {
                var body = doc.MainDocumentPart?.Document.Body;
                if (body != null)
                {
                    var text = body.InnerText;
                    _output.WriteLine($"\n=== Document Content ===\n{text}\n");

                    if (text.Contains("Error:") || text.Contains("error") || text.Contains("Error") || returnedTemplateError)
                    {
                        _output.WriteLine("\n⚠️  ERRORS FOUND IN DOCUMENT:");

                        var paragraphs = body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>();
                        foreach (var para in paragraphs)
                        {
                            var paraText = para.InnerText;
                            if (paraText.Contains("rror") || paraText.Contains("RROR"))
                            {
                                _output.WriteLine($"  - {paraText}");
                            }
                        }
                    }
                }
            }

            Assert.False(returnedTemplateError, "Template processing should not have errors");
        }

        [Fact]
        public void Debug_DA272_Premium()
        {
            var sourceDir = new DirectoryInfo("TestFiles/");
            var templateDocx = new FileInfo(Path.Combine(sourceDir.FullName, "DA272-NestedConditionalWithElse.docx"));
            var dataFile = new FileInfo(Path.Combine(sourceDir.FullName, "DA-ElseTestPremium.xml"));

            var wmlTemplate = new WmlDocument(templateDocx.FullName);
            var xmldata = XElement.Load(dataFile.FullName);

            _output.WriteLine($"Data: {xmldata}");

            var afterAssembling = Core.DocumentAssembler.AssembleDocument(wmlTemplate, xmldata, out var returnedTemplateError);
            var assembledDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "DEBUG-DA272-Premium.docx"));
            afterAssembling.SaveAs(assembledDocx.FullName);

            _output.WriteLine($"Template Error: {returnedTemplateError}");
            _output.WriteLine($"Output: {assembledDocx.FullName}");

            // Check document content for errors
            using (var doc = WordprocessingDocument.Open(assembledDocx.FullName, false))
            {
                var body = doc.MainDocumentPart?.Document.Body;
                if (body != null)
                {
                    var text = body.InnerText;
                    _output.WriteLine($"\n=== Document Content ===\n{text}\n");

                    // Find error messages
                    if (text.Contains("Error:") || text.Contains("error") || text.Contains("Error") || returnedTemplateError)
                    {
                        _output.WriteLine("\n⚠️  ERRORS FOUND IN DOCUMENT:");

                        var paragraphs = body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>();
                        foreach (var para in paragraphs)
                        {
                            var paraText = para.InnerText;
                            if (paraText.Contains("rror") || paraText.Contains("RROR"))
                            {
                                _output.WriteLine($"  - {paraText}");
                            }
                        }
                    }
                }
            }

            Assert.False(returnedTemplateError, "Template processing should not have errors");
        }

        [Fact]
        public void Debug_DA272_Standard()
        {
            var sourceDir = new DirectoryInfo("TestFiles/");
            var templateDocx = new FileInfo(Path.Combine(sourceDir.FullName, "DA272-NestedConditionalWithElse.docx"));
            var dataFile = new FileInfo(Path.Combine(sourceDir.FullName, "DA-ElseTestStandard.xml"));

            var wmlTemplate = new WmlDocument(templateDocx.FullName);
            var xmldata = XElement.Load(dataFile.FullName);

            _output.WriteLine($"Data: {xmldata}");

            var afterAssembling = Core.DocumentAssembler.AssembleDocument(wmlTemplate, xmldata, out var returnedTemplateError);
            var assembledDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "DEBUG-DA272-Standard.docx"));
            afterAssembling.SaveAs(assembledDocx.FullName);

            _output.WriteLine($"Template Error: {returnedTemplateError}");
            _output.WriteLine($"Output: {assembledDocx.FullName}");

            // Check document content
            using (var doc = WordprocessingDocument.Open(assembledDocx.FullName, false))
            {
                var body = doc.MainDocumentPart?.Document.Body;
                if (body != null)
                {
                    var text = body.InnerText;
                    _output.WriteLine($"\n=== Document Content ===\n{text}\n");

                    if (text.Contains("Error:") || text.Contains("error") || text.Contains("Error") || returnedTemplateError)
                    {
                        _output.WriteLine("\n⚠️  ERRORS FOUND IN DOCUMENT:");

                        var paragraphs = body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>();
                        foreach (var para in paragraphs)
                        {
                            var paraText = para.InnerText;
                            if (paraText.Contains("rror") || paraText.Contains("RROR"))
                            {
                                _output.WriteLine($"  - {paraText}");
                            }
                        }
                    }
                }
            }

            Assert.False(returnedTemplateError, "Template processing should not have errors");
        }
    }
}
