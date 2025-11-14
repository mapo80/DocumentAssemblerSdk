using DocumentAssembler.Core;
using DocumentAssembler.Core.Exceptions;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;
using Xunit;

namespace DocumentAssembler.Tests;

public class OpenXmlDocumentTests
{
    [Fact]
    public void OpenXmlPowerToolsDocument_CopyConstructor_PreservesBytes()
    {
        var original = new OpenXmlPowerToolsDocument(TestDocumentBuilder.CreateWmlDocument().DocumentByteArray);
        var copy = new OpenXmlPowerToolsDocument(original);
        var transitionalCopy = new OpenXmlPowerToolsDocument(original, convertToTransitional: true);

        Assert.Equal(original.DocumentByteArray, copy.DocumentByteArray);
        Assert.Equal("Unnamed Document", copy.GetName());
        Assert.NotEmpty(transitionalCopy.DocumentByteArray);
    }

    [Fact]
    public void OpenXmlPowerToolsDocument_SaveAndGetName_Works()
    {
        var bytes = TestDocumentBuilder.CreateWordDocument("SaveTest");
        var tempFile = Path.Combine(Path.GetTempPath(), $"save-test-{Guid.NewGuid():N}.docx");
        try
        {
            var stream = new MemoryStream();
            stream.Write(bytes, 0, bytes.Length);
            stream.Position = 0;
            var doc = new OpenXmlPowerToolsDocument(tempFile, stream);
            doc.Save();
            Assert.True(File.Exists(tempFile));
            doc.SaveAs(tempFile + "copy");
            Assert.True(File.Exists(tempFile + "copy"));
        }
        finally
        {
            if (File.Exists(tempFile)) File.Delete(tempFile);
            if (File.Exists(tempFile + "copy")) File.Delete(tempFile + "copy");
        }
    }

    [Fact]
    public void OpenXmlPowerToolsDocument_ConvertToTransitional_FromBytes()
    {
        var bytes = TestDocumentBuilder.CreateWordDocument("Transitional");
        var doc = new OpenXmlPowerToolsDocument(bytes, convertToTransitional: true);
        Assert.Equal("Unnamed Document", doc.GetName());
        Assert.NotEmpty(doc.DocumentByteArray);
    }

    [Fact]
    public void OpenXmlPowerToolsDocument_ConvertToTransitional_FromFile()
    {
        var bytes = TestDocumentBuilder.CreateWordDocument("FileConversion");
        var tempFile = Path.Combine(Path.GetTempPath(), $"transitional-{Guid.NewGuid():N}.docx");
        File.WriteAllBytes(tempFile, bytes);
        try
        {
            var doc = new OpenXmlPowerToolsDocument(tempFile, convertToTransitional: true);
            Assert.NotEmpty(doc.DocumentByteArray);
        }
        finally
        {
            if (File.Exists(tempFile))
            {
                File.Delete(tempFile);
            }
        }
    }

    [Fact]
    public void OpenXmlPowerToolsDocument_DetectsDocumentTypes()
    {
        var spreadsheet = new OpenXmlPowerToolsDocument(TestDocumentBuilder.CreateSpreadsheetDocument());
        Assert.Equal(typeof(SpreadsheetDocument), spreadsheet.GetDocumentType());

        var presentation = new OpenXmlPowerToolsDocument(TestDocumentBuilder.CreatePresentationDocument());
        Assert.Equal(typeof(PresentationDocument), presentation.GetDocumentType());
        var package = new OpenXmlPowerToolsDocument(TestDocumentBuilder.CreateWordDocument(), convertToTransitional: false);
        using var stream = new MemoryStream();
        package.WriteByteArray(stream);
        Assert.NotEqual(0, stream.Length);
        var unnamed = new OpenXmlPowerToolsDocument(stream.ToArray());
        Assert.Throws<InvalidOperationException>(() => unnamed.Save());
    }

    [Fact]
    public void OpenXmlPowerToolsDocument_ConvertsAllDocumentTypes()
    {
        var spreadsheet = new OpenXmlPowerToolsDocument(TestDocumentBuilder.CreateSpreadsheetDocument(), convertToTransitional: true);
        Assert.Equal(typeof(SpreadsheetDocument), spreadsheet.GetDocumentType());

        var presentation = new OpenXmlPowerToolsDocument(TestDocumentBuilder.CreatePresentationDocument(), convertToTransitional: true);
        Assert.Equal(typeof(PresentationDocument), presentation.GetDocumentType());

        var wordFile = TestDocumentBuilder.CreateWordDocument("ConvertAll");
        var tempFile = Path.Combine(Path.GetTempPath(), $"convert-all-{Guid.NewGuid():N}.docx");
        File.WriteAllBytes(tempFile, wordFile);
        try
        {
            var doc = new OpenXmlPowerToolsDocument(tempFile, convertToTransitional: false);
            Assert.Equal(typeof(WordprocessingDocument), doc.GetDocumentType());
        }
        finally
        {
            if (File.Exists(tempFile))
            {
                File.Delete(tempFile);
            }
        }
    }

    [Fact]
    public void OpenXmlPowerToolsDocument_FileBasedConstructors()
    {
        var bytes = TestDocumentBuilder.CreateWordDocument("FileBased");
        var tempFile = Path.Combine(Path.GetTempPath(), $"file-based-{Guid.NewGuid():N}.docx");
        File.WriteAllBytes(tempFile, bytes);
        try
        {
            var doc = new OpenXmlPowerToolsDocument(tempFile, convertToTransitional: true);
            Assert.NotEmpty(doc.DocumentByteArray);

            var stream = new MemoryStream();
            stream.Write(bytes, 0, bytes.Length);
            var docWithStream = new OpenXmlPowerToolsDocument(tempFile, stream, convertToTransitional: true);
            Assert.NotEmpty(docWithStream.DocumentByteArray);
        }
        finally
        {
            if (File.Exists(tempFile))
            {
                File.Delete(tempFile);
            }
        }
    }

    [Fact]
    public void OpenXmlMemoryStreamDocument_WordprocessingLifecycle()
    {
        using var memoryDoc = OpenXmlMemoryStreamDocument.CreateWordprocessingDocument();
        Assert.Equal(typeof(WordprocessingDocument), memoryDoc.GetDocumentType());
        using var wDoc = memoryDoc.GetWordprocessingDocument();
        Assert.NotNull(wDoc.MainDocumentPart);
        Assert.Throws<PowerToolsDocumentException>(() => memoryDoc.GetSpreadsheetDocument());
    }

    [Fact]
    public void OpenXmlMemoryStreamDocument_SpreadsheetAndPresentation()
    {
        using var spreadsheet = OpenXmlMemoryStreamDocument.CreateSpreadsheetDocument();
        Assert.Equal(typeof(SpreadsheetDocument), spreadsheet.GetDocumentType());
        using var sDoc = spreadsheet.GetSpreadsheetDocument();
        Assert.NotNull(sDoc.WorkbookPart);

        using var presentation = OpenXmlMemoryStreamDocument.CreatePresentationDocument();
        Assert.Equal(typeof(PresentationDocument), presentation.GetDocumentType());
        using var pDoc = presentation.GetPresentationDocument();
        Assert.NotNull(pDoc.PresentationPart);
    }

    [Fact]
    public void OpenXmlMemoryStreamDocument_PackageAccess()
    {
        using var packageDoc = OpenXmlMemoryStreamDocument.CreatePackage();
        var package = packageDoc.GetPackage();
        Assert.NotNull(package);
    }

    [Fact]
    public void WmlDocument_MainDocumentPart_ExposesComments()
    {
        var bytes = TestDocumentBuilder.CreateWordDocumentWithComment("With comment");
        var wml = new WmlDocument("comments.docx", bytes);
        var main = wml.MainDocumentPart;
        Assert.NotNull(main);
        Assert.NotNull(main.WordprocessingCommentsPart);
    }

    [Fact]
    public void WmlDocument_ThrowsForNonWordDocument()
    {
        var spreadsheet = new OpenXmlPowerToolsDocument(TestDocumentBuilder.CreateSpreadsheetDocument());
        Assert.Throws<PowerToolsDocumentException>(() => new WmlDocument(spreadsheet));
    }
}
