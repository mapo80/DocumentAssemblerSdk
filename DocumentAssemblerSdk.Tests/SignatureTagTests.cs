using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentAssembler.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;
using Xunit.Sdk;
using CoreDocumentAssembler = DocumentAssembler.Core.DocumentAssembler;

namespace DocumentAssembler.Tests;

public class SignatureTagTests
{
    [Fact]
    public void SignatureTag_WithSharpNotation_InsertsPlaceholderAndLabel()
    {
        var bytes = CreateDocxWithParagraph(
            new Paragraph(
                new Run(
                    new Text("<# <Signature Id=\"MainSigner\" Label=\"Firma Cliente\" Width=\"180px\" Height=\"48px\" /> #>")
                    {
                        Space = SpaceProcessingModeValues.Preserve
                    })));

        var template = new WmlDocument("signature.docx", bytes);
        var assembled = CoreDocumentAssembler.AssembleDocument(template, new XElement("Data"), out var hasError, out var summary);

        if (hasError)
        {
            var message = summary ?? "Template error";
            var dumpPath = Path.Combine(Path.GetTempPath(), $"signature-tag-{Guid.NewGuid():N}.docx");
            assembled.SaveAs(dumpPath);
            throw new XunitException($"Template error summary: {message}. Dump: {dumpPath}");
        }
        using var ms = new MemoryStream(assembled.DocumentByteArray);
        using var wordDoc = WordprocessingDocument.Open(ms, false);
        var texts = wordDoc.MainDocumentPart!.Document.Body!.Descendants<Text>().Select(t => t.Text).ToList();

        Assert.Contains(texts, t => t.Contains("Firma Cliente"));
        Assert.Contains(texts, t => t.Contains(SignaturePlaceholderSerializer.PlaceholderPrefix));
    }

    [Fact]
    public void SignatureTag_WithContentControlNotation_IsTransformed()
    {
        var signatureText = "<Signature Id=\"XmlSigner\" Label=\"Firma XML\" Width=\"150px\" Height=\"40px\" />";
        var sdt = new SdtRun(
            new SdtProperties(),
            new SdtContentRun(
                new Run(
                    new Text(signatureText)
                    {
                        Space = SpaceProcessingModeValues.Preserve
                    })));

        var bytes = CreateDocxWithParagraph(new Paragraph(sdt));
        var template = new WmlDocument("signature-sdt.docx", bytes);
        var assembled = CoreDocumentAssembler.AssembleDocument(template, new XElement("Root"), out var hasError, out var summary);

        if (hasError)
        {
            var message = summary ?? "Template error";
            var dumpPath = Path.Combine(Path.GetTempPath(), $"signature-tag-{Guid.NewGuid():N}.docx");
            assembled.SaveAs(dumpPath);
            throw new XunitException($"Template error summary: {message}. Dump: {dumpPath}");
        }

        using var ms = new MemoryStream(assembled.DocumentByteArray);
        using var doc = WordprocessingDocument.Open(ms, false);
        var texts = doc.MainDocumentPart!.Document.Body!.Descendants<Text>().Select(t => t.Text).ToList();

        Assert.Contains(texts, t => t.Contains("Firma XML"));
        Assert.Contains(texts, t => t.Contains(SignaturePlaceholderSerializer.PlaceholderPrefix));
    }

    [Fact]
    public void SignatureTag_InvalidPageHint_ReturnsTemplateError()
    {
        var bytes = CreateDocxWithParagraph(
            new Paragraph(
                new Run(
                    new Text("<# <Signature Id=\"BadSigner\" PageHint=\"zero\" /> #>")
                    {
                        Space = SpaceProcessingModeValues.Preserve
                    })));

        var template = new WmlDocument("invalid-pagehint.docx", bytes);
        _ = CoreDocumentAssembler.AssembleDocument(template, new XElement("Root"), out var hasError, out _);

        Assert.True(hasError);
    }

    [Fact]
    public void SignatureTag_InvalidWidth_ReportsError()
    {
        var bytes = CreateDocxWithParagraph(
            new Paragraph(
                new Run(
                    new Text("<# <Signature Id=\"WrongWidth\" Width=\"abc\" /> #>")
                    {
                        Space = SpaceProcessingModeValues.Preserve
                    })));

        var template = new WmlDocument("invalid-width.docx", bytes);
        _ = CoreDocumentAssembler.AssembleDocument(template, new XElement("Root"), out var hasError, out _);

        Assert.True(hasError);
    }

    private static byte[] CreateDocxWithParagraph(Paragraph paragraph)
    {
        using var ms = new MemoryStream();
        using (var document = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(paragraph));
            mainPart.Document.Save();
        }

        return ms.ToArray();
    }
}
