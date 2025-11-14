using DocumentAssembler.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace DocumentAssembler.Tests;

internal static class TestDocumentBuilder
{
    public static byte[] CreateWordDocument(string text = "Hello")
    {
        using var ms = new MemoryStream();
        using (var document = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document, true))
        {
            var mainPart = document.AddMainDocumentPart();
            var body = new Body(new Paragraph(new Run(new Text(text))));
            mainPart.Document = new Document(body);
            mainPart.Document.Save();
        }

        return ms.ToArray();
    }

    public static byte[] CreateWordDocumentWithComment(string text)
    {
        using var ms = new MemoryStream();
        using (var document = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document, true))
        {
            var mainPart = document.AddMainDocumentPart();
            var commentId = "0";
            var paragraph = new Paragraph(
                new CommentRangeStart { Id = commentId },
                new Run(new Text(text)),
                new CommentRangeEnd { Id = commentId },
                new Run(new CommentReference { Id = commentId }));
            mainPart.Document = new Document(new Body(paragraph));
            var commentsPart = mainPart.AddNewPart<WordprocessingCommentsPart>();
            commentsPart.Comments = new Comments(new Comment
            {
                Id = commentId,
                Author = "Tester",
                Date = System.DateTime.UtcNow,
                InnerXml = "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:r><w:t>Note</w:t></w:r></w:p>"
            });
            commentsPart.Comments.Save();
            mainPart.Document.Save();
        }

        return ms.ToArray();
    }

    public static byte[] CreateSpreadsheetDocument()
    {
        using var ms = new MemoryStream();
        using (var document = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook, true))
        {
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();
            workbookPart.Workbook.Save();
        }

        return ms.ToArray();
    }

    public static byte[] CreatePresentationDocument()
    {
        using var ms = new MemoryStream();
        using (var document = PresentationDocument.Create(ms, PresentationDocumentType.Presentation, true))
        {
            var presentationPart = document.AddPresentationPart();
            presentationPart.Presentation = new DocumentFormat.OpenXml.Presentation.Presentation();
            presentationPart.Presentation.Save();
        }

        return ms.ToArray();
    }

    public static WmlDocument CreateWmlDocument(string text = "Sample")
    {
        var bytes = CreateWordDocument(text);
        return new WmlDocument("sample.docx", bytes);
    }
}
