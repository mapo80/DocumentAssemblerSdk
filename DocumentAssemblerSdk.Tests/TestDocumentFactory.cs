using DocumentAssembler.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace DocumentAssembler.Tests;

internal static class TestDocumentFactory
{
    public static WmlDocument Create(string fileName, Action<WordDocumentBuilder> configure)
    {
        if (configure == null)
        {
            throw new ArgumentNullException(nameof(configure));
        }

        var builder = new WordDocumentBuilder(fileName);
        configure(builder);
        return builder.Build();
    }
}

internal sealed class WordDocumentBuilder
{
    private readonly string _fileName;
    private readonly List<XElement> _bodyElements = new List<XElement>();
    private readonly List<Action<WordprocessingDocument>> _partMutations = new List<Action<WordprocessingDocument>>();

    public WordDocumentBuilder(string fileName)
    {
        _fileName = fileName;
    }

    public WordDocumentBuilder AddBodyElement(XElement element)
    {
        if (element == null)
        {
            throw new ArgumentNullException(nameof(element));
        }

        _bodyElements.Add(new XElement(element));
        return this;
    }

    public WordDocumentBuilder AddParagraph(string text)
    {
        return AddBodyElement(new XElement(W.p,
            new XElement(W.r,
                new XElement(W.t, text ?? string.Empty))));
    }

    public WordDocumentBuilder AddDefaultHeader(params XElement[] paragraphs)
    {
        return AddHeader(HeaderFooterValues.Default, paragraphs);
    }

    public WordDocumentBuilder AddDefaultFooter(params XElement[] paragraphs)
    {
        return AddFooter(HeaderFooterValues.Default, paragraphs);
    }

    public WordDocumentBuilder Configure(Action<WordprocessingDocument> configure)
    {
        if (configure == null)
        {
            throw new ArgumentNullException(nameof(configure));
        }

        _partMutations.Add(configure);
        return this;
    }

    public WmlDocument Build()
    {
        using var ms = new MemoryStream();
        using (var wordDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text(string.Empty)))));
            mainPart.Document.Save();

            var xDoc = mainPart.GetXDocument();
            var body = xDoc.Root?.Element(W.body) ?? throw new InvalidOperationException("Document missing body.");
            body.RemoveNodes();

            if (_bodyElements.Count == 0)
            {
                body.Add(new XElement(W.p,
                    new XElement(W.r,
                        new XElement(W.t, string.Empty))));
            }
            else
            {
                foreach (var element in _bodyElements)
                {
                    body.Add(new XElement(element));
                }
            }

            if (!body.Elements().Any(e => e.Name == W.sectPr))
            {
                body.Add(new XElement(W.sectPr,
                    new XElement(W.pgSz),
                    new XElement(W.pgMar)));
            }

            mainPart.PutXDocument();

            foreach (var mutation in _partMutations)
            {
                mutation(wordDoc);
            }
        }

        return new WmlDocument(_fileName, ms.ToArray());
    }

    private WordDocumentBuilder AddHeader(HeaderFooterValues type, params XElement[] paragraphs)
    {
        _partMutations.Add(doc => AttachHeader(doc, type, paragraphs));
        return this;
    }

    private WordDocumentBuilder AddFooter(HeaderFooterValues type, params XElement[] paragraphs)
    {
        _partMutations.Add(doc => AttachFooter(doc, type, paragraphs));
        return this;
    }

    private static void AttachHeader(WordprocessingDocument document, HeaderFooterValues type, params XElement[] paragraphs)
    {
        var mainPart = document.MainDocumentPart ?? throw new InvalidOperationException("Document missing main part.");
        var headerPart = mainPart.AddNewPart<HeaderPart>();
        var headerDoc = CreatePartDocument(W.header, paragraphs);
        headerPart.PutXDocument(headerDoc);

        var relId = mainPart.GetIdOfPart(headerPart);
        var sectPr = GetOrCreateSectionProperties(mainPart);
        sectPr.Elements(W.headerReference)
            .Where(hr => (string?)hr.Attribute(W.type) == GetReferenceType(type))
            .Remove();
        sectPr.Add(new XElement(W.headerReference,
            new XAttribute(W.type, GetReferenceType(type)),
            new XAttribute(R.id, relId)));

        mainPart.PutXDocument();
    }

    private static void AttachFooter(WordprocessingDocument document, HeaderFooterValues type, params XElement[] paragraphs)
    {
        var mainPart = document.MainDocumentPart ?? throw new InvalidOperationException("Document missing main part.");
        var footerPart = mainPart.AddNewPart<FooterPart>();
        var footerDoc = CreatePartDocument(W.footer, paragraphs);
        footerPart.PutXDocument(footerDoc);

        var relId = mainPart.GetIdOfPart(footerPart);
        var sectPr = GetOrCreateSectionProperties(mainPart);
        sectPr.Elements(W.footerReference)
            .Where(fr => (string?)fr.Attribute(W.type) == GetReferenceType(type))
            .Remove();
        sectPr.Add(new XElement(W.footerReference,
            new XAttribute(W.type, GetReferenceType(type)),
            new XAttribute(R.id, relId)));

        mainPart.PutXDocument();
    }

    private static XElement GetOrCreateSectionProperties(OpenXmlPart part)
    {
        var xDoc = part.GetXDocument();
        var body = xDoc.Root?.Element(W.body) ?? throw new InvalidOperationException("Document missing body.");
        var sectPr = body.Elements(W.sectPr).FirstOrDefault();
        if (sectPr == null)
        {
            sectPr = new XElement(W.sectPr);
            body.Add(sectPr);
        }

        if (!sectPr.Elements(W.pgSz).Any())
        {
            sectPr.Add(new XElement(W.pgSz));
        }

        if (!sectPr.Elements(W.pgMar).Any())
        {
            sectPr.Add(new XElement(W.pgMar));
        }

        return sectPr;
    }

    private static XDocument CreatePartDocument(XName rootName, params XElement[] paragraphs)
    {
        var content = (paragraphs ?? Array.Empty<XElement>())
            .Select(p => new XElement(p));
        return new XDocument(
            new XElement(rootName,
                new XAttribute(XNamespace.Xmlns + "w", W.w),
                content));
    }

    private static string GetReferenceType(HeaderFooterValues type)
    {
        if (type == HeaderFooterValues.Even)
        {
            return "even";
        }

        if (type == HeaderFooterValues.First)
        {
            return "first";
        }

        return "default";
    }
}
