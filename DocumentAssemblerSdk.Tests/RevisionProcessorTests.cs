using DocumentAssembler.Core;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml.Linq;
using Xunit;

namespace DocumentAssembler.Tests;

    public class RevisionProcessorTests
    {
        private static readonly string TestFilesDirectory = Path.Combine(AppContext.BaseDirectory, "TestFiles");

        private const string Author = "SDK Tester";
        private const string SampleDate = "2024-01-01T00:00:00Z";

        public static IEnumerable<object[]> TrackedRevisionDocuments => new[]
        {
            new object[] { "DA024-TrackedRevisions.docx" },
            new object[] { "DA224-TrackedRevisions.docx" }
        };

        public static IEnumerable<object[]> AllTemplateDocuments =>
            Directory.EnumerateFiles(TestFilesDirectory, "*.docx")
                .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
                .Select(Path.GetFileName)
                .Where(name => !string.IsNullOrEmpty(name))
                .Select(name => new object[] { name! });

        [Fact]
        public void AcceptRevisions_RemovesTrackedMarkupFromAllParts()
        {
        var trackedDocument = TestDocumentFactory.Create("TrackedTemplate.docx", builder =>
        {
            builder.AddBodyElement(CreateTrackedParagraph("Body ", "Inserted ", "Deleted ", " Tail"));
            builder.AddDefaultHeader(CreateTrackedParagraph("Header ", "NewHeader ", "OldHeader ", string.Empty));
            builder.AddDefaultFooter(CreateTrackedParagraph("Footer ", "FooterAdd ", "FooterOld ", string.Empty));
        });

        Assert.True(RevisionProcessor.HasTrackedRevisions(trackedDocument));

        var accepted = RevisionProcessor.AcceptRevisions(trackedDocument);
        Assert.False(RevisionProcessor.HasTrackedRevisions(accepted));

        var bodyText = ReadBodyText(accepted);
        Assert.Contains("Body Inserted  Tail", bodyText, StringComparison.Ordinal);
        Assert.DoesNotContain("Deleted", bodyText, StringComparison.Ordinal);

        var allText = ReadAllText(accepted);
        Assert.Contains("NewHeader", allText, StringComparison.Ordinal);
        Assert.DoesNotContain("OldHeader", allText, StringComparison.Ordinal);
        Assert.Contains("FooterAdd", allText, StringComparison.Ordinal);
        Assert.DoesNotContain("FooterOld", allText, StringComparison.Ordinal);
    }

    [Fact]
    public void RejectRevisions_RestoresOriginalBodyContent()
    {
        var trackedDocument = TestDocumentFactory.Create("RejectTemplate.docx", builder =>
        {
            builder.AddBodyElement(CreateTrackedParagraph("Prelude ", "Inserted ", "Removed ", " Ending"));
        });

        var rejected = RevisionProcessor.RejectRevisions(trackedDocument);
        var bodyText = ReadBodyText(rejected);

        Assert.DoesNotContain("Inserted", bodyText, StringComparison.Ordinal);
        Assert.Contains("Removed", bodyText, StringComparison.Ordinal);
        Assert.Contains("Prelude", bodyText, StringComparison.Ordinal);
    }

        [Fact]
        public void AcceptRevisionsForElement_StripsMetadataAndEmptyNodes()
        {
        var paragraph = new XElement(W.p,
            new XAttribute(RevisionProcessor.PT.UniqueId, Guid.NewGuid().ToString("N")),
            new XElement(W.numPr),
            new XElement(W.del,
                new XAttribute(W.author, Author),
                new XElement(W.r, new XElement(W.t, "Deleted text"))),
            new XElement(W.ins,
                new XAttribute(W.author, Author),
                new XAttribute(W.date, SampleDate),
                new XElement(W.r,
                    new XAttribute(RevisionProcessor.PT.RunIds, "1"),
                    new XElement(W.t, "Accepted text")))
        );

        var accepted = RevisionProcessor.AcceptRevisionsForElement(paragraph);
        Assert.NotNull(accepted);
        Assert.Empty(accepted!.Descendants(W.del));
        Assert.Empty(accepted.Descendants(W.numPr));
        Assert.DoesNotContain(accepted.Descendants().Attributes(),
            attr => attr.Name == RevisionProcessor.PT.UniqueId || attr.Name == RevisionProcessor.PT.RunIds);
        Assert.Equal("Accepted text", string.Concat(accepted.Descendants(W.t).Select(t => (string)t)));
        }

        [Theory]
        [MemberData(nameof(TrackedRevisionDocuments))]
        public void AcceptRevisions_RemovesTrackedMarkupFromSample(string fileName)
        {
            var document = LoadTestDocument(fileName);
            var accepted = RevisionProcessor.AcceptRevisions(document);
            Assert.False(RevisionProcessor.HasTrackedRevisions(accepted));
        }

        [Theory]
        [MemberData(nameof(TrackedRevisionDocuments))]
        public void RejectRevisions_RemovesTrackedMarkupFromSample(string fileName)
        {
            var document = LoadTestDocument(fileName);
            var rejected = RevisionProcessor.RejectRevisions(document);
            Assert.False(RevisionProcessor.HasTrackedRevisions(rejected));
        }

        [Theory]
        [MemberData(nameof(AllTemplateDocuments))]
        public void AcceptRevisions_HandleSampleTemplates(string fileName)
        {
            var document = LoadTestDocument(fileName);
            var accepted = RevisionProcessor.AcceptRevisions(document);
            Assert.NotNull(accepted);
        }

        [Theory]
        [MemberData(nameof(AllTemplateDocuments))]
        public void RejectRevisions_HandleSampleTemplates(string fileName)
        {
            var document = LoadTestDocument(fileName);
            var rejected = RevisionProcessor.RejectRevisions(document);
            Assert.NotNull(rejected);
        }

        [Fact]
        public void AcceptRevisions_RemovesStyleChangeNodes()
        {
            var document = TestDocumentFactory.Create("StylesWithChanges.docx", builder =>
            {
                builder.Configure(wordDoc =>
                {
                    var stylesPart = wordDoc.MainDocumentPart!.StyleDefinitionsPart ??
                                     wordDoc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
                    var styles = new XDocument(
                        new XElement(W.styles,
                            new XAttribute(XNamespace.Xmlns + "w", W.w),
                            new XElement(W.style,
                                new XAttribute(W.type, "paragraph"),
                                new XAttribute(W.styleId, "CustomStyle"),
                                new XElement(W.name, new XAttribute(W.val, "CustomStyle")),
                                new XElement(W.pPrChange,
                                    new XAttribute(W.author, Author),
                                    new XAttribute(W.date, SampleDate)),
                                new XElement(W.rPrChange,
                                    new XAttribute(W.author, Author),
                                    new XAttribute(W.date, SampleDate)))));
                    stylesPart.PutXDocument(styles);
                });
            });

            var accepted = RevisionProcessor.AcceptRevisions(document);
            using var streamDoc = new OpenXmlMemoryStreamDocument(accepted);
            using var wDoc = streamDoc.GetWordprocessingDocument();
            var stylesDoc = wDoc.MainDocumentPart!.StyleDefinitionsPart!.GetXDocument();
            Assert.False(stylesDoc.Descendants(W.pPrChange).Any());
            Assert.False(stylesDoc.Descendants(W.rPrChange).Any());
        }

        [Fact]
        public void AcceptRevisions_AddsParagraphsToEmptyCells()
        {
            var document = TestDocumentFactory.Create("EmptyCells.docx", builder =>
            {
                builder.AddBodyElement(new XElement(W.tbl,
                    new XElement(W.tr,
                        new XElement(W.tc,
                            new XElement(W.tcPr)))));
            });

            var accepted = RevisionProcessor.AcceptRevisions(document);
            using var streamDoc = new OpenXmlMemoryStreamDocument(accepted);
            using var wDoc = streamDoc.GetWordprocessingDocument();
            var cell = wDoc.MainDocumentPart!.GetXDocument().Descendants(W.tc).First();
            Assert.True(cell.Elements(W.p).Any());
        }

        [Fact]
        public void AcceptRevisions_RemovesMoveFromAndMoveTo()
        {
            var document = TestDocumentFactory.Create("MoveRanges.docx", builder =>
            {
                builder.AddBodyElement(new XElement(W.p,
                    new XElement(W.moveFromRangeStart,
                        new XAttribute(W.id, 1)),
                    new XElement(W.moveFrom,
                        new XAttribute(W.author, Author),
                        new XAttribute(W.id, 1),
                        new XElement(W.r, new XElement(W.t, "From Text"))),
                    new XElement(W.moveFromRangeEnd,
                        new XAttribute(W.id, 1)),
                    new XElement(W.moveToRangeStart,
                        new XAttribute(W.id, 1)),
                    new XElement(W.moveTo,
                        new XAttribute(W.author, Author),
                        new XAttribute(W.id, 1),
                        new XElement(W.r, new XElement(W.t, "To Text"))),
                    new XElement(W.moveToRangeEnd,
                        new XAttribute(W.id, 1))));
            });

            var accepted = RevisionProcessor.AcceptRevisions(document);
            using var streamDoc = new OpenXmlMemoryStreamDocument(accepted);
            using var wDoc = streamDoc.GetWordprocessingDocument();
            var main = wDoc.MainDocumentPart!.GetXDocument();
            Assert.Empty(main.Descendants(W.moveFrom));
            Assert.Empty(main.Descendants(W.moveTo));
            Assert.Contains("To Text", main.Descendants(W.t).Select(t => (string)t));
        }

        [Fact]
        public void AcceptParagraphEndTagsInMoveFromTransform_CollapsesParagraphs()
        {
            var body = new XElement(W.body,
                new XElement(W.p,
                    new XElement(W.moveFromRangeStart, new XAttribute(W.id, 1)),
                    new XElement(W.r, new XElement(W.t, "From"))),
                new XElement(W.p,
                    new XElement(W.moveFromRangeEnd, new XAttribute(W.id, 1)),
                    new XElement(W.r, new XElement(W.t, "After"))));

            var result = InvokePrivateTransform<XElement>("AcceptParagraphEndTagsInMoveFromTransform", body);
            Assert.Equal(2, result.Elements(W.p).Count());
            Assert.Contains("After", string.Concat(result.Descendants(W.t).Select(t => (string)t)));
        }

        [Fact]
        public void TransformInstrTextToDelInstrText_ReplacesElement()
        {
            var instr = new XElement(W.instrText, "MERGEFIELD Sample");
            var result = InvokePrivateTransform<XElement>("TransformInstrTextToDelInstrText", instr);
            Assert.Equal(W.delInstrText, result.Name);
            Assert.Equal("MERGEFIELD Sample", result.Value);
        }

        [Fact]
        public void CoalesqueParagraphEndTagsInMoveFromTransform_MergesGroups()
        {
            var paragraph = new XElement(W.p, new XElement(W.r, new XElement(W.t, "Base")));
            var mergeParagraph = new XElement(W.p,
                new XElement(W.r, new XElement(W.t, "Merged")));

            var enumType = typeof(RevisionProcessor).GetNestedType("MoveFromCollectionType", BindingFlags.NonPublic)!;
            var groupingType = typeof(TestGrouping<,>).MakeGenericType(enumType, typeof(XElement));
            var key = Enum.ToObject(enumType, 0);
            var grouping = Activator.CreateInstance(groupingType, key, new List<XElement> { paragraph, mergeParagraph });

            var result = InvokePrivateTransform<XElement>("CoalesqueParagraphEndTagsInMoveFromTransform", paragraph, grouping);
            Assert.Contains("Merged", string.Concat(result.Descendants(W.t).Select(t => (string)t)));
        }

        [Fact]
        public void AcceptDeletedCellsTransform_MergesDeletedCellsIntoGridSpan()
        {
            var row = new XElement(W.tr,
                new XElement(W.tc,
                    new XElement(W.tcPr),
                    new XElement(W.p, new XElement(W.r, new XElement(W.t, "Keep")))),
                new XElement(W.tc,
                    new XElement(W.tcPr),
                    new XElement(W.cellDel)));

            var result = InvokePrivateTransform<XElement>("AcceptDeletedCellsTransform", row);
            Assert.Single(result.Elements(W.tc));
            Assert.Contains(result.Descendants(W.gridSpan), span => (int?)span.Attribute(W.val) == 2);
        }

        [Fact]
        public void AcceptDeletedAndMovedFromContentControlsTransform_RemovesTargets()
        {
            var contentControl = new XElement(W.sdt,
                new XElement(W.sdtContent,
                    new XElement(W.p,
                        new XElement(W.r, new XElement(W.t, "Current")))));
            var moveFrom = new XElement(W.moveFrom, new XAttribute(W.id, 1));
            var doc = new XElement(W.body, contentControl, moveFrom);

            var result = InvokePrivateTransform<XElement>(
                "AcceptDeletedAndMovedFromContentControlsTransform",
                doc,
                new[] { contentControl },
                new[] { moveFrom });
            Assert.DoesNotContain(result.Descendants(W.sdt), _ => true);
            Assert.DoesNotContain(result.Descendants(W.moveFrom), _ => true);
            Assert.Contains("Current", string.Concat(result.Descendants(W.t).Select(t => (string)t)));
        }

        [Fact]
        public void AcceptMoveFromRangesTransform_RemovesSpecifiedElements()
        {
            var moveFrom = new XElement(W.moveFrom, new XAttribute(W.id, 5),
                new XElement(W.r, new XElement(W.t, "old")));
            var doc = new XElement(W.body,
                moveFrom,
                new XElement(W.p, new XElement(W.r, new XElement(W.t, "keep"))));

            var result = InvokePrivateTransform<XElement>("AcceptMoveFromRangesTransform", doc, new[] { moveFrom });
            Assert.DoesNotContain(result.Descendants(W.moveFrom), _ => true);
            Assert.Contains("keep", string.Concat(result.Descendants(W.t).Select(t => (string)t)));
        }

        [Fact]
        public void RemoveRowsLeftEmptyByMoveFrom_DropsEmptyRows()
        {
            var emptyRow = new XElement(W.tr,
                new XElement(W.tc, new XElement(W.tcPr)));
            var contentRow = new XElement(W.tr,
                new XElement(W.tc,
                    new XElement(W.p, new XElement(W.r, new XElement(W.t, "keep")))));
            var table = new XElement(W.tbl, emptyRow, contentRow);

            var result = InvokePrivateTransform<XElement>("RemoveRowsLeftEmptyByMoveFrom", table);
            Assert.Single(result.Elements(W.tr));
            Assert.Contains("keep", string.Concat(result.Descendants(W.t).Select(t => (string)t)));
        }

        [Fact]
        public void AcceptDeletedAndMoveFromParagraphMarksTransform_RemovesDeletedParagraphs()
        {
            var paragraph = new XElement(W.p,
                new XElement(W.del,
                    new XElement(W.r, new XElement(W.t, "deleted"))),
                new XElement(W.r, new XElement(W.t, "keep")));
            var doc = new XElement(W.body, paragraph);

            var result = InvokePrivateTransform<XElement>("AcceptDeletedAndMoveFromParagraphMarksTransform", doc);
            Assert.Contains("keep", string.Concat(result.Descendants(W.t).Select(t => (string)t)));
        }

        [Fact]
        public void AcceptRevisions_HandlesComplexDocument()
        {
            var document = TestDocumentFactory.Create("Complex.docx", builder =>
            {
                builder.AddBodyElement(new XElement(W.p,
                    new XElement(W.r,
                        new XElement(W.fldChar, new XAttribute(W.fldCharType, "begin"))),
                    new XElement(W.del,
                        new XElement(W.r,
                            new XElement(W.fldChar, new XAttribute(W.fldCharType, "separate")))),
                    new XElement(W.r, new XElement(W.instrText, "MERGEFIELD SampleField")),
                    new XElement(W.ins,
                        new XElement(W.r,
                            new XElement(W.fldChar, new XAttribute(W.fldCharType, "end"))))));

                builder.AddBodyElement(new XElement(W.tbl,
                    new XElement(W.tr,
                        new XElement(W.tc,
                            new XElement(W.tcPr,
                                new XElement(W.gridSpan, new XAttribute(W.val, 1))),
                            new XElement(W.p, new XElement(W.r, new XElement(W.t, "KeepCell")))),
                        new XElement(W.tc,
                            new XElement(W.tcPr),
                            new XElement(W.cellDel)),
                        new XElement(W.tc,
                            new XElement(W.tcPr),
                            new XElement(W.cellDel)))));

                builder.AddBodyElement(new XElement(W.p,
                    new XElement(W.pPr, new XElement(W.rPr, new XElement(W.del))),
                    new XElement(W.r, new XElement(W.t, "Deleted paragraph"))));

                builder.AddBodyElement(new XElement(W.p,
                    new XElement(W.r, new XElement(W.t, "Following paragraph"))));

                builder.AddBodyElement(new XElement(W.sdt,
                    new XElement(W.sdtContent,
                        new XElement(W.p,
                            new XElement(W.moveFromRangeStart, new XAttribute(W.id, 5)),
                            new XElement(W.moveFrom,
                                new XAttribute(W.id, 5),
                                new XElement(W.r, new XElement(W.t, "Moved from sdt"))),
                            new XElement(W.moveTo,
                                new XAttribute(W.id, 5),
                                new XElement(W.r, new XElement(W.t, "Moved to sdt"))),
                            new XElement(W.moveFromRangeEnd, new XAttribute(W.id, 5))))));
            });

            var accepted = RevisionProcessor.AcceptRevisions(document);
            Assert.False(RevisionProcessor.HasTrackedRevisions(accepted));
            Assert.Contains("Following paragraph", ReadBodyText(accepted));
        }

        private static XElement CreateTrackedParagraph(string prefix, string inserted, string deleted, string suffix)
        {
            return new XElement(W.p,
                new XElement(W.r, new XElement(W.t, prefix)),
                new XElement(W.ins,
                    new XAttribute(W.id, 1),
                    new XAttribute(W.author, Author),
                    new XAttribute(W.date, SampleDate),
                    new XElement(W.r, new XElement(W.t, inserted))),
                new XElement(W.del,
                    new XAttribute(W.id, 2),
                    new XAttribute(W.author, Author),
                    new XAttribute(W.date, SampleDate),
                    new XElement(W.r, new XElement(W.t, deleted))),
                new XElement(W.r, new XElement(W.t, suffix)));
        }

    private static string ReadBodyText(WmlDocument document)
    {
        using var ms = new MemoryStream(document.DocumentByteArray);
        using var wordDoc = WordprocessingDocument.Open(ms, false);
        var main = wordDoc.MainDocumentPart?.GetXDocument();
        if (main?.Root == null)
        {
            return string.Empty;
        }

        return string.Concat(main.Descendants(W.t).Select(t => (string)t));
    }

    private static string ReadAllText(WmlDocument document)
    {
        using var ms = new MemoryStream(document.DocumentByteArray);
        using var wordDoc = WordprocessingDocument.Open(ms, false);
        var builder = new StringBuilder();

        if (wordDoc.MainDocumentPart != null)
        {
            builder.Append(ExtractText(wordDoc.MainDocumentPart));
        }

        foreach (var header in wordDoc.MainDocumentPart?.HeaderParts ?? Enumerable.Empty<HeaderPart>())
        {
            builder.Append(ExtractText(header));
        }

        foreach (var footer in wordDoc.MainDocumentPart?.FooterParts ?? Enumerable.Empty<FooterPart>())
        {
            builder.Append(ExtractText(footer));
        }

        return builder.ToString();
    }

    private static string ExtractText(OpenXmlPart part)
    {
        var xDoc = part.GetXDocument();
        return xDoc.Root == null
            ? string.Empty
            : string.Concat(xDoc.Descendants(W.t).Select(t => (string)t));
    }

    private static WmlDocument LoadTestDocument(string fileName)
    {
        var path = Path.Combine(TestFilesDirectory, fileName);
        return new WmlDocument(path);
    }

    private static T InvokePrivateTransform<T>(string methodName, params object?[] parameters)
    {
        var method = typeof(RevisionProcessor).GetMethod(methodName,
            BindingFlags.Static | BindingFlags.NonPublic);
        Assert.NotNull(method);
        var result = method!.Invoke(null, parameters);
        return (T)result!;
    }

    private sealed class TestGrouping<TKey, TElement> : IGrouping<TKey, TElement>
    {
        private readonly List<TElement> _items;
        public TKey Key { get; }

        public TestGrouping(TKey key, IEnumerable<TElement> items)
        {
            Key = key;
            _items = items.ToList();
        }

        public IEnumerator<TElement> GetEnumerator() => _items.GetEnumerator();

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
