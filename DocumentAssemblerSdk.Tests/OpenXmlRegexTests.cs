using DocumentAssembler.Core;
using DocumentAssembler.Core.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Xunit;

namespace DocumentAssembler.Tests;

public class OpenXmlRegexTests
{
    [Fact]
    public void MatchAndReplace_ShouldUpdateParagraphContent()
    {
        var paragraph = CreateParagraph("Order ", "42", " shipped");
        var regex = new Regex("42", RegexOptions.Compiled);

        var matches = OpenXmlRegex.Match(new[] { paragraph }, regex);
        Assert.Equal(1, matches);

        var callbackHits = 0;
        OpenXmlRegex.Match(new[] { paragraph }, regex, (p, match) =>
        {
            callbackHits++;
            Assert.Same(paragraph, p);
            Assert.Equal("42", match.Value);
        });
        Assert.Equal(1, callbackHits);

        var replacements = OpenXmlRegex.Replace(new[] { paragraph }, regex, "84", (_, __) => true);
        Assert.Equal(1, replacements);
        Assert.Equal("Order 84 shipped", GetParagraphText(paragraph));
    }

    [Fact]
    public void Replace_WithTrackRevisions_ShouldProduceInsAndDelRuns()
    {
        var paragraph = CreateParagraph("Alpha ", "Beta", " Gamma");
        var regex = new Regex("Beta");

        var replacements = OpenXmlRegex.Replace(new[] { paragraph }, regex, "Delta", (_, __) => true, true, "Tester");
        Assert.Equal(1, replacements);

        var inserted = paragraph.Descendants(W.ins).Single();
        Assert.Equal("Delta", string.Concat(inserted.Descendants(W.t).Select(t => (string)t)));

        var deleted = paragraph.Descendants(W.del).Single();
        Assert.Equal("Beta", string.Concat(deleted.Descendants(W.delText).Select(t => (string)t)));

        var ids = paragraph.Descendants()
            .Where(e => e.Name == W.ins || e.Name == W.del)
            .Select(e => e.Attribute(W.id))
            .Where(a => a != null)
            .Select(a => a!.Value)
            .ToList();
        Assert.NotEmpty(ids);
        Assert.All(ids, id => Assert.False(string.IsNullOrEmpty(id)));
    }

    [Fact]
    public void Replace_WithCallbackCanSkipMatches()
    {
        var paragraph = CreateParagraph("Alpha ", "Alpha ", "Alpha");
        var regex = new Regex("Alpha");

        var hits = 0;
        var replacements = OpenXmlRegex.Replace(new[] { paragraph }, regex, "Beta", (_, __) => hits++ == 0);
        Assert.Equal(3, replacements);
        Assert.Equal("Beta Alpha Alpha", GetParagraphText(paragraph));
    }

    [Fact]
    public void Replace_WithCoalesceDisabled_RetainsSplitRuns()
    {
        var baseline = CreateParagraph("Number ", "42");
        var coalesced = new XElement(baseline);
        var regex = new Regex("42");

        OpenXmlRegex.Replace(new[] { baseline }, regex, "43", (_, __) => true, coalesceContent: false);
        OpenXmlRegex.Replace(new[] { coalesced }, regex, "43", (_, __) => true);

        Assert.True(baseline.Descendants(W.r).Count() >= coalesced.Descendants(W.r).Count());
        Assert.Equal(GetParagraphText(coalesced), GetParagraphText(baseline));
    }

    [Fact]
    public void Replace_InPptxWithTracking_Throws()
    {
        var pptParagraph = new XElement(P.p + "p",
            new XElement(A.r,
                new XElement(A.t, "Slide content")));
        var regex = new Regex("Slide");

        Assert.Throws<OpenXmlPowerToolsException>(() =>
            OpenXmlRegex.Replace(new[] { pptParagraph }, regex, "Deck", (_, __) => true, true, "Tester"));
    }

    [Fact]
    public void Replace_ShouldThrowWhenContentIsNull()
    {
        Assert.Throws<ArgumentNullException>(() =>
            OpenXmlRegex.Replace(null!, new Regex("Value"), "V", (_, __) => true));
    }

    [Fact]
    public void Replace_ShouldThrowWhenRegexIsNull()
    {
        var paragraph = CreateParagraph("Alpha");
        Assert.Throws<ArgumentNullException>(() =>
            OpenXmlRegex.Replace(new[] { paragraph }, null!, "Beta", (_, __) => true));
    }

    [Fact]
    public void Replace_WithEmptyContent_ReturnsZero()
    {
        var result = OpenXmlRegex.Replace(Array.Empty<XElement>(), new Regex("Alpha"), "Beta", (_, __) => true);
        Assert.Equal(0, result);
    }

    [Fact]
    public void Replace_WithUnknownNamespace_ReturnsZero()
    {
        var custom = new XElement(XName.Get("root", "urn:test"));
        var result = OpenXmlRegex.Replace(new[] { custom }, new Regex("Alpha"), "Beta", (_, __) => true);
        Assert.Equal(0, result);
    }

    [Fact]
    public void Replace_AssignsMissingAndDuplicateRevisionIds()
    {
        var (document, paragraph) = CreateDocumentParagraph(
            new XElement(W.ins,
                new XAttribute(W.author, "A"),
                new XAttribute(W.id, 5),
                new XElement(W.r, new XElement(W.t, "Keep 1"))),
            new XElement(W.del,
                new XAttribute(W.author, "B"),
                new XElement(W.r, new XElement(W.t, "Keep 2"))),
            new XElement(W.ins,
                new XAttribute(W.author, "C"),
                new XAttribute(W.id, 5),
                new XElement(W.r, new XElement(W.t, "Keep 3"))),
            new XElement(W.r, new XElement(W.t, "Token")));

        var regex = new Regex("Token");
        var replaced = OpenXmlRegex.Replace(new[] { paragraph }, regex, "Value", (_, __) => true);
        Assert.Equal(1, replaced);

        var tracked = document
            .Descendants()
            .Where(e => e.Name == W.ins || e.Name == W.del)
            .Select(e => (Element: e, Id: (string?)e.Attribute(W.id)))
            .ToList();
        Assert.All(tracked, t => Assert.False(string.IsNullOrEmpty(t.Id)));
        var distinctCount = tracked.Select(t => t.Id).Distinct().Count();
        Assert.Equal(tracked.Count, distinctCount);
    }

    [Fact]
    public void Replace_SkipsZeroLengthMatches()
    {
        var paragraph = CreateParagraph("Alpha");
        var regex = new Regex("(?=Alpha)");
        var replacements = OpenXmlRegex.Replace(new[] { paragraph }, regex, "Beta", (_, __) => true);
        Assert.True(replacements > 0);
        Assert.Equal("Alpha", GetParagraphText(paragraph));
    }

    [Fact]
    public void Replace_TrackRevisionsRemovesAuthoredInsertions()
    {
        var (document, paragraph) = CreateDocumentParagraph(
            new XElement(W.ins,
                new XAttribute(W.author, "Tester"),
                new XElement(W.r, new XElement(W.t, "Mutable"))));

        var regex = new Regex("Mutable");
        var replacements = OpenXmlRegex.Replace(new[] { paragraph }, regex, "Replaced", (_, __) => true, true, "Tester");
        Assert.Equal(1, replacements);
        Assert.Single(document.Descendants(W.ins));
        Assert.DoesNotContain("Mutable", document.Descendants(W.t).Select(t => (string)t));
        Assert.Contains("Replaced", document.Descendants(W.t).Select(t => (string)t));
    }

    [Fact]
    public void Replace_TrackRevisionsWrapsForeignInsertions()
    {
        var paragraph = new XElement(W.p,
            new XElement(W.ins,
                new XAttribute(W.author, "Other"),
                new XElement(W.r, new XElement(W.t, "Mutable"))));

        var regex = new Regex("Mutable");
        var replacements = OpenXmlRegex.Replace(new[] { paragraph }, regex, "Replaced", (_, __) => true, true, "Tester");
        Assert.Equal(1, replacements);

        var deleted = paragraph.Descendants(W.del).First();
        Assert.NotEmpty(deleted.Descendants(W.delText));
        Assert.Contains("Replaced", paragraph.Descendants(W.ins).SelectMany(i => i.Descendants(W.t)).Select(t => (string)t));
    }

    [Fact]
    public void Replace_TrackRevisionsWithEmptyReplacementOnlyDeletes()
    {
        var paragraph = CreateParagraph("Alpha");
        var regex = new Regex("Alpha");
        var replacements = OpenXmlRegex.Replace(new[] { paragraph }, regex, string.Empty, (_, __) => true, true, "Tester");
        Assert.Equal(1, replacements);
        Assert.Contains(paragraph.Descendants(W.del),
            d => string.Concat(d.Descendants(W.delText).Select(t => (string)t)) == "Alpha");
        Assert.DoesNotContain(paragraph.Descendants(W.ins), _ => true);
    }

    [Fact]
    public void Replace_PmlParagraph_ReplacesText()
    {
        var slideParagraph = new XElement(A.p,
            new XElement(A.r,
                new XElement(A.t, "Slide content")));
        var regex = new Regex("content");
        var replacements = OpenXmlRegex.Replace(new[] { slideParagraph }, regex, "deck", (_, __) => true);
        Assert.Equal(1, replacements);
        Assert.Contains(slideParagraph.Descendants(A.t).Select(t => (string)t), text => text.Contains("deck", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Replace_NonTrackingHandlesRunsInsideInsertions()
    {
        var paragraph = new XElement(W.p,
            new XElement(W.ins,
                new XElement(W.r, new XElement(W.t, "To")),
                new XElement(W.r, new XElement(W.t, "ken")),
                new XElement(W.r, new XElement(W.t, "Suffix"))));
        var regex = new Regex("Token");
        var replacements = OpenXmlRegex.Replace(new[] { paragraph }, regex, "Done", (_, __) => true);
        Assert.Equal(1, replacements);
        Assert.Contains(paragraph.Descendants(W.t).Select(t => (string)t), text => text.Contains("Done", StringComparison.Ordinal));
        Assert.DoesNotContain(paragraph.Descendants(W.t), t => (string)t == "Token");
    }

    [Fact]
    public void Replace_WhenPatternMissing_PreservesParagraphStructure()
    {
        var paragraph = new XElement(W.p,
            new XElement(W.pPr, new XElement(W.rPr)),
            new XElement(W.r, new XElement(W.t, "Value")),
            new XElement(W.r, new XElement(W.tab)),
            new XElement(W.ins,
                new XElement(W.r, new XElement(W.t, "Nested"))));

        var regex = new Regex("Missing");
        var replacements = OpenXmlRegex.Replace(new[] { paragraph }, regex, "Other", (_, __) => true);
        Assert.Equal(0, replacements);
        Assert.Equal(2, paragraph.Elements(W.r).Count());
        Assert.Single(paragraph.Elements(W.ins));
    }

    [Fact]
    public void Replace_OnInsertionElement_ReplacesContent()
    {
        var paragraph = new XElement(W.p,
            new XElement(W.ins,
                new XAttribute(W.author, "Tester"),
                new XElement(W.r, new XElement(W.t, "Token"))));

        var regex = new Regex("Token");
        var replacements = OpenXmlRegex.Replace(new[] { paragraph }, regex, "Value", (_, __) => true);
        Assert.Equal(1, replacements);
        Assert.Contains("Value", paragraph.Descendants(W.t).Select(t => (string)t));
    }

    [Fact]
    public void Replace_OnRunElement_ReplacesContent()
    {
        var paragraph = new XElement(W.p,
            new XElement(W.r,
                new XElement(W.rPr, new XElement(W.b)),
                new XElement(W.t, "Token"),
                new XElement(W.tab)));

        var regex = new Regex("Token");
        var replacements = OpenXmlRegex.Replace(new[] { paragraph }, regex, "Value", (_, __) => true);
        Assert.Equal(1, replacements);
        Assert.Contains("Value", paragraph.Descendants(W.t).Select(t => (string)t));
    }

    private static XElement CreateParagraph(params string[] fragments)
    {
        IEnumerable<XElement> runs = fragments.Select(fragment =>
            new XElement(W.r, new XElement(W.t, fragment)));
        return new XElement(W.p, runs);
    }

    private static string GetParagraphText(XElement paragraph)
    {
        return string.Concat(paragraph.Descendants(W.t).Select(t => (string)t));
    }

    private static (XElement Document, XElement Paragraph) CreateDocumentParagraph(params object[] nodes)
    {
        var paragraph = new XElement(W.p, nodes);
        var doc = new XElement(W.document,
            new XAttribute(XNamespace.Xmlns + "w", W.w),
            new XElement(W.body, paragraph));
        return (doc, paragraph);
    }
}
