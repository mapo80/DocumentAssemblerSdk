using DocumentAssembler.Core;
using System;
using System.IO;
using System.Linq;
using Xunit;

namespace DocumentAssembler.Tests;

public class RevisionAccepterTests
{
    [Fact]
    public void AcceptRevisions_RemovesTrackedMarks()
    {
        var source = LoadTrackedDocument("DA024-TrackedRevisions.docx");
        Assert.True(RevisionAccepter.HasTrackedRevisions(source));

        var accepted = RevisionAccepter.AcceptRevisions(source);
        Assert.False(RevisionAccepter.HasTrackedRevisions(accepted));
    }

    [Fact]
    public void RejectRevisions_RemovesTrackedMarks()
    {
        var source = LoadTrackedDocument("DA224-TrackedRevisions.docx");
        Assert.True(RevisionAccepter.HasTrackedRevisions(source));

        var rejected = RevisionProcessor.RejectRevisions(source);
        Assert.False(RevisionAccepter.HasTrackedRevisions(rejected));
    }

    [Fact]
    public void PartHasTrackedRevisions_ScansParts()
    {
        var source = LoadTrackedDocument("DA024-TrackedRevisions.docx");
        using var msDoc = new OpenXmlMemoryStreamDocument(source);
        using var wordDoc = msDoc.GetWordprocessingDocument();
        var hasTracked = wordDoc.MainDocumentPart!.HeaderParts.Any(part => RevisionAccepter.PartHasTrackedRevisions(part));
        Assert.False(hasTracked); // header parts in sample have no revisions
        Assert.True(RevisionAccepter.PartHasTrackedRevisions(wordDoc.MainDocumentPart));
    }

    private static WmlDocument LoadTrackedDocument(string fileName)
    {
        var path = Path.Combine(AppContext.BaseDirectory, "TestFiles", fileName);
        return new WmlDocument(path);
    }
}
