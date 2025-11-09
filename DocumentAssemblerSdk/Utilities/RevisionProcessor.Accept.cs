using DocumentAssembler.Core.Exceptions;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace DocumentAssembler.Core
{
    public partial class RevisionProcessor
    {
        private static void AcceptRevisionsForStylesDefinitionPart(StyleDefinitionsPart stylesDefinitionsPart)
        {
            var xDoc = stylesDefinitionsPart.GetXDocument();
            var newRoot = AcceptRevisionsForStylesTransform(xDoc.Root);
            xDoc.Root.ReplaceWith(newRoot);
            stylesDefinitionsPart.PutXDocument();
        }

        private static object? AcceptRevisionsForStylesTransform(XNode node)
        {
            if (node is XElement element)
            {
                if (element.Name == W.pPrChange || element.Name == W.rPrChange)
                {
                    return null;
                }

                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => AcceptRevisionsForStylesTransform(n)));
            }
            return node;
        }

        public static void AcceptRevisionsForPart(OpenXmlPart part)
        {
            var documentElement = part.GetXDocument().Root;
            documentElement = RemoveRsidTransform(documentElement) as XElement;
            documentElement = (XElement)FixUpDeletedOrInsertedFieldCodesTransform(documentElement);
            var containsMoveFromMoveTo = documentElement.Descendants(W.moveFrom).Any();
            documentElement = AcceptMoveFromMoveToTransform(documentElement) as XElement;
            documentElement = AcceptMoveFromRanges(documentElement);
            // AcceptParagraphEndTagsInMoveFromTransform needs rewritten similar to AcceptDeletedAndMoveFromParagraphMarks
            documentElement = (XElement)AcceptParagraphEndTagsInMoveFromTransform(documentElement);
            documentElement = AcceptDeletedAndMovedFromContentControls(documentElement);
            documentElement = AcceptDeletedAndMoveFromParagraphMarks(documentElement);
            if (containsMoveFromMoveTo)
            {
                documentElement = RemoveRowsLeftEmptyByMoveFrom(documentElement) as XElement;
            }

            documentElement = AcceptAllOtherRevisionsTransform(documentElement) as XElement;
            documentElement = (XElement)AcceptDeletedCellsTransform(documentElement);
            documentElement = (XElement)MergeAdjacentTablesTransform(documentElement);
            documentElement = (XElement)AddEmptyParagraphToAnyEmptyCells(documentElement);
            documentElement.Descendants().Attributes().Where(a => a.Name == PT.UniqueId || a.Name == PT.RunIds).Remove();
            documentElement.Descendants(W.numPr).Where(np => !np.HasElements).Remove();
            var newXDoc = new XDocument(documentElement);
            part.PutXDocument(newXDoc);
        }

        private static object FixUpDeletedOrInsertedFieldCodesTransform(XNode node)
        {
            if (node is XElement element)
            {
                if (element.Name == W.p)
                {
                    // 1 other
                    // 2 w:del/w:r/w:fldChar
                    // 3 w:ins/w:r/w:fldChar
                    // 4 w:instrText

                    // formulate new paragraph, looking for 4 that has 2 (or 3) before and after.  Then put in a w:del (or w:ins), transforming w:instrText to w:delInstrText if w:del.
                    // transform 1, 2, 3 as usual

                    var groupedParaContentsKey = element.Elements().Select(e =>
                    {
                        if (e.Name == W.del && e.Elements(W.r).Elements(W.fldChar).Any())
                        {
                            return 2;
                        }

                        if (e.Name == W.ins && e.Elements(W.r).Elements(W.fldChar).Any())
                        {
                            return 3;
                        }

                        if (e.Name == W.r && e.Element(W.instrText) != null)
                        {
                            return 4;
                        }

                        return 1;
                    });

                    var zipped = element.Elements().Zip(groupedParaContentsKey, (e, k) => new { Ele = e, Key = k });

                    var grouped = zipped.GroupAdjacent(z => z.Key).ToArray();

                    var gLen = grouped.Length;

                    var newParaContents = grouped
                        .Select((g, i) =>
                        {
                            if (g.Key == 1 || g.Key == 2 || g.Key == 3)
                            {
                                return (object)g.Select(gc => FixUpDeletedOrInsertedFieldCodesTransform(gc.Ele));
                            }

                            if (g.Key == 4)
                            {
                                if (i == 0 || i == gLen - 1)
                                {
                                    return g.Select(gc => FixUpDeletedOrInsertedFieldCodesTransform(gc.Ele));
                                }

                                if (grouped[i - 1].Key == 2 &&
                                    grouped[i + 1].Key == 2)
                                {
                                    return new XElement(W.del,
                                        g.Select(gc => TransformInstrTextToDelInstrText(gc.Ele)));
                                }
                                else if (grouped[i - 1].Key == 3 &&
                                    grouped[i + 1].Key == 3)
                                {
                                    return new XElement(W.ins,
                                        g.Select(gc => FixUpDeletedOrInsertedFieldCodesTransform(gc.Ele)));
                                }
                                else
                                {
                                    return g.Select(gc => FixUpDeletedOrInsertedFieldCodesTransform(gc.Ele));
                                }
                            }
                            throw new OpenXmlPowerToolsException("Internal error");
                        });

                    var newParagraph = new XElement(W.p,
                        element.Attributes(),
                        newParaContents);
                    return newParagraph;
                }

                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => FixUpDeletedOrInsertedFieldCodesTransform(n)));
            }
            return node;
        }

        private static object TransformInstrTextToDelInstrText(XNode node)
        {
            if (node is XElement element)
            {
                if (element.Name == W.instrText)
                {
                    return new XElement(W.delInstrText,
                        element.Attributes(),
                        element.Nodes());
                }

                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => TransformInstrTextToDelInstrText(n)));
            }
            return node;
        }

        private static object AddEmptyParagraphToAnyEmptyCells(XNode node)
        {
            if (node is XElement element)
            {
                if (element.Name == W.tc && !element.Elements().Any(e => e.Name != W.tcPr))
                {
                    return new XElement(W.tc,
                        element.Attributes(),
                        element.Elements(),
                        new XElement(W.p));
                }

                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => AddEmptyParagraphToAnyEmptyCells(n)));
            }
            return node;
        }

        private static object? AcceptMoveFromMoveToTransform(XNode? node)
        {
            if (node is XElement element)
            {
                if (element.Name == W.moveTo)
                {
                    return element.Nodes().Select(n => AcceptMoveFromMoveToTransform(n));
                }

                if (element.Name == W.moveFrom)
                {
                    return null;
                }

                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => AcceptMoveFromMoveToTransform(n)));
            }
            return node;
        }

        private static XElement? AcceptMoveFromRanges(XElement? document)
        {
            // The following lists contain the elements that are between start/end elements.
            var startElementTagsInMoveFromRange = new List<XElement>();
            var endElementTagsInMoveFromRange = new List<XElement>();

            // Following are the elements that *may* be in a range that has both start and end
            // elements.
            var potentialDeletedElements =
                new Dictionary<string, PotentialInRangeElements>();

            foreach (var tag in DescendantAndSelfTags(document))
            {
                if (tag.Element.Name == W.moveFromRangeStart)
                {
                    var id = tag.Element.Attribute(W.id).Value;
                    potentialDeletedElements.Add(id, new PotentialInRangeElements());
                    continue;
                }
                if (tag.Element.Name == W.moveFromRangeEnd)
                {
                    var id = tag.Element.Attribute(W.id).Value;
                    if (potentialDeletedElements.ContainsKey(id))
                    {
                        startElementTagsInMoveFromRange.AddRange(
                            potentialDeletedElements[id].PotentialStartElementTagsInRange);
                        endElementTagsInMoveFromRange.AddRange(
                            potentialDeletedElements[id].PotentialEndElementTagsInRange);
                        potentialDeletedElements.Remove(id);
                    }
                    continue;
                }
                if (potentialDeletedElements.Count > 0)
                {
                    if (tag.TagType == TagTypeEnum.Element &&
                        (tag.Element.Name != W.moveFromRangeStart &&
                         tag.Element.Name != W.moveFromRangeEnd))
                    {
                        foreach (var id in potentialDeletedElements)
                        {
                            id.Value.PotentialStartElementTagsInRange.Add(tag.Element);
                        }

                        continue;
                    }
                    if (tag.TagType == TagTypeEnum.EmptyElement &&
                        (tag.Element.Name != W.moveFromRangeStart &&
                         tag.Element.Name != W.moveFromRangeEnd))
                    {
                        foreach (var id in potentialDeletedElements)
                        {
                            id.Value.PotentialStartElementTagsInRange.Add(tag.Element);
                            id.Value.PotentialEndElementTagsInRange.Add(tag.Element);
                        }
                        continue;
                    }
                    if (tag.TagType == TagTypeEnum.EndElement &&
                        (tag.Element.Name != W.moveFromRangeStart &&
                        tag.Element.Name != W.moveFromRangeEnd))
                    {
                        foreach (var id in potentialDeletedElements)
                        {
                            id.Value.PotentialEndElementTagsInRange.Add(tag.Element);
                        }

                        continue;
                    }
                }
            }
            var moveFromElementsToDelete = startElementTagsInMoveFromRange.Intersect(endElementTagsInMoveFromRange).ToArray();

            if (moveFromElementsToDelete.Any())
            {
                return AcceptMoveFromRangesTransform(document, moveFromElementsToDelete) as XElement;
            }

            return document;
        }

        private static object AcceptParagraphEndTagsInMoveFromTransform(XNode node)
        {
            if (node is XElement element)
            {
                if (W.BlockLevelContentContainers.Contains(element.Name))
                {
                    var groupedBodyChildren = element
                        .Elements()
                        .GroupAdjacent(c =>
                        {
                            var pi = c.GetParagraphInfo();
                            if (pi.ThisBlockContentElement != null)
                            {
                                var paragraphMarkIsInMoveFromRange =
                                    pi.ThisBlockContentElement.Elements(W.moveFromRangeStart).Any() &&
                                    !pi.ThisBlockContentElement.Elements(W.moveFromRangeEnd).Any();
                                if (paragraphMarkIsInMoveFromRange)
                                {
                                    return MoveFromCollectionType.ParagraphEndTagInMoveFromRange;
                                }
                            }
                            var previousContentElement = c.ContentElementsBeforeSelf()
                                .FirstOrDefault(e => e.GetParagraphInfo().ThisBlockContentElement != null);
                            if (previousContentElement != null)
                            {
                                var pi2 = previousContentElement.GetParagraphInfo();
                                if (c.Name == W.p &&
                                    pi2.ThisBlockContentElement.Elements(W.moveFromRangeStart).Any() &&
                                    !pi2.ThisBlockContentElement.Elements(W.moveFromRangeEnd).Any())
                                {
                                    return MoveFromCollectionType.ParagraphEndTagInMoveFromRange;
                                }
                            }
                            return MoveFromCollectionType.Other;
                        })
                        .ToList();

                    // If there is only one group, and it's key is MoveFromCollectionType.Other
                    // then there is nothing to do.
                    if (groupedBodyChildren.Count == 1 &&
                        groupedBodyChildren.First().Key == MoveFromCollectionType.Other)
                    {
                        var newElement = new XElement(element.Name,
                            element.Attributes(),
                            groupedBodyChildren.Select(g =>
                            {
                                if (g.Key == MoveFromCollectionType.Other)
                                {
                                    return g;
                                }

                                // This is a transform that produces the first element in the
                                // collection, except that the paragraph in the descendents is
                                // replaced with a new paragraph that contains all contents of the
                                // existing paragraph, plus subsequent elements in the group
                                // collection, where the paragraph in each of those groups is
                                // collapsed.
                                return CoalesqueParagraphEndTagsInMoveFromTransform(g.First(), g);
                            }));
                        return newElement;
                    }
                    else
                    {
                        return new XElement(element.Name,
                            element.Attributes(),
                            element.Nodes().Select(n =>
                                AcceptParagraphEndTagsInMoveFromTransform(n)));
                    }
                }
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => AcceptParagraphEndTagsInMoveFromTransform(n)));
            }
            return node;
        }

        private static object? AcceptAllOtherRevisionsTransform(XNode? node)
        {
            if (node is XElement element)
            {
                // Accept inserted text, inserted paragraph marks, etc.
                // Collapse all w:ins elements.

                if (element.Name == W.ins)
                {
                    return element
                        .Nodes()
                        .Select(n => AcceptAllOtherRevisionsTransform(n));
                }

                // Remove all of the following elements.  These elements are processed in:
                //   AcceptDeletedAndMovedFromContentControls
                //   AcceptMoveFromMoveToTransform
                //   AcceptDeletedAndMoveFromParagraphMarksTransform
                //   AcceptParagraphEndTagsInMoveFromTransform
                //   AcceptMoveFromRanges

                if (element.Name == W.customXmlDelRangeStart ||
                    element.Name == W.customXmlDelRangeEnd ||
                    element.Name == W.customXmlInsRangeStart ||
                    element.Name == W.customXmlInsRangeEnd ||
                    element.Name == W.customXmlMoveFromRangeStart ||
                    element.Name == W.customXmlMoveFromRangeEnd ||
                    element.Name == W.customXmlMoveToRangeStart ||
                    element.Name == W.customXmlMoveToRangeEnd ||
                    element.Name == W.moveFromRangeStart ||
                    element.Name == W.moveFromRangeEnd ||
                    element.Name == W.moveToRangeStart ||
                    element.Name == W.moveToRangeEnd)
                {
                    return null;
                }

                // Accept revisions in formatting on paragraphs.
                // Accept revisions in formatting on runs.
                // Accept revisions for applied styles to a table.
                // Accept revisions for grid revisions to a table.
                // Accept revisions for column properties.
                // Accept revisions for row properties.
                // Accept revisions for table level property exceptions.
                // Accept revisions for section properties.
                // Accept numbering revision in fields.
                // Accept deleted field code text.
                // Accept deleted literal text.
                // Accept inserted cell.

                if (element.Name == W.pPrChange ||
                    element.Name == W.rPrChange ||
                    element.Name == W.tblPrChange ||
                    element.Name == W.tblGridChange ||
                    element.Name == W.tcPrChange ||
                    element.Name == W.trPrChange ||
                    element.Name == W.tblPrExChange ||
                    element.Name == W.sectPrChange ||
                    element.Name == W.numberingChange ||
                    element.Name == W.delInstrText ||
                    element.Name == W.delText ||
                    element.Name == W.cellIns)
                {
                    return null;
                }

                // Accept revisions for deleted math control character.
                // Match m:f/m:fPr/m:ctrlPr/w:del, remove m:f.

                if (element.Name == M.f &&
                    element.Elements(M.fPr).Elements(M.ctrlPr).Elements(W.del).Any())
                {
                    return null;
                }

                // Accept revisions for deleted rows in tables.
                // Match w:tr/w:trPr/w:del, remove w:tr.

                if (element.Name == W.tr &&
                    element.Elements(W.trPr).Elements(W.del).Any())
                {
                    return null;
                }

                // Accept deleted text in paragraphs.

                if (element.Name == W.del)
                {
                    return null;
                }

                // Accept revisions for vertically merged cells.
                //   cellMerge with a parent of tcPr, with attribute w:vMerge="rest" transformed
                //     to <w:vMerge w:val="restart"/>
                //   cellMerge with a parent of tcPr, with attribute w:vMerge="cont" transformed
                //     to <w:vMerge w:val="continue"/>

                if (element.Name == W.cellMerge &&
                    element.Parent.Name == W.tcPr &&
                    (string)element.Attribute(W.vMerge) == "rest")
                {
                    return new XElement(W.vMerge,
                        new XAttribute(W.val, "restart"));
                }

                if (element.Name == W.cellMerge &&
                    element.Parent.Name == W.tcPr &&
                    (string)element.Attribute(W.vMerge) == "cont")
                {
                    return new XElement(W.vMerge,
                        new XAttribute(W.val, "continue"));
                }

                // Otherwise do identity clone.
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => AcceptAllOtherRevisionsTransform(n)));
            }
            return node;
        }

        private static object CollapseParagraphTransform(XNode node)
        {
            if (node is XElement element)
            {
                if (element.Name == W.p)
                {
                    return element.Elements().Where(e => e.Name != W.pPr);
                }

                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => CollapseParagraphTransform(n)));
            }
            return node;
        }

        private static void AnnotateBlockContentElements(XElement contentContainer)
        {
            // For convenience, there is a ParagraphInfo annotation on the contentContainer.
            // It contains the same information as the ParagraphInfo annotation on the first
            //   paragraph.
            if (contentContainer.Annotation<BlockContentInfo>() != null)
            {
                return;
            }

            var firstContentElement = contentContainer
                .Elements()
                .DescendantsAndSelf()
                .FirstOrDefault(e => e.Name == W.p || e.Name == W.tbl);
            if (firstContentElement == null)
            {
                return;
            }

            // Add the annotation on the contentContainer.
            var currentContentInfo = new BlockContentInfo()
            {
                PreviousBlockContentElement = null,
                ThisBlockContentElement = firstContentElement,
                NextBlockContentElement = null
            };
            // Add as annotation even though NextParagraph is not set yet.
            contentContainer.AddAnnotation(currentContentInfo);
            while (true)
            {
                currentContentInfo.ThisBlockContentElement.AddAnnotation(currentContentInfo);
                // Find next sibling content element.
                XElement? nextContentElement = null;
                var current = currentContentInfo.ThisBlockContentElement;
                while (true)
                {
                    nextContentElement = current
                        .ElementsAfterSelf()
                        .DescendantsAndSelf()
                        .FirstOrDefault(e => e.Name == W.p || e.Name == W.tbl);
                    if (nextContentElement != null)
                    {
                        currentContentInfo.NextBlockContentElement = nextContentElement;
                        break;
                    }
                    current = current.Parent;
                    // When we've backed up the tree to the contentContainer, we're done.
                    if (current == contentContainer)
                    {
                        return;
                    }
                }
                currentContentInfo = new BlockContentInfo()
                {
                    PreviousBlockContentElement = currentContentInfo.ThisBlockContentElement,
                    ThisBlockContentElement = nextContentElement,
                    NextBlockContentElement = null
                };
            }
        }

        private static void AnnotateRunElementsWithId(XElement element)
        {
            var runId = 0;
            foreach (var e in element.Descendants().Where(e => e.Name == W.r))
            {
                if (e.Name == W.r)
                {
                    e.Add(new XAttribute(PT.UniqueId, runId++));
                }
            }
        }

        private static void AnnotateContentControlsWithRunIds(XElement element)
        {
            var sdtId = 0;
            foreach (var e in element.Descendants(W.sdt))
            {
                e.Add(new XAttribute(PT.RunIds,
                    e.DescendantsTrimmed(W.txbxContent)
                     .Where(d => d.Name == W.r)
                     .Select(r => r.Attribute(PT.UniqueId).Value)
                     .StringConcatenate(s => s + ",")
                     .Trim(',')),
                    new XAttribute(PT.UniqueId, sdtId++));
            }
        }

        private static XElement AddBlockLevelContentControls(XElement newDocument, XElement original)
        {
            var originalContentControls = original.Descendants(W.sdt).ToList();
            var existingContentControls = newDocument.Descendants(W.sdt).ToList();
            var contentControlsToAdd = originalContentControls
                .Select(occ => occ.Attribute(PT.UniqueId).Value)
                .Except(existingContentControls
                    .Select(ecc => ecc.Attribute(PT.UniqueId).Value));
            foreach (var contentControl in originalContentControls
                .Where(occ => contentControlsToAdd.Contains(occ.Attribute(PT.UniqueId).Value)))
            {
                // TODO - Need a slight modification here.  If there is a paragraph
                // in the content control that contains no runs, then the paragraph isn't included in the
                // content control, because the following triggers off of runs.
                // To see an example of this, see example document "NumberingParagraphPropertiesChange.docxs"

                // find list of runs to surround
                var runIds = contentControl.Attribute(PT.RunIds).Value.Split(',');
                var runs = contentControl.Descendants(W.r).Where(r => runIds.Contains(r.Attribute(PT.UniqueId).Value));
                // find the runs in the new document

                var runsInNewDocument = runs.Select(r => newDocument.Descendants(W.r).First(z => z.Attribute(PT.UniqueId).Value == r.Attribute(PT.UniqueId).Value)).ToList();

                // find common ancestor
                List<XElement>? runAncestorIntersection = null;
                foreach (var run in runsInNewDocument)
                {
                    if (runAncestorIntersection == null)
                    {
                        runAncestorIntersection = run.Ancestors().ToList();
                    }
                    else
                    {
                        runAncestorIntersection = run.Ancestors().Intersect(runAncestorIntersection).ToList();
                    }
                }
                if (runAncestorIntersection == null)
                {
                    continue;
                }

                var commonAncestor = runAncestorIntersection.InDocumentOrder().Last();
                // find child of common ancestor that contains first run
                // find child of common ancestor that contains last run
                // create new common ancestor:
                //   elements before first run child
                //   add content control, and runs from first run child to last run child
                //   elements after last run child
                var firstRunChild = commonAncestor
                    .Elements()
                    .First(c => c.DescendantsAndSelf()
                        .Any(z => z.Name == W.r &&
                             z.Attribute(PT.UniqueId).Value == runsInNewDocument.First().Attribute(PT.UniqueId).Value));
                var lastRunChild = commonAncestor
                    .Elements()
                    .First(c => c.DescendantsAndSelf()
                        .Any(z => z.Name == W.r &&
                             z.Attribute(PT.UniqueId).Value == runsInNewDocument.Last().Attribute(PT.UniqueId).Value));

                // If the list of runs for the content control is exactly the list of runs for the paragraph, then
                // create the content control surrounding the paragraph, not surrounding the runs.

                if (commonAncestor.Name == W.p &&
                    commonAncestor.Elements()
                        .FirstOrDefault(e => e.Name != W.pPr && e.Name != W.commentRangeStart && e.Name != W.commentRangeEnd) == firstRunChild &&
                    commonAncestor.Elements()
                        .LastOrDefault(e => e.Name != W.pPr && e.Name != W.commentRangeStart && e.Name != W.commentRangeEnd) == lastRunChild)
                {
                    var newContentControlOrdered = new XElement(contentControl.Name,
                        contentControl.Attributes(),
                        contentControl.Elements().OrderBy(e =>
                        {
                            if (Order_sdt.ContainsKey(e.Name))
                            {
                                return Order_sdt[e.Name];
                            }

                            return 999;
                        }));

                    commonAncestor.ReplaceWith(newContentControlOrdered);
                    continue;
                }

                var elementsBeforeRange = commonAncestor
                    .Elements()
                    .TakeWhile(e => e != firstRunChild)
                    .ToList();
                var elementsInRange = commonAncestor
                    .Elements()
                    .SkipWhile(e => e != firstRunChild)
                    .TakeWhile(e => e != lastRunChild.ElementsAfterSelf().FirstOrDefault())
                    .ToList();
                var elementsAfterRange = commonAncestor
                    .Elements()
                    .SkipWhile(e => e != lastRunChild.ElementsAfterSelf().FirstOrDefault())
                    .ToList();

                // detatch from current parent
                commonAncestor.Elements().Remove();

                var newContentControl2 = new XElement(contentControl.Name,
                    contentControl.Attributes(),
                    contentControl.Elements().Where(e => e.Name != W.sdtContent),
                    new XElement(W.sdtContent, elementsInRange));

                var newContentControlOrdered2 = new XElement(newContentControl2.Name,
                    newContentControl2.Attributes(),
                    newContentControl2.Elements().OrderBy(e =>
                    {
                        if (Order_sdt.ContainsKey(e.Name))
                        {
                            return Order_sdt[e.Name];
                        }

                        return 999;
                    }));

                commonAncestor.Add(
                    elementsBeforeRange,
                    newContentControlOrdered2,
                    elementsAfterRange);
            }
            return newDocument;
        }

        private static XElement AcceptDeletedAndMoveFromParagraphMarks(XElement? element)
        {
            AnnotateRunElementsWithId(element);
            AnnotateContentControlsWithRunIds(element);
            var newElement = (XElement)AcceptDeletedAndMoveFromParagraphMarksTransform(element);
            var withBlockLevelContentControls = AddBlockLevelContentControls(newElement, element);
            return withBlockLevelContentControls;
        }

        private static object AcceptDeletedAndMoveFromParagraphMarksTransform(XNode node)
        {
            if (node is XElement element)
            {
                if (W.BlockLevelContentContainers.Contains(element.Name))
                {
                    XElement? bodySectPr = null;
                    if (element.Name == W.body)
                    {
                        bodySectPr = element.Element(W.sectPr);
                    }

                    var currentKey = 0;
                    var deletedParagraphGroupingInfo = new List<GroupingInfo>();

                    var state = 0; // 0 = in non deleted paragraphs
                                   // 1 = in deleted paragraph
                                   // 2 - paragraph following deleted paragraphs

                    foreach (var c in IterateBlockContentElements(element))
                    {
                        if (c.ThisBlockContentElement.Name == W.p)
                        {
                            var paragraphMarkIsDeletedOrMovedFrom = c
                                .ThisBlockContentElement
                                .Elements(W.pPr)
                                .Elements(W.rPr)
                                .Elements()
.Any(e => e.Name == W.del || e.Name == W.moveFrom);

                            if (paragraphMarkIsDeletedOrMovedFrom)
                            {
                                if (state == 0 || state == 2)
                                {
                                    state = 1;
                                    currentKey += 1;
                                    deletedParagraphGroupingInfo.Add(
                                        new GroupingInfo()
                                        {
                                            GroupingType = GroupingType.DeletedRange,
                                            GroupingKey = currentKey,
                                        });
                                    continue;
                                }
                                else if (state == 1)
                                {
                                    deletedParagraphGroupingInfo.Add(
                                        new GroupingInfo()
                                        {
                                            GroupingType = GroupingType.DeletedRange,
                                            GroupingKey = currentKey,
                                        });
                                    continue;
                                }
                            }

                            if (state == 0)
                            {
                                currentKey += 1;
                                deletedParagraphGroupingInfo.Add(
                                    new GroupingInfo()
                                    {
                                        GroupingType = GroupingType.Other,
                                        GroupingKey = currentKey,
                                    });
                                continue;
                            }
                            else if (state == 1)
                            {
                                state = 2;
                                deletedParagraphGroupingInfo.Add(
                                    new GroupingInfo()
                                    {
                                        GroupingType = GroupingType.DeletedRange,
                                        GroupingKey = currentKey,
                                    });
                                continue;
                            }
                            else if (state == 2)
                            {
                                state = 0;
                                currentKey += 1;
                                deletedParagraphGroupingInfo.Add(
                                    new GroupingInfo()
                                    {
                                        GroupingType = GroupingType.Other,
                                        GroupingKey = currentKey,
                                    });
                                continue;
                            }
                        }
                        else if (c.ThisBlockContentElement.Name == W.tbl || c.ThisBlockContentElement.Name.Namespace == M.m)
                        {
                            currentKey += 1;
                            deletedParagraphGroupingInfo.Add(
                                new GroupingInfo()
                                {
                                    GroupingType = GroupingType.Other,
                                    GroupingKey = currentKey,
                                });
                            state = 0;
                            continue;
                        }
                        else
                        {
                            // otherwise keep the same state, put in the same group, and continue
                            deletedParagraphGroupingInfo.Add(
                                new GroupingInfo()
                                {
                                    GroupingType = GroupingType.Other,
                                    GroupingKey = currentKey,
                                });
                            continue;
                        }
                    }

                    var zipped = IterateBlockContentElements(element).Zip(deletedParagraphGroupingInfo, (blc, gi) => new
                    {
                        BlockLevelContent = blc,
                        GroupingInfo = gi,
                    });

                    var groupedParagraphs = zipped
                        .GroupAdjacent(z => z.GroupingInfo.GroupingKey);

                    // Create a new block level content container.
                    var newBlockLevelContentContainer = new XElement(element.Name,
                        element.Attributes(),
                        element.Elements().Where(e => e.Name == W.tcPr),
                        groupedParagraphs.Select((g, i) =>
                        {
                            if (g.First().GroupingInfo.GroupingType == GroupingType.DeletedRange)
                            {
                                var newParagraph = new XElement(W.p,
                                    g.Last().BlockLevelContent.ThisBlockContentElement.Elements(W.pPr),
                                    g.Select(z => CollapseParagraphTransform(z.BlockLevelContent.ThisBlockContentElement)));

                                // if this contains the last paragraph in the document, and if there is no content,
                                // and if the paragraph mark is deleted, then nuke the paragraph.
                                var allIsDeleted = AllParaContentIsDeleted(newParagraph);
                                if (allIsDeleted &&
                                    g.Last().BlockLevelContent.ThisBlockContentElement.Elements(W.pPr).Elements(W.rPr).Elements(W.del).Any() &&
                                    (g.Last().BlockLevelContent.NextBlockContentElement == null ||
                                     g.Last().BlockLevelContent.NextBlockContentElement.Name == W.tbl))
                                {
                                    return null;
                                }

                                return (object)newParagraph;
                            }
                            else
                            {
                                return g.Select(z =>
                                {
                                    var newEle = new XElement(z.BlockLevelContent.ThisBlockContentElement.Name,
                                        z.BlockLevelContent.ThisBlockContentElement.Attributes(),
                                        z.BlockLevelContent.ThisBlockContentElement.Nodes().Select(n => AcceptDeletedAndMoveFromParagraphMarksTransform(n)));
                                    return newEle;
                                });
                            }
                        }),
                        bodySectPr);

                    return newBlockLevelContentContainer;
                }

                // Otherwise, identity clone.
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => AcceptDeletedAndMoveFromParagraphMarksTransform(n)));
            }
            return node;
        }

        private enum DeletedCellCollectionType
        {
            DeletedCell,
            Other
        }

        // For each table row, group deleted cells plus the cell before any deleted cell.
        // Produce a new cell that has gridSpan set appropriately for group, and clone everything
        // else.
        private static object AcceptDeletedCellsTransform(XNode? node)
        {
            if (node is XElement element)
            {
                if (element.Name == W.tr)
                {
                    var groupedCells = element
                        .Elements()
                        .GroupAdjacent(e =>
                        {
                            var cellAfter = e.ElementsAfterSelf(W.tc).FirstOrDefault();
                            var cellAfterIsDeleted = cellAfter != null &&
                                cellAfter.Descendants(W.cellDel).Any();
                            if (e.Name == W.tc &&
                                (cellAfterIsDeleted || e.Descendants(W.cellDel).Any()))
                            {
                                var a = new
                                {
                                    CollectionType = DeletedCellCollectionType.DeletedCell,
                                    Disambiguator = new[] { e }
                                        .Concat(e.SiblingsBeforeSelfReverseDocumentOrder())
                                        .FirstOrDefault(z => z.Name == W.tc &&
                                            !z.Descendants(W.cellDel).Any())
                                };
                                return a;
                            }
                            var a2 = new
                            {
                                CollectionType = DeletedCellCollectionType.Other,
                                Disambiguator = e
                            };
                            return a2;
                        });
                    var tr = new XElement(W.tr,
                        element.Attributes(),
                        groupedCells.Select(g =>
                        {
                            if (g.Key.CollectionType == DeletedCellCollectionType.DeletedCell
                                && g.First().Descendants(W.cellDel).Any())
                            {
                                return null;
                            }

                            if (g.Key.CollectionType == DeletedCellCollectionType.Other)
                            {
                                return g;
                            }

                            var gridSpanElement = g
                                .First()
                                .Elements(W.tcPr)
                                .Elements(W.gridSpan)
                                .FirstOrDefault();
                            var gridSpan = gridSpanElement != null ?
                                (int)gridSpanElement.Attribute(W.val) :
                                1;
                            var newGridSpan = gridSpan + g.Count() - 1;
                            var currentTcPr = g.First().Elements(W.tcPr).FirstOrDefault();
                            var newTcPr = new XElement(W.tcPr,
                                currentTcPr?.Attributes(),
                                new XElement(W.gridSpan,
                                    new XAttribute(W.val, newGridSpan)),
                                currentTcPr.Elements().Where(e => e.Name != W.gridSpan));
                            var orderedTcPr = new XElement(W.tcPr,
                                newTcPr.Elements().OrderBy(e =>
                                {
                                    if (Order_tcPr.ContainsKey(e.Name))
                                    {
                                        return Order_tcPr[e.Name];
                                    }

                                    return 999;
                                }));
                            var newTc = new XElement(W.tc,
                                orderedTcPr,
                                g.First().Elements().Where(e => e.Name != W.tcPr));
                            return (object)newTc;
                        }));
                    return tr;
                }

                // Identity clone
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => AcceptDeletedCellsTransform(n)));
            }
            return node;
        }

        private static bool AllParaContentIsDeleted(XElement p)
        {
            // needs collapse
            // dir, bdo, sdt, ins, moveTo, smartTag
            var testP = CollapseTransform(p) as XElement;

            var childElements = testP?.Elements();
            var contentElements = childElements
                .Where(ce =>
                {
                    var b = IsRunContent(ce.Name);
                    if (b != null)
                    {
                        return (bool)b;
                    }

                    throw new Exception("Internal error 20, found element " + ce.Name.ToString());
                });
            if (contentElements.Any())
            {
                return false;
            }

            return true;
        }

        private static object? CollapseTransform(XNode node)
        {
            if (node is XElement element)
            {
                if (element.Name == W.dir ||
                    element.Name == W.bdr ||
                    element.Name == W.ins ||
                    element.Name == W.moveTo ||
                    element.Name == W.smartTag)
                {
                    return element.Elements();
                }

                if (element.Name == W.sdt)
                {
                    return element.Elements(W.sdtContent).Elements();
                }

                if (element.Name == W.pPr)
                {
                    return null;
                }

                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => CollapseTransform(n)));
            }
            return node;
        }

        private static bool? IsRunContent(XName ceName)
        {
            // is content
            // r, fldSimple, hyperlink, oMath, oMathPara, subDoc
            if (ceName == W.r ||
                ceName == W.fldSimple ||
                ceName == W.hyperlink ||
                ceName == W.subDoc ||
                ceName == W.smartTag ||
                ceName == W.smartTagPr ||
                ceName.Namespace == M.m)
            {
                return true;
            }

            // not content
            // bookmarkStart, bookmarkEnd, commentRangeStart, commentRangeEnd, del, moveFrom, proofErr
            if (ceName == W.bookmarkStart ||
                ceName == W.bookmarkEnd ||
                ceName == W.commentRangeStart ||
                ceName == W.commentRangeEnd ||
                ceName == W.customXmlDelRangeStart ||
                ceName == W.customXmlDelRangeEnd ||
                ceName == W.customXmlInsRangeStart ||
                ceName == W.customXmlInsRangeEnd ||
                ceName == W.customXmlMoveFromRangeStart ||
                ceName == W.customXmlMoveFromRangeEnd ||
                ceName == W.customXmlMoveToRangeStart ||
                ceName == W.customXmlMoveToRangeEnd ||
                ceName == W.del ||
                ceName == W.moveFrom ||
                ceName == W.moveFromRangeStart ||
                ceName == W.moveFromRangeEnd ||
                ceName == W.moveToRangeStart ||
                ceName == W.moveToRangeEnd ||
                ceName == W.permStart ||
                ceName == W.permEnd ||
                ceName == W.proofErr)
            {
                return false;
            }

            return null;
        }

        private static object? AcceptDeletedAndMovedFromContentControlsTransform(XNode node,
            XElement[] contentControlElementsToCollapse,
            XElement[] moveFromElementsToDelete)
        {
            if (node is XElement element)
            {
                if (element.Name == W.sdt && contentControlElementsToCollapse.Contains(element))
                {
                    return element
                        .Element(W.sdtContent)
                        .Nodes()
                        .Select(n => AcceptDeletedAndMovedFromContentControlsTransform(
                            n, contentControlElementsToCollapse, moveFromElementsToDelete));
                }

                if (moveFromElementsToDelete.Contains(element))
                {
                    return null;
                }

                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => AcceptDeletedAndMovedFromContentControlsTransform(
                        n, contentControlElementsToCollapse, moveFromElementsToDelete)));
            }
            return node;
        }

        private static XElement? AcceptDeletedAndMovedFromContentControls(XElement documentRootElement)
        {
            // The following lists contain the elements that are between start/end elements.
            var startElementTagsInDeleteRange = new List<XElement>();
            var endElementTagsInDeleteRange = new List<XElement>();
            var startElementTagsInMoveFromRange = new List<XElement>();
            var endElementTagsInMoveFromRange = new List<XElement>();

            // Following are the elements that *may* be in a range that has both start and end
            // elements.
            var potentialDeletedElements =
                new Dictionary<string, PotentialInRangeElements>();
            var potentialMoveFromElements =
                new Dictionary<string, PotentialInRangeElements>();

            foreach (var tag in DescendantAndSelfTags(documentRootElement))
            {
                if (tag.Element.Name == W.customXmlDelRangeStart)
                {
                    var id = tag.Element.Attribute(W.id).Value;
                    potentialDeletedElements.Add(id, new PotentialInRangeElements());
                    continue;
                }
                if (tag.Element.Name == W.customXmlDelRangeEnd)
                {
                    var id = tag.Element.Attribute(W.id).Value;
                    if (potentialDeletedElements.ContainsKey(id))
                    {
                        startElementTagsInDeleteRange.AddRange(
                            potentialDeletedElements[id].PotentialStartElementTagsInRange);
                        endElementTagsInDeleteRange.AddRange(
                            potentialDeletedElements[id].PotentialEndElementTagsInRange);
                        potentialDeletedElements.Remove(id);
                    }
                    continue;
                }
                if (tag.Element.Name == W.customXmlMoveFromRangeStart)
                {
                    var id = tag.Element.Attribute(W.id).Value;
                    potentialMoveFromElements.Add(id, new PotentialInRangeElements());
                    continue;
                }
                if (tag.Element.Name == W.customXmlMoveFromRangeEnd)
                {
                    var id = tag.Element.Attribute(W.id).Value;
                    if (potentialMoveFromElements.ContainsKey(id))
                    {
                        startElementTagsInMoveFromRange.AddRange(
                            potentialMoveFromElements[id].PotentialStartElementTagsInRange);
                        endElementTagsInMoveFromRange.AddRange(
                            potentialMoveFromElements[id].PotentialEndElementTagsInRange);
                        potentialMoveFromElements.Remove(id);
                    }
                    continue;
                }
                if (tag.Element.Name == W.sdt)
                {
                    if (tag.TagType == TagTypeEnum.Element)
                    {
                        foreach (var id in potentialDeletedElements)
                        {
                            id.Value.PotentialStartElementTagsInRange.Add(tag.Element);
                        }

                        foreach (var id in potentialMoveFromElements)
                        {
                            id.Value.PotentialStartElementTagsInRange.Add(tag.Element);
                        }

                        continue;
                    }
                    if (tag.TagType == TagTypeEnum.EmptyElement)
                    {
                        foreach (var id in potentialDeletedElements)
                        {
                            id.Value.PotentialStartElementTagsInRange.Add(tag.Element);
                            id.Value.PotentialEndElementTagsInRange.Add(tag.Element);
                        }
                        foreach (var id in potentialMoveFromElements)
                        {
                            id.Value.PotentialStartElementTagsInRange.Add(tag.Element);
                            id.Value.PotentialEndElementTagsInRange.Add(tag.Element);
                        }
                        continue;
                    }
                    if (tag.TagType == TagTypeEnum.EndElement)
                    {
                        foreach (var id in potentialDeletedElements)
                        {
                            id.Value.PotentialEndElementTagsInRange.Add(tag.Element);
                        }

                        foreach (var id in potentialMoveFromElements)
                        {
                            id.Value.PotentialEndElementTagsInRange.Add(tag.Element);
                        }

                        continue;
                    }
                    throw new PowerToolsInvalidDataException("Should not have reached this point.");
                }
                if (potentialMoveFromElements.Any() &&
                    tag.Element.Name != W.moveFromRangeStart &&
                    tag.Element.Name != W.moveFromRangeEnd &&
                    tag.Element.Name != W.customXmlMoveFromRangeStart &&
                    tag.Element.Name != W.customXmlMoveFromRangeEnd)
                {
                    if (tag.TagType == TagTypeEnum.Element)
                    {
                        foreach (var id in potentialMoveFromElements)
                        {
                            id.Value.PotentialStartElementTagsInRange.Add(tag.Element);
                        }

                        continue;
                    }
                    if (tag.TagType == TagTypeEnum.EmptyElement)
                    {
                        foreach (var id in potentialMoveFromElements)
                        {
                            id.Value.PotentialStartElementTagsInRange.Add(tag.Element);
                            id.Value.PotentialEndElementTagsInRange.Add(tag.Element);
                        }
                        continue;
                    }
                    if (tag.TagType == TagTypeEnum.EndElement)
                    {
                        foreach (var id in potentialMoveFromElements)
                        {
                            id.Value.PotentialEndElementTagsInRange.Add(tag.Element);
                        }

                        continue;
                    }
                }
            }

            var contentControlElementsToCollapse = startElementTagsInDeleteRange
                .Intersect(endElementTagsInDeleteRange)
                .ToArray();
            var elementsToDeleteBecauseMovedFrom = startElementTagsInMoveFromRange
                .Intersect(endElementTagsInMoveFromRange)
                .ToArray();
            if (contentControlElementsToCollapse.Length > 0 ||
                elementsToDeleteBecauseMovedFrom.Length > 0)
            {
                var newDoc = AcceptDeletedAndMovedFromContentControlsTransform(documentRootElement,
                    contentControlElementsToCollapse, elementsToDeleteBecauseMovedFrom);
                return newDoc as XElement;
            }
            else
            {
                return documentRootElement;
            }
        }

        private static object? AcceptMoveFromRangesTransform(XNode node,
            XElement[] elementsToDelete)
        {
            if (node is XElement element)
            {
                if (elementsToDelete.Contains(element))
                {
                    return null;
                }

                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n =>
                        AcceptMoveFromRangesTransform(n, elementsToDelete)));
            }
            return node;
        }

        private static object CoalesqueParagraphEndTagsInMoveFromTransform(XNode node,
            IGrouping<MoveFromCollectionType, XElement> g)
        {
            if (node is XElement element)
            {
                if (element.Name == W.p)
                {
                    return new XElement(W.p,
                        element.Attributes(),
                        element.Elements(),
                        g.Skip(1).Select(p => CollapseParagraphTransform(p)));
                }
                else
                {
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Nodes().Select(n =>
                            CoalesqueParagraphEndTagsInMoveFromTransform(n, g)));
                }
            }
            return node;
        }

        private static object? RemoveRowsLeftEmptyByMoveFrom(XNode node)
        {
            if (node is XElement element)
            {
                if (element.Name == W.tr)
                {
                    var nonEmptyCells = element.Elements(W.tc).Any(tc => tc.Elements().Any(tcc => BlockLevelElements.Contains(tcc.Name)));
                    if (nonEmptyCells)
                    {
                        return new XElement(element.Name,
                            element.Attributes(),
                            element.Nodes().Select(n => RemoveRowsLeftEmptyByMoveFrom(n)));
                    }
                    return null;
                }

                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => RemoveRowsLeftEmptyByMoveFrom(n)));
            }
            return node;
        }

        private enum MoveFromCollectionType
        {
            ParagraphEndTagInMoveFromRange,
            Other
        };

        private enum GroupingType
        {
            DeletedRange,
            Other,
        };

        private class GroupingInfo
        {
            public GroupingType GroupingType;
            public int GroupingKey;
        };

        private static readonly XName[] BlockLevelElements = new[] {
            W.p,
            W.tbl,
            W.sdt,
            W.del,
            W.ins,
            M.oMath,
            M.oMathPara,
            W.moveTo,
        };

        private static readonly Dictionary<XName, int> Order_sdt = new Dictionary<XName, int>
        {
            { W.sdtPr, 10 },
            { W.sdtEndPr, 20 },
            { W.sdtContent, 30 },
            { W.bookmarkStart, 40 },
            { W.bookmarkEnd, 50 },
        };

    }
}
