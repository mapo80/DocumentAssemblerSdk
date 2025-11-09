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
        public static WmlDocument RejectRevisions(WmlDocument document)
        {
            using var streamDoc = new OpenXmlMemoryStreamDocument(document);
            using (var doc = streamDoc.GetWordprocessingDocument())
            {
                RejectRevisions(doc);
            }
            return streamDoc.GetModifiedWmlDocument();
        }

        public static void RejectRevisions(WordprocessingDocument doc)
        {
            RejectRevisionsForPart(doc.MainDocumentPart);
            foreach (var part in doc.MainDocumentPart.HeaderParts)
            {
                RejectRevisionsForPart(part);
            }

            foreach (var part in doc.MainDocumentPart.FooterParts)
            {
                RejectRevisionsForPart(part);
            }

            if (doc.MainDocumentPart.EndnotesPart != null)
            {
                RejectRevisionsForPart(doc.MainDocumentPart.EndnotesPart);
            }

            if (doc.MainDocumentPart.FootnotesPart != null)
            {
                RejectRevisionsForPart(doc.MainDocumentPart.FootnotesPart);
            }

            if (doc.MainDocumentPart.StyleDefinitionsPart != null)
            {
                RejectRevisionsForStylesDefinitionPart(doc.MainDocumentPart.StyleDefinitionsPart);
            }

            ReverseRevisions(doc);
            AcceptRevisionsForPart(doc.MainDocumentPart);
            foreach (var part in doc.MainDocumentPart.HeaderParts)
            {
                AcceptRevisionsForPart(part);
            }

            foreach (var part in doc.MainDocumentPart.FooterParts)
            {
                AcceptRevisionsForPart(part);
            }

            if (doc.MainDocumentPart.EndnotesPart != null)
            {
                AcceptRevisionsForPart(doc.MainDocumentPart.EndnotesPart);
            }

            if (doc.MainDocumentPart.FootnotesPart != null)
            {
                AcceptRevisionsForPart(doc.MainDocumentPart.FootnotesPart);
            }

            if (doc.MainDocumentPart.StyleDefinitionsPart != null)
            {
                AcceptRevisionsForStylesDefinitionPart(doc.MainDocumentPart.StyleDefinitionsPart);
            }
        }

        public static WmlDocument AcceptRevisions(WmlDocument document)
        {
            using var streamDoc = new OpenXmlMemoryStreamDocument(document);
            using (var doc = streamDoc.GetWordprocessingDocument())
            {
                AcceptRevisions(doc);
            }
            return streamDoc.GetModifiedWmlDocument();
        }

        public static void AcceptRevisions(WordprocessingDocument doc)
        {
            AcceptRevisionsForPart(doc.MainDocumentPart);
            foreach (var part in doc.MainDocumentPart.HeaderParts)
            {
                AcceptRevisionsForPart(part);
            }

            foreach (var part in doc.MainDocumentPart.FooterParts)
            {
                AcceptRevisionsForPart(part);
            }

            if (doc.MainDocumentPart.EndnotesPart != null)
            {
                AcceptRevisionsForPart(doc.MainDocumentPart.EndnotesPart);
            }

            if (doc.MainDocumentPart.FootnotesPart != null)
            {
                AcceptRevisionsForPart(doc.MainDocumentPart.FootnotesPart);
            }

            if (doc.MainDocumentPart.StyleDefinitionsPart != null)
            {
                AcceptRevisionsForStylesDefinitionPart(doc.MainDocumentPart.StyleDefinitionsPart);
            }
        }

        public static XElement? AcceptRevisionsForElement(XElement element)
        {
            var rElement = element;
            rElement = RemoveRsidTransform(rElement) as XElement;
            rElement = AcceptMoveFromMoveToTransform(rElement) as XElement;
            rElement = AcceptAllOtherRevisionsTransform(rElement) as XElement;
            rElement?.Descendants().Attributes().Where(a => a.Name == PT.UniqueId || a.Name == PT.RunIds).Remove();
            rElement?.Descendants(W.numPr).Where(np => !np.HasElements).Remove();
            return rElement;
        }

        public static readonly XName[] TrackedRevisionsElements = new[]
        {
            W.cellDel,
            W.cellIns,
            W.cellMerge,
            W.customXmlDelRangeEnd,
            W.customXmlDelRangeStart,
            W.customXmlInsRangeEnd,
            W.customXmlInsRangeStart,
            W.del,
            W.delInstrText,
            W.delText,
            W.ins,
            W.moveFrom,
            W.moveFromRangeEnd,
            W.moveFromRangeStart,
            W.moveTo,
            W.moveToRangeEnd,
            W.moveToRangeStart,
            W.numberingChange,
            W.pPrChange,
            W.rPrChange,
            W.sectPrChange,
            W.tblGridChange,
            W.tblPrChange,
            W.tblPrExChange,
            W.tcPrChange,
            W.trPrChange,
        };

        public static bool PartHasTrackedRevisions(OpenXmlPart part)
        {
            return part.GetXDocument()
                .Descendants()
                .Any(e => TrackedRevisionsElements.Contains(e.Name));
        }

        public static bool HasTrackedRevisions(WmlDocument document)
        {
            using var streamDoc = new OpenXmlMemoryStreamDocument(document);
            using var wdoc = streamDoc.GetWordprocessingDocument();
            return RevisionAccepter.HasTrackedRevisions(wdoc);
        }

        public static bool HasTrackedRevisions(WordprocessingDocument doc)
        {
            if (PartHasTrackedRevisions(doc.MainDocumentPart))
            {
                return true;
            }

            foreach (var part in doc.MainDocumentPart.HeaderParts)
            {
                if (PartHasTrackedRevisions(part))
                {
                    return true;
                }
            }

            foreach (var part in doc.MainDocumentPart.FooterParts)
            {
                if (PartHasTrackedRevisions(part))
                {
                    return true;
                }
            }

            if (doc.MainDocumentPart.EndnotesPart != null)
            {
                if (PartHasTrackedRevisions(doc.MainDocumentPart.EndnotesPart))
                {
                    return true;
                }
            }

            if (doc.MainDocumentPart.FootnotesPart != null)
            {
                if (PartHasTrackedRevisions(doc.MainDocumentPart.FootnotesPart))
                {
                    return true;
                }
            }

            return false;
        }

        public static class PT
        {
            public static readonly XNamespace pt = "http://www.codeplex.com/PowerTools/2009/RevisionAccepter";
            public static readonly XName UniqueId = pt + "UniqueId";
            public static readonly XName RunIds = pt + "RunIds";
        }
    }

    public partial class WmlDocument : OpenXmlPowerToolsDocument
    {
        public WmlDocument AcceptRevisions(WmlDocument document)
        {
            return RevisionAccepter.AcceptRevisions(document);
        }

        public bool HasTrackedRevisions(WmlDocument document)
        {
            return RevisionAccepter.HasTrackedRevisions(document);
        }
    }

    public static class RevisionAccepterExtensions
    {
        private static void InitializeParagraphInfo(XElement contentContext)
        {
            if (!(W.BlockLevelContentContainers.Contains(contentContext.Name)))
            {
                throw new ArgumentException(
                    "GetParagraphInfo called for element that is not child of content container");
            }

            XElement? prev = null;
            foreach (var content in contentContext.Elements())
            {
                // This may return null, indicating that there is no descendant paragraph.  For
                // example, comment elements have no descendant elements.
                var paragraph = content
                    .DescendantsAndSelf()
                    .FirstOrDefault(e => e.Name == W.p || e.Name == W.tc || e.Name == W.txbxContent);
                if (paragraph != null &&
                    (paragraph.Name == W.tc || paragraph.Name == W.txbxContent))
                {
                    paragraph = null;
                }

                var pi = new BlockContentInfo()
                {
                    PreviousBlockContentElement = prev,
                    ThisBlockContentElement = paragraph
                };
                content.AddAnnotation(pi);
                prev = content;
            }
        }

        public static BlockContentInfo GetParagraphInfo(this XElement contentElement)
        {
            var paragraphInfo = contentElement.Annotation<BlockContentInfo>();
            if (paragraphInfo != null)
            {
                return paragraphInfo;
            }

            InitializeParagraphInfo(contentElement.Parent);
            return contentElement.Annotation<BlockContentInfo>();
        }

        public static IEnumerable<XElement> ContentElementsBeforeSelf(this XElement element)
        {
            var current = element;
            while (true)
            {
                var pi = current.GetParagraphInfo();
                if (pi.PreviousBlockContentElement == null)
                {
                    yield break;
                }

                yield return pi.PreviousBlockContentElement;
                current = pi.PreviousBlockContentElement;
            }
        }
    }

    // Markup that this code processes:
//
// delText
//   Method: AcceptAllOtherRevisionsTransform
//   Sample document: MovedText.docx
//   Reviewed: zeyad ***************************
//   Semantics:
//     Remove these elements.
//   Reject:
//     Transform to w:t element
//
// del (deleted run content)
//   Method: AcceptAllOtherRevisionsTransform
//   Reviewed: zeyad ***************************
//   Semantics:
//     Remove these elements and descendant elements.
//   Reject:
//     Transform to w:ins element
//     Then Accept
//
// ins (inserted run content)
//   Method: AcceptAllOtherRevisionsTransform
//   Sample document: InsertedParagraphsAndRuns.docx
//   Reviewed: zeyad ***************************
//   Semantics:
//     Collapse these elements.
//   Reject:
//     Transform to w:del element, and child w:t transform to w:delText element
//     Then Accept
//
// ins (inserted paragraph)
//   Method: AcceptAllOtherRevisionsTransform
//   Sample document: InsertedParagraphsAndRuns.docx
//   Reviewed: zeyad ***************************
//   Semantics:
//     Remove these elements.
//   Reject:
//     Transform to w:del element
//     Then Accept
//
// del (deleted paragraph mark)
//   Method: AcceptDeletedAndMoveFromParagraphMarksTransform
//   Sample document: VariousTableRevisions.docx (deleted paragraph mark in paragraph in
//     content control)
//   Reviewed: tristan and zeyad ****************************************
//   Semantics:
//     Find all adjacent paragraps that have this element.
//     Group adjacent paragraphs plus the paragraph following paragraph that has this element.
//     Replace grouped paragraphs with a new paragraph containing the content from all grouped
//       paragraphs.  Use the paragraph properties from the first paragraph in the group.
//   Reject:
//     Transform to w:ins element
//     Then Accept
//
// del (deleted table row)
//   Method: AcceptAllOtherRevisionsTransform
//   Sample document: VariousTableRevisions.docx
//   Reviewed: zeyad ***************************
//   Semantics:
//     Match w:tr/w:trPr/w:del, remove w:tr.
//   Reject:
//     Transform to w:ins
//     Then Accept
//
// ins (inserted table row)
//   Method: AcceptAllOtherRevisionsTransform
//   Sample document: VariousTableRevisions.docx
//   Reviewed: zeyad ***************************
//   Semantics:
//     Remove these elements.
//   Reject:
//     Transform to w:del
//     Then Accept
//
// del (deleted math control character)
//   Method: AcceptAllOtherRevisionsTransform
//   Sample document: DeletedMathControlCharacter.docx
//   Reviewed: zeyad ***************************
//   Semantics:
//     Match m:f/m:fPr/m:ctrlPr/w:del, remove m:f.
//   Reject:
//     Transform to w:ins
//     Then Accept
//
// ins (inserted math control character)
//   Method: AcceptAllOtherRevisionsTransform
//   Sample document: InsertedMathControlCharacter.docx
//   Reviewed: zeyad ***************************
//   Semantics:
//     Remove these elements.
//   Reject:
//     Transform to w:del
//     Then Accept
//
// moveTo (move destination paragraph mark)
//   Method: AcceptMoveFromMoveToTransform
//   Sample document: MovedText.docx
//   Reviewed: zeyad ***************************
//   Semantics:
//     Remove these elements.
//   Reject:
//     Transform to moveFrom
//     Then Accept
//
// moveTo (move destination run content)
//   Method: AcceptMoveFromMoveToTransform
//   Sample document: MovedText.docx
//   Reviewed: zeyad ***************************
//   Semantics:
//     Collapse these elements.
//   Reject:
//     Transform to moveFrom
//     Then Accept
//
// moveFrom (move source paragraph mark)
//   Methods: AcceptDeletedAndMoveFromParagraphMarksTransform, AcceptParagraphEndTagsInMoveFromTransform
//   Sample document: MovedText.docx
//   Reviewed: tristan and zeyad ****************************************
//   Semantics:
//     Find all adjacent paragraps that have this element or deleted paragraph mark.
//     Group adjacent paragraphs plus the paragraph following paragraph that has this element.
//     Replace grouped paragraphs with a new paragraph containing the content from all grouped
//       paragraphs.
//     This is handled in the same code that handles del (deleted paragraph mark).
//   Reject:
//     Transform to moveTo
//     Then Accept
//
// moveFrom (move source run content)
//   Method: AcceptMoveFromMoveToTransform
//   Sample document: MovedText.docx
//   Reviewed: zeyad ***************************
//   Semantics:
//     Remove these elements.
//   Reject:
//     Transform to moveTo
//     Then Accept
//
// moveFromRangeStart
// moveFromRangeEnd
//   Method: AcceptMoveFromRanges
//   Sample document: MovedText.docx
//   Semantics:
//     Find pairs of elements.  Remove all elements that have both start and end tags in a
//       range.
//   Reject:
//     Transform to moveToRangeStart, moveToRangeEnd
//     Then Accept
//
// moveToRangeStart
// moveToRangeEnd
//   Method: AcceptAllOtherRevisionsTransform
//   Sample document: MovedText.docx
//   Reviewed: zeyad ***************************
//   Semantics:
//     Remove these elements.
//   Reject:
//     Transform to moveFromRangeStart, moveFromRangeEnd
//     Then Accept
//
// customXmlDelRangeStart
// customXmlDelRangeEnd
// customXmlMoveFromRangeStart
// customXmlMoveFromRangeEnd
//   Method: AcceptDeletedAndMovedFromContentControls
//   Reviewed: tristan and zeyad ****************************************
//   Semantics:
//     Find pairs of start/end elements, matching id attributes.  Collapse sdt
//       elements that have both start and end tags in a range.
//   Reject:
//     Transform to customXmlInsRangeStart, customXmlInsRangeEnd, customXmlMoveToRangeStart, customXmlMoveToRangeEnd
//     Then Accept
//
// customXmlInsRangeStart
// customXmlInsRangeEnd
// customXmlMoveToRangeStart
// customXmlMoveToRangeEnd
//   Method: AcceptAllOtherRevisionsTransform
//   Reviewed: tristan and zeyad ****************************************
//   Semantics:
//     Remove these elements.
//   Reject:
//     Transform to customXmlDelRangeStart, customXmlDelRangeEnd, customXmlMoveFromRangeStart, customXmlMoveFromRangeEnd
//     Then Accept
//
// delInstrText (deleted field code)
//   Method: AcceptAllOtherRevisionsTransform
//   Sample document: NumberingParagraphPropertiesChange.docx
//   Reviewed: zeyad ***************************
//   Semantics:
//     Remove these elements.
//   Reject:
//     Transform to instrText
//     Then Accept
//     Note that instrText must be transformed to delInstrText when in a w:ins, in the same fashion that w:t must be transformed to w:delText when in w:ins
//
// ins (inserted numbering properties)
//   Method: AcceptAllOtherRevisionsTransform
//   Sample document: InsertedNumberingProperties.docx
//   Reviewed: zeyad ***************************
//   Semantics:
//     Remove these elements.
//   Reject
//     Remove the containing w:numPr
//
// pPrChange (revision information for paragraph properties)
//   Method: AcceptAllOtherRevisionsTransform
//   Sample document: ParagraphAndRunPropertyRevisions.docx
//   Reviewed: zeyad ***************************
//   Semantics:
//     Remove these elements.
//   Reject:
//     Replace pPr with the pPr in pPrChange
//
// rPrChange (revision information for run properties)
//   Method: AcceptAllOtherRevisionsTransform
//   Sample document: ParagraphAndRunPropertyRevisions.docx
//   Sample document: VariousTableRevisions.docx
//   Reviewed: zeyad ***************************
//   Semantics:
//     Remove these elements.
//   Reject:
//     Replace rPr with the rPr in rPrChange
//
// rPrChange (revision information for run properties on the paragraph mark)
//   Method: AcceptAllOtherRevisionsTransform
//   Sample document: ParagraphAndRunPropertyRevisions.docx
//   Reviewed: zeyad ***************************
//   Semantics:
//     Remove these elements.
//   Reject:
//     Replace rPr with the rPr in rPrChange.
//
// numberingChange (previous numbering field properties)
//   Method: AcceptAllOtherRevisionsTransform
//   Sample document: NumberingFieldPropertiesChange.docx
//   Semantics:
//     Remove these elements.
//   Reject:
//     Remove these elements.
//     These are there for numbering created via fields, and are not important.
//
// numberingChange (previous paragraph numbering properties)
//   Method: AcceptAllOtherRevisionsTransform
//   Sample document: NumberingFieldPropertiesChange.docx
//   Semantics:
//     Remove these elements.
//   Reject:
//     Remove these elements.
//
// sectPrChange
//   Method: AcceptAllOtherRevisionsTransform
//   Sample document: SectionPropertiesChange.docx
//   Reviewed: zeyad ***************************
//   Semantics:
//     Remove these elements.
//   Reject:
//     Replace sectPr with the sectPr in sectPrChange
//
// tblGridChange
//   Method: AcceptAllOtherRevisionsTransform
//   Sample document: TableGridChange.docx
//   Sample document: VariousTableRevisions.docx
//   Reviewed: zeyad ***************************
//   Semantics:
//     Remove these elements.
//   Reject:
//     Replace tblGrid with the tblGrid in tblGridChange
//
// tblPrChange
//   Method: AcceptAllOtherRevisionsTransform
//   Sample document: TableGridChange.docx
//   Sample document: VariousTableRevisions.docx
//   Reviewed: zeyad ***************************
//   Semantics:
//     Remove these elements.
//   Reject:
//     Replace tblPr with the tblPr in tblPrChange
//
// tblPrExChange
//   Method: AcceptAllOtherRevisionsTransform
//   Sample document: VariousTableRevisions.docx
//   Reviewed: zeyad ***************************
//   Semantics:
//     Remove these elements.
//   Reject:
//     Replace tblPrEx with the tblPrEx in tblPrExChange
//
// tcPrChange
//   Method: AcceptAllOtherRevisionsTransform
//   Sample document: TableGridChange.docx
//   Sample document: VariousTableRevisions.docx
//   Reviewed: zeyad ***************************
//   Semantics:
//     Remove these elements.
//   Reject:
//     Replace tcPr with the tcPr in tcPrChange
//
// trPrChange
//   Method: AcceptAllOtherRevisionsTransform
//   Sample document: VariousTableRevisions.docx
//   Reviewed: zeyad ***************************
//   Semantics:
//     Remove these elements.
//   Reject:
//     Replace trPr with the trPr in trPrChange
//
// celDel
//   Method: AcceptDeletedCellsTransform
//   Sample document: HorizontallyMergedCells.docx
//   Semantics:
//     Group consecutive deleted cells, and remove them.
//     Adjust the cell before deleted cells:
//       Increase gridSpan by the number of deleted cells that are removed.
//   Reject:
//     Remove this element
//
// celIns
//   Method: AcceptAllOtherRevisionsTransform
//   Sample document: HorizontallyMergedCells11.docx
//   Semantics:
//     Remove these elements.
//   Reject:
//     If a w:tc contains w:tcPr/w:cellIns, then remove the cell
//
// cellMerge
//   Method: AcceptAllOtherRevisionsTransform
//   Sample document: MergedCell.docx
//   Semantics:
//     Transform cellMerge with a parent of tcPr, with attribute w:vMerge="rest"
//       to <w:vMerge w:val="restart"/>.
//     Transform cellMerge with a parent of tcPr, with attribute w:vMerge="cont"
//       to <w:vMerge w:val="continue"/>
//
// The following items need to be addressed in a future release:
// - inserted run inside deleted paragraph - moveTo is same as insert
// - must increase w:val attribute of the w:gridSpan element of the
//   cell immediately preceding the group of deleted cells by the
//   ***sum*** of the values of the w:val attributes of w:gridSpan
//   elements of each of the deleted cells.
}
