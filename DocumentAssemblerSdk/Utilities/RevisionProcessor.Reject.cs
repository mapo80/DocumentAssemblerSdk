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
        private static void RejectRevisionsForPart(OpenXmlPart part)
        {
            var xDoc = part.GetXDocument();
            var newRoot = RejectRevisionsForPartTransform(xDoc.Root);
            xDoc.Root.ReplaceWith(newRoot);
            part.PutXDocument();
        }

        private static object? RejectRevisionsForPartTransform(XNode node)
        {
            if (node is XElement element)
            {
                // Inserted Numbering Properties

                if (element.Name == W.numPr && element.Element(W.ins) != null)
                {
                    return null;
                }

                // Paragraph properties change

                if (element.Name == W.pPr &&
                    element.Element(W.pPrChange) != null)
                {
                    var pPr = element.Element(W.pPrChange).Element(W.pPr);
                    if (pPr == null)
                    {
                        pPr = new XElement(W.pPr);
                    }

                    var new_pPr = new XElement(pPr); // clone it
                    new_pPr.Add(RejectRevisionsForPartTransform(element.Element(W.rPr)));
                    return RejectRevisionsForPartTransform(new_pPr);
                }

                // Run properties change

                if (element.Name == W.rPr &&
                    element.Element(W.rPrChange) != null)
                {
                    var new_rPr = element.Element(W.rPrChange).Element(W.rPr);
                    return RejectRevisionsForPartTransform(new_rPr);
                }

                // Field code numbering change

                if (element.Name == W.numberingChange)
                {
                    return null;
                }

                // Change w:sectPr

                if (element.Name == W.sectPr &&
                    element.Element(W.sectPrChange) != null)
                {
                    var newSectPr = element.Element(W.sectPrChange).Element(W.sectPr);
                    return RejectRevisionsForPartTransform(newSectPr);
                }

                // tblGridChange

                if (element.Name == W.tblGrid &&
                    element.Element(W.tblGridChange) != null)
                {
                    var newTblGrid = element.Element(W.tblGridChange).Element(W.tblGrid);
                    return RejectRevisionsForPartTransform(newTblGrid);
                }

                // tcPrChange

                if (element.Name == W.tcPr &&
                    element.Element(W.tcPrChange) != null)
                {
                    var newTcPr = element.Element(W.tcPrChange).Element(W.tcPr);
                    return RejectRevisionsForPartTransform(newTcPr);
                }

                // trPrChange
                if (element.Name == W.trPr &&
                    element.Element(W.trPrChange) != null)
                {
                    var newTrPr = element.Element(W.trPrChange).Element(W.trPr);
                    return RejectRevisionsForPartTransform(newTrPr);
                }

                // tblPrExChange

                if (element.Name == W.tblPrEx &&
                    element.Element(W.tblPrExChange) != null)
                {
                    var newTblPrEx = element.Element(W.tblPrExChange).Element(W.tblPrEx);
                    return RejectRevisionsForPartTransform(newTblPrEx);
                }

                // tblPrChange

                if (element.Name == W.tblPr &&
                    element.Element(W.tblPrChange) != null)
                {
                    var newTrPr = element.Element(W.tblPrChange).Element(W.tblPr);
                    return RejectRevisionsForPartTransform(newTrPr);
                }

                // tblPrChange

                if (element.Name == W.cellDel ||
                    element.Name == W.cellMerge)
                {
                    return null;
                }

                if (element.Name == W.tc &&
                    element.Elements(W.tcPr).Elements(W.cellIns).Any())
                {
                    return null;
                }

                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => RejectRevisionsForPartTransform(n)));
            }
            return node;
        }

        private static void RejectRevisionsForStylesDefinitionPart(StyleDefinitionsPart stylesDefinitionsPart)
        {
            var xDoc = stylesDefinitionsPart.GetXDocument();
            var newRoot = RejectRevisionsForStylesTransform(xDoc.Root);
            xDoc.Root.ReplaceWith(newRoot);
            stylesDefinitionsPart.PutXDocument();
        }

        private static object RejectRevisionsForStylesTransform(XNode node)
        {
            if (node is XElement element)
            {
                if (element.Name == W.pPr &&
                    element.Element(W.pPrChange) != null)
                {
                    var new_pPr = element.Element(W.pPrChange).Element(W.pPr);
                    return RejectRevisionsForStylesTransform(new_pPr);
                }

                if (element.Name == W.rPr &&
                    element.Element(W.rPrChange) != null)
                {
                    var new_rPr = element.Element(W.rPrChange).Element(W.rPr);
                    return RejectRevisionsForStylesTransform(new_rPr);
                }

                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => RejectRevisionsForStylesTransform(n)));
            }
            return node;
        }

    }
}
