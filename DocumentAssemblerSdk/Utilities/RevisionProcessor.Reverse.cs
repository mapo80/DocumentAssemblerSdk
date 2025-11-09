using DocumentAssembler.Core.Exceptions;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace DocumentAssembler.Core
{
    internal class ReverseRevisionsInfo
    {
        public bool InInsert;
    }

    public partial class RevisionProcessor
    {
        private static void ReverseRevisions(WordprocessingDocument doc)
        {
            ReverseRevisionsForPart(doc.MainDocumentPart);
            foreach (var part in doc.MainDocumentPart.HeaderParts)
            {
                ReverseRevisionsForPart(part);
            }

            foreach (var part in doc.MainDocumentPart.FooterParts)
            {
                ReverseRevisionsForPart(part);
            }

            if (doc.MainDocumentPart.EndnotesPart != null)
            {
                ReverseRevisionsForPart(doc.MainDocumentPart.EndnotesPart);
            }

            if (doc.MainDocumentPart.FootnotesPart != null)
            {
                ReverseRevisionsForPart(doc.MainDocumentPart.FootnotesPart);
            }
        }

        private static void ReverseRevisionsForPart(OpenXmlPart part)
        {
            var xDoc = part.GetXDocument();
            var rri = new ReverseRevisionsInfo
            {
                InInsert = false
            };
            var newRoot = (XElement)ReverseRevisionsTransform(xDoc.Root, rri);
            newRoot = RemoveRsidTransform(newRoot) as XElement;
            xDoc.Root.ReplaceWith(newRoot);
            part.PutXDocument();
        }

        private static object? RemoveRsidTransform(XNode node)
        {
            if (node is XElement element)
            {
                if (element.Name == W.rsid)
                {
                    return null;
                }

                return new XElement(element.Name,
                    element.Attributes().Where(a => a.Name != W.rsid &&
                        a.Name != W.rsidDel &&
                        a.Name != W.rsidP &&
                        a.Name != W.rsidR &&
                        a.Name != W.rsidRDefault &&
                        a.Name != W.rsidRPr &&
                        a.Name != W.rsidSect &&
                        a.Name != W.rsidTr),
                    element.Nodes().Select(n => RemoveRsidTransform(n)));
            }
            return node;
        }

        private static object MergeAdjacentTablesTransform(XNode node)
        {
            if (node is XElement element)
            {
                if (element.Element(W.tbl) != null)
                {
                    var grouped = element
                        .Elements()
                        .GroupAdjacent(e =>
                        {
                            if (e.Name != W.tbl)
                            {
                                return "";
                            }

                            var bidiVisual = e.Elements(W.tblPr).Elements(W.bidiVisual).FirstOrDefault();
                            var bidiVisString = bidiVisual == null ? "" : "|bidiVisual";
                            var key = "tbl" + bidiVisString;
                            return key;
                        });

                    var newContent = grouped
                        .Select(g =>
                        {
                            if (g.Key == "" || g.Count() == 1)
                            {
                                return (object)g;
                            }

                            var rolled = g
                                .Select(tbl =>
                                {
                                    var gridCols = tbl
                                        .Elements(W.tblGrid)
                                        .Elements(W.gridCol)
                                        .Attributes(W._w)
                                        .Select(a => (int)a)
                                        .Rollup(0, (s, i) => s + i);
                                    return gridCols;
                                })
                                .SelectMany(m => m)
                                .Distinct()
                                .OrderBy(w => w)
                                .ToArray();
                            var newTable = new XElement(W.tbl,
                                g.First().Elements(W.tblPr),
                                new XElement(W.tblGrid,
                                    rolled.Select((r, i) =>
                                    {
                                        int v;
                                        if (i == 0)
                                        {
                                            v = r;
                                        }
                                        else
                                        {
                                            v = r - rolled[i - 1];
                                        }

                                        return new XElement(W.gridCol,
                                            new XAttribute(W._w, v));
                                    })),
                                g.Select(tbl =>
                                {
                                    var fixedWidthsTbl = FixWidths(tbl);
                                    var newRows = fixedWidthsTbl.Elements(W.tr)
                                        .Select(tr =>
                                        {
                                            var newRow = new XElement(W.tr,
                                                tr.Attributes(),
                                                tr.Elements().Where(e => e.Name != W.tc),
                                                tr.Elements(W.tc).Select(tc =>
                                                {
                                                    var w = (int?)tc
                                                        .Elements(W.tcPr)
                                                        .Elements(W.tcW)
                                                        .Attributes(W._w)
                                                        .FirstOrDefault();
                                                    if (w == null)
                                                    {
                                                        return tc;
                                                    }

                                                    var cellsToLeft = tc
                                                        .Parent
                                                        .Elements(W.tc)
                                                        .TakeWhile(btc => btc != tc);
                                                    var widthToLeft = 0;
                                                    if (cellsToLeft.Any())
                                                    {
                                                        widthToLeft = cellsToLeft
                                                        .Elements(W.tcPr)
                                                        .Elements(W.tcW)
                                                        .Attributes(W._w)
                                                        .Select(wi => (int)wi)
                                                        .Sum();
                                                    }

                                                    var rolledPairs = new[] { new
                                                        {
                                                            GridValue = 0,
                                                            Index = 0,
                                                        }}
                                                        .Concat(
                                                            rolled
                                                            .Select((r, i) => new
                                                            {
                                                                GridValue = r,
                                                                Index = i + 1,
                                                            }));
                                                    var start = rolledPairs
                                                        .FirstOrDefault(t => t.GridValue >= widthToLeft);
                                                    if (start != null)
                                                    {
                                                        var gridsRequired = rolledPairs
                                                            .Skip(start.Index)
                                                            .TakeWhile(rp => rp.GridValue - start.GridValue < w)
                                                            .Count();
                                                        var tcPr = new XElement(W.tcPr,
                                                                tc.Elements(W.tcPr).Elements().Where(e => e.Name != W.gridSpan),
                                                                gridsRequired != 1 ?
                                                                    new XElement(W.gridSpan,
                                                                        new XAttribute(W.val, gridsRequired)) :
                                                                    null);
                                                        var orderedTcPr = new XElement(W.tcPr,
                                                            tcPr.Elements().OrderBy(e =>
                                                            {
                                                                if (Order_tcPr.ContainsKey(e.Name))
                                                                {
                                                                    return Order_tcPr[e.Name];
                                                                }

                                                                return 999;
                                                            }));
                                                        var newCell = new XElement(W.tc,
                                                            orderedTcPr,
                                                            tc.Elements().Where(e => e.Name != W.tcPr));
                                                        return newCell;
                                                    }
                                                    return tc;
                                                }));
                                            return newRow;
                                        });
                                    return newRows;
                                }));
                            return newTable;
                        });
                    return new XElement(element.Name,
                        element.Attributes(),
                        newContent);
                }

                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => MergeAdjacentTablesTransform(n)));
            }
            return node;
        }

        private static object ReverseRevisionsTransform(XNode node, ReverseRevisionsInfo rri)
        {
            if (node is XElement element)
            {
                var parent = element
                    .Ancestors()
                    .FirstOrDefault(a => a.Name != W.sdtContent && a.Name != W.sdt && a.Name != W.smartTag);

                // Deleted run

                if (element.Name == W.del &&
                    parent.Name == W.p)
                {
                    return new XElement(W.ins,
                        element.Nodes().Select(n => ReverseRevisionsTransform(n, rri)));
                }

                // Deleted paragraph mark

                if (element.Name == W.del &&
                    parent.Name == W.rPr &&
                    parent.Parent.Name == W.pPr)
                {
                    return new XElement(W.ins);
                }

                // Inserted paragraph mark

                if (element.Name == W.ins &&
                    parent.Name == W.rPr &&
                    parent.Parent.Name == W.pPr)
                {
                    return new XElement(W.del);
                }

                // Inserted run

                if (element.Name == W.ins &&
                    parent.Name == W.p)
                {
                    return new XElement(W.del,
                        element.Nodes().Select(n => ReverseRevisionsTransform(n, rri)));
                }

                // Deleted table row

                if (element.Name == W.del &&
                    parent.Name == W.trPr)
                {
                    return new XElement(W.ins);
                }

                // Inserted table row

                if (element.Name == W.ins &&
                    parent.Name == W.trPr)
                {
                    return new XElement(W.del);
                }

                // Deleted math control character

                if (element.Name == W.del &&
                    parent.Name == M.r)
                {
                    return new XElement(W.ins,
                        element.Nodes().Select(n => ReverseRevisionsTransform(n, rri)));
                }

                // Inserted math control character

                if (element.Name == W.ins &&
                    parent.Name == M.r)
                {
                    return new XElement(W.del,
                        element.Nodes().Select(n => ReverseRevisionsTransform(n, rri)));
                }

                // moveFrom / moveTo

                if (element.Name == W.moveFrom)
                {
                    return new XElement(W.moveTo,
                        element.Attributes(),
                        element.Nodes().Select(n => ReverseRevisionsTransform(n, rri)));
                }
                if (element.Name == W.moveFromRangeStart)
                {
                    return new XElement(W.moveToRangeStart,
                        element.Attributes(),
                        element.Nodes().Select(n => ReverseRevisionsTransform(n, rri)));
                }
                if (element.Name == W.moveFromRangeEnd)
                {
                    return new XElement(W.moveToRangeEnd,
                        element.Attributes(),
                        element.Nodes().Select(n => ReverseRevisionsTransform(n, rri)));
                }
                if (element.Name == W.moveTo)
                {
                    return new XElement(W.moveFrom,
                        element.Attributes(),
                        element.Nodes().Select(n => ReverseRevisionsTransform(n, rri)));
                }
                if (element.Name == W.moveToRangeStart)
                {
                    return new XElement(W.moveFromRangeStart,
                        element.Attributes(),
                        element.Nodes().Select(n => ReverseRevisionsTransform(n, rri)));
                }
                if (element.Name == W.moveToRangeEnd)
                {
                    return new XElement(W.moveFromRangeEnd,
                        element.Attributes(),
                        element.Nodes().Select(n => ReverseRevisionsTransform(n, rri)));
                }

                // Deleted content control

                if (element.Name == W.customXmlDelRangeStart)
                {
                    return new XElement(W.customXmlInsRangeStart,
                        element.Attributes(),
                        element.Nodes().Select(n => ReverseRevisionsTransform(n, rri)));
                }
                if (element.Name == W.customXmlDelRangeEnd)
                {
                    return new XElement(W.customXmlInsRangeEnd,
                        element.Attributes(),
                        element.Nodes().Select(n => ReverseRevisionsTransform(n, rri)));
                }

                // Inserted content control

                if (element.Name == W.customXmlInsRangeStart)
                {
                    return new XElement(W.customXmlDelRangeStart,
                        element.Attributes(),
                        element.Nodes().Select(n => ReverseRevisionsTransform(n, rri)));
                }
                if (element.Name == W.customXmlInsRangeEnd)
                {
                    return new XElement(W.customXmlDelRangeEnd,
                        element.Attributes(),
                        element.Nodes().Select(n => ReverseRevisionsTransform(n, rri)));
                }

                // Moved content control

                if (element.Name == W.customXmlMoveFromRangeStart)
                {
                    return new XElement(W.customXmlMoveToRangeStart,
                        element.Attributes(),
                        element.Nodes().Select(n => ReverseRevisionsTransform(n, rri)));
                }
                if (element.Name == W.customXmlMoveFromRangeEnd)
                {
                    return new XElement(W.customXmlMoveToRangeEnd,
                        element.Attributes(),
                        element.Nodes().Select(n => ReverseRevisionsTransform(n, rri)));
                }
                if (element.Name == W.customXmlMoveToRangeStart)
                {
                    return new XElement(W.customXmlMoveFromRangeStart,
                        element.Attributes(),
                        element.Nodes().Select(n => ReverseRevisionsTransform(n, rri)));
                }
                if (element.Name == W.customXmlMoveToRangeEnd)
                {
                    return new XElement(W.customXmlMoveFromRangeEnd,
                        element.Attributes(),
                        element.Nodes().Select(n => ReverseRevisionsTransform(n, rri)));
                }

                // Deleted field code
                if (element.Name == W.delInstrText)
                {
                    return new XElement(W.instrText,
                        element.Attributes(), // pulls in xml:space attribute
                        element.Nodes().Select(n => ReverseRevisionsTransform(n, rri)));
                }

                // Change inserted instrText element to w:delInstrText
                if (element.Name == W.instrText && rri.InInsert)
                {
                    return new XElement(W.delInstrText,
                        element.Attributes(),
                        element.Nodes().Select(n => ReverseRevisionsTransform(n, rri)));
                }

                // Change inserted text element to w:delText
                if (element.Name == W.t && rri.InInsert)
                {
                    return new XElement(W.delText,
                        element.Attributes(),
                        element.Nodes().Select(n => ReverseRevisionsTransform(n, rri)));
                }

                // Change w:delText to w:t
                if (element.Name == W.delText)
                {
                    return new XElement(W.t,
                        element.Attributes(), // pulls in xml:space attribute
                        element.Nodes().Select(n => ReverseRevisionsTransform(n, rri)));
                }

                // Identity transform
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => ReverseRevisionsTransform(n, rri)));
            }
            return node;
        }

        private static readonly Dictionary<XName, int> Order_tcPr = new Dictionary<XName, int>
        {
            { W.cnfStyle, 10 },
            { W.tcW, 20 },
            { W.gridSpan, 30 },
            { W.hMerge, 40 },
            { W.vMerge, 50 },
            { W.tcBorders, 60 },
            { W.shd, 70 },
            { W.noWrap, 80 },
            { W.tcMar, 90 },
            { W.textDirection, 100 },
            { W.tcFitText, 110 },
            { W.vAlign, 120 },
            { W.hideMark, 130 },
            { W.headers, 140 },
        };

        private static XElement FixWidths(XElement tbl)
        {
            var newTbl = new XElement(tbl);
            var gridLines = tbl.Elements(W.tblGrid).Elements(W.gridCol).Attributes(W._w).Select(w => (int)w).ToArray();
            foreach (var tr in newTbl.Elements(W.tr))
            {
                var used = 0;
                var lastUsed = -1;
                foreach (var tc in tr.Elements(W.tc))
                {
                    var tcW = tc.Elements(W.tcPr).Elements(W.tcW).Attributes(W._w).FirstOrDefault();
                    if (tcW != null)
                    {
                        var gridSpan = (int?)tc.Elements(W.tcPr).Elements(W.gridSpan).Attributes(W.val).FirstOrDefault();

                        if (gridSpan == null)
                        {
                            gridSpan = 1;
                        }

                        var z = Math.Min(gridLines.Length - 1, lastUsed + (int)gridSpan);
                        var w = gridLines.Where((g, i) => i > lastUsed && i <= z).Sum();
                        tcW.Value = w.ToString();

                        lastUsed += (int)gridSpan;
                        used += (int)gridSpan;
                    }
                }
            }
            return newTbl;
        }

    }
}
