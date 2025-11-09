using DocumentAssembler.Core.FontMetric;
using DocumentFormat.OpenXml.Packaging;
using SkiaSharp;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace DocumentAssembler.Core
{
    public static class PtOpenXmlExtensions
    {
        public static XDocument GetXDocument(this OpenXmlPart part)
        {
            if (part == null)
            {
                throw new ArgumentNullException("part");
            }

            var partXDocument = part.Annotation<XDocument>();
            if (partXDocument != null)
            {
                return partXDocument;
            }

            using (var partStream = part.GetStream())
            {
                if (partStream.Length == 0)
                {
                    partXDocument = new XDocument
                    {
                        Declaration = new XDeclaration("1.0", "UTF-8", "yes")
                    };
                }
                else
                {
                    using var partXmlReader = XmlReader.Create(partStream);
                    partXDocument = XDocument.Load(partXmlReader);
                }
            }

            part.AddAnnotation(partXDocument);
            return partXDocument;
        }

        public static XDocument GetXDocument(this OpenXmlPart part, out XmlNamespaceManager? namespaceManager)
        {
            if (part == null)
            {
                throw new ArgumentNullException("part");
            }

            namespaceManager = part.Annotation<XmlNamespaceManager>();
            var partXDocument = part.Annotation<XDocument>();
            if (partXDocument != null)
            {
                if (namespaceManager != null)
                {
                    return partXDocument;
                }

                namespaceManager = GetManagerFromXDocument(partXDocument);
                part.AddAnnotation(namespaceManager);

                return partXDocument;
            }

            using var partStream = part.GetStream();
            if (partStream.Length == 0)
            {
                partXDocument = new XDocument
                {
                    Declaration = new XDeclaration("1.0", "UTF-8", "yes")
                };

                part.AddAnnotation(partXDocument);

                return partXDocument;
            }
            else
            {
                using var partXmlReader = XmlReader.Create(partStream);
                partXDocument = XDocument.Load(partXmlReader);
                namespaceManager = new XmlNamespaceManager(partXmlReader.NameTable);

                part.AddAnnotation(partXDocument);
                part.AddAnnotation(namespaceManager);

                return partXDocument;
            }
        }

        public static void PutXDocument(this OpenXmlPart part)
        {
            if (part == null)
            {
                throw new ArgumentNullException("part");
            }

            var partXDocument = part.GetXDocument();
            if (partXDocument != null)
            {
                using var partStream = part.GetStream(FileMode.Create, FileAccess.Write);
                using var partXmlWriter = XmlWriter.Create(partStream);
                partXDocument.Save(partXmlWriter);
            }
        }

        public static void PutXDocument(this OpenXmlPart part, XDocument document)
        {
            if (part == null)
            {
                throw new ArgumentNullException("part");
            }

            if (document == null)
            {
                throw new ArgumentNullException("document");
            }

            using (var partStream = part.GetStream(FileMode.Create, FileAccess.Write))
            using (var partXmlWriter = XmlWriter.Create(partStream))
            {
                document.Save(partXmlWriter);
            }

            part.RemoveAnnotations<XDocument>();
            part.AddAnnotation(document);
        }

        private static XmlNamespaceManager GetManagerFromXDocument(XDocument xDocument)
        {
            var reader = xDocument.CreateReader();
            var newXDoc = XDocument.Load(reader);

            var rootElement = xDocument.Elements().FirstOrDefault();
            if (rootElement != null && newXDoc.Root != null)
            {
                rootElement.ReplaceWith(newXDoc.Root);
            }

            var nameTable = reader.NameTable;
            var namespaceManager = new XmlNamespaceManager(nameTable);
            return namespaceManager;
        }

        public static IEnumerable<OpenXmlPart?> ContentParts(this WordprocessingDocument doc)
        {
            yield return doc.MainDocumentPart;

            foreach (var hdr in doc.MainDocumentPart?.HeaderParts ?? new HeaderPart[0])
            {
                yield return hdr;
            }

            foreach (var ftr in doc.MainDocumentPart?.FooterParts ?? new FooterPart[0])
            {
                yield return ftr;
            }

            if (doc.MainDocumentPart?.FootnotesPart != null)
            {
                yield return doc.MainDocumentPart.FootnotesPart;
            }

            if (doc.MainDocumentPart?.EndnotesPart != null)
            {
                yield return doc.MainDocumentPart.EndnotesPart;
            }
        }

    }

    public static class XmlUtil
    {
        public static XAttribute? GetXmlSpaceAttribute(string value)
        {
            return value.Length > 0 && (value[0] == ' ' || value[value.Length - 1] == ' ')
                ? new XAttribute(XNamespace.Xml + "space", "preserve")
                : null;
        }

        public static XAttribute? GetXmlSpaceAttribute(char value)
        {
            return value == ' ' ? new XAttribute(XNamespace.Xml + "space", "preserve") : null;
        }
    }

    public static class WordprocessingMLUtil
    {
        private static readonly HashSet<string> UnknownFonts = new HashSet<string>();

        private static readonly List<XName> AdditionalRunContainerNames = new List<XName>
        {
            W.w + "bdo",
            W.customXml,
            W.dir,
            W.fldSimple,
            W.hyperlink,
            W.moveFrom,
            W.moveTo,
            W.sdtContent
        };

        public static XElement CoalesceAdjacentRunsWithIdenticalFormatting(XElement runContainer)
        {
            const string dontConsolidate = "DontConsolidate";

            var groupedAdjacentRunsWithIdenticalFormatting =
                runContainer
                    .Elements()
                    .GroupAdjacent(ce =>
                    {
                        if (ce.Name == W.r)
                        {
                            if (ce.Elements().Count(e => e.Name != W.rPr) != 1)
                            {
                                return dontConsolidate;
                            }

                            if (ce.Attribute(PtOpenXml.AbstractNumId) != null)
                            {
                                return dontConsolidate;
                            }

                            var rPr = ce.Element(W.rPr);
                            var rPrString = rPr != null ? rPr.ToString(SaveOptions.None) : string.Empty;

                            if (ce.Element(W.t) != null)
                            {
                                return "Wt" + rPrString;
                            }

                            if (ce.Element(W.instrText) != null)
                            {
                                return "WinstrText" + rPrString;
                            }

                            return dontConsolidate;
                        }

                        if (ce.Name == W.ins)
                        {
                            if (ce.Elements(W.del).Any())
                            {
                                return dontConsolidate;
                            }

                            // w:ins/w:r/w:t
                            if (ce.Elements().Elements().Count(e => e.Name != W.rPr) != 1 ||
                                !ce.Elements().Elements(W.t).Any())
                            {
                                return dontConsolidate;
                            }

                            var dateIns2 = ce.Attribute(W.date);

                            var authorIns2 = (string?)ce.Attribute(W.author) ?? string.Empty;
                            var dateInsString2 = dateIns2 != null
                                ? ((DateTime)dateIns2).ToString("s")
                                : string.Empty;

                            var idIns2 = (string?)ce.Attribute(W.id) ?? string.Empty;

                            return "Wins2" +
                                   authorIns2 +
                                   dateInsString2 +
                                   idIns2 +
                                   ce.Elements()
                                       .Elements(W.rPr)
                                       .Select(rPr => rPr.ToString(SaveOptions.None))
                                       .StringConcatenate();
                        }

                        if (ce.Name == W.del)
                        {
                            if (ce.Elements(W.r).Elements().Count(e => e.Name != W.rPr) != 1 ||
                                !ce.Elements().Elements(W.delText).Any())
                            {
                                return dontConsolidate;
                            }

                            var dateDel2 = ce.Attribute(W.date);

                            var authorDel2 = (string?)ce.Attribute(W.author) ?? string.Empty;
                            var dateDelString2 = dateDel2 != null ? ((DateTime)dateDel2).ToString("s") : string.Empty;

                            return "Wdel" +
                                   authorDel2 +
                                   dateDelString2 +
                                   ce.Elements(W.r)
                                       .Elements(W.rPr)
                                       .Select(rPr => rPr.ToString(SaveOptions.None))
                                       .StringConcatenate();
                        }

                        return dontConsolidate;
                    });

            var runContainerWithConsolidatedRuns = new XElement(runContainer.Name,
                runContainer.Attributes(),
                groupedAdjacentRunsWithIdenticalFormatting.Select(g =>
                {
                    if (g.Key == dontConsolidate)
                    {
                        return (object)g;
                    }

                    var textValue = g
                        .Select(r =>
                            r.Descendants()
                                .Where(d => d.Name == W.t || d.Name == W.delText || d.Name == W.instrText)
                                .Select(d => d.Value)
                                .StringConcatenate())
                        .StringConcatenate();
                    var xs = XmlUtil.GetXmlSpaceAttribute(textValue);

                    if (g.First().Name == W.r)
                    {
                        if (g.First().Element(W.t) != null)
                        {
                            var statusAtt =
                                g.Select(r => r.Descendants(W.t).Take(1).Attributes(PtOpenXml.Status));
                            return new XElement(W.r,
                                g.First().Attributes(),
                                g.First().Elements(W.rPr),
                                new XElement(W.t, statusAtt, xs, textValue));
                        }

                        if (g.First().Element(W.instrText) != null)
                        {
                            return new XElement(W.r,
                                g.First().Attributes(),
                                g.First().Elements(W.rPr),
                                new XElement(W.instrText, xs, textValue));
                        }
                    }

                    if (g.First().Name == W.ins)
                    {
                        var firstR = g.First().Element(W.r);
                        return new XElement(W.ins,
                            g.First().Attributes(),
                            new XElement(W.r,
                                firstR?.Attributes(),
                                g.First().Elements(W.r).Elements(W.rPr),
                                new XElement(W.t, xs, textValue)));
                    }

                    if (g.First().Name == W.del)
                    {
                        var firstR = g.First().Element(W.r);
                        return new XElement(W.del,
                            g.First().Attributes(),
                            new XElement(W.r,
                                firstR?.Attributes(),
                                g.First().Elements(W.r).Elements(W.rPr),
                                new XElement(W.delText, xs, textValue)));
                    }
                    return g;
                }));

            // Process w:txbxContent//w:p
            foreach (var txbx in runContainerWithConsolidatedRuns.Descendants(W.txbxContent))
            {
                foreach (var txbxPara in txbx.DescendantsTrimmed(W.txbxContent).Where(d => d.Name == W.p))
                {
                    var newPara = CoalesceAdjacentRunsWithIdenticalFormatting(txbxPara);
                    txbxPara.ReplaceWith(newPara);
                }
            }

            // Process additional run containers.
            var runContainers = runContainerWithConsolidatedRuns
                .Descendants()
                .Where(d => AdditionalRunContainerNames.Contains(d.Name))
                .ToList();
            foreach (var container in runContainers)
            {
                var newContainer = CoalesceAdjacentRunsWithIdenticalFormatting(container);
                container.ReplaceWith(newContainer);
            }

            return runContainerWithConsolidatedRuns;
        }
    }

}
