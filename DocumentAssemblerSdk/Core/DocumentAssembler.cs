using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using SkiaSharp;
using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Schema;
using System.Xml.XPath;

namespace DocumentAssembler.Core
{
    public partial class DocumentAssembler
    {
        public static WmlDocument AssembleDocument(WmlDocument templateDoc, XmlDocument data, out bool templateError) =>
            AssembleDocument(templateDoc, data, out templateError, out _);

        public static WmlDocument AssembleDocument(WmlDocument templateDoc, XmlDocument data, out bool templateError, out string? templateErrorSummary)
        {
            var xDoc = data.GetXDocument();
            if (xDoc.Root == null)
            {
                throw new ArgumentException("Data document does not have a root element.", nameof(data));
            }
            return AssembleDocument(templateDoc, xDoc.Root, out templateError, out templateErrorSummary);
        }

        public static WmlDocument AssembleDocument(WmlDocument templateDoc, XElement data, out bool templateError) =>
            AssembleDocument(templateDoc, data, out templateError, out _);

        public static WmlDocument AssembleDocument(WmlDocument templateDoc, XElement data, out bool templateError, out string? templateErrorSummary)
        {
            var assembledDocument = AssembleDocumentInternal(templateDoc, data, out var templateErrorDetails);
            templateError = templateErrorDetails.HasError;
            templateErrorSummary = templateErrorDetails.GetErrorSummary();
            return assembledDocument;
        }

        private static WmlDocument AssembleDocumentInternal(WmlDocument templateDoc, XElement data, out TemplateError templateErrorDetails)
        {
            var byteArray = templateDoc.DocumentByteArray;
            using var mem = new MemoryStream();
            mem.Write(byteArray, 0, byteArray.Length);
            var te = new TemplateError();
            using (var wordDoc = WordprocessingDocument.Open(mem, true))
            {
                if (RevisionAccepter.HasTrackedRevisions(wordDoc))
                {
                    throw new OpenXmlPowerToolsException("Invalid DocumentAssembler template - contains tracked revisions");
                }

                var evaluationContext = new XPathEvaluationContext();
                foreach (var part in wordDoc.ContentParts())
                {
                    if (part != null)
                    {
                        ProcessTemplatePart(data, te, part, evaluationContext);
                    }
                }
            }
            templateErrorDetails = te;
            var assembledDocument = new WmlDocument("TempFileName.docx", mem.ToArray());
            return assembledDocument;
        }

        private static void ProcessTemplatePart(XElement data, TemplateError te, OpenXmlPart part, XPathEvaluationContext evaluationContext)
        {
            var xDoc = part.GetXDocument();
            if (xDoc.Root == null)
            {
                return;
            }

            var xDocRoot = RemoveGoBackBookmarks(xDoc.Root);

            // content controls in cells can surround the W.tc element, so transform so that such content controls are within the cell content
            xDocRoot = (XElement)NormalizeContentControlsInCells(xDocRoot);

            xDocRoot = (XElement)TransformToMetadata(xDocRoot, data, te);

            // Table might have been placed at run-level, when it should be at block-level, so fix this.
            // Repeat, EndRepeat, Conditional, EndConditional are allowed at run level, but only if there is a matching pair
            // if there is only one Repeat, EndRepeat, Conditional, EndConditional, then move to block level.
            // if there is a matching pair, then is OK.
            xDocRoot = (XElement)ForceBlockLevelAsAppropriate(xDocRoot, te);

            NormalizeTablesRepeatAndConditional(xDocRoot, te);

            // do the actual content replacement
            xDocRoot = ContentReplacementTransform(xDocRoot, data, te, part, evaluationContext) as XElement;

            // Note: Error collection is done during processing. Errors are indicated by:
            // 1. The templateError boolean flag (te.HasError)
            // 2. Inline error placeholders in the document (e.g., "[ERROR: Missing field]")
            // 3. The comprehensive error summary available via te.GetErrorSummary()
            //
            // We don't insert the error summary as paragraphs to avoid schema validation issues.
            // Users can access the full error list programmatically via the TemplateError object.

            xDoc.Elements().First().ReplaceWith(xDocRoot);
            part.PutXDocument();
            return;
        }

        private static readonly XName[] s_MetaToForceToBlock = new XName[] {
            PA.Conditional,
            PA.EndConditional,
            PA.Repeat,
            PA.EndRepeat,
            PA.Table,
        };

        private static object ForceBlockLevelAsAppropriate(XNode node, TemplateError te)
        {
            if (node is XElement element)
            {
                if (element.Name == W.p)
                {
                    var childMeta = element.Elements().Where(n => s_MetaToForceToBlock.Contains(n.Name)).ToList();
                    if (childMeta.Count == 1)
                    {
                        var child = childMeta.First();
                        var otherTextInParagraph = element.Elements(W.r).Elements(W.t).Select(t => (string)t).StringConcatenate().Trim();
                        if (otherTextInParagraph != "")
                        {
                            var newPara = new XElement(element);
                            var newMeta = newPara.Elements().First(n => s_MetaToForceToBlock.Contains(n.Name));
                            newMeta.ReplaceWith(CreateRunErrorMessage("Error: Unmatched metadata can't be in paragraph with other text", te));
                            return newPara;
                        }
                        var meta = new XElement(child.Name,
                            child.Attributes(),
                            new XElement(W.p,
                                element.Attributes(),
                                element.Elements(W.pPr),
                                child.Elements()));
                        return meta;
                    }
                    var count = childMeta.Count;
                    if (count % 2 == 0)
                    {
                        if (childMeta.Where(c => c.Name == PA.Repeat).Count() != childMeta.Where(c => c.Name == PA.EndRepeat).Count())
                        {
                            return CreateContextErrorMessage(element, "Error: Mismatch Repeat / EndRepeat at run level", te);
                        }

                        if (childMeta.Where(c => c.Name == PA.Conditional).Count() != childMeta.Where(c => c.Name == PA.EndConditional).Count())
                        {
                            return CreateContextErrorMessage(element, "Error: Mismatch Conditional / EndConditional at run level", te);
                        }

                        return new XElement(element.Name,
                            element.Attributes(),
                            element.Nodes().Select(n => ForceBlockLevelAsAppropriate(n, te)));
                    }
                    else
                    {
                        return CreateContextErrorMessage(element, "Error: Invalid metadata at run level", te);
                    }
                }
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => ForceBlockLevelAsAppropriate(n, te)));
            }
            return node;
        }

        private static XElement RemoveGoBackBookmarks(XElement xElement)
        {
            var cloneXDoc = new XElement(xElement);
            while (true)
            {
                var bm = cloneXDoc.DescendantsAndSelf(W.bookmarkStart).FirstOrDefault(b => (string?)b.Attribute(W.name) == "_GoBack");
                if (bm == null)
                {
                    break;
                }

                var id = (string?)bm.Attribute(W.id);
                if (id != null)
                {
                    var endBm = cloneXDoc.DescendantsAndSelf(W.bookmarkEnd).FirstOrDefault(b => (string?)b.Attribute(W.id) == id);
                    bm.Remove();
                    endBm?.Remove();
                }
                else
                {
                    bm.Remove();
                }
            }
            return cloneXDoc;
        }

        // this transform inverts content controls that surround W.tc elements.  After transforming, the W.tc will contain
        // the content control, which contains the paragraph content of the cell.
        private static object NormalizeContentControlsInCells(XNode node)
        {
            if (node is XElement element)
            {
                if (element.Name == W.sdt && element.Parent?.Name == W.tr)
                {
                    var newCell = new XElement(W.tc,
                        element.Elements(W.tc).Elements(W.tcPr),
                        new XElement(W.sdt,
                            element.Elements(W.sdtPr),
                            element.Elements(W.sdtEndPr),
                            new XElement(W.sdtContent,
                                element.Elements(W.sdtContent).Elements(W.tc).Elements().Where(e => e.Name != W.tcPr))));
                    return newCell;
                }
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => NormalizeContentControlsInCells(n)));
            }
            return node;
        }

        // The following method is written using tree modification, not RPFT, because it is easier to write in this fashion.
        // These types of operations are not as easy to write using RPFT.
        // Unless you are completely clear on the semantics of LINQ to XML DML, do not make modifications to this method.
        private static void NormalizeTablesRepeatAndConditional(XElement xDoc, TemplateError te)
        {
            var tables = xDoc.Descendants(PA.Table).ToList();
            foreach (var table in tables)
            {
                var followingElement = table.ElementsAfterSelf().FirstOrDefault(e => e.Name == W.tbl || e.Name == W.p);
                if (followingElement == null || followingElement.Name != W.tbl)
                {
                    table.ReplaceWith(CreateParaErrorMessage("Table metadata is not immediately followed by a table", te));
                    continue;
                }
                // remove superflous paragraph from Table metadata
                table.RemoveNodes();
                // detach w:tbl from parent, and add to Table metadata
                followingElement.Remove();
                table.Add(followingElement);
            }

            var repeatDepth = 0;
            var conditionalDepth = 0;
            foreach (var metadata in xDoc.Descendants().Where(d =>
                    d.Name == PA.Repeat ||
                    d.Name == PA.Conditional ||
                    d.Name == PA.EndRepeat ||
                    d.Name == PA.EndConditional ||
                    d.Name == PA.Else))
            {
                if (metadata.Name == PA.Repeat)
                {
                    ++repeatDepth;
                    metadata.Add(new XAttribute(PA.Depth, repeatDepth));
                    continue;
                }
                if (metadata.Name == PA.EndRepeat)
                {
                    metadata.Add(new XAttribute(PA.Depth, repeatDepth));
                    --repeatDepth;
                    continue;
                }
                if (metadata.Name == PA.Conditional)
                {
                    ++conditionalDepth;
                    metadata.Add(new XAttribute(PA.Depth, conditionalDepth));
                    continue;
                }
                if (metadata.Name == PA.Else)
                {
                    // Else is at the same depth as its containing Conditional
                    metadata.Add(new XAttribute(PA.Depth, conditionalDepth));
                    continue;
                }
                if (metadata.Name == PA.EndConditional)
                {
                    metadata.Add(new XAttribute(PA.Depth, conditionalDepth));
                    --conditionalDepth;
                    continue;
                }
            }

            while (true)
            {
                var didReplace = false;
                foreach (var metadata in xDoc.Descendants().Where(d => (d.Name == PA.Repeat || d.Name == PA.Conditional) && d.Attribute(PA.Depth) != null).ToList())
                {
                    var depthAttr = metadata.Attribute(PA.Depth);
                    if (depthAttr == null)
                    {
                        continue;
                    }
                    var depth = (int)depthAttr;
                    XName? matchingEndName = null;
                    if (metadata.Name == PA.Repeat)
                    {
                        matchingEndName = PA.EndRepeat;
                    }
                    else if (metadata.Name == PA.Conditional)
                    {
                        matchingEndName = PA.EndConditional;
                    }

                    if (matchingEndName == null)
                    {
                        throw new OpenXmlPowerToolsException("Internal error");
                    }

                    var matchingEnd = metadata.ElementsAfterSelf(matchingEndName).FirstOrDefault(end => {
                        var endDepthAttr = end.Attribute(PA.Depth);
                        return endDepthAttr != null && (int)endDepthAttr == depth;
                    });
                    if (matchingEnd == null)
                    {
                        metadata.ReplaceWith(CreateParaErrorMessage(string.Format("{0} does not have matching {1}", metadata.Name.LocalName, matchingEndName.LocalName), te));
                        continue;
                    }
                    metadata.RemoveNodes();
                    var contentBetween = metadata.ElementsAfterSelf().TakeWhile(after => after != matchingEnd).ToList();
                    foreach (var item in contentBetween)
                    {
                        item.Remove();
                    }

                    contentBetween = contentBetween.Where(n => n.Name != W.bookmarkStart && n.Name != W.bookmarkEnd).ToList();

                    // Handle Else within Conditional blocks
                    if (metadata.Name == PA.Conditional)
                    {
                        // Find the Else element at the same depth level as this Conditional
                        var elseElement = contentBetween.FirstOrDefault(e => {
                            if (e.Name != PA.Else) return false;
                            var elseDepthAttr = e.Attribute(PA.Depth);
                            return elseDepthAttr != null && (int)elseDepthAttr == depth;
                        });
                        if (elseElement != null)
                        {
                            var indexOfElse = contentBetween.IndexOf(elseElement);
                            var beforeElse = contentBetween.Take(indexOfElse).ToList();
                            var afterElse = contentBetween.Skip(indexOfElse + 1).ToList();

                            // Add content before Else directly to Conditional
                            metadata.Add(beforeElse);

                            // Update Else element to contain content after it
                            elseElement.RemoveNodes();
                            elseElement.Add(afterElse);

                            // Remove Depth attribute from Else
                            elseElement.Attributes(PA.Depth).Remove();

                            // Add Else as a child of Conditional
                            metadata.Add(elseElement);
                        }
                        else
                        {
                            // No Else, just add all content to Conditional
                            metadata.Add(contentBetween);
                        }
                    }
                    else
                    {
                        metadata.Add(contentBetween);
                    }

                    metadata.Attributes(PA.Depth).Remove();
                    matchingEnd.Remove();
                    didReplace = true;
                    break;
                }
                if (!didReplace)
                {
                    break;
                }
            }

            foreach (var element in xDoc.Descendants(PA.EndRepeat).ToList())
            {
                var error = CreateContextErrorMessage(element, "Error: EndRepeat without matching Repeat", te);
                element.ReplaceWith(error);
            }
            foreach (var element in xDoc.Descendants(PA.EndConditional).ToList())
            {
                var error = CreateContextErrorMessage(element, "Error: EndConditional without matching Conditional", te);
                element.ReplaceWith(error);
            }
        }

        private static object? ContentReplacementTransform(XNode node, XElement data, TemplateError templateError, OpenXmlPart owningPart, XPathEvaluationContext evaluationContext)
        {
            if (node is XElement element)
            {
                if (element.Name == PA.Content)
                {
                    var para = element.Descendants(W.p).FirstOrDefault();
                    var run = element.Descendants(W.r).FirstOrDefault();

                    var xPath = (string?)element.Attribute(PA.Select);
                    if (xPath == null)
                    {
                        return CreateContextErrorMessage(element, "Content: Select attribute is required", templateError);
                    }

                    var optionalString = (string?)element.Attribute(PA.Optional);
                    // Default is true (optional by default), unless explicitly set to false
                    var optional = optionalString == null || !bool.TryParse(optionalString, out var optionalValue) || optionalValue;

                    // EvaluateXPathToString now collects errors instead of throwing
                    string newValue = EvaluateXPathToString(data, xPath, optional, templateError, evaluationContext);

                    if (para != null)
                    {
                        var template = GetParagraphTemplate(para, run ?? para.Elements(W.r).FirstOrDefault());
                        var p = new XElement(W.p,
                            template.ParagraphProperties != null ? new XElement(template.ParagraphProperties) : null);
                        var firstLine = true;
                        foreach (var line in newValue.Split('\n'))
                        {
                            var lineRun = new XElement(template.RunPrototype);
                            var textNode = lineRun.Element(W.t);
                            if (textNode != null)
                            {
                                textNode.Value = line;
                            }
                            else
                            {
                                lineRun.Add(new XElement(W.t, line));
                            }
                            if (!firstLine)
                            {
                                lineRun.AddFirst(new XElement(W.br));
                            }
                            p.Add(lineRun);
                            firstLine = false;
                        }
                        return p;
                    }
                    else if (run != null)
                    {
                        var list = new List<XElement>();
                        foreach (var line in newValue.Split('\n'))
                        {
                            list.Add(new XElement(W.r,
                                run.Elements().Where(e => e.Name != W.t),
                                (list.Count > 0) ? new XElement(W.br) : null,
                                new XElement(W.t, line)));
                        }
                        return list;
                    }
                    else
                    {
                        return CreateContextErrorMessage(element, "Content: Unable to find paragraph or run context", templateError);
                    }
                }
                if (element.Name == PA.Image)
                {
                    var xPath = (string?)element.Attribute(PA.Select);
                    if (xPath == null)
                    {
                        return CreateContextErrorMessage(element, "Image: Select attribute is required", templateError);
                    }

                    var optionalString = (string?)element.Attribute(PA.Optional);
                    // Default is true (optional by default), unless explicitly set to false
                    var optional = optionalString == null || !bool.TryParse(optionalString, out var optionalValue) || optionalValue;
                    var alignString = (string?)element.Attribute(PA.Align);
                    var widthAttr = (string?)element.Attribute(PA.Width);
                    var heightAttr = (string?)element.Attribute(PA.Height);
                    var maxWidthAttr = (string?)element.Attribute(PA.MaxWidth);
                    var maxHeightAttr = (string?)element.Attribute(PA.MaxHeight);

                    // EvaluateXPathToString now collects errors instead of throwing
                    string base64Content = EvaluateXPathToString(data, xPath, optional, templateError, evaluationContext);

                    if (string.IsNullOrEmpty(base64Content))
                    {
                        return null;
                    }

                    byte[] imageBytes;
                    try
                    {
                        imageBytes = Convert.FromBase64String(base64Content);
                    }
                    catch (FormatException e)
                    {
                        return CreateContextErrorMessage(element, "Image: " + e.Message, templateError);
                    }

                    if (!TryGetJustification(alignString, out var justification, out var justificationError))
                    {
                        return CreateContextErrorMessage(element, justificationError, templateError);
                    }

                    if (owningPart == null)
                    {
                        throw new OpenXmlPowerToolsException("Image: owning part is not available.");
                    }

                    if (!TryCalculateImageDimensions(imageBytes, widthAttr, heightAttr, maxWidthAttr, maxHeightAttr, out var widthEmu, out var heightEmu, out var sizeError))
                    {
                        return CreateContextErrorMessage(element, sizeError, templateError);
                    }

                    var imagePart = AddImagePart(owningPart);
                    using (var stream = new MemoryStream(imageBytes))
                    {
                        imagePart.FeedData(stream);
                    }

                    var relationshipId = owningPart.GetIdOfPart(imagePart);
                    var docPrId = GetNextDocPrId(owningPart);
                    var imageElement = CreateImageElement(relationshipId, docPrId, widthEmu, heightEmu, justification);
                    return imageElement;
                }
                if (element.Name == PA.Repeat)
                {
                    var selector = (string?)element.Attribute(PA.Select);
                    if (selector == null)
                    {
                        return CreateContextErrorMessage(element, "Repeat: Select attribute is required", templateError);
                    }

                    var optionalString = (string?)element.Attribute(PA.Optional);
                    // Default is true (optional by default), unless explicitly set to false
                    var optional = optionalString == null || !bool.TryParse(optionalString, out var optionalValue) || optionalValue;

                    IEnumerable<XElement> repeatingData;
                    try
                    {
                        repeatingData = EvaluateXPathElements(data, selector, evaluationContext).ToList();
                    }
                    catch (XPathException e)
                    {
                        return CreateContextErrorMessage(element, "XPathException: " + e.Message, templateError);
                    }
                    if (!repeatingData.Any())
                    {
                        if (optional)
                        {
                            return null;
                        }
                        return CreateContextErrorMessage(element, "Repeat: Select returned no data", templateError);
                    }
                    var repeatChildren = element.Elements().ToList();
                    var newContent = repeatingData.Select(d =>
                        {
                            var content = repeatChildren
                                .Select(e => ContentReplacementTransform(e, d, templateError, owningPart, evaluationContext))
                                .ToList();
                            return content;
                        })
                        .ToList();
                    return newContent;
                }
                if (element.Name == PA.Table)
                {
                    var selectAttr = (string?)element.Attribute(PA.Select);
                    if (selectAttr == null)
                    {
                        return CreateContextErrorMessage(element, "Table: Select attribute is required", templateError);
                    }

                    IEnumerable<XElement> tableData;
                    try
                    {
                        tableData = EvaluateXPathElements(data, selectAttr, evaluationContext).ToList();
                    }
                    catch (XPathException e)
                    {
                        return CreateContextErrorMessage(element, "XPathException: " + e.Message, templateError);
                    }
                    if (!tableData.Any())
                    {
                        return CreateContextErrorMessage(element, "Table Select returned no data", templateError);
                    }

                    var table = element.Element(W.tbl);
                    if (table == null)
                    {
                        return CreateContextErrorMessage(element, "Table: Unable to find table element", templateError);
                    }

                    var protoRow = table.Elements(W.tr).Skip(1).FirstOrDefault();
                    var footerRowsBeforeTransform = table
                        .Elements(W.tr)
                        .Skip(2)
                        .ToList();
                    var footerRows = footerRowsBeforeTransform
                        .Select(x => ContentReplacementTransform(x, data, templateError, owningPart, evaluationContext))
                        .ToList();
                    if (protoRow == null)
                    {
                        return CreateContextErrorMessage(element, string.Format("Table does not contain a prototype row"), templateError);
                    }

                    protoRow.Descendants(W.bookmarkStart).Remove();
                    protoRow.Descendants(W.bookmarkEnd).Remove();
                    var tablePrefixNodes = table.Elements().Where(e => e.Name != W.tr).ToList();
                    var cellTemplates = protoRow.Elements(W.tc)
                        .Select(tc => CreateTableCellTemplate(tc, templateError))
                        .ToList();
                    var newTable = new XElement(W.tbl,
                        tablePrefixNodes,
                        table.Elements(W.tr).FirstOrDefault(),
                        tableData.Select(d =>
                            new XElement(W.tr,
                                protoRow.Elements().Where(r => r.Name != W.tc),
                                cellTemplates.Select(ct => BuildTableCell(ct, d, templateError, evaluationContext)))),
                        footerRows
                    );
                    return newTable;
                }
                if (element.Name == PA.Conditional)
                {
                    var xPath = (string?)element.Attribute(PA.Select);
                    if (xPath == null)
                    {
                        return CreateContextErrorMessage(element, "Conditional: Select attribute is required", templateError);
                    }

                    var match = (string?)element.Attribute(PA.Match);
                    var notMatch = (string?)element.Attribute(PA.NotMatch);

                    if (match == null && notMatch == null)
                    {
                        return CreateContextErrorMessage(element, "Conditional: Must specify either Match or NotMatch", templateError);
                    }

                    if (match != null && notMatch != null)
                    {
                        return CreateContextErrorMessage(element, "Conditional: Cannot specify both Match and NotMatch", templateError);
                    }

                    // EvaluateXPathToString now collects errors instead of throwing
                    string? testValue = EvaluateXPathToString(data, xPath, false, templateError, evaluationContext);

                    var conditionIsTrue = (match != null && testValue == match) || (notMatch != null && testValue != notMatch);

                    // Find Else element if present
                    var elseElement = element.Elements(PA.Else).FirstOrDefault();

                    if (conditionIsTrue)
                    {
                        // Process all child elements except Else
                        var content = element.Elements().Where(e => e.Name != PA.Else).Select(e => ContentReplacementTransform(e, data, templateError, owningPart, evaluationContext));
                        return content;
                    }
                    else
                    {
                        // Condition is false
                        if (elseElement != null)
                        {
                            // Process content inside Else
                            var elseContent = elseElement.Elements().Select(e => ContentReplacementTransform(e, data, templateError, owningPart, evaluationContext));
                            return elseContent;
                        }
                        // No Else, return null
                        return null;
                    }
                }
                if (element.Name == PA.Signature)
                {
                    return BuildSignaturePlaceholder(element, templateError);
                }
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => ContentReplacementTransform(n, data, templateError, owningPart, evaluationContext)));
            }
            return node;
        }

        private static object CreateContextErrorMessage(XElement element, string errorMessage, TemplateError templateError)
        {
            var para = element.Descendants(W.p).FirstOrDefault();
            var errorRun = CreateRunErrorMessage(errorMessage, templateError);
            if (para != null)
            {
                return new XElement(W.p, errorRun);
            }
            else
            {
                return errorRun;
            }
        }

        private static XElement CreateRunErrorMessage(string errorMessage, TemplateError templateError)
        {
            templateError.HasError = true;
            var errorRun = new XElement(W.r,
                new XElement(W.rPr,
                    new XElement(W.color, new XAttribute(W.val, "FF0000")),
                    new XElement(W.highlight, new XAttribute(W.val, "yellow"))),
                    new XElement(W.t, errorMessage));
            return errorRun;
        }

        private static XElement CreateParaErrorMessage(string errorMessage, TemplateError templateError)
        {
            templateError.HasError = true;
            var errorPara = new XElement(W.p,
                new XElement(W.r,
                    new XElement(W.rPr,
                        new XElement(W.color, new XAttribute(W.val, "FF0000")),
                        new XElement(W.highlight, new XAttribute(W.val, "yellow"))),
                        new XElement(W.t, errorMessage)));
            return errorPara;
        }

        private static readonly ConcurrentDictionary<string, XPathExpression> s_XPathExpressionCache = new();
        private static readonly ConditionalWeakTable<XElement, ParagraphRunTemplate> s_ParagraphTemplateCache = new();

        private readonly struct EvaluationCacheKey : IEquatable<EvaluationCacheKey>
        {
            public XElement Data { get; }
            public string XPath { get; }
            public bool Optional { get; }

            public EvaluationCacheKey(XElement data, string xPath, bool optional)
            {
                Data = data;
                XPath = xPath;
                Optional = optional;
            }

            public bool Equals(EvaluationCacheKey other) =>
                ReferenceEquals(Data, other.Data) &&
                StringComparer.Ordinal.Equals(XPath, other.XPath) &&
                Optional == other.Optional;

            public override bool Equals(object? obj) => obj is EvaluationCacheKey other && Equals(other);

            public override int GetHashCode() =>
                HashCode.Combine(RuntimeHelpers.GetHashCode(Data), StringComparer.Ordinal.GetHashCode(XPath), Optional);
        }

        private sealed class XPathEvaluationContext
        {
            private readonly Dictionary<EvaluationCacheKey, string> _cache = new();
            private readonly Dictionary<(XElement Data, string XPath), XElement[]> _elementCache = new();

            public bool TryGet(EvaluationCacheKey key, out string value) => _cache.TryGetValue(key, out value);

            public void Store(EvaluationCacheKey key, string value) => _cache[key] = value;

            public bool TryGetElements(XElement data, string xpath, out XElement[] elements) =>
                _elementCache.TryGetValue((data, xpath), out elements);

            public void StoreElements(XElement data, string xpath, XElement[] elements) =>
                _elementCache[(data, xpath)] = elements;
        }

        private sealed class ParagraphRunTemplate
        {
            public XElement? ParagraphProperties { get; }
            public XElement RunPrototype { get; }

            public ParagraphRunTemplate(XElement? paragraphProperties, XElement runPrototype)
            {
                ParagraphProperties = paragraphProperties;
                RunPrototype = runPrototype;
            }
        }

        private static ParagraphRunTemplate GetParagraphTemplate(XElement paragraph, XElement? run)
        {
            return s_ParagraphTemplateCache.GetValue(paragraph, key =>
            {
                var paragraphProps = key.Element(W.pPr);
                var runPrototype = new XElement(W.r);
                var runProps = run?.Element(W.rPr);
                if (runProps != null)
                {
                    runPrototype.Add(new XElement(runProps));
                }
                else
                {
                    runPrototype.Add(new XElement(W.rPr));
                }
                runPrototype.Add(new XElement(W.t));
                return new ParagraphRunTemplate(paragraphProps, runPrototype);
            });
        }

        private static string EvaluateXPathToString(XElement element, string xPath, bool optional, TemplateError templateError, XPathEvaluationContext evaluationContext)
        {
            var cacheKey = new EvaluationCacheKey(element, xPath, optional);
            if (evaluationContext.TryGet(cacheKey, out var cachedValue))
            {
                return cachedValue;
            }

            if (string.IsNullOrWhiteSpace(xPath))
            {
                evaluationContext.Store(cacheKey, string.Empty);
                return string.Empty;
            }

            object xPathSelectResult;
            try
            {
                var navigator = element.CreateNavigator();
                var baseExpression = s_XPathExpressionCache.GetOrAdd(xPath, key => XPathExpression.Compile(key));
                var expression = baseExpression.Clone();
                expression.SetContext(navigator);
                xPathSelectResult = navigator.Evaluate(expression);
            }
            catch (XPathException e)
            {
                // Collect XPath syntax errors instead of throwing
                var errorMsg = "XPathException: " + e.Message;
                templateError.AddError(errorMsg);
                var invalidResult = "[ERROR: Invalid XPath]";
                evaluationContext.Store(cacheKey, invalidResult);
                return invalidResult;
            }

            string result;
            if ((xPathSelectResult is IEnumerable) && !(xPathSelectResult is string))
            {
                var selectedData = ((IEnumerable)xPathSelectResult).Cast<object>().ToList();
                if (!selectedData.Any())
                {
                    if (optional)
                    {
                        result = string.Empty;
                    }
                    else
                    {
                        // Collect missing field error instead of throwing
                        templateError.AddMissingField(xPath);
                        result = "[ERROR: Missing field]";
                    }
                    evaluationContext.Store(cacheKey, result);
                    return result;
                }
                if (selectedData.Count > 1)
                {
                    // Collect multiple results error instead of throwing
                    var errorMsg = string.Format("XPath expression ({0}) returned more than one node", xPath);
                    templateError.AddError(errorMsg);
                    result = "[ERROR: Multiple results]";
                    evaluationContext.Store(cacheKey, result);
                    return result;
                }

                var selectedDatum = selectedData.First();
                if (selectedDatum is XElement element1)
                {
                    result = element1.Value;
                    evaluationContext.Store(cacheKey, result);
                    return result;
                }

                if (selectedDatum is XAttribute attribute)
                {
                    result = attribute.Value;
                    evaluationContext.Store(cacheKey, result);
                    return result;
                }

                if (selectedDatum is XPathNavigator navigator)
                {
                    result = navigator.Value;
                    evaluationContext.Store(cacheKey, result);
                    return result;
                }
            }

            result = xPathSelectResult.ToString() ?? string.Empty;
            evaluationContext.Store(cacheKey, result);
            return result;
        }

        private static XElement[] EvaluateXPathElements(XElement element, string xPath, XPathEvaluationContext evaluationContext)
        {
            if (evaluationContext.TryGetElements(element, xPath, out var cached))
            {
                return cached;
            }

            var navigator = element.CreateNavigator();
            var baseExpression = s_XPathExpressionCache.GetOrAdd(xPath, key => XPathExpression.Compile(key));
            var expression = baseExpression.Clone();
            expression.SetContext(navigator);
            var iterator = navigator.Select(expression);
            var buffer = new List<XElement>();
            while (iterator.MoveNext())
            {
                var current = iterator.Current;
                if (current?.UnderlyingObject is XElement xElement)
                {
                    buffer.Add(xElement);
                }
            }

            var result = buffer.ToArray();
            evaluationContext.StoreElements(element, xPath, result);
            return result;
        }

        private sealed record TableCellTemplate(
            XElement[] NonParagraphNodes,
            XElement? ParagraphProperties,
            XElement? RunProperties,
            string XPath,
            bool HasParagraph);

        private static TableCellTemplate CreateTableCellTemplate(XElement tc, TemplateError templateError)
        {
            var paragraph = tc.Elements(W.p).FirstOrDefault();
            var nonParagraphNodes = tc.Elements().Where(z => z.Name != W.p).Select(node => new XElement(node)).ToArray();
            if (paragraph == null)
            {
                return new TableCellTemplate(nonParagraphNodes, null, null, string.Empty, false);
            }
            var paragraphProperties = paragraph.Element(W.pPr);
            var cellRun = paragraph.Elements(W.r).FirstOrDefault();
            var runProperties = cellRun?.Element(W.rPr);
            return new TableCellTemplate(nonParagraphNodes, paragraphProperties, runProperties, paragraph.Value, true);
        }

        private static XElement BuildTableCell(TableCellTemplate template, XElement data, TemplateError templateError, XPathEvaluationContext evaluationContext)
        {
            if (!template.HasParagraph)
            {
                return new XElement(W.tc,
                    template.NonParagraphNodes,
                    new XElement(W.p,
                        CreateRunErrorMessage("Table cell does not contain a paragraph", templateError)));
            }

            var newValue = EvaluateXPathToString(data, template.XPath, false, templateError, evaluationContext);
            var paragraphProps = template.ParagraphProperties != null ? new XElement(template.ParagraphProperties) : null;
            var runProps = template.RunProperties != null ? new XElement(template.RunProperties) : new XElement(W.rPr);
            var run = new XElement(W.r,
                runProps,
                new XElement(W.t, newValue));
            var paragraphElement = new XElement(W.p,
                paragraphProps,
                run);
            return new XElement(W.tc,
                template.NonParagraphNodes.Select(node => new XElement(node)),
                paragraphElement);
        }

        // Nested classes
        private static class PA
        {
            public static readonly XName Content = "Content";
            public static readonly XName Image = "Image";
            public static readonly XName Table = "Table";
            public static readonly XName Repeat = "Repeat";
            public static readonly XName EndRepeat = "EndRepeat";
            public static readonly XName Conditional = "Conditional";
            public static readonly XName Else = "Else";
            public static readonly XName EndConditional = "EndConditional";
            public static readonly XName Signature = "Signature";
            public static readonly XName Select = "Select";
            public static readonly XName Optional = "Optional";
            public static readonly XName Align = "Align";
            public static readonly XName Width = "Width";
            public static readonly XName Height = "Height";
            public static readonly XName MaxWidth = "MaxWidth";
            public static readonly XName MaxHeight = "MaxHeight";
            public static readonly XName Match = "Match";
            public static readonly XName NotMatch = "NotMatch";
            public static readonly XName Depth = "Depth";
            public static readonly XName Id = "Id";
            public static readonly XName Label = "Label";
            public static readonly XName PageHint = "PageHint";
        }

        private class TemplateError
        {
            public bool HasError;
            public List<string> MissingFields { get; } = new List<string>();
            public List<string> AllErrors { get; } = new List<string>();

            public void AddMissingField(string xpath)
            {
                HasError = true;
                if (!MissingFields.Contains(xpath))
                {
                    MissingFields.Add(xpath);
                    AllErrors.Add($"Missing field: {xpath}");
                }
            }

            public void AddError(string errorMessage)
            {
                HasError = true;
                if (!AllErrors.Contains(errorMessage))
                {
                    AllErrors.Add(errorMessage);
                }
            }

            public string GetErrorSummary()
            {
                if (!HasError || AllErrors.Count == 0)
                {
                    return string.Empty;
                }

                if (MissingFields.Count > 0)
                {
                    return $"Template errors found:\n" +
                           $"- Missing fields ({MissingFields.Count}): {string.Join(", ", MissingFields)}\n" +
                           $"- Total errors: {AllErrors.Count}";
                }

                return $"Template errors found ({AllErrors.Count}):\n- " + string.Join("\n- ", AllErrors);
            }
        }
    }
}
