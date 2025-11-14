using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace DocumentAssembler.Core
{
    /// <summary>
    /// High-performance XML schema extractor for DocumentAssembler templates.
    /// Analyzes DOCX templates and generates the required XML data structure.
    /// </summary>
    public class TemplateSchemaExtractor
    {
        /// <summary>
        /// Result of schema extraction containing XML template and metadata
        /// </summary>
        public class SchemaExtractionResult
        {
            /// <summary>
            /// Generated XML template with placeholders
            /// </summary>
            public string XmlTemplate { get; set; } = string.Empty;

            /// <summary>
            /// List of all fields discovered in the template
            /// </summary>
            public List<FieldInfo> Fields { get; set; } = new List<FieldInfo>();

            /// <summary>
            /// Root element name (extracted from the most common parent path)
            /// </summary>
            public string RootElementName { get; set; } = "Data";

            /// <summary>
            /// Generated XSD markup (optional-aware)
            /// </summary>
            public string XsdMarkup { get; set; } = string.Empty;

            /// <summary>
            /// Generates a formatted, indented XML string
            /// </summary>
            public string ToFormattedXml()
            {
                try
                {
                    var doc = XDocument.Parse(XmlTemplate);
                    return doc.ToString();
                }
                catch
                {
                    return XmlTemplate;
                }
            }

            /// <summary>
            /// Generates a formatted, indented XSD string
            /// </summary>
            public string ToFormattedXsd()
            {
                try
                {
                    var doc = XDocument.Parse(XsdMarkup);
                    return doc.ToString();
                }
                catch
                {
                    return XsdMarkup;
                }
            }
        }

        /// <summary>
        /// Information about a discovered field in the template
        /// </summary>
        public class FieldInfo
        {
            /// <summary>
            /// Full XPath expression (e.g., "Customer/Name")
            /// </summary>
            public string XPath { get; set; } = string.Empty;

            /// <summary>
            /// Tag type (Content, Image, Repeat, Table, Conditional)
            /// </summary>
            public string TagType { get; set; } = string.Empty;

            /// <summary>
            /// Whether the field is optional (defaults to true)
            /// </summary>
            public bool IsOptional { get; set; } = true;

            /// <summary>
            /// Whether this field is part of a repeating collection
            /// </summary>
            public bool IsRepeating { get; set; } = false;

            /// <summary>
            /// Parent XPath for nested structures
            /// </summary>
            public string? ParentXPath { get; set; }

            /// <summary>
            /// Element name (last part of XPath)
            /// </summary>
            public string ElementName
            {
                get
                {
                    var parts = XPath.Split('/');
                    return parts.Length > 0 ? parts[^1] : XPath;
                }
            }

            /// <summary>
            /// Additional attributes (Match, NotMatch, Align, etc.)
            /// </summary>
            public Dictionary<string, string> Attributes { get; set; } = new Dictionary<string, string>();

            /// <summary>
            /// Whether the field targets an attribute (e.g. Order/@Number)
            /// </summary>
            public bool IsAttribute { get; set; }
        }

        /// <summary>
        /// Metadata describing a MailMerge field discovered in a template.
        /// </summary>
        public sealed record MailMergeField(string FieldName, string XPath);

        /// <summary>
        /// Extracts XML schema from a DOCX template document.
        /// Ultra-fast single-pass algorithm with caching.
        /// </summary>
        /// <param name="templateDoc">The template document to analyze</param>
        /// <returns>Schema extraction result with XML template and metadata</returns>
        public static SchemaExtractionResult ExtractXmlSchema(WmlDocument templateDoc)
        {
            var fields = new Dictionary<string, FieldInfo>(StringComparer.OrdinalIgnoreCase);
            var repeatingPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            var byteArray = templateDoc.DocumentByteArray;
            using var mem = new System.IO.MemoryStream(byteArray, false);
            using (var wordDoc = WordprocessingDocument.Open(mem, false))
            {
                // Process all content parts in parallel for maximum performance
                var allFields = new List<FieldInfo>();

                foreach (var part in wordDoc.ContentParts())
                {
                    if (part?.RootElement != null)
                    {
                        ExtractFieldsFromPart(part, allFields, repeatingPaths);
                    }
                }

                // Merge fields efficiently (last occurrence wins for metadata)
                foreach (var field in allFields)
                {
                    var key = field.XPath;
                    if (!fields.ContainsKey(key) || !field.IsOptional)
                    {
                        fields[key] = field;
                    }
                }
            }

            // Build hierarchical XML structure efficiently
            var sortedFields = fields.Values.OrderBy(f => f.XPath).ToList();
            var result = new SchemaExtractionResult
            {
                Fields = sortedFields
            };

            result.RootElementName = DetermineRootElementName(sortedFields);
            var fieldTree = BuildFieldTree(sortedFields, repeatingPaths);
            result.XmlTemplate = sortedFields.Count == 0 ? "<Data />" : BuildXmlFromTree(fieldTree);
            result.XsdMarkup = BuildXsdTemplate(fieldTree, result.RootElementName);

            return result;
        }

        /// <summary>
        /// Extracts fields from a single document part with high performance
        /// </summary>
        private static void ExtractFieldsFromPart(OpenXmlPart part, List<FieldInfo> fields, HashSet<string> repeatingPaths)
        {
            var xDoc = part.GetXDocument();
            if (xDoc.Root == null)
            {
                return;
            }

            var contextStack = new Stack<string>();
            contextStack.Push(string.Empty);

            foreach (var tag in EnumerateMetadataTags(xDoc.Root))
            {
                if (tag == null)
                {
                    continue;
                }

                switch (tag.Name)
                {
                    case "EndRepeat":
                        if (contextStack.Count > 1)
                        {
                            contextStack.Pop();
                        }
                        continue;
                    case "EndConditional":
                    case "Else":
                        continue;
                }

                string? select = null;
                if (tag.RequiresSelect && !tag.Attributes.TryGetValue("Select", out select))
                {
                    continue;
                }

                var currentContext = contextStack.Peek();
                var resolvedPath = tag.RequiresSelect ? ResolveXPath(select!, currentContext) : string.Empty;

                switch (tag.Name)
                {
                    case "Repeat":
                        if (string.IsNullOrEmpty(resolvedPath))
                        {
                            continue;
                        }
                        repeatingPaths.Add(resolvedPath);
                        fields.Add(CreateFieldInfo(resolvedPath, tag, isRepeating: true));
                        contextStack.Push(resolvedPath);
                        break;

                    case "Table":
                        if (string.IsNullOrEmpty(resolvedPath))
                        {
                            continue;
                        }
                        repeatingPaths.Add(resolvedPath);
                        fields.Add(CreateFieldInfo(resolvedPath, tag, isRepeating: true));
                        break;

                    case "Content":
                    case "Image":
                    case "Conditional":
                    case "Signature":
                        if (string.IsNullOrEmpty(resolvedPath))
                        {
                            continue;
                        }
                        fields.Add(CreateFieldInfo(resolvedPath, tag));
                        break;
                }
            }

            if (xDoc.Root != null)
            {
                foreach (var mailMergeField in EnumerateMailMergeFields(xDoc.Root))
                {
                    var info = new FieldInfo
                    {
                        XPath = mailMergeField.XPath,
                        TagType = "MailMerge",
                        IsOptional = true,
                        ParentXPath = GetParentPath(mailMergeField.XPath),
                        Attributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                        {
                            ["FieldName"] = mailMergeField.FieldName
                        },
                        IsRepeating = false,
                        IsAttribute = mailMergeField.XPath.Contains("/@") ||
                                      mailMergeField.XPath.StartsWith("@", StringComparison.Ordinal)
                    };
                    fields.Add(info);
                }
            }
        }

        private static FieldInfo CreateFieldInfo(string xpath, ParsedTag tag, bool isRepeating = false)
        {
            var attributes = new Dictionary<string, string>(tag.Attributes, StringComparer.OrdinalIgnoreCase);
            var isOptional = true;
            if (attributes.TryGetValue("Optional", out var optionalText) &&
                bool.TryParse(optionalText, out var optionalValue))
            {
                isOptional = optionalValue;
            }

            var normalizedPath = SanitizeXPath(xpath);

            return new FieldInfo
            {
                XPath = normalizedPath,
                TagType = tag.Name,
                IsOptional = isOptional,
                ParentXPath = GetParentPath(normalizedPath),
                Attributes = attributes,
                IsRepeating = isRepeating,
                IsAttribute = normalizedPath.Contains("/@") || normalizedPath.StartsWith("@", StringComparison.Ordinal)
            };
        }

        private sealed class ParsedTag
        {
            public string Name { get; init; } = string.Empty;
            public Dictionary<string, string> Attributes { get; init; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            public string RawText { get; init; } = string.Empty;
            public bool RequiresSelect =>
                Name is "Content" or "Image" or "Table" or "Repeat" or "Conditional" or "Signature";
        }

        private static readonly HashSet<string> s_SupportedTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Content",
            "Image",
            "Table",
            "Repeat",
            "EndRepeat",
            "Conditional",
            "Else",
            "EndConditional",
            "Signature",
        };

        private static IEnumerable<ParsedTag?> EnumerateMetadataTags(XElement root)
        {
            foreach (var element in root.Descendants())
            {
                if (element.Name == W.sdt)
                {
                    foreach (var tag in ExtractTagsFromElement(element))
                    {
                        yield return tag;
                    }
                }
                else if (element.Name == W.p &&
                         !element.Ancestors(W.sdt).Any() &&
                         !element.Descendants(W.sdt).Any())
                {
                    foreach (var tag in ExtractTagsFromElement(element))
                    {
                        yield return tag;
                    }
                }
            }
        }

        private static IEnumerable<ParsedTag?> ExtractTagsFromElement(XElement element)
        {
            var text = string.Concat(element.Descendants(W.t).Select(t => t.Value));
            var normalizedText = NormalizeMetadataText(text);
            if (string.IsNullOrWhiteSpace(normalizedText))
            {
                yield break;
            }

            foreach (var token in SplitIntoTagStrings(normalizedText))
            {
                if (TryParseTag(token, out var parsed))
                {
                    yield return parsed;
                }
            }
        }

        private static string NormalizeMetadataText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.Empty;
            }

            var normalized = text.Trim()
                .Replace('\u201C', '"')
                .Replace('\u201D', '"')
                .Replace('\u2018', '\'')
                .Replace('\u2019', '\'')
                .Replace("\r", " ")
                .Replace("\n", " ");

            return System.Net.WebUtility.HtmlDecode(normalized);
        }

        private static IEnumerable<string> SplitIntoTagStrings(string text)
        {
            var tags = new List<string>();
            var index = 0;
            while (index < text.Length)
            {
                var start = text.IndexOf('<', index);
                if (start == -1)
                {
                    break;
                }

                if (start + 1 < text.Length && text[start + 1] == '#')
                {
                    var end = text.IndexOf("#>", start + 2, StringComparison.Ordinal);
                    if (end == -1)
                    {
                        break;
                    }
                    tags.Add(text.Substring(start, end + 2 - start));
                    index = end + 2;
                }
                else
                {
                    var end = text.IndexOf('>', start + 1);
                    if (end == -1)
                    {
                        break;
                    }
                    tags.Add(text.Substring(start, end + 1 - start));
                    index = end + 1;
                }
            }

            return tags;
        }

        private static bool TryParseTag(string rawTag, out ParsedTag? parsedTag)
        {
            parsedTag = null;
            if (string.IsNullOrWhiteSpace(rawTag))
            {
                return false;
            }

            var normalized = NormalizeTagMarkup(rawTag);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return false;
            }

            XElement element;
            try
            {
                element = XElement.Parse(normalized);
            }
            catch
            {
                return false;
            }

            var name = element.Name.LocalName;
            if (!s_SupportedTags.Contains(name))
            {
                return false;
            }

            parsedTag = new ParsedTag
            {
                Name = name,
                Attributes = element.Attributes()
                    .ToDictionary(a => a.Name.LocalName, a => a.Value, StringComparer.OrdinalIgnoreCase),
                RawText = rawTag
            };

            return true;
        }

        private static string? NormalizeTagMarkup(string rawTag)
        {
            var trimmed = rawTag.Trim();
            if (string.IsNullOrEmpty(trimmed) || !trimmed.StartsWith("<", StringComparison.Ordinal))
            {
                return null;
            }

            if (trimmed.StartsWith("<#", StringComparison.Ordinal))
            {
                if (!trimmed.EndsWith("#>", StringComparison.Ordinal))
                {
                    return null;
                }

                var inner = trimmed.Substring(2, trimmed.Length - 4).Trim();
                if (string.IsNullOrEmpty(inner))
                {
                    return null;
                }

                if (!inner.EndsWith("/>", StringComparison.Ordinal))
                {
                    inner = inner.TrimEnd('/');
                    inner = $"{inner} />";
                }

                return $"<{inner.TrimStart('<')}";
            }

            return trimmed;
        }

        private static string ResolveXPath(string select, string currentContext)
        {
            if (string.IsNullOrWhiteSpace(select))
            {
                return string.Empty;
            }

            var trimmed = select.Trim();
            var isAbsolute = trimmed.StartsWith("/", StringComparison.Ordinal);
            var isRelative = trimmed.StartsWith(".", StringComparison.Ordinal) ||
                             trimmed.StartsWith("@", StringComparison.Ordinal) ||
                             trimmed.StartsWith("..", StringComparison.Ordinal);

            List<string> baseSegments;
            if (isAbsolute)
            {
                baseSegments = new List<string>();
                trimmed = trimmed.TrimStart('/');
            }
            else if (isRelative)
            {
                baseSegments = SplitPathSegments(currentContext);
            }
            else
            {
                baseSegments = new List<string>();
            }

            var relativeSegments = SplitPathSegments(trimmed);
            var combined = CombineSegments(baseSegments, relativeSegments);
            return string.Join("/", combined).Trim('/');
        }

        private static List<string> CombineSegments(List<string> baseSegments, List<string> relativeSegments)
        {
            var combined = new List<string>(baseSegments);
            foreach (var segment in relativeSegments)
            {
                if (segment == ".")
                {
                    continue;
                }
                if (segment == "..")
                {
                    if (combined.Count > 0)
                    {
                        combined.RemoveAt(combined.Count - 1);
                    }
                    continue;
                }
                combined.Add(segment);
            }
            return combined;
        }

        private static List<string> SplitPathSegments(string path)
        {
            var segments = new List<string>();
            if (string.IsNullOrWhiteSpace(path))
            {
                return segments;
            }

            var normalized = path.Replace("\\", "/");
            var tokens = normalized.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var token in tokens)
            {
                var trimmed = TrimPredicates(token.Trim());
                if (string.IsNullOrEmpty(trimmed))
                {
                    continue;
                }
                segments.Add(trimmed);
            }

            return segments;
        }

        private static string TrimPredicates(string segment)
        {
            var bracketIndex = segment.IndexOf('[');
            if (bracketIndex >= 0)
            {
                return segment.Substring(0, bracketIndex);
            }
            return segment;
        }

        private static string SanitizeXPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return string.Empty;
            }

            var segments = SplitPathSegments(path);
            return string.Join("/", segments);
        }

        private static string? GetParentPath(string xpath)
        {
            if (string.IsNullOrWhiteSpace(xpath))
            {
                return null;
            }

            var index = xpath.LastIndexOf('/');
            if (index <= 0)
            {
                return null;
            }

            return xpath.Substring(0, index);
        }

        /// <summary>
        /// Builds a tree representation of the discovered fields
        /// </summary>
        private static XmlNode BuildFieldTree(List<FieldInfo> fields, HashSet<string> repeatingPaths)
        {
            var root = new XmlNode
            {
                Name = "Data",
                Children = new Dictionary<string, XmlNode>(StringComparer.OrdinalIgnoreCase)
            };

            foreach (var field in fields)
            {
                var parts = field.XPath.Split('/')
                    .Select(p => p.Trim())
                    .Where(p => !string.IsNullOrWhiteSpace(p) && p != "." && !p.StartsWith("-") && !char.IsDigit(p[0]))
                    .ToList();

                if (parts.Count == 0)
                {
                    continue;
                }

                string? attributeSegment = null;
                if (parts[^1].StartsWith("@", StringComparison.Ordinal))
                {
                    attributeSegment = parts[^1];
                }

                var traversalCount = attributeSegment != null ? parts.Count - 1 : parts.Count;
                var currentNode = root;
                var pathSegments = new List<string>();

                for (int i = 0; i < traversalCount; i++)
                {
                    var part = parts[i];
                    pathSegments.Add(part);
                    var pathKey = string.Join("/", pathSegments);

                    if (!currentNode.Children.TryGetValue(part, out var childNode))
                    {
                        childNode = new XmlNode
                        {
                            Name = part,
                            Children = new Dictionary<string, XmlNode>(StringComparer.OrdinalIgnoreCase),
                            IsRepeating = repeatingPaths.Contains(pathKey)
                        };
                        currentNode.Children[part] = childNode;
                    }
                    else if (!childNode.IsRepeating && repeatingPaths.Contains(pathKey))
                    {
                        childNode.IsRepeating = true;
                    }

                    currentNode = childNode;

                    var isLeafNode = i == traversalCount - 1 && attributeSegment == null;
                    if (isLeafNode)
                    {
                        currentNode.IsOptional = field.IsOptional;
                        if (field.TagType == "Content" ||
                            field.TagType == "Image" ||
                            field.TagType == "MailMerge")
                        {
                            currentNode.IsLeaf = true;
                            currentNode.FieldType = field.TagType;
                        }
                    }
                }

                if (attributeSegment != null)
                {
                    var attributeName = attributeSegment.Substring(1);
                    if (!string.IsNullOrWhiteSpace(attributeName))
                    {
                        currentNode.AddOrUpdateAttribute(attributeName, field.IsOptional);
                    }
                }
            }

            return root;
        }

        /// <summary>
        /// Builds XML markup from the tree representation
        /// </summary>
        private static string BuildXmlFromTree(XmlNode root)
        {
            var sb = new StringBuilder();
            BuildXmlString(root, sb, 0);
            return sb.ToString();
        }

        /// <summary>
        /// Builds optional-aware XSD markup from the tree representation
        /// </summary>
        private static string BuildXsdTemplate(XmlNode root, string rootElementName)
        {
            var startNode = root;
            if (!string.Equals(rootElementName, root.Name, StringComparison.OrdinalIgnoreCase) &&
                root.Children.TryGetValue(rootElementName, out var candidate))
            {
                startNode = candidate;
            }

            var sb = new StringBuilder();
            sb.AppendLine(@"<?xml version=""1.0"" encoding=""utf-8""?>");
            sb.AppendLine(@"<xs:schema xmlns:xs=""http://www.w3.org/2001/XMLSchema"">");
            BuildXsdElement(startNode, sb, 1, true);
            sb.AppendLine("</xs:schema>");
            return sb.ToString();
        }

        /// <summary>
        /// Recursively builds XSD elements
        /// </summary>
        private static void BuildXsdElement(XmlNode node, StringBuilder sb, int indent, bool isRoot)
        {
            var indentStr = new string(' ', indent * 2);
            var occurrenceAttrs = BuildXsdOccurrenceAttributes(node, isRoot);
            var hasChildren = node.Children.Count > 0;
            var hasAttributes = node.Attributes.Count > 0;
            var isLeaf = node.IsLeaf && !hasChildren;

            if (isLeaf && !hasAttributes)
            {
                var type = node.FieldType == "Image" ? "xs:base64Binary" : "xs:string";
                sb.AppendLine($"{indentStr}<xs:element name=\"{node.Name}\" type=\"{type}\"{occurrenceAttrs} />");
                return;
            }

            if (!hasChildren && !hasAttributes)
            {
                sb.AppendLine($"{indentStr}<xs:element name=\"{node.Name}\" type=\"xs:string\"{occurrenceAttrs} />");
                return;
            }

            if (isLeaf && !hasChildren)
            {
                var baseType = node.FieldType == "Image" ? "xs:base64Binary" : "xs:string";
                sb.AppendLine($"{indentStr}<xs:element name=\"{node.Name}\"{occurrenceAttrs}>");
                sb.AppendLine($"{indentStr}  <xs:complexType>");
                sb.AppendLine($"{indentStr}    <xs:simpleContent>");
                sb.AppendLine($"{indentStr}      <xs:extension base=\"{baseType}\">");
                AppendXsdAttributes(node, sb, indent + 4);
                sb.AppendLine($"{indentStr}      </xs:extension>");
                sb.AppendLine($"{indentStr}    </xs:simpleContent>");
                sb.AppendLine($"{indentStr}  </xs:complexType>");
                sb.AppendLine($"{indentStr}</xs:element>");
                return;
            }

            sb.AppendLine($"{indentStr}<xs:element name=\"{node.Name}\"{occurrenceAttrs}>");
            sb.AppendLine($"{indentStr}  <xs:complexType>");
            if (hasChildren)
            {
                sb.AppendLine($"{indentStr}    <xs:sequence>");
                foreach (var child in node.Children.Values.OrderBy(n => n.Name))
                {
                    BuildXsdElement(child, sb, indent + 3, false);
                }
                sb.AppendLine($"{indentStr}    </xs:sequence>");
            }
            AppendXsdAttributes(node, sb, indent + 2);
            sb.AppendLine($"{indentStr}  </xs:complexType>");
            sb.AppendLine($"{indentStr}</xs:element>");
        }

        private static void AppendXsdAttributes(XmlNode node, StringBuilder sb, int indent)
        {
            if (node.Attributes.Count == 0)
            {
                return;
            }

            var indentStr = new string(' ', indent * 2);
            foreach (var attribute in node.Attributes.OrderBy(a => a.Name))
            {
                var use = attribute.IsOptional ? "optional" : "required";
                sb.AppendLine($"{indentStr}<xs:attribute name=\"{attribute.Name}\" type=\"xs:string\" use=\"{use}\" />");
            }
        }

        /// <summary>
        /// Returns occurrence attributes for optional/repeating nodes
        /// </summary>
        private static string BuildXsdOccurrenceAttributes(XmlNode node, bool isRoot)
        {
            var attrs = new StringBuilder();
            if (!isRoot && node.IsOptional)
            {
                attrs.Append(" minOccurs=\"0\"");
            }

            if (node.IsRepeating)
            {
                attrs.Append(" maxOccurs=\"unbounded\"");
            }

            return attrs.ToString();
        }

        /// <summary>
        /// Recursively builds XML string with proper indentation
        /// </summary>
        private static void BuildXmlString(XmlNode node, StringBuilder sb, int indent)
        {
            if (node.Name == "Data")
            {
                sb.AppendLine("<Data>");
                foreach (var child in node.Children.Values.OrderBy(n => n.Name))
                {
                    BuildXmlString(child, sb, indent + 1);
                }
                sb.Append("</Data>");
                return;
            }

            var indentStr = new string(' ', indent * 2);
            var attributeMarkup = node.Attributes.Count == 0
                ? string.Empty
                : " " + string.Join(" ", node.Attributes.OrderBy(a => a.Name).Select(a => $"{a.Name}=\"[value]\""));
            var optionalComment = node.IsOptional ? " <!-- Optional -->" : string.Empty;
            var optionalAttributeComment = BuildOptionalAttributeComment(node);

            if (node.IsLeaf && node.Children.Count == 0)
            {
                var placeholder = node.FieldType == "Image" ? "[Base64 image]" : "[value]";
                var comment = node.FieldType == "Image"
                    ? (node.IsOptional ? " <!-- Optional, Base64 encoded image -->" : " <!-- Base64 encoded image -->")
                    : optionalComment;
                sb.AppendLine($"{indentStr}<{node.Name}{attributeMarkup}>{placeholder}</{node.Name}>{comment}");
                return;
            }

            if (node.IsRepeating)
            {
                for (var sample = 0; sample < 2; sample++)
                {
                    sb.AppendLine($"{indentStr}<{node.Name}{attributeMarkup}> <!-- Repeating -->");
                    foreach (var child in node.Children.Values.OrderBy(n => n.Name))
                    {
                        BuildXmlString(child, sb, indent + 1);
                    }
                    sb.AppendLine($"{indentStr}</{node.Name}>");
                }
                return;
            }

            if (node.Children.Count > 0)
            {
                sb.AppendLine($"{indentStr}<{node.Name}{attributeMarkup}>");
                foreach (var child in node.Children.Values.OrderBy(n => n.Name))
                {
                    BuildXmlString(child, sb, indent + 1);
                }
                sb.AppendLine($"{indentStr}</{node.Name}>{optionalComment}{optionalAttributeComment}");
                return;
            }

            sb.AppendLine($"{indentStr}<{node.Name}{attributeMarkup} />{optionalComment}{optionalAttributeComment}");
        }

        private static string BuildOptionalAttributeComment(XmlNode node)
        {
            var optionalAttributes = node.Attributes
                .Where(a => a.IsOptional)
                .Select(a => a.Name)
                .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (optionalAttributes.Count == 0)
            {
                return string.Empty;
            }

            return $" <!-- Optional attributes: {string.Join(", ", optionalAttributes)} -->";
        }

        /// <summary>
        /// Determines the root element name from field paths
        /// </summary>
        private static string DetermineRootElementName(List<FieldInfo> fields)
        {
            if (fields.Count == 0) return "Data";

            // Find the most common root element
            var rootCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            foreach (var field in fields)
            {
                var parts = field.XPath.Split('/');

                // Find first valid element (skip "." for relative paths like "./Customer")
                string? root = null;
                foreach (var part in parts)
                {
                    var trimmedPart = part.Trim();
                    if (string.IsNullOrEmpty(trimmedPart) || trimmedPart == ".")
                    {
                        continue; // Skip empty parts and current directory marker
                    }

                    // Skip invalid XML element names
                    if (trimmedPart.StartsWith("-") || char.IsDigit(trimmedPart[0]))
                    {
                        continue;
                    }

                    root = trimmedPart;
                    break; // Found first valid element
                }

                if (root != null)
                {
                    rootCounts[root] = rootCounts.GetValueOrDefault(root, 0) + 1;
                }
            }

            if (rootCounts.Count > 0)
            {
                // If there's exactly one unique root, use it
                if (rootCounts.Count == 1)
                {
                    return rootCounts.First().Key;
                }

                // Check if any fields have hierarchical structure (meaningful nesting beyond current dir)
                // e.g., "./Orders/Order" has hierarchy, but "./Name" doesn't
                var hasHierarchy = fields.Any(f =>
                {
                    var parts = f.XPath.Split('/').Where(p => !string.IsNullOrWhiteSpace(p) && p != ".").ToArray();
                    return parts.Length > 1;
                });

                // If there are multiple different roots all with count=1 AND no hierarchical structure
                // then there's no clear winner - use "Data"
                // This handles truly flat templates like: ./Field1, ./Field2, ./Field3
                if (rootCounts.All(kvp => kvp.Value == 1) && !hasHierarchy)
                {
                    return "Data";
                }

                // Otherwise, return the most common root element (or first one if tied)
                return rootCounts.OrderByDescending(kvp => kvp.Value).First().Key;
            }

            return "Data";
        }

        /// <summary>
        /// Internal node structure for XML tree building
        /// </summary>
        private class XmlNode
        {
            public string Name { get; set; } = string.Empty;
            public Dictionary<string, XmlNode> Children { get; set; } = new Dictionary<string, XmlNode>(StringComparer.OrdinalIgnoreCase);
            public bool IsLeaf { get; set; }
            public bool IsRepeating { get; set; }
            public bool IsOptional { get; set; }
            public string FieldType { get; set; } = string.Empty;
            public List<AttributeInfo> Attributes { get; } = new List<AttributeInfo>();

            public void AddOrUpdateAttribute(string name, bool isOptional)
            {
                var existing = Attributes.FirstOrDefault(a => string.Equals(a.Name, name, StringComparison.OrdinalIgnoreCase));
                if (existing == null)
                {
                    Attributes.Add(new AttributeInfo
                    {
                        Name = name,
                        IsOptional = isOptional
                    });
                }
                else if (!isOptional)
                {
                    existing.IsOptional = false;
                }
            }
        }

        private class AttributeInfo
        {
            public string Name { get; set; } = string.Empty;
            public bool IsOptional { get; set; } = true;
        }

        /// <summary>
        /// Extracts the list of MailMerge fields (MERGEFIELD) defined in a template.
        /// </summary>
        public static IReadOnlyList<MailMergeField> ExtractMailMergeFields(WmlDocument templateDoc)
        {
            var results = new List<MailMergeField>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            var byteArray = templateDoc.DocumentByteArray;
            using var mem = new System.IO.MemoryStream(byteArray, false);
            using (var wordDoc = WordprocessingDocument.Open(mem, false))
            {
                foreach (var part in wordDoc.ContentParts())
                {
                    if (part?.RootElement == null)
                    {
                        continue;
                    }

                    var xDoc = part.GetXDocument();
                    if (xDoc.Root == null)
                    {
                        continue;
                    }

                    foreach (var field in EnumerateMailMergeFields(xDoc.Root))
                    {
                        if (seen.Add(field.FieldName))
                        {
                            results.Add(field);
                        }
                    }
                }
            }

            return results;
        }

        private static IEnumerable<MailMergeField> EnumerateMailMergeFields(XElement root)
        {
            foreach (var instruction in EnumerateMailMergeInstructions(root))
            {
                var fieldName = ParseMailMergeFieldName(instruction);
                if (string.IsNullOrWhiteSpace(fieldName))
                {
                    continue;
                }

                var xpath = NormalizeMailMergeXPath(fieldName);
                if (string.IsNullOrEmpty(xpath))
                {
                    continue;
                }

                yield return new MailMergeField(fieldName, xpath);
            }
        }

        private static IEnumerable<string> EnumerateMailMergeInstructions(XElement root)
        {
            foreach (var simpleField in root.Descendants(W.fldSimple))
            {
                var instruction = (string?)simpleField.Attribute(W.instr);
                if (!string.IsNullOrWhiteSpace(instruction))
                {
                    yield return instruction!;
                }
            }

            var capturing = false;
            var buffer = new StringBuilder();

            foreach (var element in root.Descendants())
            {
                if (element.Name == W.fldChar)
                {
                    var type = (string?)element.Attribute(W.fldCharType);
                    if (string.Equals(type, "begin", StringComparison.OrdinalIgnoreCase))
                    {
                        capturing = true;
                        buffer.Clear();
                    }
                    else if (string.Equals(type, "separate", StringComparison.OrdinalIgnoreCase))
                    {
                        if (capturing && buffer.Length > 0)
                        {
                            yield return buffer.ToString();
                        }
                        capturing = false;
                        buffer.Clear();
                    }
                    else if (string.Equals(type, "end", StringComparison.OrdinalIgnoreCase))
                    {
                        capturing = false;
                        buffer.Clear();
                    }
                }
                else if (capturing && element.Name == W.instrText)
                {
                    var text = element.Value;
                    if (!string.IsNullOrEmpty(text))
                    {
                        buffer.Append(text);
                    }
                }
            }
        }

        private static readonly Regex s_MergeFieldRegex = new(@"MERGEFIELD\s+(?:""(?<quoted>[^""]+)""|(?<simple>[^\s\\]+))", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static string? ParseMailMergeFieldName(string instruction)
        {
            if (string.IsNullOrWhiteSpace(instruction))
            {
                return null;
            }

            var match = s_MergeFieldRegex.Match(instruction);
            if (!match.Success)
            {
                return null;
            }

            return match.Groups["quoted"].Success
                ? match.Groups["quoted"].Value.Trim()
                : match.Groups["simple"].Value.Trim();
        }

        private static string? NormalizeMailMergeXPath(string fieldName)
        {
            if (string.IsNullOrWhiteSpace(fieldName))
            {
                return null;
            }

            var normalized = fieldName.Trim().Trim('"');
            normalized = normalized.Replace("\\", "/");
            normalized = normalized.Replace(".", "/");
            normalized = normalized.Replace(" ", string.Empty);
            normalized = normalized.Trim('/');

            if (string.IsNullOrWhiteSpace(normalized))
            {
                return null;
            }

            return SanitizeXPath(normalized);
        }
    }
}
