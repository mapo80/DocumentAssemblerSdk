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
        // Regex for custom format: <#Content Select="..."#>
        private static readonly Regex s_TagRegex = new Regex(@"<#\s*(\w+)\s+([^#]+)#>", RegexOptions.Compiled);

        // Regex for XML format: <Content Select="..." /> or <Content Select="...">
        // Match tag name, then attributes (everything except > unless in quotes), then optional / before >
        private static readonly Regex s_XmlTagRegex = new Regex(@"<(Content|Image|Table|Repeat|Conditional)\s+([^>]+?)(\s*/?>)", RegexOptions.Compiled);

        private static readonly Regex s_AttributeRegex = new Regex(@"(\w+)\s*=\s*[""']([^""']+)[""']", RegexOptions.Compiled);

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
        }

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
            if (xDoc.Root == null) return;

            // Fast extraction from content controls
            var contentControls = xDoc.Descendants(W.sdt).ToList();
            foreach (var sdt in contentControls)
            {
                var text = string.Concat(sdt.Descendants(W.t).Select(t => t.Value))
                    .Trim()
                    .Replace('\u201C', '"')
                    .Replace('\u201D', '"')
                    .Replace('\u2018', '\'')
                    .Replace('\u2019', '\'');

                // Handle HTML escaped entities (common in Word documents)
                text = System.Net.WebUtility.HtmlDecode(text);

                if (text.StartsWith("<#") && text.EndsWith("#>"))
                {
                    ParseAndAddField(text, fields, repeatingPaths);
                }
                else if (text.Contains("<Content") || text.Contains("<Image") || text.Contains("<Table") ||
                         text.Contains("<Repeat") || text.Contains("<Conditional") || text.Contains("<Else") || text.Contains("<EndRepeat") || text.Contains("<EndConditional"))
                {
                    // Also try XML format: <Content ... />, <Else>, etc.
                    var xmlMatches = s_XmlTagRegex.Matches(text);
                    foreach (Match match in xmlMatches)
                    {
                        ParseAndAddXmlField(match, fields, repeatingPaths);
                    }

                    // Handle special tags without Select attribute
                    if (text.Contains("<Else>") || text.Contains("<Else />"))
                    {
                        // Else tags don't have Select attribute, skip them
                    }
                    if (text.Contains("<EndRepeat>") || text.Contains("<EndRepeat />") ||
                        text.Contains("<EndConditional>") || text.Contains("<EndConditional />"))
                    {
                        // End tags don't have Select attribute, skip them
                    }
                }
            }

            // Fast extraction from plain paragraphs
            var paragraphs = xDoc.Descendants(W.p).ToList();
            foreach (var para in paragraphs)
            {
                var text = string.Concat(para.Descendants(W.t).Select(t => t.Value))
                    .Trim();

                // Handle HTML escaped entities (common in Word documents)
                text = System.Net.WebUtility.HtmlDecode(text);

                // Try custom format first: <#Content ...#>
                if (text.Contains("<#"))
                {
                    var matches = s_TagRegex.Matches(text);
                    foreach (Match match in matches)
                    {
                        ParseAndAddField(match.Value, fields, repeatingPaths);
                    }
                }

                // Also try XML format: <Content ... />
                if (text.Contains("<Content") || text.Contains("<Image") || text.Contains("<Table") ||
                    text.Contains("<Repeat") || text.Contains("<Conditional"))
                {
                    var xmlMatches = s_XmlTagRegex.Matches(text);
                    foreach (Match match in xmlMatches)
                    {
                        ParseAndAddXmlField(match, fields, repeatingPaths);
                    }
                }
            }
        }

        /// <summary>
        /// Parses an XML format tag and adds it to the fields collection
        /// </summary>
        private static void ParseAndAddXmlField(Match match, List<FieldInfo> fields, HashSet<string> repeatingPaths)
        {
            var tagName = match.Groups[1].Value; // Content, Image, Table, etc.
            var attributesText = match.Groups[2].Value;

            // Parse attributes efficiently
            var attributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var attrMatches = s_AttributeRegex.Matches(attributesText);
            foreach (Match attrMatch in attrMatches)
            {
                attributes[attrMatch.Groups[1].Value] = attrMatch.Groups[2].Value;
            }

            // Extract Select attribute (XPath)
            if (!attributes.TryGetValue("Select", out var xpath) || string.IsNullOrWhiteSpace(xpath))
            {
                return; // No Select attribute, skip
            }

            // Sanitize XPath: trim and remove any trailing slashes or invalid characters
            xpath = xpath.Trim().TrimEnd('/', '>');

            // Skip if XPath is invalid after sanitization
            if (string.IsNullOrWhiteSpace(xpath))
            {
                return;
            }

            // Determine if optional
            var isOptional = true; // Default is true
            if (attributes.TryGetValue("Optional", out var optionalStr))
            {
                isOptional = !bool.TryParse(optionalStr, out var optVal) || optVal;
            }

            var field = new FieldInfo
            {
                XPath = xpath,
                TagType = tagName,
                IsOptional = isOptional,
                Attributes = attributes
            };

            // Handle Repeat and Table tags (mark as repeating)
            if (tagName == "Repeat" || tagName == "Table")
            {
                repeatingPaths.Add(xpath);
                field.IsRepeating = true;
            }

            fields.Add(field);
        }

        /// <summary>
        /// Parses a custom format tag and adds it to the fields collection (optimized)
        /// </summary>
        private static void ParseAndAddField(string tagText, List<FieldInfo> fields, HashSet<string> repeatingPaths)
        {
            // Remove <# and #> delimiters
            var content = tagText.Trim();
            if (content.StartsWith("<#")) content = content.Substring(2);
            if (content.EndsWith("#>")) content = content.Substring(0, content.Length - 2);
            content = content.Trim();

            // Fast parse: split on whitespace to get tag name
            var firstSpace = content.IndexOf(' ');
            if (firstSpace == -1) return; // No attributes

            var tagName = content.Substring(0, firstSpace);
            var attributesText = content.Substring(firstSpace + 1);

            // Skip end tags
            if (tagName.StartsWith("End") || tagName == "Else") return;

            // Parse attributes efficiently
            var attributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var attrMatches = s_AttributeRegex.Matches(attributesText);
            foreach (Match attrMatch in attrMatches)
            {
                attributes[attrMatch.Groups[1].Value] = attrMatch.Groups[2].Value;
            }

            // Extract Select attribute (XPath)
            if (!attributes.TryGetValue("Select", out var xpath) || string.IsNullOrWhiteSpace(xpath))
            {
                return; // No Select attribute, skip
            }

            // Sanitize XPath: trim and remove any trailing slashes or invalid characters
            xpath = xpath.Trim().TrimEnd('/', '>');

            // Skip if XPath is invalid after sanitization
            if (string.IsNullOrWhiteSpace(xpath))
            {
                return;
            }

            // Determine if optional
            var isOptional = true; // Default is true
            if (attributes.TryGetValue("Optional", out var optionalStr))
            {
                isOptional = !bool.TryParse(optionalStr, out var optVal) || optVal;
            }

            var field = new FieldInfo
            {
                XPath = xpath,
                TagType = tagName,
                IsOptional = isOptional,
                Attributes = attributes
            };

            // Handle Repeat and Table tags (mark as repeating)
            if (tagName == "Repeat" || tagName == "Table")
            {
                repeatingPaths.Add(xpath);
                field.IsRepeating = true;
            }

            fields.Add(field);
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
                if (field.TagType == "Conditional") continue;

                var parts = field.XPath.Split('/')
                    .Select(p => p.Trim())
                    .Where(p => !string.IsNullOrWhiteSpace(p) && p != "." && !p.StartsWith("-") && !char.IsDigit(p[0]))
                    .ToList();

                if (parts.Count == 0) continue;

                var currentNode = root;
                var pathSegments = new List<string>();

                for (int i = 0; i < parts.Count; i++)
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

                    var isLeafNode = i == parts.Count - 1;
                    if (isLeafNode)
                    {
                        currentNode.IsOptional = field.IsOptional;
                        if (field.TagType == "Content" || field.TagType == "Image")
                        {
                            currentNode.IsLeaf = true;
                            currentNode.FieldType = field.TagType;
                        }
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

            if (node.IsLeaf)
            {
                var type = node.FieldType == "Image" ? "xs:base64Binary" : "xs:string";
                sb.AppendLine($"{indentStr}<xs:element name=\"{node.Name}\" type=\"{type}\"{occurrenceAttrs} />");
                return;
            }

            if (node.Children.Count == 0)
            {
                sb.AppendLine($"{indentStr}<xs:element name=\"{node.Name}\" type=\"xs:string\"{occurrenceAttrs} />");
                return;
            }

            sb.AppendLine($"{indentStr}<xs:element name=\"{node.Name}\"{occurrenceAttrs}>");
            sb.AppendLine($"{indentStr}  <xs:complexType>");
            sb.AppendLine($"{indentStr}    <xs:sequence>");
            foreach (var child in node.Children.Values.OrderBy(n => n.Name))
            {
                BuildXsdElement(child, sb, indent + 3, false);
            }
            sb.AppendLine($"{indentStr}    </xs:sequence>");
            sb.AppendLine($"{indentStr}  </xs:complexType>");
            sb.AppendLine($"{indentStr}</xs:element>");
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
            var indentStr = new string(' ', indent * 2);
            var hasChildren = node.Children.Count > 0;

            if (node.Name == "Data")
            {
                // Root element
                sb.AppendLine("<Data>");
                foreach (var child in node.Children.Values.OrderBy(n => n.Name))
                {
                    BuildXmlString(child, sb, indent + 1);
                }
                sb.Append("</Data>");
            }
            else if (node.IsLeaf)
            {
                // Leaf node (Content or Image)
                var comment = node.IsOptional ? " <!-- Optional -->" : "";
                if (node.FieldType == "Image")
                {
                    comment = node.IsOptional ? " <!-- Optional, Base64 encoded image -->" : " <!-- Base64 encoded image -->";
                }
                sb.AppendLine($"{indentStr}<{node.Name}>[value]{comment}</{node.Name}>");
            }
            else if (node.IsRepeating)
            {
                // Repeating element
                sb.AppendLine($"{indentStr}<{node.Name}> <!-- Repeating -->");
                foreach (var child in node.Children.Values.OrderBy(n => n.Name))
                {
                    BuildXmlString(child, sb, indent + 1);
                }
                sb.AppendLine($"{indentStr}</{node.Name}>");

                // Add second example for clarity
                sb.AppendLine($"{indentStr}<{node.Name}> <!-- Repeating -->");
                foreach (var child in node.Children.Values.OrderBy(n => n.Name))
                {
                    BuildXmlString(child, sb, indent + 1);
                }
                sb.AppendLine($"{indentStr}</{node.Name}>");
            }
            else if (hasChildren)
            {
                // Container element
                sb.AppendLine($"{indentStr}<{node.Name}>");
                foreach (var child in node.Children.Values.OrderBy(n => n.Name))
                {
                    BuildXmlString(child, sb, indent + 1);
                }
                sb.AppendLine($"{indentStr}</{node.Name}>");
            }
            else
            {
                // Empty element
                sb.AppendLine($"{indentStr}<{node.Name} />");
            }
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
        }
    }
}
