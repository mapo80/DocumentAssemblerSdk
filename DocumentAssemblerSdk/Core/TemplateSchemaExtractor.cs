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
        private static readonly Regex s_XmlTagRegex = new Regex(@"<(Content|Image|Table|Repeat|Conditional)\s+([^/>]+)(?:/?>)", RegexOptions.Compiled);

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
            var result = new SchemaExtractionResult
            {
                Fields = fields.Values.OrderBy(f => f.XPath).ToList()
            };

            result.XmlTemplate = BuildXmlTemplate(result.Fields, repeatingPaths);
            result.RootElementName = DetermineRootElementName(result.Fields);

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
        /// Builds XML template from discovered fields with optimal performance
        /// </summary>
        private static string BuildXmlTemplate(List<FieldInfo> fields, HashSet<string> repeatingPaths)
        {
            if (fields.Count == 0)
            {
                return "<Data />";
            }

            // Build hierarchical structure efficiently using a tree
            var root = new XmlNode { Name = "Data", Children = new Dictionary<string, XmlNode>(StringComparer.OrdinalIgnoreCase) };

            foreach (var field in fields)
            {
                // Skip Conditional and other non-data tags
                if (field.TagType == "Conditional") continue;

                var parts = field.XPath.Split('/');
                var currentNode = root;

                for (int i = 0; i < parts.Length; i++)
                {
                    var part = parts[i].Trim();
                    if (string.IsNullOrWhiteSpace(part)) continue;

                    if (!currentNode.Children.ContainsKey(part))
                    {
                        currentNode.Children[part] = new XmlNode
                        {
                            Name = part,
                            Children = new Dictionary<string, XmlNode>(StringComparer.OrdinalIgnoreCase),
                            IsRepeating = repeatingPaths.Contains(string.Join("/", parts.Take(i + 1)))
                        };
                    }

                    currentNode = currentNode.Children[part];

                    // Mark as leaf if this is the last part and it's a content field
                    if (i == parts.Length - 1 && (field.TagType == "Content" || field.TagType == "Image"))
                    {
                        currentNode.IsLeaf = true;
                        currentNode.IsOptional = field.IsOptional;
                        currentNode.FieldType = field.TagType;
                    }
                }
            }

            // Generate XML string efficiently
            var sb = new StringBuilder();
            BuildXmlString(root, sb, 0);
            return sb.ToString();
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
                if (parts.Length > 0 && !string.IsNullOrWhiteSpace(parts[0]))
                {
                    var root = parts[0].Trim();
                    rootCounts[root] = rootCounts.GetValueOrDefault(root, 0) + 1;
                }
            }

            if (rootCounts.Count > 0)
            {
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
            public bool IsOptional { get; set; } = true;
            public string FieldType { get; set; } = string.Empty;
        }
    }
}
