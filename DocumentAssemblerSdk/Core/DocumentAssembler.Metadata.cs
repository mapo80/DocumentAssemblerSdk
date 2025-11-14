using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Schema;

namespace DocumentAssembler.Core
{
    /// <summary>
    /// DocumentAssembler partial class - Metadata transformation functionality
    /// </summary>
    public partial class DocumentAssembler
    {
        private static readonly List<string> s_AliasList = new List<string>()
        {
            "Content",
            "Table",
            "Repeat",
            "EndRepeat",
            "Conditional",
            "Else",
            "EndConditional",
            "Signature",
        };

        private static object TransformToMetadata(XNode node, XElement data, TemplateError te)
        {
            if (node is XElement element)
            {
                if (element.Name == W.sdt)
                {
                    var alias = (string?)element.Elements(W.sdtPr).Elements(W.alias).Attributes(W.val).FirstOrDefault();
                    if (alias == null || alias == "" || s_AliasList.Contains(alias))
                    {
                        var ccContents = element
                            .DescendantsTrimmed(W.txbxContent)
                            .Where(e => e.Name == W.t)
                            .Select(t => (string?)t)
                            .Where(s => s != null)
                            .Select(s => s!)
                            .StringConcatenate()
                            .Trim()
                            .Replace('\u201C', '"')
                            .Replace('\u201D', '"')
                            .Replace('\u2018', '\'')
                            .Replace('\u2019', '\'');
                        if (ccContents.StartsWith("<"))
                        {
                            var xml = TransformXmlTextToMetadata(te, ccContents);
                            if (xml.Name == W.p || xml.Name == W.r)  // this means there was an error processing the XML.
                            {
                                if (element.Parent?.Name == W.p)
                                {
                                    return xml.Elements(W.r);
                                }

                                return xml;
                            }
                            if (alias != null && xml.Name.LocalName != alias)
                            {
                                if (element.Parent?.Name == W.p)
                                {
                                    return CreateRunErrorMessage("Error: Content control alias does not match metadata element name", te);
                                }
                                else
                                {
                                    return CreateParaErrorMessage("Error: Content control alias does not match metadata element name", te);
                                }
                            }
                            xml.Add(element.Elements(W.sdtContent).Elements());
                            return xml;
                        }
                        return new XElement(element.Name,
                            element.Attributes(),
                            element.Nodes().Select(n => TransformToMetadata(n, data, te)));
                    }
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Nodes().Select(n => TransformToMetadata(n, data, te)));
                }
                if (element.Name == W.p)
                {
                    var paraContents = element
                        .DescendantsTrimmed(W.txbxContent)
                        .Where(e => e.Name == W.t)
                        .Select(t => (string)t)
                        .StringConcatenate()
                        .Trim();
                    var occurances = paraContents.Select((c, i) => paraContents.Substring(i)).Count(sub => sub.StartsWith("<#"));
                    if (paraContents.StartsWith("<#") && paraContents.EndsWith("#>") && occurances == 1)
                    {
                        var xmlText = paraContents.Substring(2, paraContents.Length - 4).Trim()
                            .Replace('\u201C', '"').Replace('\u201D', '"')
                            .Replace('\u2018', '\'').Replace('\u2019', '\'');
                        var xml = TransformXmlTextToMetadata(te, xmlText);
                        if (xml.Name == W.p || xml.Name == W.r)
                        {
                            return xml;
                        }

                        xml.Add(element);
                        return xml;
                    }
                    if (paraContents.Contains("<#"))
                    {
                        var runReplacementInfo = new List<RunReplacementInfo>();
                        var thisGuid = Guid.NewGuid().ToString();
                        var r = new Regex("<#.*?#>");
                        XElement? xml = null;
                        OpenXmlRegex.Replace(new[] { element }, r, thisGuid, (para, match) =>
                        {
                            var matchString = match.Value.Trim();
                            var xmlText = matchString.Substring(2, matchString.Length - 4).Trim()
                                .Replace('\u201C', '"').Replace('\u201D', '"')
                                .Replace('\u2018', '\'').Replace('\u2019', '\'');
                            try
                            {
                                xml = XElement.Parse(xmlText);
                            }
                            catch (XmlException e)
                            {
                                var rri = new RunReplacementInfo()
                                {
                                    Xml = null,
                                    XmlExceptionMessage = "XmlException: " + e.Message,
                                    SchemaValidationMessage = null,
                                };
                                runReplacementInfo.Add(rri);
                                return true;
                            }
                            var schemaError = ValidatePerSchema(xml);
                            if (schemaError != null)
                            {
                                var rri = new RunReplacementInfo()
                                {
                                    Xml = null,
                                    XmlExceptionMessage = null,
                                    SchemaValidationMessage = "Schema Validation Error: " + schemaError,
                                };
                                runReplacementInfo.Add(rri);
                                return true;
                            }
                            var rri2 = new RunReplacementInfo()
                            {
                                Xml = xml,
                                XmlExceptionMessage = null,
                                SchemaValidationMessage = null,
                            };
                            runReplacementInfo.Add(rri2);
                            return true;
                        }, false);

                        var newPara = new XElement(element);
                        foreach (var rri in runReplacementInfo)
                        {
                            var runToReplace = newPara.Descendants(W.r).FirstOrDefault(rn => rn.Value == thisGuid && rn.Parent?.Name != PA.Content);
                            if (runToReplace == null)
                            {
                                throw new OpenXmlPowerToolsException("Internal error");
                            }

                            if (rri.XmlExceptionMessage != null)
                            {
                                runToReplace.ReplaceWith(CreateRunErrorMessage(rri.XmlExceptionMessage, te));
                            }
                            else if (rri.SchemaValidationMessage != null)
                            {
                                runToReplace.ReplaceWith(CreateRunErrorMessage(rri.SchemaValidationMessage, te));
                            }
                            else if (rri.Xml != null)
                            {
                                var newXml = new XElement(rri.Xml);
                                newXml.Add(runToReplace);
                                runToReplace.ReplaceWith(newXml);
                            }
                        }
                        var coalescedParagraph = WordprocessingMLUtil.CoalesceAdjacentRunsWithIdenticalFormatting(newPara);
                        return coalescedParagraph;
                    }
                }

                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => TransformToMetadata(n, data, te)));
            }
            return node;
        }

        private static XElement TransformXmlTextToMetadata(TemplateError te, string xmlText)
        {
            XElement xml;
            try
            {
                xml = XElement.Parse(xmlText);
            }
            catch (XmlException e)
            {
                return CreateParaErrorMessage("XmlException: " + e.Message, te);
            }
            var schemaError = ValidatePerSchema(xml);
            if (schemaError != null)
            {
                return CreateParaErrorMessage("Schema Validation Error: " + schemaError, te);
            }

            return xml;
        }

        private class RunReplacementInfo
        {
            public XElement? Xml;
            public string? XmlExceptionMessage;
            public string? SchemaValidationMessage;
        }

        private static string? ValidatePerSchema(XElement element)
        {
            if (s_PASchemaSets == null)
            {
                var schemaSets = new Dictionary<XName, PASchemaSet>()
                {
                    {
                        PA.Content,
                        new PASchemaSet() {
                            XsdMarkup =
                              @"<xs:schema attributeFormDefault='unqualified' elementFormDefault='qualified' xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                  <xs:element name='Content'>
                                    <xs:complexType>
                                      <xs:attribute name='Select' type='xs:string' use='required' />
                                      <xs:attribute name='Optional' type='xs:boolean' use='optional' />
                                    </xs:complexType>
                                  </xs:element>
                                </xs:schema>",
                        }
                    },
                    {
                        PA.Image,
                        new PASchemaSet() {
                            XsdMarkup =
                              @"<xs:schema attributeFormDefault='unqualified' elementFormDefault='qualified' xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                  <xs:element name='Image'>
                                    <xs:complexType>
                                      <xs:attribute name='Select' type='xs:string' use='required' />
                                      <xs:attribute name='Optional' type='xs:boolean' use='optional' />
                                      <xs:attribute name='Align' type='xs:string' use='optional' />
                                      <xs:attribute name='Width' type='xs:string' use='optional' />
                                      <xs:attribute name='Height' type='xs:string' use='optional' />
                                      <xs:attribute name='MaxWidth' type='xs:string' use='optional' />
                                      <xs:attribute name='MaxHeight' type='xs:string' use='optional' />
                                    </xs:complexType>
                                  </xs:element>
                                </xs:schema>",
                        }
                    },
                    {
                        PA.Table,
                        new PASchemaSet() {
                            XsdMarkup =
                              @"<xs:schema attributeFormDefault='unqualified' elementFormDefault='qualified' xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                  <xs:element name='Table'>
                                    <xs:complexType>
                                      <xs:attribute name='Select' type='xs:string' use='required' />
                                    </xs:complexType>
                                  </xs:element>
                                </xs:schema>",
                        }
                    },
                    {
                        PA.Repeat,
                        new PASchemaSet() {
                            XsdMarkup =
                              @"<xs:schema attributeFormDefault='unqualified' elementFormDefault='qualified' xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                  <xs:element name='Repeat'>
                                    <xs:complexType>
                                      <xs:attribute name='Select' type='xs:string' use='required' />
                                      <xs:attribute name='Optional' type='xs:boolean' use='optional' />
                                    </xs:complexType>
                                  </xs:element>
                                </xs:schema>",
                        }
                    },
                    {
                        PA.EndRepeat,
                        new PASchemaSet() {
                            XsdMarkup =
                              @"<xs:schema attributeFormDefault='unqualified' elementFormDefault='qualified' xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                  <xs:element name='EndRepeat' />
                                </xs:schema>",
                        }
                    },
                    {
                        PA.Conditional,
                        new PASchemaSet() {
                            XsdMarkup =
                              @"<xs:schema attributeFormDefault='unqualified' elementFormDefault='qualified' xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                  <xs:element name='Conditional'>
                                    <xs:complexType>
                                      <xs:attribute name='Select' type='xs:string' use='required' />
                                      <xs:attribute name='Match' type='xs:string' use='optional' />
                                      <xs:attribute name='NotMatch' type='xs:string' use='optional' />
                                    </xs:complexType>
                                  </xs:element>
                                </xs:schema>",
                        }
                    },
                    {
                        PA.Else,
                        new PASchemaSet() {
                            XsdMarkup =
                              @"<xs:schema attributeFormDefault='unqualified' elementFormDefault='qualified' xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                  <xs:element name='Else' />
                                </xs:schema>",
                        }
                    },
                    {
                        PA.EndConditional,
                        new PASchemaSet() {
                            XsdMarkup =
                              @"<xs:schema attributeFormDefault='unqualified' elementFormDefault='qualified' xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                  <xs:element name='EndConditional' />
                                </xs:schema>",
                        }
                    },
                    {
                        PA.Signature,
                        new PASchemaSet()
                        {
                            XsdMarkup =
                              @"<xs:schema attributeFormDefault='unqualified' elementFormDefault='qualified' xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                  <xs:element name='Signature'>
                                    <xs:complexType>
                                      <xs:attribute name='Id' type='xs:string' use='required' />
                                      <xs:attribute name='Label' type='xs:string' use='optional' />
                                      <xs:attribute name='Width' type='xs:string' use='optional' />
                                      <xs:attribute name='Height' type='xs:string' use='optional' />
                                      <xs:attribute name='PageHint' type='xs:int' use='optional' />
                                      <xs:attribute name='Optional' type='xs:boolean' use='optional' />
                                    </xs:complexType>
                                  </xs:element>
                                </xs:schema>",
                        }
                    },
                };
                foreach (var item in schemaSets)
                {
                    var itemPAss = item.Value;
                    var schemas = new XmlSchemaSet();
                    schemas.Add("", XmlReader.Create(new StringReader(itemPAss.XsdMarkup)));
                    itemPAss.SchemaSet = schemas;
                }
                s_PASchemaSets = schemaSets;
            }
            if (!s_PASchemaSets.ContainsKey(element.Name))
            {
                return string.Format("Invalid XML: {0} is not a valid element", element.Name.LocalName);
            }
            var paSchemaSet = s_PASchemaSets[element.Name];
            if (paSchemaSet.SchemaSet == null)
            {
                return "Internal error: Schema set not initialized";
            }

            var d = new XDocument(element);
            string? message = null;
            d.Validate(paSchemaSet.SchemaSet, (sender, e) =>
            {
                if (message == null)
                {
                    message = e.Message;
                }
            }, true);
            if (message != null)
            {
                return message;
            }

            return null;
        }

        private static Dictionary<XName, PASchemaSet>? s_PASchemaSets;

        private class PASchemaSet
        {
            public string XsdMarkup = string.Empty;
            public XmlSchemaSet? SchemaSet;
        }
    }
}
