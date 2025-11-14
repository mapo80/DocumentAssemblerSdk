using System;
using DocumentAssembler.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Example09_AllTags;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Xunit;

namespace DocumentAssemblerSdk.Tests;

public class Example09AllTagsTests
{
    [Fact]
    public void Schema_ShouldExpose_All_Complex_Paths_And_Json_Mapping()
    {
        var templateDoc = new WmlDocument(GetTemplatePath());
        var schema = TemplateSchemaExtractor.ExtractXmlSchema(templateDoc);
        Assert.Equal("Report", schema.RootElementName);

        var expectedPaths = new[]
        {
            "Report/Customer/FullName",
            "Report/Customer/LoyaltyScore",
            "Report/Highlights/Highlight/Title",
            "Report/Departments/Department/Budget/Allocated",
            "Report/Departments/Department/Achievements/Achievement",
            "Report/Orders/Order/@code",
            "Report/Orders/Order/DeliveryWindow/Start",
            "Report/Milestones/Milestone/@code",
            "Report/Attachments/Attachment/Label",
            "Report/Approvals/PrimarySigner"
        };

        var flattened = schema.Fields.Select(f => f.XPath).ToList();
        foreach (var path in expectedPaths)
        {
            Assert.Contains(path, flattened, StringComparer.OrdinalIgnoreCase);
        }

        var sampleData = Example09DataFactory.CreateSampleData();
        var missingPaths = schema.Fields
            .Where(field => !DataContainsPath(sampleData, field.XPath))
            .Select(field => field.XPath)
            .ToList();
        Assert.Empty(missingPaths);

        var jsonSample = BuildJsonSample(schema.XmlTemplate);
        using var json = JsonDocument.Parse(jsonSample);
        Assert.True(json.RootElement.TryGetProperty("Report", out var report));
        Assert.True(report.TryGetProperty("Customer", out var _));
        Assert.True(report.TryGetProperty("Highlights", out var highlights));
        Assert.True(highlights.TryGetProperty("Highlight", out var highlightArray));
        Assert.Equal(JsonValueKind.Array, highlightArray.ValueKind);
        Assert.True(report.TryGetProperty("Departments", out var departments));
        Assert.True(departments.TryGetProperty("Department", out var departmentArray));
        Assert.Equal(JsonValueKind.Array, departmentArray.ValueKind);
    }

    [Fact]
    public void DocumentAssembler_Should_Render_All_Sections()
    {
        var templateDoc = new WmlDocument(GetTemplatePath());
        var sampleData = Example09DataFactory.CreateSampleData();

        var assembled = DocumentAssembler.Core.DocumentAssembler.AssembleDocument(
            templateDoc,
            sampleData,
            out var templateError,
            out var templateSummary);

        Assert.False(templateError, templateSummary);

        var text = ExtractPlainText(assembled);
        Assert.Contains("Mario Rossi", text, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Innovation Sprint", text, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Digital Factory", text, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("ORD-2042", text, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Accesso prioritario a laboratori", text, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Deploy piattaforma wave 2", text, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Executive memo", text, StringComparison.OrdinalIgnoreCase);
    }

    private static string GetTemplatePath()
    {
        var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
        return Path.Combine(
            repoRoot,
            "DocumentAssembler",
            "DocumentAssemblerSdk.Examples",
            "Example09_AllTags",
            "TemplateAllTagsDocument.docx");
    }

    private static string ExtractPlainText(WmlDocument document)
    {
        using var ms = new MemoryStream();
        ms.Write(document.DocumentByteArray, 0, document.DocumentByteArray.Length);
        ms.Position = 0;

        using var word = WordprocessingDocument.Open(ms, false);
        var builder = new StringBuilder();
        foreach (var text in word.MainDocumentPart!.Document.Body!.Descendants<Text>())
        {
            if (!string.IsNullOrWhiteSpace(text.Text))
            {
                builder.Append(text.Text);
                builder.Append(' ');
            }
        }

        return builder.ToString();
    }

    private static bool DataContainsPath(XElement root, string xpath)
    {
        if (string.IsNullOrWhiteSpace(xpath))
        {
            return true;
        }

        var segments = xpath.Split('/', StringSplitOptions.RemoveEmptyEntries);
        IEnumerable<XElement> current = new[] { root };

        foreach (var segment in segments)
        {
            if (segment == ".")
            {
                continue;
            }

            if (segment.StartsWith("@", StringComparison.Ordinal))
            {
                var attrName = segment[1..];
                return current.Any(e => e.Attribute(attrName) != null);
            }

            current = current.SelectMany(e => e.Elements(segment));
            if (!current.Any())
            {
                return false;
            }
        }

        return true;
    }

    private static string BuildJsonSample(string xmlTemplate)
    {
        if (string.IsNullOrWhiteSpace(xmlTemplate))
        {
            return "{}";
        }

        var document = XDocument.Parse(xmlTemplate);
        if (document.Root == null)
        {
            return "{}";
        }

        var payload = ConvertElement(document.Root);
        return JsonSerializer.Serialize(payload, new JsonSerializerOptions
        {
            WriteIndented = true
        });
    }

    private static object ConvertElement(XElement element)
    {
        if (!element.HasElements)
        {
            return "sample";
        }

        var groups = element.Elements().GroupBy(e => e.Name.LocalName, StringComparer.Ordinal);
        var dict = new Dictionary<string, object?>(StringComparer.Ordinal);

        foreach (var group in groups)
        {
            var children = group.Select(ConvertElement).ToList();
            dict[group.Key] = children.Count == 1 ? children[0] : children;
        }

        return dict;
    }
}
