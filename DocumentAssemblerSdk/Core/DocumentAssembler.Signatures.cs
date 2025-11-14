using System;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace DocumentAssembler.Core
{
    public partial class DocumentAssembler
    {
        private const string DefaultSignatureLabel = "Signature";

        private static object? BuildSignaturePlaceholder(XElement element, TemplateError templateError)
        {
            var id = (string?)element.Attribute(PA.Id);
            if (string.IsNullOrWhiteSpace(id))
            {
                return CreateContextErrorMessage(element, "Signature: Id attribute is required", templateError);
            }

            var label = (string?)element.Attribute(PA.Label);
            if (string.IsNullOrWhiteSpace(label))
            {
                label = DefaultSignatureLabel;
            }

            int? pageHint = null;
            var pageHintAttr = (string?)element.Attribute(PA.PageHint);
            if (!string.IsNullOrWhiteSpace(pageHintAttr))
            {
                if (int.TryParse(pageHintAttr, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsedPage) && parsedPage > 0)
                {
                    pageHint = parsedPage;
                }
                else
                {
                    return CreateContextErrorMessage(element, "Signature: PageHint must be a positive integer", templateError);
                }
            }

            if (!TryConvertLengthToPoints((string?)element.Attribute(PA.Width), out var widthPoints, out var widthError))
            {
                return CreateContextErrorMessage(element, widthError, templateError);
            }

            if (!TryConvertLengthToPoints((string?)element.Attribute(PA.Height), out var heightPoints, out var heightError))
            {
                return CreateContextErrorMessage(element, heightError, templateError);
            }

            var placeholderMetadata = new SignaturePlaceholderMetadata(
                id.Trim(),
                label!.Trim(),
                widthPoints,
                heightPoints,
                pageHint);

            var encodedPlaceholder = SignaturePlaceholderSerializer.CreatePlaceholder(placeholderMetadata);
            var paragraph = element.Descendants(W.p).FirstOrDefault();
            var run = element.Descendants(W.r).FirstOrDefault();

            if (paragraph != null)
            {
                var template = GetParagraphTemplate(paragraph, run ?? paragraph.Elements(W.r).FirstOrDefault());
                var resultParagraph = new XElement(W.p);
                if (template.ParagraphProperties != null)
                {
                    resultParagraph.Add(new XElement(template.ParagraphProperties));
                }

                resultParagraph.Add(CreateSignatureRun(template.RunPrototype, $"{label} ____________________", false));
                resultParagraph.Add(CreateSignatureRun(template.RunPrototype, encodedPlaceholder, true));
                return resultParagraph;
            }

            if (run != null)
            {
                var visibleRun = CreateSignatureRun(run, $"{label} ____________________", false);
                var metaRun = CreateSignatureRun(run, encodedPlaceholder, true);
                return new[] { visibleRun, metaRun };
            }

            return CreateContextErrorMessage(element, "Signature: Unable to find paragraph or run context", templateError);
        }

        private static XElement CreateSignatureRun(XElement runPrototype, string text, bool isMetadata)
        {
            var newRun = new XElement(runPrototype.Name, runPrototype.Attributes(), runPrototype.Elements().Where(e => e.Name != W.t));
            var textNode = new XElement(W.t, text);
            textNode.SetAttributeValue(XNamespace.Xml + "space", "preserve");
            newRun.Add(textNode);

            if (isMetadata)
            {
                var runProperties = newRun.Element(W.rPr);
                if (runProperties == null)
                {
                    runProperties = new XElement(W.rPr);
                    newRun.AddFirst(runProperties);
                }

                // Nasconde visivamente il placeholder, ma lo lascia nel PDF per l'elaborazione successiva
                runProperties.Add(new XElement(W.color, new XAttribute(W.val, "FFFFFF")));
                runProperties.Add(new XElement(W.sz, new XAttribute(W.val, 2)));
                runProperties.Add(new XElement(W.szCs, new XAttribute(W.val, 2)));
            }

            return newRun;
        }

        private static bool TryConvertLengthToPoints(string? rawValue, out double? points, out string errorMessage)
        {
            points = null;
            errorMessage = string.Empty;
            if (string.IsNullOrWhiteSpace(rawValue))
            {
                return true;
            }

            if (!TryParseLengthToEmu(rawValue, out var emuValue, out errorMessage))
            {
                errorMessage = errorMessage.Replace("Image", "Signature", StringComparison.Ordinal);
                return false;
            }

            points = emuValue.HasValue ? emuValue.Value / EmusPerPoint : null;
            return true;
        }
    }
}
