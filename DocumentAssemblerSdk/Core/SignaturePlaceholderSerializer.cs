using System;
using System.Text;
using System.Text.Json;

namespace DocumentAssembler.Core;

/// <summary>
/// Serializza e deserializza i metadati dei punti firma inseriti nei documenti DOCX.
/// I metadati vengono codificati in una stringa Base64 racchiusa tra i delimitatori
/// [[DA_SIGN:: ... ]], così da poter essere individuati dopo la conversione in PDF.
/// </summary>
public static class SignaturePlaceholderSerializer
{
    public const string PlaceholderPrefix = "[[DA_SIGN::";
    public const string PlaceholderSuffix = "]]";

    private static readonly JsonSerializerOptions s_JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = false
    };

    public static string CreatePlaceholder(SignaturePlaceholderMetadata metadata)
    {
        if (metadata == null)
        {
            throw new ArgumentNullException(nameof(metadata));
        }

        var json = JsonSerializer.Serialize(metadata, s_JsonOptions);
        var payload = Convert.ToBase64String(Encoding.UTF8.GetBytes(json));
        return $"{PlaceholderPrefix}{payload}{PlaceholderSuffix}";
    }

    public static bool TryParse(string text, out SignaturePlaceholderMetadata? metadata)
    {
        metadata = null;
        if (string.IsNullOrWhiteSpace(text) ||
            !text.StartsWith(PlaceholderPrefix, StringComparison.Ordinal) ||
            !text.EndsWith(PlaceholderSuffix, StringComparison.Ordinal))
        {
            return false;
        }

        try
        {
            var payload = text.Substring(PlaceholderPrefix.Length,
                text.Length - PlaceholderPrefix.Length - PlaceholderSuffix.Length);
            var bytes = Convert.FromBase64String(payload);
            metadata = JsonSerializer.Deserialize<SignaturePlaceholderMetadata>(bytes, s_JsonOptions);
            return metadata != null;
        }
        catch
        {
            metadata = null;
            return false;
        }
    }
}

/// <summary>
/// DTO serializzato dentro il placeholder dei punti firma.
/// Width/Height sono espressi in punti tipografici (1/72 di pollice).
/// </summary>
public sealed record SignaturePlaceholderMetadata(
    string Id,
    string Label,
    double? WidthPoints,
    double? HeightPoints,
    int? PageHint);
