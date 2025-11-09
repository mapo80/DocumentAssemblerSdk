#!/usr/bin/env dotnet-script
#r "nuget: DocumentFormat.OpenXml, 3.3.0"

using System;
using System.IO;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

// Load test files
var testDir = "/Users/politom/Documents/Workspace/personal/OpenXmlPowerTools/DocumentAssembler/DocumentAssemblerSdk.Tests/TestFiles";
var templatePath = Path.Combine(testDir, "DA018-SmartQuotes.docx");
var dataPath = Path.Combine(testDir, "DA-Data.xml");

Console.WriteLine($"Template: {templatePath}");
Console.WriteLine($"Data: {dataPath}");
Console.WriteLine($"Template exists: {File.Exists(templatePath)}");
Console.WriteLine($"Data exists: {File.Exists(dataPath)}");

// Open template and check for smart quotes
if (File.Exists(templatePath))
{
    using (var doc = WordprocessingDocument.Open(templatePath, false))
    {
        var mainPart = doc.MainDocumentPart;
        if (mainPart != null)
        {
            using (var reader = new StreamReader(mainPart.GetStream()))
            {
                var content = reader.ReadToEnd();
                Console.WriteLine("\n=== Template Content (first 2000 chars) ===");
                Console.WriteLine(content.Substring(0, Math.Min(2000, content.Length)));

                // Check for smart quotes
                if (content.Contains('"') || content.Contains('"'))
                {
                    Console.WriteLine("\n✓ Template contains smart quotes");
                }
            }
        }
    }
}

// Load and display XML data
if (File.Exists(dataPath))
{
    var xmlData = XElement.Load(dataPath);
    Console.WriteLine("\n=== XML Data ===");
    Console.WriteLine(xmlData.ToString().Substring(0, Math.Min(1000, xmlData.ToString().Length)));
}
