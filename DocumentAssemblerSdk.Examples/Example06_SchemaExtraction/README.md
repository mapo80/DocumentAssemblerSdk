# Example 06: XML Schema Extraction

This example demonstrates how to **extract the required XML schema** from a DocumentAssembler template (.docx). This is incredibly useful when you have a template and need to understand what XML data structure is required to populate it.

## What This Example Does

1. **Analyzes a template document** to find all DocumentAssembler tags (`<#Content#>`, `<#Image#>`, `<#Repeat#>`, etc.)
2. **Extracts field information** including:
   - XPath expressions
   - Tag types (Content, Image, Repeat, Table, Conditional)
   - Whether fields are optional or required
   - Whether structures are repeating
3. **Generates an XML template** with placeholder values and helpful comments
4. **Provides detailed field analysis** grouped by type

## Use Cases

- **Reverse-engineering templates** - Understand what data a template needs
- **API integration** - Generate XML schemas for automated document generation
- **Documentation** - Create data specifications from templates
- **Validation** - Ensure your XML data matches template requirements
- **Onboarding** - Help new developers understand template structure

## Running the Example

```bash
cd DocumentAssemblerSdk.Examples/Example06_SchemaExtraction
dotnet run
```

## Output

The example produces:

1. **Console output** showing:
   - Number of fields discovered
   - Root element name
   - Generated XML template
   - Detailed field analysis by type
   - Summary statistics

2. **XML schema file** (`Schema_TemplateDocument.xml`) containing:
   - Well-formed XML structure
   - Placeholder values (`[value]`)
   - Comments indicating optional fields
   - Comments indicating repeating structures
   - Comments for special field types (e.g., Base64 images)

## Example Output

```
=== Example 06: XML Schema Extraction ===

Example 1: Basic Template Schema Extraction
--------------------------------------------
Analyzing template: TemplateDocument.docx

Discovered 15 field(s)
Root element: Customer

Generated XML Template:
----------------------
<Data>
  <Customer>
    <Name>[value]</Name>
    <Email>[value] <!-- Optional --></Email>
    <Address>
      <Street>[value]</Street>
      <City>[value]</City>
      <Zip>[value]</Zip>
    </Address>
    <Orders> <!-- Repeating -->
      <Product>[value]</Product>
      <Quantity>[value]</Quantity>
    </Orders>
    <Orders> <!-- Repeating -->
      <Product>[value]</Product>
      <Quantity>[value]</Quantity>
    </Orders>
  </Customer>
</Data>

Schema saved to: Schema_TemplateDocument.xml


Example 2: Detailed Field Analysis
-----------------------------------
Analyzing template: TemplateDocument.docx

Content Fields (10):
----------------------------------------
  XPath: Customer/Name
    - Element: Name
    - Optional: True
    - Repeating: False

  XPath: Customer/Email
    - Element: Email
    - Optional: True
    - Repeating: False

  ...

Repeat Fields (1):
----------------------------------------
  XPath: Customer/Orders
    - Element: Orders
    - Optional: True
    - Repeating: True

Summary:
--------
Total fields: 15
Optional fields: 14
Required fields: 1
Repeating structures: 1
Content fields: 10
Image fields: 1
Table fields: 1
Repeat fields: 1
Conditional fields: 2
```

## Performance

The schema extraction is **ultra-fast**:
- **Single-pass algorithm** - reads the document only once
- **Compiled regex** - extremely fast tag parsing
- **Optimized data structures** - efficient field merging and tree building
- **Typical extraction time**: < 100ms for most templates

## API Usage

You can use the extractor programmatically:

```csharp
using DocumentAssembler.Core;

// Load template
var templateDoc = new WmlDocument("MyTemplate.docx");

// Extract schema
var result = TemplateSchemaExtractor.ExtractXmlSchema(templateDoc);

// Access results
Console.WriteLine($"Root: {result.RootElementName}");
Console.WriteLine($"Fields: {result.Fields.Count}");
Console.WriteLine($"XML:\n{result.ToFormattedXml()}");
Console.WriteLine($"XSD:\n{result.ToFormattedXsd()}");

// Analyze fields
foreach (var field in result.Fields)
{
    Console.WriteLine($"{field.TagType}: {field.XPath}");
    Console.WriteLine($"  Optional: {field.IsOptional}");
    Console.WriteLine($"  Repeating: {field.IsRepeating}");
}
```

## Features

- **Automatic optional detection** - Respects the `Optional` attribute (defaults to true)
- **Repeating structure detection** - Identifies `<#Repeat#>` and `<#Table#>` collections
- **Hierarchical XML generation** - Builds proper nested structure from XPath expressions
- **Smart root element naming** - Detects the most common root element name
- **Conditional field handling** - Detects but doesn't include conditionals as data fields
- **Image placeholder support** - Identifies image fields and adds Base64 comments
- **Attribute preservation** - Captures all tag attributes (Match, NotMatch, Align, etc.)
- **XSD output** - Emits optional-aware XSD (`minOccurs`/`maxOccurs`) for DTO/validation generation

## Notes

- Fields are **optional by default** unless explicitly marked with `Optional="false"`
- **Conditional tags** are detected but not included in the data XML (they're for logic, not data)
- **Repeating structures** (Repeat/Table) show two examples in the generated XML for clarity
- The generated XML is a **template/guide** - replace `[value]` with actual data
- **Case-insensitive** comparison is used for field merging and lookup

## Integration with DocumentAssembler

This feature complements the main DocumentAssembler workflow:

1. **Design your template** in Word with DocumentAssembler tags
2. **Extract the XML schema** using `TemplateSchemaExtractor`
3. **Prepare your data** to match the extracted schema
4. **Assemble the document** using `DocumentAssembler.AssembleDocument`

This "reverse engineering" capability makes template development and integration much faster!
