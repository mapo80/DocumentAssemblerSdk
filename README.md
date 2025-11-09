# DocumentAssembler SDK

A lightweight, standalone SDK for assembling Word documents from templates with XML data binding.

## Project Origin

This SDK was **extracted from [OpenXmlPowerTools](https://github.com/EricWhiteDev/Open-Xml-PowerTools)**, a comprehensive toolkit for working with Open XML documents. We isolated the DocumentAssembler module to create a focused, maintainable library specifically for template-based document generation.

### What We Did

- **Extracted** the DocumentAssembler module from OpenXmlPowerTools
- **Modernized** to .NET 8.0 with nullable reference types enabled
- **Cleaned up** 228 lines of dead code (unused classes, methods, and constructors)
- **Fixed** all nullable warnings for improved type safety
- **Split** large files into maintainable partial classes for better code organization
- **Maintained** 100% test coverage (107/107 tests passing)
- **Enhanced** with comprehensive image placeholder support (sizing, alignment, aspect ratio preservation)

## Project Objective

The DocumentAssembler SDK enables **template-based document generation** by merging Word templates (.docx) with XML data. This allows you to:

- Generate dynamic reports, contracts, and documents from templates
- Populate Word documents with data from databases, APIs, or XML files
- Create personalized documents at scale (letters, certificates, invoices)
- Maintain document formatting and styling while varying content
- Conditionally include/exclude sections based on data

**Key Use Case**: Replace manual document creation with automated, data-driven generation while preserving professional Word formatting.

## Supported Template Tags

The SDK supports six types of template tags, all using the format `<#TagName ... #>`:

### 1. Content (Simple Data Binding)

Insert a single value from your XML data.

**Syntax**: `<#Content Select="XPathExpression"#>`

**Example**:
```xml
<!-- XML Data -->
<Customer>
  <Name>John Doe</Name>
  <Email>john@example.com</Email>
</Customer>
```

```
<!-- Template -->
Dear <#Content Select="Customer/Name"#>,
Your email is: <#Content Select="Customer/Email"#>

<!-- Output -->
Dear John Doe,
Your email is: john@example.com
```

### 2. Table (Dynamic Table Generation)

Generate table rows by iterating over a collection in your XML data. The table in your template must have at least two rows: a header row and a prototype row that will be repeated for each data item.

**Syntax**: `<#Table Select="XPathToCollection"#>` (placed in the first cell of the prototype row)

**Example**:
```xml
<!-- XML Data -->
<Orders>
  <Order>
    <Product>Laptop</Product>
    <Quantity>2</Quantity>
    <Price>1200.00</Price>
  </Order>
  <Order>
    <Product>Mouse</Product>
    <Quantity>5</Quantity>
    <Price>25.00</Price>
  </Order>
</Orders>
```

```
<!-- Template (in Word table) -->
| Product                          | Quantity                       | Price                         |
|----------------------------------|--------------------------------|-------------------------------|
| <#Table Select="Orders/Order"#> |                                |                               |
| <#Content Select="Product"#>    | <#Content Select="Quantity"#> | <#Content Select="Price"#>   |

<!-- Output -->
| Product | Quantity | Price    |
|---------|----------|----------|
| Laptop  | 2        | 1200.00  |
| Mouse   | 5        | 25.00    |
```

### 3. Conditional (Conditional Inclusion)

Include or exclude content based on whether an XML value matches a specific condition.

**Syntax**:
- `<#Conditional Select="XPath" Match="Value"#>...<#EndConditional#>` (include if matches)
- `<#Conditional Select="XPath" NotMatch="Value"#>...<#EndConditional#>` (include if doesn't match)

**Example**:
```xml
<!-- XML Data -->
<Customer>
  <Country>USA</Country>
  <Premium>true</Premium>
</Customer>
```

```
<!-- Template -->
<#Conditional Select="Customer/Country" Match="USA"#>
Shipping: Free domestic shipping within the USA.
<#EndConditional#>

<#Conditional Select="Customer/Premium" Match="true"#>
As a premium member, you receive 20% off all purchases.
<#EndConditional#>

<#Conditional Select="Customer/Country" NotMatch="USA"#>
International shipping rates apply.
<#EndConditional#>

<!-- Output (for the above data) -->
Shipping: Free domestic shipping within the USA.
As a premium member, you receive 20% off all purchases.
```

### 4. Repeat (Repeating Content Blocks)

Repeat a section of content (paragraphs, tables, etc.) for each item in a collection.

**Syntax**: `<#Repeat Select="XPathToCollection"#>...<#EndRepeat#>`

**Example**:
```xml
<!-- XML Data -->
<Departments>
  <Department>
    <Name>Engineering</Name>
    <HeadCount>50</HeadCount>
  </Department>
  <Department>
    <Name>Sales</Name>
    <HeadCount>30</HeadCount>
  </Department>
</Departments>
```

```
<!-- Template -->
<#Repeat Select="Departments/Department"#>
Department: <#Content Select="Name"#>
Headcount: <#Content Select="HeadCount"#>
---
<#EndRepeat#>

<!-- Output -->
Department: Engineering
Headcount: 50
---
Department: Sales
Headcount: 30
---
```

### 5. Image (Dynamic Image Insertion)

Insert images from base64-encoded data with optional sizing and alignment metadata.

**Syntax**: `<#Image Select="XPathToBase64Data"#>`

**Supported Metadata Attributes**:
- `Align`: left, center, right, or justify
- `Width`: explicit width (e.g., "300px", "5cm")
- `Height`: explicit height (e.g., "200px", "3in")
- `MaxWidth`: maximum width constraint
- `MaxHeight`: maximum height constraint

**Example**:
```xml
<!-- XML Data -->
<Product>
  <Name>Laptop</Name>
  <Photo Align="center" MaxWidth="400px" MaxHeight="300px">
    iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJ...
  </Photo>
</Product>
```

```
<!-- Template -->
Product: <#Content Select="Product/Name"#>

<#Image Select="Product/Photo"#>

<!-- Output -->
A Word document with the product name and a centered image
scaled to fit within 400x300px while preserving aspect ratio.
```

### 6. Optional (Optional Placeholders)

Designate a placeholder as optional, preventing error messages if the data is missing.

**Syntax**: `<#Content Select="XPath" Optional="true"#>`

**Example**:
```xml
<!-- XML Data -->
<Customer>
  <Name>John Doe</Name>
  <!-- MiddleName is missing -->
  <LastName>Doe</LastName>
</Customer>
```

```
<!-- Template -->
Full Name: <#Content Select="Customer/Name"#> <#Content Select="Customer/MiddleName" Optional="true"#> <#Content Select="Customer/LastName"#>

<!-- Output -->
Full Name: John Doe  Doe
(No error for missing MiddleName)
```

## Quick Start

### Installation

Add a project reference to DocumentAssemblerSdk in your .csproj:

```xml
<ItemGroup>
  <ProjectReference Include="../DocumentAssemblerSdk/DocumentAssemblerSdk.csproj" />
</ItemGroup>
```

### Basic Usage

```csharp
using DocumentAssembler.Core;
using System.Xml;

// Load your template
var templateDoc = new WmlDocument("Template.docx");

// Load your data
var xmlDoc = new XmlDocument();
xmlDoc.Load("Data.xml");

// Assemble the document
var assembledDoc = DocumentAssembler.AssembleDocument(
    templateDoc,
    xmlDoc,
    out bool templateError
);

// Save the result
assembledDoc.SaveAs("Output.docx");

if (templateError)
{
    Console.WriteLine("There were errors in the template - check the output document");
}
```

## Project Structure

```
DocumentAssembler/
├── DocumentAssemblerSdk/           # Main library
├── DocumentAssemblerSdk.Tests/     # 107 unit tests (100% passing)
└── README.md                        # This file
```

## Building the Project

```bash
# Build the library
dotnet build DocumentAssemblerSdk/DocumentAssemblerSdk.csproj

# Run tests (107/107 passing)
dotnet test DocumentAssemblerSdk.Tests/DocumentAssemblerSdk.Tests.csproj
```

## Dependencies

- **DocumentFormat.OpenXml** (3.3.0) - Core OpenXML functionality
- **SkiaSharp** (3.119.1) - Image processing for the Image placeholder feature
- **SkiaSharp.NativeAssets.Linux.NoDependencies** (3.119.1) - Cross-platform support

## Target Framework

- **.NET 8.0** with nullable reference types enabled
- **Main Namespace**: `DocumentAssembler.Core`

## License

This project maintains the original MIT license from OpenXmlPowerTools.
