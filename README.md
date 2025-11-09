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
- **Maintained** 100% test coverage (117/117 tests passing)
- **Enhanced** with comprehensive image placeholder support (sizing, alignment, aspect ratio preservation)
- **Added** Else tag for if-else logic in Conditional blocks with full nested support

## Project Objective

The DocumentAssembler SDK enables **template-based document generation** by merging Word templates (.docx) with XML data. This allows you to:

- Generate dynamic reports, contracts, and documents from templates
- Populate Word documents with data from databases, APIs, or XML files
- Create personalized documents at scale (letters, certificates, invoices)
- Maintain document formatting and styling while varying content
- Conditionally include/exclude sections based on data

**Key Use Case**: Replace manual document creation with automated, data-driven generation while preserving professional Word formatting.

## Supported Template Tags

The SDK supports seven types of template tags, all using the format `<#TagName ... #>`:

| Tag | Purpose | When to Use |
|-----|---------|-------------|
| **Content** | Insert a single value from XML data | Display simple data fields (names, emails, dates, etc.) |
| **Table** | Generate dynamic table rows from a collection | Create tables with variable number of rows (order items, invoices, etc.) |
| **Conditional** | Include/exclude content based on conditions | Show different content based on data values (membership status, country-specific text) |
| **Else** | Provide alternative content when condition is false | Create if-else logic within Conditional blocks (show premium vs standard content) |
| **Repeat** | Repeat content blocks for each item in a collection | Generate multiple paragraphs or sections (department summaries, product listings) |
| **Image** | Insert images from base64-encoded data | Display dynamic images with size and alignment control (product photos, signatures) |
| **Optional** | Mark placeholders as optional to avoid errors | Handle potentially missing data gracefully (middle names, optional fields) |

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
- `<#Conditional Select="XPath" Match="Value"#>...<#Else#>...<#EndConditional#>` (if-else structure)

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

#### Conditional with Else (If-Else Logic)

The `<#Else#>` tag provides an **alternative content block** when the condition is false. This is **optional** and must be placed between a `<#Conditional#>` and its matching `<#EndConditional#>`.

**Syntax**:
```
<#Conditional Select="XPath" Match="Value"#>
  Content shown when condition is TRUE
<#Else#>
  Content shown when condition is FALSE
<#EndConditional#>
```

**Key Points**:
- `<#Else#>` is **optional** - you can use Conditional without it
- `<#Else#>` must be inside a Conditional block (between Conditional and EndConditional)
- You can nest Conditionals with Else inside other Conditionals
- Works with both `Match` and `NotMatch` attributes
- Supports both **string** and **numeric** value matching

---

**Example 1: Basic If-Else with String Matching**

```xml
<!-- XML Data -->
<Customer>
  <MembershipType>Premium</MembershipType>
</Customer>
```

```
<!-- Template -->
<#Conditional Select="Customer/MembershipType" Match="Premium"#>
✓ You are a PREMIUM member! Benefits:
  • 20% discount on all purchases
  • Free shipping worldwide
  • Priority customer support
<#Else#>
You are a Standard member. Benefits:
  • 5% discount on purchases
  • Standard shipping rates apply
<#EndConditional#>

Thank you for being our customer!
```

```
<!-- Output (when MembershipType = "Premium") -->
✓ You are a PREMIUM member! Benefits:
  • 20% discount on all purchases
  • Free shipping worldwide
  • Priority customer support

Thank you for being our customer!

<!-- Output (when MembershipType = "Standard" or any other value) -->
You are a Standard member. Benefits:
  • 5% discount on purchases
  • Standard shipping rates apply

Thank you for being our customer!
```

---

**Example 2: If-Else with NotMatch**

```xml
<!-- XML Data -->
<Customer>
  <Country>Canada</Country>
</Customer>
```

```
<!-- Template -->
<#Conditional Select="Customer/Country" NotMatch="USA"#>
📦 International shipping:
  • Shipping costs apply based on location
  • Estimated delivery: 10-15 business days
  • Customs fees may apply
<#Else#>
📦 Domestic USA shipping:
  • FREE shipping on all orders!
  • Estimated delivery: 2-3 business days
<#EndConditional#>
```

```
<!-- Output (when Country = "Canada", "UK", etc. - anything except "USA") -->
📦 International shipping:
  • Shipping costs apply based on location
  • Estimated delivery: 10-15 business days
  • Customs fees may apply

<!-- Output (when Country = "USA") -->
📦 Domestic USA shipping:
  • FREE shipping on all orders!
  • Estimated delivery: 2-3 business days
```

---

**Example 3: Nested Conditionals with Else (Advanced)**

You can nest Conditional blocks with Else inside other Conditionals for complex logic.

```xml
<!-- XML Data -->
<Customer>
  <MembershipType>Premium</MembershipType>
  <Points>5000</Points>
</Customer>
```

```
<!-- Template -->
<#Conditional Select="Customer/MembershipType" Match="Premium"#>
🌟 Premium Member Benefits:

  <#Conditional Select="Customer/Points" Match="5000"#>
  ⭐ PLATINUM TIER - You've reached 5000 points!
    • Exclusive access to VIP lounge
    • Personal account manager
    • 25% discount on all purchases
  <#Else#>
  You have Premium status
    • 20% discount on purchases
    • Priority support
  <#EndConditional#>

<#Else#>
Standard Member Benefits:
  • 5% discount on purchases
  • Email support available
<#EndConditional#>
```

```
<!-- Output (when MembershipType = "Premium" AND Points = 5000) -->
🌟 Premium Member Benefits:

  ⭐ PLATINUM TIER - You've reached 5000 points!
    • Exclusive access to VIP lounge
    • Personal account manager
    • 25% discount on all purchases

<!-- Output (when MembershipType = "Premium" AND Points ≠ 5000) -->
🌟 Premium Member Benefits:

  You have Premium status
    • 20% discount on purchases
    • Priority support

<!-- Output (when MembershipType ≠ "Premium") -->
Standard Member Benefits:
  • 5% discount on purchases
  • Email support available
```

---

**Example 4: Numeric Value Matching**

The Conditional tag works with numeric values as well as strings.

```xml
<!-- XML Data -->
<Order>
  <TotalAmount>150</TotalAmount>
</Order>
```

```
<!-- Template -->
<#Conditional Select="Order/TotalAmount" Match="150"#>
🎉 Your order qualifies for a special bonus!
<#Else#>
Add more items to unlock bonus rewards.
<#EndConditional#>
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

### 6. Optional (Handling Missing Fields)

**NEW DEFAULT BEHAVIOR**: Fields are now **optional by default**. Missing fields will appear empty in the output without causing errors.

To make a field **required** (causing an error if missing), explicitly set `Optional="false"`:

```
<#Content Select="Customer/Name" Optional="false"#>
```

**Supported on**:
- `Content` (text placeholders)
- `Image` (image placeholders)
- `Repeat` (repeating blocks)

---

**Syntax**:
- `<#Content Select="XPath"#>` → Optional by default (no error if missing)
- `<#Content Select="XPath" Optional="false"#>` → Required (error if missing)
- `<#Content Select="XPath" Optional="true"#>` → Explicitly optional (for clarity)

---

**Example 1: Missing Field with Default Behavior (SUCCEEDS)**

```xml
<!-- XML Data -->
<Customer>
  <Name>John Doe</Name>
  <!-- MiddleName is MISSING -->
  <LastName>Smith</LastName>
</Customer>
```

```
<!-- Template (fields are optional by default) -->
Full Name: <#Content Select="Customer/Name"#> <#Content Select="Customer/MiddleName"#> <#Content Select="Customer/LastName"#>

<!-- Output: SUCCESS ✅ -->
Full Name: John Doe  Smith
(MiddleName is missing, appears as empty - no error!)
```

---

**Example 2: Required Field Missing (FAILS)**

```xml
<!-- XML Data -->
<Customer>
  <Email>john@example.com</Email>
  <!-- Name is MISSING -->
</Customer>
```

```
<!-- Template with Optional="false" to REQUIRE the field -->
Customer Name: <#Content Select="Customer/Name" Optional="false"#>

<!-- Result: ERROR ❌ -->
XPathException: XPath expression (Customer/Name) returned no results
Document generation FAILS
(Use Optional="false" when a field is absolutely required)
```

---

**Example 3: Image (Optional by Default)**

```xml
<!-- XML Data -->
<Product>
  <Name>Basic Widget</Name>
  <!-- Photo is MISSING -->
</Product>
```

```
<!-- Template (Image is optional by default) -->
Product: <#Content Select="Product/Name"#>
<#Image Select="Product/Photo" MaxWidth="200px"#>

<!-- Output: SUCCESS ✅ -->
Product: Basic Widget
(No image shown, no error - images are optional by default)
```

---

**Example 4: Repeat (Optional by Default)**

```xml
<!-- XML Data -->
<Report>
  <Title>Annual Summary</Title>
  <!-- Items collection is MISSING -->
</Report>
```

```
<!-- Template (Repeat is optional by default) -->
<#Repeat Select="Report/Items/Item"#>
- <#Content Select="./Name"#>
<#EndRepeat#>

<!-- Output: SUCCESS ✅ -->
(No items repeated, no error - Repeat is optional by default)
```

---

**Best Practices**:
1. ✅ **Fields are optional by default** - no need to add `Optional="true"` unless for clarity
2. ✅ **Use `Optional="false"` for critical fields** that must always be present (IDs, names, required info)
3. ✅ **Test your templates with data that has missing fields** to verify graceful handling
4. ✅ **Combine with Conditional tags** for more complex logic when fields are missing
5. 📝 **Backward compatibility**: Old templates with `Optional="true"` continue to work identically

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

# Run tests (117/117 passing)
dotnet test DocumentAssemblerSdk.Tests/DocumentAssemblerSdk.Tests.csproj
```

## Generating Test Documents

For developers who need to create test .docx files with content controls, we provide a Python script:

```bash
# Generate all test documents
python3 generate_test_docx.py

# Generate to a specific directory
python3 generate_test_docx.py /path/to/output/dir

# View help
python3 generate_test_docx.py --help
```

**What it generates**:
- `DA270-ConditionalWithElse.docx` - Basic Conditional with Else using Match
- `DA271-ConditionalWithElseNotMatch.docx` - Conditional with Else using NotMatch
- `DA272-NestedConditionalWithElse.docx` - Nested Conditionals with Else

**Requirements**:
- Python 3.x (no external dependencies needed)
- Source template: `DocumentAssemblerSdk.Examples/Example01_Basic/TemplateDocument.docx`

The script creates valid .docx files by copying the template structure and injecting custom document.xml with proper content controls. This is useful for creating test files without manually editing Word documents.

## Dependencies

- **DocumentFormat.OpenXml** (3.3.0) - Core OpenXML functionality
- **SkiaSharp** (3.119.1) - Image processing for the Image placeholder feature
- **SkiaSharp.NativeAssets.Linux.NoDependencies** (3.119.1) - Cross-platform support

## Target Framework

- **.NET 8.0** with nullable reference types enabled
- **Main Namespace**: `DocumentAssembler.Core`

## License

This project maintains the original MIT license from OpenXmlPowerTools.
