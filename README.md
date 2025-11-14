# DocumentAssembler SDK

A lightweight, standalone SDK for assembling Word documents from templates with XML data binding. The library ships with tooling, templates, tests, and helper scripts so you can automate DOCX generation end-to-end.

## Project Origin

This SDK was **extracted from [Codeuctivity/OpenXmlPowerTools](https://github.com/Codeuctivity/OpenXmlPowerTools)** to give the DocumentAssembler module its own home. We kept the original MIT license, modernized the codebase, and curated only the pieces required for template-based generation.

### What We Modernized

- **Targeted .NET 10.0** with nullable reference types, implicit usings, and analyzers enabled across the SDK, tests, examples, and tools.
- **Split the original mega-file** into partial classes (`DocumentAssembler.cs`, `DocumentAssembler.Images.cs`, `DocumentAssembler.Signatures.cs`, `DocumentAssembler.Metadata.cs`) for maintainability.
- **Cleaned up** hundreds of unused symbols and dead code while keeping the OpenXmlPowerTools primitives we still rely on (`WmlDocument`, `OpenXmlRegex`, `RevisionProcessor`, `UnicodeMapper`, etc.).
- **Hardened metadata parsing**: inline `<#...#>` tokens and content controls are normalized, aliases validated, and invalid XML is turned into inline error runs instead of throwing.
- **Upgraded the image pipeline** to SkiaSharp 3.119.1 with PNG/JPEG/GIF fallbacks, width/height/max-dimension attributes, and paragraph-level alignment metadata.
- **Introduced PDF-ready signature placeholders** (`<Signature .../>`) backed by `SignaturePlaceholderSerializer` so DOCX → PDF pipelines can promote them to AcroForm fields.
- **Added a high-performance `TemplateSchemaExtractor`** that emits XML templates, optional-aware XSDs, and even lists native MailMerge fields.
- **Expanded documentation and examples** (15 sample projects plus generation scripts) so every tag/feature has a runnable template.
- **Maintained a broad xUnit suite** that covers DocumentAssembler core logic, schema extraction, revision processors, regex utilities, Unicode mapping, and signature tags with **85.7 % line coverage / 72.5 % branch coverage** (see "Testing & Quality").

## Project Objective

The DocumentAssembler SDK enables **template-based document generation** by merging Word templates (.docx) with XML data. This allows you to:

- Generate dynamic reports, contracts, and documents from templates
- Populate Word documents with data from databases, APIs, or XML files
- Create personalized documents at scale (letters, certificates, invoices)
- Maintain document formatting and styling while varying content
- Conditionally include/exclude sections based on data

**Key Use Case**: Replace manual document creation with automated, data-driven generation while preserving professional Word formatting.

## Supported Template Tags

The SDK supports eight template tags using the `<#TagName ... #>` syntax, plus optional attributes that control behavior.

| Tag | Purpose | When to Use |
|-----|---------|-------------|
| **Content** | Insert a single value from XML data | Display simple data fields (names, emails, dates, etc.) |
| **Table** | Generate dynamic table rows from a collection | Create tables with variable number of rows (order items, invoices, etc.) |
| **Conditional** | Include/exclude content based on conditions | Show different content based on data values (membership status, country-specific text) |
| **Else** | Provide alternative content when condition is false | Create if-else logic within Conditional blocks (show premium vs standard content) |
| **Repeat** | Repeat content blocks for each item in a collection | Generate multiple paragraphs or sections (department summaries, product listings) |
| **Image** | Insert images from base64-encoded data | Display dynamic images with size and alignment control (product photos, signatures) |
| **Signature** | Emit PDF-ready signature placeholders | Reserve signing spots with labels, size hints, and metadata |
| **Optional** | Flag placeholders as optional/required | Handle potentially missing data gracefully (middle names, optional fields) |

---

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
- `Align`: left, center, right, justify, start, end, distribute
- `Width` / `Height`: explicit dimensions (`px`, `pt`, `cm`, `mm`, `in`, `emu`)
- `MaxWidth` / `MaxHeight`: clamp dimensions while preserving aspect ratio

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

**NEW DEFAULT BEHAVIOR**: Fields are **optional by default**. Missing fields render as empty content without failing the assembly. Explicitly set `Optional="false"` when a field must exist.

```
<#Content Select="Customer/Name" Optional="false"#>
```

**Supported on**:
- `Content` (text placeholders)
- `Image` (image placeholders)
- `Repeat` (repeating blocks)

See the examples in this section for success/failure scenarios.

### 7. Signature (PDF-ready Signature Blocks)

Embed visible signature lines plus hidden metadata so downstream PDF pipelines can promote them to `/Sig` AcroForm widgets.

**Syntax**:
```
<#Signature Id="Responsabile"
            Label="Firma Responsabile"
            Width="220px"
            Height="60px"
            PageHint="2" />
```

**Attributes**:
- `Id` (**required**): unique identifier used as the PDF field name.
- `Label`: visible text shown next to the signature line (defaults to "Signature").
- `Width` / `Height`: optional size hints (any `px/cm/mm/in/pt/emu` unit).
- `PageHint`: optional positive integer to help external tools position the signature in multi-page PDFs.

When the assembler encounters `<Signature>`, it emits two runs:
1. A visible line such as `Firma Responsabile ____________________` so Word users see the signing spot.
2. A hidden run containing a Base64 payload wrapped in `[[DA_SIGN::...]]` (see `SignaturePlaceholderSerializer`). PDF converters scan for that token and create sealed signature fields with the captured metadata.

See [Example08_Signature](DocumentAssemblerSdk.Examples/Example08_Signature/) for a runnable template.

## Under the Hood

Inside `DocumentAssembler.Core` the pipeline:

- Loads a `WmlDocument` into memory, rejects templates that still have tracked revisions (`RevisionAccepter.HasTrackedRevisions`).
- Traverses **all content parts** (main document, headers, footers, footnotes, endnotes) so metadata can live anywhere inside the DOCX package.
- Normalizes content controls (`TransformToMetadata`) and inline `<#...#>` tokens, enforcing alias consistency and replacing malformed XML with inline highlighted error paragraphs.
- Lifts run-level metadata (`ForceBlockLevelAsAppropriate`) and fixes mis-leveled tables/conditionals before replacing content.
- Uses `XPathEvaluationContext` to cache XPath evaluations per data node, drastically reducing repeated XPath compilation and enabling consistent error reporting.
- Collects missing fields, invalid XPath expressions, and schema mismatches in a `TemplateError` object; callers can inspect the boolean `templateError` and the string `templateErrorSummary` returned by `AssembleDocument`.
- Streams images directly into OpenXML `ImagePart`s, auto-incrementing drawing IDs with `ImageIdTracker` so headers/footers remain valid.
- Serializes signature metadata via `SignaturePlaceholderSerializer` (Base64-encoded JSON inside `[[DA_SIGN::...]]`).

### Supporting Infrastructure

- `DocumentAssemblerSdk/Documents` provides `OpenXmlPowerToolsDocument` + `WmlDocument` wrappers for easy LINQ-to-XML traversal of parts.
- `DocumentAssemblerSdk/Utilities` includes:
  - `OpenXmlRegex` for regex replacements on OpenXML runs (with optional revision tracking).
  - `RevisionAccepter` and `RevisionProcessor` helpers to accept/reject tracked changes.
  - `UnicodeMapper`, `FontMetric`, `PtOpenXmlUtil`, and namespace helpers used by complex templates (see Example10 fonts).
- `Exceptions/` contains domain exceptions (`OpenXmlPowerToolsException`, `PowerToolsDocumentException`).

## Template Schema Extraction & Mail Merge Discovery

Use `TemplateSchemaExtractor` to reverse-engineer templates, generate optional-aware XML/XSD stubs, and enumerate MailMerge fields.

### Use Cases

- **Reverse-engineer templates** - Understand what data a template needs
- **API integration** - Generate XML schemas for automated document generation
- **Documentation** - Create data specifications from templates
- **Validation** - Ensure your XML data matches template requirements
- **Onboarding** - Help new developers understand template structure
- **MailMerge auditing** - Detect `MERGEFIELD` instructions alongside content controls

### Usage

```csharp
using DocumentAssembler.Core;

// Load your template
var templateDoc = new WmlDocument("MyTemplate.docx");

// Extract the schema
var result = TemplateSchemaExtractor.ExtractXmlSchema(templateDoc);

Console.WriteLine(result.ToFormattedXml());
Console.WriteLine(result.ToFormattedXsd());
Console.WriteLine($"Fields: {result.Fields.Count}");
Console.WriteLine($"Root: {result.RootElementName}");
```

### Mail Merge inspection helper

```csharp
var mergeFields = TemplateSchemaExtractor.ExtractMailMergeFields(templateDoc);
foreach (var field in mergeFields)
{
    Console.WriteLine($"{field.FieldName} => {field.XPath}");
}
```

The extractor walks every `MERGEFIELD` instruction, normalizes the MailMerge path (e.g. `«Customer.FirstName»` → `Customer/FirstName`), and exposes it via `FieldInfo` (`TagType == "MailMerge"`) and via the standalone `ExtractMailMergeFields` API. See Examples 11–15 for fixtures derived from real MailMerge templates.

### Example Output

Given a template with:
```
<#Content Select="Customer/Name"#>
<#Content Select="Customer/Email"#>
<#Repeat Select="Customer/Orders"#>
  <#Content Select="Product"#>
  <#Content Select="Quantity"#>
<#EndRepeat#>
```

you will get a fully formatted XML + XSD stub (see `Sample XSD Output` below). Example06 and Example07 dump actual console output.

### Sample XSD Output

```xml
<?xml version="1.0" encoding="utf-8"?>
<xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema">
  <xs:element name="Data">
    <xs:complexType>
      <xs:sequence>
        <xs:element name="Customer" minOccurs="0" maxOccurs="1">
          <xs:complexType>
            <xs:sequence>
              <xs:element name="Name" type="xs:string" minOccurs="0" maxOccurs="1" />
              <xs:element name="Email" type="xs:string" minOccurs="0" maxOccurs="1" />
              <xs:element name="Orders" minOccurs="0" maxOccurs="unbounded">
                <xs:complexType>
                  <xs:sequence>
                    <xs:element name="Product" type="xs:string" minOccurs="0" maxOccurs="1" />
                    <xs:element name="Quantity" type="xs:string" minOccurs="0" maxOccurs="1" />
                  </xs:sequence>
                </xs:complexType>
              </xs:element>
            </xs:sequence>
          </xs:complexType>
        </xs:element>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
</xs:schema>
```

### Features

- ✅ **Automatic optional detection** - Respects the `Optional` attribute (defaults to true)
- ✅ **Repeating structure detection** - Identifies `<#Repeat#>` and `<#Table#>` collections
- ✅ **Hierarchical XML generation** - Builds proper nested structure from XPath expressions
- ✅ **Smart root element naming** - Detects the most common root element name
- ✅ **Conditional field handling** - Detects but doesn't include conditionals as data fields
- ✅ **Image placeholder support** - Identifies image fields and adds Base64 comments
- ✅ **Attribute preservation** - Captures tag attributes (Match, NotMatch, Align, etc.)
- ✅ **XSD output** - Emits optional-aware XSD (`minOccurs`/`maxOccurs`) for DTO/validation generation
- ✅ **MailMerge discovery** - Tags MailMerge fields (`TagType == "MailMerge"`) and exposes them via `ExtractMailMergeFields`

See:
- [Example06_SchemaExtraction](DocumentAssemblerSdk.Examples/Example06_SchemaExtraction/) for basic extraction.
- [Example07_ComplexSchemaExtraction](DocumentAssemblerSdk.Examples/Example07_ComplexSchemaExtraction/) for a full real-world template with 35+ fields.
- Examples 11–15 for MailMerge-only, mixed, and CC-only templates that exercise the MailMerge inspector.

## Example Suite

| Example | Path | Highlights |
|---------|------|------------|
| Example 01 – Basic | `DocumentAssemblerSdk.Examples/Example01_Basic` | Minimal console app that assembles a simple template (content/table tags). |
| Example 02 – Intermediate | `DocumentAssemblerSdk.Examples/Example02_Intermediate` | Adds nested data and demonstrates error reporting/out file production. |
| Example 03 – Advanced | `DocumentAssemblerSdk.Examples/Example03_Advanced` | Handles bigger templates with Repeat + Table combos. |
| Example 04 – Images | `DocumentAssemblerSdk.Examples/Example04_Images` | Shows Base64 image placeholders with width/height metadata. |
| Example 05 – Conditional + Else | `DocumentAssemblerSdk.Examples/Example05_ConditionalElse` | Uses generated templates to showcase `<#Else#>` blocks. |
| Example 06 – Schema Extraction | `DocumentAssemblerSdk.Examples/Example06_SchemaExtraction` | Runs `TemplateSchemaExtractor` against a basic template. |
| Example 07 – Complex Schema | `DocumentAssemblerSdk.Examples/Example07_ComplexSchemaExtraction` | Stress-tests schema extraction with nested repeats, tables, and images. |
| Example 08 – Signature Placeholders | `DocumentAssemblerSdk.Examples/Example08_Signature` | Demonstrates the `<Signature>` tag and PDF-ready placeholders. |
| Example 09 – All Tags | `DocumentAssemblerSdk.Examples/Example09_AllTags` | End-to-end showcase for every tag plus Italian copy and sample JSON data. |
| Example 10 – Fonts | `DocumentAssemblerSdk.Examples/Example10_Fonts` | Validates font propagation + barcode fonts using custom SDT definitions. |
| Example 11 – MailMerge ATemplate | `DocumentAssemblerSdk.Examples/Example11_MailMerge_ATemplate` | Ensures MailMerge fields are detected in `ATemplate.docx`. |
| Example 12 – MailMerge Edited | `DocumentAssemblerSdk.Examples/Example12_MailMerge_Edited` | Validates MailMerge detection in edited Word templates. |
| Example 13 – MailMerge WordDocxWithALineBreak | `DocumentAssemblerSdk.Examples/Example13_MailMerge_WordDocxWithALineBreak` | Confirms no MailMerge fields exist when content controls are used. |
| Example 14 – MailMerge template-cc | `DocumentAssemblerSdk.Examples/Example14_MailMerge_TemplateCc` | Ensures CC-only templates yield zero MailMerge hits. |
| Example 15 – MailMerge template-cc-tag | `DocumentAssemblerSdk.Examples/Example15_MailMerge_TemplateCcTag` | Same as Example14 but with tagged controls for regression coverage. |

`Example09DataFactory.cs` and `sample-data.json` provide in-repo data generators so tests can assert the assembled output deterministically.

## Tooling & Helper Scripts

- `PerfMeasurementTool/` – CLI that assembles simple & complex templates ten times, drops the first run, and reports warm performance numbers.
- `generate_test_docx.py` – Builds DOCX fixtures (`DA270`–`DA272`) used by the test suite to validate nested Conditional/Else flows.
- `create_example05_templates.py`, `create_example08_signature.py`, `create_example09_all_tags.py`, `create_example10_fonts.py` – Rebuild the example templates directly from XML snippets.
- `requirements.txt` – Python dependency list (`python-docx==1.1.2`).
- `test_da272.sh` – Helper script that runs only the DA272 nested conditional tests via `dotnet test --filter`.
- `debug_test.csx` – Dotnet-script utility that inspects DOCX/XML fixtures during debugging sessions.

All helper scripts assume the repo root as the working directory.

## Document Generation Performance Baseline

`PerfMeasurementTool` is a CLI that assembles two templates and reports timing statistics in `Release` mode. Each scenario now executes ten runs and discards the first to account for JIT/startup noise. Run the measurement with:

```
dotnet run --project PerfMeasurementTool --configuration Release
```

The CLI uses `PerfMeasurementTool/SimpleTemplate.docx` and `PerfMeasurementTool/ComplexTemplate.docx`. Because the templates are intentionally minimal, the tool prints template-warning lines but the reported durations are valid.

### Baseline snapshot (immutable)

| Scenario | Template | Raw timings (ms) | Average (drop first run) |
| --- | --- | --- | --- |
| Simple | `PerfMeasurementTool/SimpleTemplate.docx` | 91.0, 2.0, 1.0, 0.9, 0.9, 1.0, 0.9, 1.3, 3.0, 0.9 | 1.3 |
| Complex | `PerfMeasurementTool/ComplexTemplate.docx` | 1.3, 1.2, 1.2, 1.4, 1.4, 1.8, 1.2, 1.2, 1.2, 1.4 | 1.3 |

> **This section is a permanent performance baseline. Do not edit these figures unless you replace the snapshot as part of a deliberate optimization milestone.**

## Building & Requirements

```bash
# Install .NET 10 SDK (preview as of today) then:
dotnet build DocumentAssemblerSdk/DocumentAssemblerSdk.csproj
```

> 💡 If you're still on .NET 9, set `DOTNET_ROLL_FORWARD=LatestMajor` before running CLI commands so the tooling will pick the latest installed SDK.

## Testing & Quality

```bash
# Run the entire suite
DOTNET_ROLL_FORWARD=LatestMajor dotnet test DocumentAssemblerSdk.Tests/DocumentAssemblerSdk.Tests.csproj

# Collect coverage via coverlet / XPlat collector
DOTNET_ROLL_FORWARD=LatestMajor dotnet test DocumentAssemblerSdk.Tests/DocumentAssemblerSdk.Tests.csproj --collect:"XPlat Code Coverage"
```

What we test:

- `DocumentAssemblerTests` – >70 theory cases covering headers, nested repeats, tables, invalid XPath, optional defaults, and tracked revision detection.
- `EnhancedErrorReportingTests`, `MissingFieldsTests`, `ExceptionTests` – Ensure TemplateError summaries, exceptions, and missing-field diagnostics remain stable.
- `SignatureTagTests` – Validate `<Signature>` attribute validation and placeholder serialization logic.
- `Example09AllTagsTests` – Guard the showcase template/sample data pair.
- `TemplateSchemaExtractorTests` – Cover XML tree building, optional propagation, mail merge detection, attribute handling, and huge template scenarios.
- `OpenXmlRegexTests`, `RevisionAccepterTests`, `RevisionProcessor*Tests`, `UnicodeMapperTests` – Regression tests for the auxiliary utilities we ship.

**Coverage snapshot**: `DocumentAssemblerSdk.Tests/TestResults/5fb1ed5a-85fb-46e0-b0db-41a2274ebe42/coverage.cobertura.xml` reports **85.7 % line coverage (6337/7394 lines)** and **72.5 % branch coverage (1684/2322 branches)** using the `XPlat Code Coverage` collector. Re-run the command above to refresh the report.

Temporary files land under `DocumentAssemblerSdk.Tests/TestResults/` (gitignored) and can be cleaned at any time.

## Project Structure

```
DocumentAssembler/
├── DocumentAssembler.sln
├── DocumentAssemblerSdk/               # Main library (Core, Documents, Exceptions, Utilities)
├── DocumentAssemblerSdk.Tests/         # xUnit test suite + fixtures + helpers
├── DocumentAssemblerSdk.Examples/      # 15 sample projects (see table above)
├── PerfMeasurementTool/                # CLI used for perf baselines
├── create_example05_templates.py       # Template builders
├── create_example08_signature.py
├── create_example09_all_tags.py
├── create_example10_fonts.py
├── generate_test_docx.py               # DOCX fixture generator
├── requirements.txt                    # Python dependencies for generators
├── test_da272.sh                       # Focused test runner
├── debug_test.csx                      # dotnet-script debugger helper
└── README.md                           # This file
```

## Dependencies

- **DocumentFormat.OpenXml** 3.3.0 – Rich OpenXML DOM APIs.
- **SkiaSharp** 3.119.1 (+ `SkiaSharp.NativeAssets.Linux.NoDependencies`) – Image decoding and measurement, cross-platform safe.
- **xUnit 2.9.3 + coverlet.collector 6.0.0** – Test + coverage stack.

## Target Framework

- **net10.0** across SDK, examples, tests, and tools.
- **Main Namespace**: `DocumentAssembler.Core`

## Generating Test Documents

For developers who need to create additional DOCX fixtures with content controls:

```bash
# Generate DA270/DA271/DA272 into the default test folder
python3 generate_test_docx.py

# Target a custom directory
python3 generate_test_docx.py /tmp/doc-out
```

The script copies `Example01_Basic/TemplateDocument.docx`, injects Conditional + Else tags, and saves the result under `DocumentAssemblerSdk.Tests/TestFiles/`.

## Acknowledgements

- ❤️ **[Codeuctivity/OpenXmlPowerTools](https://github.com/Codeuctivity/OpenXmlPowerTools)** – The original authors of DocumentAssembler, `OpenXmlRegex`, revision processors, and the supporting infrastructure that we extracted and modernized here.
- 🙏 **[chrisfcarroll/MailMerge](https://github.com/chrisfcarroll/MailMerge/tree/main)** – Source of the MailMerge sample templates that inspired the fixtures we validate through Examples 11–15 and the schema extractor tests.

## License

This project maintains the original MIT license from OpenXmlPowerTools.
